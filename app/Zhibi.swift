// 执笔 (AI reply drafting) for new Outlook for Mac, styled after Outlook's Copilot: a round button left of the
// Copilot button and a chat pane docked over the Copilot pane. Both are floating panels that follow
// Outlook's main window.
// The pane drafts replies with Resources/scripts/ai_reply.py (claude -p) and, on request, opens Outlook's
// reply window with the formatted body. It never sends mail.
import ApplicationServices
import Carbon.HIToolbox
import Cocoa
import ServiceManagement
import WebKit

let outlookID = "com.microsoft.Outlook"
let pipeline = Bundle.main.resourcePath! + "/scripts/ai_reply.py"
let supportDir = NSHomeDirectory() + "/Library/Application Support/Zhibi"

// MARK: - config, log

func expand(_ p: String) -> String { (p as NSString).expandingTildeInPath }

struct Config {
    var outDir = supportDir + "/drafts"
    var logFile = "~/Library/Logs/Zhibi/zhibi.log"

    static func load() -> Config {
        var c = Config()
        if let d = try? Data(contentsOf: URL(fileURLWithPath: supportDir + "/config.json")),
           let j = try? JSONSerialization.jsonObject(with: d) as? [String: Any] {
            if let v = j["out_dir"] as? String { c.outDir = v }
            if let v = j["log_file"] as? String { c.logFile = v }
        }
        return c
    }
}

let config = Config.load()
let stamp: DateFormatter = {
    let f = DateFormatter()
    f.locale = Locale(identifier: "en_US_POSIX")
    f.dateFormat = "yyyy-MM-dd HH:mm:ss"
    return f
}()

func log(_ msg: String) {
    let path = expand(config.logFile)
    try? FileManager.default.createDirectory(atPath: (path as NSString).deletingLastPathComponent, withIntermediateDirectories: true)
    let data = "\(stamp.string(from: Date())) \(msg)\n".data(using: .utf8)!
    if let h = FileHandle(forWritingAtPath: path) {
        h.seekToEndOfFile()
        h.write(data)
        try? h.close()
    } else {
        FileManager.default.createFile(atPath: path, contents: data)
    }
}

/// outlook-cli's own interpreter (has exchangelib and the keyring login). The npm build of outlook-cli is a
/// frozen binary without an importable package, so only pip / pipx / uv installs count.
func pythonPath() -> String? {
    for cli in [NSHomeDirectory() + "/.local/bin/outlook-cli", "/opt/homebrew/bin/outlook-cli", "/usr/local/bin/outlook-cli"] {
        guard let text = try? String(contentsOfFile: cli, encoding: .utf8), text.hasPrefix("#!") else { continue }
        let line = text.split(separator: "\n", maxSplits: 1).first.map(String.init) ?? ""
        let py = String(line.dropFirst(2)).trimmingCharacters(in: .whitespaces)
        if (py as NSString).lastPathComponent.hasPrefix("python") { return py }
    }
    return nil
}

// MARK: - AX helpers

func attr(_ e: AXUIElement, _ n: String) -> AnyObject? {
    var v: AnyObject?
    return AXUIElementCopyAttributeValue(e, n as CFString, &v) == .success ? v : nil
}
func str(_ e: AXUIElement, _ n: String) -> String { (attr(e, n) as? String) ?? "" }
func role(_ e: AXUIElement) -> String { str(e, kAXRoleAttribute) }
func ident(_ e: AXUIElement) -> String { str(e, "AXIdentifier") }
func kids(_ e: AXUIElement) -> [AXUIElement] { (attr(e, kAXChildrenAttribute) as? [AXUIElement]) ?? [] }

@_silgen_name("_AXUIElementGetWindow")
func axWindowID(_ e: AXUIElement, _ id: UnsafeMutablePointer<CGWindowID>) -> AXError

// Read-only window-server calls (private, but long stable) to tell which Space a window is in.
private let cgs = dlopen(nil, RTLD_NOW)
private let cgsConnection = dlsym(cgs, "CGSMainConnectionID").map { unsafeBitCast($0, to: (@convention(c) () -> Int32).self)() }
private let cgsActiveSpace = dlsym(cgs, "CGSGetActiveSpace").map { unsafeBitCast($0, to: (@convention(c) (Int32) -> UInt64).self) }
private let cgsSpacesForWindows = dlsym(cgs, "CGSCopySpacesForWindows")
    .map { unsafeBitCast($0, to: (@convention(c) (Int32, Int32, CFArray) -> Unmanaged<CFArray>?).self) }

/// Whether the window is in the Space now showing; nil when the window server does not say.
func inActiveSpace(_ id: CGWindowID) -> Bool? {
    guard let c = cgsConnection, let active = cgsActiveSpace, let spacesFor = cgsSpacesForWindows,
          let spaces = spacesFor(c, 7, [NSNumber(value: id)] as CFArray)?.takeRetainedValue() as? [UInt64], !spaces.isEmpty else { return nil }
    return spaces.contains(active(c))
}

func frame(_ e: AXUIElement) -> CGRect? {
    guard let p = attr(e, kAXPositionAttribute), let z = attr(e, kAXSizeAttribute) else { return nil }
    var pt = CGPoint.zero, sz = CGSize.zero
    AXValueGetValue(p as! AXValue, .cgPoint, &pt)
    AXValueGetValue(z as! AXValue, .cgSize, &sz)
    return sz.width > 0 && sz.height > 0 ? CGRect(origin: pt, size: sz) : nil
}

// Mail bodies, the message list and the folder tree are huge; the chrome we need is never inside them.
let heavyRoles: Set<String> = ["AXWebArea", "AXTable", "AXOutline"]

func findAll(_ root: AXUIElement, depth: Int, limit: Int = .max, _ pred: (AXUIElement) -> Bool) -> [AXUIElement] {
    var out: [AXUIElement] = []
    var queue: [(AXUIElement, Int)] = [(root, 0)]
    var i = 0
    while i < queue.count && out.count < limit {
        let (e, d) = queue[i]
        i += 1
        if pred(e) { out.append(e) }
        if d < depth && !heavyRoles.contains(role(e)) { queue += kids(e).map { ($0, d + 1) } }
    }
    return out
}
func find(_ root: AXUIElement, depth: Int, _ pred: (AXUIElement) -> Bool) -> AXUIElement? {
    findAll(root, depth: depth, limit: 1, pred).first
}

/// AX uses a top-left origin on the primary screen; Cocoa uses bottom-left.
func cocoaRect(_ r: CGRect) -> NSRect {
    let h = NSScreen.screens.first?.frame.height ?? 0
    return NSRect(x: r.minX, y: h - r.maxY, width: r.width, height: r.height)
}

// MARK: - Outlook tracking

struct MailHeader {
    let subject: String
    let from: String
    let folder: String
    var key: String { subject + "|" + from }
}

final class Tracker {
    private var pid: pid_t = 0
    private var axApp: AXUIElement?
    private(set) var mainWin: AXUIElement?
    private var anchor: AXUIElement?
    private(set) var anchorFrame = CGRect.zero
    private var blocker: CGFloat = 0
    private var rescanAt: Date?
    private var header: AXUIElement?
    private var replyButton: AXUIElement?
    private var split: AXUIElement?
    private var nextSearch = Date.distantPast
    private var nextPaneSearch = Date.distantPast
    private var windowID: CGWindowID = 0
    private(set) var fullScreen = false
    /// Changes when Outlook's main window moves to another Space (new window, full screen on/off).
    var spaceKey: String { "\(windowID)/\(fullScreen)" }

    var outlook: NSRunningApplication? { NSRunningApplication.runningApplications(withBundleIdentifier: outlookID).first }

    private func reset() {
        axApp = nil; mainWin = nil; anchor = nil; header = nil; replyButton = nil; split = nil; pid = 0
        windowID = 0; fullScreen = false
    }

    /// Main window state from the window server, since AX lists no windows while Outlook sits in another Space.
    /// nil = gone; true = its Space is the one showing (a Space switch shows both while it animates).
    func mainShowing() -> Bool? {
        guard windowID != 0 else { return mainWin == nil ? nil : true }
        guard let info = (CGWindowListCopyWindowInfo(.optionIncludingWindow, windowID) as? [[String: Any]])?.first else { return nil }
        return inActiveSpace(windowID) ?? (info[kCGWindowIsOnscreen as String] as? Bool) ?? false
    }

    private func alive(_ e: AXUIElement?) -> Bool { e.map { !role($0).isEmpty } ?? false }

    /// Round button frame in AX coordinates, or nil when it should be hidden.
    func placement() -> CGRect? {
        guard let app = outlook else { reset(); return nil }
        if app.processIdentifier != pid {
            reset()
            pid = app.processIdentifier
            let ax = AXUIElementCreateApplication(pid)
            AXUIElementSetMessagingTimeout(ax, 0.5)
            axApp = ax
        }
        guard let ax = axApp, let fw = attr(ax, kAXFocusedWindowAttribute) else { return nil }
        let focused = fw as! AXUIElement
        if mainWin == nil || !CFEqual(mainWin!, focused) {
            // Compose windows, dialogs and Settings have no search bar.
            guard find(focused, depth: 3, { ident($0) == "Search Bar" }) != nil else { return nil }
            mainWin = focused; anchor = nil; header = nil; replyButton = nil; split = nil
            var wid: CGWindowID = 0
            windowID = axWindowID(focused, &wid) == .success ? wid : 0
        }
        guard let main = mainWin else { return nil }
        fullScreen = (attr(main, "AXFullScreen") as? Bool) ?? false

        let now = Date()
        if (!alive(header) || !alive(anchor)) && now >= nextSearch {
            nextSearch = now.addingTimeInterval(1)
            if !alive(header) {
                header = find(main, depth: 9) { ident($0) == "HeaderMainContainer" }
                replyButton = header == nil ? nil : find(main, depth: 10) { ident($0) == "messageHeaderContextualControlsReplyButton" }
            }
            if !alive(anchor) { scanToolbar(main) }
        }
        guard alive(header) else { return nil }

        if let a = anchor, let af = frame(a) {
            if af != anchorFrame { anchorFrame = af; rescanAt = now.addingTimeInterval(0.6) }
            // The ribbon reflows after a resize; re-measure free space once it settles.
            if let t = rescanAt, now >= t { rescanAt = nil; scanToolbar(main) }
            for d: CGFloat in [40, 30] where af.minX - 10 - d >= blocker + 6 {
                return CGRect(x: af.minX - 10 - d, y: af.midY - d / 2, width: d, height: d)
            }
        }
        // No room in the toolbar: sit left of the reading pane's reply icon.
        if !alive(replyButton) { replyButton = find(main, depth: 10) { ident($0) == "messageHeaderContextualControlsReplyButton" } }
        if let r = replyButton, let rf = frame(r) {
            return CGRect(x: rf.minX - 8 - 28, y: rf.midY - 14, width: 28, height: 28)
        }
        return nil
    }

    /// Where the chat pane goes: over Outlook's right-hand side pane (Copilot, an add-in, 我的一天),
    /// or a right-hand column over the reading pane when no side pane is open.
    func paneFrame() -> CGRect? {
        guard let main = mainWin, let wf = frame(main) else { return nil }
        let now = Date()
        if !alive(split) && now >= nextPaneSearch {
            nextPaneSearch = now.addingTimeInterval(1)
            split = kids(main).first { role($0) == "AXSplitGroup" }
        }
        if let sg = split, let sf = frame(sg) {
            // The side pane is whatever sits right of the last splitter, if that splitter is far right.
            let edge = kids(sg).filter { role($0) == "AXSplitter" }.compactMap { e -> CGFloat? in
                var p = CGPoint.zero
                guard let v = attr(e, kAXPositionAttribute) else { return nil }
                AXValueGetValue(v as! AXValue, .cgPoint, &p)
                return p.x
            }.max() ?? 0
            if edge > sf.minX + sf.width * 0.55 && sf.maxX - edge >= 240 {
                return CGRect(x: edge + 4, y: sf.minY + 2, width: sf.maxX - edge - 12, height: sf.height - 10)
            }
            return CGRect(x: sf.maxX - 8 - 360, y: sf.minY + 2, width: 360, height: sf.height - 10)
        }
        let top = anchorFrame == .zero ? wf.minY + 104 : anchorFrame.maxY + 6
        return CGRect(x: wf.maxX - 8 - 360, y: top, width: 360, height: wf.maxY - 8 - top)
    }

    /// Anchor = the toolbar's Copilot button, else the side-pane toggle at the far right of the toolbar row;
    /// blocker = right edge of the last ribbon command left of the anchor.
    private func scanToolbar(_ main: AXUIElement) {
        guard let wf = frame(main) else { return }
        let top = wf.minY + 130
        let buttons = findAll(main, depth: 9) { e in
            let r = role(e)
            guard r == "AXButton" || r == "AXMenuButton" || r == "AXPopUpButton" || r == "AXCheckBox" else { return false }
            guard let f = frame(e) else { return false }
            return f.minY < top && f.minY > wf.minY + 40
        }
        let labelled = buttons.map { b -> (AXUIElement, CGRect, String) in
            let text = [kAXDescriptionAttribute, kAXTitleAttribute, kAXHelpAttribute].map { str(b, $0) }.joined(separator: " ")
            return (b, frame(b) ?? .zero, text)
        }
        let toolbarRow = labelled.filter { $0.1.minY < wf.minY + 104 }
        guard let a = toolbarRow.first(where: { $0.2.localizedCaseInsensitiveContains("copilot") })
            ?? toolbarRow.filter({ $0.1.minX > wf.maxX - 80 }).max(by: { $0.1.minX < $1.1.minX }) else {
            anchor = nil
            return
        }
        anchor = a.0
        anchorFrame = a.1
        blocker = labelled
            .filter { abs($0.1.midY - a.1.midY) < 10 && $0.1.maxX <= a.1.minX - 1 }
            .map { $0.1.maxX }.max() ?? 0
    }

    func readHeader() -> MailHeader? {
        guard let main = mainWin else { return nil }
        let h = alive(header) ? header! : find(main, depth: 9) { ident($0) == "HeaderMainContainer" }
        guard let head = h else { return nil }
        let subject = kids(head).first { role($0) == "AXStaticText" }.map { str($0, kAXValueAttribute) } ?? ""
        let from = find(head, depth: 4) { ident($0) == "messageHeaderFromContent" }.map { str($0, kAXValueAttribute) } ?? ""
        // Title looks like "收件箱 • name@example.com".
        let parts = str(main, kAXTitleAttribute).components(separatedBy: " • ")
        let folder = parts.count > 1 ? parts[0] : ""
        guard !subject.isEmpty else { return nil }
        return MailHeader(subject: subject, from: from, folder: folder.hasPrefix("正在搜索") ? "" : folder)
    }
}

// MARK: - panels

final class KeyPanel: NSPanel {
    override var canBecomeKey: Bool { true }
}

func floatingPanel(nonActivating: Bool) -> KeyPanel {
    let style: NSWindow.StyleMask = nonActivating ? [.borderless, .nonactivatingPanel] : [.borderless]
    let p = KeyPanel(contentRect: .zero, styleMask: style, backing: .buffered, defer: false)
    p.isOpaque = false
    p.backgroundColor = .clear
    p.hasShadow = false
    p.level = .floating
    p.hidesOnDeactivate = false
    p.isReleasedWhenClosed = false
    // Live in Outlook's Space (often a full-screen one), not all Spaces, so Space switches carry the panels along.
    p.collectionBehavior = [.managed, .fullScreenAuxiliary, .ignoresCycle]
    return p
}

/// 执笔's mark: a fountain-pen nib, tip down-left, drawn into `rect` (also used for the app icon).
func drawNib(in rect: NSRect) {
    let k = rect.width / 32
    let t = NSAffineTransform()
    // The nib is defined in a 32-unit y-down box like the SVG in pane.html.
    t.translateX(by: rect.midX, yBy: rect.midY)
    t.scaleX(by: k, yBy: -k)
    t.rotate(byDegrees: 45)
    t.translateX(by: -16, yBy: -16)
    let body = NSBezierPath()
    body.move(to: NSPoint(x: 10.5, y: 3))
    body.line(to: NSPoint(x: 21.5, y: 3))
    body.curve(to: NSPoint(x: 25, y: 14), controlPoint1: NSPoint(x: 21.5, y: 7.5), controlPoint2: NSPoint(x: 25, y: 9.5))
    body.line(to: NSPoint(x: 16, y: 30))
    body.line(to: NSPoint(x: 7, y: 14))
    body.curve(to: NSPoint(x: 10.5, y: 3), controlPoint1: NSPoint(x: 7, y: 9.5), controlPoint2: NSPoint(x: 10.5, y: 7.5))
    body.close()
    let hole = NSBezierPath(ovalIn: NSRect(x: 14, y: 13, width: 4, height: 4))
    let slit = NSBezierPath()
    slit.move(to: NSPoint(x: 16, y: 16))
    slit.line(to: NSPoint(x: 16, y: 30))
    for path in [body, hole, slit] { path.transform(using: t as AffineTransform) }
    NSGradient(colors: [NSColor(srgbRed: 0.06, green: 0.42, blue: 0.74, alpha: 1),
                        NSColor(srgbRed: 0.07, green: 0.64, blue: 0.71, alpha: 1)])?.draw(in: body, angle: -45)
    NSColor.white.set()
    hole.fill()
    slit.lineWidth = 1.6 * k
    slit.lineCapStyle = .round
    slit.stroke()
}

/// White disc with the nib, sized like Outlook's Copilot button; a spinning arc while busy.
final class RoundButton: NSView {
    static let margin: CGFloat = 4  // room for the drop shadow

    var busy = false { didSet { if busy != oldValue { updateRing() } } }
    var open = false { didSet { if open != oldValue { needsDisplay = true } } }
    var onClick: (() -> Void)?
    var onMenu: ((NSEvent) -> Void)?
    private var hovering = false
    private let ring = CAShapeLayer()

    override init(frame: NSRect) {
        super.init(frame: frame)
        wantsLayer = true
        ring.fillColor = nil
        ring.strokeColor = NSColor(srgbRed: 0.06, green: 0.42, blue: 0.74, alpha: 1).cgColor
        ring.lineWidth = 2
        ring.lineCap = .round
        ring.strokeEnd = 0.3
        ring.isHidden = true
        layer?.addSublayer(ring)
        toolTip = "执笔 · AI 回信（⌃⌥R）"
    }
    required init?(coder: NSCoder) { fatalError() }

    override func layout() {
        super.layout()
        let disc = bounds.insetBy(dx: Self.margin, dy: Self.margin)
        ring.frame = bounds
        ring.path = CGPath(ellipseIn: disc.insetBy(dx: 1, dy: 1), transform: nil)
        updateRing()
    }

    private func updateRing() {
        ring.isHidden = !busy
        ring.removeAllAnimations()
        guard busy else { return }
        let spin = CABasicAnimation(keyPath: "transform.rotation.z")
        spin.fromValue = 0
        spin.toValue = -2 * Double.pi
        spin.duration = 0.9
        spin.repeatCount = .infinity
        ring.add(spin, forKey: "spin")
    }

    override func draw(_ dirtyRect: NSRect) {
        let disc = bounds.insetBy(dx: Self.margin, dy: Self.margin)
        NSGraphicsContext.saveGraphicsState()
        let shadow = NSShadow()
        shadow.shadowBlurRadius = 3
        shadow.shadowOffset = NSSize(width: 0, height: -1)
        shadow.shadowColor = NSColor.black.withAlphaComponent(0.16)
        shadow.set()
        let fill = open ? NSColor(srgbRed: 0.91, green: 0.94, blue: 0.99, alpha: 1) : hovering ? NSColor(white: 0.965, alpha: 1) : .white
        fill.setFill()
        NSBezierPath(ovalIn: disc).fill()
        NSGraphicsContext.restoreGraphicsState()
        let side = (disc.width * 0.62).rounded()
        drawNib(in: NSRect(x: disc.midX - side / 2, y: disc.midY - side / 2, width: side, height: side))
    }

    override func updateTrackingAreas() {
        trackingAreas.forEach(removeTrackingArea)
        addTrackingArea(NSTrackingArea(rect: bounds, options: [.mouseEnteredAndExited, .activeAlways, .inVisibleRect], owner: self))
    }
    override func mouseEntered(with event: NSEvent) { hovering = true; needsDisplay = true }
    override func mouseExited(with event: NSEvent) { hovering = false; needsDisplay = true }
    override func acceptsFirstMouse(for event: NSEvent?) -> Bool { true }
    override func mouseDown(with event: NSEvent) {}
    override func mouseUp(with event: NSEvent) {
        if bounds.contains(convert(event.locationInWindow, from: nil)) { onClick?() }
    }
    override func rightMouseDown(with event: NSEvent) { onMenu?(event) }
}

final class FirstMouseWebView: WKWebView {
    override func acceptsFirstMouse(for event: NSEvent?) -> Bool { true }
}

/// The chat pane: app/pane.html in a WKWebView, talking to Swift through `app` messages.
final class Pane: NSObject, WKScriptMessageHandler, WKNavigationDelegate {
    private(set) var panel = floatingPanel(nonActivating: false)
    let web: FirstMouseWebView
    var onMessage: (([String: Any]) -> Void)?
    private var ready = false
    private var pending: [String] = []

    override init() {
        let cfg = WKWebViewConfiguration()
        web = FirstMouseWebView(frame: .zero, configuration: cfg)
        super.init()
        cfg.userContentController.add(self, name: "app")
        web.setValue(false, forKey: "drawsBackground")
        web.navigationDelegate = self
        panel.contentView = web
        let url = Bundle.main.url(forResource: "pane", withExtension: "html")!
        web.loadFileURL(url, allowingReadAccessTo: url.deletingLastPathComponent())
    }

    /// Calls pane.<fn>(args…) in the page; args are JSON-encoded.
    func call(_ fn: String, _ args: Any...) {
        let data = (try? JSONSerialization.data(withJSONObject: args)) ?? Data("[]".utf8)
        let list = String(decoding: data, as: UTF8.self)
        let js = "pane.\(fn)(...\(list))"
        if ready { web.evaluateJavaScript(js) } else { pending.append(js) }
    }

    func userContentController(_ controller: WKUserContentController, didReceive message: WKScriptMessage) {
        guard let body = message.body as? [String: Any] else { return }
        if body["type"] as? String == "ready" {
            ready = true
            pending.forEach { web.evaluateJavaScript($0) }
            pending = []
        }
        onMessage?(body)
    }

    func webView(_ webView: WKWebView, decidePolicyFor action: WKNavigationAction, decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        // Links inside a draft open in the browser; the pane itself never navigates away.
        if action.navigationType == .linkActivated, let url = action.request.url {
            if ["http", "https", "mailto"].contains(url.scheme ?? "") { NSWorkspace.shared.open(url) }
            return decisionHandler(.cancel)
        }
        decisionHandler(action.request.url?.isFileURL == true ? .allow : .cancel)
    }

    /// Same page in a new window, which lands in the active Space.
    func rehome() {
        let p = floatingPanel(nonActivating: false)
        panel.close()
        panel.contentView = NSView()
        p.contentView = web
        panel = p
    }

    func focus() {
        NSApp.activate()
        panel.makeKeyAndOrderFront(nil)
        panel.makeFirstResponder(web)
        call("focus")
    }
}

// MARK: - app

struct Draft {
    let base: String
    let html: String
    let md: String
    let mailKey: String
    let mode: String
    let request: [String]
    var selected = false
}

final class App: NSObject, NSApplicationDelegate {
    let tracker = Tracker()
    var buttonPanel = floatingPanel(nonActivating: true)
    let button = RoundButton(frame: NSRect(x: 0, y: 0, width: 48, height: 48))
    let pane = Pane()
    var paneOpen = false
    var job: Process?
    var drafts: [Int: Draft] = [:]
    var lastDraft: [String: Int] = [:]
    var chatMailKey = ""
    var chatEmpty = true
    var nextID = 1
    var shownSubject: String?
    var nextMailCheck = Date.distantPast
    var hotKey: EventHotKeyRef?
    var trusted = false
    var visible = false
    var offSpaceSince: Date?
    var homeKey = ""

    func applicationDidFinishLaunching(_ note: Notification) {
        button.autoresizingMask = [.width, .height]
        buttonPanel.contentView = button
        button.onClick = { [weak self] in self?.togglePane() }
        button.onMenu = { [weak self] e in self?.showMenu(e) }
        pane.onMessage = { [weak self] m in self?.handle(m) }

        trusted = AXIsProcessTrustedWithOptions([kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true] as CFDictionary)
        log("start; accessibility=\(trusted)")
        installHotKey()
        NSWorkspace.shared.notificationCenter.addObserver(self, selector: #selector(appActivated(_:)),
                                                          name: NSWorkspace.didActivateApplicationNotification, object: nil)
        NSWorkspace.shared.notificationCenter.addObserver(self, selector: #selector(spaceChanged(_:)),
                                                          name: NSWorkspace.activeSpaceDidChangeNotification, object: nil)
        Timer.scheduledTimer(withTimeInterval: 0.15, repeats: true) { [weak self] _ in self?.tick() }
        syncLoginItem()
    }

    /// outlookai://reply opens the pane; outlookai://reply?points=…[&to=sender] also sends right away
    /// (for Raycast / Shortcuts; use `open -g` so Outlook stays in front).
    func application(_ application: NSApplication, open urls: [URL]) {
        for url in urls where url.scheme == "outlookai" {
            log("url \(url.absoluteString)")
            let query = URLComponents(url: url, resolvingAgainstBaseURL: false)?.queryItems ?? []
            guard let points = query.first(where: { $0.name == "points" })?.value else { openPane(); continue }
            paneOpen = true
            send(points, all: query.first { $0.name == "to" }?.value != "sender")
        }
    }

    // MARK: tracking

    func tick() {
        // Activation notifications can arrive before frontmostApplication changes, so re-check every tick.
        updateHotKey()
        if !trusted {
            trusted = AXIsProcessTrusted()
            if !trusted { return }
            log("accessibility granted")
        }
        // Our own pane being active counts as "in Outlook" too.
        let inOutlook = NSWorkspace.shared.frontmostApplication?.bundleIdentifier == outlookID || NSApp.isActive
        let place = inOutlook ? tracker.placement() : nil
        guard let p = place else {
            offSpaceSince = nil
            if visible && (inOutlook || !parked()) {
                buttonPanel.orderOut(nil)
                pane.panel.orderOut(nil)
                visible = false
            }
            return
        }
        // Outlook is in front but its Space is not showing yet; a panel ordered in now would land in the old Space.
        guard tracker.mainShowing() ?? false else { return }
        if [buttonPanel, pane.panel].contains(where: { $0.isVisible && !$0.isOnActiveSpace }) {
            // Outlook moved Space (full screen on/off), or a Space switch is still settling: wait that out first.
            let since = offSpaceSince ?? Date()
            offSpaceSince = since
            if Date().timeIntervalSince(since) < 0.6 { return }
            rehome()
        } else if tracker.spaceKey != homeKey && !homeKey.isEmpty {
            // Hidden panels would come back in Outlook's old Space.
            rehome()
        }
        homeKey = tracker.spaceKey
        offSpaceSince = nil
        visible = true
        let r = cocoaRect(p.insetBy(dx: -RoundButton.margin, dy: -RoundButton.margin))
        if buttonPanel.frame != r { buttonPanel.setFrame(r, display: true) }
        if !buttonPanel.isVisible { buttonPanel.orderFrontRegardless() }
        button.open = paneOpen
        guard paneOpen, let pf = tracker.paneFrame() else {
            if pane.panel.isVisible { pane.panel.orderOut(nil) }
            return
        }
        let f = cocoaRect(pf)
        if pane.panel.frame != f { pane.panel.setFrame(f, display: true) }
        if !pane.panel.isVisible { pane.panel.orderFrontRegardless() }
        if Date() >= nextMailCheck {
            nextMailCheck = Date().addingTimeInterval(0.5)
            let subject = tracker.readHeader()?.subject ?? ""
            if subject != shownSubject { shownSubject = subject; pane.call("setMail", subject) }
        }
    }

    /// With another app in front, the panels can stay where they are when that is not Outlook's Space (they slide
    /// away and back with it) or when Outlook is full screen (that app is only an overlay or split-view neighbour).
    func parked() -> Bool {
        guard let app = tracker.outlook, !app.isHidden, let showing = tracker.mainShowing() else { return false }
        return buttonPanel.isOnActiveSpace ? tracker.fullScreen && showing : true
    }

    /// A window keeps its Space across orderOut/orderFront, so moving the panels means new windows.
    func rehome() {
        let b = floatingPanel(nonActivating: true)
        buttonPanel.close()
        buttonPanel.contentView = NSView()
        b.contentView = button
        buttonPanel = b
        pane.rehome()
        log("panels moved to Outlook's Space")
    }

    // MARK: pane

    func togglePane() {
        if paneOpen && pane.panel.isKeyWindow && NSApp.isActive { return closePane() }
        openPane()
    }

    func openPane() {
        paneOpen = true
        button.open = true
        nextMailCheck = .distantPast
        tick()
        pane.focus()
    }

    func closePane() {
        paneOpen = false
        button.open = false
        pane.panel.orderOut(nil)
        if NSApp.isActive { returnToOutlook() }
    }

    func returnToOutlook() {
        guard let app = tracker.outlook else { return }
        NSApp.yieldActivation(to: app)
        app.activate()
    }

    func handle(_ m: [String: Any]) {
        let id = m["id"] as? Int ?? 0
        switch m["type"] as? String ?? "" {
        case "ready": pane.call("setMail", shownSubject ?? "")
        case "send": send(m["text"] as? String ?? "", all: m["all"] as? Bool ?? true)
        case "refine":
            guard let d = drafts[id], job == nil else { return }
            let text = m["text"] as? String ?? ""
            pane.call("user", text)
            generate(["generate", "--revise", d.base, "--points", text] + (d.mode == "reply" ? ["--reply"] : []), mailKey: d.mailKey)
        case "regen":
            guard let d = drafts[id], job == nil else { return }
            pane.call("user", "重新生成")
            generate(d.request, mailKey: d.mailKey)
        case "open": openReply(id)
        case "copy": copy(id)
        case "new":
            pane.call("clear")
            lastDraft = [:]
            chatMailKey = ""
            chatEmpty = true
            pane.call("focus")
        case "close": closePane()
        case "openSettings":
            run(["settings"] + (m["refresh"] as? Bool == true ? ["--refresh"] : []), track: false) { [weak self] ok, res in
                if ok, let d = res["data"] { self?.pane.call("showSettings", d) }
                else { self?.pane.call("settingsSaved", ["ok": false, "message": (res["error"] as? [String: Any])?["message"] ?? "读取设置失败"]) }
            }
        case "saveSettings":
            let values = m["values"] as? [String: Any] ?? [:]
            let input = (try? JSONSerialization.data(withJSONObject: values)) ?? Data()
            let ping = m["ping"] as? Bool == true
            run(["save-settings"], stdin: input, track: false) { [weak self] ok, res in
                guard let self else { return }
                var d = res["data"] as? [String: Any] ?? [:]
                d["ok"] = ok
                if !ok { d["message"] = (res["error"] as? [String: Any])?["message"] ?? "未知错误" }
                self.pane.call("settingsSaved", d)
                guard ok, ping else { return }
                self.run(["ping"], track: false) { ok, res in
                    var r = res["data"] as? [String: Any] ?? [:]
                    r["ok"] = ok
                    if !ok { r["message"] = (res["error"] as? [String: Any])?["message"] ?? "未知错误" }
                    self.pane.call("pingResult", r)
                }
            }
        default: break
        }
    }

    func send(_ text: String, all: Bool) {
        guard job == nil else { return }
        guard let h = tracker.readHeader() else { return NSSound.beep() }
        if h.key != chatMailKey {
            if !chatEmpty { pane.call("divider", "回复：\(h.subject)") }
            chatMailKey = h.key
        }
        chatEmpty = false
        pane.call("user", text.isEmpty ? "按原文拟一版" : text)
        let modeArgs = all ? [] : ["--reply"]
        // A follow-up on a mail that already has a draft rewrites that draft.
        if !text.isEmpty, let prev = lastDraft[h.key], let d = drafts[prev] {
            generate(["generate", "--revise", d.base, "--points", text] + modeArgs, mailKey: h.key)
        } else {
            generate(["generate", "--subject", h.subject, "--from", h.from, "--folder", h.folder, "--points", text] + modeArgs,
                     mailKey: h.key)
        }
    }

    func generate(_ args: [String], mailKey: String) {
        let id = nextID
        nextID += 1
        pane.call("start", id, args.contains("--revise") ? "改写中…" : "读取中…")
        setBusy(true)
        run(args, onStage: { [weak self] s in self?.pane.call("status", id, s + "…") }) { [weak self] ok, res in
            guard let self else { return }
            self.setBusy(false)
            let data = res["data"] as? [String: Any] ?? [:]
            guard ok, let base = data["base"] as? String,
                  let html = try? String(contentsOfFile: base + ".html", encoding: .utf8) else {
                let err = res["error"] as? [String: Any] ?? [:]
                let code = err["code"] as? String ?? ""
                let message: String = {
                    switch code {
                    case "E_CANCELLED": return "已停止。"
                    case "E_NOT_FOUND": return "在邮箱里没找到这封邮件（按主题和发件人查找）。"
                    case "E_AI": return "AI 生成失败：\(err["message"] as? String ?? "")"
                    default: return "生成失败：\(err["message"] as? String ?? "未知错误")"
                    }
                }()
                self.drafts[id] = Draft(base: "", html: "", md: "", mailKey: mailKey, mode: "", request: args)
                return self.pane.call("error", id, message, code != "E_CANCELLED")
            }
            let md = (try? String(contentsOfFile: base + ".md", encoding: .utf8)) ?? ""
            let mode = data["mode"] as? String ?? "reply-all"
            self.drafts[id] = Draft(base: base, html: html, md: md, mailKey: mailKey, mode: mode, request: args)
            self.lastDraft[mailKey] = id
            self.pane.call("draft", id, [
                "html": html,
                "placeholders": data["placeholders"] as? [String] ?? [],
                "note": self.recipientsNote(base: base, mode: mode),
                "notesHtml": data["notes_html"] as? String ?? "",
                "openLabel": mode == "reply-all" ? "打开全部答复窗口" : "打开答复窗口",
            ])
        }
    }

    func recipientsNote(base: String, mode: String) -> String {
        guard let d = try? Data(contentsOf: URL(fileURLWithPath: base + ".info.json")),
              let j = (try? JSONSerialization.jsonObject(with: d)) as? [String: Any],
              let info = j["data"] as? [String: Any] else { return "" }
        let people = ((info["to"] as? [[String: String]] ?? []) + (info["cc"] as? [[String: String]] ?? []))
            .map { ($0["name"]?.isEmpty == false ? $0["name"]! : $0["email"] ?? "") }
        let who = people.count > 2 ? "\(people.prefix(2).joined(separator: "、")) 等 \(people.count) 人" : people.joined(separator: "、")
        var note = "\(mode == "reply-all" ? "全部答复" : "答复") · 发给 \(who)"
        let files = (info["original"] as? [String: Any])?["attachment_files"] as? [[String: Any]] ?? []
        let read = files.filter { !($0["text"] as? String ?? "").isEmpty }.count
        if !files.isEmpty { note += " · 已读附件 \(read)/\(files.count)" }
        return note
    }

    func openReply(_ id: Int) {
        guard job == nil, let d = drafts[id], !d.base.isEmpty else { return }
        pane.call("opened", id, "正在打开…", true)
        setBusy(true)
        // Bring Outlook back first so the helper needs no Apple Event (and no Automation prompt).
        returnToOutlook()
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.4) {
            self.run(["compose", "--base", d.base] + (d.selected ? ["--selected"] : [])) { [weak self] ok, res in
                guard let self else { return }
                self.setBusy(false)
                if ok { return self.pane.call("opened", id, "已打开回复窗口", true) }
                let err = res["error"] as? [String: Any] ?? [:]
                switch err["code"] as? String ?? "" {
                case "E_NOT_FOUND":
                    self.drafts[id]?.selected = true
                    self.pane.call("opened", id, "当前列表里没找到原邮件，在 Outlook 里选中它后再点一次", false)
                case "E_CONFLICT":
                    self.pane.call("opened", id, "这封邮件已有打开的回复窗口，先关掉或放弃它", false)
                default:
                    self.pane.call("opened", id, "打开失败：\(err["message"] as? String ?? "未知错误")", false)
                }
            }
        }
    }

    func copy(_ id: Int) {
        guard let d = drafts[id], !d.base.isEmpty else { return }
        let pb = NSPasteboard.general
        pb.clearContents()
        pb.setString(d.html, forType: .html)
        pb.setString(d.md, forType: .string)
    }

    func setBusy(_ b: Bool) {
        button.busy = b
        pane.call("setBusy", b)
    }

    /// Runs ai_reply.py; "@stage" lines on stderr report progress, stdout is one JSON object.
    func run(_ args: [String], stdin: Data? = nil, track: Bool = true, onStage: @escaping (String) -> Void = { _ in },
             done: @escaping (Bool, [String: Any]) -> Void) {
        guard let py = pythonPath() else {
            return done(false, ["error": ["code": "E_CONFIG", "message": "找不到 Python 版 outlook-cli，安装方法见 https://github.com/ElvisLabs/zhibi"]])
        }
        let p = Process()
        p.executableURL = URL(fileURLWithPath: py)
        p.arguments = [pipeline] + args
        var env = ProcessInfo.processInfo.environment
        env["PATH"] = NSHomeDirectory() + "/.local/bin:/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin:/usr/sbin:/sbin"
        env["PYTHONUTF8"] = "1"
        env["PYTHONDONTWRITEBYTECODE"] = "1"   // scripts live in the signed bundle
        p.environment = env
        let out = Pipe(), err = Pipe()
        p.standardOutput = out
        p.standardError = err
        let inPipe = Pipe()
        p.standardInput = stdin == nil ? FileHandle.nullDevice : inPipe
        var errText = ""
        var partial = ""
        // Drain stdout while the child runs: `settings` returns signatures with embedded images, far
        // more than a pipe buffer, and a child blocked on a full pipe never exits.
        var outData = Data()
        let drained = DispatchGroup()
        drained.enter()
        err.fileHandleForReading.readabilityHandler = { h in
            let data = h.availableData
            if data.isEmpty { h.readabilityHandler = nil; return }
            let chunk = String(decoding: data, as: UTF8.self)
            DispatchQueue.main.async {
                errText += chunk
                partial += chunk
                while let nl = partial.firstIndex(of: "\n") {
                    let line = String(partial[..<nl])
                    partial = String(partial[partial.index(after: nl)...])
                    if line.hasPrefix("@stage ") { onStage(String(line.dropFirst(7))) }
                }
            }
        }
        p.terminationHandler = { [weak self] proc in
            err.fileHandleForReading.readabilityHandler = nil
            drained.notify(queue: .main) {
                if track { self?.job = nil }
                let text = String(decoding: outData, as: UTF8.self)
                let json = (try? JSONSerialization.jsonObject(with: Data(text.utf8))) as? [String: Any]
                    ?? ["ok": false, "error": ["code": "E_INTERNAL", "message": String(errText.suffix(400))]]
                let ok = proc.terminationStatus == 0 && (json["ok"] as? Bool ?? false)
                log("\(args.first ?? "") exit=\(proc.terminationStatus) \(text.trimmingCharacters(in: .whitespacesAndNewlines).prefix(800))")
                if !ok && !errText.isEmpty { log("stderr: \(errText.suffix(2000))") }
                done(ok, json)
            }
        }
        log("run \(args.joined(separator: " | "))")
        do {
            try p.run()
            DispatchQueue.global().async {
                outData = out.fileHandleForReading.readDataToEndOfFile()
                drained.leave()
            }
            if track { job = p }
            if let data = stdin {
                inPipe.fileHandleForWriting.write(data)
                try? inPipe.fileHandleForWriting.close()
            }
        } catch {
            drained.leave()
            done(false, ["error": ["code": "E_INTERNAL", "message": error.localizedDescription]])
        }
    }

    // MARK: menu

    func showMenu(_ event: NSEvent) {
        let menu = NSMenu()
        menu.addItem(item(paneOpen ? "关闭执笔面板" : "打开执笔面板（⌃⌥R）") { [weak self] in self?.togglePane() })
        if job != nil { menu.addItem(item("停止生成") { [weak self] in self?.job?.terminate() }) }
        menu.addItem(.separator())
        menu.addItem(item("打开输出目录") { NSWorkspace.shared.open(URL(fileURLWithPath: expand(config.outDir))) })
        menu.addItem(item("查看日志") { NSWorkspace.shared.open(URL(fileURLWithPath: expand(config.logFile))) })
        let login = item("登录时自动启动") { [weak self] in self?.toggleLoginItem() }
        login.state = SMAppService.mainApp.status == .enabled ? .on : .off
        menu.addItem(login)
        menu.addItem(.separator())
        menu.addItem(item("退出执笔") { NSApp.terminate(nil) })
        NSMenu.popUpContextMenu(menu, with: event, for: button)
    }

    private var menuActions: [() -> Void] = []
    func item(_ title: String, _ action: @escaping () -> Void) -> NSMenuItem {
        menuActions.append(action)
        let it = NSMenuItem(title: title, action: #selector(menuTap(_:)), keyEquivalent: "")
        it.target = self
        it.tag = menuActions.count - 1
        return it
    }
    @objc func menuTap(_ sender: NSMenuItem) { menuActions[sender.tag]() }

    // MARK: login item

    /// Turns on start-at-login on the first launch. After the app moved (e.g. a Homebrew install) it re-registers
    /// so login opens this copy, but keeps it off if the user turned it off.
    func syncLoginItem() {
        let d = UserDefaults.standard, path = Bundle.main.bundlePath, svc = SMAppService.mainApp
        let known = d.string(forKey: "loginItemPath")
        guard known != path else { return }
        d.set(path, forKey: "loginItemPath")
        let wasOn = svc.status == .enabled
        if wasOn { try? svc.unregister() }
        if wasOn || known == nil { try? svc.register() }
        log("login item: \(svc.status.rawValue) \(path)")
    }

    func toggleLoginItem() {
        if SMAppService.mainApp.status == .enabled {
            try? SMAppService.mainApp.unregister()
        } else {
            try? SMAppService.mainApp.register()
        }
    }

    // MARK: hot key ⌃⌥R, registered only while Outlook or the pane is in front

    func installHotKey() {
        var spec = EventTypeSpec(eventClass: OSType(kEventClassKeyboard), eventKind: UInt32(kEventHotKeyPressed))
        InstallEventHandler(GetApplicationEventTarget(), { _, _, _ in
            DispatchQueue.main.async { (NSApp.delegate as? App)?.hotKeyPressed() }
            return noErr
        }, 1, &spec, nil, nil)
    }

    @objc func appActivated(_ note: Notification) { tick() }

    @objc func spaceChanged(_ note: Notification) { tick() }

    func updateHotKey() {
        let front = NSWorkspace.shared.frontmostApplication?.bundleIdentifier
        let want = front == outlookID || front == Bundle.main.bundleIdentifier
        if want && hotKey == nil {
            let id = EventHotKeyID(signature: OSType(0x4F41_4952), id: 1)
            let status = RegisterEventHotKey(UInt32(kVK_ANSI_R), UInt32(controlKey | optionKey), id, GetApplicationEventTarget(), 0, &hotKey)
            if status != noErr { log("hot key ⌃⌥R not registered: \(status)") }
        } else if !want, let k = hotKey {
            UnregisterEventHotKey(k)
            hotKey = nil
        }
    }

    func hotKeyPressed() {
        if visible { togglePane() } else { NSSound.beep() }
    }
}

let app = NSApplication.shared
let delegate = App()
app.delegate = delegate
app.setActivationPolicy(.accessory)
app.run()
