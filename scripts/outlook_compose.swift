// Drive new Outlook for Mac through the Accessibility API: open the reply window of a
// given mail and paste a formatted HTML body at the top. Never presses Send.
//
//   outlook_compose reply --subject "<original subject>" --sender <email> [--name <display name>]
//                         [--folder <folder name>] --html body.html [--text body.md] [--all]
//                         [--selected] [--timeout 20]
//
// Prints one JSON object; exit 0 on success, 2 bad args, 3 mail not found, 4 permission, 5 UI failure.
import ApplicationServices
import Cocoa

let bundleID = "com.microsoft.Outlook"

// MARK: - output

var steps: [String] = []
func step(_ s: String) { steps.append(s) }

func emit(_ obj: [String: Any], _ code: Int32) -> Never {
    let data = try! JSONSerialization.data(withJSONObject: obj, options: [.sortedKeys])
    print(String(data: data, encoding: .utf8)!)
    exit(code)
}

func fail(_ code: String, _ message: String, _ exitCode: Int32, _ extra: [String: Any] = [:]) -> Never {
    var err: [String: Any] = ["code": code, "message": message]
    extra.forEach { err[$0] = $1 }
    emit(["ok": false, "error": err, "steps": steps], exitCode)
}

// MARK: - AX helpers

func attr(_ e: AXUIElement, _ n: String) -> AnyObject? {
    var v: AnyObject?
    return AXUIElementCopyAttributeValue(e, n as CFString, &v) == .success ? v : nil
}
func s(_ e: AXUIElement, _ n: String) -> String { (attr(e, n) as? String) ?? "" }
func role(_ e: AXUIElement) -> String { s(e, kAXRoleAttribute) }
func ident(_ e: AXUIElement) -> String { s(e, "AXIdentifier") }
func kids(_ e: AXUIElement) -> [AXUIElement] { (attr(e, kAXChildrenAttribute) as? [AXUIElement]) ?? [] }
func press(_ e: AXUIElement) -> Bool { AXUIElementPerformAction(e, kAXPressAction as CFString) == .success }

func find(_ root: AXUIElement, depth: Int = 14, _ pred: (AXUIElement) -> Bool) -> AXUIElement? {
    var queue: [(AXUIElement, Int)] = [(root, 0)]
    while !queue.isEmpty {
        let (e, d) = queue.removeFirst()
        if pred(e) { return e }
        // Mail bodies are huge web trees; never descend into them while searching chrome.
        if d < depth && role(e) != "AXWebArea" { queue += kids(e).map { ($0, d + 1) } }
    }
    return nil
}

func wait<T>(_ timeout: Double, _ poll: Double = 0.2, _ body: () -> T?) -> T? {
    let end = Date().addingTimeInterval(timeout)
    repeat {
        if let v = body() { return v }
        RunLoop.current.run(until: Date().addingTimeInterval(poll))
    } while Date() < end
    return nil
}

func norm(_ x: String) -> String {
    x.replacingOccurrences(of: "\u{00a0}", with: " ")
        .components(separatedBy: .whitespacesAndNewlines).filter { !$0.isEmpty }.joined(separator: " ")
}

let prefixRE = try! NSRegularExpression(pattern: "^\\s*((re|fw|fwd|回复|答复|转发)\\s*[:：]\\s*)+", options: [.caseInsensitive])
func core(_ subject: String) -> String {
    let n = norm(subject)
    return prefixRE.stringByReplacingMatches(in: n, range: NSRange(n.startIndex..., in: n), withTemplate: "")
}

// MARK: - args

var opts: [String: String] = [:]
var flags = Set<String>()
var args = Array(CommandLine.arguments.dropFirst())
guard let command = args.first else { fail("E_USAGE", "usage: outlook_compose reply --subject S --sender EMAIL --html FILE [--all]", 2) }
args.removeFirst()
var i = 0
while i < args.count {
    let a = args[i]
    if a == "--all" || a == "--selected" { flags.insert(String(a.dropFirst(2))); i += 1; continue }
    guard a.hasPrefix("--"), i + 1 < args.count else { fail("E_USAGE", "bad argument \(a)", 2) }
    opts[String(a.dropFirst(2))] = args[i + 1]
    i += 2
}
guard command == "reply" else { fail("E_USAGE", "unknown command \(command)", 2) }
guard let subject = opts["subject"], let htmlPath = opts["html"] else { fail("E_USAGE", "--subject and --html are required", 2) }
let senderEmail = (opts["sender"] ?? "").lowercased()
let senderName = opts["name"] ?? ""
let timeout = Double(opts["timeout"] ?? "20") ?? 20
guard let html = try? String(contentsOfFile: htmlPath, encoding: .utf8) else { fail("E_USAGE", "cannot read \(htmlPath)", 2) }
let plain = opts["text"].flatMap { try? String(contentsOfFile: $0, encoding: .utf8) } ?? ""

guard AXIsProcessTrusted() else {
    fail("E_PERMISSION", "terminal app lacks Accessibility permission (System Settings > Privacy & Security > Accessibility)", 4)
}

// MARK: - Outlook app

func outlookApp() -> NSRunningApplication? { NSRunningApplication.runningApplications(withBundleIdentifier: bundleID).first }

if outlookApp() == nil {
    step("launch Outlook")
    NSWorkspace.shared.open(URL(fileURLWithPath: "/Applications/Microsoft Outlook.app"))
}
guard let app = wait(30, 0.5, { outlookApp() }) else { fail("E_UI", "Outlook did not start", 5) }
let ax = AXUIElementCreateApplication(app.processIdentifier)

func activate() -> Bool {
    let ready = { NSWorkspace.shared.frontmostApplication?.bundleIdentifier == bundleID && attr(ax, kAXWindowsAttribute) != nil }
    // Skipping the Apple Event when Outlook is already in front avoids an Automation prompt for the button app.
    if ready() { return true }
    // NSRunningApplication.activate does not switch to a full-screen Space; an Apple Event does.
    var err: NSDictionary?
    NSAppleScript(source: "tell application id \"\(bundleID)\" to activate")?.executeAndReturnError(&err)
    if err != nil { app.activate(options: [.activateAllWindows]) }
    return wait(5) { ready() ? true : nil } ?? false
}

guard activate() else { fail("E_UI", "could not bring Outlook to front", 5) }

func windows() -> [AXUIElement] {
    var list = (attr(ax, kAXWindowsAttribute) as? [AXUIElement]) ?? []
    if list.isEmpty, let w = attr(ax, kAXMainWindowAttribute) { list = [w as! AXUIElement] }
    return list
}
func searchBar(_ w: AXUIElement) -> AXUIElement? { find(w, depth: 3) { ident($0) == "Search Bar" } }
func mainWindow() -> AXUIElement? { windows().first { searchBar($0) != nil } }

// A full-screen Outlook only shows its windows to AX once its Space is active.
var main = wait(5, 0.3) { mainWindow() }
if main == nil {
    step("reopen main window")
    NSWorkspace.shared.open(URL(fileURLWithPath: "/Applications/Microsoft Outlook.app"))
    main = wait(10, 0.5) { mainWindow() }
}
guard let mainWin = main else { fail("E_UI", "Outlook main window not found", 5) }

// MARK: - locate the mail

func header() -> (subject: String, from: String)? {
    guard let h = find(mainWin, depth: 8, { ident($0) == "HeaderMainContainer" }) else { return nil }
    let subj = kids(h).first { role($0) == "AXStaticText" }.map { s($0, kAXValueAttribute) } ?? ""
    let from = find(h, depth: 4) { ident($0) == "messageHeaderFromContent" }.map { s($0, kAXValueAttribute) } ?? ""
    return (subj, from)
}

func headerMatches() -> Bool {
    guard let h = header() else { return false }
    let subjOK = norm(h.subject) == norm(subject)
    let fromOK = senderEmail.isEmpty || h.from.lowercased().contains(senderEmail)
    return subjOK && fromOK
}

func messageTable() -> AXUIElement? {
    find(mainWin, depth: 8) { role($0) == "AXTable" && s($0, kAXDescriptionAttribute) == "邮件列表" }
}

func rowText(_ row: AXUIElement) -> String { norm(kids(row).map { s($0, kAXDescriptionAttribute) }.joined(separator: " ")) }

func rowMatches(_ text: String) -> Bool {
    guard text.contains(norm(subject)) || text.contains(core(subject)) else { return false }
    return senderName.isEmpty || text.contains(norm(senderName))
}

func selectFromVisibleRows() -> Bool {
    guard let table = messageTable() else { return false }
    let rows = (attr(table, "AXVisibleRows") as? [AXUIElement]) ?? []
    for row in rows where rowMatches(rowText(row)) {
        AXUIElementSetAttributeValue(table, kAXSelectedRowsAttribute as CFString, [row] as CFArray)
        AXUIElementSetAttributeValue(row, kAXSelectedAttribute as CFString, kCFBooleanTrue)
        if wait(3, 0.2, { headerMatches() ? true : nil }) != nil { return true }
    }
    return false
}

func searching() -> Bool {
    let t = s(mainWin, kAXTitleAttribute)
    return t.hasPrefix("正在搜索") || t.hasPrefix("Searching")
}

func clearSearch() {
    // Use Outlook's own cancel button; synthetic keys could land in Copilot or another app.
    guard searching(), let cancel = find(mainWin, depth: 8, { ident($0) == "Cancel Search Button" }) else { return }
    _ = press(cancel)
    _ = wait(3, 0.2) { searching() ? nil : true }
}

func selectFolder(_ name: String) -> Bool {
    guard let outline = find(mainWin, depth: 5, { role($0) == "AXOutline" }) else { return false }
    for row in (attr(outline, kAXRowsAttribute) as? [AXUIElement]) ?? [] {
        let label = kids(row).map { s($0, kAXDescriptionAttribute) }.first ?? ""
        guard label.components(separatedBy: ";").first?.trimmingCharacters(in: .whitespaces) == name else { continue }
        AXUIElementSetAttributeValue(outline, kAXSelectedRowsAttribute as CFString, [row] as CFArray)
        return wait(5, 0.2) { s(mainWin, kAXTitleAttribute).hasPrefix(name) ? true : nil } ?? false
    }
    return false
}

if flags.contains("selected") {
    // User picked the mail by hand; only make sure it is the same thread.
    guard let h = header(), norm(h.subject).contains(core(subject)) else {
        fail("E_NOT_FOUND", "selected mail is not in this thread", 3, ["header_subject": header()?.subject ?? ""])
    }
    step("using selected mail: \(h.subject)")
} else if headerMatches() {
    step("target mail already selected")
} else {
    clearSearch()
    if let folder = opts["folder"], !folder.isEmpty {
        step(selectFolder(folder) ? "opened folder \(folder)" : "folder \(folder) not in sidebar")
    }
    guard selectFromVisibleRows() else {
        fail("E_NOT_FOUND", "mail not in the visible message list; select it in Outlook and rerun with --selected", 3)
    }
    step("selected target in message list")
}

// MARK: - open the reply window

func isCompose(_ w: AXUIElement) -> Bool {
    find(w, depth: 3) { role($0) == "AXButton" && s($0, kAXDescriptionAttribute) == "发送" } != nil
}

let before = Set(windows().map { s($0, kAXTitleAttribute) })
if let open = windows().first(where: { norm(s($0, kAXTitleAttribute)).contains(core(subject)) && isCompose($0) }) {
    // Never paste into a reply the user may already be writing.
    AXUIElementPerformAction(open, kAXRaiseAction as CFString)
    fail("E_CONFLICT", "a reply window for this mail is already open; close or discard it first", 6, ["window": s(open, kAXTitleAttribute)])
}
let buttonID = flags.contains("all") ? "messageHeaderContextualControlsReplyAllButton" : "messageHeaderContextualControlsReplyButton"
guard let replyButton = find(mainWin, depth: 8, { ident($0) == buttonID }) else { fail("E_UI", "reply button not found", 5) }
_ = activate()
_ = press(replyButton)
step(flags.contains("all") ? "pressed 全部答复" : "pressed 答复")

guard let compose = wait(timeout, 0.3, { () -> AXUIElement? in
    windows().first { w in
        let t = s(w, kAXTitleAttribute)
        return !before.contains(t) && norm(t).contains(core(subject)) && isCompose(w)
    }
}) else { fail("E_UI", "reply window did not open", 5) }
let composeTitle = s(compose, kAXTitleAttribute)
step("reply window: \(composeTitle)")

// MARK: - paste the body

// The editor is the AXTextArea directly under the compose window's AXWebArea.
guard let body = wait(10, 0.3, { () -> AXUIElement? in
    guard let web = find(compose, depth: 8, { role($0) == "AXWebArea" }) else { return nil }
    return kids(web).first { role($0) == "AXTextArea" }
}) else { fail("E_UI", "message body not found", 5) }
AXUIElementPerformAction(compose, kAXRaiseAction as CFString)
AXUIElementSetAttributeValue(body, kAXFocusedAttribute as CFString, kCFBooleanTrue)
var top = CFRange(location: 0, length: 0)
if let range = AXValueCreate(.cfRange, &top) { AXUIElementSetAttributeValue(body, kAXSelectedTextRangeAttribute as CFString, range) }

let pb = NSPasteboard.general
let saved: [[NSPasteboard.PasteboardType: Data]] = (pb.pasteboardItems ?? []).map { item in
    var d: [NSPasteboard.PasteboardType: Data] = [:]
    for t in item.types { if let data = item.data(forType: t) { d[t] = data } }
    return d
}
func restoreClipboard() {
    pb.clearContents()
    pb.writeObjects(saved.map { dict -> NSPasteboardItem in
        let it = NSPasteboardItem()
        dict.forEach { it.setData($1, forType: $0) }
        return it
    })
}
pb.clearContents()
pb.setString(html, forType: .html)
pb.setString(plain, forType: .string)

func pasteMenuItem() -> AXUIElement? {
    guard let bar = attr(ax, kAXMenuBarAttribute) else { return nil }
    for top in kids(bar as! AXUIElement) where s(top, kAXTitleAttribute) == "编辑" || s(top, kAXTitleAttribute) == "Edit" {
        for menu in kids(top) {
            if let item = kids(menu).first(where: { ["粘贴", "Paste"].contains(s($0, kAXTitleAttribute)) }) { return item }
        }
    }
    return nil
}

let beforeLen = s(body, kAXValueAttribute).count
guard activate(), let pasteItem = pasteMenuItem(), (attr(pasteItem, kAXEnabledAttribute) as? Bool) == true, press(pasteItem) else {
    restoreClipboard()
    fail("E_UI", "paste failed; the reply window is open, paste manually", 5, ["window": composeTitle])
}
let pasted = wait(5, 0.2) { s(body, kAXValueAttribute).count > beforeLen + 10 ? true : nil } != nil
RunLoop.current.run(until: Date().addingTimeInterval(0.5))
restoreClipboard()
step(pasted ? "pasted formatted body" : "paste not confirmed")

emit([
    "ok": pasted,
    "data": ["window": composeTitle, "mode": flags.contains("all") ? "reply-all" : "reply"],
    "steps": steps,
], pasted ? 0 : 5)
