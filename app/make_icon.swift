// Render AppIcon.iconset (blue rounded square + white pen nib) for iconutil.
import Cocoa

let out = CommandLine.arguments[1]
try? FileManager.default.createDirectory(atPath: out, withIntermediateDirectories: true)
for (px, name) in [(16, "16x16"), (32, "16x16@2x"), (32, "32x32"), (64, "32x32@2x"), (128, "128x128"),
                   (256, "128x128@2x"), (256, "256x256"), (512, "256x256@2x"), (512, "512x512"), (1024, "512x512@2x")] {
    let s = CGFloat(px)
    let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: px, pixelsHigh: px, bitsPerSample: 8, samplesPerPixel: 4,
                               hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
    NSGraphicsContext.saveGraphicsState()
    NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
    let inset = s * 0.1
    let rect = NSRect(x: inset, y: inset, width: s - 2 * inset, height: s - 2 * inset)
    NSGradient(starting: NSColor(srgbRed: 0.07, green: 0.64, blue: 0.71, alpha: 1),
               ending: NSColor(srgbRed: 0.06, green: 0.42, blue: 0.74, alpha: 1))!
        .draw(in: NSBezierPath(roundedRect: rect, xRadius: rect.width * 0.22, yRadius: rect.width * 0.22), angle: -90)
    // White nib on the blue tile, same shape as drawNib() in Zhibi.swift.
    let k = rect.width * 0.62 / 32
    let t = NSAffineTransform()
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
    NSColor.white.setFill()
    body.fill()
    // Hole and slit in the tile's mid-gradient colour.
    let tile = NSColor(srgbRed: 0.065, green: 0.53, blue: 0.725, alpha: 1)
    tile.set()
    hole.fill()
    slit.lineWidth = 1.6 * k
    slit.lineCapStyle = .round
    slit.stroke()
    NSGraphicsContext.restoreGraphicsState()
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: "\(out)/icon_\(name).png"))
}
