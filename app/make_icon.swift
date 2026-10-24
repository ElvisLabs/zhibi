// Render AppIcon.iconset for iconutil: a fountain pen writing on an envelope, on a blue macOS-style tile.
// Drawn once at 1024 px (Core Graphics shadows ignore the CTM) and scaled down for the smaller sizes.
import Cocoa

let cs = CGColorSpace(name: CGColorSpace.sRGB)!

func rgb(_ hex: UInt32, _ alpha: CGFloat = 1) -> CGColor {
    CGColor(srgbRed: CGFloat((hex >> 16) & 255) / 255, green: CGFloat((hex >> 8) & 255) / 255,
            blue: CGFloat(hex & 255) / 255, alpha: alpha)
}

func gradient(_ stops: [(CGFloat, UInt32)]) -> CGGradient {
    CGGradient(colorsSpace: cs, colors: stops.map { rgb($0.1) } as CFArray, locations: stops.map { $0.0 })!
}

extension CGContext {
    func fill(_ path: CGPath, _ color: CGColor, shadow: (CGSize, CGFloat, CGColor)? = nil) {
        saveGState()
        if let (offset, blur, color) = shadow { setShadow(offset: offset, blur: blur, color: color) }
        addPath(path)
        setFillColor(color)
        fillPath()
        restoreGState()
    }

    func fill(_ path: CGPath, _ g: CGGradient, from: CGPoint, to: CGPoint) {
        saveGState()
        addPath(path)
        clip()
        drawLinearGradient(g, start: from, end: to, options: [.drawsBeforeStartLocation, .drawsAfterEndLocation])
        restoreGState()
    }
}

func drawTile(_ ctx: CGContext) {
    let tile = CGRect(x: 100, y: 100, width: 824, height: 824)
    let path = CGPath(roundedRect: tile, cornerWidth: 185, cornerHeight: 185, transform: nil)
    ctx.fill(path, rgb(0x1E6FE0), shadow: (CGSize(width: 0, height: -10), 28, rgb(0, 0.28)))
    ctx.fill(path, gradient([(0, 0x5CC2FF), (1, 0x1660E8)]), from: CGPoint(x: 0, y: tile.maxY), to: CGPoint(x: 0, y: tile.minY))
    // Soft glow at the top, as if lit from above.
    ctx.saveGState()
    ctx.addPath(path)
    ctx.clip()
    let glow = CGGradient(colorsSpace: cs, colors: [rgb(0xFFFFFF, 0.28), rgb(0xFFFFFF, 0)] as CFArray, locations: [0, 1])!
    ctx.drawRadialGradient(glow, startCenter: CGPoint(x: 512, y: 900), startRadius: 0,
                           endCenter: CGPoint(x: 512, y: 900), endRadius: 520, options: [])
    ctx.restoreGState()
}

func drawEnvelope(_ ctx: CGContext) {
    let env = CGRect(x: 200, y: 300, width: 560, height: 380)
    let path = CGPath(roundedRect: env, cornerWidth: 34, cornerHeight: 34, transform: nil)
    ctx.fill(path, rgb(0xFFFFFF), shadow: (CGSize(width: 0, height: -14), 30, rgb(0x0A2A6B, 0.35)))
    ctx.saveGState()
    ctx.addPath(path)
    ctx.clip()
    ctx.fill(path, gradient([(0, 0xFFFFFF), (1, 0xE6ECF5)]), from: CGPoint(x: 0, y: env.maxY), to: CGPoint(x: 0, y: env.minY))
    // Bottom folds from the lower corners, a shade darker.
    let fold = CGMutablePath()
    fold.move(to: CGPoint(x: env.minX, y: env.minY))
    fold.addLine(to: CGPoint(x: env.midX, y: env.minY + env.height * 0.46))
    fold.addLine(to: CGPoint(x: env.maxX, y: env.minY))
    fold.closeSubpath()
    ctx.fill(fold, rgb(0xDDE5F0, 0.55))
    // Top flap, casting a soft shadow on the folds.
    let flap = CGMutablePath()
    let flapY = env.minY + env.height * 0.40
    flap.move(to: CGPoint(x: env.minX, y: env.maxY))
    flap.addLine(to: CGPoint(x: env.midX - 26, y: flapY))
    flap.addQuadCurve(to: CGPoint(x: env.midX + 26, y: flapY), control: CGPoint(x: env.midX, y: env.minY + env.height * 0.34))
    flap.addLine(to: CGPoint(x: env.maxX, y: env.maxY))
    flap.closeSubpath()
    ctx.fill(flap, rgb(0xFFFFFF), shadow: (CGSize(width: 0, height: -6), 14, rgb(0x1B3A70, 0.22)))
    ctx.fill(flap, gradient([(0, 0xFFFFFF), (1, 0xF1F5FA)]), from: CGPoint(x: 0, y: env.maxY), to: CGPoint(x: 0, y: flapY))
    ctx.restoreGState()
}

/// Handwriting on the envelope that ends at the nib tip.
func drawInk(_ ctx: CGContext, to tip: CGPoint) {
    let ink = CGMutablePath()
    ink.move(to: CGPoint(x: 258, y: 380))
    ink.addCurve(to: CGPoint(x: 300, y: 438), control1: CGPoint(x: 276, y: 392), control2: CGPoint(x: 306, y: 420))
    ink.addCurve(to: CGPoint(x: 290, y: 372), control1: CGPoint(x: 294, y: 458), control2: CGPoint(x: 276, y: 420))
    ink.addCurve(to: CGPoint(x: 340, y: 384), control1: CGPoint(x: 300, y: 350), control2: CGPoint(x: 326, y: 360))
    ink.addCurve(to: CGPoint(x: 382, y: 382), control1: CGPoint(x: 350, y: 400), control2: CGPoint(x: 366, y: 404))
    ink.addCurve(to: CGPoint(x: 424, y: 380), control1: CGPoint(x: 394, y: 362), control2: CGPoint(x: 408, y: 360))
    ink.addCurve(to: tip, control1: CGPoint(x: 436, y: 398), control2: CGPoint(x: 456, y: 396))
    ctx.addPath(ink)
    ctx.setStrokeColor(rgb(0x1E4FB8))
    ctx.setLineWidth(8)
    ctx.setLineCap(.round)
    ctx.setLineJoin(.round)
    ctx.strokePath()
}

/// A fountain pen drawn along its own axis: x from the nib tip, y across (+y is the lit side).
func drawPen(_ ctx: CGContext, tip: CGPoint) {
    let length: CGFloat = 520, w: CGFloat = 74, nibLen: CGFloat = 120, sectionLen: CGFloat = 70
    ctx.saveGState()
    ctx.translateBy(x: tip.x, y: tip.y)
    ctx.rotate(by: .pi / 4)
    func across(_ path: CGPath, _ stops: [(CGFloat, UInt32)], _ h: CGFloat) {
        ctx.fill(path, gradient(stops), from: CGPoint(x: 0, y: -h), to: CGPoint(x: 0, y: h))
    }
    func taper(_ end: CGFloat, _ h: CGFloat) -> CGMutablePath {
        let p = CGMutablePath()
        p.move(to: .zero)
        p.addCurve(to: CGPoint(x: end, y: h), control1: CGPoint(x: nibLen * 0.45, y: h / 3), control2: CGPoint(x: nibLen * 0.7, y: h))
        return p
    }
    // Silhouette for the shadow; shadow offsets are in device space, so this falls down and to the right.
    let body = taper(nibLen, w * 0.36)
    body.addLine(to: CGPoint(x: length - w / 2, y: w / 2))
    body.addArc(center: CGPoint(x: length - w / 2, y: 0), radius: w / 2, startAngle: .pi / 2, endAngle: -.pi / 2, clockwise: true)
    body.addLine(to: CGPoint(x: nibLen, y: -w * 0.36))
    body.addCurve(to: .zero, control1: CGPoint(x: nibLen * 0.7, y: -w * 0.36), control2: CGPoint(x: nibLen * 0.45, y: -w * 0.12))
    body.closeSubpath()
    ctx.fill(body, rgb(0), shadow: (CGSize(width: 18, height: -10), 22, rgb(0x0A1F4D, 0.38)))

    let nib = taper(nibLen + 4, w * 0.34)
    nib.addLine(to: CGPoint(x: nibLen + 4, y: -w * 0.34))
    nib.addCurve(to: .zero, control1: CGPoint(x: nibLen * 0.7, y: -w * 0.34), control2: CGPoint(x: nibLen * 0.45, y: -w * 0.1))
    nib.closeSubpath()
    across(nib, [(0, 0x9C6B1E), (0.35, 0xE9C46A), (0.55, 0xFFF0B8), (0.75, 0xD9A93F), (1, 0x8A5A14)], w * 0.34)
    // Slit and breather hole.
    let hole = nibLen * 0.62
    ctx.setStrokeColor(rgb(0x5E3B08, 0.9))
    ctx.setLineWidth(3.2)
    ctx.move(to: CGPoint(x: 4, y: 0))
    ctx.addLine(to: CGPoint(x: hole, y: 0))
    ctx.strokePath()
    ctx.setFillColor(rgb(0x5E3B08, 0.9))
    ctx.fillEllipse(in: CGRect(x: hole - 7, y: -7, width: 14, height: 14))

    let section = CGPath(rect: CGRect(x: nibLen, y: -w * 0.4, width: sectionLen, height: w * 0.8), transform: nil)
    across(section, [(0, 0x05070C), (0.62, 0x3A4252), (0.75, 0x6B7486), (1, 0x0B0E15)], w * 0.4)
    let barrel = CGMutablePath()
    barrel.move(to: CGPoint(x: nibLen + sectionLen, y: -w / 2))
    barrel.addLine(to: CGPoint(x: length - w / 2, y: -w / 2))
    barrel.addArc(center: CGPoint(x: length - w / 2, y: 0), radius: w / 2, startAngle: -.pi / 2, endAngle: .pi / 2, clockwise: false)
    barrel.addLine(to: CGPoint(x: nibLen + sectionLen, y: w / 2))
    barrel.closeSubpath()
    across(barrel, [(0, 0x0B1630), (0.45, 0x1F3A78), (0.72, 0x5D86D6), (0.82, 0x9DBBF2), (0.9, 0x3C62B0), (1, 0x112552)], w / 2)
    for x in [nibLen + sectionLen, length - 150] {
        let band = CGPath(rect: CGRect(x: x, y: -w / 2, width: 16, height: w), transform: nil)
        across(band, [(0, 0x8A5A14), (0.6, 0xF3D27E), (0.78, 0xFFF4C9), (1, 0x9C6B1E)], w / 2)
    }
    let clip = CGPath(roundedRect: CGRect(x: length - 140, y: w * 0.18, width: 112, height: 14), cornerWidth: 7, cornerHeight: 7,
                      transform: nil)
    across(clip, [(0, 0xB98A2E), (0.5, 0xFFF0B8), (1, 0xC99A3E)], 7)
    ctx.restoreGState()
}

let out = CommandLine.arguments[1]
try? FileManager.default.createDirectory(atPath: out, withIntermediateDirectories: true)
func bitmap(_ px: Int) -> CGContext {
    CGContext(data: nil, width: px, height: px, bitsPerComponent: 8, bytesPerRow: 0, space: cs,
              bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
}
let big = bitmap(1024)
let tip = CGPoint(x: 472, y: 366)
drawTile(big)
big.saveGState()
big.addPath(CGPath(roundedRect: CGRect(x: 100, y: 100, width: 824, height: 824), cornerWidth: 185, cornerHeight: 185, transform: nil))
big.clip()
drawEnvelope(big)
drawInk(big, to: tip)
drawPen(big, tip: tip)
big.restoreGState()
let master = big.makeImage()!
for (px, name) in [(16, "16x16"), (32, "16x16@2x"), (32, "32x32"), (64, "32x32@2x"), (128, "128x128"),
                   (256, "128x128@2x"), (256, "256x256"), (512, "256x256@2x"), (512, "512x512"), (1024, "512x512@2x")] {
    let ctx = bitmap(px)
    ctx.interpolationQuality = .high
    ctx.draw(master, in: CGRect(x: 0, y: 0, width: px, height: px))
    let rep = NSBitmapImageRep(cgImage: ctx.makeImage()!)
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: "\(out)/icon_\(name).png"))
}
