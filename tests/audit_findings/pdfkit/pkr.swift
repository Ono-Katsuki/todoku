// PDFKit / Quartz renderer + annotation dump. No windows, no dialogs.
// usage: pkr render <pdf> <page1based> <out.png> <dpi> <pdfkit|quartz|pdfview>
//        pkr dump <pdf>                  (JSON: annotations + page text per page)
//        pkr ocr <png> ...               (Vision text per image, one JSON line each)
//        pkr batch <jobs.tsv>            (each line: pdf \t page \t png \t dpi \t mode)
import Foundation
import PDFKit
import AppKit
import Vision

func viewSize(_ page: PDFPage) -> (CGFloat, CGFloat) {
    let box = page.bounds(for: .cropBox)
    let rot = page.rotation
    return (rot % 180 == 0) ? (box.width, box.height) : (box.height, box.width)
}

func renderBitmap(_ page: PDFPage, _ dpi: CGFloat, _ out: String, quartz: Bool) {
    let (vw, vh) = viewSize(page)
    let k = dpi / 72.0
    let w = Int((vw * k).rounded()), h = Int((vh * k).rounded())
    let cs = CGColorSpace(name: CGColorSpace.sRGB)!
    guard let ctx = CGContext(data: nil, width: w, height: h, bitsPerComponent: 8, bytesPerRow: 0, space: cs,
                              bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else { return }
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    ctx.fill(CGRect(x: 0, y: 0, width: w, height: h))
    ctx.scaleBy(x: k, y: k)
    if quartz {
        let cg = page.pageRef!
        let t = cg.getDrawingTransform(.cropBox, rect: CGRect(x: 0, y: 0, width: vw, height: vh), rotate: 0, preserveAspectRatio: true)
        ctx.concatenate(t)
        ctx.drawPDFPage(cg)
    } else {
        NSGraphicsContext.saveGraphicsState()
        NSGraphicsContext.current = NSGraphicsContext(cgContext: ctx, flipped: false)
        page.draw(with: .cropBox, to: ctx)
        NSGraphicsContext.restoreGraphicsState()
    }
    let img = ctx.makeImage()!
    let rep = NSBitmapImageRep(cgImage: img)
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: out))
}

func renderPDFView(_ doc: PDFDocument, _ pageIndex: Int, _ dpi: CGFloat, _ out: String) {
    // An off-screen PDFView (never ordered front) showing one page at the given scale.
    let page = doc.page(at: pageIndex)!
    let (vw, vh) = viewSize(page)
    let k = dpi / 72.0
    let frame = NSRect(x: 0, y: 0, width: vw * k, height: vh * k)
    let v = PDFView(frame: frame)
    v.document = doc
    v.displayMode = .singlePage
    v.displaysPageBreaks = false
    v.pageShadowsEnabled = false
    v.autoScales = false
    v.scaleFactor = k
    v.backgroundColor = .white
    v.go(to: page)
    v.layoutDocumentView()
    RunLoop.current.run(until: Date(timeIntervalSinceNow: 0.3))
    let rep = v.bitmapImageRepForCachingDisplay(in: v.bounds)!
    v.cacheDisplay(in: v.bounds, to: rep)
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: out))
}

func dump(_ path: String) {
    guard let doc = PDFDocument(url: URL(fileURLWithPath: path)) else { print("{\"error\":\"open\"}"); return }
    var pages: [[String: Any]] = []
    for i in 0..<doc.pageCount {
        let p = doc.page(at: i)!
        var ann: [[String: Any]] = []
        for a in p.annotations {
            var d: [String: Any] = ["type": a.type ?? "", "bounds": [a.bounds.minX, a.bounds.minY, a.bounds.maxX, a.bounds.maxY]]
            d["contents"] = a.contents ?? ""
            d["name"] = a.fieldName ?? ""
            d["widgetValue"] = a.widgetStringValue ?? ""
            d["widgetType"] = a.widgetFieldType.rawValue
            d["readOnly"] = a.isReadOnly
            d["shouldDisplay"] = a.shouldDisplay
            d["shouldPrint"] = a.shouldPrint
            d["hasAppearance"] = a.hasAppearanceStream
            d["font"] = a.font?.fontName ?? ""
            d["fontSize"] = a.font?.pointSize ?? 0
            d["align"] = a.alignment.rawValue
            d["state"] = a.buttonWidgetState.rawValue
            ann.append(d)
        }
        pages.append(["annots": ann, "text": p.string ?? ""])
    }
    let data = try! JSONSerialization.data(withJSONObject: ["pages": pages], options: [.sortedKeys])
    print(String(data: data, encoding: .utf8)!)
}

func ocr(_ path: String) -> [String] {
    guard let img = NSImage(contentsOfFile: path), let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else { return [] }
    let req = VNRecognizeTextRequest()
    req.recognitionLevel = .accurate
    req.recognitionLanguages = ["ja-JP", "en-US"]
    req.usesLanguageCorrection = false
    try? VNImageRequestHandler(cgImage: cg, options: [:]).perform([req])
    return (req.results ?? []).compactMap { $0.topCandidates(1).first?.string }
}

let a = CommandLine.arguments
switch a[1] {
case "dump": dump(a[2])
case "ocr":
    for p in a[2...] {
        let d = try! JSONSerialization.data(withJSONObject: ["path": p, "text": ocr(p)], options: [])
        print(String(data: d, encoding: .utf8)!)
    }
case "render":
    let doc = PDFDocument(url: URL(fileURLWithPath: a[2]))!
    let i = Int(a[3])! - 1
    if a[6] == "pdfview" { renderPDFView(doc, i, CGFloat(Double(a[5])!), a[4]) }
    else { renderBitmap(doc.page(at: i)!, CGFloat(Double(a[5])!), a[4], quartz: a[6] == "quartz") }
case "batch":
    let lines = try! String(contentsOfFile: a[2], encoding: .utf8).split(separator: "\n")
    var cur = ""; var doc: PDFDocument? = nil
    for l in lines {
        let f = l.split(separator: "\t").map(String.init)
        if f[0] != cur { cur = f[0]; doc = PDFDocument(url: URL(fileURLWithPath: cur)) }
        guard let d = doc, let pg = d.page(at: Int(f[1])! - 1) else { print("FAIL", l); continue }
        autoreleasepool {
            if f[4] == "pdfview" { renderPDFView(d, Int(f[1])! - 1, CGFloat(Double(f[3])!), f[2]) }
            else { renderBitmap(pg, CGFloat(Double(f[3])!), f[2], quartz: f[4] == "quartz") }
        }
    }
default: print("?")
}
