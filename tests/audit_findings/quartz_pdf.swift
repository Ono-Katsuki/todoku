// macOS の Quartz（CoreText）で、見本の文を PDF にする（「プレビュー」「PDF として保存」と同じ作り方）。
// 使い方: swift quartz_pdf.swift IN.tsv OUT.pdf   （IN.tsv は「名前<TAB>フォント<TAB>文」の行）
import Foundation
import CoreText
import CoreGraphics

let args = CommandLine.arguments
let lines = try! String(contentsOfFile: args[1], encoding: .utf8).split(separator: "\n")
var box = CGRect(x: 0, y: 0, width: 595, height: 842)
let ctx = CGContext(URL(fileURLWithPath: args[2]) as CFURL, mediaBox: &box, nil)!
ctx.beginPDFPage(nil)
var y: CGFloat = 800
for l in lines {
    let p = l.split(separator: "\t", omittingEmptySubsequences: false).map(String.init)
    if p.count < 3 { continue }
    let label = CTFontCreateWithName("Helvetica" as CFString, 11, nil)
    let font = CTFontCreateWithName(p[1] as CFString, 13, nil)
    let s = NSMutableAttributedString(string: p[0] + " | ", attributes: [NSAttributedString.Key(kCTFontAttributeName as String): label])
    s.append(NSAttributedString(string: p[2], attributes: [NSAttributedString.Key(kCTFontAttributeName as String): font]))
    let line = CTLineCreateWithAttributedString(s)
    ctx.textPosition = CGPoint(x: 40, y: y)
    CTLineDraw(line, ctx)
    y -= 24
}
ctx.endPDFPage()
ctx.closePDF()
