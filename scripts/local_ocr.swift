import Foundation
import Vision
import AppKit
import PDFKit

func recognize(_ image: CGImage) throws -> String {
    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.recognitionLanguages = ["zh-Hans", "en-US"]
    request.usesLanguageCorrection = true
    try VNImageRequestHandler(cgImage: image).perform([request])
    return (request.results ?? []).compactMap { $0.topCandidates(1).first?.string }.joined(separator: "\n")
}
do {
    guard CommandLine.arguments.count >= 2 else { throw NSError(domain:"OCR",code:1) }
    let url = URL(fileURLWithPath: CommandLine.arguments[1])
    let image: CGImage
    if url.pathExtension.lowercased() == "pdf" {
        let number = Int(CommandLine.arguments.dropFirst(2).first ?? "1") ?? 1
        guard let pdf = PDFDocument(url:url), let page = pdf.page(at:number-1) else { throw NSError(domain:"PDF",code:2) }
        let bounds = page.bounds(for:.mediaBox)
        let scale = min(3.0, 3000.0 / max(bounds.width,bounds.height))
        let bitmap = page.thumbnail(of:NSSize(width:bounds.width*scale,height:bounds.height*scale),for:.mediaBox)
        guard let cg = bitmap.cgImage(forProposedRect:nil,context:nil,hints:nil) else { throw NSError(domain:"PDF",code:3) }
        image = cg
    } else {
        guard let source = NSImage(contentsOf:url), let cg = source.cgImage(forProposedRect:nil,context:nil,hints:nil) else { throw NSError(domain:"Image",code:4) }
        image = cg
    }
    let data = try JSONSerialization.data(withJSONObject:["text":try recognize(image)])
    FileHandle.standardOutput.write(data)
} catch {
    FileHandle.standardError.write(Data("本地 OCR 无法读取此图像或页面".utf8))
    exit(1)
}
