import AppKit
import Foundation
import PDFKit

struct RequestPayload: Codable {
  let pdfPath: String
  let outputDir: String
  let dpi: Double?
}

struct PageImage: Codable {
  let pageIndex: Int
  let path: String
  let width: Int
  let height: Int
}

struct ResponsePayload: Codable {
  let success: Bool
  let pages: [PageImage]?
  let error: String?
}

func render(page: PDFPage, dpi: CGFloat) throws -> (NSBitmapImageRep, Int, Int) {
  let bounds = page.bounds(for: .mediaBox)
  let scale = dpi / 72.0
  let width = max(1, Int(ceil(bounds.width * scale)))
  let height = max(1, Int(ceil(bounds.height * scale)))

  guard let bitmap = NSBitmapImageRep(
    bitmapDataPlanes: nil,
    pixelsWide: width,
    pixelsHigh: height,
    bitsPerSample: 8,
    samplesPerPixel: 4,
    hasAlpha: true,
    isPlanar: false,
    colorSpaceName: .deviceRGB,
    bytesPerRow: 0,
    bitsPerPixel: 0
  ) else {
    throw NSError(domain: "word-cursor.pdf", code: 1, userInfo: [NSLocalizedDescriptionKey: "无法创建位图"])
  }

  bitmap.size = NSSize(width: bounds.width, height: bounds.height)
  NSGraphicsContext.saveGraphicsState()
  defer { NSGraphicsContext.restoreGraphicsState() }

  guard let context = NSGraphicsContext(bitmapImageRep: bitmap) else {
    throw NSError(domain: "word-cursor.pdf", code: 2, userInfo: [NSLocalizedDescriptionKey: "无法创建图形上下文"])
  }

  NSGraphicsContext.current = context
  let cg = context.cgContext
  cg.setFillColor(NSColor.white.cgColor)
  cg.fill(CGRect(x: 0, y: 0, width: width, height: height))
  cg.scaleBy(x: scale, y: scale)
  cg.translateBy(x: 0, y: bounds.height)
  cg.scaleBy(x: 1, y: -1)
  page.draw(with: .mediaBox, to: cg)

  return (bitmap, width, height)
}

do {
  let input = FileHandle.standardInput.readDataToEndOfFile()
  let request = try JSONDecoder().decode(RequestPayload.self, from: input)
  let pdfURL = URL(fileURLWithPath: request.pdfPath)
  let outputURL = URL(fileURLWithPath: request.outputDir, isDirectory: true)
  try FileManager.default.createDirectory(at: outputURL, withIntermediateDirectories: true)

  guard let document = PDFDocument(url: pdfURL) else {
    throw NSError(domain: "word-cursor.pdf", code: 3, userInfo: [NSLocalizedDescriptionKey: "无法打开 PDF"])
  }

  let dpi = CGFloat(request.dpi ?? 144)
  var pages: [PageImage] = []

  for pageIndex in 0..<document.pageCount {
    guard let page = document.page(at: pageIndex) else { continue }
    let (bitmap, width, height) = try render(page: page, dpi: dpi)
    guard let pngData = bitmap.representation(using: .png, properties: [:]) else {
      throw NSError(domain: "word-cursor.pdf", code: 4, userInfo: [NSLocalizedDescriptionKey: "无法导出 PNG"])
    }
    let fileURL = outputURL.appendingPathComponent(String(format: "page-%03d.png", pageIndex + 1))
    try pngData.write(to: fileURL)
    pages.append(PageImage(pageIndex: pageIndex + 1, path: fileURL.path, width: width, height: height))
  }

  let response = ResponsePayload(success: true, pages: pages, error: nil)
  let encoded = try JSONEncoder().encode(response)
  FileHandle.standardOutput.write(encoded)
} catch {
  let response = ResponsePayload(success: false, pages: nil, error: error.localizedDescription)
  if let encoded = try? JSONEncoder().encode(response) {
    FileHandle.standardOutput.write(encoded)
  }
  exit(1)
}
