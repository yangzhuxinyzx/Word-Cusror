import AppKit
import Foundation

let pointsToPixels = 96.0 / 72.0

struct MeasureEntry: Codable {
  let id: String
  let text: String
  let fontFamily: String
  let fontSize: Double
  let fontWeight: CodableWeight?
  let fontStyle: String?
  let letterSpacing: Double?
  let lineHeight: Double?
}

enum CodableWeight: Codable {
  case string(String)
  case number(Double)

  init(from decoder: Decoder) throws {
    let container = try decoder.singleValueContainer()
    if let stringValue = try? container.decode(String.self) {
      self = .string(stringValue)
      return
    }
    self = .number(try container.decode(Double.self))
  }

  func encode(to encoder: Encoder) throws {
    var container = encoder.singleValueContainer()
    switch self {
    case .string(let value):
      try container.encode(value)
    case .number(let value):
      try container.encode(value)
    }
  }
}

struct RequestPayload: Codable {
  let mode: String?
  let entries: [MeasureEntry]?
  let fonts: [String]?
}

struct MeasureResult: Codable {
  let id: String
  let width: Double
  let ascent: Double
  let descent: Double
  let lineHeight: Double
  let baseline: Double
  let resolvedFontFamily: String?
  let usedFallback: Bool
}

struct FontAvailabilityResult: Codable {
  let name: String
  let available: Bool
  let resolvedName: String?
}

struct ResponsePayload: Codable {
  let success: Bool
  let measurements: [MeasureResult]?
  let fonts: [FontAvailabilityResult]?
  let error: String?
}

func readStdin() -> Data {
  FileHandle.standardInput.readDataToEndOfFile()
}

func normalizedCandidateNames(from raw: String) -> [String] {
  raw
    .split(separator: ",")
    .map { $0.trimmingCharacters(in: .whitespacesAndNewlines).trimmingCharacters(in: CharacterSet(charactersIn: "\"'")) }
    .filter { !$0.isEmpty }
}

func resolvedWeight(_ weight: CodableWeight?) -> NSFont.Weight {
  switch weight {
  case .string(let value):
    switch value.lowercased() {
    case "bold":
      return .bold
    default:
      return .regular
    }
  case .number(let value):
    if value >= 700 {
      return .bold
    }
    if value >= 500 {
      return .medium
    }
    return .regular
  case .none:
    return .regular
  }
}

func resolveFont(_ family: String, size: CGFloat, weight: NSFont.Weight, italic: Bool) -> (font: NSFont, resolvedName: String?, usedFallback: Bool) {
  let candidates = normalizedCandidateNames(from: family)
  let availableFonts = Set(NSFontManager.shared.availableFonts)
  let availableFamilies = Set(NSFontManager.shared.availableFontFamilies)

  for candidate in candidates {
    if let direct = NSFont(name: candidate, size: size) {
      let manager = NSFontManager.shared
      let weighted = manager.font(withFamily: direct.familyName ?? candidate, traits: [], weight: Int(weight.rawValue * 1000), size: size) ?? direct
      let finalFont = italic ? manager.convert(weighted, toHaveTrait: .italicFontMask) : weighted
      return (finalFont, finalFont.familyName ?? candidate, false)
    }

    if availableFonts.contains(candidate) || availableFamilies.contains(candidate) {
      if let named = NSFont(name: candidate, size: size) {
        return (named, named.familyName ?? candidate, false)
      }
      let manager = NSFontManager.shared
      if let familyFont = manager.font(withFamily: candidate, traits: italic ? [.italicFontMask] : [], weight: Int(weight.rawValue * 1000), size: size) {
        return (familyFont, familyFont.familyName ?? candidate, false)
      }
    }
  }

  let system = NSFont.systemFont(ofSize: size, weight: weight)
  let fallback = italic ? NSFontManager.shared.convert(system, toHaveTrait: .italicFontMask) : system
  return (fallback, fallback.familyName, true)
}

func measure(entry: MeasureEntry) -> MeasureResult {
  let size = CGFloat(entry.fontSize)
  let italic = (entry.fontStyle ?? "normal").lowercased() == "italic"
  let resolved = resolveFont(entry.fontFamily, size: size, weight: resolvedWeight(entry.fontWeight), italic: italic)
  let attributes: [NSAttributedString.Key: Any] = [
    .font: resolved.font,
    .kern: NSNumber(value: entry.letterSpacing ?? 0)
  ]
  let attributed = NSAttributedString(string: entry.text, attributes: attributes)
  let width = ceil(attributed.size().width * pointsToPixels)
  let ascent = ceil(resolved.font.ascender * pointsToPixels)
  let descent = ceil(abs(resolved.font.descender) * pointsToPixels)
  let naturalLineHeight = ceil((resolved.font.ascender - resolved.font.descender + resolved.font.leading) * pointsToPixels)
  let explicitLineHeight = entry.lineHeight.map { ceil(CGFloat($0)) }
  let lineHeight = Double(explicitLineHeight ?? naturalLineHeight)

  return MeasureResult(
    id: entry.id,
    width: Double(width),
    ascent: Double(ascent),
    descent: Double(descent),
    lineHeight: lineHeight,
    baseline: Double(ascent),
    resolvedFontFamily: resolved.resolvedName,
    usedFallback: resolved.usedFallback
  )
}

func fontAvailability(for fontName: String) -> FontAvailabilityResult {
  let resolved = resolveFont(fontName, size: 12, weight: .regular, italic: false)
  let candidates = normalizedCandidateNames(from: fontName)
  let matched = !resolved.usedFallback || candidates.contains { ($0.caseInsensitiveCompare(resolved.resolvedName ?? "") == .orderedSame) }
  return FontAvailabilityResult(
    name: fontName,
    available: matched,
    resolvedName: resolved.resolvedName
  )
}

do {
  let payloadData = readStdin()
  let payload = try JSONDecoder().decode(RequestPayload.self, from: payloadData)
  let mode = payload.mode ?? "measure"

  let response: ResponsePayload
  if mode == "font-check" {
    let fonts = (payload.fonts ?? []).map(fontAvailability(for:))
    response = ResponsePayload(success: true, measurements: nil, fonts: fonts, error: nil)
  } else {
    let measurements = (payload.entries ?? []).map(measure(entry:))
    response = ResponsePayload(success: true, measurements: measurements, fonts: nil, error: nil)
  }

  let encoded = try JSONEncoder().encode(response)
  if let output = String(data: encoded, encoding: .utf8) {
    FileHandle.standardOutput.write(output.data(using: .utf8)!)
  }
} catch {
  let response = ResponsePayload(success: false, measurements: nil, fonts: nil, error: error.localizedDescription)
  if let encoded = try? JSONEncoder().encode(response) {
    FileHandle.standardOutput.write(encoded)
  }
  exit(1)
}
