import AppKit
import WebKit
import UniformTypeIdentifiers
import Foundation

let appName = "XLSX Lite"
let supportedExtensions = ["xlsx", "xlsm"]
let italian = (Locale.preferredLanguages.first ?? "").lowercased().hasPrefix("it")
func L(_ it: String, _ en: String) -> String { italian ? it : en }
func failure(_ message: String, _ code: Int = 1) -> NSError { NSError(domain: appName, code: code, userInfo: [NSLocalizedDescriptionKey: message]) }

func esc(_ s: String) -> String {
    guard s.contains(where: { "&<>\"'".contains($0) }) else { return s }
    var out = ""; out.reserveCapacity(s.utf8.count + 16)
    for ch in s { switch ch { case "&": out += "&amp;"; case "<": out += "&lt;"; case ">": out += "&gt;"; case "\"": out += "&quot;"; case "'": out += "&#39;"; default: out.append(ch) } }
    return out
}
extension XMLNode {
    func nodes(_ path: String) -> [XMLNode] { (try? nodes(forXPath: path)) ?? [] }
    func attr(_ name: String) -> String { (self as? XMLElement)?.attribute(forName: name)?.stringValue ?? "" }
    func val(_ path: String) -> String { nodes(path).first?.stringValue ?? "" }
}
extension XMLElement {
    // Direct child lookups: much faster than XPath when called for every cell.
    func child(_ name: String) -> XMLElement? { children?.first { $0.localName == name } as? XMLElement }
    func kids(_ name: String) -> [XMLElement] { children?.compactMap { $0.localName == name ? $0 as? XMLElement : nil } ?? [] }
}
func column(_ ref: String) -> Int { ref.prefix(while: { $0.isASCII && $0.isLetter }).uppercased().unicodeScalars.reduce(0) { $0 * 26 + Int($1.value) - 64 } }
func rowNumber(_ ref: String) -> Int { Int(ref.drop(while: { $0.isLetter })) ?? 0 }
func columnName(_ c: Int) -> String { var v = c, s = ""; while v > 0 { v -= 1; s = String(UnicodeScalar(65 + v % 26)!) + s; v /= 26 }; return s }
func cellKey(_ row: Int, _ col: Int) -> Int { row * 16_385 + col }
/// Text of a shared string or inline string, skipping phonetic runs.
func richText(_ e: XMLElement) -> String {
    var out = ""
    for child in e.children ?? [] {
        if child.localName == "t" { out += child.stringValue ?? "" }
        else if child.localName == "r", let t = (child as? XMLElement)?.child("t") { out += t.stringValue ?? "" }
    }
    return out
}

enum ValueKind { case general, text, date, number }
struct CellStyle {
    var css = ""
    var kind = ValueKind.general
    var formatter: Formatter?
    var prefix = "", suffix = ""

    mutating func setFormat(_ format: String) {
        let lower = format.lowercased()
        if format.isEmpty || lower == "general" { return }
        if format == "@" { kind = .text; return }
        let section = String(format.split(separator: ";", omittingEmptySubsequences: false).first ?? "")
        let clean = section.replacingOccurrences(of: "\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.|_.|\\*.", with: "", options: .regularExpression).lowercased()
        if clean.contains("y") || clean.contains("d") || clean.contains("h") {
            let f = DateFormatter(); f.locale = .current; f.timeZone = TimeZone(secondsFromGMT: 0)
            var template = ""
            if clean.contains("y") || clean.contains("d") { template = clean.contains("d") ? (clean.contains("y") ? "ddMMyyyy" : "ddMM") : "MMyyyy" }
            if clean.contains("h") { template += clean.contains("s") ? "jjmmss" : "jjmm" }
            f.setLocalizedDateFormatFromTemplate(template)
            kind = .date; formatter = f; return
        }
        guard clean.contains("0") || clean.contains("#") else { return }
        let f = NumberFormatter(); f.locale = .current
        f.numberStyle = clean.contains("%") ? .percent : .decimal
        let digits = clean.split(separator: ".").dropFirst().first?.prefix(while: { $0 == "0" || $0 == "#" }).count ?? 0
        f.minimumFractionDigits = digits; f.maximumFractionDigits = digits
        f.usesGroupingSeparator = clean.contains(",")
        for symbol in ["€", "$", "£", "¥", "CHF"] where section.contains(symbol) {
            let firstDigit = section.firstIndex(where: { $0 == "0" || $0 == "#" }) ?? section.endIndex
            if let at = section.range(of: symbol), at.lowerBound < firstDigit { prefix = symbol } else { suffix = " " + symbol }
            break
        }
        kind = .number; formatter = f
    }
}

struct Sheet { let name: String; let path: String; let hidden: Bool }
struct Rendered { let html: String; let rows: Int; let columns: Int; let shownRows: Int }
enum CellKind { case text, number, formula }
struct SheetIndex { var cells: [Int: XMLElement] = [:]; var hiddenRows = Set<Int>(); var maxRow = 0, maxCol = 0 }

final class Book {
    let url: URL
    var documents: [String: XMLDocument] = [:]
    var dirty = false
    var strings: [String] = []; var sheets: [Sheet] = []; var styles: [CellStyle] = []; var date1904 = false
    var activeSheet = 0
    private var indexes: [Int: SheetIndex] = [:]
    private var memo: [Int: FValue] = [:]
    private var evaluating = Set<Int>()
    func invalidate() { indexes.removeAll(); memo.removeAll() }
    static let general: NumberFormatter = { let f = NumberFormatter(); f.locale = .current; f.numberStyle = .decimal; f.usesGroupingSeparator = false; f.maximumFractionDigits = 10; return f }()

    init(_ url: URL) throws {
        self.url = url
        let workbook = try xml("xl/workbook.xml")
        date1904 = ["1", "true"].contains(workbook.val("/workbook/workbookPr/@date1904"))
        let rels = try xml("xl/_rels/workbook.xml.rels").nodes("/Relationships/Relationship")
        for sheet in workbook.nodes("/workbook/sheets/sheet") {
            let id = sheet.attr("r:id")
            if let rel = rels.first(where: { $0.attr("Id") == id }) {
                let target = rel.attr("Target")
                let path = target.hasPrefix("/") ? String(target.dropFirst()) : "xl/" + target
                sheets.append(Sheet(name: sheet.attr("name"), path: path, hidden: ["hidden", "veryHidden"].contains(sheet.attr("state"))))
            }
        }
        if sheets.isEmpty { throw failure(L("Il file non contiene fogli leggibili.", "The file has no readable sheets.")) }
        let active = Int(workbook.val("/workbook/bookViews/workbookView/@activeTab")) ?? 0
        activeSheet = sheets.indices.contains(active) && !sheets[active].hidden ? active : (sheets.firstIndex { !$0.hidden } ?? 0)
        if let shared = try? xml("xl/sharedStrings.xml"), let root = shared.rootElement() { strings = root.kids("si").map(richText) }
        if let doc = try? xml("xl/styles.xml") {
            let fonts = doc.nodes("/styleSheet/fonts/font"), fills = doc.nodes("/styleSheet/fills/fill")
            var formats = [14: "dd/mm/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm", 19: "h:mm:ss", 20: "h:mm", 21: "h:mm:ss", 22: "dd/mm/yyyy h:mm",
                           9: "0%", 10: "0.00%", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 11: "0.00E+00", 37: "#,##0", 38: "#,##0", 39: "#,##0.00", 40: "#,##0.00", 49: "@"]
            for f in doc.nodes("/styleSheet/numFmts/numFmt") { formats[Int(f.attr("numFmtId")) ?? -1] = f.attr("formatCode") }
            for x in doc.nodes("/styleSheet/cellXfs/xf") {
                var s = CellStyle(); s.setFormat(formats[Int(x.attr("numFmtId")) ?? 0] ?? "")
                if let i = Int(x.attr("fontId")), fonts.indices.contains(i) {
                    let f = fonts[i]
                    if !f.nodes("b").isEmpty && f.val("b/@val") != "0" { s.css += "font-weight:600;" }
                    if !f.nodes("i").isEmpty && f.val("i/@val") != "0" { s.css += "font-style:italic;" }
                    if let c = color(f.nodes("color").first), c != "#000000" { s.css += "color:\(c);" }
                }
                if let i = Int(x.attr("fillId")), fills.indices.contains(i), fills[i].val("patternFill/@patternType") == "solid", let c = color(fills[i].nodes("patternFill/fgColor").first) { s.css += "background:\(c);" }
                let align = x.val("alignment/@horizontal")
                if ["left", "right", "center"].contains(align) { s.css += "text-align:\(align);" }
                styles.append(s)
            }
        }
    }
    func color(_ node: XMLNode?) -> String? {
        guard let node else { return nil }
        let rgb = node.attr("rgb"); if rgb.count >= 6 { return "#" + rgb.suffix(6) }
        let colors = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"]
        if let i = Int(node.attr("theme")), colors.indices.contains(i) { return colors[i] }
        return nil
    }
    func xml(_ member: String) throws -> XMLDocument {
        let process = Process(); process.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
        let escaped = member.replacingOccurrences(of: "[", with: "\\[").replacingOccurrences(of: "*", with: "\\*").replacingOccurrences(of: "?", with: "\\?")
        process.arguments = ["-p", url.path, escaped]
        let pipe = Pipe(); process.standardOutput = pipe; process.standardError = FileHandle.nullDevice
        try process.run(); let data = pipe.fileHandleForReading.readDataToEndOfFile(); process.waitUntilExit()
        guard process.terminationStatus == 0 else { throw failure(L("File non valido, protetto da password o non leggibile.", "The file is invalid, password-protected or unreadable."), 2) }
        return try XMLDocument(data: data, options: [.nodePreserveWhitespace, .nodeLoadExternalEntitiesNever])
    }
    func sheetDocument(_ index: Int) throws -> XMLDocument {
        let path = sheets[index].path
        if let doc = documents[path] { return doc }
        let doc = try xml(path); documents[path] = doc; return doc
    }
    func style(_ cell: XMLElement) -> CellStyle {
        let idx = Int(cell.attribute(forName: "s")?.stringValue ?? "") ?? 0
        return styles.indices.contains(idx) ? styles[idx] : CellStyle()
    }
    func formatNumber(_ n: Double, _ style: CellStyle, raw: String) -> String {
        switch style.kind {
        case .text: return raw
        case .date:
            guard let f = style.formatter as? DateFormatter else { return raw }
            return f.string(from: Date(timeIntervalSince1970: (n - (date1904 ? 24107 : 25569)) * 86400))
        case .number:
            return style.prefix + ((style.formatter as? NumberFormatter)?.string(from: NSNumber(value: n)) ?? raw) + style.suffix
        case .general:
            if n == n.rounded(), abs(n) < 1e15 { return String(Int64(n)) }
            return Book.general.string(from: NSNumber(value: n)) ?? raw
        }
    }
    func format(_ value: FValue, _ style: CellStyle) -> String {
        switch value {
        case .number(let n): return n.isFinite ? formatNumber(n, style, raw: value.text) : "#NUM!"
        case .bool(let b): return b ? L("VERO", "TRUE") : L("FALSO", "FALSE")
        case .list(let l): return format(l.first ?? .empty, style)
        default: return value.text
        }
    }
    /// The formula text of a cell that this app can calculate itself (not shared/array formulas).
    func formulaText(_ cell: XMLElement) -> String? {
        guard let f = cell.child("f"), let text = f.stringValue, !text.isEmpty, !["array", "dataTable"].contains(f.attr("t")) else { return nil }
        return text
    }
    func display(_ cell: XMLElement, sheet: Int) -> String {
        if cell.child("f") != nil, cell.child("v")?.stringValue?.isEmpty ?? true || dirty, let value = try? computed(cell, sheet: sheet) {
            return format(value, style(cell))
        }
        let type = cell.attribute(forName: "t")?.stringValue ?? ""
        if type == "inlineStr" { return cell.child("is").map(richText) ?? "" }
        let v = cell.child("v")?.stringValue ?? ""
        if v.isEmpty { return cell.child("f") != nil ? L("↻ si ricalcola in Excel", "↻ recalculated in Excel") : "" }
        switch type {
        case "s": return Int(v).flatMap { strings.indices.contains($0) ? strings[$0] : nil } ?? v
        case "b": return v == "1" ? L("VERO", "TRUE") : L("FALSO", "FALSE")
        case "str", "e": return v
        default: break
        }
        guard let n = Double(v), n.isFinite else { return v }
        return formatNumber(n, style(cell), raw: v)
    }
    func index(_ sheet: Int) throws -> SheetIndex {
        if let cached = indexes[sheet] { return cached }
        var result = SheetIndex(), r = 0
        for row in try sheetDocument(sheet).rootElement()?.child("sheetData")?.kids("row") ?? [] {
            r = Int(row.attr("r")) ?? r + 1
            if ["1", "true"].contains(row.attr("hidden")) { result.hiddenRows.insert(r) }
            var c = 0
            for cell in row.kids("c") {
                let ref = cell.attr("r"); c = ref.isEmpty ? c + 1 : column(ref)
                guard c > 0, cell.child("v") != nil || cell.child("is") != nil || cell.child("f") != nil else { continue }
                result.cells[cellKey(r, c)] = cell; result.maxRow = max(result.maxRow, r); result.maxCol = max(result.maxCol, c)
            }
        }
        indexes[sheet] = result
        return result
    }
    /// Value of a formula cell computed by this app; throws when the formula uses something unsupported.
    func computed(_ cell: XMLElement, sheet: Int) throws -> FValue {
        guard let text = formulaText(cell) else { throw FormulaUnsupported() }
        let ref = cell.attr("r"), key = sheet << 40 | cellKey(rowNumber(ref), column(ref))
        if let cached = memo[key] { return cached }
        guard !evaluating.contains(key) else { return .error("#REF!") }
        evaluating.insert(key); defer { evaluating.remove(key) }
        let value = try FormulaParser(book: self, sheet: sheet).evaluate(text)
        memo[key] = value
        return value
    }
    func value(_ sheet: Int, _ row: Int, _ col: Int) throws -> FValue {
        guard let cell = try index(sheet).cells[cellKey(row, col)] else { return .empty }
        if cell.child("f") != nil, cell.child("v")?.stringValue?.isEmpty ?? true || dirty, let v = try? computed(cell, sheet: sheet) { return v }
        let type = cell.attribute(forName: "t")?.stringValue ?? ""
        let v = cell.child("v")?.stringValue ?? ""
        switch type {
        case "inlineStr": return .text(cell.child("is").map(richText) ?? "")
        case "s": return .text(Int(v).flatMap { strings.indices.contains($0) ? strings[$0] : nil } ?? v)
        case "b": return .bool(v == "1")
        case "str": return .text(v)
        case "e": return .error(v)
        default: return v.isEmpty ? (cell.child("f") != nil ? .error("#N/A") : .empty) : (Double(v).map { .number($0) } ?? .text(v))
        }
    }
    func rangeValues(_ sheet: Int, _ r1: Int, _ c1: Int, _ r2: Int, _ c2: Int) throws -> [FValue] {
        let idx = try index(sheet)
        let lastRow = min(r2, max(idx.maxRow, r1)), lastCol = min(c2, max(idx.maxCol, c1))
        guard (lastRow - r1 + 1) * (lastCol - c1 + 1) <= 2_000_000 else { throw FormulaUnsupported() }
        var out: [FValue] = []
        for r in r1...lastRow { for c in c1...lastCol { out.append(try value(sheet, r, c)) } }
        return out
    }
    func isNumeric(_ cell: XMLElement) -> Bool {
        let type = cell.attribute(forName: "t")?.stringValue ?? "n"
        guard type == "n", let v = cell.child("v")?.stringValue, Double(v) != nil else { return false }
        let idx = Int(cell.attribute(forName: "s")?.stringValue ?? "") ?? 0
        return !(styles.indices.contains(idx) && styles[idx].kind == .text)
    }
    func cell(_ index: Int, _ ref: String) throws -> XMLElement? {
        guard ref.range(of: "^[A-Z]{1,3}[1-9][0-9]{0,6}$", options: .regularExpression) != nil else { return nil }
        return try sheetDocument(index).nodes("/worksheet/sheetData/row/c[@r='\(ref)']").first as? XMLElement
    }
    func setCell(_ index: Int, _ ref: String, _ value: String, kind: CellKind) throws {
        guard ref.range(of: "^[A-Z]{1,3}[1-9][0-9]{0,6}$", options: .regularExpression) != nil else { return }
        let doc = try sheetDocument(index)
        guard doc.nodes("/worksheet/sheetProtection").isEmpty else { throw failure(L("Questo foglio è protetto. Rimuovi la protezione in Excel prima di modificarlo.", "This sheet is protected. Remove the protection in Excel before editing it.")) }
        guard let data = doc.nodes("/worksheet/sheetData").first as? XMLElement else { return }
        let r = rowNumber(ref)
        let row: XMLElement
        if let existing = data.nodes("row[@r='\(r)']").first as? XMLElement { row = existing }
        else { row = XMLElement(name: "row"); row.addAttribute(XMLNode.attribute(withName: "r", stringValue: String(r)) as! XMLNode); let place = (data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0) > r } ?? data.childCount; data.insertChild(row, at: place) }
        let c: XMLElement
        if let existing = try cell(index, ref) { c = existing }
        else { c = XMLElement(name: "c"); c.addAttribute(XMLNode.attribute(withName: "r", stringValue: ref) as! XMLNode); let place = (row.children ?? []).firstIndex { column($0.attr("r")) > column(ref) } ?? row.childCount; row.insertChild(c, at: place) }
        if let f = c.nodes("f").first, ["shared", "array", "dataTable"].contains(f.attr("t")) { throw failure(L("Questa cella fa parte di una formula condivisa o matriciale: modificala in Excel.", "This cell is part of a shared or array formula: edit it in Excel."), 2) }
        for node in c.nodes("f|v|is") { node.detach() }; c.removeAttribute(forName: "t")
        if !value.isEmpty {
            if kind == .formula { c.addChild(XMLElement(name: "f", stringValue: normalizeFormula(value))) }
            else if kind == .number { c.addChild(XMLElement(name: "v", stringValue: value)) }
            else { c.addAttribute(XMLNode.attribute(withName: "t", stringValue: "inlineStr") as! XMLNode); let inline = XMLElement(name: "is"); let t = XMLElement(name: "t", stringValue: value); t.addAttribute(XMLNode.attribute(withName: "xml:space", stringValue: "preserve") as! XMLNode); inline.addChild(t); c.addChild(inline) }
        }
        dirty = true
        invalidate()
    }
    func save(to destination: URL) throws {
        if !dirty { if destination.standardizedFileURL != url.standardizedFileURL { try Data(contentsOf: url).write(to: destination, options: .atomic) }; return }
        let fm = FileManager.default, temp = fm.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try fm.createDirectory(at: temp, withIntermediateDirectories: true); defer { try? fm.removeItem(at: temp) }
        let archive = temp.appendingPathComponent("edited.xlsx"), staging = temp.appendingPathComponent("contents")
        try fm.copyItem(at: url, to: archive); try fm.createDirectory(at: staging, withIntermediateDirectories: true)
        var members = [String]()
        func write(_ path: String, _ doc: XMLDocument) throws {
            guard !path.hasPrefix("/"), !path.split(separator: "/").contains("..") else { throw failure(L("Percorso interno non valido.", "Invalid internal path."), 3) }
            let target = staging.appendingPathComponent(path); try fm.createDirectory(at: target.deletingLastPathComponent(), withIntermediateDirectories: true)
            try doc.xmlData(options: []).write(to: target); members.append(path)
        }
        for i in sheets.indices where documents[sheets[i].path] != nil {
            let doc = try XMLDocument(data: sheetDocument(i).xmlData, options: [.nodePreserveWhitespace])
            // Discard stale formula caches: Excel recomputes them when the file is opened.
            for cell in doc.nodes("/worksheet/sheetData/row/c[f]") { for v in cell.nodes("v") { v.detach() } }
            try write(sheets[i].path, doc)
        }
        let workbook = try xml("xl/workbook.xml")
        let calc = (workbook.nodes("/workbook/calcPr").first as? XMLElement) ?? XMLElement(name: "calcPr")
        if calc.parent == nil { workbook.rootElement()?.addChild(calc) }
        for (key, value) in [("calcMode", "auto"), ("fullCalcOnLoad", "1"), ("forceFullCalc", "1")] { calc.removeAttribute(forName: key); calc.addAttribute(XMLNode.attribute(withName: key, stringValue: value) as! XMLNode) }
        try write("xl/workbook.xml", workbook)
        let rels = try xml("xl/_rels/workbook.xml.rels")
        for r in rels.nodes("/Relationships/Relationship") where r.attr("Type").hasSuffix("/calcChain") { r.detach() }
        try write("xl/_rels/workbook.xml.rels", rels)
        let types = try xml("[Content_Types].xml")
        for t in types.nodes("/Types/Override") where t.attr("PartName") == "/xl/calcChain.xml" { t.detach() }
        try write("[Content_Types].xml", types)
        func zip(_ args: [String], acceptMissing: Bool = false) throws {
            let process = Process(); process.executableURL = URL(fileURLWithPath: "/usr/bin/zip"); process.arguments = args; process.currentDirectoryURL = staging; process.standardOutput = FileHandle.nullDevice; process.standardError = FileHandle.nullDevice
            try process.run(); process.waitUntilExit()
            if process.terminationStatus != 0 && !(acceptMissing && process.terminationStatus == 12) { throw failure(L("Impossibile creare il file. L'originale non è stato modificato.", "Could not write the file. The original was not changed."), 4) }
        }
        try zip(["-q", "-d", archive.path, "xl/calcChain.xml"], acceptMissing: true)
        try zip(["-q", archive.path] + members)
        _ = try Book(archive)
        try Data(contentsOf: archive).write(to: destination, options: .atomic)
    }

    func render(_ index: Int) throws -> Rendered {
        let doc = try sheetDocument(index)
        guard let root = doc.rootElement() else { throw failure(L("Foglio non leggibile.", "Unreadable sheet.")) }
        memo.removeAll()
        let sheetIndex = try self.index(index)
        let grid = sheetIndex.cells, hiddenRows = sheetIndex.hiddenRows
        var maxRow = sheetIndex.maxRow, maxCol = sheetIndex.maxCol
        maxRow = max(maxRow, 1); maxCol = min(max(maxCol, 1), 1000)
        var widths = [Int: Double](); var hiddenCols = Set<Int>()
        for c in root.child("cols")?.kids("col") ?? [] {
            let lo = Int(c.attr("min")) ?? 1, hi = min(maxCol, Int(c.attr("max")) ?? 1)
            guard lo <= hi else { continue }
            for i in lo...hi {
                widths[i] = min(620, max(48, (Double(c.attr("width")) ?? 12) * 7 + 8))
                if ["1", "true"].contains(c.attr("hidden")) { hiddenCols.insert(i) }
            }
        }
        let columns = (1...maxCol).filter { !hiddenCols.contains($0) }
        let allRows = (1...maxRow).filter { !hiddenRows.contains($0) }
        // Keep very large sheets responsive: render at most ~400k cells.
        let rows = Array(allRows.prefix(max(200, 400_000 / max(1, columns.count))))
        let lastRow = rows.last ?? 1
        var merges: [Int: (Int, Int)] = [:]; var covered = Set<Int>()
        for merge in root.child("mergeCells")?.kids("mergeCell") ?? [] {
            let parts = merge.attr("ref").split(separator: ":").map(String.init); guard parts.count == 2 else { continue }
            let c1 = column(parts[0]), c2 = min(maxCol, column(parts[1])), r1 = rowNumber(parts[0]), r2 = min(lastRow, rowNumber(parts[1]))
            guard c1 >= 1, r1 >= 1, c1 <= c2, r1 <= r2 else { continue }
            merges[cellKey(r1, c1)] = ((c1...c2).filter { !hiddenCols.contains($0) }.count, (r1...r2).filter { !hiddenRows.contains($0) }.count)
            for rr in r1...r2 { for cc in c1...c2 where rr != r1 || cc != c1 { covered.insert(cellKey(rr, cc)) } }
        }
        var offsets = [Int: Double](); var offset = 40.0
        for c in columns { offsets[c] = offset; offset += widths[c] ?? 110 }
        let totalWidth = offset
        let pane = root.child("sheetViews")?.child("sheetView")?.child("pane")
        let frozen = pane?.attr("state").hasPrefix("frozen") == true
        let frozenRows = frozen ? min(5, Int(pane?.attr("ySplit") ?? "") ?? 0) : 0
        let frozenCols = frozen ? min(3, Int(pane?.attr("xSplit") ?? "") ?? 0) : 0
        let filterParts = (root.child("autoFilter")?.attr("ref") ?? "").split(separator: ":").map(String.init)
        let filterStart = filterParts.first ?? "", filterEnd = filterParts.last ?? ""
        let filterRow = rowNumber(filterStart), filterLastRow = min(lastRow, max(rowNumber(filterEnd), filterRow))
        let marginTop = 28 + frozenRows * 27, marginLeft = Int(offsets[columns.first { $0 > frozenCols } ?? 1] ?? 40) + 4
        let strings = ["search": L("Cerca valore…", "Search values…"), "all": L("Tutti i valori", "All values"), "empty": L("(Vuote)", "(Blanks)"), "more": L("Altri valori: usa la ricerca", "More values: use search")]
        let stringsJSON = String(data: try JSONSerialization.data(withJSONObject: strings), encoding: .utf8)!

        var html = """
        <!doctype html><html><head><meta charset="utf-8"><style>
        :root{color-scheme:light}*{box-sizing:border-box}body{margin:0;background:#fff;color:#202124;font:13px -apple-system,system-ui;font-variant-numeric:tabular-nums}
        table{border-collapse:separate;border-spacing:0;table-layout:fixed;width:\(Int(totalWidth))px}
        td,th{border-right:1px solid #dfe3e9;border-bottom:1px solid #e5e7eb;background:#fff;height:27px;padding:0 8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:middle;scroll-margin:\(marginTop)px 4px 4px \(marginLeft)px}
        th{background:#f3f4f6;color:#7b8088;font-size:10px;font-weight:500;text-align:center;padding:0}
        thead th{position:sticky;top:0;z-index:6;height:24px}tbody th{position:sticky;left:0;z-index:4}
        td.n{text-align:right}tr.heading td{height:50px;font-size:11px;line-height:14px;white-space:normal;padding:4px 6px}
        td.sel{outline:2px solid #2979e8;outline-offset:-2px}td.sel::selection{background:transparent}tr:hover td{box-shadow:inset 0 0 0 999px #8080800c}::selection{background:#b9d9ff}
        td.fh{cursor:pointer;padding-right:18px;position:relative}td.fh::after{content:'▾';position:absolute;right:5px;top:50%;transform:translateY(-50%);color:#7b8088}td.fh.on::after{content:'●';color:#06f}
        #filter-menu{position:fixed;z-index:100;background:#fff;color:#222;border:1px solid #ddd;border-radius:9px;box-shadow:0 5px 25px #0003;width:230px;padding:6px;max-height:350px;overflow:auto;font:13px -apple-system,system-ui}
        #filter-menu button{display:block;border:0;background:transparent;border-radius:4px;text-align:left;width:100%;padding:7px;color:#222;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}#filter-menu button:hover{background:#e6efff}
        #filter-menu input{width:100%;padding:7px;border:1px solid #ddd;border-radius:4px;margin-bottom:4px}#filter-menu .more{padding:7px;color:#888;font-size:12px}
        @media print{thead th,tbody th,td{position:static!important}td.sel{outline:none}tr:hover td{box-shadow:none}td.fh::after{content:''}}
        </style></head><body><table><colgroup><col style="width:40px">
        """
        html.reserveCapacity(rows.count * columns.count * 24 + 4096)
        for c in columns { html += "<col style='width:\(Int(widths[c] ?? 110))px'>" }
        html += "</colgroup><thead><tr><th style='left:0;z-index:9'></th>"
        for c in columns { html += "<th\(c <= frozenCols ? " style='left:\(Int(offsets[c] ?? 40))px;z-index:8'" : "")>\(columnName(c))</th>" }
        html += "</tr></thead><tbody>"
        for (i, r) in rows.enumerated() {
            let isFrozen = r <= frozenRows, top = 24 + i * 27
            html += "<tr data-row='\(r)'\(r == frozenRows && frozenRows > 0 ? " class='heading'" : "")><th\(isFrozen ? " style='top:\(top)px;z-index:8'" : "")>\(r)</th>"
            for c in columns {
                let key = cellKey(r, c); if covered.contains(key) { continue }
                let cell = grid[key]
                var css = ""
                if let cell, let s = Int(cell.attribute(forName: "s")?.stringValue ?? ""), styles.indices.contains(s) { css = styles[s].css }
                if isFrozen { css += "position:sticky;top:\(top)px;z-index:5;" }
                if c <= frozenCols { css += "position:sticky;left:\(Int(offsets[c] ?? 40))px;z-index:\(isFrozen ? 7 : 3);" }
                var classes: [String] = []
                if r == filterRow && c >= column(filterStart) && c <= column(filterEnd) { classes.append("fh") }
                if let cell, isNumeric(cell) || (cell.child("f") != nil && { if case .number? = try? computed(cell, sheet: index) { return true }; return false }()) { classes.append("n") }
                html += "<td data-c='\(c)'"
                if !classes.isEmpty { html += " class='\(classes.joined(separator: " "))'" }
                if let span = merges[key] { html += " colspan='\(span.0)' rowspan='\(span.1)'" }
                if !css.isEmpty { html += " style='\(esc(css))'" }
                if let cell {
                    let value = display(cell, sheet: index)
                    if value.contains("\n") { html += " data-full='\(esc(value))'" }
                    if value.count > 28 { html += " title='\(esc(value))'" }
                    html += ">\(esc(value.replacingOccurrences(of: "\n", with: " ")))</td>"
                } else { html += "></td>" }
            }
            html += "</tr>"
        }
        html += """
        </tbody></table><script>
        const T=\(stringsJSON),FR=\(filterRow),FL=\(filterLastRow);let filters={},current=null;
        const full=td=>td.dataset.full??td.textContent;
        function colName(c){let s='';while(c>0){c--;s=String.fromCharCode(65+c%26)+s;c=Math.floor(c/26)}return s}
        const refOf=td=>colName(+td.dataset.c)+td.parentElement.dataset.row;
        const cellAt=(tr,c)=>tr.querySelector(`td[data-c="${c}"]`);
        const post=m=>window.webkit?.messageHandlers.cell.postMessage(m);
        function closeFilter(){document.getElementById('filter-menu')?.remove()}
        function dataRows(){return [...document.querySelectorAll('tbody tr')].filter(r=>{const n=+r.dataset.row;return n>FR&&n<=FL})}
        function applyFilters(){const keys=Object.entries(filters);dataRows().forEach(row=>{row.hidden=keys.some(([c,v])=>{const td=cellAt(row,c);return (td?full(td):'')!==v})});document.querySelectorAll('td.fh').forEach(td=>td.classList.toggle('on',filters[td.dataset.c]!==undefined));closeFilter()}
        window.clearFilters=()=>{filters={};applyFilters()};
        function select(td,scroll){current?.classList.remove('sel');current=td;td.classList.add('sel');const s=getSelection(),range=document.createRange();range.selectNodeContents(td);s.removeAllRanges();s.addRange(range);if(scroll)td.scrollIntoView({block:'nearest',inline:'nearest'});post({ref:refOf(td),text:full(td)})}
        window.setCells=list=>{let chosen=null;for(const [row,c,value,pick] of list){const tr=document.querySelector(`tr[data-row="${row}"]`),td=tr&&cellAt(tr,c);if(!td)continue;td.textContent=value.replaceAll('\\n',' ');if(value.includes('\\n'))td.dataset.full=value;else delete td.dataset.full;td.title=value.length>28?value:'';if(pick)chosen=td}if(Object.keys(filters).length)applyFilters();if(chosen)select(chosen)};
        function showFilter(td){closeFilter();const c=td.dataset.c,seen=new Set();dataRows().forEach(row=>{const x=cellAt(row,c);seen.add(x?full(x):'')});
        const values=[...seen].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));const menu=document.createElement('div');menu.id='filter-menu';const input=document.createElement('input');input.placeholder=T.search;const list=document.createElement('div');menu.append(input,list);
        const add=(label,on,fn)=>{const b=document.createElement('button');b.textContent=(on?'✓ ':'')+label;b.title=label;b.onclick=fn;list.append(b)};
        const draw=()=>{list.textContent='';const q=input.value.toLowerCase();if(!q)add(T.all,filters[c]===undefined,()=>{delete filters[c];applyFilters()});let n=0;for(const v of values){if(q&&!v.toLowerCase().includes(q))continue;if(n++>=300){const more=document.createElement('div');more.className='more';more.textContent=T.more;list.append(more);break}add(v||T.empty,filters[c]===v,()=>{filters[c]=v;applyFilters()})}};
        input.oninput=draw;draw();document.body.append(menu);const box=td.getBoundingClientRect();menu.style.left=Math.max(6,Math.min(box.left,innerWidth-236))+'px';menu.style.top=Math.max(6,Math.min(box.bottom+3,innerHeight-356))+'px';input.focus()}
        function neighbor(td,dx,dy){if(dx){let x=dx>0?td.nextElementSibling:td.previousElementSibling;while(x&&x.tagName!=='TD')x=dx>0?x.nextElementSibling:x.previousElementSibling;return x}
        let row=td.parentElement;const c=+td.dataset.c;while(true){row=dy>0?row.nextElementSibling:row.previousElementSibling;if(!row)return null;if(row.hidden)continue;let x=cellAt(row,c);if(!x){for(const y of row.children)if(y.tagName==='TD'&&+y.dataset.c<=c)x=y}if(x)return x}}
        document.addEventListener('click',e=>{if(e.target.closest('#filter-menu'))return;const td=e.target.closest('td');if(!td){closeFilter();return}select(td);if(td.classList.contains('fh'))showFilter(td);else closeFilter()});
        document.addEventListener('dblclick',e=>{const td=e.target.closest('td');if(td&&!td.classList.contains('fh'))post({ref:refOf(td),text:full(td),edit:'1'})});
        document.addEventListener('keydown',e=>{if(e.target.closest&&e.target.closest('#filter-menu')){if(e.key==='Escape')closeFilter();return}
        if(e.key==='Escape'){closeFilter();return}if(!current||e.metaKey||e.ctrlKey||e.altKey)return;
        if(e.key==='Enter'||e.key==='F2'){e.preventDefault();post({ref:refOf(current),text:full(current),edit:'1'});return}
        const m={ArrowUp:[0,-1],ArrowDown:[0,1],ArrowLeft:[-1,0],ArrowRight:[1,0],Tab:[e.shiftKey?-1:1,0]}[e.key];if(!m)return;e.preventDefault();const next=neighbor(current,m[0],m[1]);if(next)select(next,true)});
        </script></body></html>
        """
        return Rendered(html: html, rows: allRows.count, columns: columns.count, shownRows: rows.count)
    }
}

// MARK: - Formulas
// A small evaluator for the most common worksheet formulas. It is used for new formulas and after edits;
// anything it doesn't know (other functions, names, shared formulas) keeps Excel's saved value.

enum FValue {
    case number(Double), text(String), bool(Bool), error(String), empty, list([FValue])
    var number: Double? {
        switch self { case .number(let n): return n; case .bool(let b): return b ? 1 : 0; case .empty: return 0
        case .text(let s): return Double(s.trimmingCharacters(in: .whitespaces)); default: return nil }
    }
    var text: String {
        switch self { case .number(let n): return n == n.rounded() && abs(n) < 1e15 ? String(Int64(n)) : String(n); case .text(let s): return s
        case .bool(let b): return b ? "TRUE" : "FALSE"; case .error(let e): return e; case .empty: return ""; case .list(let l): return l.first?.text ?? "" }
    }
    var truthy: Bool? { if case .text(let s) = self { return ["TRUE", "VERO"].contains(s.uppercased()) ? true : (["FALSE", "FALSO"].contains(s.uppercased()) ? false : nil) }; return number.map { $0 != 0 } }
    var isError: Bool { if case .error = self { return true }; return false }
    var items: [FValue] { if case .list(let l) = self { return l }; return [self] }
}
struct FormulaUnsupported: Error {}
struct FormulaInvalid: Error {}

enum FToken: Equatable { case num(Double), str(String), ref(Int, Int, Int), name(String), op(String), lp, rp, comma, colon }

let italianFunctions = ["SOMMA": "SUM", "MEDIA": "AVERAGE", "CONTA.NUMERI": "COUNT", "CONTA.VALORI": "COUNTA", "ARROTONDA": "ROUND", "ARROTONDA.PER.ECC": "ROUNDUP",
    "ARROTONDA.PER.DIF": "ROUNDDOWN", "ASS": "ABS", "SE": "IF", "E": "AND", "O": "OR", "NON": "NOT", "SE.ERRORE": "IFERROR", "SOMMA.SE": "SUMIF", "CONTA.SE": "COUNTIF",
    "MEDIA.SE": "AVERAGEIF", "CONCATENA": "CONCATENATE", "LUNGHEZZA": "LEN", "MAIUSC": "UPPER", "MINUSC": "LOWER", "ANNULLA.SPAZI": "TRIM", "OGGI": "TODAY",
    "RADQ": "SQRT", "POTENZA": "POWER", "RESTO": "MOD", "INT": "INT", "MEDIANA": "MEDIAN", "PRODOTTO": "PRODUCT", "VERO": "TRUE", "FALSO": "FALSE"]

/// "=somma(A1;A2*1,5)" → "SUM(A1,A2*1.5)": English names and separators, as stored in the file.
func normalizeFormula(_ input: String) -> String {
    var s = input.trimmingCharacters(in: .whitespacesAndNewlines)
    if s.hasPrefix("=") { s.removeFirst() }
    var quote: Character? = nil, semicolons = false
    for ch in s { if ch == "\"" || ch == "'" { quote = quote == nil ? ch : (quote == ch ? nil : quote) } else if quote == nil && ch == ";" { semicolons = true } }
    var out = "", word = ""; quote = nil
    func flush(_ next: Character?) {
        guard !word.isEmpty else { return }
        let upper = word.uppercased()
        out += next == "(" || ["VERO", "FALSO"].contains(upper) ? (italianFunctions[upper] ?? upper) : upper
        word = ""
    }
    for ch in s {
        if let q = quote { out.append(ch); if ch == q { quote = nil }; continue }
        if ch == "\"" || ch == "'" { flush(ch); quote = ch; out.append(ch); continue }
        if ch.isLetter || ch.isNumber || ch == "." || ch == "_" || ch == "$" { word.append(ch); continue }
        if semicolons && ch == "," { word.append("."); continue }
        flush(ch)
        out.append(semicolons && ch == ";" ? "," : ch)
    }
    flush(nil)
    return out
}

final class FormulaParser {
    let book: Book, sheet: Int
    var tokens: [FToken] = [], pos = 0
    init(book: Book, sheet: Int) { self.book = book; self.sheet = sheet }

    func tokenize(_ s: String) throws {
        let chars = Array(s); var i = 0
        func refToken(_ word: String, sheet: Int) -> FToken? {
            let clean = word.replacingOccurrences(of: "$", with: "").uppercased()
            guard clean.range(of: "^[A-Z]{1,3}[0-9]{1,7}$", options: .regularExpression) != nil else { return nil }
            return .ref(sheet, rowNumber(clean), column(clean))
        }
        func sheetIndex(_ name: String) throws -> Int { guard let i = book.sheets.firstIndex(where: { $0.name.caseInsensitiveCompare(name) == .orderedSame }) else { throw FormulaUnsupported() }; return i }
        while i < chars.count {
            let ch = chars[i]
            if ch == " " { i += 1; continue }
            if ch == "\"" {
                var text = ""; i += 1
                while i < chars.count { if chars[i] == "\"" { if i + 1 < chars.count && chars[i + 1] == "\"" { text.append("\""); i += 2; continue }; break }; text.append(chars[i]); i += 1 }
                guard i < chars.count else { throw FormulaInvalid() }
                i += 1; tokens.append(.str(text)); continue
            }
            if ch == "'" {
                var name = ""; i += 1
                while i < chars.count { if chars[i] == "'" { if i + 1 < chars.count && chars[i + 1] == "'" { name.append("'"); i += 2; continue }; break }; name.append(chars[i]); i += 1 }
                guard i + 1 < chars.count, chars[i + 1] == "!" else { throw FormulaInvalid() }
                i += 2; var word = ""
                while i < chars.count, chars[i].isLetter || chars[i].isNumber || chars[i] == "$" { word.append(chars[i]); i += 1 }
                guard let t = refToken(word, sheet: try sheetIndex(name)) else { throw FormulaUnsupported() }
                tokens.append(t); continue
            }
            if ch.isNumber || (ch == "." && i + 1 < chars.count && chars[i + 1].isNumber) {
                var num = ""
                while i < chars.count, chars[i].isNumber || chars[i] == "." || ((chars[i] == "E" || chars[i] == "e") && i + 1 < chars.count && (chars[i + 1].isNumber || chars[i + 1] == "-" || chars[i + 1] == "+")) {
                    if chars[i] == "E" || chars[i] == "e" { num.append("e"); num.append(chars[i + 1]); i += 2; continue }
                    num.append(chars[i]); i += 1
                }
                guard let n = Double(num) else { throw FormulaInvalid() }
                tokens.append(.num(n)); continue
            }
            if ch.isLetter || ch == "$" || ch == "_" {
                var word = ""
                while i < chars.count, chars[i].isLetter || chars[i].isNumber || chars[i] == "." || chars[i] == "_" || chars[i] == "$" { word.append(chars[i]); i += 1 }
                if i < chars.count && chars[i] == "!" {
                    i += 1; var ref = ""
                    while i < chars.count, chars[i].isLetter || chars[i].isNumber || chars[i] == "$" { ref.append(chars[i]); i += 1 }
                    guard let t = refToken(ref, sheet: try sheetIndex(word)) else { throw FormulaUnsupported() }
                    tokens.append(t)
                } else if let t = refToken(word, sheet: sheet), !(i < chars.count && chars[i] == "(") { tokens.append(t) }
                else { tokens.append(.name(word.uppercased())) }
                continue
            }
            let two = i + 1 < chars.count ? String([ch, chars[i + 1]]) : ""
            if ["<>", "<=", ">="].contains(two) { tokens.append(.op(two)); i += 2; continue }
            switch ch {
            case "(": tokens.append(.lp); case ")": tokens.append(.rp); case ",": tokens.append(.comma); case ":": tokens.append(.colon)
            case "+", "-", "*", "/", "^", "&", "=", "<", ">", "%": tokens.append(.op(String(ch)))
            default: throw FormulaInvalid()
            }
            i += 1
        }
    }
    func evaluate(_ formula: String) throws -> FValue {
        tokens = []; pos = 0
        try tokenize(formula)
        let value = try comparison()
        guard pos == tokens.count else { throw FormulaInvalid() }
        if case .list(let l) = value { return l.first ?? .empty }
        return value
    }
    var peek: FToken? { pos < tokens.count ? tokens[pos] : nil }
    func take(_ t: FToken) -> Bool { if peek == t { pos += 1; return true }; return false }

    func comparison() throws -> FValue {
        var left = try concat()
        while case .op(let o)? = peek, ["=", "<>", "<", ">", "<=", ">="].contains(o) {
            pos += 1; let right = try concat()
            if left.isError { return left }; if right.isError { return right }
            left = .bool(compare(left, o, right))
        }
        return left
    }
    func concat() throws -> FValue {
        var left = try additive()
        while take(.op("&")) { let right = try additive(); if left.isError { return left }; if right.isError { return right }; left = .text(left.text + right.text) }
        return left
    }
    func arithmetic(_ a: FValue, _ b: FValue, _ f: (Double, Double) -> FValue) -> FValue {
        if a.isError { return a }; if b.isError { return b }
        guard let x = a.number, let y = b.number else { return .error("#VALUE!") }
        return f(x, y)
    }
    func additive() throws -> FValue {
        var left = try term()
        while case .op(let o)? = peek, o == "+" || o == "-" { pos += 1; let right = try term(); left = arithmetic(left, right) { .number(o == "+" ? $0 + $1 : $0 - $1) } }
        return left
    }
    func term() throws -> FValue {
        var left = try power()
        while case .op(let o)? = peek, o == "*" || o == "/" { pos += 1; let right = try power(); left = arithmetic(left, right) { o == "*" ? .number($0 * $1) : ($1 == 0 ? .error("#DIV/0!") : .number($0 / $1)) } }
        return left
    }
    func power() throws -> FValue {
        var left = try unary()
        while take(.op("^")) { let right = try unary(); left = arithmetic(left, right) { .number(Foundation.pow($0, $1)) } }
        return left
    }
    func unary() throws -> FValue {
        if take(.op("-")) { return arithmetic(.number(0), try unary()) { .number($0 - $1) } }
        if take(.op("+")) { return try unary() }
        var value = try primary()
        while take(.op("%")) { value = arithmetic(value, .number(100)) { .number($0 / $1) } }
        return value
    }
    func primary() throws -> FValue {
        guard let t = peek else { throw FormulaInvalid() }
        pos += 1
        switch t {
        case .num(let n): return .number(n)
        case .str(let s): return .text(s)
        case .lp: let v = try comparison(); guard take(.rp) else { throw FormulaInvalid() }; return v
        case .ref(let s, let r, let c):
            if take(.colon) {
                guard case .ref(_, let r2, let c2)? = peek else { throw FormulaInvalid() }
                pos += 1
                return .list(try book.rangeValues(s, min(r, r2), min(c, c2), max(r, r2), max(c, c2)))
            }
            return try book.value(s, r, c)
        case .name(let name):
            if name == "TRUE" || name == "FALSE" { _ = take(.lp) && take(.rp); return .bool(name == "TRUE") }
            guard take(.lp) else { throw FormulaUnsupported() }
            var args: [FValue] = []
            if !take(.rp) {
                repeat { args.append(peek == .comma || peek == .rp ? .empty : try comparison()) } while take(.comma)
                guard take(.rp) else { throw FormulaInvalid() }
            }
            return try call(italianFunctions[name] ?? name, args)
        default: throw FormulaInvalid()
        }
    }
    func compare(_ a: FValue, _ op: String, _ b: FValue) -> Bool {
        let order: Int
        if case .text(let x) = a, case .text(let y) = b { order = x.caseInsensitiveCompare(y).rawValue }
        else if let x = a.number, let y = b.number, !(a.isTextual || b.isTextual) { order = x < y ? -1 : (x > y ? 1 : 0) }
        else { order = a.text.caseInsensitiveCompare(b.text).rawValue }
        switch op { case "=": return order == 0; case "<>": return order != 0; case "<": return order < 0; case ">": return order > 0; case "<=": return order <= 0; default: return order >= 0 }
    }
    func matches(_ value: FValue, _ criterion: FValue) -> Bool {
        var text = criterion.text, op = "="
        for candidate in ["<>", "<=", ">=", "<", ">", "="] where text.hasPrefix(candidate) { op = candidate; text = String(text.dropFirst(candidate.count)); break }
        let target: FValue = Double(text).map { .number($0) } ?? .text(text)
        if case .empty = value { return op == "=" ? text.isEmpty : (op == "<>" ? !text.isEmpty : false) }
        return compare(value, op, target)
    }
    func numbers(_ args: [FValue]) throws -> [Double] {
        var out: [Double] = []
        for arg in args {
            if case .list(let l) = arg { for v in l { if case .error = v { throw FormulaErrorValue(v) }; if case .number(let n) = v { out.append(n) } } }
            else if arg.isError { throw FormulaErrorValue(arg) }
            else if case .empty = arg { continue }
            else if let n = arg.number { out.append(n) } else { throw FormulaErrorValue(.error("#VALUE!")) }
        }
        return out
    }
    func call(_ name: String, _ args: [FValue]) throws -> FValue {
        func arg(_ i: Int) -> FValue { args.indices.contains(i) ? args[i] : .empty }
        func num(_ i: Int) throws -> Double { guard let n = arg(i).number else { throw FormulaErrorValue(arg(i).isError ? arg(i) : .error("#VALUE!")) }; return n }
        do {
            switch name {
            case "SUM": return .number(try numbers(args).reduce(0, +))
            case "PRODUCT": return .number(try numbers(args).reduce(1, *))
            case "AVERAGE": let n = try numbers(args); return n.isEmpty ? .error("#DIV/0!") : .number(n.reduce(0, +) / Double(n.count))
            case "MIN": return .number(try numbers(args).min() ?? 0)
            case "MAX": return .number(try numbers(args).max() ?? 0)
            case "MEDIAN": let n = try numbers(args).sorted(); guard !n.isEmpty else { return .error("#NUM!") }; return .number(n.count % 2 == 1 ? n[n.count / 2] : (n[n.count / 2 - 1] + n[n.count / 2]) / 2)
            case "COUNT": return .number(Double(args.flatMap(\.items).filter { if case .number = $0 { return true }; return false }.count))
            case "COUNTA": return .number(Double(args.flatMap(\.items).filter { if case .empty = $0 { return false }; return true }.count))
            case "ROUND", "ROUNDUP", "ROUNDDOWN":
                let x = try num(0), d = args.count > 1 ? try num(1) : 0, m = Foundation.pow(10, d.rounded(.towardZero))
                let scaled = x * m
                let r = name == "ROUND" ? scaled.rounded(.toNearestOrAwayFromZero) : (name == "ROUNDUP" ? (scaled < 0 ? scaled.rounded(.down) : scaled.rounded(.up)) : scaled.rounded(.towardZero))
                return .number(r / m)
            case "INT": return .number(try num(0).rounded(.down))
            case "ABS": return .number(abs(try num(0)))
            case "SQRT": let x = try num(0); return x < 0 ? .error("#NUM!") : .number(x.squareRoot())
            case "POWER": return .number(Foundation.pow(try num(0), try num(1)))
            case "MOD": let d = try num(1); guard d != 0 else { return .error("#DIV/0!") }; let x = try num(0); return .number(x - d * (x / d).rounded(.down))
            case "IF": if arg(0).isError { return arg(0) }; guard let c = arg(0).truthy else { return .error("#VALUE!") }; return c ? arg(1) : (args.count > 2 ? arg(2) : .bool(false))
            case "IFERROR": return arg(0).isError ? arg(1) : arg(0)
            case "AND", "OR":
                let flags = args.flatMap(\.items).compactMap { v -> Bool? in if case .empty = v { return nil }; return v.truthy }
                return .bool(name == "AND" ? !flags.contains(false) : flags.contains(true))
            case "NOT": guard let c = arg(0).truthy else { return .error("#VALUE!") }; return .bool(!c)
            case "CONCATENATE", "CONCAT": return .text(args.flatMap(\.items).map(\.text).joined())
            case "LEN": return .number(Double(arg(0).text.count))
            case "UPPER": return .text(arg(0).text.uppercased())
            case "LOWER": return .text(arg(0).text.lowercased())
            case "TRIM": return .text(arg(0).text.split(separator: " ").joined(separator: " "))
            case "TODAY": return .number((Date().timeIntervalSince1970 / 86400).rounded(.down) + 25569)
            case "SUMIF", "COUNTIF", "AVERAGEIF":
                let range = arg(0).items, values = args.count > 2 ? arg(2).items : range
                let picked = range.indices.filter { matches(range[$0], arg(1)) }
                if name == "COUNTIF" { return .number(Double(picked.count)) }
                let nums = picked.compactMap { i -> Double? in guard values.indices.contains(i), case .number(let n) = values[i] else { return nil }; return n }
                if name == "SUMIF" { return .number(nums.reduce(0, +)) }
                return nums.isEmpty ? .error("#DIV/0!") : .number(nums.reduce(0, +) / Double(nums.count))
            default: throw FormulaUnsupported()
            }
        } catch let error as FormulaErrorValue { return error.value }
    }
}
struct FormulaErrorValue: Error { let value: FValue; init(_ v: FValue) { value = v } }
extension FValue { var isTextual: Bool { if case .text = self { return true }; return false } }

func welcomePage() -> String {
    """
    <!doctype html><html><head><meta charset="utf-8"><style>:root{color-scheme:light dark}html,body{height:100%;margin:0}
    body{display:flex;align-items:center;justify-content:center;font:13px -apple-system,system-ui;color:#86868b;background:#fff;-webkit-user-select:none}
    @media(prefers-color-scheme:dark){body{background:#1e1e1e;color:#98989d}h1{color:#f5f5f7!important}.box{border-color:#3a3a3c!important}}
    .box{text-align:center;padding:44px 64px;border:2px dashed #d2d2d7;border-radius:18px}h1{font-size:21px;font-weight:500;color:#1d1d1f;margin:14px 0 8px}
    svg{width:52px;height:52px;opacity:.55}</style></head><body><div class="box">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16M15 4v16"/></svg>
    <h1>\(L("Apri un foglio di calcolo", "Open a spreadsheet"))</h1><div>\(L("Trascina qui un file .xlsx oppure premi ⌘O", "Drop an .xlsx file here or press ⌘O"))</div></div></body></html>
    """
}

final class SheetWebView: WKWebView {
    var openURL: ((URL) -> Void)?
    private func fileURL(_ info: NSDraggingInfo) -> URL? {
        let urls = info.draggingPasteboard.readObjects(forClasses: [NSURL.self], options: [.urlReadingFileURLsOnly: true]) as? [URL] ?? []
        return urls.first { supportedExtensions.contains($0.pathExtension.lowercased()) }
    }
    override func draggingEntered(_ sender: NSDraggingInfo) -> NSDragOperation { fileURL(sender) != nil ? .copy : [] }
    override func draggingUpdated(_ sender: NSDraggingInfo) -> NSDragOperation { fileURL(sender) != nil ? .copy : [] }
    override func prepareForDragOperation(_ sender: NSDraggingInfo) -> Bool { fileURL(sender) != nil }
    override func performDragOperation(_ sender: NSDraggingInfo) -> Bool {
        guard let url = fileURL(sender) else { return false }
        DispatchQueue.main.async { self.openURL?(url) }
        return true
    }
}

final class App: NSObject, NSApplicationDelegate, WKNavigationDelegate, WKScriptMessageHandler, NSWindowDelegate, NSMenuItemValidation {
    var window: NSWindow!; var web: SheetWebView!; var picker: NSPopUpButton!; var status: NSTextField!; var book: Book?
    var zoomLabel: NSButton!; var search: NSSearchField!; var detail: NSTextView!; var notice: NSTextField!
    let progress = NSProgressIndicator()
    let recentMenu = NSMenu(title: L("Apri recenti", "Open Recent"))
    var selectedRef = ""
    let editUndo = UndoManager()
    var zoom = 0.85
    var pendingURL: URL?
    var loadID = UUID()
    var loading = false
    let loader = DispatchQueue(label: "xlsxlite.load", qos: .userInitiated)
    var hint: String { L("Invio o doppio clic: modifica · Frecce: muoviti · Intestazioni ▾: filtra", "Return or double-click: edit · Arrows: move · Headers ▾: filter") }

    func updateEditState() {
        window.isDocumentEdited = book?.dirty ?? false
        notice.stringValue = (book?.dirty == true) ? L("Modifiche non salvate · Le formule si ricalcolano quando apri il file in Excel", "Unsaved changes · Formulas recalculate when the file is opened in Excel") : hint
    }
    @objc func clearFilters() { web.evaluateJavaScript("window.clearFilters&&clearFilters()", completionHandler: nil) }
    @objc func undoEdit() { if let modal = NSApp.modalWindow { modal.firstResponder?.undoManager?.undo(); return }; editUndo.undo(); updateEditState() }
    @objc func redoEdit() { if let modal = NSApp.modalWindow { modal.firstResponder?.undoManager?.redo(); return }; editUndo.redo(); updateEditState() }
    func validateMenuItem(_ item: NSMenuItem) -> Bool {
        switch item.action {
        case #selector(undoEdit): return NSApp.modalWindow != nil || editUndo.canUndo
        case #selector(redoEdit): return NSApp.modalWindow != nil || editUndo.canRedo
        case #selector(saveCopy), #selector(printSheet), #selector(clearFilters), #selector(nextSheet), #selector(previousSheet): return book != nil && !loading
        case #selector(editCell): return book != nil && !selectedRef.isEmpty && !loading
        default: return true
        }
    }
    func restoreSheet(_ index: Int, _ data: Data, dirty: Bool) {
        guard let book else { return }
        do {
            let old = try book.sheetDocument(index).xmlData, wasDirty = book.dirty
            editUndo.registerUndo(withTarget: self) { $0.restoreSheet(index, old, dirty: wasDirty) }
            book.documents[book.sheets[index].path] = try XMLDocument(data: data, options: [.nodePreserveWhitespace]); book.dirty = dirty; book.invalidate()
            picker.selectItem(at: index); changeSheet(); updateEditState()
        } catch { show(error) }
    }
    @objc func editCell() {
        guard let book, !selectedRef.isEmpty, !loading else { return }
        do {
            let index = picker.indexOfSelectedItem, ref = selectedRef, cell = try book.cell(index, ref)
            let alert = NSAlert(); alert.messageText = L("Modifica \(ref)", "Edit \(ref)")
            alert.informativeText = cell?.child("f") != nil
                ? L("Questa cella contiene una formula: applicando la modifica la sostituisci con un valore. Le altre formule si ricalcolano in Excel.", "This cell contains a formula: applying replaces it with a value. Other formulas recalculate in Excel.")
                : L("Le formule che dipendono da questa cella si ricalcolano quando apri il file salvato in Excel.", "Formulas that depend on this cell recalculate when the saved file is opened in Excel.")
            alert.addButton(withTitle: L("Applica", "Apply")); alert.addButton(withTitle: L("Annulla", "Cancel"))
            let accessory = NSView(frame: NSRect(x: 0, y: 0, width: 440, height: 160))
            let kind = NSPopUpButton(frame: NSRect(x: 0, y: 132, width: 440, height: 26))
            kind.addItems(withTitles: [L("Testo", "Text"), L("Numero (decimali con virgola o punto)", "Number"), L("Formula (es. =SOMMA(A1:A10) o =SUM(A1:A10))", "Formula (e.g. =SUM(A1:A10))")]); accessory.addSubview(kind)
            let scroll = NSScrollView(frame: NSRect(x: 0, y: 0, width: 440, height: 122)); scroll.borderType = .bezelBorder; scroll.hasVerticalScroller = true
            let input = NSTextView(frame: scroll.bounds); input.isRichText = false; input.font = .systemFont(ofSize: 14); input.isVerticallyResizable = true; input.autoresizingMask = [.width]; input.textContainer?.widthTracksTextView = true; input.allowsUndo = true
            input.isAutomaticQuoteSubstitutionEnabled = false; input.isAutomaticDashSubstitutionEnabled = false; input.isAutomaticTextReplacementEnabled = false
            if let cell, let formula = book.formulaText(cell) { kind.selectItem(at: 2); input.string = "=" + formula }
            else if let cell, book.isNumeric(cell) { kind.selectItem(at: 1); input.string = cell.child("v")?.stringValue ?? "" }
            else { input.string = cell.map { book.display($0, sheet: index) } ?? "" }
            input.selectAll(nil)
            scroll.documentView = input; accessory.addSubview(scroll); alert.accessoryView = accessory; alert.window.initialFirstResponder = input
            while alert.runModal() == .alertFirstButtonReturn {
                var value = input.string
                var cellKind: CellKind = [.text, .number, .formula][kind.indexOfSelectedItem]
                if value.hasPrefix("=") && value.count > 1 { cellKind = .formula }
                if cellKind == .formula {
                    // Reject typos now; formulas with functions this app doesn't know are kept for Excel.
                    do { _ = try FormulaParser(book: book, sheet: index).evaluate(normalizeFormula(value)) }
                    catch is FormulaInvalid { alert.informativeText = L("La formula non è valida. Esempi: =A1+B1, =SOMMA(A1:A10), =SE(B2>100;\"sì\";\"no\")", "The formula is not valid. Examples: =A1+B1, =SUM(A1:A10), =IF(B2>100,\"yes\",\"no\")"); continue }
                    catch {}
                } else if cellKind == .number && !value.isEmpty {
                    value = value.trimmingCharacters(in: .whitespacesAndNewlines).replacingOccurrences(of: ",", with: ".")
                    guard let n = Double(value), n.isFinite else { alert.informativeText = L("Inserisci un numero senza separatori delle migliaia (es. 50000 o 3,65).", "Enter a number without thousands separators (e.g. 50000 or 3.65)."); continue }
                    value = n == n.rounded() && abs(n) < 1e15 ? String(Int64(n)) : String(n)
                }
                let old = try book.sheetDocument(index).xmlData, wasDirty = book.dirty
                try book.setCell(index, ref, value, kind: cellKind)
                editUndo.registerUndo(withTarget: self) { $0.restoreSheet(index, old, dirty: wasDirty) }; editUndo.setActionName(L("modifica cella", "Cell Edit"))
                refreshFormulas(index, select: ref)
                updateEditState(); break
            }
            window.makeFirstResponder(web)
        } catch { show(error) }
    }
    /// Pushes the edited cell and every recalculated formula on the sheet to the page, keeping scroll and filters.
    func refreshFormulas(_ index: Int, select ref: String) {
        guard let book, let sheetIndex = try? book.index(index) else { return }
        var updates: [[Any]] = []
        for (key, cell) in sheetIndex.cells where cell.child("f") != nil { updates.append([key / 16_385, key % 16_385, book.display(cell, sheet: index), 0]) }
        let edited = (try? book.cell(index, ref)).flatMap { $0 }.map { book.display($0, sheet: index) } ?? ""
        updates.append([rowNumber(ref), column(ref), edited, 1])
        guard let data = try? JSONSerialization.data(withJSONObject: updates), let json = String(data: data, encoding: .utf8) else { return }
        web.evaluateJavaScript("setCells(\(json))", completionHandler: nil)
        showDetail(ref, edited)
    }
    func showDetail(_ ref: String, _ text: String) {
        var line = "\(ref)   \(text)"
        if let book, let cell = (try? book.cell(picker.indexOfSelectedItem, ref)).flatMap({ $0 }), let formula = book.formulaText(cell) { line += "      =" + formula }
        detail.string = line
    }
    @discardableResult @objc func saveCopy() -> Bool {
        guard let book else { return false }
        let ext = book.url.pathExtension.lowercased()
        let panel = NSSavePanel(); panel.allowedContentTypes = [UTType(filenameExtension: ext) ?? .data]
        panel.nameFieldStringValue = book.url.deletingPathExtension().lastPathComponent + L(" - modificato.", " - edited.") + ext
        panel.directoryURL = book.url.deletingLastPathComponent(); panel.title = L("Salva una copia", "Save a Copy")
        guard panel.runModal() == .OK, let target = panel.url else { return false }
        do {
            try book.save(to: target)
            let index = picker.indexOfSelectedItem
            let newBook = try Book(target); self.book = newBook; editUndo.removeAllActions()
            window.title = target.lastPathComponent; window.representedURL = target
            UserDefaults.standard.set(target.path, forKey: "lastFile"); NSDocumentController.shared.noteNewRecentDocumentURL(target); refreshRecentMenu()
            picker.selectItem(at: index); changeSheet(); updateEditState(); return true
        } catch { show(error); return false }
    }
    func confirmUnsaved() -> Bool {
        guard book?.dirty == true else { return true }
        let alert = NSAlert(); alert.messageText = L("Salvare le modifiche?", "Save your changes?"); alert.informativeText = L("Ci sono modifiche non salvate in questo file.", "This file has unsaved changes.")
        alert.addButton(withTitle: L("Salva una copia…", "Save a Copy…")); alert.addButton(withTitle: L("Annulla", "Cancel")); alert.addButton(withTitle: L("Non salvare", "Don't Save"))
        switch alert.runModal() { case .alertFirstButtonReturn: return saveCopy(); case .alertThirdButtonReturn: return true; default: return false }
    }
    func windowShouldClose(_ sender: NSWindow) -> Bool { let close = confirmUnsaved(); if close { book?.dirty = false }; return close }
    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply { confirmUnsaved() ? .terminateNow : .terminateCancel }

    func applyZoom(_ value: Double) {
        zoom = min(2, max(0.5, value)); web.pageZoom = zoom; web.magnification = 1
        zoomLabel.title = "\(Int((zoom * 100).rounded()))%"
        UserDefaults.standard.set(zoom, forKey: "pageZoom")
    }
    @objc func zoomIn() { applyZoom(zoom + 0.1) }
    @objc func zoomOut() { applyZoom(zoom - 0.1) }
    @objc func resetZoom() { applyZoom(1) }
    @objc func focusSearch() { window.makeFirstResponder(search) }
    @objc func findNext() { find(backwards: false) }
    @objc func findPrevious() { find(backwards: true) }
    func find(backwards: Bool) {
        guard !search.stringValue.isEmpty, book != nil else { return }
        let config = WKFindConfiguration(); config.wraps = true; config.backwards = backwards; config.caseSensitive = false
        web.find(search.stringValue, configuration: config) { result in
            self.notice.stringValue = result.matchFound ? self.hint : L("Nessun risultato in questo foglio", "No results in this sheet")
            if !result.matchFound { NSSound.beep() }
        }
    }
    @objc func nextSheet() { moveSheet(1) }
    @objc func previousSheet() { moveSheet(-1) }
    func moveSheet(_ step: Int) {
        guard book != nil, picker.numberOfItems > 1 else { return }
        picker.selectItem(at: (picker.indexOfSelectedItem + step + picker.numberOfItems) % picker.numberOfItems); changeSheet()
    }
    @objc func printSheet() {
        guard let book, !loading else { return }
        let info = NSPrintInfo.shared.copy() as! NSPrintInfo
        info.orientation = .landscape; info.horizontalPagination = .fit; info.verticalPagination = .automatic
        info.leftMargin = 24; info.rightMargin = 24; info.topMargin = 24; info.bottomMargin = 24
        let operation = web.printOperation(with: info)
        operation.view?.frame = web.bounds
        operation.jobTitle = book.url.deletingPathExtension().lastPathComponent + " – " + (picker.titleOfSelectedItem ?? "")
        operation.runModal(for: window, delegate: nil, didRun: nil, contextInfo: nil)
    }
    func userContentController(_ userContentController: WKUserContentController, didReceive message: WKScriptMessage) {
        guard let data = message.body as? [String: String] else { return }
        selectedRef = data["ref"] ?? ""
        showDetail(selectedRef, data["text"] ?? "")
        if data["edit"] == "1" { DispatchQueue.main.async { self.editCell() } }
    }

    func buildMenu() {
        let main = NSMenu()
        func submenu(_ title: String) -> NSMenu { let item = NSMenuItem(title: title, action: nil, keyEquivalent: ""); let sub = NSMenu(title: title); item.submenu = sub; main.addItem(item); return sub }
        func add(_ menu: NSMenu, _ title: String, _ action: Selector, _ key: String = "", _ mods: NSEvent.ModifierFlags = .command, target: AnyObject? = nil) {
            let item = menu.addItem(withTitle: title, action: action, keyEquivalent: key); item.keyEquivalentModifierMask = mods; item.target = target
        }
        let app = submenu(appName)
        add(app, L("Informazioni su \(appName)", "About \(appName)"), #selector(NSApplication.orderFrontStandardAboutPanel(_:)))
        app.addItem(.separator())
        add(app, L("Nascondi \(appName)", "Hide \(appName)"), #selector(NSApplication.hide(_:)), "h")
        add(app, L("Nascondi altre", "Hide Others"), #selector(NSApplication.hideOtherApplications(_:)), "h", [.command, .option])
        app.addItem(.separator())
        add(app, L("Esci da \(appName)", "Quit \(appName)"), #selector(NSApplication.terminate(_:)), "q")
        let file = submenu("File")
        add(file, L("Apri…", "Open…"), #selector(openFile), "o", target: self)
        let recent = NSMenuItem(title: recentMenu.title, action: nil, keyEquivalent: ""); recent.submenu = recentMenu; file.addItem(recent); refreshRecentMenu()
        file.addItem(.separator())
        add(file, L("Chiudi", "Close"), #selector(NSWindow.performClose(_:)), "w")
        add(file, L("Salva una copia…", "Save a Copy…"), #selector(saveCopy), "s", target: self)
        file.addItem(.separator())
        add(file, L("Stampa…", "Print…"), #selector(printSheet), "p", target: self)
        let edit = submenu(L("Modifica", "Edit"))
        add(edit, L("Annulla", "Undo"), #selector(undoEdit), "z", target: self)
        add(edit, L("Ripristina", "Redo"), #selector(redoEdit), "z", [.command, .shift], target: self)
        edit.addItem(.separator())
        add(edit, L("Taglia", "Cut"), #selector(NSText.cut(_:)), "x")
        add(edit, L("Copia", "Copy"), #selector(NSText.copy(_:)), "c")
        add(edit, L("Incolla", "Paste"), #selector(NSText.paste(_:)), "v")
        add(edit, L("Seleziona tutto", "Select All"), #selector(NSText.selectAll(_:)), "a")
        edit.addItem(.separator())
        add(edit, L("Modifica cella…", "Edit Cell…"), #selector(editCell), target: self)
        add(edit, L("Rimuovi filtri", "Clear Filters"), #selector(clearFilters), target: self)
        edit.addItem(.separator())
        add(edit, L("Cerca…", "Find…"), #selector(focusSearch), "f", target: self)
        add(edit, L("Trova successivo", "Find Next"), #selector(findNext), "g", target: self)
        add(edit, L("Trova precedente", "Find Previous"), #selector(findPrevious), "g", [.command, .shift], target: self)
        let view = submenu(L("Vista", "View"))
        add(view, L("Ingrandisci", "Zoom In"), #selector(zoomIn), "+", target: self)
        add(view, L("Riduci", "Zoom Out"), #selector(zoomOut), "-", target: self)
        add(view, L("Dimensioni reali", "Actual Size"), #selector(resetZoom), "0", target: self)
        view.addItem(.separator())
        add(view, L("Foglio successivo", "Next Sheet"), #selector(nextSheet), "]", [.command, .option], target: self)
        add(view, L("Foglio precedente", "Previous Sheet"), #selector(previousSheet), "[", [.command, .option], target: self)
        view.addItem(.separator())
        add(view, L("Schermo intero", "Enter Full Screen"), #selector(NSWindow.toggleFullScreen(_:)), "f", [.command, .control])
        let windowMenu = submenu(L("Finestra", "Window"))
        add(windowMenu, L("Contrai", "Minimize"), #selector(NSWindow.performMiniaturize(_:)), "m")
        add(windowMenu, L("Ridimensiona", "Zoom"), #selector(NSWindow.performZoom(_:)))
        NSApp.windowsMenu = windowMenu
        NSApp.mainMenu = main
    }
    func refreshRecentMenu() {
        recentMenu.removeAllItems()
        for url in NSDocumentController.shared.recentDocumentURLs.prefix(10) {
            let item = recentMenu.addItem(withTitle: url.lastPathComponent, action: #selector(openRecent(_:)), keyEquivalent: "")
            item.target = self; item.representedObject = url; item.toolTip = url.path
        }
        if recentMenu.items.isEmpty { recentMenu.addItem(withTitle: L("Nessun file recente", "No Recent Files"), action: nil, keyEquivalent: "").isEnabled = false }
        else { recentMenu.addItem(.separator()); recentMenu.addItem(withTitle: L("Cancella elenco", "Clear Menu"), action: #selector(clearRecent), keyEquivalent: "").target = self }
    }
    @objc func openRecent(_ sender: NSMenuItem) { if let url = sender.representedObject as? URL { load(url) } }
    @objc func clearRecent() { NSDocumentController.shared.clearRecentDocuments(nil); refreshRecentMenu() }

    func applicationDidFinishLaunching(_ notification: Notification) {
        buildMenu()
        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 1180, height: 760), styleMask: [.titled, .closable, .miniaturizable, .resizable], backing: .buffered, defer: false)
        window.delegate = self; window.title = appName; window.minSize = NSSize(width: 900, height: 400); window.isReleasedWhenClosed = false
        window.center(); window.setFrameAutosaveName("ReaderWindow")
        let content = window.contentView!
        let open = NSButton(title: L("Apri…", "Open…"), target: self, action: #selector(openFile)); open.toolTip = L("Apri un file .xlsx (⌘O)", "Open an .xlsx file (⌘O)")
        progress.style = .spinning; progress.controlSize = .small; progress.isDisplayedWhenStopped = false
        picker = NSPopUpButton(); picker.target = self; picker.action = #selector(changeSheet); picker.toolTip = L("Foglio", "Sheet")
        status = NSTextField(labelWithString: ""); status.textColor = .secondaryLabelColor; status.font = .systemFont(ofSize: 11)
        search = NSSearchField(); search.placeholderString = L("Cerca nel foglio", "Search sheet"); search.target = self; search.action = #selector(findNext); search.sendsSearchStringImmediately = false; search.sendsWholeSearchString = true
        let minus = NSButton(title: "−", target: self, action: #selector(zoomOut)); minus.toolTip = L("Riduci (⌘−)", "Zoom out (⌘−)")
        let plus = NSButton(title: "+", target: self, action: #selector(zoomIn)); plus.toolTip = L("Ingrandisci (⌘+)", "Zoom in (⌘+)")
        zoomLabel = NSButton(title: "85%", target: self, action: #selector(resetZoom)); zoomLabel.toolTip = L("Dimensioni reali (⌘0)", "Actual size (⌘0)")
        let editButton = NSButton(title: L("Modifica", "Edit"), target: self, action: #selector(editCell)); editButton.toolTip = L("Modifica la cella selezionata (Invio)", "Edit the selected cell (Return)")
        let saveButton = NSButton(title: L("Salva…", "Save…"), target: self, action: #selector(saveCopy)); saveButton.toolTip = L("Salva una copia (⌘S)", "Save a copy (⌘S)")
        let resetFilters = NSButton(title: L("Rimuovi filtri", "Clear filters"), target: self, action: #selector(clearFilters))
        let spacer = NSView(); spacer.setContentHuggingPriority(.defaultLow, for: .horizontal)
        let bar = NSStackView(views: [open, progress, picker, editButton, saveButton, resetFilters, spacer, search, minus, zoomLabel, plus]); bar.spacing = 8; bar.orientation = .horizontal
        let footer = NSScrollView(); footer.hasVerticalScroller = true; footer.borderType = .noBorder
        detail = NSTextView(); detail.isEditable = false; detail.isSelectable = true; detail.font = .systemFont(ofSize: 12); detail.textColor = .secondaryLabelColor; detail.backgroundColor = .windowBackgroundColor; detail.textContainerInset = NSSize(width: 12, height: 6)
        detail.isVerticallyResizable = true; detail.isHorizontallyResizable = false; detail.autoresizingMask = [.width]; detail.textContainer?.widthTracksTextView = true
        detail.string = L("Seleziona una cella per leggerne il contenuto completo", "Select a cell to read its full content")
        footer.documentView = detail
        notice = NSTextField(labelWithString: hint); notice.font = .systemFont(ofSize: 11); notice.textColor = .secondaryLabelColor; notice.lineBreakMode = .byTruncatingTail
        notice.setContentCompressionResistancePriority(.defaultLow, for: .horizontal)
        let statusBar = NSStackView(views: [status, NSView(), notice]); statusBar.orientation = .horizontal; statusBar.spacing = 12
        let config = WKWebViewConfiguration(); config.userContentController.add(self, name: "cell")
        web = SheetWebView(frame: .zero, configuration: config); web.navigationDelegate = self; web.allowsMagnification = false
        web.openURL = { [weak self] url in self?.load(url) }
        let savedZoom = UserDefaults.standard.double(forKey: "pageZoom"); applyZoom(savedZoom > 0 ? savedZoom : 0.85)
        for view in [bar, web!, footer, statusBar] { view.translatesAutoresizingMaskIntoConstraints = false; content.addSubview(view) }
        NSLayoutConstraint.activate([
            bar.topAnchor.constraint(equalTo: content.topAnchor, constant: 12), bar.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 16), bar.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -16), bar.heightAnchor.constraint(equalToConstant: 30),
            web.topAnchor.constraint(equalTo: bar.bottomAnchor, constant: 12), web.leadingAnchor.constraint(equalTo: content.leadingAnchor), web.trailingAnchor.constraint(equalTo: content.trailingAnchor), web.bottomAnchor.constraint(equalTo: footer.topAnchor),
            footer.leadingAnchor.constraint(equalTo: content.leadingAnchor), footer.trailingAnchor.constraint(equalTo: content.trailingAnchor), footer.bottomAnchor.constraint(equalTo: statusBar.topAnchor, constant: -2), footer.heightAnchor.constraint(equalToConstant: 46),
            statusBar.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 12), statusBar.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -12), statusBar.bottomAnchor.constraint(equalTo: content.bottomAnchor, constant: -6), statusBar.heightAnchor.constraint(equalToConstant: 16),
            picker.widthAnchor.constraint(equalToConstant: 170), search.widthAnchor.constraint(equalToConstant: 170), zoomLabel.widthAnchor.constraint(equalToConstant: 54)])
        web.loadHTMLString(welcomePage(), baseURL: nil)
        window.makeKeyAndOrderFront(nil); NSApp.activate(ignoringOtherApps: true)
        if let url = pendingURL { pendingURL = nil; load(url) }
        else if let saved = UserDefaults.standard.string(forKey: "lastFile"), FileManager.default.fileExists(atPath: saved) { load(URL(fileURLWithPath: saved)) }
    }
    func application(_ sender: NSApplication, openFiles filenames: [String]) {
        if let f = filenames.first { let url = URL(fileURLWithPath: f); if window == nil { pendingURL = url } else { load(url) } }
        sender.reply(toOpenOrPrint: .success)
    }
    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { true }
    @objc func openFile() {
        let panel = NSOpenPanel(); panel.allowedContentTypes = supportedExtensions.compactMap { UTType(filenameExtension: $0) }; panel.canChooseDirectories = false; panel.allowsMultipleSelection = false
        if let url = book?.url { panel.directoryURL = url.deletingLastPathComponent() }
        if panel.runModal() == .OK, let url = panel.url { load(url) }
    }
    func setLoading(_ on: Bool, _ message: String = "") {
        loading = on
        if on { progress.startAnimation(nil); status.stringValue = message } else { progress.stopAnimation(nil) }
    }
    func load(_ url: URL) {
        guard confirmUnsaved() else { return }
        let request = UUID(); loadID = request
        setLoading(true, L("Apertura di \(url.lastPathComponent)…", "Opening \(url.lastPathComponent)…"))
        loader.async {
            let result = Result { () -> (Book, Rendered) in let book = try Book(url); return (book, try book.render(book.activeSheet)) }
            DispatchQueue.main.async {
                guard self.loadID == request else { return }
                self.setLoading(false)
                switch result {
                case .success(let (newBook, rendered)):
                    self.book = newBook; self.editUndo.removeAllActions(); self.selectedRef = ""
                    self.picker.removeAllItems()
                    for sheet in newBook.sheets { self.picker.addItem(withTitle: sheet.name + (sheet.hidden ? L(" (nascosto)", " (hidden)") : "")) }
                    self.picker.selectItem(at: newBook.activeSheet)
                    self.window.title = url.lastPathComponent; self.window.representedURL = url
                    UserDefaults.standard.set(url.path, forKey: "lastFile")
                    NSDocumentController.shared.noteNewRecentDocumentURL(url); self.refreshRecentMenu()
                    self.show(rendered); self.updateEditState()
                case .failure(let error):
                    self.status.stringValue = self.book == nil ? "" : self.status.stringValue
                    let alert = NSAlert(); alert.messageText = L("Impossibile aprire \(url.lastPathComponent)", "Can't open \(url.lastPathComponent)"); alert.informativeText = error.localizedDescription
                    alert.beginSheetModal(for: self.window)
                }
            }
        }
    }
    func show(_ rendered: Rendered) {
        let number = { (n: Int) in NumberFormatter.localizedString(from: NSNumber(value: n), number: .decimal) }
        var text = L("\(number(rendered.rows)) righe × \(rendered.columns) colonne", "\(number(rendered.rows)) rows × \(rendered.columns) columns")
        if let book, book.sheets.count > 1 { text = L("\(book.sheets.count) fogli · ", "\(book.sheets.count) sheets · ") + text }
        if rendered.shownRows < rendered.rows { text += L(" · mostrate le prime \(number(rendered.shownRows))", " · showing the first \(number(rendered.shownRows))") }
        status.stringValue = text
        detail.string = L("Seleziona una cella per leggerne il contenuto completo", "Select a cell to read its full content")
        web.loadHTMLString(rendered.html, baseURL: nil)
        window.makeFirstResponder(web)
    }
    @objc func changeSheet() {
        guard let book else { return }
        let index = picker.indexOfSelectedItem, request = UUID(); loadID = request; selectedRef = ""
        setLoading(true, L("Caricamento del foglio…", "Loading sheet…"))
        loader.async {
            let result = Result { try book.render(index) }
            DispatchQueue.main.async {
                guard self.loadID == request else { return }
                self.setLoading(false)
                switch result { case .success(let rendered): self.show(rendered); case .failure(let error): self.show(error) }
            }
        }
    }
    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) { webView.pageZoom = zoom }
    func show(_ error: Error) { NSAlert(error: error).runModal() }
    func webView(_ webView: WKWebView, decidePolicyFor action: WKNavigationAction, decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        if let url = action.request.url, url.isFileURL, supportedExtensions.contains(url.pathExtension.lowercased()) { decisionHandler(.cancel); load(url); return }
        decisionHandler(action.navigationType == .other ? .allow : .cancel)
    }
}

if CommandLine.arguments.count > 2 && CommandLine.arguments[1] == "--check" {
    do {
        let start = Date(), book = try Book(URL(fileURLWithPath: CommandLine.arguments[2]))
        for (i, s) in book.sheets.enumerated() {
            let r = try book.render(i)
            print("OK \(s.name): \(r.rows) rows × \(r.columns) cols, \(r.html.utf8.count / 1024) KB HTML")
            if CommandLine.arguments.count > 3 { try r.html.write(toFile: CommandLine.arguments[3] + "/sheet\(i).html", atomically: true, encoding: .utf8) }
        }
        print(String(format: "Time: %.2f s", Date().timeIntervalSince(start)))
    } catch { fputs("\(error.localizedDescription)\n", stderr); exit(1) }
} else {
    let app = NSApplication.shared; let delegate = App(); app.delegate = delegate; app.setActivationPolicy(.regular)
    if CommandLine.arguments.count > 1, !CommandLine.arguments[1].hasPrefix("-") { delegate.pendingURL = URL(fileURLWithPath: CommandLine.arguments[1]) }
    #if DEBUG
    if CommandLine.arguments.count > 2, CommandLine.arguments[1] == "--smoke-test" { scheduleSheetTests(delegate, fixture: URL(fileURLWithPath: CommandLine.arguments[2])) }
    #endif
    app.run()
}
