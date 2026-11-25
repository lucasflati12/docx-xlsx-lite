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
    func set(_ name: String, _ value: String) { removeAttribute(forName:name); addAttribute(XMLNode.attribute(withName:name,stringValue:value) as! XMLNode) }
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
    var numberFormat = "General", foreground = "#000000", background = "#ffffff", alignment = "left"
    var bold = false, italic = false, strike = false, wrap = false, explicitAlignment = false, fontSize = 11.0

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

enum StyleChange {
    case number(String), decimals(Int), flag(String,Bool), fontSize(Double), fontDelta(Double)
    case foreground(String), fill(String), alignment(String), wrap(Bool), copy(Int)
}
func styleSignature(_ e: XMLElement) -> String {
    let attrs = (e.attributes ?? []).map { ($0.name ?? "")+"="+($0.stringValue ?? "") }.sorted().joined(separator:"|")
    return (e.localName ?? "")+"["+attrs+"]"+(e.children?.compactMap { $0 as? XMLElement }.map(styleSignature).sorted().joined() ?? "")
}

struct Sheet { let name: String; let path: String; let hidden: Bool }
struct Rendered { let html: String; let rows: Int; let columns: Int; let shownRows: Int }
enum CellKind { case text, number, formula }
struct SheetIndex { var cells: [Int: XMLElement] = [:]; var blankCells: [Int:XMLElement] = [:]; var hiddenRows = Set<Int>(); var maxRow = 0, maxCol = 0 }

/// The edit grammar intentionally accepts either decimal separator, without ambiguous grouping.
func automaticNumber(_ input: String) -> Double? {
    var s = input.trimmingCharacters(in: .whitespacesAndNewlines)
    if let c = s.first, "€$£¥".contains(c) { s.removeFirst(); s = s.trimmingCharacters(in: .whitespaces) }
    let percent = s.hasSuffix("%")
    if percent { s.removeLast(); s = s.trimmingCharacters(in: .whitespaces) }
    guard s.range(of: "^[+-]?(?:[0-9]+(?:[.,][0-9]*)?|[.,][0-9]+)(?:[eE][+-]?[0-9]+)?$",options:.regularExpression) != nil,
          let n = Double(s.replacingOccurrences(of:",",with:".")), n.isFinite else { return nil }
    return percent ? n/100 : n
}
struct CellEdit { let ref: String; let value: String; let kind: CellKind }
struct CellSnapshot { let ref: String; let cell: XMLElement?; let rowExisted: Bool }
struct CellRange {
    var r1 = 1, c1 = 1, r2 = 1, c2 = 1
    var count: Int { (r2-r1+1)*(c2-c1+1) }
    var refs: [String] { (r1...r2).flatMap { r in (c1...c2).map { columnName($0)+String(r) } } }
}

/// Fill Down shifts relative row references; quoted strings and sheet names stay literal.
func fillFormula(_ formula: String, down delta: Int, right horizontal: Int = 0) -> String {
    let re = try! NSRegularExpression(pattern:"(?<![A-Za-z0-9_.])([$]?[A-Za-z]{1,3})([$]?)([1-9][0-9]{0,6})(?![A-Za-z0-9_!(])")
    func shift(_ chunk: String) -> String {
        let result = NSMutableString(string:chunk)
        for m in re.matches(in:chunk,range:NSRange(chunk.startIndex...,in:chunk)).reversed() {
            let ns = chunk as NSString
            let colText = ns.substring(with:m.range(at:1)), absoluteRow = ns.substring(with:m.range(at:2)) == "$", absoluteCol = colText.hasPrefix("$")
            let row = (Int(ns.substring(with:m.range(at:3))) ?? 0)+(absoluteRow ? 0 : delta)
            let col = column(colText.replacingOccurrences(of:"$",with:""))+(absoluteCol ? 0 : horizontal)
            result.replaceCharacters(in:m.range,with:row>0 && row<=1048576 && col>0 && col<=16384 ? (absoluteCol ? "$" : "")+columnName(col)+(absoluteRow ? "$" : "")+String(row) : "#REF!")
        }
        return result as String
    }
    var out = "", chunk = "", quoted: Character?
    for ch in formula {
        if let quote = quoted { out.append(ch); if ch == quote { quoted = nil }; continue }
        if ch == "\"" || ch == "'" { out += shift(chunk); chunk = ""; quoted = ch; out.append(ch) }
        else { chunk.append(ch) }
    }
    return out+shift(chunk)
}

/// Every document action owns a closed top-level group. Never mix Cocoa's per-event
/// implicit group (which can remain open across WK callbacks) with these explicit groups.
final class SheetUndoManager: UndoManager {
    override init() { super.init(); groupsByEvent = false }
    func record<T: AnyObject>(_ target: T, name: String, undo: @escaping (T) -> Void) {
        precondition(groupingLevel == 0 && !isUndoing && !isRedoing)
        beginUndoGrouping()
        registerUndo(withTarget:target,handler:undo)
        setActionName(name)
        endUndoGrouping()
        precondition(groupingLevel == 0)
    }
}

final class Book {
    var url: URL
    var untitled = false
    var temporaryDirectory: URL?
    deinit { if let dir = temporaryDirectory { try? FileManager.default.removeItem(at:dir) } }
    var documents: [String: XMLDocument] = [:]
    var dirty = false
    var calculationTime = Date()
    var strings: [String] = []; var sheets: [Sheet] = []; var styles: [CellStyle] = []; var date1904 = false
    var activeSheet = 0
    private var indexes: [Int: SheetIndex] = [:]
    private var memo: [Int: FValue] = [:]
    private var evaluating = Set<Int>()
    func invalidate() { indexes.removeAll(); memo.removeAll(); calculationTime = Date() }
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
        if let doc = try? xml("xl/styles.xml") { readStyles(doc) }
    }
    func readStyles(_ doc: XMLDocument) {
        styles = []
            let fonts = doc.nodes("/styleSheet/fonts/font"), fills = doc.nodes("/styleSheet/fills/fill")
            var formats = Book.builtinFormats
            for f in doc.nodes("/styleSheet/numFmts/numFmt") { formats[Int(f.attr("numFmtId")) ?? -1] = f.attr("formatCode") }
            for x in doc.nodes("/styleSheet/cellXfs/xf") {
                var s = CellStyle(); s.numberFormat = formats[Int(x.attr("numFmtId")) ?? 0] ?? "General"; s.setFormat(s.numberFormat)
                if let i = Int(x.attr("fontId")), fonts.indices.contains(i) {
                    let f = fonts[i]
                    if !f.nodes("b").isEmpty && f.val("b/@val") != "0" { s.bold = true; s.css += "font-weight:600;" }
                    if !f.nodes("i").isEmpty && f.val("i/@val") != "0" { s.italic = true; s.css += "font-style:italic;" }
                    if !f.nodes("strike").isEmpty && f.val("strike/@val") != "0" { s.strike = true; s.css += "text-decoration:line-through;" }
                    if let size = Double(f.val("sz/@val")), size>=1,size<=409 { s.fontSize = size; if size != 11 { s.css += "font-size:\(size)pt;" } }
                    if let c = color(f.nodes("color").first) { s.foreground = c; if c != "#000000" { s.css += "color:\(c);" } }
                }
                if let i = Int(x.attr("fillId")), fills.indices.contains(i), fills[i].val("patternFill/@patternType") == "solid", let c = color(fills[i].nodes("patternFill/fgColor").first) { s.background = c; s.css += "background:\(c);" }
                let align = x.val("alignment/@horizontal")
                if ["left", "right", "center"].contains(align) { s.alignment = align; s.explicitAlignment = true; s.css += "text-align:\(align);" }
                if ["1","true"].contains(x.val("alignment/@wrapText")) { s.wrap = true; s.css += "white-space:normal;overflow-wrap:anywhere;" }
                styles.append(s)
            }
    }
    func color(_ node: XMLNode?) -> String? {
        guard let node else { return nil }
        let rgb = node.attr("rgb"); if rgb.count >= 6, rgb.suffix(6).range(of:"^[0-9A-Fa-f]{6}$",options:.regularExpression) != nil { return "#" + rgb.suffix(6) }
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
        case .list(let l), .range(let l, _, _): return format(l.first ?? .empty, style)
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
                guard c > 0 else { continue }
                if cell.child("v") == nil && cell.child("is") == nil && cell.child("f") == nil { result.blankCells[cellKey(r,c)] = cell; if !cell.attr("s").isEmpty && cell.attr("s") != "0" { result.maxRow = max(result.maxRow,r); result.maxCol = max(result.maxCol,c) }; continue }
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
        _ = try index(sheet)
        let lastRow = r2, lastCol = c2
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
        let idx = try self.index(index); return idx.cells[cellKey(rowNumber(ref),column(ref))] ?? idx.blankCells[cellKey(rowNumber(ref),column(ref))]
    }
    func raw(_ sheet: Int, _ ref: String) throws -> String {
        guard let cell = try cell(sheet,ref) else { return "" }
        if let f = cell.child("f") { return "=" + (f.stringValue ?? "") }
        return try value(sheet,rowNumber(ref),column(ref)).text
    }
    func automaticEdit(_ sheet: Int, _ ref: String, _ text: String) throws -> CellEdit {
        guard column(ref) > 0, column(ref) <= 16384, rowNumber(ref) > 0, rowNumber(ref) <= 1048576 else { throw failure(L("Intervallo fuori dal foglio.", "Range outside the worksheet.")) }
        if text.hasPrefix("=") {
            do { let parser = FormulaParser(book:self,sheet:sheet); parser.validating = true; _ = try parser.evaluate(normalizeFormula(text)) }
            catch is FormulaInvalid { throw failure(L("Formula non valida: controlla parentesi, operatori e separatori degli argomenti.", "Invalid formula: check parentheses, operators and argument separators.")) }
            return CellEdit(ref:ref,value:normalizeFormula(text),kind:.formula)
        }
        if let n = automaticNumber(text) { return CellEdit(ref:ref,value:FValue.number(n).text,kind:.number) }
        return CellEdit(ref:ref,value:text,kind:.text)
    }
    func snapshots(_ sheet: Int, refs: [String]) throws -> [CellSnapshot] {
        let idx = try index(sheet), rows = Set(try sheetDocument(sheet).rootElement()?.child("sheetData")?.kids("row").compactMap { Int($0.attr("r")) } ?? [])
        var seen = Set<String>()
        return refs.filter { seen.insert($0).inserted }.map { ref in
            let key = cellKey(rowNumber(ref),column(ref)), cell = idx.cells[key] ?? idx.blankCells[key]
            return CellSnapshot(ref:ref,cell:cell?.copy() as? XMLElement,rowExisted:rows.contains(rowNumber(ref)))
        }
    }
    func restoreCells(_ sheet: Int, _ snapshots: [CellSnapshot]) throws {
        guard let data = try sheetDocument(sheet).rootElement()?.child("sheetData") else { return }
        let idx = try index(sheet)
        var rows = Dictionary(uniqueKeysWithValues:data.kids("row").map { (Int($0.attr("r")) ?? 0,$0) })
        for snapshot in snapshots {
            let r = rowNumber(snapshot.ref), c = column(snapshot.ref), key = cellKey(r,c)
            (idx.cells[key] ?? idx.blankCells[key])?.detach()
            if let saved = snapshot.cell {
                let row: XMLElement
                if let existing = rows[r] { row = existing }
                else { row = XMLElement(name:"row"); row.set("r",String(r)); let at = (data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0)>r } ?? data.childCount; data.insertChild(row,at:at); rows[r] = row }
                let at = (row.children ?? []).firstIndex { column($0.attr("r"))>c } ?? row.childCount
                row.insertChild(saved.copy() as! XMLElement,at:at)
            } else if !snapshot.rowExisted, let row = rows[r], row.kids("c").isEmpty { row.detach(); rows.removeValue(forKey:r) }
        }
        dirty = true; invalidate()
    }
    func setCell(_ index: Int, _ ref: String, _ value: String, kind: CellKind) throws {
        try applyEdits(index,[CellEdit(ref:ref,value:value,kind:kind)])
    }
    /// Preflight the complete batch before changing XML. A rejected paste cannot leave partial edits.
    func applyEdits(_ index: Int, _ edits: [CellEdit]) throws {
        let doc = try sheetDocument(index)
        guard doc.nodes("/worksheet/sheetProtection").isEmpty else { throw failure(L("Questo foglio è protetto. Rimuovi la protezione in Excel prima di modificarlo.", "This sheet is protected. Remove the protection in Excel before editing it.")) }
        guard let data = doc.rootElement()?.child("sheetData") else { return }
        let idx = try self.index(index)
        for edit in edits {
            guard column(edit.ref) > 0, column(edit.ref) <= 16384, rowNumber(edit.ref) > 0, rowNumber(edit.ref) <= 1048576 else { throw failure(L("Intervallo fuori dal foglio.", "Range outside the worksheet.")) }
            if let f = idx.cells[cellKey(rowNumber(edit.ref),column(edit.ref))]?.child("f"), ["shared","array","dataTable"].contains(f.attr("t")) { throw failure(L("L'intervallo contiene formule condivise o matriciali: modificalo in Excel.", "The range contains shared or array formulas: edit it in Excel.")) }
        }
        var rows = Dictionary(uniqueKeysWithValues:data.kids("row").map { (Int($0.attr("r")) ?? 0,$0) })
        var cells: [Int:XMLElement] = [:]
        // Include styled empty cells, which the value index deliberately omits.
        let touchedRows = Set(edits.map { rowNumber($0.ref) })
        for r in touchedRows { for c in rows[r]?.kids("c") ?? [] { cells[cellKey(r,column(c.attr("r")))] = c } }
        for edit in edits {
            let r = rowNumber(edit.ref), col = column(edit.ref), key = cellKey(r,col)
            let row: XMLElement
            if let existing = rows[r] { row = existing }
            else { row = XMLElement(name:"row"); row.set("r",String(r)); let place = (data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0)>r } ?? data.childCount; data.insertChild(row,at:place); rows[r] = row }
            let c: XMLElement
            if let existing = cells[key] { c = existing }
            else { c = XMLElement(name:"c"); c.set("r",edit.ref); let place = (row.children ?? []).firstIndex { column($0.attr("r"))>col } ?? row.childCount; row.insertChild(c,at:place); cells[key] = c }
            for node in c.kids("f") + c.kids("v") + c.kids("is") { node.detach() }; c.removeAttribute(forName:"t")
            if !edit.value.isEmpty {
                switch edit.kind {
                case .formula: c.addChild(XMLElement(name:"f",stringValue:normalizeFormula(edit.value)))
                case .number: c.addChild(XMLElement(name:"v",stringValue:edit.value))
                case .text:
                    c.set("t","inlineStr"); let inline = XMLElement(name:"is"), t = XMLElement(name:"t",stringValue:edit.value); t.set("xml:space","preserve"); inline.addChild(t); c.addChild(inline)
                }
            }
        }
        doc.rootElement()?.child("dimension")?.detach() // Excel derives the used range on reopening.
        dirty = true; invalidate()
    }
    static let builtinFormats = [0:"General",1:"0",2:"0.00",3:"#,##0",4:"#,##0.00",9:"0%",10:"0.00%",11:"0.00E+00",14:"dd/mm/yyyy",15:"d-mmm-yy",16:"d-mmm",17:"mmm-yy",18:"h:mm",19:"h:mm:ss",20:"h:mm",21:"h:mm:ss",22:"dd/mm/yyyy h:mm",37:"#,##0",38:"#,##0",39:"#,##0.00",40:"#,##0.00",49:"@"]
    func styleDocument() throws -> XMLDocument {
        if let doc = documents["xl/styles.xml"] { return doc }
        let doc: XMLDocument
        if let existing = try? xml("xl/styles.xml") { doc = existing }
        else { let blank = try Book.blank(); doc = try blank.xml("xl/styles.xml") }
        documents["xl/styles.xml"] = doc; return doc
    }
    func editable(_ sheet: Int) throws {
        guard try sheetDocument(sheet).rootElement()?.child("sheetProtection") == nil else { throw failure(L("Il foglio è protetto.","The sheet is protected.")) }
    }
    /// Styles are immutable pool entries. Undo restores cell s attributes; unused pool entries are harmless.
    func formatCells(_ sheet: Int, refs: [String], change: StyleChange) throws {
        try editable(sheet)
        guard !refs.isEmpty, let root = try styleDocument().rootElement() else { return }
        func container(_ name: String) -> XMLElement {
            if let e = root.child(name) { return e }
            let e = XMLElement(name:name), order = ["numFmts","fonts","fills","borders","cellStyleXfs","cellXfs","cellStyles","dxfs","tableStyles","colors","extLst"]
            let at = (root.children ?? []).firstIndex { (order.firstIndex(of:$0.localName ?? "") ?? 99) > (order.firstIndex(of:name) ?? 99) } ?? root.childCount
            root.insertChild(e,at:at); return e
        }
        func poolID(_ pool: XMLElement, _ entry: XMLElement) -> Int {
            let signature = styleSignature(entry), entries = pool.children?.compactMap { $0 as? XMLElement } ?? []
            if let i = entries.firstIndex(where:{ styleSignature($0) == signature }) { return i }
            pool.addChild(entry); pool.set("count",String(entries.count+1)); return entries.count
        }
        func formatID(_ code: String) -> Int {
            if let id = Book.builtinFormats.first(where:{ [0,1,2,3,4,9,10,11,20,21,49].contains($0.key) && $0.value == code })?.key { return id }
            let formats = container("numFmts")
            if let e = formats.kids("numFmt").first(where:{ $0.attr("formatCode") == code }) { return Int(e.attr("numFmtId"))! }
            let id = max(164,(formats.kids("numFmt").compactMap { Int($0.attr("numFmtId")) }.max() ?? 163)+1)
            let e = XMLElement(name:"numFmt"); e.set("numFmtId",String(id)); e.set("formatCode",code); formats.addChild(e); formats.set("count",String(formats.kids("numFmt").count)); return id
        }
        let fonts = container("fonts"), fills = container("fills"), xfs = container("cellXfs")
        var mapped: [Int:Int] = [:]
        let needed = try refs.filter { try cell(sheet,$0) == nil }
        if !needed.isEmpty { try applyEdits(sheet,needed.map { CellEdit(ref:$0,value:"",kind:.text) }) }
        for ref in refs {
            guard let c = try cell(sheet,ref) else { continue }
            let oldID = Int(c.attr("s")) ?? 0
            if let id = mapped[oldID] { c.set("s",String(id)); continue }
            if case .copy(let id) = change { guard xfs.kids("xf").indices.contains(id) else { continue }; c.set("s",String(id)); mapped[oldID] = id; continue }
            let all = xfs.kids("xf"), xf = (all.indices.contains(oldID) ? all[oldID].copy() as? XMLElement : nil) ?? XMLElement(name:"xf")
            let style = styles.indices.contains(oldID) ? styles[oldID] : CellStyle()
            switch change {
            case .number(let code): xf.set("numFmtId",String(formatID(code))); xf.set("applyNumberFormat","1")
            case .decimals(let delta):
                var code = style.numberFormat
                if let range = code.range(of:"[0#][0#,]*(?:\\.[0#]+)?",options:.regularExpression) {
                    let digits = String(code[range]), parts = digits.split(separator:".",omittingEmptySubsequences:false)
                    let count = min(10,max(0,(parts.count>1 ? parts[1].count : 0)+delta))
                    code.replaceSubrange(range,with:String(parts[0])+(count>0 ? "."+String(repeating:"0",count:count) : ""))
                } else { code = delta>0 ? "0.0" : "0" }
                xf.set("numFmtId",String(formatID(code))); xf.set("applyNumberFormat","1")
            case .flag, .fontSize, .fontDelta, .foreground:
                let fontID = Int(xf.attr("fontId")) ?? 0, entries = fonts.kids("font")
                let font = (entries.indices.contains(fontID) ? entries[fontID].copy() as? XMLElement : nil) ?? XMLElement(name:"font")
                func replace(_ name: String, _ attr: String? = nil, _ value: String? = nil) {
                    font.kids(name).forEach { $0.detach() }; let e = XMLElement(name:name); if let a = attr, let v = value { e.set(a,v) }; font.addChild(e)
                }
                switch change {
                case .flag(let name,let on): font.kids(name).forEach { $0.detach() }; if on { font.addChild(XMLElement(name:name)) }
                case .foreground(let rgb): replace("color","rgb","FF"+rgb.replacingOccurrences(of:"#",with:""))
                case .fontSize(let size): replace("sz","val",String(min(72,max(6,size))))
                case .fontDelta(let delta): replace("sz","val",String(min(72,max(6,style.fontSize+delta))))
                default: break
                }
                xf.set("fontId",String(poolID(fonts,font))); xf.set("applyFont","1")
            case .fill(let rgb):
                let fill = XMLElement(name:"fill"), pattern = XMLElement(name:"patternFill"), fg = XMLElement(name:"fgColor"), bg = XMLElement(name:"bgColor")
                pattern.set("patternType","solid"); fg.set("rgb","FF"+rgb.replacingOccurrences(of:"#",with:"")); bg.set("indexed","64"); pattern.addChild(fg); pattern.addChild(bg); fill.addChild(pattern)
                xf.set("fillId",String(poolID(fills,fill))); xf.set("applyFill","1")
            case .alignment(let horizontal):
                let alignment = xf.child("alignment") ?? XMLElement(name:"alignment"); if alignment.parent == nil { xf.addChild(alignment) }; alignment.set("horizontal",horizontal); xf.set("applyAlignment","1")
            case .wrap(let on):
                let alignment = xf.child("alignment") ?? XMLElement(name:"alignment"); if alignment.parent == nil { xf.addChild(alignment) }; alignment.set("wrapText",on ? "1" : "0"); xf.set("applyAlignment","1")
            case .copy: break
            }
            let id = poolID(xfs,xf); c.set("s",String(id)); mapped[oldID] = id
        }
        readStyles(try styleDocument()); dirty = true; invalidate()
    }
    func autofillEdits(_ sheet: Int, source: CellRange, target: CellRange) throws -> [CellEdit] {
        let down = target.r2>source.r2
        guard target.r1 == source.r1,target.c1 == source.c1, (down ? target.c2 == source.c2 : target.r2 == source.r2), target.count<=400000 else { return [] }
        var result: [CellEdit] = []
        let length = down ? source.r2-source.r1+1 : source.c2-source.c1+1
        for lane in (down ? source.c1...source.c2 : source.r1...source.r2) {
            var seeds: [FValue] = [], raws: [String] = []
            for j in 0..<length { let r = down ? source.r1+j : lane, c = down ? lane : source.c1+j; seeds.append(try value(sheet,r,c)); raws.append(try raw(sheet,columnName(c)+String(r))) }
            let nums = seeds.compactMap { v -> Double? in if case .number(let n) = v { return n }; return nil }
            var step: Double?
            if length>=2, nums.count == length { let d = nums[1]-nums[0]; if (1..<length).allSatisfy({ abs((nums[$0]-nums[$0-1])-d)<1e-9*max(1,abs(d)) }) { step = d } }
            let last = down ? target.r2-source.r1 : target.c2-source.c1
            guard last>=length else { continue }
            for j in length...last {
                let seed = j%length, r = down ? source.r1+j : lane, c = down ? lane : source.c1+j, ref = columnName(c)+String(r)
                if let step { result.append(CellEdit(ref:ref,value:FValue.number(nums[0]+Double(j)*step).text,kind:.number)) }
                else if raws[seed].hasPrefix("=") { result.append(CellEdit(ref:ref,value:fillFormula(String(raws[seed].dropFirst()),down:down ? j-seed : 0,right:down ? 0 : j-seed),kind:.formula)) }
                else { result.append(CellEdit(ref:ref,value:raws[seed],kind: { if case .number = seeds[seed] { return .number }; return .text }())) }
            }
        }
        return result
    }
    func dataRegion(_ sheet: Int, selection: CellRange) throws -> CellRange {
        if selection.count>1 { return selection }
        let idx = try index(sheet), start = cellKey(selection.r1,selection.c1)
        guard idx.cells[start] != nil else { throw failure(L("Seleziona una cella nell'area dei dati.","Select a cell in the data region.")) }
        var region = selection, seen = Set([start]), todo = [start], i = 0
        while i<todo.count {
            let key = todo[i]; i += 1; let r = key/16385,c = key%16385
            region.r1 = min(region.r1,r); region.r2 = max(region.r2,r); region.c1 = min(region.c1,c); region.c2 = max(region.c2,c)
            for (rr,cc) in [(r-1,c),(r+1,c),(r,c-1),(r,c+1)] where rr>0 && cc>0 && cc<=16384 {
                let next = cellKey(rr,cc); if idx.cells[next] != nil && seen.insert(next).inserted { todo.append(next) }
            }
        }
        return region
    }
    func toggleFilter(_ sheet: Int, selection: CellRange) throws {
        try editable(sheet); guard let root = try sheetDocument(sheet).rootElement() else { return }
        if let filter = root.child("autoFilter") { filter.detach() }
        else {
            let range = try dataRegion(sheet,selection:selection), filter = XMLElement(name:"autoFilter")
            filter.set("ref",columnName(range.c1)+String(range.r1)+":"+columnName(range.c2)+String(range.r2))
            let before = ["sortState","dataConsolidate","customSheetViews","mergeCells","phoneticPr","conditionalFormatting","dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","drawing","legacyDrawing","extLst"]
            let at = (root.children ?? []).firstIndex { before.contains($0.localName ?? "") } ?? root.childCount; root.insertChild(filter,at:at)
        }
        dirty = true; invalidate()
    }
    func aggregateFormula(_ sheet: Int, ref: String, function: String) throws -> String {
        let r = rowNumber(ref), c = column(ref)
        func numeric(_ row: Int, _ col: Int) throws -> Bool { guard row>0,col>0 else { return false }; if case .number = try value(sheet,row,col) { return true }; return false }
        if try numeric(r-1,c) { var first = r-1; while try numeric(first-1,c) { first -= 1 }; return "="+function+"("+columnName(c)+String(first)+":"+columnName(c)+String(r-1)+")" }
        if try numeric(r,c-1) { var first = c-1; while try numeric(r,first-1) { first -= 1 }; return "="+function+"("+columnName(first)+String(r)+":"+columnName(c-1)+String(r)+")" }
        throw failure(L("Non ci sono numeri contigui sopra o a sinistra della cella.","There are no contiguous numbers above or to the left of this cell."))
    }
    func sort(_ sheet: Int, column col: Int, row selectedRow: Int, ascending: Bool) throws {
        let doc = try sheetDocument(sheet), idx = try index(sheet)
        guard doc.rootElement()?.child("sheetProtection") == nil else { throw failure(L("Il foglio è protetto.", "The sheet is protected.")) }
        let filter = doc.rootElement()?.child("autoFilter")?.attr("ref").split(separator:":").map(String.init) ?? []
        var first = 1, last = idx.maxRow
        if let lo = filter.first, let hi = filter.last {
            guard col >= column(lo), col <= column(hi) else { throw failure(L("Seleziona una colonna nell'intervallo del filtro.", "Select a column inside the filter range.")) }
            first = rowNumber(lo); last = rowNumber(hi)
        } else {
            // Find the contiguous block containing the selection. The first row is its header.
            let occupied = Set(idx.cells.keys.map { $0/16385 })
            first = max(1,selectedRow); while first>1 && occupied.contains(first-1) { first -= 1 }
            last = max(first,selectedRow); while occupied.contains(last+1) { last += 1 }
        }
        let formulaRows = idx.cells.filter { $0.key/16385 >= first && $0.key/16385 <= last && $0.value.child("f") != nil }
        guard formulaRows.isEmpty else { throw failure(L("Impossibile ordinare: l'intervallo contiene formule. I riferimenti delle formule non verrebbero aggiornati.", "Cannot sort: the region contains formulas. Formula references would not be updated.")) }
        for merge in doc.rootElement()?.child("mergeCells")?.kids("mergeCell") ?? [] {
            let ends = merge.attr("ref").split(separator:":").map(String.init)
            if let a = ends.first, let b = ends.last, rowNumber(a) <= last, rowNumber(b) >= first { throw failure(L("Impossibile ordinare un intervallo con celle unite.", "Cannot sort a region containing merged cells.")) }
        }
        guard last > first else { return }
        guard let data = doc.rootElement()?.child("sheetData") else { return }
        let rows = data.kids("row").filter { let n = Int($0.attr("r")) ?? 0; return n>first && n<=last }
        var values: [Int:FValue] = [:]
        for row in rows { let r = Int(row.attr("r"))!; values[r] = try value(sheet,r,col) }
        let sorted = rows.enumerated().sorted { a,b in
            let x = values[Int(a.element.attr("r"))!] ?? .empty, y = values[Int(b.element.attr("r"))!] ?? .empty
            if x.text.isEmpty != y.text.isEmpty { return !x.text.isEmpty }
            let xn: Double? = { if case .number(let n) = x { return n }; return nil }(), yn: Double? = { if case .number(let n) = y { return n }; return nil }()
            if (xn != nil) != (yn != nil) { return xn != nil }
            let cmp: Int = xn != nil && yn != nil ? (xn! < yn! ? -1 : (xn! > yn! ? 1 : 0)) : x.text.caseInsensitiveCompare(y.text).rawValue
            return cmp == 0 ? a.offset<b.offset : (ascending ? cmp<0 : cmp>0)
        }.map(\.element)
        rows.forEach { $0.detach() }
        let insertion = (data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0)>first } ?? data.childCount
        for (i,row) in sorted.enumerated() { let r = first+1+i; row.set("r",String(r)); for c in row.kids("c") { c.set("r",columnName(column(c.attr("r")))+String(r)) }; data.insertChild(row,at:insertion+i) }
        dirty = true; invalidate()
    }
    static func blank() throws -> Book {
        let fm = FileManager.default, dir = fm.temporaryDirectory.appendingPathComponent("xlsxlite-"+UUID().uuidString)
        try fm.createDirectory(at:dir,withIntermediateDirectories:true)
        let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        let parts = [
            "[Content_Types].xml": "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>",
            "_rels/.rels": "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"\(rel)/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
            "xl/workbook.xml": "<workbook xmlns=\"\(ns)\" xmlns:r=\"\(rel)\"><sheets><sheet name=\"\(L("Foglio1","Sheet1"))\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            "xl/_rels/workbook.xml.rels": "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"\(rel)/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"\(rel)/styles\" Target=\"styles.xml\"/></Relationships>",
            "xl/worksheets/sheet1.xml": "<worksheet xmlns=\"\(ns)\"><sheetData/></worksheet>",
            "xl/styles.xml": "<styleSheet xmlns=\"\(ns)\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>"
        ]
        let contents = dir.appendingPathComponent("contents"); defer { try? fm.removeItem(at:contents) }
        for (path,text) in parts { let dest = contents.appendingPathComponent(path); try fm.createDirectory(at:dest.deletingLastPathComponent(),withIntermediateDirectories:true); try ("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"+text).write(to:dest,atomically:true,encoding:.utf8) }
        let target = dir.appendingPathComponent("Untitled.xlsx"), process = Process(); process.executableURL = URL(fileURLWithPath:"/usr/bin/zip"); process.arguments = ["-q","-r",target.path,"."]; process.currentDirectoryURL = contents
        try process.run(); process.waitUntilExit(); guard process.terminationStatus == 0 else { throw failure(L("Impossibile creare il file.", "Could not create workbook.")) }
        let book = try Book(target); book.untitled = true; book.temporaryDirectory = dir; return book
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
            for cell in doc.nodes("/worksheet/sheetData/row/c[f]") { if let c = cell as? XMLElement, (try? computed(c,sheet:i)) != nil { for v in c.nodes("v") { v.detach() } } }
            try write(sheets[i].path, doc)
        }
        if let styles = documents["xl/styles.xml"] { try write("xl/styles.xml",styles) }
        let workbook = try xml("xl/workbook.xml")
        let calc = (workbook.nodes("/workbook/calcPr").first as? XMLElement) ?? XMLElement(name: "calcPr")
        if calc.parent == nil { workbook.rootElement()?.addChild(calc) }
        for (key, value) in [("calcMode", "auto"), ("fullCalcOnLoad", "1"), ("forceFullCalc", "1")] { calc.removeAttribute(forName: key); calc.addAttribute(XMLNode.attribute(withName: key, stringValue: value) as! XMLNode) }
        try write("xl/workbook.xml", workbook)
        let rels = try xml("xl/_rels/workbook.xml.rels")
        for r in rels.nodes("/Relationships/Relationship") where r.attr("Type").hasSuffix("/calcChain") { r.detach() }
        if documents["xl/styles.xml"] != nil, !rels.nodes("/Relationships/Relationship").contains(where:{ $0.attr("Type").hasSuffix("/styles") }) {
            let node = XMLElement(name:"Relationship"), used = Set(rels.nodes("/Relationships/Relationship").map { $0.attr("Id") }); var n = 1; while used.contains("rId\(n)") { n += 1 }
            node.set("Id","rId\(n)"); node.set("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"); node.set("Target","styles.xml"); rels.rootElement()?.addChild(node)
        }
        try write("xl/_rels/workbook.xml.rels", rels)
        let types = try xml("[Content_Types].xml")
        for t in types.nodes("/Types/Override") where t.attr("PartName") == "/xl/calcChain.xml" { t.detach() }
        if documents["xl/styles.xml"] != nil, !types.nodes("/Types/Override").contains(where:{ $0.attr("PartName") == "/xl/styles.xml" }) {
            let node = XMLElement(name:"Override"); node.set("PartName","/xl/styles.xml"); node.set("ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"); types.rootElement()?.addChild(node)
        }
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
        memo.removeAll(); calculationTime = Date()
        let sheetIndex = try self.index(index)
        let grid = sheetIndex.cells, hiddenRows = sheetIndex.hiddenRows
        var maxRow = sheetIndex.maxRow, maxCol = sheetIndex.maxCol
        maxCol = min(max(maxCol, maxRow <= 1000 ? 26 : 1), 1000); maxRow = max(maxRow, 100)
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
        let marginTop = 70 + frozenRows * 27, marginLeft = Int(offsets[columns.first { $0 > frozenCols } ?? 1] ?? 40) + 4
        let strings = ["search": L("Cerca valore…", "Search values…"), "all": L("Tutti i valori", "All values"), "empty": L("(Vuote)", "(Blanks)"), "more": L("Altri valori: usa la ricerca", "More values: use search"), "asc": L("Ordina A → Z","Sort A → Z"), "desc": L("Ordina Z → A","Sort Z → A"), "sum": L("Somma","Sum"), "average": L("Media","Average"), "count": L("Conteggio","Count"), "edit": L("Modifica cella","Edit cell")]
        let stringsJSON = String(data: try JSONSerialization.data(withJSONObject: strings), encoding: .utf8)!

        let styleRules = styles.enumerated().filter { !$0.element.css.isEmpty }.map { "td.s\($0.offset){\($0.element.css)}" }.joined()
        var numericJS = "[", firstNumber = true; numericJS.reserveCapacity(rows.count*columns.count*8)
        var formulaRaw: [String:String] = [:]
        var html = """
        <!doctype html><html><head><meta charset="utf-8"><style>
        :root{color-scheme:light}*{box-sizing:border-box}body{margin:0;background:#fff;color:#202124;font:13px -apple-system,system-ui;font-variant-numeric:tabular-nums}
        table{border-collapse:separate;border-spacing:0;table-layout:fixed;width:\(Int(totalWidth))px}
        td,th{border-right:1px solid #dfe3e9;border-bottom:1px solid #e5e7eb;background:#fff;height:27px;padding:0 8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:middle;scroll-margin:\(marginTop)px 4px 4px \(marginLeft)px}
        th{background:#f3f4f6;color:#7b8088;font-size:10px;font-weight:500;text-align:center;padding:0}
        thead th{position:sticky;top:42px;z-index:6;height:24px}tbody th{position:sticky;left:0;z-index:4}
        td[n]{text-align:right}\(styleRules)
        td.sel{box-shadow:inset 0 0 0 9999px #1a73e81a!important}td.active{box-shadow:none!important}td.sel::selection{background:transparent}tr:hover td{box-shadow:inset 0 0 0 999px #8080800c}::selection{background:#b9d9ff}
        td.fh{cursor:pointer;padding-right:18px;position:relative}td.fh::after{content:'▾';position:absolute;right:5px;top:50%;transform:translateY(-50%);color:#7b8088}td.fh.on::after{content:'●';color:#06f}
        #filter-menu{position:fixed;z-index:100;background:#fff;color:#222;border:1px solid #ddd;border-radius:9px;box-shadow:0 5px 25px #0003;width:230px;padding:6px;max-height:350px;overflow:auto;font:13px -apple-system,system-ui}
        #filter-menu button{display:block;border:0;background:transparent;border-radius:4px;text-align:left;width:100%;padding:7px;color:#222;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}#filter-menu button:hover{background:#e6efff}
        #filter-menu input{width:100%;padding:7px;border:1px solid #ddd;border-radius:4px;margin-bottom:4px}#filter-menu .more{padding:7px;color:#888;font-size:12px}
        body{padding-bottom:24px}#formula-bar{position:sticky;top:0;left:0;z-index:30;height:42px;background:#f8f9fa;display:flex;align-items:center;gap:8px;padding:5px;width:100vw;border-bottom:1px solid #ddd}#name{width:110px}#formula{flex:1;min-width:80px;height:29px;resize:none;white-space:pre}#formula-bar :is(input,textarea){font:13px -apple-system;padding:5px;border:1px solid #ddd;border-radius:3px}#message{color:#b3261e;font-size:12px;max-width:35vw}#editor{resize:none;position:absolute;z-index:40;border:2px solid #2979e8;background:white;color:#202124;font:13px -apple-system;padding:0 6px;outline:none}#selection-border{position:absolute;pointer-events:none;z-index:12;border:2px solid #1a73e8}#fill-handle{position:absolute;width:8px;height:8px;right:-5px;bottom:-5px;border:1px solid white;background:#1a73e8;cursor:crosshair;pointer-events:auto}th.selected-header{background:#d3e3fd!important;color:#174ea6;font-weight:700}body.paint td{cursor:crosshair}#stats{position:fixed;bottom:0;right:0;z-index:35;background:#f8f9fa;padding:4px 12px;font-size:12px}table{user-select:none}#formula-bar :is(input,textarea){user-select:text}
        @media print{#formula-bar,#stats,#editor,#message,#selection-border{display:none!important}thead th,tbody th,td{position:static!important}td.sel{outline:none;box-shadow:none!important}th.selected-header{background:#f3f4f6!important}tr:hover td{box-shadow:none}td.fh::after{content:''}}
        </style></head><body><div id="formula-bar"><input id="name" aria-label="\(L("Cella o intervallo","Cell or range"))" placeholder="A1"><span style="font:italic 16px Georgia;color:#747775">fx</span><textarea id="formula" rows="1" wrap="off" aria-label="\(L("Barra della formula","Formula bar"))"></textarea><span id="message" role="alert"></span></div><div id="stats"></div><table><colgroup><col style="width:40px">
        """
        html.reserveCapacity(rows.count * columns.count * 48 + 4096)
        for c in columns { html += "<col data-c='\(c)' style='width:\(Int(widths[c] ?? 110))px'>" }
        html += "</colgroup><thead><tr><th style='left:0;z-index:9'></th>"
        for c in columns { html += "<th data-c='\(c)'\(c <= frozenCols ? " style='left:\(Int(offsets[c] ?? 40))px;z-index:8'" : "")>\(columnName(c))</th>" }
        html += "</tr></thead><tbody>"
        for (i, r) in rows.enumerated() {
            let isFrozen = r <= frozenRows, top = 66 + i * 27
            html += "<tr data-row='\(r)'\(r == frozenRows && frozenRows > 0 ? " class='heading'" : "")><th\(isFrozen ? " style='top:\(top)px;z-index:8'" : "")>\(r)</th>"
            for c in columns {
                let key = cellKey(r, c); if covered.contains(key) { if !firstNumber { numericJS += "," }; firstNumber = false; numericJS += "null"; continue }
                let cell = grid[key] ?? sheetIndex.blankCells[key]
                let styleID = cell.flatMap { Int($0.attr("s")) } ?? 0
                let sty = styles.indices.contains(styleID) ? styles[styleID] : CellStyle()
                var css = "", classes: [String] = []
                if !sty.css.isEmpty { classes.append("s\(styleID)") }
                if isFrozen { css += "position:sticky;top:\(top)px;z-index:5;" }
                if c <= frozenCols { css += "position:sticky;left:\(Int(offsets[c] ?? 40))px;z-index:\(isFrozen ? 7 : 3);" }
                if r == filterRow && c >= column(filterStart) && c <= column(filterEnd) { classes.append("fh") }
                var numeric: Double?
                var value = ""
                if let cell {
                    value = display(cell,sheet:index)
                    if let f = cell.child("f") {
                        formulaRaw[String(key)] = "="+(f.stringValue ?? "")
                        if case .number(let n)? = try? self.value(index,r,c), n.isFinite { numeric = n }
                    } else if ["","n"].contains(cell.attr("t")), let text = cell.child("v")?.stringValue, let n = Double(text), n.isFinite { numeric = n }
                }
                if !firstNumber { numericJS += "," }; firstNumber = false; numericJS += numeric.map { FValue.number($0).text } ?? "null"
                html += "<td data-c='\(c)'"
                if numeric != nil && sty.kind != .text { html += " n" }
                if !classes.isEmpty { html += " class='\(classes.joined(separator:" "))'" }
                if let span = merges[key] { html += " colspan='\(span.0)' rowspan='\(span.1)'" }
                if !css.isEmpty { html += " style='\(css)'" }
                if value.contains("\n") { html += " data-full='\(esc(value))'" }
                html += ">\(esc(value.replacingOccurrences(of:"\n",with:" ")))</td>"
            }
            html += "</tr>"
        }
        numericJS += "]"
        let formulaJSON = String(data:try JSONSerialization.data(withJSONObject:formulaRaw),encoding:.utf8)!.replacingOccurrences(of:"<",with:"\\u003c")
        html += """
        </tbody></table><script>
        const T=\(stringsJSON);let FR=\(filterRow),FL=\(filterLastRow),filterC1=\(column(filterStart)),filterC2=\(column(filterEnd));let filters={},current=null,anchor=null,selected=new Set(),dragging=false,editing=null,pendingMove=null,busy=false;
        const HR=new Set(\(hiddenRows.sorted())),HC=new Set(\(hiddenCols.sorted()));
        const full=td=>td.dataset.full??td.textContent;
        function colName(c){let s='';while(c>0){c--;s=String.fromCharCode(65+c%26)+s;c=Math.floor(c/26)}return s}
        const refOf=td=>colName(+td.dataset.c)+td.parentElement.dataset.row;
        const mutationActions=new Set(['commit','clear','cut','paste','fillDown','autofill','undo','redo','sort']);
        const post=m=>{if(mutationActions.has(m.action)){if(busy)return;busy=true}window.webkit?.messageHandlers.cell.postMessage(m)};
        const rows=[...document.querySelectorAll('tbody tr')],rowMap=new Map(rows.map(r=>[+r.dataset.row,r]));
        const cellAt=(tr,c)=>tr?.querySelector(`td[data-c="${c}"]`);
        const at=(r,c)=>cellAt(rowMap.get(r),c);
        const bar=document.getElementById('formula'),nameBox=document.getElementById('name'),stats=document.getElementById('stats');
        const COLS=\(columns),numbers=new Map(),raws=new Map(Object.entries(\(formulaJSON)).map(([k,v])=>[+k,v]));
        const keyOf=td=>+td.parentElement.dataset.row*16385+(+td.dataset.c);
        (()=>{const values=\(numericJS);let i=0;for(const row of rows)for(const c of COLS){const n=values[i++];if(n!==null)numbers.set(+row.dataset.row*16385+c,n)}})();
        const raw=td=>raws.get(keyOf(td))??(numbers.has(keyOf(td))?String(numbers.get(keyOf(td))):full(td));
        function closeFilter(){document.getElementById('filter-menu')?.remove()}
        function dataRows(){return rows.filter(r=>{const n=+r.dataset.row;return n>FR&&n<=FL})}
        function applyFilters(){const keys=Object.entries(filters);dataRows().forEach(row=>{row.hidden=keys.some(([c,v])=>{const td=cellAt(row,c);return (td?full(td):'')!==v})});document.querySelectorAll('td.fh').forEach(td=>td.classList.toggle('on',filters[td.dataset.c]!==undefined));closeFilter()}
        window.setFilterRegion=(r1,c1,r2,c2)=>{if(FR===r1&&FL===r2&&filterC1===c1&&filterC2===c2)return;for(const row of rows)row.hidden=false;filters={};FR=r1;FL=r2;filterC1=c1;filterC2=c2;document.querySelectorAll('td.fh').forEach(td=>td.classList.remove('fh','on'));if(r1)for(let c=c1;c<=c2;c++)at(r1,c)?.classList.add('fh');selectionFurniture()};
        window.clearFilters=()=>{filters={};applyFilters()};
        function bounds(){const a=anchor??current,b=current;return {r1:Math.min(+a.parentElement.dataset.row,+b.parentElement.dataset.row),r2:Math.max(+a.parentElement.dataset.row,+b.parentElement.dataset.row),c1:Math.min(+a.dataset.c,+b.dataset.c),c2:Math.max(+a.dataset.c,+b.dataset.c)}}
        function updateStats(){let sum=0,count=0;for(const td of selected){if(numbers.has(keyOf(td))){sum+=numbers.get(keyOf(td));count++}}stats.dataset.sum=sum;stats.dataset.average=count?sum/count:0;stats.dataset.count=count;const fmt=n=>n.toLocaleString(undefined,{maximumFractionDigits:10});stats.textContent=selected.size<2?'':`${T.sum}: ${fmt(sum)}   ${T.average}: ${count?fmt(sum/count):'—'}   ${T.count}: ${fmt(count)}`}
        const selectionBorder=document.createElement('div');selectionBorder.id='selection-border';const handle=document.createElement('div');handle.id='fill-handle';selectionBorder.append(handle);document.body.append(selectionBorder);
        let selectedHeaders=[],filling=null,fillTarget=null,paintArmed=false,pendingFill=null;
        function selectionFurniture(){if(!current)return;for(const th of selectedHeaders)th.classList.remove('selected-header');selectedHeaders=[];const b=bounds();for(const th of document.querySelectorAll('thead th[data-c]'))if(+th.dataset.c>=b.c1&&+th.dataset.c<=b.c2)selectedHeaders.push(th);for(let r=b.r1;r<=b.r2;r++){const th=rowMap.get(r)?.querySelector('th');if(th)selectedHeaders.push(th)}for(const th of selectedHeaders)th.classList.add('selected-header');const a=at(b.r1,b.c1),z=at(b.r2,b.c2);if(!a||!z){selectionBorder.hidden=true;return}const x=a.getBoundingClientRect(),y=z.getBoundingClientRect();selectionBorder.hidden=false;Object.assign(selectionBorder.style,{left:x.left+scrollX+'px',top:x.top+scrollY+'px',width:y.right-x.left+'px',height:y.bottom-x.top+'px'})}
        window.armPaint=on=>{paintArmed=on;document.body.classList.toggle('paint',on)};
        function finishPaint(){if(!paintArmed)return;paintArmed=false;document.body.classList.remove('paint');post({action:'paint'})}
        handle.addEventListener('mousedown',e=>{if(busy)return;e.preventDefault();e.stopPropagation();filling=bounds();fillTarget=filling});
        function notifySelection(){selectionFurniture();const b=bounds();post({action:'select',ref:refOf(current),text:full(current),...Object.fromEntries(Object.entries(b).map(([k,v])=>[k,String(v)]))});nameBox.value=selected.size>1?colName(b.c1)+b.r1+':'+colName(b.c2)+b.r2:refOf(current);bar.value=raw(current);updateStats()}
        function select(td,scroll=false,extend=false){if(!td)return;if(editing)cancelEdit();if(!extend||!anchor)anchor=td;current?.classList.remove('active');current=td;current.classList.add('active');const b=bounds(),next=new Set();for(let r=b.r1;r<=b.r2;r++){const row=rowMap.get(r);if(!row||row.hidden)continue;for(let c=b.c1;c<=b.c2;c++){const x=cellAt(row,c);if(x)next.add(x)}}for(const x of selected)if(!next.has(x))x.classList.remove('sel');for(const x of next)if(!selected.has(x))x.classList.add('sel');selected=next;const s=getSelection(),range=document.createRange();range.selectNodeContents(td);s.removeAllRanges();s.addRange(range);if(scroll)td.scrollIntoView({block:'nearest',inline:'nearest'});notifySelection()}
        function ensure(r,c){let width=Math.max(...[...document.querySelectorAll('thead th')].map(th=>+th.dataset.c||0));const table=document.querySelector('table');while(width<c){width++;const col=document.createElement('col');col.style.width='110px';col.dataset.c=width;document.querySelector('colgroup').append(col);const th=document.createElement('th');th.textContent=colName(width);th.dataset.c=width;document.querySelector('thead tr').append(th);for(const row of rows){const td=document.createElement('td');td.dataset.c=width;row.append(td)}}table.style.width=[...document.querySelectorAll('colgroup col')].reduce((n,x)=>n+parseFloat(x.style.width),0)+'px';let last=+(rows.at(-1)?.dataset.row??0);while(last<r){last++;const row=document.createElement('tr');row.dataset.row=last;const th=document.createElement('th');th.textContent=last;row.append(th);for(let cc=1;cc<=width;cc++){const td=document.createElement('td');td.dataset.c=cc;row.append(td)}document.querySelector('tbody').append(row);rows.push(row);rowMap.set(last,row)}return at(r,c)}
        window.setCells=list=>{for(const [r,c,value,pick,rawValue,n,css] of list){if(HR.has(r)||HC.has(c))continue;let td=at(r,c);if(!td&&rowMap.has(r)&&c<=Math.max(...[...document.querySelectorAll('thead th')].map(th=>+th.dataset.c||0)))continue;if(!td&&!pick)continue;if(!td)td=ensure(r,c);if(!td)continue;td.textContent=value.replaceAll('\\n',' ');if(value.includes('\\n'))td.dataset.full=value;else delete td.dataset.full;raws.set(keyOf(td),rawValue??value);if(n!==null&&n!==undefined)numbers.set(keyOf(td),n);else numbers.delete(keyOf(td));td.toggleAttribute('n',n!==null&&n!==undefined);if(css!==undefined){for(const cls of [...td.classList])if(/^s[0-9]+$/.test(cls))td.classList.remove(cls);for(const key of ['font-weight','font-style','text-decoration','font-size','color','background','text-align','white-space','overflow-wrap'])td.style.removeProperty(key);const temp=document.createElement('span');temp.style.cssText=css;for(const key of Array.from(temp.style))td.style.setProperty(key,temp.style.getPropertyValue(key))};td.title=value.length>28?value:''}if(Object.keys(filters).length)applyFilters();if(current){bar.value=raw(current);updateStats();selectionFurniture()}};
        function showFilter(td){closeFilter();const c=td.dataset.c,seen=new Set();dataRows().forEach(row=>{const x=cellAt(row,c);seen.add(x?full(x):'')});const values=[...seen].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true})),menu=document.createElement('div');menu.id='filter-menu';const input=document.createElement('input');input.placeholder=T.search;const list=document.createElement('div');menu.append(input,list);const add=(label,on,fn)=>{const b=document.createElement('button');b.textContent=(on?'✓ ':'')+label;b.title=label;b.onclick=fn;list.append(b)};const draw=()=>{list.textContent='';add(T.asc,false,()=>{post({action:'sort',ascending:'1',col:c});closeFilter()});add(T.desc,false,()=>{post({action:'sort',ascending:'0',col:c});closeFilter()});const q=input.value.toLowerCase();if(!q)add(T.all,filters[c]===undefined,()=>{delete filters[c];applyFilters()});let n=0;for(const v of values){if(q&&!v.toLowerCase().includes(q))continue;if(n++>=300){const more=document.createElement('div');more.className='more';more.textContent=T.more;list.append(more);break}add(v||T.empty,filters[c]===v,()=>{filters[c]=v;applyFilters()})}};input.oninput=draw;draw();document.body.append(menu);const box=td.getBoundingClientRect();menu.style.left=Math.max(6,Math.min(box.left,innerWidth-236))+'px';menu.style.top=Math.max(6,Math.min(box.bottom+3,innerHeight-356))+'px';input.focus()}
        function neighbor(td,dx,dy){const r=+td.parentElement.dataset.row,c=+td.dataset.c;if(dx){let x=dx>0?td.nextElementSibling:td.previousElementSibling;while(x&&x.tagName!=='TD')x=dx>0?x.nextElementSibling:x.previousElementSibling;return x??(dx>0?ensure(r,c+(+td.colSpan||1)):null)}let rr=r;while(true){rr+=dy;if(rr<1)return null;let row=rowMap.get(rr);if(!row){if(dy>0)return ensure(rr,c);return null}if(!row.hidden){let x=cellAt(row,c);if(!x)for(const td of row.querySelectorAll('td'))if(+td.dataset.c<=c)x=td;return x}}}
        function cancelEdit(){editing?.remove();editing=null;bar.value=current?raw(current):'';document.activeElement?.blur()}
        function beginEdit(value){if(!current)return;closeFilter();cancelEdit();const td=current,input=document.createElement('textarea');input.rows=1;input.wrap='off';input.id='editor';input.value=value===undefined?raw(td):value;input.setAttribute('aria-label',T.edit);document.body.append(input);editing=input;const box=td.getBoundingClientRect();Object.assign(input.style,{left:(box.left+scrollX)+'px',top:(box.top+scrollY)+'px',width:Math.max(box.width,80)+'px',height:box.height+'px'});input.focus();if(value!==undefined)input.setSelectionRange(input.value.length,input.value.length)}
        function commit(value,move){if(!current||busy)return;pendingMove=move;post({action:'commit',ref:refOf(current),value})}
        window.editResult=(ok,message)=>{busy=false;document.getElementById('message').textContent=message??'';if(!ok){pendingMove=null;pendingFill=null;editing?.focus();selectionFurniture();return}cancelEdit();if(pendingFill){const fill=pendingFill;pendingFill=null;select(at(fill.r1,fill.c1));select(at(fill.r2,fill.c2),false,true)}else if(current&&pendingMove){const next=neighbor(current,...pendingMove);if(next)select(next,true)}else if(current)notifySelection();pendingMove=null};
        bar.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();commit(bar.value,null)}if(e.key==='Escape'){e.preventDefault();cancelEdit()}});
        nameBox.addEventListener('keydown',e=>{if(e.key!=='Enter')return;const parts=nameBox.value.toUpperCase().split(':'),parse=s=>{const m=/^([A-Z]{1,3})([1-9][0-9]*)$/.exec(s);if(!m)return null;let c=0;for(const ch of m[1])c=c*26+ch.charCodeAt(0)-64;return [+m[2],c]};const a=parse(parts[0]),b=parse(parts[1]??parts[0]);if(a&&b&&Math.max(a[0],b[0])<=1048576&&Math.max(a[1],b[1])<=1000&&Math.max(a[0],b[0])*Math.max(a[1],b[1])<=400000){select(ensure(...a));select(ensure(...b),true,true);nameBox.blur()}});
        document.addEventListener('mousedown',e=>{if(e.button!==0||e.target.closest('input,textarea,#filter-menu'))return;const td=e.target.closest('td');if(td){e.preventDefault();if(editing){commit(editing.value,null);return}select(td,false,e.shiftKey);dragging=true}else if(e.target.tagName==='TH'){e.preventDefault();const th=e.target;if(th.closest('thead')&&th.dataset.c){const c=+th.dataset.c;select(at(+rows[0].dataset.row,c));select(at(+rows.at(-1).dataset.row,c),false,true)}else if(th.parentElement.dataset.row){const row=th.parentElement;select(row.querySelector('td'));select(row.querySelector('td:last-child'),false,true)}}});
        document.addEventListener('mouseover',e=>{if(filling){const td=e.target.closest?.('td');if(td){const r=+td.parentElement.dataset.row,c=+td.dataset.c;fillTarget=r>filling.r2&&r-filling.r2>=c-filling.c2?{...filling,r2:r}:{...filling,c2:Math.max(c,filling.c2)};const z=at(fillTarget.r2,fillTarget.c2);if(z){const a=at(filling.r1,filling.c1).getBoundingClientRect(),b=z.getBoundingClientRect();selectionBorder.style.width=b.right-a.left+'px';selectionBorder.style.height=b.bottom-a.top+'px'}}return}const tip=e.target.closest?.('td');if(tip&&!tip.title&&full(tip).length>28)tip.title=full(tip);if(dragging){const td=e.target.closest('td');if(td&&td!==current)select(td,false,true)}});document.addEventListener('mouseup',()=>{dragging=false;if(filling){const source=filling,target=fillTarget;filling=null;fillTarget=null;pendingFill=target;post({action:'autofill',source:JSON.stringify(source),target:JSON.stringify(target)});return}finishPaint()});document.addEventListener('keyup',e=>{if(e.key.startsWith('Arrow')||e.key==='Tab')finishPaint()});
        document.addEventListener('click',e=>{if(e.target.closest('input,textarea,#filter-menu'))return;const td=e.target.closest('td');if(!td){closeFilter();return}if(!selected.has(td)||!e.isTrusted)select(td,false,e.shiftKey);if(td.classList.contains('fh'))showFilter(td);else closeFilter();if(!e.isTrusted)finishPaint()});
        document.addEventListener('dblclick',e=>{const th=e.target.closest('thead th');if(th?.dataset.c){autoWidth(+th.dataset.c);return}const td=e.target.closest('td');if(td&&!td.classList.contains('fh')){select(td);beginEdit()}});
        document.addEventListener('keydown',e=>{if(e.defaultPrevented)return;if(e.target.closest?.('#filter-menu')){if(e.key==='Escape')closeFilter();return}if(e.target===bar||e.target===nameBox)return;if(editing){if(e.key==='Escape'){e.preventDefault();cancelEdit()}else if(e.key==='Enter'||e.key==='Tab'){e.preventDefault();commit(editing.value,e.key==='Enter'?[0,e.shiftKey?-1:1]:[e.shiftKey?-1:1,0])}return}if(e.target.matches?.('input,textarea'))return;
        if(e.isTrusted&&e.metaKey)return;if((e.metaKey||e.ctrlKey)&&['c','x','v','z','a','d'].includes(e.key.toLowerCase())){e.preventDefault();const key=e.key.toLowerCase();if(key==='a'){select(rows[0].querySelector('td'));select(rows.at(-1).querySelector('td:last-child'),false,true)}else post({action:({c:'copy',x:'cut',v:'paste',z:e.shiftKey?'redo':'undo',d:'fillDown'})[key]});return}if(e.key==='Escape'){closeFilter();armPaint(false);post({action:'cancelPaint'});return}if(!current||e.metaKey||e.ctrlKey||e.altKey)return;if(e.key==='Enter'||e.key==='F2'){e.preventDefault();beginEdit();return}if(e.key==='Delete'||e.key==='Backspace'){e.preventDefault();e.stopPropagation();if(e.repeat)return;post({action:'clear'});return}const m={ArrowUp:[0,-1],ArrowDown:[0,1],ArrowLeft:[-1,0],ArrowRight:[1,0],Tab:[e.shiftKey?-1:1,0]}[e.key];if(m){e.preventDefault();const next=neighbor(current,...m);if(next)select(next,true,e.shiftKey&&e.key!=='Tab');return}if(e.key.length===1){e.preventDefault();beginEdit(e.key)}});
        window.gridSelectAll=()=>{const input=document.activeElement;if(input?.matches('input,textarea')){input.select();return}select(rows[0].querySelector('td'));select(rows.at(-1).querySelector('td:last-child'),false,true)};
        window.gridUndo=redo=>{if(document.activeElement?.matches('input,textarea'))document.execCommand(redo?'redo':'undo');else post({action:redo?'redo':'undo'})};
        window.pasteInput=text=>{const input=document.activeElement;if(input?.matches('input,textarea')){input.setRangeText(text,input.selectionStart,input.selectionEnd,'end');input.dispatchEvent(new Event('input',{bubbles:true}))}};
        window.gridClipboard=action=>{const input=document.activeElement;if(input?.matches('input,textarea')){if(action==='paste')post({action:'pasteInput'});else{post({action:'copyText',value:input.value.slice(input.selectionStart,input.selectionEnd)});if(action==='cut')input.setRangeText('',input.selectionStart,input.selectionEnd,'end')}return}post({action})};
        let firstFrozen=\(frozenRows > 0 ? "true" : "false");const FC=\(frozenCols);
        window.toggleFreeze=()=>{firstFrozen=!firstFrozen;for(const row of rows.slice(0,5)){const r=+row.dataset.row;row.classList.toggle('heading',firstFrozen&&r===1);for(const td of row.children){const frozen=firstFrozen&&r===1;td.style.position=frozen||td.tagName==='TH'||+td.dataset.c<=FC?'sticky':'';td.style.top=frozen?'66px':'';td.style.zIndex=frozen?'5':(+td.dataset.c<=FC?'3':'')}}};
        function autoWidth(c){const col=document.querySelector(`colgroup col[data-c="${c}"]`);if(!col)return;const canvas=document.createElement('canvas'),ctx=canvas.getContext('2d');ctx.font='13px -apple-system';let width=48;for(const row of rows.slice(0,1000)){const td=cellAt(row,c);if(td){ctx.font=getComputedStyle(td).font||'13px -apple-system';width=Math.max(width,ctx.measureText(full(td)).width+24)}}col.style.width=Math.min(620,width)+'px';document.querySelector('table').style.width=[...document.querySelectorAll('colgroup col')].reduce((n,x)=>n+parseFloat(x.style.width),0)+'px'}
        window.autoWidth=autoWidth;
        addEventListener('resize',selectionFurniture);addEventListener('scroll',selectionFurniture,{passive:true});
        select(rows[0]?.querySelector('td'));
        </script></body></html>
        """
        return Rendered(html: html, rows: allRows.count, columns: columns.count, shownRows: rows.count)
    }
}

// MARK: - Formulas
// A small evaluator for the most common worksheet formulas. It is used for new formulas and after edits;
// anything it doesn't know (other functions, names, shared formulas) keeps Excel's saved value.

enum FValue {
    case number(Double), text(String), bool(Bool), error(String), empty, list([FValue]), range([FValue], Int, Int)
    var number: Double? {
        switch self { case .number(let n): return n; case .bool(let b): return b ? 1 : 0; case .empty: return 0
        case .text(let s): return Double(s.trimmingCharacters(in: .whitespaces)); default: return nil }
    }
    var text: String {
        switch self { case .number(let n): return n == n.rounded() && abs(n) < 1e15 ? String(Int64(n)) : String(n); case .text(let s): return s
        case .bool(let b): return b ? "TRUE" : "FALSE"; case .error(let e): return e; case .empty: return ""; case .list(let l), .range(let l, _, _): return l.first?.text ?? "" }
    }
    var truthy: Bool? { if case .text(let s) = self { return ["TRUE", "VERO"].contains(s.uppercased()) ? true : (["FALSE", "FALSO"].contains(s.uppercased()) ? false : nil) }; return number.map { $0 != 0 } }
    var isError: Bool { if case .error = self { return true }; return false }
    var items: [FValue] { switch self { case .list(let l), .range(let l, _, _): return l; default: return [self] } }
    var shape: (Int, Int) { if case .range(_, let r, let c) = self { return (r,c) }; return (items.count,1) }
}
struct FormulaUnsupported: Error {}
struct FormulaInvalid: Error {}

enum FToken: Equatable { case num(Double), str(String), ref(Int, Int, Int), name(String), error(String), op(String), lp, rp, comma, colon }

let italianFunctions = ["SOMMA": "SUM", "MEDIA": "AVERAGE", "CONTA.NUMERI": "COUNT", "CONTA.VALORI": "COUNTA", "ARROTONDA": "ROUND", "ARROTONDA.PER.ECC": "ROUNDUP",
    "ARROTONDA.PER.DIF": "ROUNDDOWN", "ASS": "ABS", "SE": "IF", "E": "AND", "O": "OR", "NON": "NOT", "SE.ERRORE": "IFERROR", "SOMMA.SE": "SUMIF", "CONTA.SE": "COUNTIF",
    "MEDIA.SE": "AVERAGEIF", "CONCATENA": "CONCATENATE", "LUNGHEZZA": "LEN", "MAIUSC": "UPPER", "MINUSC": "LOWER", "ANNULLA.SPAZI": "TRIM", "OGGI": "TODAY",
    "VAL.ERRORE": "ISERROR", "VAL.ERR": "ISERR", "VAL.NON.DISP": "ISNA", "VAL.NUMERO": "ISNUMBER", "VAL.TESTO": "ISTEXT", "VAL.VUOTO": "ISBLANK", "VAL.LOGICO": "ISLOGICAL", "DATA.MESE": "EDATE", "FINE.MESE": "EOMONTH",
    "SOMMA.PIÙ.SE": "SUMIFS", "CONTA.PIÙ.SE": "COUNTIFS", "MEDIA.PIÙ.SE": "AVERAGEIFS", "MAX.PIÙ.SE": "MAXIFS", "MIN.PIÙ.SE": "MINIFS", "PIÙ.SE": "IFS",
    "CERCA.VERT": "VLOOKUP", "CERCA.ORIZZ": "HLOOKUP", "CERCA.X": "XLOOKUP", "INDICE": "INDEX", "CONFRONTA": "MATCH", "SINISTRA": "LEFT", "DESTRA": "RIGHT", "STRINGA.ESTRAI": "MID",
    "TROVA": "FIND", "RICERCA": "SEARCH", "SOSTITUISCI": "SUBSTITUTE", "TESTO.UNISCI": "TEXTJOIN", "VALORE": "VALUE", "DATA": "DATE", "ANNO": "YEAR", "MESE": "MONTH", "GIORNO": "DAY",
    "ADESSO": "NOW", "GIORNO.SETTIMANA": "WEEKDAY", "CONTA.VUOTE": "COUNTBLANK", "RANGO": "RANK", "RADQ": "SQRT", "POTENZA": "POWER", "RESTO": "MOD", "INT": "INT", "MEDIANA": "MEDIAN", "PRODOTTO": "PRODUCT", "VERO": "TRUE", "FALSO": "FALSE"]

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
    let characters = Array(s)
    for (i,ch) in characters.enumerated() {
        if let q = quote { out.append(ch); if ch == q { quote = nil }; continue }
        if ch == "\"" || ch == "'" { flush(ch); quote = ch; out.append(ch); continue }
        if ch.isLetter || ch.isNumber || ch == "." || ch == "_" || ch == "$" { word.append(ch); continue }
        if semicolons && ch == "," { word.append("."); continue }
        flush(ch.isWhitespace ? characters.dropFirst(i+1).first(where: { !$0.isWhitespace }) : ch)
        out.append(semicolons && ch == ";" ? "," : ch)
    }
    flush(nil)
    return out
}

final class FormulaParser {
    // Check the complete token stream before evaluating arguments: an unknown function
    // anywhere, even in an unused branch or IFERROR, requires the workbook's cached result.
    static let supportedFunctions: Set<String> = [
        "ABS", "AND", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "CONCAT", "CONCATENATE", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "DATE", "DAY", "EDATE", "EOMONTH", "FALSE", "FIND", "HLOOKUP", "IF", "IFERROR", "IFS", "INDEX", "INT", "ISBLANK", "ISERR", "ISERROR", "ISLOGICAL", "ISNA", "ISNUMBER", "ISTEXT", "LEFT", "LEN", "LOWER", "MATCH", "MAX", "MAXIFS", "MEDIAN", "MID", "MIN", "MINIFS", "MOD", "MONTH", "NOT", "NOW", "OR", "POWER", "PRODUCT", "RANK", "RANK.EQ", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "SEARCH", "SQRT", "SUBSTITUTE", "SUM", "SUMIF", "SUMIFS", "TEXTJOIN", "TODAY", "TRIM", "TRUE", "UPPER", "VALUE", "VLOOKUP", "WEEKDAY", "XLOOKUP", "YEAR"
    ]
    static func functionName(_ name: String) -> String {
        let clean = name.uppercased().replacingOccurrences(of:"_XLFN.",with:"")
        return italianFunctions[clean] ?? clean
    }
    let book: Book, sheet: Int
    var tokens: [FToken] = [], pos = 0
    var validating = false
    var patterns: [String: NSRegularExpression] = [:]
    init(book: Book, sheet: Int) { self.book = book; self.sheet = sheet }

    func tokenize(_ s: String) throws {
        let chars = Array(s); var i = 0
        func refToken(_ word: String, sheet: Int) -> FToken? {
            let clean = word.replacingOccurrences(of: "$", with: "").uppercased()
            guard clean.range(of: "^[A-Z]{1,3}([1-9][0-9]{0,6})?$", options: .regularExpression) != nil else { return nil }
            return .ref(sheet, rowNumber(clean), column(clean))
        }
        func sheetIndex(_ name: String) throws -> Int { guard let i = book.sheets.firstIndex(where: { $0.name.caseInsensitiveCompare(name) == .orderedSame }) else { throw FormulaUnsupported() }; return i }
        while i < chars.count {
            let ch = chars[i]
            if ch.isWhitespace { i += 1; continue }
            if ch == "#" {
                let rest = String(chars[i...]); guard let e = ["#DIV/0!", "#VALUE!", "#REF!", "#N/A", "#NUM!", "#NAME?", "#NULL!"].first(where: { rest.hasPrefix($0) }) else { throw FormulaInvalid() }
                tokens.append(.error(e)); i += e.count; continue
            }
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
                let isCall = chars.dropFirst(i).first(where:{ !$0.isWhitespace }) == "("
                if i < chars.count && chars[i] == "!" {
                    i += 1; var ref = ""
                    while i < chars.count, chars[i].isLetter || chars[i].isNumber || chars[i] == "$" { ref.append(chars[i]); i += 1 }
                    guard let t = refToken(ref, sheet: try sheetIndex(word)) else { throw FormulaUnsupported() }
                    tokens.append(t)
                } else if let t = refToken(word, sheet: sheet), (rowNumber(word.replacingOccurrences(of: "$", with: "")) > 0 || (i < chars.count && chars[i] == ":") || tokens.last == .colon), !isCall { tokens.append(t) }
                else {
                    if isCall && !validating && !Self.supportedFunctions.contains(Self.functionName(word)) { throw FormulaUnsupported() }
                    tokens.append(.name(word.uppercased()))
                }
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
        if !validating {
            for i in tokens.indices where i+1<tokens.count && tokens[i+1] == .lp {
                if case .name(let name) = tokens[i], !Self.supportedFunctions.contains(Self.functionName(name)) { throw FormulaUnsupported() }
            }
        }
        let value = try comparison()
        guard pos == tokens.count else { throw FormulaInvalid() }
        if case .range(let l, _, _) = value { return l.first ?? .empty }; if case .list(let l) = value { return l.first ?? .empty }
        return value
    }
    var peek: FToken? { pos < tokens.count ? tokens[pos] : nil }
    func take(_ t: FToken) -> Bool { if peek == t { pos += 1; return true }; return false }

    func comparison() throws -> FValue {
        var left = try concat()
        while case .op(let o)? = peek, ["=", "<>", "<", ">", "<=", ">="].contains(o) {
            pos += 1; let right = try concat()
            if !left.isError { left = right.isError ? right : .bool(compare(left,o,right)) }
        }
        return left
    }
    func concat() throws -> FValue {
        var left = try additive()
        while take(.op("&")) { let right = try additive(); if !left.isError { left = right.isError ? right : .text(left.text+right.text) } }
        return left
    }
    func arithmetic(_ a: FValue, _ b: FValue, _ f: (Double, Double) -> FValue) -> FValue {
        if a.isError { return a }; if b.isError { return b }
        guard let x = a.number, let y = b.number else { return .error("#VALUE!") }
        return f(x, y).excelValue
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
        case .num(let n): return FValue.number(n).excelValue
        case .str(let s): return .text(s)
        case .error(let e): return .error(e)
        case .lp: let v = try comparison(); guard take(.rp) else { throw FormulaInvalid() }; return v
        case .ref(let s, let r, let c):
            if take(.colon) {
                guard case .ref(let s2, let r2, let c2)? = peek, s2 == sheet || s2 == s, (r == 0) == (r2 == 0) else { throw FormulaInvalid() }
                pos += 1
                if validating { return .empty }
                guard c <= 16384, c2 <= 16384, r <= 1048576, r2 <= 1048576 else { return .error("#REF!") }
                let lo = r == 0 ? 1 : min(r,r2), hi = r == 0 ? max(1, try book.index(s).maxRow) : max(r,r2)
                let left = min(c,c2), right = max(c,c2)
                return .range(try book.rangeValues(s, lo, left, hi, right), hi-lo+1, right-left+1)
            }
            if validating { return .empty }
            guard r > 0, r <= 1048576, c <= 16384 else { return .error("#REF!") }
            return try book.value(s, r, c)
        case .name(let rawName):
            let name = Self.functionName(rawName)
            if name == "TRUE" || name == "FALSE" { _ = take(.lp) && take(.rp); return .bool(name == "TRUE") }
            guard take(.lp) else { if validating { return .empty }; throw FormulaUnsupported() }
            var args: [FValue] = []
            if !take(.rp) {
                repeat { args.append(peek == .comma || peek == .rp ? .empty : try comparison()) } while take(.comma)
                guard take(.rp) else { throw FormulaInvalid() }
            }
            return validating ? .empty : try call(name,args)
        default: throw FormulaInvalid()
        }
    }
    func compare(_ a: FValue, _ op: String, _ b: FValue) -> Bool {
        func scalar(_ v: FValue) -> FValue { switch v { case .range(let l,_,_), .list(let l): return l.first ?? .empty; default: return v } }
        func zero(like v: FValue) -> FValue { switch v { case .text: return .text(""); case .bool: return .bool(false); default: return .number(0) } }
        var x = scalar(a), y = scalar(b)
        if case .empty = x { x = zero(like:y) }; if case .empty = y { y = zero(like:x) }
        func typeOrder(_ v: FValue) -> Int { switch v { case .number: return 0; case .text: return 1; case .bool: return 2; default: return 3 } }
        let order: Int
        switch (x,y) {
        case (.number(let a),.number(let b)): order = a<b ? -1 : (a>b ? 1 : 0)
        case (.text(let a),.text(let b)): order = a.caseInsensitiveCompare(b).rawValue
        case (.bool(let a),.bool(let b)): order = a == b ? 0 : (a ? 1 : -1)
        default: order = typeOrder(x)<typeOrder(y) ? -1 : (typeOrder(x)>typeOrder(y) ? 1 : x.text.caseInsensitiveCompare(y.text).rawValue)
        }
        switch op { case "=": return order == 0; case "<>": return order != 0; case "<": return order < 0; case ">": return order > 0; case "<=": return order <= 0; default: return order >= 0 }
    }
    func matches(_ value: FValue, _ criterion: FValue) -> Bool {
        var text = criterion.text, op = "="
        for candidate in ["<>", "<=", ">=", "<", ">", "="] where text.hasPrefix(candidate) { op = candidate; text = String(text.dropFirst(candidate.count)); break }
        if (op == "=" || op == "<>"), text.contains("*") || text.contains("?") || text.contains("~") {
            guard case .text(let actual) = value else { return op == "<>" }
            let yes = wildcard(text, actual)
            return op == "=" ? yes : !yes
        }
        let target: FValue = automaticNumber(text).map { .number($0) } ?? .text(text)
        if case .empty = value { if automaticNumber(text) == 0 { return compare(.number(0),op,.number(0)) }; return op == "=" ? text.isEmpty : (op == "<>" ? !text.isEmpty : false) }
        if case .number = target, case .text(let s) = value {
            guard let n = automaticNumber(s) else { return op == "<>" }; return compare(.number(n),op,target)
        }
        if case .number = target, case .bool(let b) = value { return compare(.number(b ? 1 : 0),op,target) }
        if case .text = target, !(value.isTextual) { return compare(.text(value.text),op,target) }
        return compare(value, op, target)
    }
    func wildcard(_ pattern: String, _ value: String) -> Bool {
        if patterns[pattern] == nil {
            var re = "^", escaped = false
            for ch in pattern {
                if escaped { re += NSRegularExpression.escapedPattern(for: String(ch)); escaped = false }
                else if ch == "~" { escaped = true }
                else if ch == "*" { re += ".*" } else if ch == "?" { re += "." }
                else { re += NSRegularExpression.escapedPattern(for: String(ch)) }
            }
            if escaped { re += "~" }; re += "$"
            patterns[pattern] = try? NSRegularExpression(pattern: re, options: [.caseInsensitive, .dotMatchesLineSeparators])
        }
        return patterns[pattern]?.firstMatch(in: value, range: NSRange(value.startIndex..., in: value)) != nil
    }
    func lookupIndex(_ key: FValue, _ values: [FValue], mode: Int, wildcards: Bool = true) -> Int? {
        var found: Int?
        for (i,v) in values.enumerated() {
            if v.isError { continue }
            let sameType = (key.isTextual == v.isTextual)
            if mode == 0 {
                let equal = wildcards && key.isTextual ? (v.isTextual && wildcard(key.text,v.text)) : (sameType && compare(v,"=",key))
                if equal { return i }
            } else if sameType && compare(v, mode == 1 ? "<=" : ">=", key) { found = i }
        }
        return found
    }
    func numbers(_ args: [FValue]) throws -> [Double] {
        var out: [Double] = []
        for arg in args {
            if case .range = arg { for v in arg.items { if v.isError { throw FormulaErrorValue(v) }; if case .number(let n) = v { out.append(n) } } }
            else if case .list(let l) = arg { for v in l { if case .error = v { throw FormulaErrorValue(v) }; if case .number(let n) = v { out.append(n) } } }
            else if arg.isError { throw FormulaErrorValue(arg) }
            else if case .empty = arg { continue }
            else if let n = arg.number { out.append(n) } else { throw FormulaErrorValue(.error("#VALUE!")) }
        }
        return out
    }
    /// Convert Excel serials at month boundaries, preserving the 1900 serial offset.
    /// EDATE/EOMONTH clamp to Gregorian month ends (February 1900 ends on serial 59).
    func shiftedMonth(_ start: Double, months: Int, endOfMonth: Bool) -> FValue {
        let maximum = book.date1904 ? 2957003.0 : 2958465.0
        guard start.isFinite,start>=0,start.rounded(.down)<=maximum else { return .error("#NUM!") }
        let serial = start.rounded(.down)
        var calendar = Calendar(identifier:.gregorian); calendar.timeZone = TimeZone(secondsFromGMT:0)!
        var year: Int, month: Int, day: Int
        if !book.date1904 && serial == 0 { year = 1900; month = 1; day = 0 }
        else if !book.date1904 && serial == 60 { year = 1900; month = 2; day = 29 }
        else {
            let adjusted = !book.date1904 && serial<60 ? serial+1 : serial
            let date = Date(timeIntervalSince1970:(adjusted-(book.date1904 ? 24107 : 25569))*86400)
            let parts = calendar.dateComponents([.year,.month,.day],from:date)
            year = parts.year!; month = parts.month!; day = parts.day!
        }
        let target = year*12+month-1+months
        guard target >= (book.date1904 ? 1904 : 1900)*12, target < 10000*12 else { return .error("#NUM!") }
        year = target/12; month = target%12+1
        guard let first = calendar.date(from:DateComponents(year:year,month:month,day:1)), let days = calendar.range(of:.day,in:.month,for:first)?.count else { return .error("#NUM!") }
        let lastDay = days
        let raw = first.timeIntervalSince1970/86400+(book.date1904 ? 24107 : 25569)
        let result = raw-(!book.date1904 && raw<61 ? 1 : 0)+Double((endOfMonth ? lastDay : min(day,lastDay))-1)
        return result>=0 && result<=maximum ? .number(result) : .error("#NUM!")
    }
    func call(_ rawName: String, _ args: [FValue]) throws -> FValue {
        let name = Self.functionName(rawName)
        guard Self.supportedFunctions.contains(name) else { throw FormulaUnsupported() }
        return try calculate(name,args).excelValue
    }
    private func calculate(_ name: String, _ args: [FValue]) throws -> FValue {
        func arg(_ i: Int) -> FValue { args.indices.contains(i) ? args[i] : .empty }
        func num(_ i: Int) throws -> Double { guard let n = arg(i).number else { throw FormulaErrorValue(arg(i).isError ? arg(i) : .error("#VALUE!")) }; return n }
        func integer(_ i: Int) throws -> Int { let n = try num(i); guard n.isFinite, abs(n) < 1e9 else { throw FormulaErrorValue(.error("#NUM!")) }; return Int(n) }
        do {
            if !["IF", "IFERROR", "IFS", "COUNT", "COUNTA", "COUNTBLANK", "SUMIF", "COUNTIF", "AVERAGEIF", "SUMIFS", "COUNTIFS", "AVERAGEIFS", "MAXIFS", "MINIFS", "XLOOKUP", "ISERROR", "ISERR", "ISNA", "ISNUMBER", "ISTEXT", "ISBLANK", "ISLOGICAL"].contains(name), let e = args.first(where: \.isError) { return e }
            let counts: [String:ClosedRange<Int>] = ["ISERROR":1...1,"ISERR":1...1,"ISNA":1...1,"ISNUMBER":1...1,"ISTEXT":1...1,"ISBLANK":1...1,"ISLOGICAL":1...1,"EDATE":2...2,"EOMONTH":2...2,"TRUE":0...0,"FALSE":0...0,"SUMIFS":3...255,"COUNTIFS":2...254,"AVERAGEIFS":3...255,"MAXIFS":3...255,"MINIFS":3...255,"IFS":2...254,"VLOOKUP":3...4,"HLOOKUP":3...4,"XLOOKUP":3...4,"INDEX":2...3,"MATCH":2...3,"LEFT":1...2,"RIGHT":1...2,"MID":3...3,"FIND":2...3,"SEARCH":2...3,"SUBSTITUTE":3...4,"TEXTJOIN":3...254,"VALUE":1...1,"DATE":3...3,"YEAR":1...1,"MONTH":1...1,"DAY":1...1,"NOW":0...0,"TODAY":0...0,"WEEKDAY":1...2,"COUNTBLANK":1...1,"RANK":2...3,"RANK.EQ":2...3]
            if let allowed = counts[name], !allowed.contains(args.count) { return .error("#VALUE!") }
            switch name {
            case "TRUE": return .bool(true)
            case "FALSE": return .bool(false)
            case "ISERROR", "ISERR", "ISNA", "ISNUMBER", "ISTEXT", "ISBLANK", "ISLOGICAL":
                func check(_ value: FValue) -> FValue {
                    switch name {
                    case "ISERROR": return .bool(value.isError)
                    case "ISERR": if case .error(let error) = value { return .bool(error != "#N/A") }; return .bool(false)
                    case "ISNA": if case .error("#N/A") = value { return .bool(true) }; return .bool(false)
                    case "ISNUMBER": if case .number = value { return .bool(true) }; return .bool(false)
                    case "ISTEXT": if case .text = value { return .bool(true) }; return .bool(false)
                    case "ISBLANK": if case .empty = value { return .bool(true) }; return .bool(false)
                    default: if case .bool = value { return .bool(true) }; return .bool(false)
                    }
                }
                if case .range(let values,let r,let c) = arg(0) { return .range(values.map(check),r,c) }
                return check(arg(0))
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
            case "IF": if arg(0).isError { return arg(0) }; guard let c = arg(0).truthy else { return .error("#VALUE!") }; let result = c ? arg(1) : (args.count > 2 ? arg(2) : .bool(false)); if case .empty = result { return .number(0) }; return result
            case "IFERROR": let result = arg(0).isError ? arg(1) : arg(0); if case .empty = result { return .number(0) }; return result
            case "AND", "OR":
                if let e = args.flatMap(\.items).first(where: \.isError) { return e }
                let flags = args.flatMap(\.items).compactMap { v -> Bool? in if case .empty = v { return nil }; return v.truthy }
                return .bool(name == "AND" ? !flags.contains(false) : flags.contains(true))
            case "NOT": guard let c = arg(0).truthy else { return .error("#VALUE!") }; return .bool(!c)
            case "CONCATENATE", "CONCAT": return .text(args.flatMap(\.items).map(\.text).joined())
            case "LEN": return .number(Double(arg(0).text.count))
            case "UPPER": return .text(arg(0).text.uppercased())
            case "LOWER": return .text(arg(0).text.lowercased())
            case "TRIM": return .text(arg(0).text.split(separator: " ").joined(separator: " "))
            case "SUMIF", "COUNTIF", "AVERAGEIF":
                let range = arg(0).items, values = args.count > 2 ? arg(2).items : range
                let picked = range.indices.filter { matches(range[$0], arg(1)) }
                if arg(1).isError { return arg(1) }
                if name == "COUNTIF" { return .number(Double(picked.count)) }
                 if let e = picked.compactMap({ values.indices.contains($0) ? values[$0] : nil }).first(where: \.isError) { return e }
                let nums = picked.compactMap { i -> Double? in guard values.indices.contains(i), case .number(let n) = values[i] else { return nil }; return n }
                if name == "SUMIF" { return .number(nums.reduce(0, +)) }
                return nums.isEmpty ? .error("#DIV/0!") : .number(nums.reduce(0, +) / Double(nums.count))
            case "SUMIFS", "COUNTIFS", "AVERAGEIFS", "MAXIFS", "MINIFS":
                let start = name == "COUNTIFS" ? 0 : 1
                guard args.count >= start+2, (args.count-start)%2 == 0 else { return .error("#VALUE!") }
                let base = arg(start).items, shape = arg(start).shape
                let values = name == "COUNTIFS" ? base : arg(0).items
                guard name == "COUNTIFS" || arg(0).shape == shape else { return .error("#VALUE!") }
                for j in stride(from: start, to: args.count, by: 2) {
                    if arg(j).shape != shape { return .error("#VALUE!") }; if arg(j+1).isError { return arg(j+1) }
                }
                let ranges = stride(from: start, to: args.count, by: 2).map { (arg($0).items,arg($0+1)) }
                var picked: [FValue] = [], count = 0
                for i in base.indices where ranges.allSatisfy({ matches($0.0[i], $0.1) }) { count += 1; picked.append(values[i]) }
                if name == "COUNTIFS" { return .number(Double(count)) }
                if let e = picked.first(where: \.isError) { return e }
                let ns = picked.compactMap { v -> Double? in if case .number(let n) = v { return n }; return nil }
                if name == "MAXIFS" { return .number(ns.max() ?? 0) }; if name == "MINIFS" { return .number(ns.min() ?? 0) }
                if name == "AVERAGEIFS" { return ns.isEmpty ? .error("#DIV/0!") : .number(ns.reduce(0,+)/Double(ns.count)) }
                return .number(ns.reduce(0,+))
            case "IFS":
                guard args.count >= 2, args.count%2 == 0 else { return .error("#VALUE!") }
                for j in stride(from: 0, to: args.count, by: 2) { if arg(j).isError { return arg(j) }; guard let b = arg(j).truthy else { return .error("#VALUE!") }; if b { if case .empty = arg(j+1) { return .number(0) }; return arg(j+1) } }
                return .error("#N/A")
            case "COUNTBLANK": return .number(Double(arg(0).items.filter { if case .empty = $0 { return true }; if case .text("") = $0 { return true }; return false }.count))
            case "RANK", "RANK.EQ":
                let x = try num(0), ascending = args.count > 2 ? (try num(2)) != 0 : false
                if let e = arg(1).items.first(where: \.isError) { return e }
                let ns = arg(1).items.compactMap { v -> Double? in if case .number(let n) = v { return n }; return nil }
                guard ns.contains(x) else { return .error("#N/A") }
                return .number(Double(1 + ns.filter { ascending ? $0 < x : $0 > x }.count))
            case "INDEX":
                let (rows,cols) = arg(0).shape, values = arg(0).items
                var r = try integer(1), c = args.count > 2 ? try integer(2) : (rows>1 && cols>1 ? 0 : 1)
                if rows == 1 && args.count == 2 { c = r; r = 1 }
                guard r >= 0, c >= 0 else { return .error("#VALUE!") }
                guard r <= rows, c <= cols else { return .error("#REF!") }
                if r == 0 || c == 0 {
                    let rs = r == 0 ? Array(1...rows) : [r], cs = c == 0 ? Array(1...cols) : [c]
                    return .range(rs.flatMap { rr in cs.map { cc in values[(rr-1)*cols+cc-1] } }, rs.count,cs.count)
                }
                let v = values[(r-1)*cols+c-1]; if case .empty = v { return .number(0) }; return v
            case "MATCH":
                let mode = args.count > 2 ? try integer(2) : 1
                guard [-1,0,1].contains(mode), arg(1).shape.0 == 1 || arg(1).shape.1 == 1 else { return .error("#N/A") }
                return lookupIndex(arg(0), arg(1).items, mode: mode).map { .number(Double($0+1)) } ?? .error("#N/A")
            case "VLOOKUP", "HLOOKUP":
                let (rows,cols) = arg(1).shape, values = arg(1).items, vertical = name == "VLOOKUP"
                let n = try integer(2), limit = vertical ? cols : rows
                guard n >= 1 else { return .error("#VALUE!") }; guard n <= limit else { return .error("#REF!") }
                let keys = vertical ? (0..<rows).map { values[$0*cols] } : Array(values.prefix(cols))
                let mode = args.count > 3 && arg(3).truthy == false ? 0 : 1
                guard let i = lookupIndex(arg(0), keys, mode: mode) else { return .error("#N/A") }
                let v = values[vertical ? i*cols+n-1 : (n-1)*cols+i]; if case .empty = v { return .number(0) }; return v
            case "XLOOKUP":
                if let e = args.prefix(3).first(where: \.isError) { return e }
                guard arg(1).shape == arg(2).shape, arg(1).shape.0 == 1 || arg(1).shape.1 == 1 else { return .error("#VALUE!") }
                guard let i = lookupIndex(arg(0), arg(1).items, mode: 0, wildcards: false) else { return args.count > 3 ? arg(3) : .error("#N/A") }
                let v = arg(2).items[i]; if case .empty = v { return .number(0) }; return v
            case "LEFT", "RIGHT", "MID":
                let chars = Array(arg(0).text)
                let start = name == "MID" ? try integer(1) : 1
                let count = name == "MID" ? try integer(2) : (args.count > 1 ? try integer(1) : 1)
                guard start >= 1, count >= 0 else { return .error("#VALUE!") }
                if name == "RIGHT" { return .text(String(chars.suffix(count))) }
                return .text(String(chars.dropFirst(min(chars.count,start-1)).prefix(count)))
            case "FIND", "SEARCH":
                let needle = arg(0).text, text = arg(1).text, start = args.count > 2 ? try integer(2) : 1
                guard start > 0, start <= text.count+1 else { return .error("#VALUE!") }
                if needle.isEmpty { return .number(Double(start)) }
                let tail = String(text.dropFirst(start-1))
                if name == "SEARCH" {
                    var pattern = "", escaped = false
                    for ch in needle { if escaped { pattern += NSRegularExpression.escapedPattern(for: String(ch)); escaped = false } else if ch == "~" { escaped = true } else if ch == "*" { pattern += ".*?" } else if ch == "?" { pattern += "." } else { pattern += NSRegularExpression.escapedPattern(for: String(ch)) } }
                    if escaped { pattern += "~" }
                    if let range = tail.range(of: pattern, options: [.regularExpression,.caseInsensitive]) { return .number(Double(start+tail.distance(from: tail.startIndex,to: range.lowerBound))) }
                } else if let range = tail.range(of: needle) { return .number(Double(start+tail.distance(from: tail.startIndex,to: range.lowerBound))) }
                return .error("#VALUE!")
            case "SUBSTITUTE":
                let text = arg(0).text, old = arg(1).text, new = arg(2).text
                if old.isEmpty { return .text(text) }
                if args.count < 4 { return .text(text.replacingOccurrences(of: old,with: new)) }
                let instance = try integer(3); guard instance > 0 else { return .error("#VALUE!") }
                var result = "", tail = text[...], count = 0
                while let range = tail.range(of: old) { count += 1; result += tail[..<range.lowerBound]; result += count == instance ? new : old; tail = tail[range.upperBound...] }
                return .text(result + tail)
            case "TEXTJOIN":
                guard let ignore = arg(1).truthy else { return .error("#VALUE!") }
                let values = args.dropFirst(2).flatMap(\.items)
                if let e = values.first(where: \.isError) { return e }
                let result = values.filter { !ignore || !$0.text.isEmpty }.map(\.text).joined(separator: arg(0).text)
                return result.count > 32767 ? .error("#VALUE!") : .text(result)
            case "VALUE":
                var text = arg(0).text.trimmingCharacters(in:.whitespacesAndNewlines)
                let negative = text.hasPrefix("(") && text.hasSuffix(")")
                if negative { text = String(text.dropFirst().dropLast()) }
                if let first = text.first, "€$£¥".contains(first) { text.removeFirst(); text = text.trimmingCharacters(in:.whitespaces) }
                let percent = text.hasSuffix("%")
                if percent { text.removeLast(); text = text.trimmingCharacters(in:.whitespaces) }
                let f = NumberFormatter(); f.locale = .current; f.numberStyle = .decimal
                let group = NSRegularExpression.escapedPattern(for:f.groupingSeparator ?? ","), decimal = NSRegularExpression.escapedPattern(for:f.decimalSeparator ?? ".")
                let pattern = "^[+-]?[0-9]{1,3}(?:"+group+"[0-9]{3})+(?:"+decimal+"[0-9]+)?$"
                var number: Double?
                if text.range(of:pattern,options:.regularExpression) != nil { number = f.number(from:text)?.doubleValue }
                if number == nil { number = automaticNumber(text) }
                if let n = number { return .number(n*(negative ? -1 : 1)/(percent ? 100 : 1)) }
                let parts = text.split(separator:":",omittingEmptySubsequences:false)
                if (2...3).contains(parts.count), let h = Double(parts[0]), let m = Double(parts[1]), let sec = parts.count == 3 ? Double(parts[2]) : 0, h>=0,h<24,m>=0,m<60,sec>=0,sec<60 { return .number((h*3600+m*60+sec)/86400) }
                let df = DateFormatter(); df.locale = .current; df.timeZone = TimeZone(secondsFromGMT:0); df.isLenient = false; df.dateStyle = .short
                var date = df.date(from:text)
                if date == nil { df.dateStyle = .none; df.dateFormat = "yyyy-MM-dd"; date = df.date(from:text) }
                if let date {
                    let raw = date.timeIntervalSince1970/86400 + (book.date1904 ? 24107 : 25569)
                    return .number(raw-(!book.date1904 && raw<61 ? 1 : 0))
                }
                return .error("#VALUE!")
            case "EDATE", "EOMONTH":
                var start = arg(0).number
                if start == nil, case .text = arg(0) { start = try call("VALUE",[arg(0)]).number }
                guard let date = start else { return .error("#VALUE!") }
                return shiftedMonth(date,months:try integer(1),endOfMonth:name == "EOMONTH")
            case "DATE":
                var y = try integer(0); let m = try integer(1), d = try integer(2)
                if y >= 0 && y < 1900 { y += 1900 }
                guard y >= 1900, y <= 9999, abs(m) < 120000, abs(d) < 3660000 else { return .error("#NUM!") }
                var cal = Calendar(identifier: .gregorian); cal.timeZone = TimeZone(secondsFromGMT: 0)!
                guard let jan = cal.date(from:DateComponents(year:y,month:1,day:1)), let month = cal.date(byAdding:.month,value:m-1,to:jan) else { return .error("#NUM!") }
                let raw = month.timeIntervalSince1970/86400 + (book.date1904 ? 24107 : 25569)
                let serial = raw - (!book.date1904 && raw < 61 ? 1 : 0) + Double(d-1)
                return serial >= (book.date1904 ? -1461 : 0) && serial < (book.date1904 ? 2957004 : 2958466) ? .number(serial) : .error("#NUM!")
            case "YEAR", "MONTH", "DAY", "WEEKDAY":
                let serial = try num(0).rounded(.down)
                guard serial >= 0, serial < 2958466 else { return .error("#NUM!") }
                var cal = Calendar(identifier: .gregorian); cal.timeZone = TimeZone(secondsFromGMT: 0)!
                let adjusted = !book.date1904 && serial < 60 ? serial+1 : serial
                let date = Date(timeIntervalSince1970:(adjusted-(book.date1904 ? 24107 : 25569))*86400)
                if name == "WEEKDAY" {
                    let mode = args.count > 1 ? try integer(1) : 1
                    let wd = !book.date1904 && serial < 61 ? (Int(serial)+6)%7 : cal.component(.weekday,from:date)-1
                    if mode == 1 { return .number(Double(wd+1)) }; if mode == 3 { return .number(Double((wd+6)%7)) }
                    let first = mode == 2 || mode == 11 ? 1 : (mode >= 12 && mode <= 17 ? mode-10 : -1)
                    guard first >= 0 else { return .error("#NUM!") }; return .number(Double((wd-first+7)%7+1))
                }
                if !book.date1904 && serial == 60 { return .number(name == "YEAR" ? 1900 : (name == "MONTH" ? 2 : 29)) }
                if !book.date1904 && serial == 0 { return .number(name == "YEAR" ? 1900 : (name == "MONTH" ? 1 : 0)) }
                return .number(Double(cal.component(name == "YEAR" ? .year : (name == "MONTH" ? .month : .day),from:date)))
            case "NOW", "TODAY":
                let now = book.calculationTime, local = now.timeIntervalSince1970 + Double(TimeZone.current.secondsFromGMT(for:now))
                let serial = local/86400 + (book.date1904 ? 24107 : 25569)
                return .number(name == "TODAY" ? serial.rounded(.down) : serial)
            default: throw FormulaUnsupported()
            }
        } catch let error as FormulaErrorValue { return error.value }
    }
}
struct FormulaErrorValue: Error { let value: FValue; init(_ v: FValue) { value = v } }
extension FValue {
    var excelValue: FValue {
        switch self {
        case .number(let n): return n.isFinite ? self : .error("#NUM!")
        case .range(let values,let r,let c): return .range(values.map(\.excelValue),r,c)
        case .list(let values): return .list(values.map(\.excelValue))
        default: return self
        }
    }
    var isTextual: Bool { if case .text = self { return true }; return false } }

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

/// TSV quoting matches Excel/Numbers, including embedded tabs, newlines and doubled quotes.
func parseTSV(_ input: String) -> [[String]] {
    let chars = Array(input.replacingOccurrences(of:"\r\n",with:"\n").replacingOccurrences(of:"\r",with:"\n"))
    var result: [[String]] = [], row: [String] = [], field = "", quoted = false, i = 0
    while i<chars.count {
        let ch = chars[i]
        if ch == "\"" && (quoted || field.isEmpty) {
            if quoted && i+1<chars.count && chars[i+1] == "\"" { field.append(ch); i += 1 }
            else { quoted.toggle() }
        } else if !quoted && (ch == "\t" || ch == "\n") {
            row.append(field); field = ""
            if ch == "\n" { result.append(row); row = [] }
        } else { field.append(ch) }
        i += 1
    }
    if !field.isEmpty || !row.isEmpty || result.isEmpty { row.append(field); result.append(row) }
    return result
}

/// Keep the native color panel/tracking/accessibility, with Sheets-style glyphs and a color underline.
final class ToolbarColorWell: NSColorWell {
    var isTextGlyph: Bool { identifier?.rawValue == "foreground" }
    var colorBarRect: NSRect { NSRect(x:4,y:2,width:max(0,bounds.width-8),height:3) }
    override func draw(_ dirtyRect: NSRect) {
        if isActive {
            NSColor.controlAccentColor.withAlphaComponent(0.12).setFill()
            NSBezierPath(roundedRect:bounds.insetBy(dx:1,dy:1),xRadius:4,yRadius:4).fill()
        }
        let ink = isEnabled ? NSColor.labelColor : .disabledControlTextColor
        if isTextGlyph {
            let attributes: [NSAttributedString.Key:Any] = [.font:NSFont.systemFont(ofSize:16,weight:.semibold),.foregroundColor:ink]
            let text = "A" as NSString, size = text.size(withAttributes:attributes)
            text.draw(at:NSPoint(x:(bounds.width-size.width)/2,y:6),withAttributes:attributes)
        } else if let symbol = NSImage(systemSymbolName:"paintbucket.fill",accessibilityDescription:toolTip),
                  let image = symbol.withSymbolConfiguration(NSImage.SymbolConfiguration(paletteColors:[ink])) {
            image.draw(in:NSRect(x:5,y:8,width:17,height:16))
        } else {
            // A small native vector fallback keeps the bucket available on macOS 13 too.
            ink.setStroke(); ink.setFill()
            let bucket = NSBezierPath(); bucket.move(to:NSPoint(x:6,y:15)); bucket.line(to:NSPoint(x:13,y:22)); bucket.line(to:NSPoint(x:21,y:14)); bucket.line(to:NSPoint(x:14,y:8)); bucket.close(); bucket.lineWidth = 1.3; bucket.stroke()
            let handle = NSBezierPath(); handle.move(to:NSPoint(x:9,y:18)); handle.curve(to:NSPoint(x:16,y:20),controlPoint1:NSPoint(x:6,y:25),controlPoint2:NSPoint(x:12,y:27)); handle.lineWidth = 1.3; handle.stroke()
            let drop = NSBezierPath(); drop.move(to:NSPoint(x:22,y:13)); drop.curve(to:NSPoint(x:22,y:7),controlPoint1:NSPoint(x:19,y:10),controlPoint2:NSPoint(x:19,y:7)); drop.curve(to:NSPoint(x:22,y:13),controlPoint1:NSPoint(x:25,y:7),controlPoint2:NSPoint(x:25,y:10)); drop.fill()
        }
        (isEnabled ? color : color.withAlphaComponent(0.4)).setFill()
        NSBezierPath(roundedRect:colorBarRect,xRadius:0.5,yRadius:0.5).fill()
    }
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
    var zoomLabel: NSPopUpButton!; var search: NSSearchField!; var detail: NSTextView!; var notice: NSTextField!
    let progress = NSProgressIndicator()
    let recentMenu = NSMenu(title: L("Apri recenti", "Open Recent"))
    var selectedRef = ""
    var selectedRange = CellRange()
    var lastEditError = ""
    let editUndo = SheetUndoManager()
    var committedActions = 0
    var savedGeneration = 0
    var zoom = 1.0
    var pendingURL: URL?
    var loadID = UUID()
    var loading = false
    let loader = DispatchQueue(label: "xlsxlite.load", qos: .userInitiated)
    var hint: String { L("Invio o doppio clic: modifica · Frecce: muoviti · Intestazioni ▾: filtra", "Return or double-click: edit · Arrows: move · Headers ▾: filter") }

    func updateEditState() {
        window.isDocumentEdited = book?.dirty ?? false
        notice.stringValue = (book?.dirty == true) ? L("Modifiche non salvate", "Unsaved changes") : hint
    }
    @objc func clearFilters() { web.evaluateJavaScript("window.clearFilters&&clearFilters()", completionHandler: nil) }
    @objc func undoMenu() { if let text = window.firstResponder as? NSTextView { text.undoManager?.undo(); return }; web.evaluateJavaScript("gridUndo(false)",completionHandler:nil) }
    @objc func redoMenu() { if let text = window.firstResponder as? NSTextView { text.undoManager?.redo(); return }; web.evaluateJavaScript("gridUndo(true)",completionHandler:nil) }
    @objc func selectAllCells() { if let text = window.firstResponder as? NSTextView { text.selectAll(nil); return }; web.evaluateJavaScript("gridSelectAll()",completionHandler:nil) }
    @objc func undoEdit() { editUndo.undo(); updateEditState(); web.evaluateJavaScript("editResult(true,'')",completionHandler:nil) }
    @objc func redoEdit() { editUndo.redo(); updateEditState(); web.evaluateJavaScript("editResult(true,'')",completionHandler:nil) }
    func validateMenuItem(_ item: NSMenuItem) -> Bool {
        switch item.action {
        case #selector(undoEdit): return NSApp.modalWindow != nil || editUndo.canUndo
        case #selector(redoEdit): return NSApp.modalWindow != nil || editUndo.canRedo
        case #selector(saveFile), #selector(saveCopy), #selector(printSheet), #selector(clearFilters), #selector(nextSheet), #selector(previousSheet): return book != nil && !loading
        case #selector(editCell): return book != nil && !selectedRef.isEmpty && !loading
        default: return true
        }
    }
    func restoreSheet(_ index: Int, _ data: Data, dirty: Bool, refs: [String] = [], generation: Int? = nil) {
        guard let book else { return }
        do {
            let old = try book.sheetDocument(index).xmlData, wasDirty = book.dirty, inverseGeneration = savedGeneration
            editUndo.registerUndo(withTarget:self) { $0.restoreSheet(index,old,dirty:wasDirty,refs:refs,generation:inverseGeneration) }
            book.documents[book.sheets[index].path] = try XMLDocument(data:data,options:[.nodePreserveWhitespace]); book.dirty = dirty || (generation != nil && generation != savedGeneration); book.invalidate()
            refreshFormulas(picker.indexOfSelectedItem,refs:picker.indexOfSelectedItem == index ? refs : [])
            updateEditState(); syncFilter(); updateFormatControls()
        } catch { editError(error) }
    }
    @objc func editCell() { web.evaluateJavaScript("beginEdit()",completionHandler:nil) }
    func editError(_ error: Error) {
        lastEditError = error.localizedDescription
        let json = jsonString(lastEditError)
        web.evaluateJavaScript("editResult(false,\(json))",completionHandler:nil)
    }
    func jsonString(_ text: String) -> String { String(data:try! JSONSerialization.data(withJSONObject:text,options:.fragmentsAllowed),encoding:.utf8)! }
    func restoreCells(_ index: Int, _ snapshots: [CellSnapshot], dirty: Bool, generation: Int) {
        guard let book else { return }
        do {
            let refs = snapshots.map(\.ref), inverse = try book.snapshots(index,refs:refs), wasDirty = book.dirty, inverseGeneration = savedGeneration
            try book.restoreCells(index,snapshots)
            book.dirty = dirty || generation != savedGeneration
            editUndo.registerUndo(withTarget:self) { $0.restoreCells(index,inverse,dirty:wasDirty,generation:inverseGeneration) }
            refreshFormulas(picker.indexOfSelectedItem,refs:picker.indexOfSelectedItem == index ? refs : []); updateEditState(); updateFormatControls()
        } catch { editError(error) }
    }
    func commitEdits(_ edits: [CellEdit]) throws {
        guard let book else { return }
        guard !edits.isEmpty else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil); return }
        let index = picker.indexOfSelectedItem
        let edits = try edits.filter { edit in
            let raw = try book.raw(index,edit.ref)
            if edit.value.isEmpty { return !raw.isEmpty }
            let cell = try book.cell(index,edit.ref), type = cell?.attr("t") ?? ""
            switch edit.kind { case .formula: return raw != "="+normalizeFormula(edit.value)
            case .number: return !["","n"].contains(type) || raw != edit.value
            case .text: return !["inlineStr","s","str"].contains(type) || raw != edit.value }
        }
        guard !edits.isEmpty else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil); return }
        let refs = edits.map(\.ref), old = try book.snapshots(index,refs:refs), wasDirty = book.dirty, generation = savedGeneration
        try book.applyEdits(index,edits)
        let changed = try book.snapshots(index,refs:refs)
        guard zip(old,changed).contains(where: { $0.cell?.xmlString != $1.cell?.xmlString }) else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil); return }
        committedActions += 1
        editUndo.record(self,name:L("Modifica celle","Edit Cells")) { $0.restoreCells(index,old,dirty:wasDirty,generation:generation) }
        refreshFormulas(index,refs:refs); updateEditState(); lastEditError = ""
        web.evaluateJavaScript("editResult(true,'')",completionHandler:nil)
    }
    /// Refresh only the affected cells plus all formula cells, without replacing the document.
    func refreshFormulas(_ index: Int, select ref: String) { refreshFormulas(index,refs:[ref]) }
    func refreshFormulas(_ index: Int, refs: [String]) {
        guard let book, let idx = try? book.index(index) else { return }
        let editedKeys = Set(refs.map { cellKey(rowNumber($0),column($0)) }); var keys = editedKeys
        for (key,cell) in idx.cells where cell.child("f") != nil { keys.insert(key) }
        var updates: [[Any]] = []; updates.reserveCapacity(keys.count)
        for key in keys {
            let r = key/16385, c = key%16385, cell = idx.cells[key] ?? idx.blankCells[key]
            let value = cell.map { book.display($0,sheet:index) } ?? ""
            let raw = (try? book.raw(index,columnName(c)+String(r))) ?? ""
            let numeric: Any
            if case .number(let n)? = try? book.value(index,r,c) { numeric = n } else { numeric = NSNull() }
            updates.append([r,c,value,editedKeys.contains(key) ? 1 : 0,raw,numeric,cell.map { book.style($0).css } ?? ""])
        }
        guard let data = try? JSONSerialization.data(withJSONObject:updates), let json = String(data:data,encoding:.utf8) else { return }
        web.evaluateJavaScript("setCells(\(json))",completionHandler:nil)
        if !selectedRef.isEmpty { let selected = idx.cells[cellKey(rowNumber(selectedRef),column(selectedRef))] ?? idx.blankCells[cellKey(rowNumber(selectedRef),column(selectedRef))]; showDetail(selectedRef,selected.map { book.display($0,sheet:index) } ?? "") }
    }
    func selectedEdits(_ text: String) throws -> [CellEdit] {
        guard let book else { return [] }; let i = picker.indexOfSelectedItem
        return try selectedRange.refs.map { try book.automaticEdit(i,$0,text) }
    }
    @objc func clearSelection() { do { try commitEdits(selectedEdits("")) } catch { editError(error) } }
    func clipboardText() throws -> String {
        guard let book else { return "" }; let i = picker.indexOfSelectedItem, range = selectedRange
        func quote(_ s: String) -> String { s.contains(where: { "\t\n\r\"".contains($0) }) ? "\""+s.replacingOccurrences(of:"\"",with:"\"\"")+"\"" : s }
        return try (range.r1...range.r2).map { r in try (range.c1...range.c2).map { c in quote(try book.raw(i,columnName(c)+String(r))) }.joined(separator:"\t") }.joined(separator:"\n")
    }
    @objc func copySelection() {
        do { let text = try clipboardText(); let pb = NSPasteboard.general; pb.clearContents(); pb.setString(text,forType:.string); pb.setString(text,forType:.tabularText); pb.setString(text,forType:NSPasteboard.PasteboardType("public.utf8-tab-separated-values-text")) } catch { editError(error) }
    }
    @objc func cutSelection() { copySelection(); clearSelection() }
    @objc func copyCells() { if let text = window.firstResponder as? NSTextView { text.copy(nil); return }; web.evaluateJavaScript("gridClipboard('copy')",completionHandler:nil) }
    @objc func cutCells() { if let text = window.firstResponder as? NSTextView { text.cut(nil); return }; web.evaluateJavaScript("gridClipboard('cut')",completionHandler:nil) }
    @objc func pasteCells() { if let text = window.firstResponder as? NSTextView { text.paste(nil); return }; web.evaluateJavaScript("gridClipboard('paste')",completionHandler:nil) }
    @objc func pasteSelection() { if let text = NSPasteboard.general.string(forType:.string) { do { try pasteText(text) } catch { editError(error) } } else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil) } }
    func pasteText(_ text: String) throws {
        guard let book else { return }; let rows = parseTSV(text), i = picker.indexOfSelectedItem
        if rows.count == 1, rows[0].count == 1 { try commitEdits(selectedEdits(rows[0][0])); return }
        var edits: [CellEdit] = []
        for (r,row) in rows.enumerated() { for (c,value) in row.enumerated() { edits.append(try book.automaticEdit(i,columnName(selectedRange.c1+c)+String(selectedRange.r1+r),value)) } }
        try commitEdits(edits)
    }
    @objc func sortAscending() { sortSelection(true) }
    @objc func sortDescending() { sortSelection(false) }
    func sortSelection(_ ascending: Bool, column col: Int? = nil) {
        guard let book else { return }; let i = picker.indexOfSelectedItem
        do {
            let old = try book.sheetDocument(i).xmlData, wasDirty = book.dirty, idx = try book.index(i), generation = savedGeneration
            try book.sort(i,column:col ?? selectedRange.c1,row:selectedRange.r1,ascending:ascending)
            let newIndex = try book.index(i), refs = Set(idx.cells.keys).union(newIndex.cells.keys).map { columnName($0%16385)+String($0/16385) }
            guard try book.sheetDocument(i).xmlData != old else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil); return }; committedActions += 1
            editUndo.record(self,name:L("Ordina","Sort")) { $0.restoreSheet(i,old,dirty:wasDirty,refs:refs,generation:generation) }
            refreshFormulas(i,refs:refs); updateEditState(); lastEditError = ""
            web.evaluateJavaScript("editResult(true,'')",completionHandler:nil)
        } catch { editError(error) }
    }
    @objc func newWorkbook() {
        guard confirmUnsaved() else { return }
        do {
            loadID = UUID(); setLoading(false); let newBook = try Book.blank(); book = newBook; editUndo.removeAllActions(); selectedRef = ""; selectedRange = CellRange(); paintStyle = nil
            picker.removeAllItems(); picker.addItem(withTitle:newBook.sheets[0].name)
            window.title = L("Senza titolo","Untitled"); window.representedURL = nil
            show(try newBook.render(0)); updateEditState()
        } catch { show(error) }
    }
    func persist(to target: URL) throws {
        guard let book else { return }
        try book.save(to:target)
        book.url = target; book.untitled = false; book.dirty = false; savedGeneration += 1
        for i in book.sheets.indices { if let doc = book.documents[book.sheets[i].path] { for c in doc.nodes("/worksheet/sheetData/row/c[f]") { if let cell = c as? XMLElement, (try? book.computed(cell,sheet:i)) != nil { c.nodes("v").forEach { $0.detach() } } } } }; book.invalidate()
        window.title = target.lastPathComponent; window.representedURL = target
        UserDefaults.standard.set(target.path,forKey:"lastFile"); NSDocumentController.shared.noteNewRecentDocumentURL(target); refreshRecentMenu(); updateEditState()
    }
    @discardableResult @objc func saveFile() -> Bool {
        guard let book else { return false }
        if book.untitled { return saveCopy() }
        do { try persist(to:book.url); return true } catch { show(error); return false }
    }
    @objc func fillDown() {
        guard let book, selectedRange.r2>selectedRange.r1 else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil); return }
        do {
            var edits: [CellEdit] = []; let i = picker.indexOfSelectedItem
            for c in selectedRange.c1...selectedRange.c2 { let value = try book.raw(i,columnName(c)+String(selectedRange.r1)); for r in (selectedRange.r1+1)...selectedRange.r2 { edits.append(try book.automaticEdit(i,columnName(c)+String(r),value.hasPrefix("=") ? "="+fillFormula(String(value.dropFirst()),down:r-selectedRange.r1) : value)) } }
            try commitEdits(edits)
        } catch { editError(error) }
    }
    @objc func toggleFreeze() { web.evaluateJavaScript("toggleFreeze()",completionHandler:nil) }
    @objc func autoWidth() { web.evaluateJavaScript("autoWidth(current?+current.dataset.c:1)",completionHandler:nil) }
    var formatControls: [String:NSControl] = [:]
    var paintStyle: Int?
    var selectedStyle: CellStyle {
        guard let book, let cell = try? book.cell(picker.indexOfSelectedItem,selectedRef) else { return CellStyle() }
        return book.style(cell)
    }
    func updateFormatControls() {
        let s = selectedStyle
        (formatControls["bold"] as? NSButton)?.state = s.bold ? .on : .off
        (formatControls["italic"] as? NSButton)?.state = s.italic ? .on : .off
        (formatControls["strike"] as? NSButton)?.state = s.strike ? .on : .off
        (formatControls["wrap"] as? NSButton)?.state = s.wrap ? .on : .off
        (formatControls["fontSize"] as? NSTextField)?.stringValue = FValue.number(s.fontSize).text
        (formatControls["foreground"] as? NSColorWell)?.color = colorFromHex(s.foreground)
        (formatControls["fill"] as? NSColorWell)?.color = colorFromHex(s.background)
        for key in ["foreground","fill"] { formatControls[key]?.needsDisplay = true }
        if let popup = formatControls["alignment"] as? NSPopUpButton {
            var effective = s.alignment
            if !s.explicitAlignment, let book, case .number? = try? book.value(picker.indexOfSelectedItem,rowNumber(selectedRef),column(selectedRef)), s.kind != .text { effective = "right" }
            let symbol: String, description: String
            switch effective {
            case "center": symbol = "text.aligncenter"; description = L("Allinea al centro","Align Center")
            case "right": symbol = "text.alignright"; description = L("Allinea a destra","Align Right")
            default: symbol = "text.alignleft"; description = L("Allinea a sinistra","Align Left")
            }
            let image = NSImage(systemSymbolName:symbol,accessibilityDescription:description)
            popup.item(at:0)?.title = ""; popup.item(at:0)?.image = image; popup.image = image; popup.imagePosition = .imageOnly
            for item in popup.menu?.items.dropFirst() ?? [] { item.state = (item.representedObject as? String) == "align:"+effective ? .on : .off }
        }
        let filter = (try? book?.sheetDocument(picker.indexOfSelectedItem).rootElement()?.child("autoFilter")) != nil
        (formatControls["filter"] as? NSButton)?.state = filter ? .on : .off
    }
    func colorFromHex(_ text: String) -> NSColor {
        let n = UInt32(text.replacingOccurrences(of:"#",with:""),radix:16) ?? 0
        return NSColor(srgbRed:Double((n>>16)&255)/255,green:Double((n>>8)&255)/255,blue:Double(n&255)/255,alpha:1)
    }
    func hexColor(_ color: NSColor) -> String {
        let c = color.usingColorSpace(.sRGB) ?? .black
        return String(format:"#%02X%02X%02X",Int((c.redComponent*255).rounded()),Int((c.greenComponent*255).rounded()),Int((c.blueComponent*255).rounded()))
    }
    func formatSelection(_ change: StyleChange) {
        guard let book, !loading, !selectedRef.isEmpty else { return }
        do {
            let index = picker.indexOfSelectedItem, refs = selectedRange.refs, old = try book.snapshots(index,refs:refs), dirty = book.dirty, generation = savedGeneration
            try book.formatCells(index,refs:refs,change:change)
            let new = try book.snapshots(index,refs:refs)
            guard zip(old,new).contains(where:{ $0.cell?.xmlString != $1.cell?.xmlString }) else { book.dirty = dirty; updateFormatControls(); return }
            committedActions += 1
            editUndo.record(self,name:L("Formato celle","Format Cells")) { $0.restoreCells(index,old,dirty:dirty,generation:generation) }
            refreshFormulas(index,refs:refs); updateEditState(); updateFormatControls()
            web.evaluateJavaScript("busy=false;document.getElementById('message').textContent='';selectionFurniture()",completionHandler:nil)
        } catch { editError(error) }
    }
    @objc func formatAction(_ sender: NSControl) {
        switch sender.identifier?.rawValue ?? "" {
        case "bold": formatSelection(.flag("b",!selectedStyle.bold))
        case "italic": formatSelection(.flag("i",!selectedStyle.italic))
        case "strike": formatSelection(.flag("strike",!selectedStyle.strike))
        case "currency": formatSelection(.number("#,##0.00 \"€\""))
        case "percent": formatSelection(.number("0.00%"))
        case "decLess": formatSelection(.decimals(-1))
        case "decMore": formatSelection(.decimals(1))
        case "fontLess": formatSelection(.fontDelta(-1))
        case "fontMore": formatSelection(.fontDelta(1))
        case "fontSize": if let field = sender as? NSTextField, let size = Double(field.stringValue), size.isFinite { formatSelection(.fontSize(size)); window.makeFirstResponder(web) } else { updateFormatControls() }
        case "foreground": if let well = sender as? NSColorWell { formatSelection(.foreground(hexColor(well.color))) }
        case "fill": if let well = sender as? NSColorWell { formatSelection(.fill(hexColor(well.color))) }
        case "wrap": formatSelection(.wrap(!selectedStyle.wrap))
        case "paint":
            guard let book, !selectedRef.isEmpty else { return }
            let cell = try? book.cell(picker.indexOfSelectedItem,selectedRef)
            paintStyle = Int(cell?.attr("s") ?? "") ?? 0; web.evaluateJavaScript("armPaint(true)",completionHandler:nil)
        case "filter": toggleAutoFilter()
        default: break
        }
    }
    @objc func formatMenu(_ sender: NSMenuItem) {
        let command = sender.representedObject as? String ?? ""
        if command.hasPrefix("number:") { formatSelection(.number(String(command.dropFirst(7)))) }
        else if command.hasPrefix("align:") { formatSelection(.alignment(String(command.dropFirst(6)))) }
        else if command.hasPrefix("aggregate:") { insertAggregate(String(command.dropFirst(10))) }
        else if command == "bold" { formatSelection(.flag("b",!selectedStyle.bold)) }
        else if command == "italic" { formatSelection(.flag("i",!selectedStyle.italic)) }
        else if command == "strike" { formatSelection(.flag("strike",!selectedStyle.strike)) }
        else if command == "wrap" { formatSelection(.wrap(!selectedStyle.wrap)) }
    }
    func insertAggregate(_ function: String) {
        guard let book, !selectedRef.isEmpty else { return }
        do { let i = picker.indexOfSelectedItem; try commitEdits([book.automaticEdit(i,selectedRef,book.aggregateFormula(i,ref:selectedRef,function:function))]) } catch { editError(error) }
    }
    func syncFilter() {
        guard let book else { return }
        let parts = ((try? book.sheetDocument(picker.indexOfSelectedItem).rootElement()?.child("autoFilter")?.attr("ref")) ?? "").split(separator:":").map(String.init)
        let first = parts.first ?? "", last = parts.last ?? ""
        web.evaluateJavaScript("window.setFilterRegion&&setFilterRegion(\(rowNumber(first)),\(column(first)),\(rowNumber(last)),\(column(last)))",completionHandler:nil)
        updateFormatControls()
    }
    @objc func toggleAutoFilter() {
        guard let book, !loading else { return }
        do {
            let i = picker.indexOfSelectedItem, old = try book.sheetDocument(i).xmlData, dirty = book.dirty, generation = savedGeneration
            try book.toggleFilter(i,selection:selectedRange); committedActions += 1
            editUndo.record(self,name:L("Filtro","Filter")) { $0.restoreSheet(i,old,dirty:dirty,generation:generation) }
            syncFilter(); updateEditState()
        } catch { editError(error) }
    }
    @objc func openHelp() { NSWorkspace.shared.open(URL(string:"https://github.com/lucasflati12/docx-xlsx-lite")!) }
    @objc func zoomChoice(_ sender: NSPopUpButton) { applyZoom(Double(sender.titleOfSelectedItem?.replacingOccurrences(of:"%",with:"") ?? "100")! / 100) }
    func makeFormatToolbar() -> NSScrollView {
        var views: [NSView] = []
        func button(_ id: String, _ symbol: String?, _ title: String, _ tip: String, _ action: Selector? = nil, toggle: Bool = false) -> NSButton {
            let b = NSButton(title:title,target:self,action:action ?? #selector(formatAction(_:))); b.identifier = NSUserInterfaceItemIdentifier(id); b.toolTip = tip; b.isBordered = false; b.bezelStyle = .regularSquare; b.controlSize = .small; b.font = .systemFont(ofSize:12)
            if let symbol { b.image = NSImage(systemSymbolName:symbol,accessibilityDescription:tip); b.imagePosition = .imageOnly }
            if toggle { b.setButtonType(.pushOnPushOff) }
            b.translatesAutoresizingMaskIntoConstraints = false; b.widthAnchor.constraint(equalToConstant:title.count>2 && symbol == nil ? 38 : 27).isActive = true; b.heightAnchor.constraint(equalToConstant:27).isActive = true
            formatControls[id] = b; views.append(b); return b
        }
        func separator() { let sep = NSBox(); sep.boxType = .separator; sep.translatesAutoresizingMaskIntoConstraints = false; sep.widthAnchor.constraint(equalToConstant:1).isActive = true; sep.heightAnchor.constraint(equalToConstant:20).isActive = true; views.append(sep) }
        func popup(_ id: String, _ title: String, _ items: [(String,String)], width: CGFloat) -> NSPopUpButton {
            let p = NSPopUpButton(frame:.zero,pullsDown:true); p.addItem(withTitle:title); p.isBordered = false; p.font = .systemFont(ofSize:12); p.identifier = NSUserInterfaceItemIdentifier(id); p.toolTip = title
            for (label,value) in items { let item = NSMenuItem(title:label,action:#selector(formatMenu(_:)),keyEquivalent:""); item.target = self; item.representedObject = value; p.menu?.addItem(item) }
            p.translatesAutoresizingMaskIntoConstraints = false; p.widthAnchor.constraint(equalToConstant:width).isActive = true; formatControls[id] = p; views.append(p); return p
        }
        _ = button("undo","arrow.uturn.backward","",L("Annulla (⌘Z)","Undo (⌘Z)"),#selector(undoMenu))
        _ = button("redo","arrow.uturn.forward","",L("Ripristina (⇧⌘Z)","Redo (⇧⌘Z)"),#selector(redoMenu))
        _ = button("print","printer","",L("Stampa (⌘P)","Print (⌘P)"),#selector(printSheet))
        _ = button("paint","paintbrush","",L("Copia formato: applica alla prossima selezione","Paint format: apply to the next selection"))
        zoomLabel = NSPopUpButton(); zoomLabel.addItems(withTitles:["50%","75%","85%","100%","125%","150%","200%"]); zoomLabel.selectItem(withTitle:"100%"); zoomLabel.target = self; zoomLabel.action = #selector(zoomChoice(_:)); zoomLabel.isBordered = false; zoomLabel.toolTip = L("Zoom","Zoom"); zoomLabel.translatesAutoresizingMaskIntoConstraints = false; zoomLabel.widthAnchor.constraint(equalToConstant:66).isActive = true; formatControls["zoom"] = zoomLabel; views.append(zoomLabel)
        separator()
        _ = button("currency",nil,"€",L("Formato valuta","Currency format")); _ = button("percent",nil,"%",L("Formato percentuale","Percent format"))
        _ = button("decLess",nil,".0←",L("Diminuisci cifre decimali","Decrease decimal places")); _ = button("decMore",nil,".00→",L("Aumenta cifre decimali","Increase decimal places"))
        _ = popup("numberFormat","123",[(L("Automatico","Automatic"),"General"),(L("Numero","Number"),"0.00"),(L("Percentuale","Percent"),"0.00%"),(L("Valuta","Currency"),"#,##0.00 \"€\""),(L("Data","Date"),"dd/mm/yyyy"),(L("Ora","Time"),"hh:mm:ss"),(L("Testo","Text"),"@")].map { ($0.0,"number:"+$0.1) },width:48)
        separator()
        _ = button("fontLess",nil,"−",L("Riduci dimensione carattere","Decrease font size"))
        let size = NSTextField(string:"11"); size.identifier = NSUserInterfaceItemIdentifier("fontSize"); size.target = self; size.action = #selector(formatAction(_:)); size.alignment = .center; size.font = .systemFont(ofSize:12); size.toolTip = L("Dimensione carattere","Font size"); size.translatesAutoresizingMaskIntoConstraints = false; size.widthAnchor.constraint(equalToConstant:34).isActive = true; formatControls["fontSize"] = size; views.append(size)
        _ = button("fontMore",nil,"+",L("Aumenta dimensione carattere","Increase font size"))
        separator()
        _ = button("bold","bold","B",L("Grassetto","Bold"),toggle:true); _ = button("italic","italic","I",L("Corsivo","Italic"),toggle:true); _ = button("strike","strikethrough","S",L("Barrato","Strikethrough"),toggle:true)
        for (id,label,color) in [("foreground",L("Colore testo","Text color"),NSColor.labelColor),("fill",L("Colore riempimento","Fill color"),NSColor.systemYellow)] { let well = ToolbarColorWell(frame:NSRect(x:0,y:0,width:27,height:27)); well.identifier = NSUserInterfaceItemIdentifier(id); well.color = color; well.target = self; well.action = #selector(formatAction(_:)); well.isContinuous = false; well.toolTip = label; well.setAccessibilityLabel(label); well.isBordered = false; well.translatesAutoresizingMaskIntoConstraints = false; well.widthAnchor.constraint(equalToConstant:27).isActive = true; formatControls[id] = well; views.append(well) }
        separator()
        let alignment = popup("alignment",L("Allinea","Align"),[(L("Sinistra","Left"),"align:left"),(L("Centro","Center"),"align:center"),(L("Destra","Right"),"align:right")],width:42); alignment.item(at:0)?.title = ""; alignment.imagePosition = .imageOnly; alignment.toolTip = L("Allineamento orizzontale","Horizontal alignment"); alignment.image = NSImage(systemSymbolName:"text.alignleft",accessibilityDescription:L("Allinea a sinistra","Align Left")); alignment.item(at:0)?.image = alignment.image
        _ = button("wrap","arrow.turn.down.left","",L("Testo a capo","Wrap text"),toggle:true)
        separator()
        _ = button("filter","line.3.horizontal.decrease.circle","",L("Crea/rimuovi filtro","Create/remove filter"),toggle:true)
        _ = popup("aggregate","Σ",[(L("Somma","Sum"),"SUM"),(L("Media","Average"),"AVERAGE"),(L("Conta numeri","Count numbers"),"COUNT"),(L("Massimo","Maximum"),"MAX"),(L("Minimo","Minimum"),"MIN")].map { ($0.0,"aggregate:"+$0.1) },width:42)
        let stack = NSStackView(views:views); stack.orientation = .horizontal; stack.alignment = .centerY; stack.spacing = 5; stack.edgeInsets = NSEdgeInsets(top:4,left:8,bottom:4,right:8)
        let scroll = NSScrollView(); scroll.hasHorizontalScroller = true; scroll.autohidesScrollers = true; scroll.scrollerStyle = .overlay; scroll.drawsBackground = false; scroll.wantsLayer = true; scroll.layer?.backgroundColor = NSColor.controlBackgroundColor.cgColor; scroll.layer?.cornerRadius = 14
        stack.frame = NSRect(x:0,y:0,width:max(950,stack.fittingSize.width),height:36); scroll.documentView = stack
        return scroll
    }
    func showDetail(_ ref: String, _ text: String) {
        var line = "\(ref)   \(text)"
        if let book, let cell = (try? book.cell(picker.indexOfSelectedItem, ref)).flatMap({ $0 }), let formula = book.formulaText(cell) { line += "      =" + formula }
        detail.string = line
        updateFormatControls()
    }
    @discardableResult @objc func saveCopy() -> Bool {
        guard let book else { return false }
        let ext = book.url.pathExtension.lowercased()
        let panel = NSSavePanel(); panel.allowedContentTypes = [UTType(filenameExtension: ext) ?? .data]
        panel.nameFieldStringValue = book.untitled ? L("Senza titolo.xlsx","Untitled.xlsx") : book.url.deletingPathExtension().lastPathComponent + L(" - copia.", " - copy.") + ext
        panel.directoryURL = book.untitled ? FileManager.default.urls(for:.documentDirectory,in:.userDomainMask).first : book.url.deletingLastPathComponent(); panel.title = L("Salva una copia", "Save a Copy")
        guard panel.runModal() == .OK, let target = panel.url else { return false }
        do { try persist(to:target); return true } catch { show(error); return false }
    }

    func confirmUnsaved() -> Bool {
        guard book?.dirty == true else { return true }
        let alert = NSAlert(); alert.messageText = L("Salvare le modifiche?", "Save your changes?"); alert.informativeText = L("Ci sono modifiche non salvate in questo file.", "This file has unsaved changes.")
        alert.addButton(withTitle: L("Salva", "Save")); alert.addButton(withTitle: L("Annulla", "Cancel")); alert.addButton(withTitle: L("Non salvare", "Don't Save"))
        switch alert.runModal() { case .alertFirstButtonReturn: return saveFile(); case .alertThirdButtonReturn: return true; default: return false }
    }
    func windowShouldClose(_ sender: NSWindow) -> Bool { let close = confirmUnsaved(); if close { book?.dirty = false }; return close }
    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply { confirmUnsaved() ? .terminateNow : .terminateCancel }

    func applyZoom(_ value: Double) {
        zoom = min(2, max(0.5, value)); web.pageZoom = zoom; web.magnification = 1
        let title = "\(Int((zoom * 100).rounded()))%"; if zoomLabel.item(withTitle:title) == nil { zoomLabel.addItem(withTitle:title) }; zoomLabel.selectItem(withTitle:title)
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
        guard let data = message.body as? [String:String] else { return }
        switch data["action"] ?? "select" {
        case "select":
            selectedRef = data["ref"] ?? ""
            selectedRange = CellRange(r1:Int(data["r1"] ?? "") ?? rowNumber(selectedRef),c1:Int(data["c1"] ?? "") ?? column(selectedRef),r2:Int(data["r2"] ?? "") ?? rowNumber(selectedRef),c2:Int(data["c2"] ?? "") ?? column(selectedRef))
            showDetail(selectedRef,data["text"] ?? "")
        case "commit":
            guard let book, let ref = data["ref"] else { return }
            do { try commitEdits([book.automaticEdit(picker.indexOfSelectedItem,ref,data["value"] ?? "")]) } catch { editError(error) }
        case "clear": clearSelection()
        case "copy": copySelection()
        case "cut": cutSelection()
        case "paste": pasteSelection()
        case "undo": undoEdit()
        case "redo": redoEdit()
        case "sort": sortSelection(data["ascending"] == "1",column:Int(data["col"] ?? ""))
        case "fillDown": fillDown()
        case "paint": if let style = paintStyle { paintStyle = nil; formatSelection(.copy(style)); web.evaluateJavaScript("armPaint(false)",completionHandler:nil) }
        case "cancelPaint": paintStyle = nil
        case "autofill":
            guard let book else { return }
            func range(_ key: String) -> CellRange? { guard let text = data[key], let bytes = text.data(using:.utf8), let obj = (try? JSONSerialization.jsonObject(with:bytes)) as? [String:Int],let r1=obj["r1"],let c1=obj["c1"],let r2=obj["r2"],let c2=obj["c2"] else { return nil }; return CellRange(r1:r1,c1:c1,r2:r2,c2:c2) }
            if let source = range("source"), let target = range("target") { do { try commitEdits(book.autofillEdits(picker.indexOfSelectedItem,source:source,target:target)) } catch { editError(error) } }
            else { web.evaluateJavaScript("editResult(true,'')",completionHandler:nil) }
        case "copyText": NSPasteboard.general.clearContents(); NSPasteboard.general.setString(data["value"] ?? "",forType:.string)
        case "pasteInput": if let text = NSPasteboard.general.string(forType:.string) { web.evaluateJavaScript("pasteInput(\(jsonString(text)))",completionHandler:nil) }
        default: break
        }
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
        add(file, L("Nuovo", "New"), #selector(newWorkbook), "n", target:self)
        add(file, L("Apri…", "Open…"), #selector(openFile), "o", target: self)
        let recent = NSMenuItem(title: recentMenu.title, action: nil, keyEquivalent: ""); recent.submenu = recentMenu; file.addItem(recent); refreshRecentMenu()
        file.addItem(.separator())
        add(file, L("Chiudi", "Close"), #selector(NSWindow.performClose(_:)), "w")
        add(file, L("Salva", "Save"), #selector(saveFile), "s", target:self)
        add(file, L("Salva una copia…", "Save a Copy…"), #selector(saveCopy), "s", [.command,.shift], target:self)
        file.addItem(.separator())
        add(file, L("Stampa…", "Print…"), #selector(printSheet), "p", target: self)
        let edit = submenu(L("Modifica", "Edit"))
        add(edit, L("Annulla", "Undo"), #selector(undoMenu), "z", target: self)
        add(edit, L("Ripristina", "Redo"), #selector(redoMenu), "z", [.command, .shift], target: self)
        edit.addItem(.separator())
        add(edit, L("Taglia", "Cut"), #selector(cutCells), "x", target:self)
        add(edit, L("Copia", "Copy"), #selector(copyCells), "c", target:self)
        add(edit, L("Incolla", "Paste"), #selector(pasteCells), "v", target:self)
        add(edit, L("Seleziona tutto", "Select All"), #selector(selectAllCells), "a", target:self)
        edit.addItem(.separator())
        add(edit, L("Modifica cella…", "Edit Cell…"), #selector(editCell), target: self)
        add(edit, L("Rimuovi filtri", "Clear Filters"), #selector(clearFilters), target: self)
        edit.addItem(.separator())
        add(edit, L("Cerca…", "Find…"), #selector(focusSearch), "f", target: self)
        add(edit, L("Trova successivo", "Find Next"), #selector(findNext), "g", target: self)
        add(edit, L("Trova precedente", "Find Previous"), #selector(findPrevious), "g", [.command, .shift], target: self)
        let view = submenu(L("Visualizza", "View"))
        add(view,L("Blocca/sblocca prima riga","Freeze/Unfreeze First Row"),#selector(toggleFreeze),target:self)
        add(view,L("Adatta larghezza colonna","Fit Column Width"),#selector(autoWidth),target:self)
        add(view, L("Ingrandisci", "Zoom In"), #selector(zoomIn), "+", target: self)
        add(view, L("Riduci", "Zoom Out"), #selector(zoomOut), "-", target: self)
        add(view, L("Dimensioni reali", "Actual Size"), #selector(resetZoom), "0", target: self)
        view.addItem(.separator())
        add(view, L("Foglio successivo", "Next Sheet"), #selector(nextSheet), "]", [.command, .option], target: self)
        add(view, L("Foglio precedente", "Previous Sheet"), #selector(previousSheet), "[", [.command, .option], target: self)
        view.addItem(.separator())
        add(view, L("Schermo intero", "Enter Full Screen"), #selector(NSWindow.toggleFullScreen(_:)), "f", [.command, .control])
        let insert = submenu(L("Inserisci","Insert"))
        for (title,name) in [(L("Somma","Sum"),"SUM"),(L("Media","Average"),"AVERAGE"),(L("Conta numeri","Count numbers"),"COUNT"),(L("Massimo","Maximum"),"MAX"),(L("Minimo","Minimum"),"MIN")] { let item = insert.addItem(withTitle:title,action:#selector(formatMenu(_:)),keyEquivalent:""); item.target = self; item.representedObject = "aggregate:"+name }
        let format = submenu(L("Formato","Format"))
        for (title,key,command) in [(L("Grassetto","Bold"),"b","bold"),(L("Corsivo","Italic"),"i","italic"),(L("Barrato","Strikethrough"),"","strike"),(L("Testo a capo","Wrap text"),"","wrap")] { let item = format.addItem(withTitle:title,action:#selector(formatMenu(_:)),keyEquivalent:key); item.target = self; item.representedObject = command }
        format.addItem(.separator())
        for (title,command) in [(L("Automatico","Automatic"),"number:General"),(L("Numero","Number"),"number:0.00"),(L("Percentuale","Percent"),"number:0.00%"),(L("Valuta","Currency"),"number:#,##0.00 \"€\""),(L("Data","Date"),"number:dd/mm/yyyy"),(L("Ora","Time"),"number:hh:mm:ss"),(L("Testo","Text"),"number:@"),(L("Allinea a sinistra","Align Left"),"align:left"),(L("Allinea al centro","Align Center"),"align:center"),(L("Allinea a destra","Align Right"),"align:right")] { let item = format.addItem(withTitle:title,action:#selector(formatMenu(_:)),keyEquivalent:""); item.target = self; item.representedObject = command }
        let data = submenu(L("Dati","Data"))
        add(data,L("Ordina A → Z","Sort A → Z"),#selector(sortAscending),target:self)
        add(data,L("Ordina Z → A","Sort Z → A"),#selector(sortDescending),target:self)
        add(data,L("Riempi in basso","Fill Down"),#selector(fillDown),"d",target:self)
        add(data,L("Crea/rimuovi filtro","Create/Remove Filter"),#selector(toggleAutoFilter),target:self)
        let tools = submenu(L("Strumenti","Tools"))
        add(tools,L("Cerca…","Find…"),#selector(focusSearch),target:self)
        add(tools,L("Riempi in basso","Fill Down"),#selector(fillDown),target:self)
        let windowMenu = submenu(L("Finestra", "Window"))
        add(windowMenu, L("Contrai", "Minimize"), #selector(NSWindow.performMiniaturize(_:)), "m")
        add(windowMenu, L("Ridimensiona", "Zoom"), #selector(NSWindow.performZoom(_:)))
        let help = submenu(L("Guida","Help")); add(help,L("Guida di XLSX Lite","XLSX Lite Help"),#selector(openHelp),target:self); NSApp.helpMenu = help
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
        let saveButton = NSButton(title:L("Salva","Save"),target:self,action:#selector(saveFile)); saveButton.toolTip = L("Salva (⌘S)","Save (⌘S)")
        let spacer = NSView(); spacer.setContentHuggingPriority(.defaultLow,for:.horizontal)
        let bar = NSStackView(views:[open,progress,picker,saveButton,spacer,search]); bar.spacing = 8; bar.orientation = .horizontal
        let formatting = makeFormatToolbar()
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
        let savedZoom = UserDefaults.standard.double(forKey: "pageZoom"); applyZoom(savedZoom > 0 ? savedZoom : 1)
        for view in [bar, formatting, web!, footer, statusBar] { view.translatesAutoresizingMaskIntoConstraints = false; content.addSubview(view) }
        NSLayoutConstraint.activate([
            bar.topAnchor.constraint(equalTo: content.topAnchor, constant: 12), bar.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 16), bar.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -16), bar.heightAnchor.constraint(equalToConstant: 30),
            formatting.topAnchor.constraint(equalTo:bar.bottomAnchor,constant:6), formatting.leadingAnchor.constraint(equalTo:content.leadingAnchor,constant:10), formatting.trailingAnchor.constraint(equalTo:content.trailingAnchor,constant:-10), formatting.heightAnchor.constraint(equalToConstant:38),
            web.topAnchor.constraint(equalTo: formatting.bottomAnchor, constant: 4), web.leadingAnchor.constraint(equalTo: content.leadingAnchor), web.trailingAnchor.constraint(equalTo: content.trailingAnchor), web.bottomAnchor.constraint(equalTo: footer.topAnchor),
            footer.leadingAnchor.constraint(equalTo: content.leadingAnchor), footer.trailingAnchor.constraint(equalTo: content.trailingAnchor), footer.bottomAnchor.constraint(equalTo: statusBar.topAnchor, constant: -2), footer.heightAnchor.constraint(equalToConstant: 46),
            statusBar.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 12), statusBar.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -12), statusBar.bottomAnchor.constraint(equalTo: content.bottomAnchor, constant: -6), statusBar.heightAnchor.constraint(equalToConstant: 16),
            picker.widthAnchor.constraint(equalToConstant: 170), search.widthAnchor.constraint(equalToConstant: 170)])
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
                    self.book = newBook; self.paintStyle = nil; self.editUndo.removeAllActions(); self.selectedRef = ""; self.selectedRange = CellRange()
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
    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) { webView.pageZoom = zoom; if paintStyle != nil { web.evaluateJavaScript("armPaint(true)",completionHandler:nil) } }
    func show(_ error: Error) { NSAlert(error: error).runModal() }
    func webView(_ webView: WKWebView, decidePolicyFor action: WKNavigationAction, decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        if let url = action.request.url, url.isFileURL, supportedExtensions.contains(url.pathExtension.lowercased()) { decisionHandler(.cancel); load(url); return }
        decisionHandler(action.navigationType == .other ? .allow : .cancel)
    }
}

#if DEBUG
if CommandLine.arguments.count > 2 && CommandLine.arguments[1] == "--model-test" {
    runSheetModelTests(URL(fileURLWithPath:CommandLine.arguments[2])); exit(0)
}
#endif
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
