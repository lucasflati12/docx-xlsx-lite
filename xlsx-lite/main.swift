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
// Reuse the surrounding part's namespace prefix when adding OOXML nodes.
func partElement(_ name: String, like context: XMLElement, text: String? = nil) -> XMLElement {
    let prefix = (context.name ?? "").split(separator:":")
    return XMLElement(name:prefix.count>1 ? String(prefix[0])+":"+name : name,stringValue:text)
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
struct SheetIndex { var cells: [Int: XMLElement] = [:]; var blankCells: [Int:XMLElement] = [:]; var sharedMasters: [String:XMLElement] = [:]; var hiddenRows = Set<Int>(); var maxRow = 0, maxCol = 0 }

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

// MARK: - Structural references and package transactions
func gridActionTitle(rows: Bool, count: Int, deleting: Bool, after: Bool = false) -> String {
    let n = count == 0 ? "%N" : String(count)
    if deleting {
        if count == 1 { return rows ? L("Elimina riga","Delete row") : L("Elimina colonna","Delete column") }
        return rows ? L("Elimina \(n) righe","Delete \(n) rows") : L("Elimina \(n) colonne","Delete \(n) columns")
    }
    if rows { return after ? (count == 1 ? L("Inserisci 1 riga sotto","Insert 1 row below") : L("Inserisci \(n) righe sotto","Insert \(n) rows below")) : (count == 1 ? L("Inserisci 1 riga sopra","Insert 1 row above") : L("Inserisci \(n) righe sopra","Insert \(n) rows above")) }
    return after ? (count == 1 ? L("Inserisci 1 colonna a destra","Insert 1 column to the right") : L("Inserisci \(n) colonne a destra","Insert \(n) columns to the right")) : (count == 1 ? L("Inserisci 1 colonna a sinistra","Insert 1 column to the left") : L("Inserisci \(n) colonne a sinistra","Insert \(n) columns to the left"))
}
struct GridChange {
    let rows: Bool, at: Int, count: Int, deleting: Bool
    var limit: Int { rows ? 1048576 : 16384 }
    func point(_ n: Int, clamp: Bool = false) -> Int? {
        if !deleting { return n>=at ? n+count : n }
        if n<at { return n }; if n>=at+count { return n-count }; return clamp ? at : nil
    }
    func interval(_ lo: Int, _ hi: Int) -> (Int,Int)? {
        if !deleting { return (lo>=at ? lo+count : lo,hi>=at ? hi+count : hi) }
        let a = lo>=at && lo<at+count ? at+count : lo
        let b = hi>=at && hi<at+count ? at-1 : hi
        guard a<=b else { return nil }; return (a>=at+count ? a-count : a,b>=at+count ? b-count : b)
    }
}
struct GridPoint {
    var row: Int?, col: Int?, absoluteRow = false, absoluteCol = false
    init?(_ text: String) {
        let s = text.uppercased(), parts = s.split(separator:"$",omittingEmptySubsequences:true)
        let clean = parts.joined()
        let letters = String(clean.prefix(while: { $0.isLetter })), digits = String(clean.dropFirst(letters.count))
        guard !clean.isEmpty, letters.isEmpty || letters.range(of:"^[A-Z]{1,3}$",options:.regularExpression) != nil,
              digits.isEmpty || (Int(digits) ?? 0)>0 else { return nil }
        col = letters.isEmpty ? nil : column(letters); row = digits.isEmpty ? nil : Int(digits)
        absoluteCol = !letters.isEmpty && s.hasPrefix("$")
        absoluteRow = !digits.isEmpty && (letters.isEmpty ? s.hasPrefix("$") : s.contains("$"+digits))
        guard (col ?? 1)<=16384,(row ?? 1)<=1048576 else { return nil }
    }
    var text: String { (col.map { (absoluteCol ? "$" : "")+columnName($0) } ?? "")+(row.map { (absoluteRow ? "$" : "")+String($0) } ?? "") }
}
struct GridRef {
    var first: GridPoint, last: GridPoint, range: Bool
    init?(_ text: String) {
        let parts = text.split(separator:":",omittingEmptySubsequences:false)
        guard parts.count<=2,let a = GridPoint(String(parts[0])),let b = GridPoint(String(parts.last!)),
              (a.row == nil) == (b.row == nil),(a.col == nil) == (b.col == nil),parts.count == 2 || (a.row != nil && a.col != nil) else { return nil }
        first = a; last = b; range = parts.count == 2
    }
    var text: String { first.text+(range ? ":"+last.text : "") }
    var top: Int { min(first.row ?? 1,last.row ?? 1) }; var left: Int { min(first.col ?? 1,last.col ?? 1) }
    func shifted(_ change: GridChange, bounded: Bool = false) throws -> GridRef? {
        let a = change.rows ? first.row : first.col, b = change.rows ? last.row : last.col
        guard let x = a,let y = b else { return self }
        guard let (lo,hi) = change.interval(min(x,y),max(x,y)) else { return nil }
        if !bounded && !(min(x,y) == 1 && max(x,y) == change.limit) && hi>change.limit { throw failure(L("Lo spostamento supererebbe i limiti del foglio Excel.","The shift would exceed the Excel worksheet limits.")) }
        guard lo<=change.limit else { return nil }
        var result = self
        let p = x<=y ? lo : min(change.limit,hi), q = x<=y ? min(change.limit,hi) : lo
        if change.rows { result.first.row = p; result.last.row = q } else { result.first.col = p; result.last.col = q }
        return result
    }
    func copied(down: Int, right: Int) -> GridRef? {
        var r = self
        if let n = r.first.row,!r.first.absoluteRow { r.first.row = n+down }; if let n = r.last.row,!r.last.absoluteRow { r.last.row = n+down }
        if let n = r.first.col,!r.first.absoluteCol { r.first.col = n+right }; if let n = r.last.col,!r.last.absoluteCol { r.last.col = n+right }
        for p in [r.first,r.last] { if (p.row ?? 1)<1 || (p.row ?? 1)>1048576 || (p.col ?? 1)<1 || (p.col ?? 1)>16384 { return nil } }
        return r
    }
}
enum RefScanner {
    // Double-quoted Excel strings are handled separately; single quotes belong to sheet names.
    static let regex = try! NSRegularExpression(pattern:"(?<![A-Za-z0-9_.$])(?:(('(?:[^']|'')+'|[\\p{L}_][\\p{L}0-9_.]*)!))?((?:[$]?[A-Za-z]{1,3}[$]?[1-9][0-9]{0,6})(?::[$]?[A-Za-z]{1,3}[$]?[1-9][0-9]{0,6})?|[$]?[A-Za-z]{1,3}:[$]?[A-Za-z]{1,3}|[$]?[1-9][0-9]{0,6}:[$]?[1-9][0-9]{0,6})(?![A-Za-z0-9_.]|\\s*\\()")
    static func transform(_ input: String, _ replace: (String?,GridRef) throws -> String?) rethrows -> String {
        let chars = Array(input); var i = 0, out = "", chunk = "", sheetQuote = false
        func flush() throws {
            let ns = chunk as NSString, matches = regex.matches(in:chunk,range:NSRange(location:0,length:ns.length))
            let result = NSMutableString(string:chunk)
            for m in matches.reversed() {
                guard let ref = GridRef(ns.substring(with:m.range(at:3))) else { continue }
                let rawSheet = m.range(at:2).location == NSNotFound ? nil : ns.substring(with:m.range(at:2))
                if let replacement = try replace(rawSheet,ref) { result.replaceCharacters(in:m.range,with:(rawSheet.map { $0+"!" } ?? "")+replacement) }
            }
            out += result as String; chunk = ""
        }
        while i<chars.count {
            if chars[i] == "'" { sheetQuote.toggle(); chunk.append(chars[i]); i += 1; continue }
            if chars[i] == "\"" && !sheetQuote {
                try flush(); out.append(chars[i]); i += 1
                while i<chars.count { let ch = chars[i]; out.append(ch); i += 1; if ch == "\"" { if i<chars.count && chars[i] == "\"" { out.append(chars[i]); i += 1 } else { break } } }
            } else { chunk.append(chars[i]); i += 1 }
        }
        try flush(); return out
    }
    static func sheetName(_ raw: String) -> String { raw.hasPrefix("'") ? String(raw.dropFirst().dropLast()).replacingOccurrences(of:"''",with:"'") : raw }
    static func unquoted(_ text: String) -> String {
        var out = "", quoted = false
        for ch in text { if ch == "\"" { quoted.toggle() }; if !quoted { out.append(ch) } }; return out
    }
}
func partTarget(_ base: String, _ target: String) -> String {
    let path = target.hasPrefix("/") ? target : "/"+(base as NSString).deletingLastPathComponent+"/"+target
    return String((path as NSString).standardizingPath.drop(while: { $0 == "/" }))
}
func relationshipsPart(_ path: String) -> String { (path as NSString).deletingLastPathComponent+"/_rels/"+(path as NSString).lastPathComponent+".rels" }
func elements(_ root: XMLElement) -> [XMLElement] {
    var result = [root], i = 0
    while i<result.count { result += result[i].children?.compactMap { $0 as? XMLElement } ?? []; i += 1 }; return result
}
struct PackageSnapshot { let documents: [String:Data]; let removed: Set<String>; let memberNames: Set<String>; let dirty: Bool }

func fillFormula(_ formula: String, down delta: Int, right horizontal: Int = 0) -> String {
    RefScanner.transform(formula) { _,ref in ref.copied(down:delta,right:horizontal)?.text ?? "#REF!" }
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

// MARK: - Charts (DrawingML in, SVG out)
let chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"
let drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
let drawingMainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
let relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
struct ChartSeries {
    var name: String, labels: [String], values: [Double?], x: [Double?], color: String, markers: Bool
}
struct ChartModel {
    var id: String, path: String, drawing: String, anchor: Int, title: String, kind: String, grouping: String
    var series: [ChartSeries], frame: CGRect, lastRow: Int, lastCol: Int
}
struct GridGeometry {
    var widths: [Int:Double] = [:], heights: [Int:Double] = [:], hiddenRows = Set<Int>(), hiddenCols = Set<Int>()
    init(_ root: XMLElement) {
        for col in root.child("cols")?.kids("col") ?? [] {
            let lo = Int(col.attr("min")) ?? 1,hi = min(16384,Int(col.attr("max")) ?? 1); guard lo<=hi else { continue }
            for c in lo...hi { widths[c] = min(620,max(48,(Double(col.attr("width")) ?? 12)*7+8)); if ["1","true"].contains(col.attr("hidden")) { hiddenCols.insert(c) } }
        }
        for row in root.child("sheetData")?.kids("row") ?? [] {
            let r = Int(row.attr("r")) ?? 0
            if let h = Double(row.attr("ht")),h>0 { heights[r] = h*4/3 }
            if ["1","true"].contains(row.attr("hidden")) { hiddenRows.insert(r) }
        }
    }
    func width(_ col: Int) -> Double { hiddenCols.contains(col) ? 0 : widths[col] ?? 110 }
    func height(_ row: Int) -> Double { hiddenRows.contains(row) ? 0 : heights[row] ?? 27 }
    func point(col: Int,row: Int,colOff: Double = 0,rowOff: Double = 0) -> CGPoint {
        var x = 40.0,y = 66.0
        if col>0 { for c in 1...min(col,16384) { x += width(c) } }; if row>0 { for r in 1...min(row,1048576) { y += height(r) } }
        return CGPoint(x:x+colOff/9525,y:y+rowOff/9525)
    }
    func coordinate(_ pixel: Double, rows: Bool) -> (Int,Int) {
        var remaining = max(0,pixel-(rows ? 66 : 40)),i = 1,limit = rows ? 1048576 : 16384
        while i<limit { let n = rows ? height(i) : width(i); if n>0 && remaining<n { break }; remaining -= n; i += 1 }
        return (i-1,Int((remaining*9525).rounded()))
    }
}
func chartNumber(_ n: Double) -> String { String(format:"%.2f",locale:Locale(identifier:"en_US_POSIX"),n.isFinite ? n : 0) }
let chartPalette = ["#4472C4","#ED7D31","#A5A5A5","#FFC000","#5B9BD5","#70AD47","#264478","#9E480E"]
func chartSVG(_ chart: ChartModel) -> String {
    let w = Double(chart.frame.width),h = Double(chart.frame.height)
    func n(_ x: Double) -> String { chartNumber(x) }
    func text(_ x: Double,_ y: Double,_ value: String,_ anchor: String = "middle",_ size: Int = 10) -> String { "<text x='\(n(x))' y='\(n(y))' text-anchor='\(anchor)' font-size='\(size)'>\(esc(value))</text>" }
    var svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100%' height='100%' viewBox='0 0 \(n(w)) \(n(h))' role='img' aria-label='\(esc(chart.title))' style='font-family:-apple-system,system-ui;color:#333'><title>\(esc(chart.title))</title><rect width='100%' height='100%' fill='white'/>"+text(w/2,24,chart.title,"middle",14)
    guard chart.kind != "unsupported" else { return svg+text(w/2,h/2,L("Anteprima non disponibile per questo tipo di grafico","Preview unavailable for this chart type"),"middle",11)+"</svg>" }
    let count = chart.series.map { $0.values.count }.max() ?? 0
    guard count>0,!chart.series.isEmpty else { return svg+text(w/2,h/2,L("Nessun dato","No data"))+"</svg>" }
    if chart.kind == "pie" || chart.kind == "doughnut" {
        let s = chart.series[0], values = s.values.map { abs($0 ?? 0) },total = values.reduce(0,+)
        let radius = max(15,min((w-155)/2,(h-55)/2)),cx = max(radius+10,(w-145)/2),cy = (h+25)/2,inner = chart.kind == "doughnut" ? radius*0.55 : 0
        var angle = -Double.pi/2
        for (i,v) in values.enumerated() where v>0 && total>0 {
            let end = angle+v/total*2*Double.pi,mid = (angle+end)/2
            func point(_ r: Double,_ a: Double) -> String { n(cx+r*cos(a))+" "+n(cy+r*sin(a)) }
            // Two arcs also represent a complete, single-slice circle correctly.
            var path = "M "+point(radius,angle)+" A \(n(radius)) \(n(radius)) 0 0 1 "+point(radius,mid)+" A \(n(radius)) \(n(radius)) 0 0 1 "+point(radius,end)
            if inner>0 { path += " L "+point(inner,end)+" A \(n(inner)) \(n(inner)) 0 0 0 "+point(inner,mid)+" A \(n(inner)) \(n(inner)) 0 0 0 "+point(inner,angle)+" Z" } else { path += " L \(n(cx)) \(n(cy)) Z" }
            let label = i<s.labels.count ? s.labels[i] : String(i+1),color = chartPalette[i%chartPalette.count]
            svg += "<path data-mark='slice' d='\(path)' fill='\(color)' stroke='white'><title>\(esc(label)): \(esc(Book.general.string(from:NSNumber(value:v)) ?? String(v)))</title></path>"
            angle = end
        }
        for i in values.indices { let label = i<s.labels.count ? s.labels[i] : String(i+1),y = 47+Double(i)*15; svg += "<rect x='\(n(w-137))' y='\(n(y-8))' width='9' height='9' fill='\(chartPalette[i%chartPalette.count])'/>"+text(w-122,y,String(label.prefix(20)),"start") }
        return svg+"</svg>"
    }
    let stacked = chart.grouping.lowercased().contains("stacked"),percent = chart.grouping.lowercased().contains("percent")
    var values = chart.series.map(\.values)
    if percent {
        for i in 0..<count {
            let positive = values.reduce(0.0) { $0+max(0,i<$1.count ? $1[i] ?? 0 : 0) },negative = values.reduce(0.0) { $0+abs(min(0,i<$1.count ? $1[i] ?? 0 : 0)) }
            for j in values.indices where i<values[j].count { if let v = values[j][i] { values[j][i] = v>=0 ? (positive>0 ? v/positive : 0) : (negative>0 ? v/negative : 0) } }
        }
    }
    var low = 0.0,high = 0.0
    for i in 0..<count { var pos = 0.0,neg = 0.0; for s in values { let v = i<s.count ? s[i] ?? 0 : 0; if stacked { if v>=0 { pos += v } else { neg += v } } else { pos = max(pos,v); neg = min(neg,v) } }; low = min(low,neg); high = max(high,pos) }
    func ticks(_ minimum: Double,_ maximum: Double) -> (Double,Double,[Double]) {
        let span = maximum>minimum ? maximum-minimum : max(1,abs(minimum)),raw = span/5,power = pow(10,floor(log10(max(raw,1e-200)))),unit = raw/power
        let step = (unit<=1 ? 1.0 : unit<=2 ? 2 : unit<=5 ? 5 : 10)*power,a = floor(minimum/step)*step,b = ceil((maximum>minimum ? maximum : minimum+span)/step)*step
        let count = min(20,max(1,Int(((b-a)/step).rounded())))
        return (a,b,(0...count).map { a+Double($0)*step })
    }
    let (minimum,maximum,ys) = ticks(low,high),left = chart.kind == "bar" ? 85.0 : 53,right = w-20,top = 43.0,bottom = max(90,h-62)
    func y(_ v: Double) -> Double { bottom-(v-minimum)/max(1e-200,maximum-minimum)*(bottom-top) }
    func vx(_ v: Double) -> Double { left+(v-minimum)/max(1e-200,maximum-minimum)*(right-left) }
    func label(_ v: Double) -> String { percent ? "\(Int((v*100).rounded()))%" : Book.general.string(from:NSNumber(value:v)) ?? String(v) }
    for tick in ys {
        if chart.kind == "bar" { svg += "<path d='M \(n(vx(tick))) \(n(top)) V \(n(bottom))' stroke='#e5e7eb' fill='none'/>"+text(vx(tick),bottom+16,label(tick)) }
        else { svg += "<path d='M \(n(left)) \(n(y(tick))) H \(n(right))' stroke='#e5e7eb' fill='none'/>"+text(left-7,y(tick)+3,label(tick),"end") }
    }
    svg += "<path d='M \(n(left)) \(n(top)) V \(n(bottom)) H \(n(right))' stroke='#9ca3af' fill='none'/>"
    let labels = chart.series.first?.labels ?? []
    func cat(_ i: Int) -> String { i<labels.count ? labels[i] : String(i+1) }
    if chart.kind == "column" || chart.kind == "bar" {
        let step = (chart.kind == "bar" ? bottom-top : right-left)/Double(count),group = max(1,Double(chart.series.count))
        var pos = [Double](repeating:0,count:count),neg = pos
        for (j,s) in chart.series.enumerated() { for i in s.values.indices {
            guard let value = values[j][i] else { continue }; let start = stacked ? (value>=0 ? pos[i] : neg[i]) : 0,end = start+value
            if value>=0 { pos[i] = end } else { neg[i] = end }
            let thickness = step*0.72/(stacked ? 1 : group),offset = step*0.14+(stacked ? 0 : Double(j)*thickness)
            let rect: CGRect
            if chart.kind == "bar" { rect = CGRect(x:min(vx(start),vx(end)),y:top+Double(i)*step+offset,width:abs(vx(end)-vx(start)),height:thickness) }
            else { rect = CGRect(x:left+Double(i)*step+offset,y:min(y(start),y(end)),width:thickness,height:abs(y(end)-y(start))) }
            svg += "<rect data-mark='bar' x='\(n(rect.minX))' y='\(n(rect.minY))' width='\(n(rect.width))' height='\(n(rect.height))' fill='\(s.color)'><title>\(esc(s.name)) · \(esc(cat(i))): \(esc(label(value)))</title></rect>"
        } }
        for i in stride(from:0,to:count,by:max(1,(count+7)/8)) { if chart.kind == "bar" { svg += text(left-6,top+(Double(i)+0.5)*step+3,String(cat(i).prefix(12)),"end") } else { svg += text(left+(Double(i)+0.5)*step,bottom+17,String(cat(i).prefix(14))) } }
    } else if chart.kind == "scatter" {
        let xs = chart.series.flatMap(\.x).compactMap { $0 },(xmin,xmax,xt) = ticks(xs.min() ?? 0,xs.max() ?? 1)
        func x(_ v: Double) -> Double { left+(v-xmin)/max(1e-200,xmax-xmin)*(right-left) }
        for v in xt { svg += text(x(v),bottom+17,label(v)) }
        for s in chart.series { for i in 0..<min(s.x.count,s.values.count) { if let xx = s.x[i],let yy = s.values[i] { svg += "<circle data-mark='point' cx='\(n(x(xx)))' cy='\(n(y(yy)))' r='3' fill='\(s.color)'><title>\(esc(s.name)): \(esc(label(xx))), \(esc(label(yy)))</title></circle>" } } }
    } else {
        var cumulative = [Double](repeating:0,count:count)
        for (j,s) in chart.series.enumerated() {
            var path = "",started = false,areaTop: [CGPoint] = [],areaBase: [CGPoint] = []
            for i in 0..<count {
                guard i<values[j].count,let v = values[j][i] else { started = false; continue }
                let x = left+Double(i)/Double(max(1,count-1))*(right-left),base = stacked ? cumulative[i] : 0,end = base+v
                cumulative[i] = end; path += (started ? " L " : " M ")+n(x)+" "+n(y(end)); started = true
                areaTop.append(CGPoint(x:x,y:y(end))); areaBase.append(CGPoint(x:x,y:y(base)))
                if s.markers && chart.kind == "line" { svg += "<circle data-mark='point' cx='\(n(x))' cy='\(n(y(end)))' r='3' fill='\(s.color)'/>" }
            }
            if chart.kind == "area",let first = areaTop.first {
                let polygon = ([first]+areaTop.dropFirst()+areaBase.reversed()).map { n($0.x)+","+n($0.y) }.joined(separator:" ")
                svg += "<polygon data-mark='area' points='\(polygon)' fill='\(s.color)' fill-opacity='0.45' stroke='\(s.color)'/>"
            } else { svg += "<path data-mark='line' d='\(path)' fill='none' stroke='\(s.color)' stroke-width='2'/>" }
        }
        for i in stride(from:0,to:count,by:max(1,(count+7)/8)) { svg += text(left+Double(i)/Double(max(1,count-1))*(right-left),bottom+17,String(cat(i).prefix(14))) }
    }
    var x = left
    for s in chart.series {
        svg += "<rect x='\(n(x))' y='\(n(h-23))' width='10' height='10' fill='\(s.color)'/>"+text(x+15,h-14,String(s.name.prefix(24)),"start")
        x += min(180,Double(s.name.count)*6+35)
    }
    return svg+"</svg>"
}
final class Book {
    var url: URL
    var untitled = false
    var temporaryDirectory: URL?
    deinit { if let dir = temporaryDirectory { try? FileManager.default.removeItem(at:dir) } }
    var documents: [String: XMLDocument] = [:]
    var removedMembers = Set<String>()
    var originalMembers: Set<String>?
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
        if let doc = documents[member] { return doc }
        guard !removedMembers.contains(member) else { throw failure(L("Parte del file eliminata.","Package part was deleted.")) }
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
    func part(_ path: String) throws -> XMLDocument {
        if let doc = documents[path] { return doc }; let doc = try xml(path); documents[path] = doc; return doc
    }
    func members() throws -> Set<String> {
        if originalMembers == nil {
            let p = Process(); p.executableURL = URL(fileURLWithPath:"/usr/bin/unzip"); p.arguments = ["-Z","-1",url.path]
            let pipe = Pipe(); p.standardOutput = pipe; p.standardError = FileHandle.nullDevice; try p.run()
            let data = pipe.fileHandleForReading.readDataToEndOfFile(); p.waitUntilExit()
            guard p.terminationStatus == 0 else { throw failure(L("Impossibile leggere il pacchetto XLSX.","Could not read the XLSX package.")) }
            originalMembers = Set(String(decoding:data,as:UTF8.self).split(separator:"\n").map(String.init))
        }
        return (originalMembers ?? []).union(documents.keys).subtracting(removedMembers)
    }
    func packageSnapshot() throws -> PackageSnapshot { PackageSnapshot(documents:documents.mapValues { $0.xmlData },removed:removedMembers,memberNames:try members(),dirty:dirty) }
    func restorePackage(_ state: PackageSnapshot) throws {
        let discarded = try members().subtracting(state.memberNames)
        documents = try state.documents.mapValues { data in
            let doc = try XMLDocument(data:data,options:[.nodePreserveWhitespace,.nodeLoadExternalEntitiesNever])
            // Foundation defaults an omitted standalone declaration to yes on reparse.
            // Restore its serialized spelling too; copying XMLDocument also mutates that flag.
            let header = String((String(data:data,encoding:.utf8) ?? String(data:data,encoding:.utf16) ?? "").prefix(200).prefix(while:{ $0 != ">" }))
            doc.isStandalone = header.range(of:"standalone\\s*=\\s*[\"']yes[\"']",options:.regularExpression) != nil
            return doc
        }
        removedMembers = state.removed.union(discarded); dirty = state.dirty; if let styles = documents["xl/styles.xml"] { readStyles(styles) }; invalidate()
    }
    func linked(_ path: String, type: String) throws -> [(String,String)] {
        let rp = relationshipsPart(path)
        guard try members().contains(rp) else { return [] }
        return try part(rp).rootElement()?.kids("Relationship").filter { $0.attr("Type").hasSuffix("/"+type) && $0.attr("TargetMode") != "External" }.map { ($0.attr("Id"),partTarget(path,$0.attr("Target"))) } ?? []
    }
    func preparePackage() throws {
        let names = try members()
        for s in sheets { _ = try sheetDocument(sheets.firstIndex(where:{ $0.path == s.path })!); if names.contains(relationshipsPart(s.path)) { _ = try part(relationshipsPart(s.path)) } }
        for name in names where name == "xl/workbook.xml" || name == "xl/_rels/workbook.xml.rels" || name == "[Content_Types].xml" || name.hasPrefix("xl/tables/") && name.hasSuffix(".xml") || name.hasPrefix("xl/charts/") && (name.hasSuffix(".xml") || name.hasSuffix(".rels")) || name.hasPrefix("xl/drawings/") && (name.hasSuffix(".xml") || name.hasSuffix(".rels")) { _ = try part(name) }
    }
    func rewrite(_ text: String, context: Int?, target: Int, change: GridChange) throws -> String {
        try RefScanner.transform(text) { rawSheet, ref in
            let applies: Bool
            if let rawSheet { applies = RefScanner.sheetName(rawSheet).caseInsensitiveCompare(sheets[target].name) == .orderedSame }
            else { applies = context == target }
            return applies ? try ref.shifted(change)?.text ?? "#REF!" : nil
        }
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
    func formulaText(_ cell: XMLElement, sheet: Int? = nil) -> String? {
        if let f = cell.child("f"), f.attr("t") == "shared", (f.stringValue ?? "").isEmpty, let sheet, let master = try? index(sheet).sharedMasters[f.attr("si")], let mf = master.child("f") {
            return fillFormula(mf.stringValue ?? "",down:rowNumber(cell.attr("r"))-rowNumber(master.attr("r")),right:column(cell.attr("r"))-column(master.attr("r")))
        }
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
                if let f = cell.child("f"), f.attr("t") == "shared", !(f.stringValue ?? "").isEmpty { result.sharedMasters[f.attr("si")] = cell }
                if cell.child("v") == nil && cell.child("is") == nil && cell.child("f") == nil { result.blankCells[cellKey(r,c)] = cell; if !cell.attr("s").isEmpty && cell.attr("s") != "0" { result.maxRow = max(result.maxRow,r); result.maxCol = max(result.maxCol,c) }; continue }
                result.cells[cellKey(r, c)] = cell; result.maxRow = max(result.maxRow, r); result.maxCol = max(result.maxCol, c)
            }
        }
        indexes[sheet] = result
        return result
    }
    /// Value of a formula cell computed by this app; throws when the formula uses something unsupported.
    func computed(_ cell: XMLElement, sheet: Int) throws -> FValue {
        guard let text = formulaText(cell,sheet:sheet) else { throw FormulaUnsupported() }
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
        if let f = cell.child("f") { return "=" + (formulaText(cell,sheet:sheet) ?? f.stringValue ?? "") }
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
                else { row = partElement("row",like:data); row.set("r",String(r)); let at = (data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0)>r } ?? data.childCount; data.insertChild(row,at:at); rows[r] = row }
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
        guard doc.rootElement()?.child("sheetProtection") == nil else { throw failure(L("Questo foglio è protetto. Rimuovi la protezione in Excel prima di modificarlo.", "This sheet is protected. Remove the protection in Excel before editing it.")) }
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
            else { row = partElement("row",like:data); row.set("r",String(r)); let place = (data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0)>r } ?? data.childCount; data.insertChild(row,at:place); rows[r] = row }
            let c: XMLElement
            if let existing = cells[key] { c = existing }
            else { c = partElement("c",like:row); c.set("r",edit.ref); let place = (row.children ?? []).firstIndex { column($0.attr("r"))>col } ?? row.childCount; row.insertChild(c,at:place); cells[key] = c }
            for node in c.kids("f") + c.kids("v") + c.kids("is") { node.detach() }; c.removeAttribute(forName:"t")
            if !edit.value.isEmpty {
                switch edit.kind {
                case .formula: c.addChild(partElement("f",like:c,text:normalizeFormula(edit.value)))
                case .number: c.addChild(partElement("v",like:c,text:edit.value))
                case .text:
                    c.set("t","inlineStr"); let inline = partElement("is",like:c), t = partElement("t",like:c,text:edit.value); t.set("xml:space","preserve"); inline.addChild(t); c.addChild(inline)
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
    func nextPart(_ prefix: String) throws -> String {
        let used = try members(); var i = 1; while used.contains(prefix+String(i)+".xml") { i += 1 }; return prefix+String(i)+".xml"
    }
    func contentType(_ path: String, _ type: String) throws {
        let root = try part("[Content_Types].xml").rootElement()!
        if root.kids("Override").contains(where:{ $0.attr("PartName") == "/"+path }) { return }
        let e = partElement("Override",like:root); e.set("PartName","/"+path); e.set("ContentType",type); root.addChild(e)
    }
    func addRelationship(_ source: String, target: String, type: String) throws -> String {
        let path = relationshipsPart(source),doc: XMLDocument
        if try members().contains(path) { doc = try part(path) }
        else { doc = try XMLDocument(xmlString:"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>",options:[]); documents[path] = doc }
        let root = doc.rootElement()!,used = Set(root.kids("Relationship").map { $0.attr("Id") }); var i = 1; while used.contains("rId\(i)") { i += 1 }
        var a = (source as NSString).deletingLastPathComponent.split(separator:"/").map(String.init),b = target.split(separator:"/").map(String.init)
        while !a.isEmpty && !b.isEmpty && a[0] == b[0] { a.removeFirst(); b.removeFirst() }
        let relative = (Array(repeating:"..",count:a.count)+b).joined(separator:"/"),e = partElement("Relationship",like:root)
        e.set("Id","rId\(i)"); e.set("Type",relationshipNamespace+"/"+type); e.set("Target",relative); root.addChild(e)
        return "rId\(i)"
    }
    func markerXML(_ tag: String, col: Int,row: Int,colOff: Int = 0,rowOff: Int = 0) -> String {
        "<xdr:\(tag)><xdr:col>\(col)</xdr:col><xdr:colOff>\(colOff)</xdr:colOff><xdr:row>\(row)</xdr:row><xdr:rowOff>\(rowOff)</xdr:rowOff></xdr:\(tag)>"
    }
    @discardableResult func createChart(_ sheet: Int, range: CellRange, kind: String) throws -> String {
        try editable(sheet)
        guard ["column","bar","line","pie"].contains(kind),range.count>1,range.count<=200000 else { throw failure(L("Seleziona un intervallo con dati per il grafico.","Select a data range for the chart.")) }
        let backup = try packageSnapshot()
        do {
            let first = try (range.c1...range.c2).map { try value(sheet,range.r1,$0) }
            let header = first.contains { if case .text(let s) = $0 { return !s.isEmpty }; return false } && first.allSatisfy { if case .text = $0 { return true }; if case .empty = $0 { return true }; return false }
            let start = range.r1+(header ? 1 : 0); guard start<=range.r2 else { throw failure(L("L'intervallo contiene solo intestazioni.","The range only contains headers.")) }
            let firstCol = try (start...range.r2).map { try value(sheet,$0,range.c1) }
            let categories = range.c1<range.c2 && firstCol.contains { if case .text = $0 { return true }; return false } && firstCol.allSatisfy { if case .text = $0 { return true }; if case .empty = $0 { return true }; return false }
            var columns: [Int] = []
            for c in (range.c1+(categories ? 1 : 0))...range.c2 { if try (start...range.r2).contains(where:{ if case .number = try value(sheet,$0,c) { return true }; return false }) { columns.append(c) } }
            if kind == "pie" { columns = Array(columns.prefix(1)) }
            guard !columns.isEmpty else { throw failure(L("L'intervallo non contiene serie numeriche.","The range contains no numeric series.")) }
            let sheetName = "'"+sheets[sheet].name.replacingOccurrences(of:"'",with:"''")+"'!"
            func reference(_ c: Int,_ r1: Int,_ r2: Int) -> String { sheetName+"$"+columnName(c)+"$\(r1):$"+columnName(c)+"$\(r2)" }
            func cache(_ tag: String,_ strings: [String?],numeric: Bool) -> String {
                "<c:\(tag)>"+(numeric ? "<c:formatCode>General</c:formatCode>" : "")+"<c:ptCount val='\(strings.count)'/>"+strings.enumerated().compactMap { i,s in s.map { "<c:pt idx='\(i)'><c:v>\(esc($0))</c:v></c:pt>" } }.joined()+"</c:\(tag)>"
            }
            var seriesXML = ""
            for (j,c) in columns.enumerated() {
                let name = header ? try raw(sheet,columnName(c)+String(range.r1)) : L("Serie \(j+1)","Series \(j+1)")
                let tx = header ? "<c:tx><c:strRef><c:f>\(esc(sheetName+"$"+columnName(c)+"$\(range.r1)"))</c:f>\(cache("strCache",[name],numeric:false))</c:strRef></c:tx>" : "<c:tx><c:v>\(esc(name))</c:v></c:tx>"
                let labels: [String?] = try (start...range.r2).map { categories ? try cell(sheet,columnName(range.c1)+String($0)).map { display($0,sheet:sheet) } ?? "" : String($0-start+1) }
                let cat = categories ? "<c:cat><c:strRef><c:f>\(esc(reference(range.c1,start,range.r2)))</c:f>\(cache("strCache",labels,numeric:false))</c:strRef></c:cat>" : "<c:cat>\(cache("strLit",labels,numeric:false))</c:cat>"
                let vals: [String?] = try (start...range.r2).map { r in if case .number(let n) = try value(sheet,r,c),n.isFinite { return FValue.number(n).text }; return nil }
                let val = "<c:val><c:numRef><c:f>\(esc(reference(c,start,range.r2)))</c:f>\(cache("numCache",vals,numeric:true))</c:numRef></c:val>"
                let marker = kind == "line" ? "<c:marker><c:symbol val='circle'/><c:size val='5'/></c:marker>" : ""
                seriesXML += "<c:ser><c:idx val='\(j)'/><c:order val='\(j)'/>"+tx+"<c:spPr><a:solidFill><a:srgbClr val='\(String(chartPalette[j%chartPalette.count].dropFirst()))'/></a:solidFill></c:spPr>"+marker+cat+val+(kind == "line" ? "<c:smooth val='0'/>" : "")+"</c:ser>"
            }
            let title = columns.count == 1 && header ? try raw(sheet,columnName(columns[0])+String(range.r1)) : L("Grafico","Chart")
            let titleXML = "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>\(esc(title))</a:t></a:r></a:p></c:rich></c:tx><c:overlay val='0'/></c:title>"
            let tag = kind == "line" ? "lineChart" : kind == "pie" ? "pieChart" : "barChart"
            let prefix = tag == "barChart" ? "<c:barDir val='\(kind == "bar" ? "bar" : "col")'/><c:grouping val='clustered'/>" : tag == "lineChart" ? "<c:grouping val='standard'/>" : ""
            let axes = kind == "pie" ? "" : "<c:catAx><c:axId val='100001'/><c:scaling><c:orientation val='minMax'/></c:scaling><c:delete val='0'/><c:axPos val='\(kind == "bar" ? "l" : "b")'/><c:tickLblPos val='nextTo'/><c:crossAx val='100002'/><c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/><c:lblOffset val='100'/></c:catAx><c:valAx><c:axId val='100002'/><c:scaling><c:orientation val='minMax'/></c:scaling><c:delete val='0'/><c:axPos val='\(kind == "bar" ? "b" : "l")'/><c:majorGridlines/><c:numFmt formatCode='General' sourceLinked='1'/><c:tickLblPos val='nextTo'/><c:crossAx val='100001'/><c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
            let tail = kind == "line" ? "<c:marker val='1'/><c:smooth val='0'/>" : kind == "pie" ? "<c:firstSliceAng val='0'/>" : "<c:gapWidth val='150'/><c:overlap val='0'/>"
            let path = try nextPart("xl/charts/chart")
            let xml = "<c:chartSpace xmlns:c='\(chartNamespace)' xmlns:a='\(drawingMainNamespace)'><c:chart>"+titleXML+"<c:autoTitleDeleted val='0'/><c:plotArea><c:layout/><c:\(tag)>"+prefix+"<c:varyColors val='\(kind == "pie" ? "1" : "0")'/>"+seriesXML+tail+(kind == "pie" ? "" : "<c:axId val='100001'/><c:axId val='100002'/>")+"</c:\(tag)>"+axes+"</c:plotArea><c:legend><c:legendPos val='b'/><c:overlay val='0'/></c:legend><c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart></c:chartSpace>"
            documents[path] = try XMLDocument(xmlString:xml,options:[]); try contentType(path,"application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            let sheetRoot = try sheetDocument(sheet).rootElement()!,drawing: String
            if let old = try linked(sheets[sheet].path,type:"drawing").first { drawing = old.1 }
            else {
                drawing = try nextPart("xl/drawings/drawing"); documents[drawing] = try XMLDocument(xmlString:"<xdr:wsDr xmlns:xdr='\(drawingNamespace)' xmlns:a='\(drawingMainNamespace)'/>",options:[])
                let id = try addRelationship(sheets[sheet].path,target:drawing,type:"drawing"),node = partElement("drawing",like:sheetRoot); node.addNamespace(XMLNode.namespace(withName:"r",stringValue:relationshipNamespace) as! XMLNode); node.set("r:id",id)
                let after = ["legacyDrawing","legacyDrawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"]
                let at = (sheetRoot.children ?? []).firstIndex { after.contains($0.localName ?? "") } ?? sheetRoot.childCount; sheetRoot.insertChild(node,at:at)
                try contentType(drawing,"application/vnd.openxmlformats-officedocument.drawing+xml")
            }
            let rid = try addRelationship(drawing,target:path,type:"chart"),dr = try part(drawing).rootElement()!
            let id = (elements(dr).filter { $0.localName == "cNvPr" }.compactMap { Int($0.attr("id")) }.max() ?? 0)+1
            let fromCol = min(16376,range.c2+1),fromRow = min(1048561,max(0,range.r1-1))
            let anchorXML = "<xdr:twoCellAnchor xmlns:xdr='\(drawingNamespace)' xmlns:a='\(drawingMainNamespace)' xmlns:c='\(chartNamespace)' xmlns:r='\(relationshipNamespace)' editAs='oneCell'>"+markerXML("from",col:fromCol,row:fromRow)+markerXML("to",col:fromCol+6,row:fromRow+12)+"<xdr:graphicFrame macro=''><xdr:nvGraphicFramePr><xdr:cNvPr id='\(id)' name='\(esc(L("Grafico \(id)","Chart \(id)")))'/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x='0' y='0'/><a:ext cx='0' cy='0'/></xdr:xfrm><a:graphic><a:graphicData uri='\(chartNamespace)'><c:chart r:id='\(rid)'/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"
            let anchor = try XMLElement(xmlString:anchorXML),at = (dr.children ?? []).firstIndex { $0.localName == "extLst" } ?? dr.childCount; dr.insertChild(anchor,at:at)
            dirty = true; invalidate()
            return drawing+"#\((dr.children?.compactMap { $0 as? XMLElement } ?? []).firstIndex(where:{ $0 === anchor })!)"
        } catch { try restorePackage(backup); throw error }
    }
    func chartAnchor(_ sheet: Int, id: String) throws -> (ChartModel,XMLElement,XMLElement) {
        guard let model = try readCharts(sheet).first(where:{ $0.id == id }),let root = try part(model.drawing).rootElement() else { throw failure(L("Grafico non trovato.","Chart not found.")) }
        let children = root.children?.compactMap { $0 as? XMLElement } ?? []
        guard children.indices.contains(model.anchor) else { throw failure(L("Ancoraggio del grafico non valido.","Invalid chart anchor.")) }; return (model,root,children[model.anchor])
    }
    func setChartFrame(_ sheet: Int, id: String, frame: CGRect) throws {
        try editable(sheet)
        guard frame.minX.isFinite,frame.minY.isFinite,frame.width.isFinite,frame.height.isFinite,frame.width>=120,frame.height>=100 else { throw failure(L("Dimensioni del grafico non valide.","Invalid chart dimensions.")) }
        let (_,root,old) = try chartAnchor(sheet,id:id),geometry = GridGeometry(try sheetDocument(sheet).rootElement()!)
        let x1 = geometry.coordinate(max(40,frame.minX),rows:false),y1 = geometry.coordinate(max(66,frame.minY),rows:true),x2 = geometry.coordinate(max(40,frame.minX)+frame.width,rows:false),y2 = geometry.coordinate(max(66,frame.minY)+frame.height,rows:true)
        let node = try XMLElement(xmlString:"<xdr:twoCellAnchor xmlns:xdr='\(drawingNamespace)' xmlns:a='\(drawingMainNamespace)' xmlns:c='\(chartNamespace)' xmlns:r='\(relationshipNamespace)' editAs='oneCell'>"+markerXML("from",col:x1.0,row:y1.0,colOff:x1.1,rowOff:y1.1)+markerXML("to",col:x2.0,row:y2.0,colOff:x2.1,rowOff:y2.1)+"</xdr:twoCellAnchor>")
        for child in old.children?.compactMap({ $0 as? XMLElement }) ?? [] where !["from","to","pos","ext"].contains(child.localName ?? "") { node.addChild(child.copy() as! XMLElement) }
        let at = (root.children ?? []).firstIndex { $0 === old }!; old.detach(); root.insertChild(node,at:at); dirty = true
    }
    func removePackagePart(_ path: String) throws {
        documents.removeValue(forKey:path); removedMembers.insert(path)
        if let root = try part("[Content_Types].xml").rootElement() { for e in root.kids("Override") where e.attr("PartName") == "/"+path { e.detach() } }
    }
    func deleteChart(_ sheet: Int, id: String) throws {
        try editable(sheet); let (model,root,anchor) = try chartAnchor(sheet,id:id)
        let rid = elements(anchor).first(where:{ $0.localName == "chart" })?.attr("r:id") ?? ""
        anchor.detach()
        if !elements(root).contains(where:{ $0.localName == "chart" && $0.attr("r:id") == rid }),let rels = try part(relationshipsPart(model.drawing)).rootElement() { rels.kids("Relationship").filter { $0.attr("Id") == rid }.forEach { $0.detach() } }
        if root.children?.contains(where:{ $0 is XMLElement && $0.localName != "extLst" }) != true {
            try sheetDocument(sheet).rootElement()?.child("drawing")?.detach()
            if let rels = try part(relationshipsPart(sheets[sheet].path)).rootElement() { rels.kids("Relationship").filter { partTarget(sheets[sheet].path,$0.attr("Target")) == model.drawing }.forEach { $0.detach() } }
            try removePackagePart(model.drawing); try removePackagePart(relationshipsPart(model.drawing))
        }
        var referenced = false
        for path in try members() where path.hasSuffix(".rels") {
            let source = ((path as NSString).deletingLastPathComponent as NSString).deletingLastPathComponent+"/"+String((path as NSString).lastPathComponent.dropLast(5))
            if try part(path).rootElement()?.kids("Relationship").contains(where:{ $0.attr("Type").hasSuffix("/chart") && partTarget(source,$0.attr("Target")) == model.path }) == true { referenced = true; break }
        }
        if !referenced { try removePackagePart(model.path); let rel = relationshipsPart(model.path); if try members().contains(rel) { try removePackagePart(rel) } }
        dirty = true; invalidate()
    }
    func chartReference(_ text: String, sheet: Int) throws -> [(FValue,String)]? {
        // An intentionally deleted reference must never revive stale cached points.
        if text.trimmingCharacters(in:.whitespacesAndNewlines).hasSuffix("#REF!") { return [] }
        var parsed: (Int,GridRef)?
        let source = text.hasPrefix("=") ? String(text.dropFirst()) : text
        _ = RefScanner.transform(source) { raw,ref in
            let canonical = (raw.map { $0+"!" } ?? "")+ref.text
            guard canonical.caseInsensitiveCompare(source) == .orderedSame else { return nil }
            let target = raw.flatMap { name in sheets.firstIndex { $0.name.caseInsensitiveCompare(RefScanner.sheetName(name)) == .orderedSame } } ?? (raw == nil ? sheet : -1)
            if target>=0 { parsed = (target,ref) }; return nil
        }
        guard let (i,ref) = parsed else { return nil }; let idx = try index(i)
        let r1 = ref.top,r2 = max(ref.first.row ?? max(1,idx.maxRow),ref.last.row ?? max(1,idx.maxRow))
        let c1 = ref.left,c2 = max(ref.first.col ?? max(1,idx.maxCol),ref.last.col ?? max(1,idx.maxCol))
        guard (r2-r1+1)*(c2-c1+1)<=200000 else { return nil }
        var result: [(FValue,String)] = []
        for r in r1...r2 { for c in c1...c2 {
            let cell = idx.cells[cellKey(r,c)]
            if let cell,cell.child("f") != nil,(try? computed(cell,sheet:i)) == nil,(cell.child("v")?.stringValue ?? "").isEmpty { result.append((.error("#UNSUPPORTED!"),"")) }
            else { let v = try value(i,r,c); result.append((v,cell.map { display($0,sheet:i) } ?? "")) }
        } }; return result
    }
    func chartCache(_ container: XMLElement?) -> [String?] {
        guard let container else { return [] }
        let points = container.kids("pt"), count = min(200000,max(Int(container.child("ptCount")?.attr("val") ?? "") ?? 0,(points.compactMap { Int($0.attr("idx")) }.max() ?? -1)+1))
        guard count>0 else { return [] }; var result = [String?](repeating:nil,count:count)
        for p in points { if let i = Int(p.attr("idx")),result.indices.contains(i) { result[i] = p.child("v")?.stringValue } }; return result
    }
    func chartValues(_ source: XMLElement?, sheet: Int, numeric: Bool) throws -> [(FValue,String)] {
        guard let source else { return [] }
        let ref = source.child("numRef") ?? source.child("strRef"), literal = source.child("numLit") ?? source.child("strLit")
        let cache = chartCache(ref?.child("numCache") ?? ref?.child("strCache") ?? literal)
        if let formula = ref?.child("f")?.stringValue,let values = try chartReference(formula,sheet:sheet) {
            return values.enumerated().map { j,pair in
                // A missing formula result can use the chart cache; an actual blank remains a gap.
                if case .error("#UNSUPPORTED!") = pair.0,j<cache.count,let s = cache[j] { return (numeric ? Double(s).map(FValue.number) ?? .empty : .text(s),s) }
                return pair
            }
        }
        return cache.map { s in guard let s else { return (.empty,"") }; return (numeric ? Double(s).map(FValue.number) ?? .empty : .text(s),s) }
    }
    func readCharts(_ sheet: Int) throws -> [ChartModel] {
        guard try sheetDocument(sheet).rootElement()?.child("drawing") != nil else { return [] }
        var models: [ChartModel] = []; let geometry = GridGeometry(try sheetDocument(sheet).rootElement()!)
        for (_,drawing) in try linked(sheets[sheet].path,type:"drawing") {
            guard let root = try part(drawing).rootElement() else { continue }
            let links = Dictionary(uniqueKeysWithValues:try linked(drawing,type:"chart"))
            for (i,anchor) in (root.children?.compactMap { $0 as? XMLElement } ?? []).enumerated() {
                guard let link = elements(anchor).first(where:{ $0.localName == "chart" }),let path = links[link.attr("r:id")],let chartSpace = try part(path).rootElement() else { continue }
                let chart = chartSpace.child("chart"),plot = chart?.child("plotArea")
                let plotTypes = plot?.children?.compactMap { $0 as? XMLElement }.filter { ($0.localName ?? "").hasSuffix("Chart") } ?? []
                let type = plotTypes.first, rawKind = plotTypes.count == 1 ? type?.localName ?? "" : "combination"
                var kind = rawKind
                if rawKind == "barChart" { kind = type?.child("barDir")?.attr("val") == "bar" ? "bar" : "column" }
                else { kind = ["lineChart":"line","pieChart":"pie","doughnutChart":"doughnut","areaChart":"area","scatterChart":"scatter"][rawKind] ?? "unsupported" }
                let titleNode = chart?.child("title"), titleText = titleNode.map { elements($0).filter { $0.localName == "t" }.compactMap(\.stringValue).joined() } ?? ""
                let titleValues = try chartValues(titleNode?.child("tx"),sheet:sheet,numeric:false)
                let title = !titleText.isEmpty ? titleText : titleValues.first?.1 ?? L("Grafico","Chart")
                var series: [ChartSeries] = []
                for (j,ser) in (type?.kids("ser") ?? []).enumerated() {
                    let names = try chartValues(ser.child("tx"),sheet:sheet,numeric:false)
                    let name = ser.child("tx")?.child("v")?.stringValue ?? names.first?.1 ?? L("Serie \(j+1)","Series \(j+1)")
                    let cat = try chartValues(ser.child("cat"),sheet:sheet,numeric:false),values = try chartValues(ser.child(kind == "scatter" ? "yVal" : "val"),sheet:sheet,numeric:true),xs = try chartValues(ser.child("xVal"),sheet:sheet,numeric:true)
                    func number(_ pair: (FValue,String)) -> Double? { if case .number(let n) = pair.0,n.isFinite { return n }; return nil }
                    let rgb = ser.child("spPr")?.child("solidFill")?.child("srgbClr")?.attr("val") ?? ""
                    let color = rgb.range(of:"^[0-9a-fA-F]{6}$",options:.regularExpression) != nil ? "#"+rgb : chartPalette[j%chartPalette.count]
                    let marker = ser.child("marker")?.child("symbol")?.attr("val")
                    let markers = kind == "scatter" || (marker != nil ? marker != "none" : type?.child("marker")?.attr("val") == "1")
                    series.append(ChartSeries(name:name,labels:cat.map(\.1),values:values.map(number),x:xs.map(number),color:color,markers:markers))
                }
                func markerPoint(_ e: XMLElement?) -> CGPoint { geometry.point(col:Int(e?.child("col")?.stringValue ?? "") ?? 0,row:Int(e?.child("row")?.stringValue ?? "") ?? 0,colOff:Double(e?.child("colOff")?.stringValue ?? "") ?? 0,rowOff:Double(e?.child("rowOff")?.stringValue ?? "") ?? 0) }
                let from = anchor.child("from"),to = anchor.child("to")
                let point: CGPoint
                if from != nil { point = markerPoint(from) } else { point = CGPoint(x:40+(Double(anchor.child("pos")?.attr("x") ?? "") ?? 0)/9525,y:66+(Double(anchor.child("pos")?.attr("y") ?? "") ?? 0)/9525) }
                let frame: CGRect
                if let to { let end = markerPoint(to); frame = CGRect(x:point.x,y:point.y,width:max(120,end.x-point.x),height:max(100,end.y-point.y)) }
                else { frame = CGRect(x:point.x,y:point.y,width:max(120,(Double(anchor.child("ext")?.attr("cx") ?? "") ?? 4572000)/9525),height:max(100,(Double(anchor.child("ext")?.attr("cy") ?? "") ?? 2857500)/9525)) }
                models.append(ChartModel(id:drawing+"#\(i)",path:path,drawing:drawing,anchor:i,title:title,kind:kind,grouping:type?.child("grouping")?.attr("val") ?? "standard",series:series,frame:frame,lastRow:Int(to?.child("row")?.stringValue ?? from?.child("row")?.stringValue ?? "") ?? 15,lastCol:Int(to?.child("col")?.stringValue ?? from?.child("col")?.stringValue ?? "") ?? 8))
            }
        }; return models
    }
    func chartHTML(_ sheet: Int) throws -> String {
        try readCharts(sheet).map { model in
            "<div class='chart-box' tabindex='0' data-chart='\(esc(model.id))' style='left:\(chartNumber(model.frame.minX))px;top:\(chartNumber(model.frame.minY))px;width:\(chartNumber(model.frame.width))px;height:\(chartNumber(model.frame.height))px'>"+chartSVG(model)+"<div class='chart-resize' aria-label='\(esc(L("Ridimensiona grafico","Resize chart")))'></div></div>"
        }.joined()
    }
    func refreshChartCaches(_ doc: XMLDocument, sheet: Int) throws {
        guard let root = doc.rootElement() else { return }
        for source in elements(root) where ["strRef","numRef"].contains(source.localName ?? "") {
            guard let f = source.child("f")?.stringValue,let values = try chartReference(f,sheet:sheet) else { continue }
            if values.contains(where:{ if case .error("#UNSUPPORTED!") = $0.0 { return true }; return false }) { continue }
            let numeric = source.localName == "numRef",name = numeric ? "numCache" : "strCache",cache = XMLElement(name:"c:"+name),count = XMLElement(name:"c:ptCount")
            if numeric { cache.addChild(XMLElement(name:"c:formatCode",stringValue:source.child(name)?.child("formatCode")?.stringValue ?? "General")) }
            count.set("val",String(values.count)); cache.addChild(count)
            for (i,pair) in values.enumerated() {
                let text: String
                if numeric { guard case .number(let n) = pair.0,n.isFinite else { continue }; text = FValue.number(n).text } else { text = pair.1 }
                let point = XMLElement(name:"c:pt"); point.set("idx",String(i)); point.addChild(XMLElement(name:"c:v",stringValue:text)); cache.addChild(point)
            }
            // The new nodes declare their chart namespace, even when the source uses a default namespace.
            cache.addNamespace(XMLNode.namespace(withName:"c",stringValue:chartNamespace) as! XMLNode)
            source.child(name)?.detach(); source.addChild(cache)
        }
    }
    func structuralPreflight(_ sheet: Int, _ change: GridChange) throws {
        try editable(sheet)
        guard change.at>0,change.count>0,change.at+change.count-1<=change.limit else { throw failure(L("Intervallo non valido.","Invalid range.")) }
        let names = try members()
        if names.contains(where:{ let p = $0.lowercased(); return p.contains("/pivot") || p.contains("/externallinks/") || p.contains("/querytables/") || p == "xl/connections.xml" }) {
            throw failure(L("Inserimento/eliminazione non disponibile: il workbook contiene pivot, collegamenti esterni o connessioni dati che non posso aggiornare con sicurezza.","Cannot insert/delete: this workbook contains pivots, external links or data connections that cannot be safely updated."))
        }
        var owners = Set<String>()
        for s in sheets { for kind in ["drawing","table"] { for (_,path) in try linked(s.path,type:kind) { guard owners.insert(path).inserted else { throw failure(L("Una parte disegno/tabella è condivisa da più fogli: lo spostamento non è sicuro.","A drawing/table part is shared across sheets; shifting it is unsafe.")) } } } }
        for (i,_) in sheets.enumerated() {
            guard let root = try sheetDocument(i).rootElement() else { continue }
            if elements(root).contains(where:{ $0.localName == "extLst" }) || root.child("customSheetViews") != nil || i == sheet && (root.child("legacyDrawing") != nil || root.localName != "worksheet") {
                throw failure(L("Il foglio contiene estensioni, viste personalizzate o disegni legacy che non posso spostare con sicurezza.","The sheet contains extensions, custom views or legacy drawings that cannot be safely shifted."))
            }
            if i == sheet, !(try linked(sheets[i].path,type:"comments")).isEmpty { throw failure(L("Sposta prima i commenti legacy in Excel.","Move legacy comments in Excel before changing the grid.")) }
            for e in elements(root) {
                if e.localName == "f",e.attr("t") == "dataTable" { throw failure(L("Le tabelle di simulazione richiedono Excel per modificare righe o colonne.","What-if data tables require Excel for row/column changes.")) }
                if i == sheet,e.localName == "f",e.attr("t") == "array",let ref = GridRef(e.attr("ref")) {
                    let a = change.rows ? ref.first.row : ref.first.col, b = change.rows ? ref.last.row : ref.last.col
                    if let a,let b {
                        let lo = min(a,b),hi = max(a,b)
                        let cuts = change.deleting ? change.at<=hi && change.at+change.count-1>=lo && !(change.at<=lo && change.at+change.count-1>=hi) : change.at>lo && change.at<=hi
                        if cuts { throw failure(L("Non posso modificare solo una parte di una formula matriciale.","Cannot change only part of an array formula.")) }
                    }
                }
            }
        }
        for (path,doc) in documents where path == "xl/workbook.xml" || path.hasPrefix("xl/tables/") || path.hasPrefix("xl/charts/") || sheets.contains(where:{ $0.path == path }) {
            guard let root = doc.rootElement() else { continue }
            for e in elements(root) where ["f","formula","formula1","formula2","definedName","calculatedColumnFormula","totalsRowFormula"].contains(e.localName ?? "") {
                let text = RefScanner.unquoted(e.stringValue ?? "")
                if text.contains("[") || text.contains("]") || text.range(of:"[A-Za-z0-9_.']+:[A-Za-z0-9_.']+!",options:.regularExpression) != nil {
                    throw failure(L("Il workbook usa riferimenti strutturati, esterni o 3D: aggiornarli automaticamente non è sicuro.","This workbook uses structured, external or 3D references; they cannot be safely rewritten."))
                }
                _ = try RefScanner.transform(text) { raw,ref in
                    if let raw,RefScanner.sheetName(raw).contains(":") { throw failure(L("I riferimenti 3D richiedono Excel.","3D references require Excel.")) }
                    if e.localName == "definedName",e.attr("localSheetId").isEmpty,raw == nil {
                        throw failure(L("Un nome globale usa riferimenti relativi senza foglio: non posso determinarne l'origine.","A global defined name uses unqualified relative references; its origin cannot be determined safely."))
                    }; return nil
                }
            }
        }
        if change.deleting && change.rows {
            for (_,path) in try linked(sheets[sheet].path,type:"table") {
                let root = try part(path).rootElement()
                if let ref = root.flatMap({ GridRef($0.attr("ref")) }),let next = try ref.shifted(change) {
                    let headers = root?.attr("headerRowCount") == "0" ? 0 : 1,totals = Int(root?.attr("totalsRowCount") ?? "") ?? 0,last = max(ref.first.row!,ref.last.row!)
                    if max(next.first.row!,next.last.row!)-next.top+1<headers+totals+1 || totals>0 && change.at<=last && change.at+change.count>last {
                        throw failure(L("L'eliminazione dell'ultima riga dati o dei totali di una tabella richiede Excel.","Deleting the last table data row or its totals row requires Excel."))
                    }
                }
                if let ref = root.flatMap({ GridRef($0.attr("ref")) }),root?.attr("headerRowCount") != "0",change.at<=ref.top,change.at+change.count>ref.top {
                    throw failure(L("L'eliminazione dell'intestazione di una tabella richiede Excel.","Deleting a table's header row requires Excel."))
                }
            }
        }
        if change.deleting && !change.rows {
            for (_,path) in try linked(sheets[sheet].path,type:"table") {
                if let root = try part(path).rootElement(),let ref = GridRef(root.attr("ref")),try ref.shifted(change) == nil {
                    throw failure(L("Per eliminare tutte le colonne di una tabella, rimuovi prima la tabella in Excel.","To delete every column of a table, remove the table in Excel first."))
                }
            }
        }
    }
    func shiftFilter(_ filter: XMLElement, change: GridChange) throws {
        guard let old = GridRef(filter.attr("ref")) else { return }
        guard let next = try old.shifted(change) else { filter.detach(); return }
        if !change.rows {
            for column in filter.kids("filterColumn") {
                if let id = Int(column.attr("colId")),let shifted = change.point(old.left+id) { column.set("colId",String(shifted-next.left)) } else { column.detach() }
            }
        }
        filter.set("ref",next.text)
    }
    func shiftScope(_ e: XMLElement, attribute: String, sheet: Int, change: GridChange) throws {
        let scopes = e.attr(attribute).split(whereSeparator: { $0.isWhitespace }).compactMap { GridRef(String($0)) }
        var next: [GridRef] = [], sourceAnchor: (Int,Int)?
        for old in scopes {
            if let mapped = try old.shifted(change) {
                if next.isEmpty {
                    var r = old.top,c = old.left
                    if change.deleting { if change.rows && r>=change.at && r<change.at+change.count { r = change.at+change.count }; if !change.rows && c>=change.at && c<change.at+change.count { c = change.at+change.count } }
                    sourceAnchor = (r,c)
                }; next.append(mapped)
            }
        }
        guard !next.isEmpty else { e.detach(); return }
        if let original = scopes.first,let anchor = sourceAnchor {
            for f in elements(e) where ["formula","formula1","formula2"].contains(f.localName ?? "") {
                let rebased = fillFormula(f.stringValue ?? "",down:anchor.0-original.top,right:anchor.1-original.left)
                f.stringValue = try rewrite(rebased,context:sheet,target:sheet,change:change)
            }
        }
        e.set(attribute,next.map(\.text).joined(separator:" "))
    }
    func shiftGrid(_ sheet: Int, change: GridChange) throws {
        try preparePackage(); try structuralPreflight(sheet,change)
        let backup = try packageSnapshot()
        do {
            // Capture expanded shared expressions before moving cells. Preserve sharing only if
            // the shifted expressions still form one consistent template; otherwise de-share them.
            var formulas: [(Int,XMLElement,String,String?)] = []
            var shared: [String:[XMLElement]] = [:], sharedRanges: [String:String] = [:]
            for i in sheets.indices {
                for cell in try index(i).cells.values {
                    guard let f = cell.child("f") else { continue }
                    let si = f.attr("t") == "shared" ? f.attr("si") : nil
                    let text: String
                    if si != nil { guard let expanded = formulaText(cell,sheet:i) else { throw failure(L("Formula condivisa senza origine leggibile.","Shared formula has no readable master.")) }; text = expanded }
                    else { text = f.stringValue ?? "" }
                    formulas.append((i,cell,text,si))
                    if let si { let key = "\(i):\(si)"; shared[key,default:[]].append(cell); if !f.attr("ref").isEmpty { sharedRanges[key] = f.attr("ref") } }
                }
            }
            let doc = try sheetDocument(sheet), root = doc.rootElement()!, data = root.child("sheetData")!
            var inferredRow = 0
            for row in data.kids("row") {
                let r = Int(row.attr("r")) ?? row.kids("c").first.map { rowNumber($0.attr("r")) } ?? inferredRow+1
                inferredRow = r
                let nextRow = change.rows ? change.point(r) : r
                guard let rr = nextRow else { row.detach(); continue }
                guard rr<=1048576 else { throw failure(L("Ci sono celle all'ultima riga del foglio.","There are cells at the worksheet's last row.")) }
                row.set("r",String(rr)); var inferred = 0
                for c in row.kids("c") {
                    inferred = c.attr("r").isEmpty ? inferred+1 : column(c.attr("r"))
                    guard let cc = change.rows ? inferred : change.point(inferred) else { c.detach(); continue }
                    guard cc<=16384 else { throw failure(L("Ci sono celle all'ultima colonna del foglio.","There are cells at the worksheet's last column.")) }
                    c.set("r",columnName(cc)+String(rr))
                }
                let cols = row.kids("c").map { column($0.attr("r")) }
                if let first = cols.min(),let last = cols.max() { row.set("spans","\(first):\(last)") } else { row.removeAttribute(forName:"spans") }
            }
            var desired: [ObjectIdentifier:String] = [:]
            for (i,cell,text,si) in formulas where cell.parent?.parent != nil {
                guard let f = cell.child("f") else { continue }
                let shifted = try rewrite(text,context:i,target:sheet,change:change)
                f.stringValue = shifted; desired[ObjectIdentifier(cell)] = shifted
                if si != nil { f.removeAttribute(forName:"t"); f.removeAttribute(forName:"si"); f.removeAttribute(forName:"ref") }
                else if i == sheet,!f.attr("ref").isEmpty,let ref = GridRef(f.attr("ref")) { if let mapped = try ref.shifted(change) { f.set("ref",mapped.text) } }
            }
            for (key,original) in shared {
                let group = original.filter { desired[ObjectIdentifier($0)] != nil }.sorted { cellKey(rowNumber($0.attr("r")),column($0.attr("r")))<cellKey(rowNumber($1.attr("r")),column($1.attr("r"))) }
                guard let master = group.first,let expression = desired[ObjectIdentifier(master)] else { continue }
                let r = rowNumber(master.attr("r")),c = column(master.attr("r")),si = String(key.split(separator:":").last!)
                if group.allSatisfy({ desired[ObjectIdentifier($0)] == fillFormula(expression,down:rowNumber($0.attr("r"))-r,right:column($0.attr("r"))-c) }) {
                    for cell in group { let f = cell.child("f")!; f.set("t","shared"); f.set("si",si); if cell !== master { f.stringValue = "" } }
                    let context = Int(key.split(separator:":")[0])!
                    let oldRange = sharedRanges[key].flatMap(GridRef.init)
                    let mapped = context == sheet ? try oldRange?.shifted(change) : oldRange
                    master.child("f")!.set("ref",mapped?.text ?? master.attr("r")+":"+group.last!.attr("r"))
                }
            }
            if !change.rows,let cols = root.child("cols") {
                for col in cols.kids("col") {
                    guard let a = Int(col.attr("min")),let b = Int(col.attr("max")),let (lo,hi) = change.interval(a,b),lo<=16384 else { col.detach(); continue }
                    col.set("min",String(lo)); col.set("max",String(min(16384,hi)))
                }
            }
            let scopedNames = ["conditionalFormatting","dataValidation","protectedRange","ignoredError"]
            for e in elements(root) where scopedNames.contains(e.localName ?? "") && !e.attr("sqref").isEmpty { try shiftScope(e,attribute:"sqref",sheet:sheet,change:change) }
            for e in elements(root) {
                let name = e.localName ?? ""
                if name == "autoFilter" { try shiftFilter(e,change:change) }
                else if ["mergeCell","hyperlink","sortState","sortCondition"].contains(name),let ref = GridRef(e.attr("ref")) {
                    if let next = try ref.shifted(change),name != "mergeCell" || next.first.text != next.last.text { e.set("ref",next.text) } else { e.detach() }
                }
                if name == "hyperlink",!e.attr("location").isEmpty { e.set("location",try rewrite(e.attr("location"),context:sheet,target:sheet,change:change)) }
                for attr in ["activeCell","topLeftCell"] where !e.attr(attr).isEmpty {
                    if var ref = GridRef(e.attr(attr)) { if change.rows { ref.first.row = change.point(ref.top,clamp:true); ref.last.row = ref.first.row } else { ref.first.col = change.point(ref.left,clamp:true); ref.last.col = ref.first.col }; e.set(attr,ref.text) }
                }
                if name == "selection",let ref = GridRef(e.attr("sqref")) { e.set("sqref",try ref.shifted(change)?.text ?? e.attr("activeCell")) }
                if name == "pane" {
                    let key = change.rows ? "ySplit" : "xSplit"
                    if let n = Int(e.attr(key)),e.attr("state").hasPrefix("frozen") { let split = max(0,(change.point(n+1,clamp:true) ?? 1)-1); e.set(key,String(split)); if (Int(e.attr("xSplit")) ?? 0) == 0 && (Int(e.attr("ySplit")) ?? 0) == 0 { e.detach() } }
                }
            }
            for parent in root.kids("mergeCells") { parent.set("count",String(parent.kids("mergeCell").count)); if parent.kids("mergeCell").isEmpty { parent.detach() } }
            for parent in root.kids("dataValidations") { parent.set("count",String(parent.kids("dataValidation").count)); if parent.kids("dataValidation").isEmpty { parent.detach() } }
            for kind in ["rowBreaks","colBreaks"] {
                if let breaks = root.child(kind) {
                    for brk in breaks.kids("brk") where (kind == "rowBreaks") == change.rows {
                        if let n = Int(brk.attr("id")),let next = change.point(n) { brk.set("id",String(next)) } else { brk.detach() }
                    }; breaks.set("count",String(breaks.kids("brk").count)); breaks.set("manualBreakCount",String(breaks.kids("brk").count))
                }
            }
            for name in try part("xl/workbook.xml").nodes("/workbook/definedNames/definedName") {
                name.stringValue = try rewrite(name.stringValue ?? "",context:Int(name.attr("localSheetId")),target:sheet,change:change)
            }
            var rewrittenCharts = Set<String>()
            for i in sheets.indices {
                if i != sheet,let other = try sheetDocument(i).rootElement() { for f in elements(other) where ["formula","formula1","formula2"].contains(f.localName ?? "") { f.stringValue = try rewrite(f.stringValue ?? "",context:i,target:sheet,change:change) } }
                if i != sheet,let other = try sheetDocument(i).rootElement() {
                    for link in other.child("hyperlinks")?.kids("hyperlink") ?? [] where !link.attr("location").isEmpty { link.set("location",try rewrite(link.attr("location"),context:i,target:sheet,change:change)) }
                }
                if try members().contains(relationshipsPart(sheets[i].path)),let rels = try part(relationshipsPart(sheets[i].path)).rootElement() {
                    for link in rels.kids("Relationship") where link.attr("Type").hasSuffix("/hyperlink") && link.attr("Target").hasPrefix("#") {
                        link.set("Target","#"+(try rewrite(String(link.attr("Target").dropFirst()),context:i,target:sheet,change:change)))
                    }
                }
                for (_,path) in try linked(sheets[i].path,type:"table") {
                    guard let table = try part(path).rootElement(),let ref = GridRef(table.attr("ref")) else { continue }
                    if i == sheet {
                        if let next = try ref.shifted(change) { table.set("ref",next.text) }
                        if let filter = table.child("autoFilter") { try shiftFilter(filter,change:change) }
                        for e in elements(table) where ["sortState","sortCondition"].contains(e.localName ?? "") { if let ref = GridRef(e.attr("ref")) { if let next = try ref.shifted(change) { e.set("ref",next.text) } else { e.detach() } } }
                        if !change.rows,let columns = table.child("tableColumns") {
                            let old = columns.kids("tableColumn")
                            if change.deleting { for (j,col) in old.enumerated() where change.point(ref.left+j) == nil { col.detach() } }
                            else if change.at>ref.left && change.at<=max(ref.first.col!,ref.last.col!) {
                                var id = (old.compactMap { Int($0.attr("id")) }.max() ?? 0)+1
                                let names = Set(old.map { $0.attr("name") }); var added: [CellEdit] = []
                                for j in 0..<change.count {
                                    var title = L("Colonna\(id)","Column\(id)"); while names.contains(title) { id += 1; title = L("Colonna\(id)","Column\(id)") }
                                    let col = partElement("tableColumn",like:columns); col.set("id",String(id)); col.set("name",title); columns.insertChild(col,at:change.at-ref.left+j); id += 1
                                    if table.attr("headerRowCount") != "0" { added.append(CellEdit(ref:columnName(change.at+j)+String(ref.top),value:title,kind:.text)) }
                                }
                                invalidate(); if !added.isEmpty { try applyEdits(sheet,added) }
                            }
                            columns.set("count",String(columns.kids("tableColumn").count))
                        }
                    }
                    for formula in elements(table) where ["calculatedColumnFormula","totalsRowFormula"].contains(formula.localName ?? "") { formula.stringValue = try rewrite(formula.stringValue ?? "",context:i,target:sheet,change:change) }
                }
                for (_,drawing) in try linked(sheets[i].path,type:"drawing") {
                    if i == sheet,let dr = try part(drawing).rootElement() {
                        for anchor in dr.children?.compactMap({ $0 as? XMLElement }) ?? [] {
                            let key = change.rows ? "row" : "col"
                            if anchor.attr("editAs") == "absolute" { continue }
                            let from = anchor.child("from")?.child(key),to = anchor.child("to")?.child(key)
                            let before = Int(from?.stringValue ?? "") ?? 0
                            if let from { from.stringValue = String(max(0,(change.point(before+1,clamp:true) ?? 1)-1)) }
                            if let to,let n = Int(to.stringValue ?? "") {
                                let moved = anchor.attr("editAs") == "oneCell" ? n+(Int(from?.stringValue ?? "") ?? before)-before : max(0,(change.point(n+1,clamp:true) ?? 1)-1)
                                to.stringValue = String(max(Int(from?.stringValue ?? "") ?? 0,moved))
                            }
                        }
                    }
                    for (_,chart) in try linked(drawing,type:"chart") where rewrittenCharts.insert(chart).inserted {
                        if let cr = try part(chart).rootElement() { for f in elements(cr) where f.localName == "f" { f.stringValue = try rewrite(f.stringValue ?? "",context:i,target:sheet,change:change) } }
                    }
                }
            }
            for (path,doc) in documents where path.hasPrefix("xl/charts/") && !rewrittenCharts.contains(path) { if let cr = doc.rootElement() { for f in elements(cr) where f.localName == "f" { f.stringValue = try rewrite(f.stringValue ?? "",context:nil,target:sheet,change:change) } } }
            root.child("dimension")?.detach()
            let cells = data.kids("row").flatMap { $0.kids("c") }
            let dim = partElement("dimension",like:root)
            if !cells.isEmpty { let rows = cells.map { rowNumber($0.attr("r")) },cols = cells.map { column($0.attr("r")) }; dim.set("ref",columnName(cols.min()!)+String(rows.min()!)+":"+columnName(cols.max()!)+String(rows.max()!)) } else { dim.set("ref","A1") }
            let at = (root.children ?? []).firstIndex { !["sheetPr"].contains($0.localName ?? "") && $0 is XMLElement } ?? 0; root.insertChild(dim,at:at)
            dirty = true; invalidate()
        } catch { try restorePackage(backup); throw error }
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
        var chartOwners: [String:Int] = [:]
        for i in sheets.indices where try sheetDocument(i).rootElement()?.child("drawing") != nil { for (_,drawing) in try linked(sheets[i].path,type:"drawing") { for (_,chart) in try linked(drawing,type:"chart") { chartOwners[chart] = i; _ = try part(chart) } } }
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
            for c in doc.rootElement()?.child("sheetData")?.kids("row").flatMap({ $0.kids("c") }) ?? [] where c.child("f") != nil { if (try? computed(c,sheet:i)) != nil { c.kids("v").forEach { $0.detach() } } }
            try write(sheets[i].path, doc)
        }
        for (path,document) in documents where !sheets.contains(where:{ $0.path == path }) && !["xl/workbook.xml","xl/_rels/workbook.xml.rels","[Content_Types].xml"].contains(path) && !removedMembers.contains(path) {
            if document.rootElement()?.localName == "chartSpace" { let copy = try XMLDocument(data:document.xmlData,options:[.nodePreserveWhitespace]); try refreshChartCaches(copy,sheet:chartOwners[path] ?? 0); try write(path,copy) }
            else { try write(path,document) }
        }
        let workbook = try XMLDocument(data:part("xl/workbook.xml").xmlData,options:[.nodePreserveWhitespace])
        let calc = (workbook.nodes("/workbook/calcPr").first as? XMLElement) ?? XMLElement(name: "calcPr")
        if calc.parent == nil { workbook.rootElement()?.addChild(calc) }
        for (key, value) in [("calcMode", "auto"), ("fullCalcOnLoad", "1"), ("forceFullCalc", "1")] { calc.removeAttribute(forName: key); calc.addAttribute(XMLNode.attribute(withName: key, stringValue: value) as! XMLNode) }
        try write("xl/workbook.xml", workbook)
        let rels = try XMLDocument(data:part("xl/_rels/workbook.xml.rels").xmlData,options:[.nodePreserveWhitespace])
        for r in rels.nodes("/Relationships/Relationship") where r.attr("Type").hasSuffix("/calcChain") { r.detach() }
        if documents["xl/styles.xml"] != nil, !rels.nodes("/Relationships/Relationship").contains(where:{ $0.attr("Type").hasSuffix("/styles") }) {
            let node = XMLElement(name:"Relationship"), used = Set(rels.nodes("/Relationships/Relationship").map { $0.attr("Id") }); var n = 1; while used.contains("rId\(n)") { n += 1 }
            node.set("Id","rId\(n)"); node.set("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"); node.set("Target","styles.xml"); rels.rootElement()?.addChild(node)
        }
        try write("xl/_rels/workbook.xml.rels", rels)
        let types = try XMLDocument(data:part("[Content_Types].xml").xmlData,options:[.nodePreserveWhitespace])
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
        try zip(["-q", "-d", archive.path, "xl/calcChain.xml"]+removedMembers.sorted(), acceptMissing: true)
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
        let geometry = GridGeometry(root)
        var maxRow = sheetIndex.maxRow, maxCol = sheetIndex.maxCol
        maxCol = min(max(maxCol, maxRow <= 1000 ? 26 : 1), 1000); maxRow = max(maxRow, 100)
        var widths = [Int: Double](); var hiddenCols = Set<Int>()
        for c in root.child("cols")?.kids("col") ?? [] {
            let lo = Int(c.attr("min")) ?? 1, hi = min(maxCol, Int(c.attr("max")) ?? 1)
            guard lo <= hi else { continue }
            for i in lo...hi {
                widths[i] = geometry.width(i)
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
        var strings = ["search": L("Cerca valore…", "Search values…"), "all": L("Tutti i valori", "All values"), "empty": L("(Vuote)", "(Blanks)"), "more": L("Altri valori: usa la ricerca", "More values: use search"), "asc": L("Ordina A → Z","Sort A → Z"), "desc": L("Ordina Z → A","Sort Z → A"), "sum": L("Somma","Sum"), "average": L("Media","Average"), "count": L("Conteggio","Count"), "edit": L("Modifica cella","Edit cell")]
        for rows in [true,false] { for after in [false,true] {
            strings[(rows ? "row" : "col")+(after ? "After" : "Before")] = gridActionTitle(rows:rows,count:0,deleting:false,after:after)
        }; strings[rows ? "rowDelete" : "colDelete"] = gridActionTitle(rows:rows,count:0,deleting:true)
            for position in ["Before","After","Delete"] { strings[(rows ? "row" : "col")+position+"One"] = gridActionTitle(rows:rows,count:1,deleting:position == "Delete",after:position == "After") }
        }
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
        body{padding-bottom:24px}#formula-bar{position:sticky;top:0;left:0;z-index:30;height:42px;background:#f8f9fa;display:flex;align-items:center;gap:8px;padding:5px;width:100vw;border-bottom:1px solid #ddd}#name{width:110px}#formula{flex:1;min-width:80px;height:29px;resize:none;white-space:pre}#formula-bar :is(input,textarea){font:13px -apple-system;padding:5px;border:1px solid #ddd;border-radius:3px}#message{color:#b3261e;font-size:12px;max-width:35vw}#editor{resize:none;position:absolute;z-index:40;border:2px solid #2979e8;background:white;color:#202124;font:13px -apple-system;padding:0 6px;outline:none}#selection-border{position:absolute;pointer-events:none;z-index:12;border:2px solid #1a73e8}#fill-handle{position:absolute;width:8px;height:8px;right:-5px;bottom:-5px;border:1px solid white;background:#1a73e8;cursor:crosshair;pointer-events:auto}tr[style]>td,tr[style]>th{height:inherit}th.selected-header{background:#d3e3fd!important;color:#174ea6;font-weight:700}body.paint td{cursor:crosshair}#stats{position:fixed;bottom:0;right:0;z-index:35;background:#f8f9fa;padding:4px 12px;font-size:12px}table{user-select:none}#formula-bar :is(input,textarea){user-select:text}
        .chart-box{position:absolute;z-index:15;background:white;box-shadow:0 1px 5px #0003;user-select:none;touch-action:none}.chart-box.selected-chart{outline:2px solid #1a73e8}.chart-box svg{display:block;pointer-events:none}.chart-resize{display:none;position:absolute;width:12px;height:12px;right:-6px;bottom:-6px;background:#1a73e8;border:2px solid white;cursor:nwse-resize}.selected-chart .chart-resize{display:block}#grid-menu{position:fixed;z-index:100;background:white;border:1px solid #ddd;border-radius:6px;padding:5px;box-shadow:0 4px 18px #0003}#grid-menu button{display:block;width:100%;border:0;background:white;padding:8px;text-align:left;font:13px -apple-system}#grid-menu button:hover{background:#e6efff}
        @media print{.chart-resize,#grid-menu{display:none!important}.chart-box{box-shadow:none;outline:none!important}#formula-bar,#stats,#editor,#message,#selection-border{display:none!important}thead th,tbody th,td{position:static!important}td.sel{outline:none;box-shadow:none!important}th.selected-header{background:#f3f4f6!important}tr:hover td{box-shadow:none}td.fh::after{content:''}}
        </style></head><body><div id="formula-bar"><input id="name" aria-label="\(L("Cella o intervallo","Cell or range"))" placeholder="A1"><span style="font:italic 16px Georgia;color:#747775">fx</span><textarea id="formula" rows="1" wrap="off" aria-label="\(L("Barra della formula","Formula bar"))"></textarea><span id="message" role="alert"></span></div><div id="stats"></div><table><colgroup><col style="width:40px">
        """
        html.reserveCapacity(rows.count * columns.count * 48 + 4096)
        for c in columns { html += "<col data-c='\(c)' style='width:\(chartNumber(widths[c] ?? 110))px'>" }
        html += "</colgroup><thead><tr><th style='left:0;z-index:9'></th>"
        for c in columns { html += "<th data-c='\(c)'\(c <= frozenCols ? " style='left:\(Int(offsets[c] ?? 40))px;z-index:8'" : "")>\(columnName(c))</th>" }
        html += "</tr></thead><tbody>"
        for r in rows {
            let isFrozen = r <= frozenRows, top = r <= frozenRows ? Int(geometry.point(col:0,row:r-1).y) : 0
            let rowHeight = geometry.heights[r].map { " style='height:\(chartNumber($0))px'" } ?? ""
            html += "<tr\(rowHeight) data-row='\(r)'\(r == frozenRows && frozenRows > 0 ? " class='heading'" : "")><th\(isFrozen ? " style='top:\(top)px;z-index:8'" : "")>\(r)</th>"
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
        </tbody></table><div id="charts">\(try chartHTML(index))</div><script>
        const T=\(stringsJSON);let FR=\(filterRow),FL=\(filterLastRow),filterC1=\(column(filterStart)),filterC2=\(column(filterEnd));let filters={},current=null,anchor=null,selected=new Set(),dragging=false,editing=null,pendingMove=null,busy=false;
        const HR=new Set(\(hiddenRows.sorted())),HC=new Set(\(hiddenCols.sorted()));
        const full=td=>td.dataset.full??td.textContent;
        function colName(c){let s='';while(c>0){c--;s=String.fromCharCode(65+c%26)+s;c=Math.floor(c/26)}return s}
        const refOf=td=>colName(+td.dataset.c)+td.parentElement.dataset.row;
        const mutationActions=new Set(['commit','clear','cut','paste','fillDown','autofill','undo','redo','sort','gridChange','chartFrame','chartDelete']);
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
        let headerDrag=null;let selectedHeaders=[],filling=null,fillTarget=null,paintArmed=false,pendingFill=null;
        function selectionFurniture(){if(!current)return;for(const th of selectedHeaders)th.classList.remove('selected-header');selectedHeaders=[];const b=bounds();for(const th of document.querySelectorAll('thead th[data-c]'))if(+th.dataset.c>=b.c1&&+th.dataset.c<=b.c2)selectedHeaders.push(th);for(let r=b.r1;r<=b.r2;r++){const th=rowMap.get(r)?.querySelector('th');if(th)selectedHeaders.push(th)}for(const th of selectedHeaders)th.classList.add('selected-header');const a=at(b.r1,b.c1),z=at(b.r2,b.c2);if(!a||!z){selectionBorder.hidden=true;return}const x=a.getBoundingClientRect(),y=z.getBoundingClientRect();selectionBorder.hidden=false;Object.assign(selectionBorder.style,{left:x.left+scrollX+'px',top:x.top+scrollY+'px',width:y.right-x.left+'px',height:y.bottom-x.top+'px'})}
        window.armPaint=on=>{paintArmed=on;document.body.classList.toggle('paint',on)};
        function finishPaint(){if(!paintArmed)return;paintArmed=false;document.body.classList.remove('paint');post({action:'paint'})}
        handle.addEventListener('mousedown',e=>{if(busy)return;e.preventDefault();e.stopPropagation();filling=bounds();fillTarget=filling});
        function notifySelection(){selectionFurniture();const b=bounds();post({action:'select',ref:refOf(current),text:full(current),...Object.fromEntries(Object.entries(b).map(([k,v])=>[k,String(v)]))});nameBox.value=selected.size>1?colName(b.c1)+b.r1+':'+colName(b.c2)+b.r2:refOf(current);bar.value=raw(current);updateStats()}
        function select(td,scroll=false,extend=false){if(!td)return;deselectChart();if(editing)cancelEdit();if(!extend||!anchor)anchor=td;current?.classList.remove('active');current=td;current.classList.add('active');const b=bounds(),next=new Set();for(let r=b.r1;r<=b.r2;r++){const row=rowMap.get(r);if(!row||row.hidden)continue;for(let c=b.c1;c<=b.c2;c++){const x=cellAt(row,c);if(x)next.add(x)}}for(const x of selected)if(!next.has(x))x.classList.remove('sel');for(const x of next)if(!selected.has(x))x.classList.add('sel');selected=next;const s=getSelection(),range=document.createRange();range.selectNodeContents(td);s.removeAllRanges();s.addRange(range);if(scroll)td.scrollIntoView({block:'nearest',inline:'nearest'});notifySelection()}
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
        document.addEventListener('mousedown',e=>{if(e.button!==0||e.target.closest('input,textarea,#filter-menu,#grid-menu,.chart-box'))return;const td=e.target.closest('td');if(td){e.preventDefault();if(editing){commit(editing.value,null);return}select(td,false,e.shiftKey);dragging=true}else if(e.target.tagName==='TH'){e.preventDefault();const th=e.target;headerDrag=headerInfo(th);if(headerDrag){if(e.shiftKey&&current){const b=bounds();headerDrag.start=headerDrag.rows?b.r1:b.c1}else headerDrag.start=headerDrag.n;selectHeaders(headerDrag,headerDrag.n)}}});
        document.addEventListener('mouseover',e=>{if(headerDrag){const h=headerInfo(e.target.closest?.('th'));if(h&&h.rows===headerDrag.rows)selectHeaders(headerDrag,h.n);return}if(filling){const td=e.target.closest?.('td');if(td){const r=+td.parentElement.dataset.row,c=+td.dataset.c;fillTarget=r>filling.r2&&r-filling.r2>=c-filling.c2?{...filling,r2:r}:{...filling,c2:Math.max(c,filling.c2)};const z=at(fillTarget.r2,fillTarget.c2);if(z){const a=at(filling.r1,filling.c1).getBoundingClientRect(),b=z.getBoundingClientRect();selectionBorder.style.width=b.right-a.left+'px';selectionBorder.style.height=b.bottom-a.top+'px'}}return}const tip=e.target.closest?.('td');if(tip&&!tip.title&&full(tip).length>28)tip.title=full(tip);if(dragging){const td=e.target.closest('td');if(td&&td!==current)select(td,false,true)}});document.addEventListener('mouseup',()=>{dragging=false;headerDrag=null;if(filling){const source=filling,target=fillTarget;filling=null;fillTarget=null;pendingFill=target;post({action:'autofill',source:JSON.stringify(source),target:JSON.stringify(target)});return}finishPaint()});document.addEventListener('keyup',e=>{if(e.key.startsWith('Arrow')||e.key==='Tab')finishPaint()});
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
        let selectedChart=null,chartDrag=null;
        function deselectChart(){document.querySelector('.selected-chart')?.classList.remove('selected-chart');selectedChart=null}
        window.setCharts=html=>{const id=selectedChart;deselectChart();document.getElementById('charts').innerHTML=html;const box=[...document.querySelectorAll('.chart-box')].find(b=>b.dataset.chart===id);if(box){selectedChart=id;box.classList.add('selected-chart');box.focus()}};
        function headerInfo(th){if(!th)return null;if(th.closest('thead')&&th.dataset.c)return {rows:false,n:+th.dataset.c};if(th.parentElement.dataset.row)return {rows:true,n:+th.parentElement.dataset.row};return null}
        function selectHeaders(h,n){const lo=Math.min(h.start??h.n,n),hi=Math.max(h.start??h.n,n);if(h.rows){select(rowMap.get(lo)?.querySelector('td'));select(rowMap.get(hi)?.querySelector('td:last-child'),false,true)}else{select(at(+rows[0].dataset.row,lo));select(at(+rows.at(-1).dataset.row,hi),false,true)}}
        document.addEventListener('contextmenu',e=>{const h=headerInfo(e.target.closest('th'));if(!h)return;e.preventDefault();if(busy||editing)return;const b=current?bounds():null;if(!b||h.n<(h.rows?b.r1:b.c1)||h.n>(h.rows?b.r2:b.c2))selectHeaders(h,h.n);document.getElementById('grid-menu')?.remove();const count=h.rows?bounds().r2-bounds().r1+1:bounds().c2-bounds().c1+1,menu=document.createElement('div');menu.id='grid-menu';for(const [suffix,position] of [['Before','before'],['After','after'],['Delete','delete']]){const button=document.createElement('button');button.textContent=T[(h.rows?'row':'col')+suffix].replaceAll('%N',String(count));if(count===1)button.textContent=T[(h.rows?'row':'col')+suffix+'One'];button.onclick=()=>{menu.remove();post({action:'gridChange',rows:h.rows?'1':'0',position})};menu.append(button)}document.body.append(menu);menu.style.left=Math.max(0,Math.min(e.clientX,innerWidth-menu.offsetWidth))+'px';menu.style.top=Math.max(0,Math.min(e.clientY,innerHeight-menu.offsetHeight))+'px'});
        document.addEventListener('mousedown',e=>{if(!e.target.closest('#grid-menu'))document.getElementById('grid-menu')?.remove();const box=e.target.closest('.chart-box');if(!box){if(!e.target.closest('#formula-bar'))deselectChart();return}if(busy||editing||e.button!==0)return;e.preventDefault();deselectChart();selectedChart=box.dataset.chart;box.classList.add('selected-chart');box.focus();chartDrag={box,x:e.clientX,y:e.clientY,left:parseFloat(box.style.left),top:parseFloat(box.style.top),width:box.offsetWidth,height:box.offsetHeight,resize:!!e.target.closest('.chart-resize')}});
        document.addEventListener('mousemove',e=>{if(!chartDrag)return;const d=chartDrag,dx=e.clientX-d.x,dy=e.clientY-d.y;if(d.resize){d.box.style.width=Math.max(120,d.width+dx)+'px';d.box.style.height=Math.max(100,d.height+dy)+'px'}else{d.box.style.left=Math.max(40,d.left+dx)+'px';d.box.style.top=Math.max(66,d.top+dy)+'px'}});
        document.addEventListener('mouseup',()=>{if(!chartDrag)return;const d=chartDrag;chartDrag=null;const frame={x:parseFloat(d.box.style.left),y:parseFloat(d.box.style.top),width:d.box.offsetWidth,height:d.box.offsetHeight};if(frame.x!==d.left||frame.y!==d.top||frame.width!==d.width||frame.height!==d.height)post({action:'chartFrame',id:d.box.dataset.chart,...Object.fromEntries(Object.entries(frame).map(([k,v])=>[k,String(v)]))})});
        document.addEventListener('keydown',e=>{if(!selectedChart||e.target.matches?.('input,textarea'))return;if(e.key==='Delete'||e.key==='Backspace'){e.preventDefault();e.stopImmediatePropagation();post({action:'chartDelete',id:selectedChart})}else if(e.key==='Escape'){e.preventDefault();deselectChart()}},true);
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
            if clean.range(of:"^[1-9][0-9]{0,6}$",options:.regularExpression) != nil { return .ref(sheet,Int(clean)!,0) }
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
                i += 2; if String(chars[i...]).hasPrefix("#REF!") { tokens.append(.error("#REF!")); i += 5; continue }; var word = ""
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
                    i += 1; if String(chars[i...]).hasPrefix("#REF!") { tokens.append(.error("#REF!")); i += 5; continue }; var ref = ""
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
        case .num(let n):
            if peek == .colon { guard n.isFinite,n == n.rounded(),n>=1,n<=1048576 else { throw FormulaInvalid() }; return try wholeRows(sheet,Int(n)) }
            return FValue.number(n).excelValue
        case .str(let s): return .text(s)
        case .error(let e): return .error(e)
        case .lp: let v = try comparison(); guard take(.rp) else { throw FormulaInvalid() }; return v
        case .ref(let s, let r, let c):
            if c == 0 { return try wholeRows(s,r) }
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
    func wholeRows(_ s: Int, _ first: Int) throws -> FValue {
        guard take(.colon),let token = peek else { throw FormulaInvalid() }
        let last: Int
        switch token { case .num(let n): guard n == n.rounded(),n>0,n<=1048576 else { throw FormulaInvalid() }; last = Int(n)
        case .ref(_,let r,0): last = r
        default: throw FormulaInvalid() }
        pos += 1; if validating { return .empty }
        guard first>0,first<=1048576 else { return .error("#REF!") }
        let width = max(1,try book.index(s).maxCol), lo = min(first,last),hi = max(first,last)
        return .range(try book.rangeValues(s,lo,1,hi,width),hi-lo+1,width)
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
        if item.action == #selector(gridMenu(_:)),let command = item.representedObject as? String {
            let parts = command.split(separator:":"),rows = parts[0] == "row",count = rows ? selectedRange.r2-selectedRange.r1+1 : selectedRange.c2-selectedRange.c1+1
            item.title = gridActionTitle(rows:rows,count:count,deleting:parts[1] == "delete",after:parts[1] == "after")
            return book != nil && !loading
        }
        if item.action == #selector(createChartMenu(_:)) { return book != nil && !loading && selectedRange.count>1 }
        switch item.action {
        case #selector(undoEdit): return NSApp.modalWindow != nil || editUndo.canUndo
        case #selector(redoEdit): return NSApp.modalWindow != nil || editUndo.canRedo
        case #selector(saveFile), #selector(saveCopy), #selector(printSheet), #selector(clearFilters), #selector(nextSheet), #selector(previousSheet): return book != nil && !loading
        case #selector(editCell): return book != nil && !selectedRef.isEmpty && !loading
        default: return true
        }
    }
    // Package actions share one undo record, including all dependent sheet/chart/table parts.
    func packageAction(_ name: String, _ action: (Book,Int) throws -> Void, reload: Bool = false) {
        guard let book,!loading else { return }
        do {
            try book.preparePackage()
            let old = try book.packageSnapshot(),generation = savedGeneration
            do { try action(book,picker.indexOfSelectedItem) } catch { try book.restorePackage(old); throw error }
            committedActions += 1
            editUndo.record(self,name:name) { $0.restoreWorkbook(old,generation:generation,reload:reload) }
            if reload { reloadGrid() } else { refreshCharts() }
            updateEditState(); lastEditError = ""; web.evaluateJavaScript("editResult(true,'')",completionHandler:nil)
        } catch { editError(error) }
    }
    func restoreWorkbook(_ state: PackageSnapshot, generation: Int, reload: Bool) {
        guard let book else { return }
        do {
            let inverse = try book.packageSnapshot(),inverseGeneration = savedGeneration
            try book.restorePackage(state); book.dirty = state.dirty || generation != savedGeneration
            editUndo.registerUndo(withTarget:self) { $0.restoreWorkbook(inverse,generation:inverseGeneration,reload:reload) }
            if reload { reloadGrid() } else { refreshCharts() }; updateEditState()
        } catch { editError(error) }
    }
    var pendingGridSelection: CellRange?
    func reloadGrid() {
        guard let book else { return }
        do { pendingGridSelection = selectedRange; show(try book.render(picker.indexOfSelectedItem)) } catch { editError(error) }
    }
    func refreshCharts() {
        guard let book,let html = try? book.chartHTML(picker.indexOfSelectedItem) else { return }
        web.evaluateJavaScript("window.setCharts&&setCharts(\(jsonString(html)))",completionHandler:nil)
    }
    func changeGrid(rows: Bool, position: String) {
        let lo = rows ? selectedRange.r1 : selectedRange.c1,hi = rows ? selectedRange.r2 : selectedRange.c2
        let deleting = position == "delete",change = GridChange(rows:rows,at:position == "after" ? hi+1 : lo,count:hi-lo+1,deleting:deleting)
        packageAction(gridActionTitle(rows:rows,count:change.count,deleting:deleting,after:position == "after"),{ book,i in try book.shiftGrid(i,change:change) },reload:true)
    }
    @objc func gridMenu(_ sender: NSMenuItem) {
        guard let command = sender.representedObject as? String else { return }; let parts = command.split(separator:":")
        changeGrid(rows:parts[0] == "row",position:String(parts[1]))
    }
    @objc func createChartMenu(_ sender: NSMenuItem) {
        guard let kind = sender.representedObject as? String else { return }
        packageAction(L("Inserisci grafico","Insert Chart"),{ book,i in _ = try book.createChart(i,range:self.selectedRange,kind:kind) })
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
        web.evaluateJavaScript("setCells(\(json))",completionHandler:nil); refreshCharts()
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
        book.url = target; book.originalMembers = nil; book.untitled = false; book.dirty = false; savedGeneration += 1
        for i in book.sheets.indices { if let doc = book.documents[book.sheets[i].path] { for c in doc.rootElement()?.child("sheetData")?.kids("row").flatMap({ $0.kids("c") }) ?? [] where c.child("f") != nil { if (try? book.computed(c,sheet:i)) != nil { c.kids("v").forEach { $0.detach() } } } } }; book.invalidate()
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
        case "gridChange": changeGrid(rows:data["rows"] == "1",position:data["position"] ?? "before")
        case "chartFrame":
            guard let id = data["id"],let x = Double(data["x"] ?? ""),let y = Double(data["y"] ?? ""),let w = Double(data["width"] ?? ""),let h = Double(data["height"] ?? "") else { return }
            packageAction(L("Sposta/ridimensiona grafico","Move/Resize Chart"),{ book,i in try book.setChartFrame(i,id:id,frame:CGRect(x:x,y:y,width:w,height:h)) })
        case "chartDelete":
            guard let id = data["id"] else { return }
            packageAction(L("Elimina grafico","Delete Chart"),{ book,i in try book.deleteChart(i,id:id) })
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
        func gridItems(_ menu: NSMenu, includeDelete: Bool) {
            for rows in [true,false] { for position in (includeDelete ? ["before","after","delete"] : ["before","after"]) {
                let item = menu.addItem(withTitle:gridActionTitle(rows:rows,count:1,deleting:position == "delete",after:position == "after"),action:#selector(gridMenu(_:)),keyEquivalent:""); item.target = self; item.representedObject = (rows ? "row:" : "col:")+position
            } }
        }
        edit.addItem(.separator()); gridItems(edit,includeDelete:true)
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
        insert.addItem(.separator()); gridItems(insert,includeDelete:false)
        let chartItem = NSMenuItem(title:L("Grafico","Chart"),action:nil,keyEquivalent:""),chartMenu = NSMenu(title:L("Grafico","Chart")); chartItem.submenu = chartMenu; insert.addItem(chartItem)
        for (title,kind) in [(L("Colonne","Column"),"column"),(L("Barre","Bar"),"bar"),(L("Linee","Line"),"line"),(L("Torta","Pie"),"pie")] { let item = chartMenu.addItem(withTitle:title,action:#selector(createChartMenu(_:)),keyEquivalent:""); item.target = self; item.representedObject = kind }
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
    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) { webView.pageZoom = zoom;
        if let range = pendingGridSelection { pendingGridSelection = nil; web.evaluateJavaScript("select(ensure(\(range.r1),\(range.c1)));select(ensure(\(range.r2),\(range.c2)),true,true)",completionHandler:nil) }
 if paintStyle != nil { web.evaluateJavaScript("armPaint(true)",completionHandler:nil) } }
    func show(_ error: Error) { NSAlert(error: error).runModal() }
    func webView(_ webView: WKWebView, decidePolicyFor action: WKNavigationAction, decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        if let url = action.request.url, url.isFileURL, supportedExtensions.contains(url.pathExtension.lowercased()) { decisionHandler(.cancel); load(url); return }
        decisionHandler(action.navigationType == .other ? .allow : .cancel)
    }
}

#if DEBUG
if CommandLine.arguments.count > 2 && CommandLine.arguments[1] == "--model-test-v13" {
    do { try runV13ModelTests(URL(fileURLWithPath:CommandLine.arguments[2])) } catch { fputs("FAIL: \(error)\n",stderr); exit(1) }; exit(0)
}
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
