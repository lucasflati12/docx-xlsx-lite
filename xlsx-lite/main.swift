import AppKit
import WebKit
import UniformTypeIdentifiers
import Foundation

func esc(_ s: String) -> String { s.replacingOccurrences(of: "&", with: "&amp;").replacingOccurrences(of: "<", with: "&lt;").replacingOccurrences(of: ">", with: "&gt;").replacingOccurrences(of: "\"", with: "&quot;") }
extension XMLNode {
    func nodes(_ path: String) -> [XMLNode] { (try? nodes(forXPath: path)) ?? [] }
    func attr(_ name: String) -> String { (self as? XMLElement)?.attribute(forName: name)?.stringValue ?? "" }
    func val(_ path: String) -> String { nodes(path).first?.stringValue ?? "" }
}
func column(_ ref: String) -> Int { ref.prefix(while: { $0.isLetter }).uppercased().utf8.reduce(0) { $0 * 26 + Int($1) - 64 } }
struct Sheet { let name: String; let path: String }
struct CellStyle { var css = ""; var format = "" }
final class Book {
    let url: URL
    var documents: [String:XMLDocument] = [:]
    var dirty = false
    var strings: [String] = []; var sheets: [Sheet] = []; var styles: [CellStyle] = []; var date1904 = false
    init(_ url: URL) throws {
        self.url = url
        let workbook = try xml("xl/workbook.xml")
        date1904 = workbook.val("/workbook/workbookPr/@date1904") == "1"
        let rels = try xml("xl/_rels/workbook.xml.rels")
        for sheet in workbook.nodes("/workbook/sheets/sheet") {
            let id = sheet.attr("r:id")
            if let rel = rels.nodes("/Relationships/Relationship").first(where: { $0.attr("Id") == id }) {
                let target = rel.attr("Target")
                sheets.append(Sheet(name: sheet.attr("name"), path: target.hasPrefix("/") ? String(target.dropFirst()) : "xl/" + target))
            }
        }
        if let shared = try? xml("xl/sharedStrings.xml") { strings = shared.nodes("/sst/si").map { $0.nodes(".//t").map { $0.stringValue ?? "" }.joined() } }
        if let doc = try? xml("xl/styles.xml") {
            let fonts = doc.nodes("/styleSheet/fonts/font"), fills = doc.nodes("/styleSheet/fills/fill")
            var formats = [14:"dd/mm/yyyy",15:"dd/mm/yyyy",16:"dd/mm",17:"mm/yyyy",18:"hh:mm",19:"hh:mm:ss",20:"hh:mm",21:"hh:mm:ss",22:"dd/mm/yyyy hh:mm",9:"0%",10:"0.00%",1:"0",2:"0.00",3:"#,##0",4:"#,##0.00"]
            for f in doc.nodes("/styleSheet/numFmts/numFmt") { formats[Int(f.attr("numFmtId")) ?? -1] = f.attr("formatCode") }
            for x in doc.nodes("/styleSheet/cellXfs/xf") {
                var s = CellStyle(); s.format = formats[Int(x.attr("numFmtId")) ?? 0] ?? ""
                if let i = Int(x.attr("fontId")), fonts.indices.contains(i) {
                    let f = fonts[i]
                    if !f.nodes("b").isEmpty { s.css += "font-weight:600;" }
                    if !f.nodes("i").isEmpty { s.css += "font-style:italic;" }
                    if let c = color(f.nodes("color").first) { s.css += "color:\(c);" }
                }
                if let i = Int(x.attr("fillId")), fills.indices.contains(i), fills[i].val("patternFill/@patternType") == "solid", let c = color(fills[i].nodes("patternFill/fgColor").first) { s.css += "background:\(c);" }
                let align = x.val("alignment/@horizontal")
                if ["left","right","center"].contains(align) { s.css += "text-align:\(align);" }
                if x.val("alignment/@wrapText") == "1" { s.css += "white-space:nowrap;" }
                styles.append(s)
            }
        }
        if sheets.isEmpty { throw NSError(domain:"Lettore", code:1, userInfo:[NSLocalizedDescriptionKey:"Il file non contiene fogli leggibili."]) }
    }
    func color(_ node: XMLNode?) -> String? {
        guard let node else { return nil }
        let rgb = node.attr("rgb"); if rgb.count >= 6 { return "#" + rgb.suffix(6) }
        let colors = ["#ffffff","#000000","#e7e6e6","#44546a","#4472c4","#ed7d31","#a5a5a5","#ffc000","#5b9bd5","#70ad47"]
        if let i = Int(node.attr("theme")), colors.indices.contains(i) { return colors[i] }
        return nil
    }
    func xml(_ member: String) throws -> XMLDocument {
        let process = Process(); process.executableURL = URL(fileURLWithPath:"/usr/bin/unzip")
        let escapedMember=member.replacingOccurrences(of:"[",with:"\\[").replacingOccurrences(of:"*",with:"\\*").replacingOccurrences(of:"?",with:"\\?")
        process.arguments = ["-p", url.path, escapedMember]
        let pipe = Pipe(); process.standardOutput = pipe; process.standardError = FileHandle.nullDevice
        try process.run(); let data = pipe.fileHandleForReading.readDataToEndOfFile(); process.waitUntilExit()
        guard process.terminationStatus == 0 else { throw NSError(domain:"Lettore",code:2,userInfo:[NSLocalizedDescriptionKey:"File Excel non valido, protetto o non leggibile."]) }
        return try XMLDocument(data:data, options: [.nodePreserveWhitespace])
    }
    func sheetDocument(_ index: Int) throws -> XMLDocument {
        let path=sheets[index].path
        if let doc=documents[path] { return doc }
        let doc=try xml(path); documents[path]=doc; return doc
    }
    func display(_ cell: XMLNode) -> String {
        let v = cell.val("v"), type = cell.attr("t")
        if !cell.nodes("f").isEmpty && v.isEmpty { return "↻ Ricalcolo in Excel" }
        if type == "s" { return Int(v).flatMap { strings.indices.contains($0) ? strings[$0] : nil } ?? v }
        if type == "inlineStr" { return cell.nodes("is//t").map { $0.stringValue ?? "" }.joined() }
        if type == "b" { return v == "1" ? "VERO" : "FALSO" }
        guard type != "str", let n = Double(v), n.isFinite else { return v }
        let idx = Int(cell.attr("s")) ?? 0
        let format = styles.indices.contains(idx) ? styles[idx].format : ""
        let clean = format.replacingOccurrences(of:"\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.",with:"",options:.regularExpression).lowercased()
        if clean.contains("yy") || clean.contains("dd") || clean.contains("hh") {
            let formatter = DateFormatter(); formatter.locale = Locale(identifier:"it_IT"); formatter.timeZone = TimeZone(secondsFromGMT:0)
            formatter.dateFormat = clean.contains("yy") || clean.contains("dd") ? (clean.contains("hh") ? "dd/MM/yyyy HH:mm" : "dd/MM/yyyy") : "HH:mm"
            return formatter.string(from: Date(timeIntervalSince1970:(n - (date1904 ? 24107 : 25569)) * 86400))
        }
        if format.isEmpty || format == "General" { return v }
        let f = NumberFormatter(); f.locale = Locale(identifier:"it_IT"); f.numberStyle = clean.contains("%") ? .percent : .decimal
        let fraction = clean.split(separator:".").dropFirst().first?.prefix(while: { $0 == "0" || $0 == "#" }).count ?? 0
        f.minimumFractionDigits = fraction; f.maximumFractionDigits = fraction; f.usesGroupingSeparator = clean.contains(",")
        return (f.string(from:NSNumber(value:n)) ?? v) + (format.contains("€") ? " €" : "")
    }
    func cell(_ index: Int,_ ref: String) throws -> XMLElement? {
        guard ref.range(of:"^[A-Z]{1,3}[1-9][0-9]{0,6}$",options:.regularExpression) != nil else { return nil }
        return try sheetDocument(index).nodes("/worksheet/sheetData/row/c[@r='\(ref)']").first as? XMLElement
    }
    func setCell(_ index: Int,_ ref: String,_ value: String, numeric: Bool) throws {
        guard ref.range(of:"^[A-Z]{1,3}[1-9][0-9]{0,6}$",options:.regularExpression) != nil else { return }
        let doc=try sheetDocument(index)
        guard doc.nodes("/worksheet/sheetProtection").isEmpty else { throw NSError(domain:"Editor",code:1,userInfo:[NSLocalizedDescriptionKey:"Questo foglio è protetto. Rimuovi la protezione in Excel prima di modificarlo."]) }
        guard let data=doc.nodes("/worksheet/sheetData").first as? XMLElement else { return }
        let rowNumber=Int(ref.drop(while:{$0.isLetter})) ?? 1
        let row: XMLElement
        if let existing=data.nodes("row[@r='\(rowNumber)']").first as? XMLElement { row=existing }
        else { row=XMLElement(name:"row"); row.addAttribute(XMLNode.attribute(withName:"r",stringValue:String(rowNumber)) as! XMLNode); let place=(data.children ?? []).firstIndex { (Int($0.attr("r")) ?? 0)>rowNumber } ?? data.childCount; data.insertChild(row,at:place) }
        let c: XMLElement
        if let existing=try cell(index,ref) { c=existing }
        else { c=XMLElement(name:"c"); c.addAttribute(XMLNode.attribute(withName:"r",stringValue:ref) as! XMLNode); let place=(row.children ?? []).firstIndex { column($0.attr("r"))>column(ref) } ?? row.childCount; row.insertChild(c,at:place) }
        if let f=c.nodes("f").first, ["shared","array","dataTable"].contains(f.attr("t")) { throw NSError(domain:"Editor",code:2,userInfo:[NSLocalizedDescriptionKey:"Questa cella appartiene a una formula condivisa o matriciale: modificala in Excel."]) }
        for node in c.nodes("f|v|is") { node.detach() }; c.removeAttribute(forName:"t")
        if !value.isEmpty {
            if numeric { c.addChild(XMLElement(name:"v",stringValue:value)) }
            else { c.addAttribute(XMLNode.attribute(withName:"t",stringValue:"inlineStr") as! XMLNode); let inline=XMLElement(name:"is"); let t=XMLElement(name:"t",stringValue:value); t.addAttribute(XMLNode.attribute(withName:"xml:space",stringValue:"preserve") as! XMLNode); inline.addChild(t); c.addChild(inline) }
        }
        dirty=true
    }
    func save(to destination: URL) throws {
        if !dirty { if destination != url { try Data(contentsOf:url).write(to:destination,options:.atomic) }; return }
        let fm=FileManager.default, temp=fm.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try fm.createDirectory(at:temp,withIntermediateDirectories:true); defer { try? fm.removeItem(at:temp) }
        let archive=temp.appendingPathComponent("edited.xlsx"), staging=temp.appendingPathComponent("contents")
        try fm.copyItem(at:url,to:archive); try fm.createDirectory(at:staging,withIntermediateDirectories:true)
        var members=[String]()
        func write(_ path:String,_ doc:XMLDocument) throws {
            guard !path.hasPrefix("/"), !path.split(separator:"/").contains("..") else { throw NSError(domain:"Editor",code:3,userInfo:[NSLocalizedDescriptionKey:"Percorso interno non valido."]) }
            let target=staging.appendingPathComponent(path); try fm.createDirectory(at:target.deletingLastPathComponent(),withIntermediateDirectories:true)
            try doc.xmlData(options:[]).write(to:target); members.append(path)
        }
        for i in sheets.indices {
            let doc=try XMLDocument(data:sheetDocument(i).xmlData,options:[.nodePreserveWhitespace])
            // Discard stale formula caches: Excel will recompute them on opening.
            for cell in doc.nodes("/worksheet/sheetData/row/c[f]") { for v in cell.nodes("v") { v.detach() } }
            try write(sheets[i].path,doc)
        }
        let workbook=try xml("xl/workbook.xml")
        let calc=(workbook.nodes("/workbook/calcPr").first as? XMLElement) ?? XMLElement(name:"calcPr")
        if calc.parent == nil { workbook.rootElement()?.addChild(calc) }
        for (key,value) in [("calcMode","auto"),("fullCalcOnLoad","1"),("forceFullCalc","1")] { calc.removeAttribute(forName:key); calc.addAttribute(XMLNode.attribute(withName:key,stringValue:value) as! XMLNode) }
        try write("xl/workbook.xml",workbook)
        let rels=try xml("xl/_rels/workbook.xml.rels")
        for r in rels.nodes("/Relationships/Relationship") where r.attr("Type").hasSuffix("/calcChain") { r.detach() }
        try write("xl/_rels/workbook.xml.rels",rels)
        let types=try xml("[Content_Types].xml")
        for t in types.nodes("/Types/Override") where t.attr("PartName")=="/xl/calcChain.xml" { t.detach() }
        try write("[Content_Types].xml",types)
        func zip(_ args:[String],acceptMissing:Bool=false) throws {
            let process=Process(); process.executableURL=URL(fileURLWithPath:"/usr/bin/zip"); process.arguments=args; process.currentDirectoryURL=staging; process.standardOutput=FileHandle.nullDevice; process.standardError=FileHandle.nullDevice
            try process.run(); process.waitUntilExit()
            if process.terminationStatus != 0 && !(acceptMissing && process.terminationStatus == 12) { throw NSError(domain:"Editor",code:4,userInfo:[NSLocalizedDescriptionKey:"Impossibile creare il file Excel. Il file originale non è stato modificato."]) }
        }
        try zip(["-q","-d",archive.path,"xl/calcChain.xml"],acceptMissing:true)
        try zip(["-q",archive.path]+members)
        _ = try Book(archive)
        try Data(contentsOf:archive).write(to:destination,options:.atomic)
    }
    func html(_ index: Int) throws -> String {
        let doc = try sheetDocument(index)
        let cells = doc.nodes("/worksheet/sheetData/row/c").filter { !$0.nodes("v|is|f").isEmpty }
        let maxCol = cells.map { column($0.attr("r")) }.max() ?? 1
        let maxRow = cells.compactMap { Int($0.attr("r").drop(while: { $0.isLetter })) }.max() ?? 1
        var grid: [String:XMLNode] = [:]; for c in cells { grid[c.attr("r")] = c }
        var merges: [String:(Int,Int)] = [:]; var skipped = Set<String>()
        func ref(_ c: Int,_ r: Int) -> String { var v=c; var s=""; while v>0 { v-=1; s=String(UnicodeScalar(65+v%26)!)+s; v/=26 }; return s+String(r) }
        for merge in doc.nodes("/worksheet/mergeCells/mergeCell") {
            let parts = merge.attr("ref").split(separator:":").map(String.init); guard parts.count == 2 else { continue }
            let c1=column(parts[0]), c2=min(maxCol,column(parts[1])); let r1=Int(parts[0].drop(while:{$0.isLetter})) ?? 1; let r2=min(maxRow,Int(parts[1].drop(while:{$0.isLetter})) ?? r1)
            if c1 > c2 || r1 > r2 { continue }; merges[parts[0]]=(c2-c1+1,r2-r1+1)
            for r in r1...r2 { for c in c1...c2 where r != r1 || c != c1 { skipped.insert(ref(c,r)) } }
        }
        var widths = [Int:Double]()
        var hidden = Set<Int>()
        for c in doc.nodes("/worksheet/cols/col") {
            let lo=Int(c.attr("min")) ?? 1, hi=min(maxCol,Int(c.attr("max")) ?? 1)
            if lo<=hi { for i in lo...hi {
                let width=Double(c.attr("width")) ?? 12
                widths[i]=min(620,max(64,width*7+8))
                if c.attr("hidden") == "1" { hidden.insert(i) }
            } }
        }
        let columns = (1...maxCol).filter { !hidden.contains($0) }
        let totalWidth = 40 + columns.reduce(0.0) { $0 + (widths[$1] ?? 120) }
        let pane = doc.nodes("/worksheet/sheetViews/sheetView/pane").first
        let frozenRows = pane?.attr("state").hasPrefix("frozen") == true ? min(5,Int(pane?.attr("ySplit") ?? "") ?? 0) : 0
        let frozenCols = pane?.attr("state").hasPrefix("frozen") == true ? min(2,Int(pane?.attr("xSplit") ?? "") ?? 0) : 0
        var offsets = [Int:Double](); var offset=40.0
        for c in columns { offsets[c]=offset; offset += widths[c] ?? 120 }
        let filterRef=doc.val("/worksheet/autoFilter/@ref")
        let filterParts=filterRef.split(separator:":").map(String.init)
        let filterStart=filterParts.first ?? "", filterEnd=filterParts.last ?? ""
        let filterRow=Int(filterStart.drop(while:{$0.isLetter})) ?? 0
        let filterLastRow=Int(filterEnd.drop(while:{$0.isLetter})) ?? 0
        var html = """
        <!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>
        :root{color-scheme:light}*{box-sizing:border-box}body{margin:0;background:#fff;color:#202124;font:13px -apple-system; font-variant-numeric:tabular-nums}
        table{border-collapse:separate;border-spacing:0;table-layout:fixed;width:\(totalWidth)px}
        td,th{padding:0;border-right:1px solid #dfe3e9;border-bottom:1px solid #e5e7eb;vertical-align:middle;background:#fff}
        .value{display:block;height:26px;line-height:26px;padding:0 8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
        th{background:#f3f4f6;color:#7b8088;font-size:10px;font-weight:500;height:24px}
        thead th{position:sticky;top:0;z-index:6}tbody th{position:sticky;left:0;z-index:4;width:40px}
        .frozen{position:sticky}.heading .value{height:50px;font-size:11px;line-height:14px;padding:4px 6px;white-space:normal;word-break:normal;overflow-wrap:normal}
        td.selected{outline:2px solid #2979e8;outline-offset:-2px}tr:hover td{box-shadow:inset 0 0 0 999px #8080800c}::selection{background:#b9d9ff}
        .filter-head{cursor:pointer}.filter-head .value{padding-right:16px;position:relative}.filter-head .value:after{content:'▾';position:absolute;right:4px;top:5px}.filter-head.active .value:after{color:#06f;content:'●'}
        #filter-menu{position:fixed;z-index:100;background:#fff;color:#222;border:1px solid #ddd;border-radius:9px;box-shadow:0 5px 25px #0003;width:220px;padding:6px;max-height:340px;overflow:auto;font:13px -apple-system}
        #filter-menu button{display:block;border:0;background:transparent;border-radius:4px;text-align:left;width:100%;padding:7px;color:#222;cursor:pointer}#filter-menu button:hover{background:#e6efff}#filter-menu button[hidden]{display:none}#filter-menu input{width:100%;padding:7px;border:1px solid #ddd;border-radius:4px;margin-bottom:4px}
        </style></head><body><table><colgroup><col style="width:40px">
        """
        for c in columns { html += "<col style='width:\(widths[c] ?? 120)px'>" }; html += "</colgroup><thead><tr><th style='left:0;z-index:9'></th>"
        for c in columns { html += "<th style='\(c <= frozenCols ? "left:\(offsets[c] ?? 40)px;z-index:8;" : "")'>\(ref(c,1).dropLast())</th>" }; html += "</tr></thead><tbody>"
        for r in 1...maxRow {
            let isFrozen = r <= frozenRows
            let top = 24 + (r-1)*27
            let rowClass = r == frozenRows ? "heading" : ""
            html += "<tr data-row='\(r)' class='\(rowClass)'><th style='\(isFrozen ? "top:\(top)px;z-index:8;" : "")'>\(r)</th>"
            for c in columns {
                let key=ref(c,r); if skipped.contains(key) { continue }
                let cell=grid[key]; let idx=Int(cell?.attr("s") ?? "") ?? 0; var css=styles.indices.contains(idx) ? styles[idx].css : ""
                let span=merges[key] ?? (1,1)
                let visibleSpan=(c..<(c+span.0)).filter { !hidden.contains($0) }.count
                if isFrozen { css += "position:sticky;top:\(top)px;z-index:5;" }
                if c <= frozenCols { css += "position:sticky;left:\(offsets[c] ?? 40)px;z-index:\(isFrozen ? 7 : 3);" }
                let value=esc(cell.map(display) ?? "").replacingOccurrences(of:"'",with:"&#39;")
                html += "<td class='\(r == filterRow && c >= column(filterStart) && c <= column(filterEnd) ? "filter-head" : "")' data-col='\(c)' data-ref='\(key)' colspan='\(visibleSpan)' rowspan='\(span.1)' title='\(value)' style='\(esc(css))'><span class='value'>\(esc((cell.map(display) ?? "").replacingOccurrences(of:"\n",with:" ")))</span></td>"
            }; html += "</tr>"
        }
        return html + """
        </tbody></table><script>
        let filters={};
        function closeFilter(){document.querySelector('#filter-menu')?.remove()}
        function applyFilters(){document.querySelectorAll('tbody tr').forEach(row=>{const r=Number(row.dataset.row);if(r<=\(filterRow)||r>\(filterLastRow))return;row.hidden=Object.entries(filters).some(([c,v])=>(row.querySelector(`[data-col="${c}"]`)?.title||'')!==v)});document.querySelectorAll('.filter-head').forEach(c=>c.classList.toggle('active',filters[c.dataset.col]!==undefined));closeFilter()}
        window.clearFilters=()=>{filters={};applyFilters()};
        function showFilter(cell){closeFilter();const col=cell.dataset.col;const values=new Set();document.querySelectorAll('tbody tr').forEach(row=>{const r=Number(row.dataset.row);if(r>\(filterRow)&&r<=\(filterLastRow))values.add(row.querySelector(`[data-col="${col}"]`)?.title||'')});const menu=document.createElement('div');menu.id='filter-menu';const input=document.createElement('input');input.placeholder='Cerca valore…';menu.append(input);const add=(label,value,all=false)=>{const b=document.createElement('button');b.textContent=(all?filters[col]===undefined:filters[col]===value)?'✓ '+label:label;b.onclick=()=>{if(all)delete filters[col];else filters[col]=value;applyFilters()};menu.append(b);return b};add('Tutti i valori','',true);const buttons=[...values].sort((a,b)=>a.localeCompare(b,'it',{numeric:true})).map(v=>add(v||'(Vuote)',v));input.oninput=()=>buttons.forEach(b=>b.hidden=!b.textContent.toLowerCase().includes(input.value.toLowerCase()));document.body.append(menu);const box=cell.getBoundingClientRect();menu.style.left=Math.max(6,Math.min(box.left,innerWidth-230))+'px';menu.style.top=Math.min(box.bottom+3,innerHeight-350)+'px';input.focus()}
        document.addEventListener('click',e=>{if(e.target.closest('#filter-menu'))return;const cell=e.target.closest('td');if(!cell){closeFilter();return}document.querySelector('.selected')?.classList.remove('selected');cell.classList.add('selected');window.webkit?.messageHandlers.cell.postMessage({ref:cell.dataset.ref,text:cell.title});if(cell.classList.contains('filter-head'))showFilter(cell);else closeFilter()});
        document.addEventListener('dblclick',e=>{const cell=e.target.closest('td');if(cell&&!cell.classList.contains('filter-head'))window.webkit?.messageHandlers.cell.postMessage({ref:cell.dataset.ref,text:cell.title,edit:'1'})});
        document.addEventListener('keydown',e=>{if(e.key==='Escape')closeFilter()});
        </script></body></html>
        """

    }
}
final class App: NSObject, NSApplicationDelegate, WKNavigationDelegate, WKScriptMessageHandler, NSWindowDelegate {
    var window: NSWindow!; var web: WKWebView!; var picker: NSPopUpButton!; var status: NSTextField!; var book: Book?
    var zoomLabel: NSButton!; var search: NSSearchField!; var detail: NSTextView!
    var selectedRef=""
    let editUndo = UndoManager()
    var zoom = 0.85
    var notice: NSTextField!
    func updateEditState() {
        window.isDocumentEdited=book?.dirty ?? false
        notice.stringValue=(book?.dirty == true) ? "Modifiche non salvate · Le formule verranno ricalcolate in Excel" : "Doppio clic: modifica · Intestazioni ▾: filtra"
    }
    @objc func clearFilters() { web.evaluateJavaScript("clearFilters()",completionHandler:nil) }
    @objc func undoEdit() { if let modal=NSApp.modalWindow { modal.firstResponder?.undoManager?.undo(); return }; editUndo.undo(); updateEditState() }
    @objc func redoEdit() { if let modal=NSApp.modalWindow { modal.firstResponder?.undoManager?.redo(); return }; editUndo.redo(); updateEditState() }
    func restoreSheet(_ index:Int,_ data:Data,dirty:Bool) {
        guard let book else { return }
        do { let old=try book.sheetDocument(index).xmlData, wasDirty=book.dirty
            editUndo.registerUndo(withTarget:self) { $0.restoreSheet(index,old,dirty:wasDirty) }
            book.documents[book.sheets[index].path]=try XMLDocument(data:data,options:[.nodePreserveWhitespace]); book.dirty=dirty
            picker.selectItem(at:index); changeSheet(); updateEditState()
        } catch { show(error) }
    }
    @objc func editCell() {
        guard let book, !selectedRef.isEmpty else { return }
        do {
            let index=picker.indexOfSelectedItem, ref=selectedRef, cell=try book.cell(index,ref)
            let alert=NSAlert(); alert.messageText="Modifica \(ref)"
            alert.informativeText = cell?.nodes("f").isEmpty == false ? "Questa cella contiene una formula. Applicando la modifica la sostituisci con un valore. Le altre formule si ricalcoleranno in Excel." : "Modifica il valore della cella. Le formule dipendenti si ricalcoleranno quando aprirai il file salvato in Excel."
            alert.addButton(withTitle:"Applica"); alert.addButton(withTitle:"Annulla")
            let accessory=NSView(frame:NSRect(x:0,y:0,width:440,height:160))
            let kind=NSPopUpButton(frame:NSRect(x:0,y:132,width:440,height:26)); kind.addItems(withTitles:["Testo","Numero (decimali con virgola o punto)"]); accessory.addSubview(kind)
            let scroll=NSScrollView(frame:NSRect(x:0,y:0,width:440,height:122)); scroll.borderType = .bezelBorder; scroll.hasVerticalScroller=true
            let input=NSTextView(frame:scroll.bounds); input.isRichText=false; input.font = .systemFont(ofSize:14); input.isVerticallyResizable=true; input.autoresizingMask=[.width]; input.textContainer?.widthTracksTextView=true
            let raw=cell?.val("v") ?? ""
            let numeric=cell != nil && ["","n"].contains(cell!.attr("t")) && Double(raw) != nil
            if numeric { kind.selectItem(at:1); input.string=raw } else { input.string=cell.map(book.display) ?? "" }
            scroll.documentView=input; accessory.addSubview(scroll); alert.accessoryView=accessory; alert.window.initialFirstResponder=input
            while alert.runModal() == .alertFirstButtonReturn {
                var value=input.string
                if kind.indexOfSelectedItem==1 && !value.isEmpty {
                    value=value.trimmingCharacters(in:.whitespacesAndNewlines).replacingOccurrences(of:",",with:".")
                    guard let n=Double(value), n.isFinite else { alert.informativeText="Inserisci un numero senza separatori delle migliaia (es. 50000 o 3,65)."; continue }
                    value=String(n)
                }
                let old=try book.sheetDocument(index).xmlData, wasDirty=book.dirty
                try book.setCell(index,ref,value,numeric:kind.indexOfSelectedItem==1)
                editUndo.registerUndo(withTarget:self) { $0.restoreSheet(index,old,dirty:wasDirty) }; editUndo.setActionName("Modifica cella")
                // Update the displayed value without resetting scroll position or filters.
                let updated=try book.cell(index,ref).map(book.display) ?? ""
                let json=String(data:try JSONSerialization.data(withJSONObject:[ref,updated]),encoding:.utf8)!
                web.evaluateJavaScript("(()=>{const [ref,value]=\(json);const c=document.querySelector('[data-ref=\\\"'+ref+'\\\"]');if(c){c.title=value;c.querySelector('.value').textContent=value.replaceAll('\\n',' ')}applyFilters()})()",completionHandler:nil)
                detail.string="\(ref)   \(updated)"; updateEditState(); break
            }
        } catch { show(error) }
    }
    @discardableResult @objc func saveCopy() -> Bool {
        guard let book else { return false }
        let panel=NSSavePanel(); panel.allowedContentTypes=[UTType(filenameExtension:"xlsx")!]; panel.nameFieldStringValue=book.url.deletingPathExtension().lastPathComponent+" - modificato.xlsx"; panel.directoryURL=book.url.deletingLastPathComponent(); panel.title="Salva file Excel"
        guard panel.runModal() == .OK, let target=panel.url else { return false }
        do { try book.save(to:target); let newBook=try Book(target); self.book=newBook; editUndo.removeAllActions(); window.title=target.lastPathComponent+" — Lettore XLSX"; window.representedURL=target; UserDefaults.standard.set(target.path,forKey:"lastFile"); changeSheet(); updateEditState(); return true } catch { show(error); return false }
    }
    func confirmUnsaved() -> Bool {
        guard book?.dirty == true else { return true }
        let alert=NSAlert(); alert.messageText="Salvare le modifiche?"; alert.informativeText="Ci sono modifiche non salvate in questo file."; alert.addButton(withTitle:"Salva con nome…"); alert.addButton(withTitle:"Annulla"); alert.addButton(withTitle:"Non salvare")
        switch alert.runModal() { case .alertFirstButtonReturn:return saveCopy(); case .alertThirdButtonReturn:return true; default:return false }
    }
    func windowShouldClose(_ sender:NSWindow)->Bool { let close=confirmUnsaved(); if close { book?.dirty=false }; return close }
    func applicationShouldTerminate(_ sender:NSApplication)->NSApplication.TerminateReply { confirmUnsaved() ? .terminateNow : .terminateCancel }

    func applyZoom(_ value: Double) {
        zoom = min(2,max(0.5,value)); web.pageZoom=zoom; web.magnification=1
        zoomLabel.title="\(Int((zoom*100).rounded()))%"
        UserDefaults.standard.set(zoom,forKey:"pageZoom")
    }
    @objc func zoomIn() { applyZoom(zoom+0.1) }
    @objc func zoomOut() { applyZoom(zoom-0.1) }
    @objc func resetZoom() { applyZoom(1) }
    @objc func focusSearch() { window.makeFirstResponder(search) }
    @objc func findNext() {
        guard !search.stringValue.isEmpty else { return }
        let config=WKFindConfiguration(); config.wraps=true
        web.find(search.stringValue,configuration:config) { result in
            self.search.toolTip=result.matchFound ? "Invio: risultato successivo" : "Nessun risultato nel foglio"
            if !result.matchFound { NSSound.beep() }
        }
    }
    func userContentController(_ userContentController: WKUserContentController,didReceive message: WKScriptMessage) {
        guard let data=message.body as? [String:String] else { return }
        selectedRef=data["ref"] ?? ""
        if data["edit"] == "1" { DispatchQueue.main.async { self.editCell() } }
        detail.string="\(data["ref"] ?? "")   \(data["text"] ?? "")"
    }
    func applicationDidFinishLaunching(_ notification: Notification) {
        let menu=NSMenu(); let root=NSMenuItem(); menu.addItem(root); let appMenu=NSMenu(); root.submenu=appMenu
        appMenu.addItem(withTitle:"Apri…",action:#selector(openFile),keyEquivalent:"o").target=self
        appMenu.addItem(withTitle:"Salva con nome…",action:#selector(saveCopy),keyEquivalent:"s").target=self
        appMenu.addItem(.separator()); appMenu.addItem(withTitle:"Esci",action:#selector(NSApplication.terminate(_:)),keyEquivalent:"q")
        let editRoot=NSMenuItem(); editRoot.title="Modifica"; menu.addItem(editRoot); let edit=NSMenu(title:"Modifica"); editRoot.submenu=edit
        edit.addItem(withTitle:"Annulla modifica cella",action:#selector(undoEdit),keyEquivalent:"z").target=self
        let redo=edit.addItem(withTitle:"Ripristina modifica cella",action:#selector(redoEdit),keyEquivalent:"z"); redo.keyEquivalentModifierMask=[.command,.shift]; redo.target=self
        edit.addItem(withTitle:"Incolla",action:#selector(NSText.paste(_:)),keyEquivalent:"v")
        edit.addItem(withTitle:"Taglia",action:#selector(NSText.cut(_:)),keyEquivalent:"x")
        edit.addItem(withTitle:"Copia",action:#selector(NSText.copy(_:)),keyEquivalent:"c"); edit.addItem(withTitle:"Seleziona tutto",action:#selector(NSText.selectAll(_:)),keyEquivalent:"a"); let viewRoot=NSMenuItem(); viewRoot.title="Vista"; menu.addItem(viewRoot); let viewMenu=NSMenu(title:"Vista"); viewRoot.submenu=viewMenu
        for (title,action,key) in [("Ingrandisci",#selector(zoomIn),"+"),("Riduci",#selector(zoomOut),"-"),("Dimensioni originali",#selector(resetZoom),"0"),("Cerca nel foglio…",#selector(focusSearch),"f"),("Trova successivo",#selector(findNext),"g")] { viewMenu.addItem(withTitle:title,action:action,keyEquivalent:key).target=self }
        NSApp.mainMenu=menu
        window=NSWindow(contentRect:NSRect(x:0,y:0,width:1180,height:760),styleMask:[.titled,.closable,.miniaturizable,.resizable],backing:.buffered,defer:false)
        window.delegate=self; window.title="Lettore XLSX"; window.minSize=NSSize(width:1050,height:400); window.center(); window.setFrameAutosaveName("ReaderWindow")
        let content=window.contentView!
        let open=NSButton(title:"Apri…",target:self,action:#selector(openFile))
        picker=NSPopUpButton(); picker.target=self; picker.action=#selector(changeSheet)
        status=NSTextField(labelWithString:"Doppio clic per modificare"); status.textColor = .secondaryLabelColor; status.font = .systemFont(ofSize:12)
        search=NSSearchField(); search.placeholderString="Cerca nel foglio"; search.target=self; search.action=#selector(findNext); search.sendsSearchStringImmediately=false; search.sendsWholeSearchString=true
        search.widthAnchor.constraint(equalToConstant:185).isActive=true
        let minus=NSButton(title:"−",target:self,action:#selector(zoomOut)); minus.toolTip="Riduci (⌘−)"
        let plus=NSButton(title:"+",target:self,action:#selector(zoomIn)); plus.toolTip="Ingrandisci (⌘+)"
        zoomLabel=NSButton(title:"85%",target:self,action:#selector(resetZoom)); zoomLabel.toolTip="Ripristina 100% (⌘0)"; zoomLabel.widthAnchor.constraint(equalToConstant:54).isActive=true
        let editButton=NSButton(title:"Modifica",target:self,action:#selector(editCell))
        let saveButton=NSButton(title:"Salva…",target:self,action:#selector(saveCopy))
        let resetFilters=NSButton(title:"Tutti",target:self,action:#selector(clearFilters)); resetFilters.toolTip="Rimuovi tutti i filtri"
        let bar=NSStackView(views:[open,picker,editButton,saveButton,resetFilters,NSView(),search,minus,zoomLabel,plus]); bar.spacing=8; bar.orientation = .horizontal
        let footer=NSScrollView(); footer.hasVerticalScroller=true; footer.borderType = .noBorder
        detail=NSTextView(); detail.isEditable=false; detail.isSelectable=true; detail.font = .systemFont(ofSize:12); detail.textColor = .secondaryLabelColor; detail.backgroundColor = .windowBackgroundColor; detail.textContainerInset=NSSize(width:12,height:6)
        detail.isVerticallyResizable=true; detail.isHorizontallyResizable=false; detail.autoresizingMask=[.width]; detail.textContainer?.widthTracksTextView=true
        detail.string="Seleziona una cella per leggere il contenuto completo · Doppio clic per modificare"
        footer.documentView=detail
        notice=NSTextField(labelWithString:"Doppio clic: modifica · Intestazioni ▾: filtra"); notice.font = .systemFont(ofSize:11); notice.textColor = .secondaryLabelColor
        let config=WKWebViewConfiguration(); config.defaultWebpagePreferences.allowsContentJavaScript=true; config.userContentController.add(self,name:"cell")
        web=WKWebView(frame:.zero,configuration:config); web.navigationDelegate=self; web.allowsMagnification=false
        let savedZoom=UserDefaults.standard.double(forKey:"pageZoom"); applyZoom(savedZoom > 0 ? savedZoom : 0.85)
        for view in [bar,web!,footer,notice!] { view.translatesAutoresizingMaskIntoConstraints=false; content.addSubview(view) }
        NSLayoutConstraint.activate([bar.topAnchor.constraint(equalTo:content.topAnchor,constant:12),bar.leadingAnchor.constraint(equalTo:content.leadingAnchor,constant:16),bar.trailingAnchor.constraint(equalTo:content.trailingAnchor,constant:-16),bar.heightAnchor.constraint(equalToConstant:30),web.topAnchor.constraint(equalTo:bar.bottomAnchor,constant:12),web.leadingAnchor.constraint(equalTo:content.leadingAnchor),web.trailingAnchor.constraint(equalTo:content.trailingAnchor),web.bottomAnchor.constraint(equalTo:footer.topAnchor),footer.leadingAnchor.constraint(equalTo:content.leadingAnchor),footer.trailingAnchor.constraint(equalTo:content.trailingAnchor),footer.bottomAnchor.constraint(equalTo:notice.topAnchor),notice.leadingAnchor.constraint(equalTo:content.leadingAnchor,constant:12),notice.trailingAnchor.constraint(equalTo:content.trailingAnchor,constant:-12),notice.bottomAnchor.constraint(equalTo:content.bottomAnchor,constant:-5),notice.heightAnchor.constraint(equalToConstant:16),footer.heightAnchor.constraint(equalToConstant:48),picker.widthAnchor.constraint(equalToConstant:220)])
        window.makeKeyAndOrderFront(nil); NSApp.activate(ignoringOtherApps:true)
        if book == nil, let saved=UserDefaults.standard.string(forKey:"lastFile"), FileManager.default.fileExists(atPath:saved) { load(URL(fileURLWithPath:saved)) }
    }
    func application(_ sender:NSApplication,openFiles filenames:[String]) { if window == nil { DispatchQueue.main.async { if let f=filenames.first { self.load(URL(fileURLWithPath:f)) } } } else if let f=filenames.first { load(URL(fileURLWithPath:f)) }; sender.reply(toOpenOrPrint:.success) }
    func applicationShouldTerminateAfterLastWindowClosed(_ sender:NSApplication)->Bool { true }
    @objc func openFile() { let p=NSOpenPanel(); p.allowedContentTypes=[UTType(filenameExtension:"xlsx")!]; p.canChooseDirectories=false; p.allowsMultipleSelection=false; if p.runModal() == .OK, let url=p.url { load(url) } }
    func load(_ url:URL) { guard confirmUnsaved() else { return }; do { let newBook=try Book(url); let html=try newBook.html(0); book=newBook; editUndo.removeAllActions(); selectedRef=""; updateEditState(); picker.removeAllItems(); picker.addItems(withTitles:newBook.sheets.map(\.name)); window.title=url.lastPathComponent + " — Lettore XLSX"; window.representedURL=url; status.stringValue="\(newBook.sheets.count) fogli · Doppio clic per modificare"; web.loadHTMLString(html,baseURL:nil); UserDefaults.standard.set(url.path,forKey:"lastFile") } catch { show(error) } }
    @objc func changeSheet() { selectedRef=""; detail.string="Seleziona una cella per leggerla · Doppio clic per modificarla"; do { if let book { web.loadHTMLString(try book.html(picker.indexOfSelectedItem),baseURL:nil) } } catch { show(error) } }
    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) { webView.pageZoom=zoom }
    func show(_ error:Error) { let a=NSAlert(error:error); a.runModal() }
    func webView(_ webView:WKWebView,decidePolicyFor navigationAction:WKNavigationAction,decisionHandler:@escaping(WKNavigationActionPolicy)->Void) { decisionHandler(navigationAction.navigationType == .linkActivated ? .cancel : .allow) }
}
if CommandLine.arguments.count > 2 && CommandLine.arguments[1] == "--check" {
    do { let b=try Book(URL(fileURLWithPath:CommandLine.arguments[2])); for (i,s) in b.sheets.enumerated() { let h=try b.html(i); print("OK \(s.name): \(h.utf8.count) bytes HTML"); if CommandLine.arguments.count > 3 { try h.write(toFile:CommandLine.arguments[3]+"/sheet\(i).html",atomically:true,encoding:.utf8) } } } catch { fputs("\(error)\n",stderr); exit(1) }
} else { let app=NSApplication.shared; let delegate=App(); app.delegate=delegate; app.setActivationPolicy(.regular); app.run() }
