import AppKit
import PDFKit
import WebKit

final class PrintWaiter: NSObject {
    var done = false
    @objc func printOperationDidRun(_ operation: NSPrintOperation, success: Bool, contextInfo: UnsafeMutableRawPointer?) { done = true }
}
let printWaiter = PrintWaiter()

// Debug-only end-to-end checks: open a workbook in the real window, navigate, edit, save a copy, print to PDF.
func scheduleLegacySheetTests(_ app: App, fixture: URL) {
    let out = fixture.deletingLastPathComponent()
    func check(_ condition: Bool, _ message: String) {
        guard condition else { fputs("FAIL: \(message)\n", stderr); exit(1) }
        print("PASS: \(message)")
    }
    func js(_ script: String, _ done: @escaping (Any?) -> Void) {
        app.web.evaluateJavaScript(script) { value, error in
            if let error { fputs("JS ERROR: \(error)\n", stderr); exit(1) }
            done(value)
        }
    }
    func waitUntil(_ condition: @escaping () -> Bool, timeout: Double = 30, _ next: @escaping () -> Void) {
        let start = Date()
        Timer.scheduledTimer(withTimeInterval: 0.1, repeats: true) { timer in
            if condition() { timer.invalidate(); DispatchQueue.main.asyncAfter(deadline: .now() + 0.5, execute: next) }
            else if Date().timeIntervalSince(start) > timeout { fputs("FAIL: timeout\n", stderr); exit(1) }
        }
    }
    DispatchQueue.main.asyncAfter(deadline: .now() + 0.3) {
        let original = try! Data(contentsOf: fixture)
        let started = Date()
        app.load(fixture)
        waitUntil({ app.book != nil && !app.loading && !app.web.isLoading }) {
            print(String(format: "INFO: opened in %.2f s", Date().timeIntervalSince(started)))
            check(app.picker.numberOfItems == 2, "Sheet picker lists all sheets")
            js("document.querySelectorAll('tbody tr').length") { rows in
                check((rows as? Int) == 20001, "All rows rendered")
                js("document.querySelector('tr[data-row=\"2\"] td[data-c=\"1\"]').click()") { _ in
                    waitUntil({ app.selectedRef == "A2" }) {
                        check(true, "Click selects a cell")
                        js("document.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowRight'}));document.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown'}))") { _ in
                            waitUntil({ app.selectedRef == "B3" }) {
                                check(true, "Arrow keys move the selection")
                                js("getSelection().toString()") { text in
                                    check((text as? String) == "Cliente 2", "Selected cell text is ready for Copy")
                                    js("(()=>{const td=document.querySelector('tr[data-row=\"1\"] td[data-c=\"3\"]');td.click();const b=[...document.querySelectorAll('#filter-menu button')].find(b=>b.textContent==='Roma');b.click();return [...document.querySelectorAll('tbody tr')].filter(r=>!r.hidden).length})()") { visible in
                                        let count = visible as? Int ?? 0
                                        check(count > 1000 && count < 15000, "Filter hides other rows (\(count) visible)")
                                        js("clearFilters();[...document.querySelectorAll('tbody tr')].filter(r=>!r.hidden).length") { all in
                                            check((all as? Int) == 20001, "Clear filters")
                                            do {
                                                let book = app.book!
                                                try book.setCell(0, "B3", "Edited ✓", kind: .text)
                                                try book.setCell(0, "D3", "42.5", kind: .number)
                                                try book.setCell(0, "H2", "=somma(D2:D4)*2", kind: .formula)
                                                try book.setCell(0, "H3", "=SE(D3>40;\"alto\";\"basso\")", kind: .formula)
                                                try book.setCell(0, "H4", "=SUMIF(C2:C20001,\"Roma\",D2:D20001)/COUNTIF(C2:C20001,\"Roma\")", kind: .formula)
                                                let d2 = try book.value(0, 2, 4).number!, d4 = try book.value(0, 4, 4).number!
                                                check(abs(try book.value(0, 2, 8).number! - (d2 + 42.5 + d4) * 2) < 1e-9, "Italian formula SOMMA with recalculated edit")
                                                check(try book.value(0, 3, 8).text == "alto", "SE with semicolons")
                                                check((try book.value(0, 4, 8).number ?? 0) > 0, "SUMIF / COUNTIF over 20,000 rows")
                                                check(try book.cell(0, "H2")?.child("f")?.stringValue == "SUM(D2:D4)*2", "Formula stored in English syntax")
                                                let total = try book.value(1, 1, 2).number ?? -1
                                                let expected = try (2...20001).reduce(0.0) { $0 + (try book.value(0, $1, 4).number ?? 0) }
                                                check(abs(total - expected) < 1e-6, "Cross-sheet SUM(Dati!D2:D20001) without a cached value")
                                                let copy = out.appendingPathComponent("edited-copy.xlsx")
                                                try? FileManager.default.removeItem(at: copy)
                                                try book.save(to: copy)
                                                let reopened = try Book(copy)
                                                check(try reopened.cell(0, "B3").map { reopened.display($0, sheet: 0) } == "Edited ✓", "Saved copy keeps text edit")
                                                check(try reopened.value(0, 3, 8).text == "alto", "Saved copy recalculates formulas")
                                                check(try reopened.cell(0, "D3")?.child("v")?.stringValue == "42.5", "Saved copy keeps number edit")
                                                check(try Data(contentsOf: fixture) == original, "Original file is unchanged")
                                            } catch { check(false, "Edit and save: \(error)") }
                                            let pdf = out.appendingPathComponent("print.pdf")
                                            try? FileManager.default.removeItem(at: pdf)
                                            let info = NSPrintInfo.shared.copy() as! NSPrintInfo
                                            info.jobDisposition = .save; info.dictionary()[NSPrintInfo.AttributeKey.jobSavingURL] = pdf
                                            info.orientation = .landscape; info.horizontalPagination = .fit
                                            let op = app.web.printOperation(with: info); op.view?.frame = app.web.bounds
                                            op.showsPrintPanel = false; op.showsProgressPanel = false
                                            op.runModal(for: app.window, delegate: printWaiter, didRun: #selector(PrintWaiter.printOperationDidRun(_:success:contextInfo:)), contextInfo: nil)
                                            waitUntil({ printWaiter.done }, timeout: 300) {
                                                check((PDFDocument(url: pdf)?.pageCount ?? 0) > 0, "Printed PDF is readable (\(PDFDocument(url: pdf)?.pageCount ?? 0) pages)")
                                                check(true, "Print to PDF")
                                                if let content = app.window.contentView, let bitmap = content.bitmapImageRepForCachingDisplay(in: content.bounds) {
                                                    content.cacheDisplay(in: content.bounds, to: bitmap)
                                                    try? bitmap.representation(using: .png, properties: [:])?.write(to: out.appendingPathComponent("xlsx-preview.png"))
                                                }
                                                print("ALL SHEET TESTS PASSED")
                                                app.book?.dirty = false
                                                NSApp.terminate(nil)
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

func sheetCheck(_ ok: Bool, _ message: String) {
    guard ok else { fputs("FAIL: \(message)\n",stderr); exit(1) }
    print("PASS: \(message)")
}

/// Runs without NSApplication as well, so the formula/file checks are usable on a headless host.
final class UndoProbe: NSObject {
    let history = SheetUndoManager()
    var value = "", commits = 0
    func apply(_ next: String) {
        guard next != value else { return }
        let old = value; value = next; commits += 1
        history.record(self,name:"Edit") { $0.restore(old) }
    }
    func restore(_ next: String) {
        let old = value; value = next
        history.registerUndo(withTarget:self) { $0.restore(old) }
    }
}

func runSheetModelTests(_ fixture: URL) {
    do {
        let probe = UndoProbe()
        sheetCheck(!probe.history.groupsByEvent,"Document undo disables implicit event grouping")
        for action in ["typing","delete","paste","cut","fill down","sort","format"] {
            probe.history.removeAllActions(); probe.value = "before"
            probe.apply("after"); probe.apply("after")
            sheetCheck(probe.history.groupingLevel == 0,"\(action): exactly one closed undo group")
            probe.history.undo()
            sheetCheck(probe.value == "before" && !probe.history.canUndo && probe.history.canRedo,"\(action): ONE undo restores the prior state")
            probe.history.redo()
            sheetCheck(probe.value == "after" && !probe.history.canRedo,"\(action): ONE redo restores the edit")
        }
        let b = try Book.blank()
        let entries = [("A1","Key"),("B1","Amount"),("C1","Group"),("A2","1"),("B2","10"),("C2","Roma"),("A3","2"),("B3","20"),("C3","Rimini"),("A4","3"),("B4","30"),("C4","Milano"),("D2","30"),("D3","20"),("D4","10"),("E2","1"),("F2","2"),("G2","3"),("E3","uno"),("F3","due"),("G3","tre")]
        try b.applyEdits(0,entries.map { try b.automaticEdit(0,$0.0,$0.1) })
        let cases: [(String,String)] = [
            ("SUMIFS(B2:B4,C2:C4,\"R*\",A2:A4,\">1\")","20"),
            ("COUNTIFS(C2:C4,\"R?m*\",B2:B4,\">=10\")","2"),
            ("AVERAGEIFS(B2:B4,C2:C4,\"R*\")","15"),
            ("MAXIFS(B2:B4,C2:C4,\"R*\")","20"),("MINIFS(B2:B4,C2:C4,\"R*\")","10"),
            ("IFS(1>2,3,2>1,4)","4"),("IFS(FALSE,1)","#N/A"),
            ("VLOOKUP(2,A2:C4,3,FALSE)","Rimini"),("VLOOKUP(2.5,A2:C4,2,TRUE)","20"),
            ("HLOOKUP(2,E2:G3,2,FALSE)","due"),("HLOOKUP(2.5,E2:G3,2)","due"),
            ("XLOOKUP(2,A2:A4,C2:C4)","Rimini"),("XLOOKUP(2,A2:A4,C2:C4,1/0)","Rimini"),("XLOOKUP(9,A2:A4,C2:C4,\"missing\")","missing"),
            ("INDEX(A2:C4,2,3)","Rimini"),("INDEX(E2:G2,2)","2"),
            ("MATCH(2,A2:A4,0)","2"),("MATCH(2.5,A2:A4,1)","2"),("MATCH(25,D2:D4,-1)","1"),
            ("LEFT(\"Ciao\",2)","Ci"),("RIGHT(\"Ciao\",2)","ao"),("MID(\"Ciao\",2,2)","ia"),
            ("FIND(\"a\",\"Ciao\")","3"),("SEARCH(\"I?\",\"Ciao\")","2"),
            ("SUBSTITUTE(\"a-a-a\",\"a\",\"b\",2)","a-b-a"),
            ("TEXTJOIN(\"|\",TRUE,C2:C4)","Roma|Rimini|Milano"),("VALUE(\"€12,5\")","12.5"),
            ("DATE(2024,2,29)","45351"),("YEAR(DATE(2024,2,29))","2024"),
            ("MONTH(DATE(2024,2,29))","2"),("DAY(DATE(2024,2,29))","29"),
            ("WEEKDAY(DATE(2024,2,29),2)","4"),("COUNTBLANK(H2:H4)","3"),("RANK(20,B2:B4)","2"),
            ("SUMIF(C2:C4,\"R*\",B2:B4)","30"),("COUNTIF(C2:C4,\"R?????\")","1"),
            ("SUM(B:B)","60"),("COUNTBLANK(H1:H8)","8"),
            ("INDEX(A2:C4,4,1)","#REF!"),("MATCH(9,A2:A4,0)","#N/A"),
            ("SUMIFS(B2:B4,A2:A3,1)","#VALUE!"),("AVERAGEIFS(B2:B4,A2:A4,99)","#DIV/0!"),
            ("LEFT(1/0,1)","#DIV/0!"),("IFERROR(VLOOKUP(9,A2:C4,2,FALSE),42)","42"),
            ("FIND(\"x\",\"abc\")","#VALUE!"),("LEFT(\"abc\",-1)","#VALUE!"),
            ("DATE(2024,13,1)","45658"),("DATE(1900,2,29)","60"),("DAY(60)","29"),
            ("COUNTIF(C2:C4,\"~*\")","0"),("XLOOKUP(\"R*\",C2:C4,B2:B4)","#N/A")
        ]
        for (formula,expected) in cases {
            let value = try FormulaParser(book:b,sheet:0).evaluate(formula)
            sheetCheck(value.text == expected,"\(formula) = \(expected) (got \(value.text))")
        }
        // Every new localized name is checked against the same expected result, using semicolons.
        let names = ["SUMIFS","COUNTIFS","AVERAGEIFS","MAXIFS","MINIFS","IFS","VLOOKUP","HLOOKUP","XLOOKUP","INDEX","MATCH","LEFT","RIGHT","MID","FIND","SEARCH","SUBSTITUTE","TEXTJOIN","VALUE","DATE","YEAR","MONTH","DAY","WEEKDAY","COUNTBLANK","RANK"]
        for name in names {
            let pair = cases.first { $0.0.hasPrefix(name+"(") }!, localized = italianFunctions.first { $0.value == name }!.key
            var quoted = false, formula = ""
            for ch in pair.0.replacingOccurrences(of:name+"(",with:localized+"(") { if ch == "\"" { quoted.toggle() }; formula.append(ch == "," && !quoted ? ";" : ch) }
            let value = try FormulaParser(book:b,sheet:0).evaluate(normalizeFormula("="+formula))
            sheetCheck(value.text == pair.1,"Italian \(localized) with semicolons")
        }
        b.calculationTime = Date(timeIntervalSince1970:0)
        let now = try FormulaParser(book:b,sheet:0).evaluate("NOW()").number!
        let italianNow = try FormulaParser(book:b,sheet:0).evaluate(normalizeFormula("=ADESSO()")).number!
        let exactSerial = 25569+Double(TimeZone.current.secondsFromGMT(for:b.calculationTime))/86400
        sheetCheck(now == exactSerial && italianNow == exactSerial,"NOW / ADESSO exact local Excel serial at fixed clock")
        let cross = try Book(fixture)
        sheetCheck(try FormulaParser(book:cross,sheet:1).evaluate("COUNT(Dati!A:A)").text == "20000","Whole-column cross-sheet range bounded to used rows")
        for bad in ["=SUM(1+)","=UNKNOWN(1+)","=1+","=SUM((1)","=\"unterminated"] {
            do { _ = try b.automaticEdit(0,"H1",bad); sheetCheck(false,"Reject \(bad)") } catch { sheetCheck(true,"Reject \(bad)") }
        }
        _ = try b.automaticEdit(0,"H1","=UNKNOWN(1,2)")
        sheetCheck(true,"Unknown functions accepted after full syntax validation")
        sheetCheck(automaticNumber(" € 12,5% ") == 0.125 && automaticNumber("hello") == nil,"Automatic decimal/currency/percent types")
        sheetCheck(parseTSV("\"a\tb\"\t\"two\nlines\"\r\n3\t4\r\n") == [["a\tb","two\nlines"],["3","4"]],"Quoted TSV round trip grammar")
        let out = fixture.deletingLastPathComponent(), blank = try Book.blank(), saved = out.appendingPathComponent("new-blank.xlsx")
        try blank.save(to:saved)
        sheetCheck(try Book(saved).sheets.count == 1,"Minimal blank package reopens with Book")
        let python = Process(); python.executableURL = URL(fileURLWithPath:"/usr/bin/env"); python.arguments = ["python3","-c","import openpyxl,sys; w=openpyxl.load_workbook(sys.argv[1]); assert len(w.sheetnames)==1",saved.path]
        try python.run(); python.waitUntilExit(); sheetCheck(python.terminationStatus == 0,"Minimal blank package opens with openpyxl")
        try blank.setCell(0,"A1","saved in place",kind:.text); try blank.save(to:saved)
        let reopened = try Book(saved)
        sheetCheck(try reopened.raw(0,"A1") == "saved in place","Atomic save in place then reopen")
        try b.sort(0,column:2,row:2,ascending:false)
        sheetCheck(try b.raw(0,"A2") == "3" && b.raw(0,"B2") == "30","Sort descending moves complete rows")
        try b.sort(0,column:2,row:2,ascending:true)
        sheetCheck(try b.raw(0,"B2") == "10","Sort ascending")
        try b.setCell(0,"B3","SUM(1,2)",kind:.formula)
        do { try b.sort(0,column:2,row:2,ascending:true); sheetCheck(false,"Refuse sorting formulas") } catch { sheetCheck(error.localizedDescription.contains(L("formule","formulas")),"Refuse sorting formulas with explanation") }
        do { try cross.sort(1,column:1,row:3,ascending:true); sheetCheck(false,"Refuse merged region") } catch { sheetCheck(true,"Refuse merged region") }
        let features = try Book(out.appendingPathComponent("features.xlsx"))
        try features.sort(0,column:1,row:2,ascending:true)
        sheetCheck(try (2...6).map { try features.raw(0,"A"+String($0)) } == ["1","3","Alpha","beta",""],"Sort: numeric before text, empty last, styles and payload move together")
        try features.sort(0,column:1,row:2,ascending:false)
        sheetCheck(try (2...6).map { try features.raw(0,"A"+String($0)) } == ["3","1","beta","Alpha",""],"Descending sort keeps numeric before text and empty last")
        try features.sort(1,column:1,row:6,ascending:true)
        sheetCheck(try features.raw(1,"A6") == "b" && features.raw(1,"A2") == "z","No-filter sort stays inside the contiguous block below its header")
        try features.setCell(2,"C2","0.125",kind:.number)
        sheetCheck(try features.cell(2,"C2")?.attr("s") != features.cell(2,"C1")?.attr("s"),"Styled blank cell retains its distinct fill style")
        let formatted = try features.display(features.cell(2,"C2")!,sheet:2)
        sheetCheck(formatted.contains("%"),"Editing preserves existing percent formatting")
        sheetCheck(try features.value(2,1,1).text == "123","Unknown function uses cached result after edits")
        for (row,error) in [(1,"#N/A"),(2,"#VALUE!"),(3,"#DIV/0!"),(4,"#REF!")] {
            sheetCheck(try FormulaParser(book:features,sheet:2).evaluate("LEFT(B\(row),1)").text == error,"Propagate \(error) through text function")
        }
        try features.setCell(2,"C1","200",kind:.number)
        sheetCheck(try features.value(2,2,1).text == "201","Known formula recalculates a stale cache after an input edit")
        let savedFeatures = out.appendingPathComponent("features-saved.xlsx")
        try features.save(to:savedFeatures); let readFeatures = try Book(savedFeatures)
        sheetCheck(try readFeatures.value(2,1,1).text == "123","Unknown function cache survives save/reopen")
        sheetCheck(try readFeatures.value(2,2,1).text == "201","Recalculated known formula survives save/reopen")
        sheetCheck(try readFeatures.xml("customXml/keep.xml").rootElement()?.stringValue == "untouched package part","Untouched XLSX package member survives save")
        let beforeProtection = try features.sheetDocument(4).xmlData
        do { try features.applyEdits(4,[CellEdit(ref:"A1",value:"changed",kind:.text),CellEdit(ref:"A2",value:"new",kind:.text)]); sheetCheck(false,"Protected batch refused") } catch { sheetCheck(try features.sheetDocument(4).xmlData == beforeProtection,"Protected batch refused without partial changes") }
        do { try features.sort(3,column:1,row:2,ascending:true); sheetCheck(false,"Merged sorting refused") } catch { sheetCheck(error.localizedDescription.contains(L("unite","merged")),"Merged sorting refused with explanation") }
        sheetCheck(fillFormula("SUM(A2,$B2,C$2,$D$2,'A2'!A2,\"A2\")",down:2) == "SUM(A4,$B4,C$2,$D$2,'A2'!A4,\"A2\")","Fill Down shifts relative rows and preserves absolute refs and quoted text")
        for (f,e) in [("DATE(1900,3,0)","60"),("DATE(1900,2,30)","61"),("YEAR(0)","1900"),("WEEKDAY(1)","1"),("SUBSTITUTE(\"abc\",\"x\",\"y\")","abc"),("TEXTJOIN(\"|\",FALSE,H2:H4)","||"),("COUNTIFS(A2:A4,\">0\",B2:B4,\"<25\")","2"),("COUNTIF(B2:B4,#N/A)","#N/A"),("FIND(\"\",\"abc\",2)","2"),("COUNTIF(A2:A4,\">=1,5\")","2"),("SUM(INDEX(A2:B4,2))","5"),("VALUE(\"12no\")","#VALUE!"),("VALUE(\"12:00\")","0.5")] {
            sheetCheck(try FormulaParser(book:b,sheet:0).evaluate(f).text == e,"Edge case \(f) = \(e)")
        }
        sheetCheck(normalizeFormula("=SOMMA (A2; A3)") == "SUM (A2, A3)","Italian function with whitespace stores English name")
        b.date1904 = true
        sheetCheck(try FormulaParser(book:b,sheet:0).evaluate("DATE(2024,2,29)").text == "43889","DATE supports the 1904 date system")
        sheetCheck(try FormulaParser(book:b,sheet:0).evaluate("YEAR(0)").text == "1904","YEAR supports the 1904 date system")
        let undoBook = try Book.blank()
        let originalCell = try XMLElement(xmlString:"<c r=\"A1\" t=\"b\"><v>1</v></c>")
        let originalRow = XMLElement(name:"row"); originalRow.set("r","1"); originalRow.addChild(originalCell)
        try undoBook.sheetDocument(0).rootElement()?.child("sheetData")?.addChild(originalRow); undoBook.invalidate()
        let snapshots = try undoBook.snapshots(0,refs:["A1","B1","A2","B2"])
        try undoBook.applyEdits(0,[undoBook.automaticEdit(0,"A1","changed"),undoBook.automaticEdit(0,"B1","2"),undoBook.automaticEdit(0,"A2","3"),undoBook.automaticEdit(0,"B2","4")])
        let redo = try undoBook.snapshots(0,refs:snapshots.map(\.ref))
        try undoBook.restoreCells(0,snapshots)
        sheetCheck(try undoBook.cell(0,"A1")?.attr("t") == "b" && undoBook.raw(0,"B1") == "" && undoBook.raw(0,"A2") == "","Cell snapshots restore exact boolean type and remove newly created cells")
        sheetCheck(try undoBook.sheetDocument(0).rootElement()?.child("sheetData")?.kids("row").count == 1,"Cell undo removes the row created by the edit")
        try undoBook.restoreCells(0,redo)
        sheetCheck(try undoBook.raw(0,"A1") == "changed" && undoBook.raw(0,"B2") == "4","Cell snapshots redo an entire multi-cell edit")
        let formatting = try Book.blank()
        try formatting.applyEdits(0,[CellEdit(ref:"A1",value:"0.125",kind:.number),CellEdit(ref:"B1",value:"12.5",kind:.number)])
        let beforeFormat = try formatting.snapshots(0,refs:["A1"])
        for (change,name) in [(StyleChange.flag("b",true),"bold"),(.flag("i",true),"italic"),(.flag("strike",true),"strike"),(.fontSize(14),"font size"),(.foreground("#123456"),"text color"),(.fill("#FABC12"),"fill"),(.alignment("center"),"horizontal alignment"),(.wrap(true),"wrap text")] {
            try formatting.formatCells(0,refs:["A1"],change:change)
            sheetCheck(try formatting.cell(0,"A1")?.attr("s").isEmpty == false,"Formatting writes an xf index: \(name)")
        }
        let rich = formatting.style(try formatting.cell(0,"A1")!)
        sheetCheck(rich.bold && rich.italic && rich.strike && rich.wrap && rich.fontSize == 14 && rich.foreground == "#123456" && rich.background == "#FABC12" && rich.alignment == "center","Style renderer reads every font/fill/alignment property")
        let xfsBefore = try formatting.styleDocument().nodes("/styleSheet/cellXfs/xf").count
        try formatting.formatCells(0,refs:["A1"],change:.flag("b",true))
        sheetCheck(try formatting.styleDocument().nodes("/styleSheet/cellXfs/xf").count == xfsBefore,"Identical existing xf is reused")
        try formatting.formatCells(0,refs:["B1"],change:.number("0.00")); try formatting.formatCells(0,refs:["B1"],change:.decimals(1))
        sheetCheck(formatting.style(try formatting.cell(0,"B1")!).numberFormat == "0.000","Increase decimals")
        try formatting.formatCells(0,refs:["B1"],change:.decimals(-1))
        sheetCheck(formatting.style(try formatting.cell(0,"B1")!).numberFormat == "0.00","Decrease decimals")
        try formatting.formatCells(0,refs:["B1"],change:.fontDelta(1)); sheetCheck(formatting.style(try formatting.cell(0,"B1")!).fontSize == 12,"Increase font size")
        try formatting.formatCells(0,refs:["B1"],change:.fontDelta(-1)); sheetCheck(formatting.style(try formatting.cell(0,"B1")!).fontSize == 11,"Decrease font size")
        for (ref,code) in [("A1","#,##0.00 \"€\""),("C1","0.00%"),("D1","#,##0.00 \"€\""),("E1","dd/mm/yyyy"),("F1","hh:mm:ss"),("G1","@"),("H1","General")] {
            try formatting.formatCells(0,refs:[ref],change:.number(code))
            sheetCheck(formatting.style(try formatting.cell(0,ref)!).numberFormat == code,"123 format menu: \(code)")
        }
        let paintID = Int(try formatting.cell(0,"A1")!.attr("s"))!
        try formatting.formatCells(0,refs:["K1","L1","K2","L2"],change:.copy(paintID))
        sheetCheck(try ["K1","L1","K2","L2"].allSatisfy { try formatting.cell(0,$0)?.attr("s") == String(paintID) },"Paint format applies the source style to the entire next selection")
        let formattedState = try formatting.snapshots(0,refs:["A1"])
        try formatting.restoreCells(0,beforeFormat)
        sheetCheck(try !formatting.style(formatting.cell(0,"A1")!).bold && formatting.raw(0,"A1") == "0.125","Undo restores formatting and preserves the value")
        try formatting.restoreCells(0,formattedState)
        let formatFile = out.appendingPathComponent("formatting.xlsx"); try formatting.save(to:formatFile)
        let verifyFormatting = Process(); verifyFormatting.executableURL = URL(fileURLWithPath:"/usr/bin/env")
        verifyFormatting.arguments = ["python3","-c","""
        import openpyxl,sys
        w=openpyxl.load_workbook(sys.argv[1]); s=w.active; c=s['A1']
        assert c.number_format == '#,##0.00 "€"'
        assert c.font.b and c.font.i and c.font.strike and c.font.sz == 14
        assert c.font.color.type == 'rgb' and c.font.color.rgb == 'FF123456'
        assert c.fill.patternType == 'solid' and c.fill.fgColor.rgb == 'FFFABC12'
        assert c.alignment.horizontal == 'center' and c.alignment.wrap_text
        assert s['B1'].number_format == '0.00' and s['C1'].number_format == '0.00%'
        assert s['D1'].number_format == '#,##0.00 "€"' and s['E1'].number_format == 'dd/mm/yyyy'
        assert s['F1'].number_format == 'hh:mm:ss' and s['G1'].number_format == '@' and s['H1'].number_format == 'General'
        assert s['K2'].style_id == c.style_id and s['L2'].font.b
        print('OPENPYXL FORMATTING PASSED')
        """,formatFile.path]
        fflush(stdout); try verifyFormatting.run(); verifyFormatting.waitUntilExit(); sheetCheck(verifyFormatting.terminationStatus == 0,"Saved formats/fonts/colors/fills/alignment/wrap pass openpyxl")
        let series = try Book.blank()
        try series.applyEdits(0,[CellEdit(ref:"A1",value:"1",kind:.number),CellEdit(ref:"A2",value:"2",kind:.number),CellEdit(ref:"A3",value:"3",kind:.number)])
        try series.applyEdits(0,series.autofillEdits(0,source:CellRange(r1:1,c1:1,r2:3,c2:1),target:CellRange(r1:1,c1:1,r2:6,c2:1)))
        sheetCheck(try series.raw(0,"A4") == "4" && series.raw(0,"A6") == "6","Fill handle extends a vertical 1,2,3 series")
        try series.applyEdits(0,[CellEdit(ref:"C8",value:"1",kind:.number),CellEdit(ref:"D8",value:"2",kind:.number),CellEdit(ref:"E8",value:"3",kind:.number)])
        try series.applyEdits(0,series.autofillEdits(0,source:CellRange(r1:8,c1:3,r2:8,c2:5),target:CellRange(r1:8,c1:3,r2:8,c2:8)))
        sheetCheck(try series.raw(0,"F8") == "4" && series.raw(0,"H8") == "6","Fill handle extends a horizontal number series")
        try series.setCell(0,"K1","A1+$B1+C$1+$D$1",kind:.formula)
        try series.applyEdits(0,series.autofillEdits(0,source:CellRange(r1:1,c1:11,r2:1,c2:11),target:CellRange(r1:1,c1:11,r2:3,c2:11)))
        sheetCheck(try series.raw(0,"K3") == "=A3+$B3+C$1+$D$1","Vertical autofill preserves absolute references")
        try series.applyEdits(0,series.autofillEdits(0,source:CellRange(r1:1,c1:11,r2:1,c2:11),target:CellRange(r1:1,c1:11,r2:1,c2:13)))
        sheetCheck(try series.raw(0,"M1") == "=C1+$B1+E$1+$D$1","Horizontal autofill shifts columns and preserves $ references")
        try series.setCell(0,"N1","hello",kind:.text)
        try series.applyEdits(0,series.autofillEdits(0,source:CellRange(r1:1,c1:14,r2:1,c2:14),target:CellRange(r1:1,c1:14,r2:3,c2:14)))
        sheetCheck(try series.raw(0,"N3") == "hello","Fill handle copies plain text")
        let filterBook = try Book.blank()
        try filterBook.applyEdits(0,[CellEdit(ref:"A1",value:"Header",kind:.text),CellEdit(ref:"B1",value:"Amount",kind:.text),CellEdit(ref:"A2",value:"one",kind:.text),CellEdit(ref:"B2",value:"1",kind:.number),CellEdit(ref:"A3",value:"two",kind:.text),CellEdit(ref:"B3",value:"2",kind:.number)])
        try filterBook.toggleFilter(0,selection:CellRange(r1:2,c1:2,r2:2,c2:2))
        sheetCheck(try filterBook.sheetDocument(0).rootElement()?.child("autoFilter")?.attr("ref") == "A1:B3","Filter toggle creates the current contiguous data region")
        let filterFile = out.appendingPathComponent("filter.xlsx"); try filterBook.save(to:filterFile)
        sheetCheck(try Book(filterFile).sheetDocument(0).rootElement()?.child("autoFilter")?.attr("ref") == "A1:B3","Filter range survives save/reopen")
        try filterBook.toggleFilter(0,selection:CellRange())
        sheetCheck(try filterBook.sheetDocument(0).rootElement()?.child("autoFilter") == nil,"Filter toggle removes autoFilter")
        for (name,result) in [("SUM","3"),("AVERAGE","1.5"),("COUNT","2"),("MAX","2"),("MIN","1")] {
            let formula = try filterBook.aggregateFormula(0,ref:"B4",function:name)
            sheetCheck(try formula == "=\(name)(B2:B3)" && FormulaParser(book:filterBook,sheet:0).evaluate(normalizeFormula(formula)).text == result,"Σ menu \(name) inserts the contiguous numbers above")
        }
        try filterBook.setCell(0,"D8","3",kind:.number); try filterBook.setCell(0,"E8","4",kind:.number)
        sheetCheck(try filterBook.aggregateFormula(0,ref:"F8",function:"SUM") == "=SUM(D8:E8)","Σ uses contiguous numbers to the left when none are above")
        let smallHTML = try filterBook.render(0)
        sheetCheck(smallHTML.columns >= 26 && smallHTML.shownRows >= 100,"Small workbook renders at least 26 columns × 100 rows")
        sheetCheck(smallHTML.html.contains("fill-handle") && smallHTML.html.contains("selected-header"),"Grid includes one fill handle and header selection styling")
        let bare = try Book(out.appendingPathComponent("styleless.xlsx"))
        try bare.formatCells(0,refs:["A1"],change:.flag("b",true)); try bare.formatCells(0,refs:["A1"],change:.number("0.00"))
        let bareFile = out.appendingPathComponent("styleless-formatted.xlsx"); try bare.save(to:bareFile)
        let bareCheck = Process(); bareCheck.executableURL = URL(fileURLWithPath:"/usr/bin/env"); bareCheck.arguments = ["python3","-c","import openpyxl,sys; c=openpyxl.load_workbook(sys.argv[1]).active['A1']; assert c.font.b and c.number_format=='0.00' and c.value==123",bareFile.path]
        try bareCheck.run(); bareCheck.waitUntilExit(); sheetCheck(bareCheck.terminationStatus == 0,"Formatting a styleless workbook creates valid styles/rels/content types")
        do { try features.formatCells(4,refs:["A1"],change:.flag("b",true)); sheetCheck(false,"Protected formatting refused") } catch { sheetCheck(try features.sheetDocument(4).xmlData == beforeProtection,"Protected formatting is refused without partial mutation") }
        let deletion = try Book.blank()
        try deletion.applyEdits(0,[deletion.automaticEdit(0,"A2","7"),deletion.automaticEdit(0,"B2","=SUM(A2,3)")])
        let formulaBeforeDelete = try deletion.snapshots(0,refs:["B2"]), deletionHistory = SheetUndoManager(), deletionTarget = NSObject()
        try deletion.applyEdits(0,[CellEdit(ref:"B2",value:"",kind:.text)])
        deletionHistory.record(deletionTarget,name:"Delete") { _ in try! deletion.restoreCells(0,formulaBeforeDelete) }
        sheetCheck(try deletion.raw(0,"B2") == "" && deletionHistory.groupingLevel == 0,"Delete regression: empty value and exactly one closed document undo group")
        deletionHistory.undo()
        sheetCheck(try deletion.raw(0,"B2") == "=SUM(A2,3)" && deletion.value(0,2,2).text == "10" && !deletionHistory.canUndo,"Delete regression: exactly ONE undo restores the original formula and 10")
        let checks = try Book.blank()
        try checks.applyEdits(0,[CellEdit(ref:"A2",value:"\"\"",kind:.formula),CellEdit(ref:"B1",value:"19",kind:.number),CellEdit(ref:"B2",value:"19",kind:.text),CellEdit(ref:"C1",value:"TRUE()",kind:.formula)])
        let round2Cases: [(String,String)] = [
            ("ISERROR(1/0)","TRUE"),("ISERROR(\"text\"+1)","TRUE"),("ISERROR(SQRT(-1))","TRUE"),("ISERROR(POWER(-1,0.5))","TRUE"),
            ("ISERROR(1E308*1E308)","TRUE"),("ISERROR(19)","FALSE"),("ISERROR(\"#VALUE!\")","FALSE"),
            ("ISERROR(MATCH(99,B1:B2,0))","TRUE"),("ISERROR(VLOOKUP(99,B1:C2,2,FALSE))","TRUE"),("ISERROR(XLOOKUP(99,B1:B2,C1:C2))","TRUE"),
            ("ISERR(1/0)","TRUE"),("ISERR(#VALUE!)","TRUE"),("ISERR(#N/A)","FALSE"),("ISERR(0)","FALSE"),
            ("ISNA(#N/A)","TRUE"),("ISNA(MATCH(99,B1:B2,0))","TRUE"),("ISNA(1/0)","FALSE"),("ISNA(\"#N/A\")","FALSE"),
            ("ISNUMBER(19)","TRUE"),("ISNUMBER(B1)","TRUE"),("ISNUMBER(\"19\")","FALSE"),("ISNUMBER(B2)","FALSE"),("ISNUMBER(TRUE)","FALSE"),("ISNUMBER(A1)","FALSE"),("ISNUMBER(1/0)","FALSE"),
            ("ISTEXT(\"19\")","TRUE"),("ISTEXT(A2)","TRUE"),("ISTEXT(\"\")","TRUE"),("ISTEXT(A1)","FALSE"),("ISTEXT(19)","FALSE"),("ISTEXT(#N/A)","FALSE"),
            ("ISBLANK(A1)","TRUE"),("ISBLANK(A2)","FALSE"),("ISBLANK(\"\")","FALSE"),("ISBLANK(0)","FALSE"),("ISBLANK(FALSE)","FALSE"),("ISBLANK(#N/A)","FALSE"),("ISBLANK(IF(TRUE,A1,1))","FALSE"),
            ("ISLOGICAL(TRUE)","TRUE"),("ISLOGICAL(C1)","TRUE"),("ISLOGICAL(1=1)","TRUE"),("ISLOGICAL(1)","FALSE"),("ISLOGICAL(\"TRUE\")","FALSE"),("ISLOGICAL(A1)","FALSE"),("ISLOGICAL(#VALUE!)","FALSE"),
            ("EDATE(DATE(2024,1,31),1)","45351"),("EDATE(DATE(2024,1,31),1.9)","45351"),("EDATE(DATE(2024,3,31),-1.9)","45351"),
            ("EDATE(DATE(2024,2,29),12)","45716"),("EDATE(DATE(2024,2,29),-12)","44985"),("EDATE(DATE(2024,2,29),-1)","45320"),("EDATE(45322.75,1)","45351"),
            ("EOMONTH(DATE(2024,2,10),0)","45351"),("EOMONTH(DATE(2024,2,10),12)","45716"),("EOMONTH(DATE(2024,3,1),-1.9)","45351"),("EOMONTH(45322.75,0)","45322"),
            ("EDATE(-1,1)","#NUM!"),("EOMONTH(-1,1)","#NUM!"),("EDATE(\"not a date\",1)","#VALUE!"),("EOMONTH(\"not a date\",1)","#VALUE!"),
            ("EDATE(DATE(9999,12,31),1)","#NUM!"),("EOMONTH(DATE(9999,12,31),1)","#NUM!"),("EDATE(1,0)","1"),("EDATE(0,1)","31"),("EOMONTH(32,0)","59"),("EDATE(31,1)","59"),
            ("IF(\"\"=0,1,2)","2"),("IF(A1=\"\",1,2)","1"),("IF(A1=0,1,2)","1"),("IF(A2=0,1,2)","2"),("IF(B2=B1,1,2)","2"),
            ("IF(\"0\"=0,1,2)","2"),("IF(\"2\">10,1,2)","1"),("IF(2<\"10\",1,2)","1"),("IF(FALSE=0,1,2)","2"),("IF(A1=FALSE,1,2)","1"),
            ("IF(\"hello\"=\"HELLO\",1,2)","1"),("COUNTIF(B1:B2,19)","2")
        ]
        for (formula,expected) in round2Cases {
            let actual = try FormulaParser(book:checks,sheet:0).evaluate(formula).text
            sheetCheck(actual == expected,"Round 2: \(formula) = \(expected) (got \(actual))")
        }
        for name in ["ISERROR","ISERR","ISNA","ISNUMBER","ISTEXT","ISBLANK","ISLOGICAL","EDATE","EOMONTH"] {
            let pair = round2Cases.first { $0.0.hasPrefix(name+"(") }!, localized = italianFunctions.first { $0.value == name }!.key
            let input = pair.0.replacingOccurrences(of:name+"(",with:localized+"(").replacingOccurrences(of:",",with:";")
            sheetCheck(try FormulaParser(book:checks,sheet:0).evaluate(normalizeFormula("="+input)).text == pair.1,"Round 2 Italian \(localized) with semicolons = \(pair.1)")
        }
        for ref in ["B1","B2","B3","B4"] {
            sheetCheck(try FormulaParser(book:features,sheet:2).evaluate("ISERROR(\(ref))").text == "TRUE","ISERROR catches a referenced error in \(ref)")
            sheetCheck(try FormulaParser(book:features,sheet:2).evaluate(normalizeFormula("=VAL.ERRORE(\(ref))")).text == "TRUE","VAL.ERRORE catches a referenced error in \(ref)")
        }
        let originalShape = "IF(ISERROR(K477/100*$D$1*0.74-M477),0,K477/100*$D$1*0.74-M477)"
        let unknownShape = originalShape.replacingOccurrences(of:"ISERROR",with:"UNKNOWN.CHECK")
        let regression = try Book.blank()
        try regression.applyEdits(0,[CellEdit(ref:"K477",value:"not a number",kind:.text),CellEdit(ref:"D1",value:"100",kind:.number),CellEdit(ref:"M477",value:"3",kind:.number)])
        sheetCheck(try FormulaParser(book:regression,sheet:0).evaluate(originalShape).text == "0","Exact real-world IF(ISERROR(K477/100*$D$1*0.74-M477),0,...) returns 0")
        let unknowns = [unknownShape,"UNKNOWN.CHECK(1/0)","UNKNOWN.CHECK(1)","IF(TRUE,0,UNKNOWN.CHECK(1/0))","IFERROR(UNKNOWN.CHECK(1/0),0)","ISERROR(UNKNOWN.CHECK(1/0))","1/0+UNKNOWN.CHECK(1)","IF(FALSE,UNKNOWN.CHECK(#VALUE!),0)","_xlfn.UNKNOWN.CHECK(#N/A)"]
        for formula in unknowns {
            do { _ = try FormulaParser(book:regression,sheet:0).evaluate(formula); sheetCheck(false,"Unknown must throw: \(formula)") }
            catch is FormulaUnsupported { sheetCheck(true,"Whole formula falls back for unknown function: \(formula)") }
        }
        do { _ = try FormulaParser(book:regression,sheet:0).call("UNKNOWN.CHECK",[.error("#VALUE!")]); sheetCheck(false,"Unknown call with error argument must throw") }
        catch is FormulaUnsupported { sheetCheck(true,"Unknown function dispatch cannot turn an argument error into a result") }
        _ = try regression.automaticEdit(0,"P477","="+unknownShape)
        sheetCheck(true,"Syntax validation still accepts an unknown function")
        try regression.setCell(0,"P477",unknownShape,kind:.formula)
        try regression.cell(0,"P477")?.addChild(XMLElement(name:"v",stringValue:"0")); regression.invalidate()
        sheetCheck(try regression.value(0,477,16).text == "0","Unknown function in the exact shape preserves cached 0 after edits")
        sheetCheck(try regression.display(regression.cell(0,"P477")!,sheet:0) == "0","Unknown function displays its cached value")
        let regressionFile = out.appendingPathComponent("round2-cache.xlsx"); try regression.save(to:regressionFile)
        sheetCheck(try Book(regressionFile).value(0,477,16).text == "0","Unknown-function cached result survives save/reopen")
        try regression.setCell(0,"K477","10",kind:.number)
        sheetCheck(try FormulaParser(book:regression,sheet:0).evaluate(originalShape).text == "4.4","Exact ISERROR shape also calculates its non-error branch")
        let ranks = [7.0,3.5,3.5,1,2]
        try checks.applyEdits(0,ranks.enumerated().map { CellEdit(ref:"J\($0.offset+1)",value:FValue.number($0.element).text,kind:.number) }+[CellEdit(ref:"J6",value:"99",kind:.text),CellEdit(ref:"J7",value:"TRUE()",kind:.formula)])
        for (formula,expected) in [("RANK(3.5,J1:J7)","2"),("RANK(3.5,J1:J7,0)","2"),("RANK(3.5,J1:J7,1)","3"),("RANK(3.5,J1:J7,-1)","3"),("RANK(3.5,J1:J7,0.1)","3"),("RANK(2,J1:J7)","4"),("RANK(7,J1:J7,1)","5"),("RANK(99,J1:J7)","#N/A"),("RANK.EQ(3.5,J1:J7,1)","3")] {
            sheetCheck(try FormulaParser(book:checks,sheet:0).evaluate(formula).text == expected,"RANK ties/order/ignored text: \(formula) = \(expected)")
        }
        sheetCheck(try FormulaParser(book:checks,sheet:0).evaluate(normalizeFormula("=RANGO(3,5;J1:J7;1)")).text == "3","Italian RANGO ties with comma decimals and ascending order")
        try checks.applyEdits(0,[CellEdit(ref:"L1",value:"45351",kind:.number),CellEdit(ref:"L2",value:"45351",kind:.number),CellEdit(ref:"L3",value:"45292",kind:.number),CellEdit(ref:"L4",value:"45382",kind:.number)])
        var localCalendar = Calendar(identifier:.gregorian); localCalendar.timeZone = .current
        checks.calculationTime = localCalendar.date(from:DateComponents(year:2024,month:1,day:31,hour:12))!
        for formula in ["IFERROR(RANK(ROUND(EDATE(TODAY(),1),0),L1:L4,IF(AND(ISNUMBER(TODAY()),OR(TRUE,FALSE)),0,1)),0)",normalizeFormula("SE.ERRORE(RANGO(ARROTONDA(DATA.MESE(OGGI();1);0);L1:L4;SE(E(VAL.NUMERO(OGGI());O(VERO;FALSO));0;1));0)")] {
            sheetCheck(try FormulaParser(book:checks,sheet:0).evaluate(formula).text == "2","RANK/TODAY/EDATE/IFERROR/ROUND/AND/OR integration = 2")
        }
        checks.date1904 = true
        sheetCheck(try FormulaParser(book:checks,sheet:0).evaluate("EDATE(DATE(2024,1,31),1)").text == "43889","EDATE honors the workbook's 1904 date system")
        sheetCheck(try FormulaParser(book:checks,sheet:0).evaluate("EOMONTH(DATE(2024,2,10),0)").text == "43889","EOMONTH honors the workbook's 1904 date system")
        try runV13ModelTests(out)
        print("ALL MODEL TESTS PASSED")
    } catch { fputs("FAIL: model tests: \(error)\n",stderr); exit(1) }
}


/// v1.3 structural changes are checked against exact package fields, not implementation mirrors.
func runV13ModelTests(_ out: URL) throws {
    let source = out.appendingPathComponent("v13.xlsx"), book = try Book(source)
    try book.preparePackage()
    let original = try book.packageSnapshot()
    func formula(_ b: Book,_ s: Int,_ ref: String) throws -> String { try b.cell(s,ref)?.child("f")?.stringValue ?? "" }
    func root(_ b: Book) throws -> XMLElement { try b.sheetDocument(0).rootElement()! }
    func sameXML(_ b: Book,_ state: PackageSnapshot) throws -> Bool {
        for (path,bytes) in state.documents { if try b.part(path).xmlData != bytes {
            try bytes.write(to:out.appendingPathComponent("undo-before.xml")); try b.part(path).xmlData.write(to:out.appendingPathComponent("undo-after.xml")); print("XML DIFFERENCE: \(path)"); return false
        } }; return true
    }
    func expect(_ change: GridChange,_ input: String,_ output: String,context: Int = 0) throws {
        sheetCheck(try book.rewrite(input,context:context,target:0,change:change) == output,"Reference rewrite: \(input) → \(output)")
    }
    let insert = GridChange(rows:true,at:3,count:2,deleting:false),delete = GridChange(rows:true,at:3,count:2,deleting:true)
    for (input,output) in [("$B$3+B3+B$3+$B3","$B$5+B5+B$5+$B5"),("SUM(B2:B4)","SUM(B2:B6)"),("SUM(3:3)+SUM(B:B)","SUM(5:5)+SUM(B:B)"),("IF(B3=1,\"B3\",Other!B3)","IF(B5=1,\"B3\",Other!B3)"),("LOG10(B3)","LOG10(B5)"),("'Data sheet'!B3+Other!B3","'Data sheet'!B5+Other!B3")] { try expect(insert,input,output) }
    for (input,output) in [("SUM(B2:B6)","SUM(B2:B4)"),("SUM(B3:B4)","SUM(#REF!)"),("$B$3","#REF!"),("SUM(3:4)","SUM(#REF!)"),("SUM(B4:B6)","SUM(B3:B4)"),("'Data sheet'!B3","'Data sheet'!#REF!")] { try expect(delete,input,output) }
    try expect(GridChange(rows:false,at:2,count:1,deleting:false),"SUM($A:$C)+$B$3","SUM($A:$D)+$C$3")
    try expect(insert,"'Data sheet'!B3+B3","'Data sheet'!B5+B3",context:1)
    sheetCheck(fillFormula("'Data sheet'!$B2+B$2+\"B2\"",down:1,right:2) == "'Data sheet'!$B3+D$2+\"B2\"","Copy uses the structural reference scanner while preserving absolute axes and strings")
    try book.shiftGrid(0,change:insert)
    sheetCheck(try formula(book,1,"A1") == "SUM('Data sheet'!$B$2:$B$6)" && formula(book,1,"A2") == "'Data sheet'!$B$5","Cross-sheet formulas and absolutes shift workbook-wide")
    sheetCheck(try book.value(0,6,4).text == "60" && book.value(1,1,1).text == "60","Shared children and cross-sheet SUM compute after insertion")
    sheetCheck(try book.cell(0,"D2")?.child("f")?.attr("ref") == "D2:D6","Shared master ref grows")
    sheetCheck(try book.cell(0,"F12")?.child("f")?.attr("ref") == "F12:F13","Array formula ref moves without splitting the array")
    let r = try root(book)
    sheetCheck(r.child("autoFilter")?.attr("ref") == "A1:C6" && r.child("mergeCells")?.child("mergeCell")?.attr("ref") == "J2:K5","Filter and merged range grow")
    sheetCheck(r.child("dataValidations")?.child("dataValidation")?.attr("sqref") == "B2:B6" && r.child("conditionalFormatting")?.attr("sqref") == "B2:B6","Data validation and conditional formatting scopes grow")
    sheetCheck(r.child("conditionalFormatting")?.child("cfRule")?.child("formula")?.stringValue == "B2>0","Conditional formatting relative formula preserves its anchor")
    sheetCheck(r.child("sheetData")?.kids("row").first(where:{ $0.attr("r") == "5" })?.attr("ht") == "30" && r.child("sheetData")?.kids("row").first(where:{ $0.attr("r") == "6" })?.attr("hidden") == "1","Row height/hidden move with rows")
    let tablePath = try book.linked(book.sheets[0].path,type:"table")[0].1
    sheetCheck(try book.part(tablePath).rootElement()?.attr("ref") == "A1:C6","Table range grows")
    let names = try book.part("xl/workbook.xml").rootElement()!.child("definedNames")!.kids("definedName")
    sheetCheck(names.first(where:{ $0.attr("name") == "SalesRange" })?.stringValue == "'Data sheet'!$B$2:$B$6","Global defined name shifts")
    sheetCheck(names.first(where:{ $0.attr("name") == "_xlnm.Print_Area" })?.stringValue == "'Data sheet'!$A$1:$C$6","Print area shifts")
    sheetCheck(r.child("sheetData")?.kids("row").first(where:{ $0.attr("r") == "5" })?.attr("spans") == "1:4","Row spans are rebuilt from moved cells")
    sheetCheck(r.child("hyperlinks")?.child("hyperlink")?.attr("ref") == "C2","Hyperlink cell scope preserves its row above insertion")
    let hyperlinkRels = try book.part(relationshipsPart(book.sheets[0].path)).rootElement()!.kids("Relationship")
    sheetCheck(hyperlinkRels.first(where:{ $0.attr("Type").hasSuffix("/hyperlink") })?.attr("Target") == "#'Data sheet'!B5","Workbook-local hyperlink relationship follows its target")
    let shiftedCharts = try book.readCharts(0)
    let initialCharts = try Book(source).readCharts(0)
    sheetCheck(shiftedCharts[1].frame.minY == initialCharts[1].frame.minY+54,"Existing drawing row anchors move with inserted rows")
    sheetCheck(shiftedCharts.count == 10 && shiftedCharts[0].series[0].values.count == 5,"Chart series references grow with inserted rows")
    try book.save(to:out.appendingPathComponent("v13-rows.xlsx"))
    try book.restorePackage(original)
    sheetCheck(try sameXML(book,original),"Structural undo restores every original XML byte")
    try book.shiftGrid(0,change:GridChange(rows:true,at:3,count:1,deleting:true))
    sheetCheck(try formula(book,1,"A2") == "'Data sheet'!#REF!" && book.value(1,2,1).text == "#REF!","Deleted cross-sheet cell computes #REF!, including with a cached formula")
    sheetCheck(try book.value(0,3,4).text == "60" && book.cell(0,"D2")?.child("f")?.attr("ref") == "D2:D3","Shared formulas shrink and keep calculating")
    sheetCheck(try root(book).child("mergeCells")?.child("mergeCell")?.attr("ref") == "J2:K2","Merge shrinks on deletion")
    try book.restorePackage(original)
    try book.shiftGrid(0,change:GridChange(rows:true,at:2,count:1,deleting:true))
    sheetCheck(try book.value(0,2,4).text == "40" && formula(book,0,"D2") == "B2*2","Deleting a shared master promotes a surviving child")
    sheetCheck(try root(book).child("conditionalFormatting")?.child("cfRule")?.child("formula")?.stringValue == "B2>0","Deleting a conditional format anchor rebases the relative formula")
    try book.restorePackage(original)
    try book.shiftGrid(0,change:GridChange(rows:false,at:2,count:1,deleting:false))
    sheetCheck(try book.part(tablePath).rootElement()?.attr("ref") == "A1:D4" && book.part(tablePath).rootElement()?.child("tableColumns")?.attr("count") == "4","Table column insertion updates column definitions and header cell")
    sheetCheck(try book.raw(0,"B1").hasPrefix(L("Colonna","Column")) && book.value(1,1,1).text == "60","Inserted table header is unique; cross-sheet values stay stable")
    sheetCheck(try root(book).child("cols")?.kids("col").contains(where:{ $0.attr("min") == "4" && $0.attr("width") == "21" }) == true,"Column widths shift")
    try book.save(to:out.appendingPathComponent("v13-cols.xlsx"))
    try book.restorePackage(original)
    try book.shiftGrid(0,change:GridChange(rows:false,at:2,count:1,deleting:true))
    sheetCheck(try book.part(tablePath).rootElement()?.child("tableColumns")?.attr("count") == "2" && formula(book,1,"A1") == "SUM('Data sheet'!#REF!)","Column deletion removes table columns and invalidates fully deleted references")
    try book.restorePackage(original)
    for change in [GridChange(rows:true,at:11,count:1,deleting:false),GridChange(rows:true,at:10,count:1,deleting:true)] {
        do { try book.shiftGrid(0,change:change); sheetCheck(false,"Refuse partial array edits") } catch { sheetCheck(try sameXML(book,original),"Refused partial array operation leaves every XML byte unchanged") }
    }
    // Insert at every data boundary, or delete pre-existing empty rows/columns.
    let invariant = try Book.blank()
    try invariant.applyEdits(0,[CellEdit(ref:"A1",value:"7",kind:.number),CellEdit(ref:"C3",value:"11",kind:.number),CellEdit(ref:"E5",value:"13",kind:.number)])
    let expressions = ["$A$1+C3+$E5","SUM(A1:E5)","SUM(A:A)","SUM(3:3)","IF(C3>0,SUM($A$1:$E$5),0)","SUM(A1:C3)*2","SUM($1:$5)"]
    for (i,f) in expressions.enumerated() { try invariant.setCell(0,"G\(i+1)",f,kind:.formula) }
    try invariant.preparePackage(); let base = try invariant.packageSnapshot()
    let values = try expressions.indices.map { try invariant.value(0,$0+1,7).text }
    var passes = 0
    for rows in [true,false] { for at in 1...9 { for deleting in [false,true] {
        if deleting && (rows ? at<=7 : [1,3,5,7].contains(at)) { continue }
        try invariant.restorePackage(base)
        let change = GridChange(rows:rows,at:at,count:1,deleting:deleting)
        try invariant.shiftGrid(0,change:change)
        for i in expressions.indices {
            let row = rows ? change.point(i+1)! : i+1,col = rows ? 7 : change.point(7)!
            sheetCheck(try invariant.value(0,row,col).text == values[i],"INVARIANT \(rows ? "row" : "column") \(deleting ? "delete empty" : "insert") \(at), formula \(i+1)")
        }; passes += 1
    } } }
    sheetCheck(passes == 25,"Invariant covers every insertion boundary and existing empty axes")
    try book.restorePackage(original)
    let models = try book.readCharts(0)
    sheetCheck(models.map(\.kind) == ["column","column","column","bar","line","pie","doughnut","area","scatter","unsupported"],"Parse existing chart families and unsupported placeholder")
    func marks(_ svg: String,_ kind: String) -> Int { svg.components(separatedBy:"data-mark='\(kind)'").count-1 }
    sheetCheck(marks(chartSVG(models[0]),"bar") == 6 && marks(chartSVG(models[4]),"point") == 6 && marks(chartSVG(models[8]),"point") == 3,"SVG has exact bar, line-marker and scatter-point counts")
    sheetCheck(marks(chartSVG(models[5]),"slice") == 3 && marks(chartSVG(models[6]),"slice") == 3 && marks(chartSVG(models[7]),"area") == 2,"SVG pie/doughnut slices and area polygons")
    sheetCheck(chartSVG(models[9]).contains(L("Anteprima non disponibile","Preview unavailable")),"Unsupported chart keeps title and a localized placeholder")
    let fallback = try XMLElement(xmlString:"<val><numRef><f>UnsupportedName</f><numCache><ptCount val='3'/><pt idx='0'><v>42</v></pt><pt idx='2'><v>7</v></pt></numCache></numRef></val>")
    let cached = try book.chartValues(fallback,sheet:0,numeric:true)
    sheetCheck(cached.count == 3 && cached[0].0.text == "42" && cached[1].0.text == "" && cached[2].0.text == "7","Unsupported chart references use sparse caches with their original point positions")
    fallback.child("numRef")?.child("f")?.stringValue = "'Data sheet'!#REF!"
    sheetCheck(try book.chartValues(fallback,sheet:0,numeric:true).isEmpty,"Deleted chart reference never resurrects cached points")
    let chartRange = CellRange(r1:1,c1:1,r2:4,c2:3)
    for kind in ["column","bar","line","pie"] { _ = try book.createChart(0,range:chartRange,kind:kind) }
    sheetCheck(try book.readCharts(0).count == 14,"Create all four chart types alongside existing charts")
    let created = try book.readCharts(0).last!,chartState = try book.packageSnapshot()
    try book.setChartFrame(0,id:created.id,frame:CGRect(x:350,y:160,width:480,height:290))
    sheetCheck(try book.readCharts(0).last!.frame == CGRect(x:350,y:160,width:480,height:290),"Move and resize chart using EMU anchors")
    try book.save(to:out.appendingPathComponent("v13-charts.xlsx"))
    let reopened = try Book(out.appendingPathComponent("v13-charts.xlsx"))
    sheetCheck(try reopened.readCharts(0).count == 14 && reopened.readCharts(0).last!.frame == CGRect(x:350,y:160,width:480,height:290),"Chart anchors and every existing/new chart survive save/reopen")
    try book.restorePackage(chartState); sheetCheck(try sameXML(book,chartState),"Move/resize undo restores exact XML")
    let deletion = try book.packageSnapshot(); try book.deleteChart(0,id:created.id)
    sheetCheck(try book.readCharts(0).count == 13,"Delete selected chart")
    try book.save(to:out.appendingPathComponent("v13-chart-deleted.xlsx"))
    sheetCheck(try Book(out.appendingPathComponent("v13-chart-deleted.xlsx")).readCharts(0).count == 13,"Chart deletion survives save/reopen")
    try book.restorePackage(deletion); sheetCheck(try sameXML(book,deletion),"Chart deletion undo restores exact XML and relationships")
    try book.setCell(0,"B2","99",kind:.number)
    sheetCheck(try book.readCharts(0)[0].series[0].values[0] == 99,"Existing chart reads live edited cells")
    try book.save(to:out.appendingPathComponent("v13-live.xlsx"))
    try book.restorePackage(original); sheetCheck(try sameXML(book,original) && book.readCharts(0).count == 10,"Chart creation undo restores exact original package XML")
    let single = try Book.blank(); try single.applyEdits(0,[CellEdit(ref:"A1",value:"Name",kind:.text),CellEdit(ref:"B1",value:"Value",kind:.text),CellEdit(ref:"A2",value:"Only",kind:.text),CellEdit(ref:"B2",value:"3",kind:.number)])
    try single.preparePackage(); let empty = try single.packageSnapshot()
    let id = try single.createChart(0,range:CellRange(r1:1,c1:1,r2:2,c2:2),kind:"pie")
    sheetCheck(marks(chartSVG(try single.readCharts(0)[0]),"slice") == 1,"Single-slice pie uses a complete circle")
    let anchorBefore = try single.packageSnapshot(),chartBefore = try single.readCharts(0)[0]
    let (_,_,anchor) = try single.chartAnchor(0,id:id); anchor.removeAttribute(forName:"editAs")
    try single.shiftGrid(0,change:GridChange(rows:true,at:3,count:1,deleting:false))
    sheetCheck(try single.readCharts(0)[0].frame.height == chartBefore.frame.height+27,"Two-cell drawing to-anchor expands across inserted rows")
    try single.restorePackage(anchorBefore)
    try single.save(to:out.appendingPathComponent("v13-new.xlsx")); try single.deleteChart(0,id:id)
    try single.save(to:out.appendingPathComponent("v13-empty.xlsx"))
    sheetCheck(try Book(out.appendingPathComponent("v13-empty.xlsx")).readCharts(0).isEmpty,"Deleting the last chart removes drawing parts and links cleanly")
    try single.restorePackage(empty); sheetCheck(try sameXML(single,empty),"Undo first chart creation restores empty workbook XML")
    let unsafe = try Book(source); try unsafe.preparePackage()
    try unsafe.cell(1,"A1")?.child("f")?.stringValue = "SUM(SalesTable[Sales])"
    let unsafeState = try unsafe.packageSnapshot()
    do { try unsafe.shiftGrid(0,change:insert); sheetCheck(false,"Refuse structured references") } catch { sheetCheck(try sameXML(unsafe,unsafeState),"Structured references refused without mutation") }
    unsafe.documents["xl/pivotTables/pivotTable1.xml"] = try XMLDocument(xmlString:"<pivotTableDefinition/>",options:[])
    do { try unsafe.shiftGrid(0,change:insert); sheetCheck(false,"Refuse pivot tables") } catch { sheetCheck(true,"Pivot tables refused explicitly") }
    print("V1.3 MODEL TESTS PASSED")
}

func scheduleSheetTests(_ app: App, fixture: URL) {
    Task { @MainActor in
        func pause(_ seconds: Double = 0.08) async { try? await Task.sleep(nanoseconds:UInt64(seconds*1e9)) }
        @MainActor func js(_ script: String) async -> Any? {
            await withCheckedContinuation { continuation in
                app.web.evaluateJavaScript(script) { value,error in
                    if let error { fputs("JS ERROR: \(error)\n",stderr); exit(1) }
                    continuation.resume(returning:value)
                }
            }
        }
        @MainActor func wait(_ description: String, _ condition: () -> Bool) async {
            let start = Date()
            while !condition() { if Date().timeIntervalSince(start)>30 { sheetCheck(false,"Timeout: \(description)") }; await pause() }
            await pause()
        }
        await pause(0.6)
        runSheetModelTests(fixture)
        app.book?.dirty = false
        let launchDeadline = Date().addingTimeInterval(10)
        while app.window == nil || NSApp.mainMenu == nil {
            if Date() > launchDeadline { fflush(stdout); fputs("UI TESTS UNAVAILABLE: macOS GUI application launch is blocked in this environment (no window/menu). Run --smoke-test outside the sandbox.\n",stderr); fflush(stdout); exit(2) }
            await pause()
        }
        let fileMenu = NSApp.mainMenu!.items.first { $0.title == "File" }!.submenu!
        let newItem = fileMenu.items.first { $0.action == #selector(App.newWorkbook) }!
        sheetCheck(newItem.keyEquivalent == "n","New workbook has ⌘N binding")
        fileMenu.performActionForItem(at:fileMenu.index(of:newItem))
        await wait("new window document",{ app.book?.untitled == true && !app.web.isLoading })
        let beforeTyping = app.committedActions
        _ = await js("select(at(2,1));document.dispatchEvent(new KeyboardEvent('keydown',{key:'7'}));document.getElementById('editor').dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))")
        await wait("typing commit",{ (try? app.book?.raw(0,"A2")) == "7" && app.selectedRef == "A3" })
        sheetCheck(app.committedActions == beforeTyping+1 && app.editUndo.groupingLevel == 0,"In-cell typing commits once; Return moves down; undo group is closed")
        _ = await js("document.dispatchEvent(new KeyboardEvent('keydown',{key:'F2'}));document.getElementById('editor').value='cancel me';document.getElementById('editor').dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))")
        sheetCheck(try! app.book!.raw(0,"A3") == "","Escape cancels in-cell editing")
        _ = await js("select(at(2,2));formula.value='=SOMMA(A2;3)';formula.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))")
        await wait("formula bar edit",{ (try? app.book?.value(0,2,2).text) == "10" })
        sheetCheck(true,"Formula bar commits Italian formula")
        _ = await js("select(at(2,2));document.dispatchEvent(new KeyboardEvent('keydown',{key:'Delete'}))")
        await wait("delete",{ (try? app.book?.raw(0,"B2")) == "" })
        sheetCheck(true,"Delete clears selected cells")
        sheetCheck(app.editUndo.groupingLevel == 0,"Delete leaves no implicit event group")
        let deleteActions = app.committedActions
        app.undoEdit()
        sheetCheck((try! app.book!.raw(0,"B2")) == "=SUM(A2,3)" && (try! app.book!.value(0,2,2).text) == "10","Exactly ONE Delete undo restores formula and value immediately")
        sheetCheck(app.committedActions == deleteActions,"Undo does not register a new edit")
        sheetCheck(true,"Delete is one undo step")
        _ = await js("select(at(3,1));post({action:'commit',ref:'A3',value:'3'})")
        await wait("A3 edit",{ (try? app.book?.raw(0,"A3")) == "3" })
        _ = await js("select(at(2,1));document.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',shiftKey:true}))")
        await wait("range selection",{ app.selectedRange.count == 2 })
        let stats = await js("[+stats.dataset.sum,+stats.dataset.average,+stats.dataset.count]") as? [Double]
        sheetCheck(stats == [10,5,2],"Range stats sum=10, average=5, count=2")
        _ = await js("document.dispatchEvent(new KeyboardEvent('keydown',{key:'c',metaKey:true}))")
        await wait("copy range",{ NSPasteboard.general.string(forType:.string) == "7\n3" })
        sheetCheck(true,"Copy range writes TSV and plain pasteboard text")
        let beforePaste = app.committedActions
        NSPasteboard.general.clearContents(); NSPasteboard.general.setString("11\t12\n21\t22",forType:.string)
        _ = await js("select(at(5,1));document.dispatchEvent(new KeyboardEvent('keydown',{key:'v',metaKey:true}))")
        await wait("paste block",{ (try? app.book?.raw(0,"B6")) == "22" })
        sheetCheck((try! app.book!.raw(0,"A5")) == "11" && app.committedActions == beforePaste+1 && app.editUndo.groupingLevel == 0,"Paste TSV block commits exactly once in a closed undo group")
        app.undoEdit()
        await wait("paste undo",{ (try? app.book?.raw(0,"A5")) == "" })
        sheetCheck(try! app.book!.raw(0,"B6") == "","Entire block paste is one undo step")
        _ = await js("select(at(5,1));select(at(6,2),false,true)")
        await wait("fill selection",{ app.selectedRange.count == 4 })
        try! app.pasteText("9")
        sheetCheck(try! app.book!.raw(0,"B6") == "9","Single value fills a multi-cell selection")
        let beforeCut = app.committedActions
        _ = await js("select(at(5,1));select(at(6,2),false,true);post({action:'cut'})")
        await wait("cut clear",{ (try? app.book?.raw(0,"A5")) == "" })
        sheetCheck(NSPasteboard.general.string(forType:.string) == "9\t9\n9\t9" && app.committedActions == beforeCut+1 && app.editUndo.groupingLevel == 0,"Cut copies full range and commits exactly one closed undo group")
        app.undoEdit()
        sheetCheck(try! app.book!.raw(0,"B6") == "9","Cut clears as one undo step")
        // Start an independent table without formulas for sort and the small additions.
        app.book?.dirty = false; app.newWorkbook()
        await wait("sort workbook",{ !app.web.isLoading })
        _ = await js("select(at(1,1))")
        await wait("A1 selection",{ app.selectedRef == "A1" })
        try! app.pasteText("Name\tAmount\nbeta\t20\nAlpha\t10\ngamma\t30")
        _ = await js("select(at(2,1))")
        await wait("sort selection",{ app.selectedRef == "A2" })
        let beforeSort = app.committedActions
        app.sortAscending()
        sheetCheck(try! app.book!.raw(0,"A2") == "Alpha","Sort A → Z, case insensitive")
        sheetCheck(app.committedActions == beforeSort+1 && app.editUndo.groupingLevel == 0,"Sort A → Z commits exactly one closed undo group")
        app.sortDescending()
        sheetCheck(try! app.book!.raw(0,"A2") == "gamma","Sort Z → A")
        app.undoEdit()
        sheetCheck(try! app.book!.raw(0,"A2") == "Alpha","Sort is one undo step")
        _ = await js("select(at(2,2));select(at(4,2),false,true)")
        await wait("fill down range",{ app.selectedRange.count == 3 })
        let beforeFillDown = app.committedActions
        app.fillDown()
        sheetCheck(try! app.book!.raw(0,"B4") == "10","Addition: Fill Down copies first row in selection")
        sheetCheck(app.committedActions == beforeFillDown+1 && app.editUndo.groupingLevel == 0,"Fill Down commits exactly one closed undo group")
        app.undoEdit(); sheetCheck((try! app.book!.raw(0,"B3")) == "20" && (try! app.book!.raw(0,"B4")) == "30","Exactly ONE Fill Down Undo restores the range")
        app.redoEdit()
        let before = await js("firstFrozen") as? Bool
        app.toggleFreeze(); await pause()
        let after = await js("firstFrozen") as? Bool
        sheetCheck(before == false && after == true,"Addition: freeze first row toggle")
        _ = await js("autoWidth(1)")
        sheetCheck((await js("parseFloat(document.querySelectorAll('colgroup col')[1].style.width)") as? Double ?? 0) >= 48,"Addition: column auto-width")
        let out = fixture.deletingLastPathComponent(), target = out.appendingPathComponent("ui-new.xlsx")
        try! app.persist(to:target)
        sheetCheck(try! Book(target).raw(0,"A2") == "Alpha","⌘N workbook saves and reopens with Book")
        let python = Process(); python.executableURL = URL(fileURLWithPath:"/usr/bin/env"); python.arguments = ["python3","-c","import openpyxl,sys; assert openpyxl.load_workbook(sys.argv[1]).active['A2'].value=='Alpha'",target.path]
        try! python.run(); python.waitUntilExit(); sheetCheck(python.terminationStatus == 0,"⌘N saved workbook reopens with openpyxl")
        try! app.commitEdits([app.book!.automaticEdit(0,"B2","99")])
        sheetCheck(app.saveFile(),"Save in place action succeeds")
        sheetCheck(try! Book(target).raw(0,"B2") == "99","Save in place then reopen")
        app.undoEdit()
        sheetCheck(app.book!.dirty && (try! app.book!.raw(0,"B2")) == "10","Undo after save marks the document unsaved")
        app.redoEdit()
        sheetCheck(!app.book!.dirty && (try! app.book!.raw(0,"B2")) == "99","Redo to the saved state clears the unsaved flag")
        try! app.commitEdits([app.book!.automaticEdit(0,"B3","=1+1")])
        app.sortAscending()
        sheetCheck(app.lastEditError.contains(L("formule","formulas")),"UI refuses sorting region with formulas")
        _ = await js("select(at(2,1));formula.value='=UNKNOWN(1+)';formula.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))")
        await wait("invalid syntax",{ app.lastEditError.contains(L("Formula non valida","Invalid formula")) })
        sheetCheck(try! app.book!.raw(0,"A2") == "Alpha","Invalid formula reports error without altering cell")
        _ = await js("scrollTo(0,600);filters={2:'99'};window.savedScroll=scrollY;window.savedFilters=JSON.stringify(filters)")
        try! app.commitEdits([app.book!.automaticEdit(0,"A4","changed")])
        await pause()
        sheetCheck(await js("scrollY===savedScroll&&JSON.stringify(filters)===savedFilters") as? Bool == true,"Cell edit preserves scroll position and active filters")
        app.book?.dirty = false; app.newWorkbook()
        await wait("format toolbar workbook",{ !app.web.isLoading })
        _ = await js("select(at(1,1))")
        await wait("format selection",{ app.selectedRef == "A1" })
        try! app.commitEdits([app.book!.automaticEdit(0,"A1","12.5")]); await pause()
        @MainActor func tap(_ id: String) { (app.formatControls[id] as! NSButton).performClick(nil) }
        @MainActor func choose(_ id: String, _ command: String) {
            let popup = app.formatControls[id] as! NSPopUpButton
            let item = popup.menu!.items.first { ($0.representedObject as? String) == command }!
            sheetCheck(item.action == #selector(App.formatMenu(_:)),"Native \(id) menu routes \(command)")
            app.formatMenu(item)
        }
        let menuTitles = NSApp.mainMenu!.items.dropFirst().map(\.title)
        sheetCheck(menuTitles == ["File",L("Modifica","Edit"),L("Visualizza","View"),L("Inserisci","Insert"),L("Formato","Format"),L("Dati","Data"),L("Strumenti","Tools"),L("Finestra","Window"),L("Guida","Help")],"Native menu order matches Sheets")
        sheetCheck(NSApp.helpMenu?.items.contains { $0.action == #selector(App.openHelp) } == true,"Help routes to the repository")
        sheetCheck((app.formatControls["print"] as? NSButton)?.action == #selector(App.printSheet),"Toolbar print uses the existing tested PDF path")
        sheetCheck(await js("document.querySelector('thead th[data-c=\"1\"]').classList.contains('selected-header')&&rowMap.get(1).querySelector('th').classList.contains('selected-header')&&getComputedStyle(document.querySelector('thead th[data-c=\"1\"]')).fontWeight==='700'") as? Bool == true,"Selected column and row headers are blue and bold")
        sheetCheck(await js("document.getElementById('fill-handle').offsetWidth===8&&getComputedStyle(selectionBorder).borderTopColor==='rgb(26, 115, 232)'") as? Bool == true,"Selection has a blue border and a square fill handle")
        sheetCheck((app.formatControls["alignment"] as? NSPopUpButton)?.image?.accessibilityDescription == L("Allinea a destra","Align Right"),"Alignment icon reflects automatic numeric right alignment")
        tap("bold"); sheetCheck(app.selectedStyle.bold && app.editUndo.groupingLevel == 0,"Native Bold is one closed undo group")
        tap("undo"); await wait("toolbar single undo",{ !app.selectedStyle.bold })
        sheetCheck(!app.selectedStyle.bold,"ONE toolbar Undo restores the format")
        tap("redo"); await wait("toolbar single redo",{ app.selectedStyle.bold })
        sheetCheck(app.selectedStyle.bold,"ONE toolbar Redo reapplies the format")
        tap("italic"); sheetCheck(app.selectedStyle.italic,"Native Italic")
        tap("strike"); sheetCheck(app.selectedStyle.strike,"Native Strikethrough")
        tap("fontMore"); sheetCheck(app.selectedStyle.fontSize == 12,"Native font size +")
        tap("fontLess"); sheetCheck(app.selectedStyle.fontSize == 11,"Native font size −")
        let fontSize = app.formatControls["fontSize"] as! NSTextField; fontSize.stringValue = "14"; app.formatAction(fontSize)
        sheetCheck(app.selectedStyle.fontSize == 14,"Native font size field")
        tap("percent"); sheetCheck(app.selectedStyle.numberFormat == "0.00%","Native %")
        tap("decMore"); sheetCheck(app.selectedStyle.numberFormat == "0.000%","Native increase decimals")
        tap("decLess"); sheetCheck(app.selectedStyle.numberFormat == "0.00%","Native decrease decimals")
        tap("currency"); sheetCheck(app.selectedStyle.numberFormat == "#,##0.00 \"€\"","Native €")
        let fg = app.formatControls["foreground"] as! NSColorWell; fg.color = app.colorFromHex("#123456"); app.formatAction(fg)
        let bg = app.formatControls["fill"] as! NSColorWell; bg.color = app.colorFromHex("#FABC12"); app.formatAction(bg)
        sheetCheck(app.selectedStyle.foreground == "#123456" && app.selectedStyle.background == "#FABC12","Native text and fill colors")
        choose("alignment","align:center"); sheetCheck(app.selectedStyle.alignment == "center","Native horizontal alignment")
        if let foreground = fg as? ToolbarColorWell, let fill = bg as? ToolbarColorWell {
            sheetCheck(foreground.isTextGlyph && !fill.isTextGlyph,"Toolbar uses an A glyph for text color and a bucket for fill")
            sheetCheck(foreground.colorBarRect.height == 3 && fill.colorBarRect.height == 3 && app.hexColor(foreground.color) == "#123456" && app.hexColor(fill.color) == "#FABC12","Both toolbar color bars track the selected cell's current colors")
            sheetCheck(foreground.toolTip == L("Colore testo","Text color") && fill.toolTip == L("Colore riempimento","Fill color"),"Color buttons retain localized tooltips")
        } else { sheetCheck(false,"Toolbar color controls keep native color wells") }
        for (alignment,description) in [("left",L("Allinea a sinistra","Align Left")),("right",L("Allinea a destra","Align Right")),("center",L("Allinea al centro","Align Center"))] {
            choose("alignment","align:"+alignment)
            let popup = app.formatControls["alignment"] as! NSPopUpButton
            sheetCheck(popup.item(at:0)?.title == "" && popup.image?.accessibilityDescription == description && popup.imagePosition == .imageOnly,"Alignment toolbar shows the current \(alignment) icon without a text label")
            sheetCheck(popup.toolTip == L("Allineamento orizzontale","Horizontal alignment"),"Alignment icon retains its tooltip")
        }
        tap("wrap"); sheetCheck(app.selectedStyle.wrap,"Native wrap toggle")
        await pause()
        sheetCheck(await js("at(1,1).style.whiteSpace==='normal'&&at(1,1).style.textAlign==='center'&&at(1,1).style.fontSize==='14pt'") as? Bool == true,"Formatting renders immediately in the existing grid")
        let paintStyle = try! app.book!.cell(0,"A1")!.attr("s"), beforePaint = app.committedActions
        tap("paint"); await pause()
        _ = await js("select(at(3,3));select(at(4,4),false,true);document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}))")
        await wait("paint selection",{ (try? app.book!.cell(0,"D4")?.attr("s")) == paintStyle })
        sheetCheck(app.committedActions == beforePaint+1,"Paint format is ONE action for the entire next range")
        app.undoEdit()
        sheetCheck((try! app.book!.cell(0,"C3")) == nil && (try! app.book!.cell(0,"D4")) == nil,"ONE paint Undo restores the complete range")
        app.redoEdit()
        sheetCheck(try! app.book!.cell(0,"D4")?.attr("s") == paintStyle,"ONE paint Redo restores the style")
        for (n,code) in ["General","0.00","0.00%","#,##0.00 \"€\"","dd/mm/yyyy","hh:mm:ss","@"].enumerated() {
            _ = await js("select(at(10,\(n+5)))")
            await wait("123 selection",{ app.selectedRange.r1 == 10 && app.selectedRange.c1 == n+5 })
            choose("numberFormat","number:"+code)
            sheetCheck(app.selectedStyle.numberFormat == code,"Native 123 applies \(code)")
        }
        app.zoomLabel.selectItem(withTitle:"125%"); app.zoomChoice(app.zoomLabel)
        sheetCheck(app.zoom == 1.25 && app.web.pageZoom == 1.25,"Native zoom popup")
        app.zoomLabel.selectItem(withTitle:"100%"); app.zoomChoice(app.zoomLabel)
        sheetCheck(app.zoom == 1,"Native zoom returns to 100%")
        _ = await js("nameBox.value='B12';nameBox.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))")
        await wait("name box jump",{ app.selectedRef == "B12" })
        sheetCheck(true,"Name box jumps to B12")
        _ = await js("select(at(1,1))")
        await wait("duplicate Delete selection",{ app.selectedRef == "A1" })
        let beforeDelete = app.committedActions
        _ = await js("document.dispatchEvent(new KeyboardEvent('keydown',{key:'Delete'}));document.dispatchEvent(new KeyboardEvent('keydown',{key:'Delete'}))")
        await wait("duplicate Delete",{ (try? app.book!.raw(0,"A1")) == "" })
        sheetCheck(app.committedActions == beforeDelete+1 && app.editUndo.groupingLevel == 0,"Duplicate/in-flight Delete cannot create an extra undo step")
        _ = await js("document.dispatchEvent(new KeyboardEvent('keydown',{key:'Delete'}))")
        await pause()
        sheetCheck(app.committedActions == beforeDelete+1,"Delete on an empty cell creates no empty undo step")
        app.undoEdit()
        sheetCheck((try! app.book!.raw(0,"A1")) == "12.5","Exactly ONE Undo restores the value after duplicate/no-op Delete")
        try! app.commitEdits([app.book!.automaticEdit(0,"A1","1"),app.book!.automaticEdit(0,"A2","2"),app.book!.automaticEdit(0,"A3","3")]); await pause()
        _ = await js("select(at(1,1));select(at(3,1),false,true)")
        await wait("vertical autofill source",{ app.selectedRange.r2 == 3 })
        let beforeFill = app.committedActions
        _ = await js("handle.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));at(6,1).dispatchEvent(new MouseEvent('mouseover',{bubbles:true}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}))")
        await wait("vertical fill handle",{ (try? app.book!.raw(0,"A6")) == "6" && app.selectedRange.r2 == 6 })
        sheetCheck(app.committedActions == beforeFill+1,"Fill handle extends 1,2,3 as ONE action")
        app.undoEdit()
        sheetCheck((try! app.book!.raw(0,"A4")) == "" && (try! app.book!.raw(0,"A6")) == "","Exactly ONE Undo restores every autofilled cell")
        app.redoEdit(); sheetCheck((try! app.book!.raw(0,"A6")) == "6","Exactly ONE autofill Redo")
        try! app.commitEdits([app.book!.automaticEdit(0,"D8","10"),app.book!.automaticEdit(0,"E8","20"),app.book!.automaticEdit(0,"F8","30")]); await pause()
        _ = await js("select(at(8,4));select(at(8,6),false,true)")
        await wait("horizontal autofill source",{ app.selectedRange.c2 == 6 })
        _ = await js("handle.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));at(8,9).dispatchEvent(new MouseEvent('mouseover',{bubbles:true}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}))")
        await wait("horizontal fill handle",{ (try? app.book!.raw(0,"I8")) == "60" })
        sheetCheck(true,"Fill handle extends a series to the right")
        try! app.commitEdits([app.book!.automaticEdit(0,"M12","=A1+$B1+C$1+$D$1")]); await pause()
        _ = await js("select(at(12,13))")
        await wait("formula autofill source",{ app.selectedRef == "M12" })
        _ = await js("handle.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));at(12,15).dispatchEvent(new MouseEvent('mouseover',{bubbles:true}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}))")
        await wait("formula fill handle",{ (try? app.book!.raw(0,"O12")) == "=C1+$B1+E$1+$D$1" })
        sheetCheck(true,"Fill handle shifts relative references while preserving $ absolutes")
        let formattedGUI = fixture.deletingLastPathComponent().appendingPathComponent("ui-formatting.xlsx")
        try! app.persist(to:formattedGUI)
        let verifyUI = Process(); verifyUI.executableURL = URL(fileURLWithPath:"/usr/bin/env"); verifyUI.arguments = ["python3","-c","import openpyxl,sys; c=openpyxl.load_workbook(sys.argv[1]).active['A1']; assert c.font.b and c.font.i and c.font.strike and c.font.sz==14; assert c.font.color.rgb=='FF123456'; assert c.fill.fgColor.rgb=='FFFABC12'; assert c.alignment.horizontal=='center' and c.alignment.wrap_text; assert c.number_format=='#,##0.00 '+chr(34)+'€'+chr(34)",formattedGUI.path]
        try! verifyUI.run(); verifyUI.waitUntilExit(); sheetCheck(verifyUI.terminationStatus == 0,"Native toolbar formatting survives save and openpyxl")
        app.book?.dirty = false; app.newWorkbook(); await wait("filter/Sigma workbook",{ !app.web.isLoading })
        _ = await js("select(at(1,1))"); await wait("filter table origin",{ app.selectedRef == "A1" })
        try! app.pasteText("Name\tAmount\none\t1\ntwo\t2"); await pause()
        _ = await js("select(at(2,2))"); await wait("filter selection",{ app.selectedRef == "B2" })
        tap("filter"); await pause()
        sheetCheck(try! app.book!.sheetDocument(0).rootElement()?.child("autoFilter")?.attr("ref") == "A1:B3","Native filter toggle finds the data region")
        sheetCheck(await js("document.querySelectorAll('td.fh').length") as? Int == 2,"Filter headers render without reloading the page")
        tap("filter"); await pause()
        sheetCheck(try! app.book!.sheetDocument(0).rootElement()?.child("autoFilter") == nil,"Native filter toggle removes the filter")
        app.undoEdit(); await pause()
        sheetCheck(await js("document.querySelectorAll('td.fh').length") as? Int == 2,"ONE Undo restores filter XML and header controls")
        _ = await js("select(at(4,2))"); await wait("Σ target",{ app.selectedRef == "B4" })
        for (function,expected) in [("SUM","3"),("AVERAGE","1.5"),("COUNT","2"),("MAX","2"),("MIN","1")] {
            let prior = try! app.book!.raw(0,"B4")
            choose("aggregate","aggregate:"+function)
            sheetCheck((try! app.book!.raw(0,"B4")) == "=\(function)(B2:B3)" && (try! app.book!.value(0,4,2).text) == expected,"Native Σ \(function)")
            app.undoEdit(); sheetCheck((try! app.book!.raw(0,"B4")) == prior,"ONE Σ \(function) Undo")
            app.redoEdit(); await pause()
        }
        app.book?.dirty = false
        app.newWorkbook(); await wait("v1.3 workbook",{ !app.web.isLoading })
        _ = await js("select(at(1,1))"); await wait("v1.3 origin",{ app.selectedRef == "A1" })
        try! app.pasteText("Name\tAmount\none\t10\ntwo\t20"); await pause()
        let gridXML = try! app.book!.sheetDocument(0).xmlData,gridActions = app.committedActions
        _ = await js("rowMap.get(2).querySelector('th').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));rowMap.get(3).querySelector('th').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0,shiftKey:true}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));rowMap.get(2).querySelector('th').dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,cancelable:true,clientX:20,clientY:100}))")
        await wait("multi-row headers",{ app.selectedRange.r1 == 2 && app.selectedRange.r2 == 3 })
        sheetCheck(await js("document.querySelector('#grid-menu button').textContent") as? String == gridActionTitle(rows:true,count:2,deleting:false),"Header context menu uses the selected row count")
        _ = await js("document.querySelector('#grid-menu button').click()")
        await wait("insert two rows",{ !app.web.isLoading && (try? app.book!.raw(0,"A4")) == "one" })
        sheetCheck(app.committedActions == gridActions+1 && app.editUndo.groupingLevel == 0,"Header insertion is one closed undo step")
        app.undoEdit(); await wait("structural undo",{ !app.web.isLoading && (try? app.book!.raw(0,"A2")) == "one" })
        sheetCheck(try! app.book!.sheetDocument(0).xmlData == gridXML,"UI structural Undo restores exact sheet XML")
        _ = await js("select(at(1,1));select(at(3,2),false,true)"); await wait("chart selection",{ app.selectedRange.count == 6 })
        let chartMenu = NSMenuItem(title:"",action:#selector(App.createChartMenu(_:)),keyEquivalent:""); chartMenu.representedObject = "column"
        let chartXML = try! app.book!.sheetDocument(0).xmlData
        app.createChartMenu(chartMenu); await pause()
        let renderedBars = await js("document.querySelectorAll('.chart-box svg [data-mark=bar]').length") as? Int
        sheetCheck((try! app.book!.readCharts(0).count) == 1 && renderedBars == 2,"Insert Chart creates DrawingML and two live SVG bars")
        let chartFrame = try! app.book!.readCharts(0)[0].frame
        _ = await js("document.querySelector('.chart-box').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0,clientX:300,clientY:100}));document.dispatchEvent(new MouseEvent('mousemove',{bubbles:true,clientX:340,clientY:130}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}))")
        await pause()
        sheetCheck((try! app.book!.readCharts(0)[0].frame.minX) == chartFrame.minX+40,"Chart dragging saves the moved anchor")
        _ = await js("const corner=document.querySelector('.chart-resize');corner.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0,clientX:500,clientY:300}));document.dispatchEvent(new MouseEvent('mousemove',{bubbles:true,clientX:540,clientY:320}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}))")
        await pause()
        sheetCheck((try! app.book!.readCharts(0)[0].frame.width) == chartFrame.width+40,"Corner dragging saves resized chart dimensions")
        _ = await js("document.querySelector('.chart-box').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0,clientX:300,clientY:100}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));document.querySelector('.chart-box').dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'Delete'}))")
        await pause(); sheetCheck((try! app.book!.readCharts(0).isEmpty),"Delete removes selected chart without clearing cells")
        app.undoEdit(); await pause(); sheetCheck((try! app.book!.readCharts(0).count) == 1,"ONE Undo restores deleted chart")
        app.undoEdit(); await pause(); sheetCheck((try! app.book!.readCharts(0)[0].frame.width) == chartFrame.width,"ONE Undo restores chart size")
        app.undoEdit(); await pause(); sheetCheck((try! app.book!.readCharts(0)[0].frame) == chartFrame,"ONE Undo restores chart position")
        app.undoEdit(); await pause(); sheetCheck((try! app.book!.readCharts(0).isEmpty) && (try! app.book!.sheetDocument(0).xmlData) == chartXML,"ONE Undo removes chart creation and restores original XML")
        app.book?.dirty = false
        print("ALL EXTENDED UI TESTS PASSED")
        scheduleLegacySheetTests(app,fixture:fixture)
    }
}
