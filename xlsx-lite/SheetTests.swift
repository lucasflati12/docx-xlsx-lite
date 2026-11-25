import AppKit
import PDFKit
import WebKit

final class PrintWaiter: NSObject {
    var done = false
    @objc func printOperationDidRun(_ operation: NSPrintOperation, success: Bool, contextInfo: UnsafeMutableRawPointer?) { done = true }
}
let printWaiter = PrintWaiter()

// Debug-only end-to-end checks: open a workbook in the real window, navigate, edit, save a copy, print to PDF.
func scheduleSheetTests(_ app: App, fixture: URL) {
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
