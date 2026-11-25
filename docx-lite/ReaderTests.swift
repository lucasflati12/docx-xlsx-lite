import AppKit

func scheduleReaderTests(_ reader: ReaderApp, fixture: URL) {
    DispatchQueue.main.asyncAfter(deadline: .now() + 0.5) {
        do {
            let original = try Data(contentsOf: fixture)
            let doc = try WordDocument(url: fixture)
            func check(_ condition: @autoclosure () -> Bool, _ message: String) {
                guard condition() else { fputs("FAIL: \(message)\n", stderr); exit(1) }
                print("PASS: \(message)")
            }
            check(doc.text.string.contains("🌈"), "Unicode and emoji")
            check(doc.wordCount > 400, "Word count")
            var imageCount = 0, tableCount = 0, links = 0, bold = 0
            let range = NSRange(location: 0, length: doc.text.length)
            doc.text.enumerateAttributes(in: range) { attrs, _, _ in
                if attrs[.attachment] != nil { imageCount += 1 }
                if attrs[.link] != nil { links += 1 }
                if let p = attrs[.paragraphStyle] as? NSParagraphStyle, !p.textBlocks.isEmpty { tableCount += 1 }
                if let f = attrs[.font] as? NSFont, f.fontDescriptor.symbolicTraits.contains(.bold) { bold += 1 }
            }
            check(imageCount == 1, "Embedded image")
            check(tableCount >= 4, "Table cells")
            check(links > 0, "Hyperlink")
            check(bold > 0, "Bold formatting")
            reader.display(doc, url: fixture)
            check(!reader.textView.isEditable && reader.textView.isSelectable, "Read-only selectable text")
            check(reader.canvas.frame.height > reader.scroll.contentSize.height, "Long document can scroll")
            reader.search.stringValue = "caffe"
            reader.updateMatches()
            check(reader.matches.count == 3, "Case and accent insensitive search")
            reader.findNext()
            check(reader.matchIndex == 0, "First search result")
            reader.findPrevious()
            check(reader.matchIndex == 2, "Search wraps backwards")
            reader.findNext()
            check(reader.matchIndex == 0, "Search wraps forwards")
            reader.search.stringValue = "nonexistent-word"
            reader.updateMatches()
            check(reader.matches.isEmpty && reader.searchStatus.stringValue == "Nessun risultato", "No search results")
            reader.search.stringValue = ""
            reader.updateMatches()
            check(reader.searchStatus.stringValue.isEmpty, "Search cleared")
            reader.applyZoom(100)
            check(reader.scroll.magnification == 3, "Maximum zoom")
            reader.applyZoom(0)
            check(reader.scroll.magnification == 0.5, "Minimum zoom")
            reader.resetZoom()
            check(reader.scroll.magnification == 1, "Reset zoom")
            reader.fitWidth()
            check(reader.canvas.frame.width >= reader.scroll.contentView.bounds.width - 1, "Fit width geometry")
            check(reader.textView.frame.minX >= 0, "Page remains in canvas")
            let emptyURL = fixture.deletingLastPathComponent().appendingPathComponent("empty.docx")
            let empty = try WordDocument(url: emptyURL)
            reader.display(empty, url: emptyURL)
            check(empty.wordCount == 0, "Empty document opens")
            check(reader.textView.frame.height > 100, "Empty document has a visible page")
            let broken = fixture.deletingLastPathComponent().appendingPathComponent("broken.docx")
            do {
                _ = try WordDocument(url: broken)
                check(false, "Broken file is rejected")
            } catch { print("PASS: Broken file is rejected") }
            let afterReading = try Data(contentsOf: fixture)
            check(afterReading == original, "Source file is unchanged")
            reader.display(doc, url: fixture)
            reader.applyZoom(1.15)
            reader.window.displayIfNeeded()
            if let content = reader.window.contentView, let bitmap = content.bitmapImageRepForCachingDisplay(in: content.bounds) {
                content.cacheDisplay(in: content.bounds, to: bitmap)
                try bitmap.representation(using: .png, properties: [:])!.write(to: fixture.deletingLastPathComponent().appendingPathComponent("reader-preview.png"))
            }
            print("ALL READER TESTS PASSED")
            NSApp.terminate(nil)
        } catch {
            fputs("TEST ERROR: \(error)\n", stderr)
            exit(1)
        }
    }
}
