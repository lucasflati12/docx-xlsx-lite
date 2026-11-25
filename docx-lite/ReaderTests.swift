import AppKit

func scheduleReaderTests(_ reader: ReaderApp, fixture: URL) {
    DispatchQueue.main.asyncAfter(deadline: .now() + 0.5) {
        do {
            let original = try Data(contentsOf: fixture)
            let doc = try WordDocument(url: fixture)
            var passed = 0
            func check(_ condition: @autoclosure () -> Bool, _ message: String) {
                guard condition() else { fputs("FAIL: \(message)\n", stderr); exit(1) }
                passed += 1
                print("PASS: \(message)")
            }
            func descendants(_ node: XMLNode, named name: String) -> [XMLNode] {
                (node.children ?? []).flatMap { child in (child.localName == name ? [child] : []) + descendants(child, named: name) }
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
            let imageData = try Data(contentsOf: fixture.deletingLastPathComponent().appendingPathComponent("image.png"))
            check(NSBitmapImageRep(data: imageData) != nil, "Image fixture decodes")
            check(tableCount >= 4, "Table cells")
            check(links > 0, "Hyperlink")
            check(bold > 0, "Bold formatting")
            reader.display(doc, url: fixture)
            check(reader.textView.isEditable && reader.textView.isSelectable, "Editable selectable text")
            check(!reader.isDirty && !reader.window.isDocumentEdited && reader.status.stringValue.contains(L("Salvato", "Saved")), "Opening rich DOCX leaves document clean and status saved")
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
            check(reader.matches.isEmpty && reader.searchStatus.stringValue == L("Nessun risultato", "No results"), "No search results")
            reader.search.stringValue = ""
            reader.updateMatches()
            check(reader.searchStatus.stringValue.isEmpty, "Search cleared")
            reader.applyZoom(100)
            check(reader.scroll.magnification == 2, "Maximum zoom")
            reader.applyZoom(0)
            check(reader.scroll.magnification == 0.5, "Minimum zoom")
            reader.resetZoom()
            check(reader.scroll.magnification == 1, "Reset zoom")
            reader.fitWidth()
            check(reader.canvas.frame.width >= reader.scroll.contentView.bounds.width - 1, "Fit width geometry")
            check(reader.textView.frame.minX >= 0, "Page remains in canvas")
            reader.textView.setSelectedRange(NSRange(location: 20, length: 10))
            reader.scroll.contentView.scroll(to: NSPoint(x: 0, y: 60))
            reader.window.displayIfNeeded()
            check(!reader.isDirty && !reader.window.isDocumentEdited, "Scrolling, zooming, searching and selecting leave document clean")
            reader.textView.setSelectedRange(NSRange(location: reader.textView.string.utf16.count, length: 0))
            reader.textView.insertText("x", replacementRange: reader.textView.selectedRange())
            reader.textView.breakUndoCoalescing()
            check(reader.isDirty && reader.window.isDocumentEdited, "Typing one character in opened rich DOCX marks it edited")
            reader.undoEdit()
            check(!reader.isDirty && !reader.window.isDocumentEdited && reader.status.stringValue.contains(L("Salvato", "Saved")), "Undo to original rich DOCX restores clean state")
            let emptyURL = fixture.deletingLastPathComponent().appendingPathComponent("empty.docx")
            let empty = try WordDocument(url: emptyURL)
            reader.display(empty, url: emptyURL)
            check(empty.wordCount == 0, "Empty document opens")
            check(reader.textView.frame.height > 100, "Empty document has a visible page")
            let broken = fixture.deletingLastPathComponent().appendingPathComponent("broken.docx")
            do {
                _ = try WordDocument(url: broken)
                check(false, "Broken file is rejected")
            } catch { check(true, "Broken file is rejected") }
            let afterReading = try Data(contentsOf: fixture)
            check(afterReading == original, "Source file is unchanged")
            // Probe the unmodified AppKit exporter separately from the image fix.
            let rawURL = fixture.deletingLastPathComponent().appendingPathComponent("native-writer.docx")
            do {
                let probe = NSMutableAttributedString(attributedString: doc.text)
                probe.removeAttribute(docxPackageKey, range: range)
                let raw = try probe.data(from: range, documentAttributes: [.documentType: NSAttributedString.DocumentType.officeOpenXML])
                try raw.write(to: rawURL)
                let native = try WordDocument(url: rawURL)
                var nativeImages = 0
                native.text.enumerateAttribute(.attachment, in: NSRange(location: 0, length: native.text.length)) { value, _, _ in if value != nil { nativeImages += 1 } }
                print(nativeImages == imageCount ? "PASS: Native AppKit writer preserves images" : "WARN: Native AppKit DOCX writer drops images; DOCX Lite adds DrawingML image parts")
            } catch { print("WARN: Native AppKit image export failed: \(error.localizedDescription); testing DOCX Lite image fix") }
            reader.display(doc, url: fixture)
            let copyURL = fixture.deletingLastPathComponent().appendingPathComponent("rich-copy.docx")
            reader.textView.setSelectedRange(NSRange(location: reader.textView.string.utf16.count, length: 0))
            reader.textView.insertText("\nCopy edit", replacementRange: reader.textView.selectedRange())
            try reader.writeDocument(to: copyURL)
            let afterCopy = try Data(contentsOf: fixture)
            check(afterCopy == original, "Save a Copy leaves original fixture unchanged")
            let imageCopy = try WordDocument(url: copyURL)
            var roundtripImages = 0
            imageCopy.text.enumerateAttribute(.attachment, in: NSRange(location: 0, length: imageCopy.text.length)) { value, _, _ in if value != nil { roundtripImages += 1 } }
            check(roundtripImages == imageCount, "Image round trip with DrawingML fix")
            check(imageCopy.text.string.contains("Copy edit"), "Copy contains edits")

            let copyXML = try XMLDocument(data: archiveMember(copyURL, "word/document.xml"), options: [.nodeLoadExternalEntitiesNever])
            let savedTables = try copyXML.nodes(forXPath: "//*[local-name()='tbl']")
            let savedRows = try savedTables.first?.nodes(forXPath: "./*[local-name()='tr']") ?? []
            let savedCells = try savedTables.first?.nodes(forXPath: "./*[local-name()='tr']/*[local-name()='tc']") ?? []
            let cellTexts = savedCells.map { descendants($0, named: "t").compactMap { $0.stringValue }.joined() }
            check(savedTables.count == 1 && savedRows.count == 2 && savedCells.count == 4 && cellTexts == ["Nome", "Valore", "Documento", "Leggerissimo"], "Save a Copy preserves a real 2x2 Word table and all four cell texts")
            var reopenedTableParagraphs = 0
            imageCopy.text.enumerateAttribute(.paragraphStyle, in: NSRange(location: 0, length: imageCopy.text.length)) { value, _, _ in
                if let style = value as? NSParagraphStyle, !style.textBlocks.isEmpty { reopenedTableParagraphs += 1 }
            }
            check(reopenedTableParagraphs >= 4, "Saved Word table reopens as native NSTextTable cells")
            check(!reader.needsRewriteWarning && doc.package?.unsupported.isEmpty == true, "Supported tables do not require an overwrite warning")
            reader.newBlank()
            check(reader.currentURL == nil && reader.window.title == L("Senza titolo", "Untitled"), "New blank untitled document")
            check(reader.textView.textContainer!.containerSize.width == 483, "New document uses A4 and margins")
            reader.textView.insertText("Undo me", replacementRange: NSRange(location: 0, length: 0))
            reader.textView.breakUndoCoalescing()
            check(reader.textView.string == "Undo me" && reader.isDirty && reader.window.isDocumentEdited, "Typing updates document and edited dot")
            reader.undoEdit()
            check(reader.textView.string.isEmpty && !reader.isDirty, "Undo restores text and saved state")
            reader.redoEdit()
            check(reader.textView.string == "Undo me", "Redo restores text")
            reader.newBlank()
            reader.textView.insertText("Bold Italic Underline\nHeading\nBullet\nNumber\nCentered\nLinked\nHighlighted\nTask\nReplace café CAFE", replacementRange: NSRange(location: 0, length: 0))
            func select(_ string: String) -> NSRange {
                let r = (reader.textView.string as NSString).range(of: string)
                reader.textView.setSelectedRange(r)
                return r
            }
            func attrs(_ string: String, in text: NSAttributedString? = nil) -> [NSAttributedString.Key: Any] {
                let text = text ?? reader.textView.attributedString()
                return text.attributes(at: (text.string as NSString).range(of: string).location, effectiveRange: nil)
            }
            _ = select("Bold"); reader.bold()
            check((attrs("Bold")[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.bold), "Bold command")
            _ = select("Italic"); reader.italicText()
            check((attrs("Italic")[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.italic), "Italic command")
            _ = select("Underline"); reader.underline()
            check((attrs("Underline")[.underlineStyle] as? Int ?? 0) == 1, "Underline command")
            _ = select("Heading"); reader.applyStyle(2)
            check((attrs("Heading")[.font] as! NSFont).pointSize == 24 && reader.stylePopup.indexOfSelectedItem == 2, "Heading style and toolbar state")
            _ = select("Bullet"); reader.bulletList()
            check(!(attrs("Bullet")[.paragraphStyle] as! NSParagraphStyle).textLists.isEmpty && reader.textView.string.contains("•"), "Bulleted list uses NSTextList and markers")
            _ = select("Number"); reader.numberedList()
            check((attrs("Number")[.paragraphStyle] as! NSParagraphStyle).textLists.last?.markerFormat.rawValue.contains("decimal") == true && reader.textView.string.contains("1."), "Numbered list uses NSTextList and markers")
            _ = select("Centered"); reader.alignCenter()
            check((attrs("Centered")[.paragraphStyle] as! NSParagraphStyle).alignment == .center, "Paragraph alignment")
            _ = select("Linked"); reader.applyLink("https://example.com/editor")
            check((attrs("Linked")[.link] as? URL)?.absoluteString == "https://example.com/editor", "Link insertion")
            reader.applyLink("")
            check(attrs("Linked")[.link] == nil, "Empty link removes link")
            reader.applyLink("https://example.com/editor")
            _ = select("Highlighted"); reader.setAttributes([.backgroundColor: NSColor.yellow])
            check(attrs("Highlighted")[.backgroundColor] != nil, "Highlight color")
            reader.setAttributes([:], remove: [.backgroundColor])
            check(attrs("Highlighted")[.backgroundColor] == nil, "Remove highlight")
            reader.setAttributes([.backgroundColor: NSColor.yellow])
            _ = select("Task"); reader.checklist()
            check(reader.textView.string.contains("☐ Task"), "Insert checklist")
            _ = select("Task"); reader.checklist()
            check(reader.textView.string.contains("☑ Task"), "Toggle checklist")
            _ = select("Task"); reader.checklist()
            check(reader.textView.string.contains("☐ Task"), "Toggle checklist back to unchecked")
            let replaced = reader.replaceAll("cafe", with: "coffee")
            check(replaced == 2 && reader.textView.string.contains("coffee coffee"), "Find and replace handles accents and case")
            let counts = reader.documentCounts()
            check(counts.words > 10 && counts.characters == reader.textView.string.count, "Word and character count")
            _ = select("Centered")
            let spacing = NSMenuItem(); spacing.tag = 150; reader.spacingChosen(spacing)
            check((attrs("Centered")[.paragraphStyle] as! NSParagraphStyle).lineHeightMultiple == 1.5, "Line spacing command")
            _ = select("Heading"); spacing.tag = 115; reader.spacingChosen(spacing)
            _ = select("Task"); reader.indentMore()
            check((attrs("Task")[.paragraphStyle] as! NSParagraphStyle).headIndent == 24, "Increase indent")
            reader.indentLess()
            check((attrs("Task")[.paragraphStyle] as! NSParagraphStyle).headIndent == 0, "Decrease indent")
            _ = select("Task"); reader.bold(); reader.setFontSize(18); reader.clearFormat()
            check((attrs("Task")[.font] as! NSFont).pointSize == 12 && !(attrs("Task")[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.bold), "Clear formatting")
            let spell = reader.textView.isContinuousSpellCheckingEnabled; reader.spellToggle()
            check(reader.textView.isContinuousSpellCheckingEnabled != spell, "Spell check toggle")
            reader.spellToggle()
            let editedURL = fixture.deletingLastPathComponent().appendingPathComponent("edited.docx")
            try reader.writeDocument(to: editedURL)
            check(!reader.isDirty && !reader.window.isDocumentEdited, "Atomic save clears edited state")
            let reopened = try WordDocument(url: editedURL)
            check(reopened.text.string.contains("coffee coffee"), "DOCX reopens with edited text")
            check((attrs("Bold", in: reopened.text)[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.bold), "Bold survives DOCX round trip")
            check((attrs("Heading", in: reopened.text)[.font] as! NSFont).pointSize == 24, "Heading size survives DOCX round trip")
            check(reopened.text.string.contains("•") && reopened.text.string.contains("1."), "List markers survive DOCX round trip")
            check(attrs("Linked", in: reopened.text)[.link] != nil, "Link survives DOCX round trip")
            check((attrs("Centered", in: reopened.text)[.paragraphStyle] as! NSParagraphStyle).alignment == .center, "Alignment survives DOCX round trip")
            check((attrs("Italic", in: reopened.text)[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.italic), "Italic survives DOCX round trip")
            check((attrs("Underline", in: reopened.text)[.underlineStyle] as? Int ?? 0) != 0, "Underline survives DOCX round trip")
            check(attrs("Highlighted", in: reopened.text)[.backgroundColor] != nil, "Highlight survives DOCX round trip")
            check((attrs("Centered", in: reopened.text)[.paragraphStyle] as! NSParagraphStyle).lineHeightMultiple == 1.5, "Line spacing survives DOCX round trip")
            check(abs((attrs("Heading", in: reopened.text)[.paragraphStyle] as! NSParagraphStyle).lineHeightMultiple - 1.15) < 0.001, "Fractional line spacing survives DOCX round trip")
            check(!(attrs("Bullet", in: reopened.text)[.paragraphStyle] as! NSParagraphStyle).textLists.isEmpty, "Native list editing restored on reopen")
            reader.display(reopened, url: editedURL)
            let wide = NSImage(size: NSSize(width: 1600, height: 800))
            wide.lockFocus(); NSColor.systemBlue.setFill(); NSRect(x: 0, y: 0, width: 1600, height: 800).fill(); wide.unlockFocus()
            reader.textView.setSelectedRange(NSRange(location: reader.textView.string.utf16.count, length: 0))
            reader.insertImage(wide.tiffRepresentation!)
            check(reader.textView.selectedRange().location == reader.textView.string.utf16.count, "Cursor follows inserted image")
            let attachment = reader.textView.textStorage!.attribute(.attachment, at: reader.textView.string.utf16.count - 1, effectiveRange: nil) as! NSTextAttachment
            check((attachment.attachmentCell as! NSTextAttachmentCell).image!.size.width <= 483, "Inserted image scales to page width")
            let insertedURL = fixture.deletingLastPathComponent().appendingPathComponent("inserted-image.docx")
            try reader.writeDocument(to: insertedURL)
            let inserted = try WordDocument(url: insertedURL)
            var insertedImages = 0
            inserted.text.enumerateAttribute(.attachment, in: NSRange(location: 0, length: inserted.text.length)) { value, _, _ in if value != nil { insertedImages += 1 } }
            check(insertedImages == 1, "Inserted image survives DOCX round trip")
            check(attrs("Linked", in: inserted.text)[.link] != nil && attrs("Highlighted", in: inserted.text)[.backgroundColor] != nil, "Repeated save preserves links and highlight beside images")
            let legacyURL = fixture.deletingLastPathComponent().appendingPathComponent("legacy.rtf")
            let legacyData = try Data(contentsOf: legacyURL)
            let legacy = try WordDocument(url: legacyURL)
            reader.display(legacy, url: legacyURL)
            check(!reader.needsRewriteWarning && reader.currentURL?.pathExtension == "rtf", "Legacy format opens for DOCX conversion")
            try reader.writeDocument(to: fixture.deletingLastPathComponent().appendingPathComponent("legacy-copy.docx"))
            let unchangedLegacy = try Data(contentsOf: legacyURL)
            check(legacyData == unchangedLegacy, "DOCX conversion leaves RTF original unchanged")
            reader.newBlank()
            reader.textView.insertText("One", replacementRange: NSRange(location: 0, length: 0))
            reader.textView.setSelectedRange(NSRange(location: 0, length: 3)); reader.numberedList()
            reader.textView.setSelectedRange(NSRange(location: reader.textView.string.utf16.count, length: 0))
            reader.textView.insertNewline(nil)
            reader.textView.insertText("Two", replacementRange: reader.textView.selectedRange())
            check(reader.textView.string.contains("2."), "Numbered list continues on Return")
            // Inspect the package itself: visible formatting alone is insufficient.
            reader.newBlank()
            reader.textView.insertText("Document title\nHeading one\nHeading two\nHeading three\nBullet one\nBullet two\nNumber one\nNumber nested\nNumber two", replacementRange: NSRange(location: 0, length: 0))
            for (index, title) in ["Document title", "Heading one", "Heading two", "Heading three"].enumerated() { _ = select(title); reader.applyStyle(index + 1) }
            _ = select("Heading one"); reader.setFontSize(18)
            let bulletsStart = (reader.textView.string as NSString).range(of: "Bullet one").location
            let bulletsEnd = NSMaxRange((reader.textView.string as NSString).range(of: "Bullet two"))
            reader.textView.setSelectedRange(NSRange(location: bulletsStart, length: bulletsEnd - bulletsStart)); reader.bulletList()
            let numbersStart = (reader.textView.string as NSString).range(of: "Number one").location
            reader.textView.setSelectedRange(NSRange(location: numbersStart, length: reader.textView.string.utf16.count - numbersStart)); reader.numberedList()
            _ = select("Number nested"); reader.indentMore()
            let semanticURL = fixture.deletingLastPathComponent().appendingPathComponent("semantic.docx")
            try reader.writeDocument(to: semanticURL)
            let semanticXML = try XMLDocument(data: archiveMember(semanticURL, "word/document.xml"), options: [.nodeLoadExternalEntitiesNever])
            let numberingXML = try XMLDocument(data: archiveMember(semanticURL, "word/numbering.xml"), options: [.nodeLoadExternalEntitiesNever])
            let stylesXML = try XMLDocument(data: archiveMember(semanticURL, "word/styles.xml"), options: [.nodeLoadExternalEntitiesNever])
            let numberedParagraphs = try semanticXML.nodes(forXPath: "//*[local-name()='p'][*[local-name()='pPr']/*[local-name()='numPr']]")
            check(numberedParagraphs.count == 5, "Bullets and numbering export as real w:numPr paragraphs")
            let formats = try numberingXML.nodes(forXPath: "//*[local-name()='numFmt']/@*[local-name()='val']").compactMap { $0.stringValue }
            check(formats.contains("bullet") && formats.contains("decimal"), "numbering.xml defines both bulleted and decimal Word lists")
            let levels = try semanticXML.nodes(forXPath: "//*[local-name()='numPr']/*[local-name()='ilvl']/@*[local-name()='val']").compactMap { $0.stringValue }
            check(levels.filter { $0 == "1" }.count == 1 && levels.filter { $0 == "0" }.count == 4, "Word list levels follow paragraph indentation")
            let listText = numberedParagraphs.map { descendants($0, named: "t").compactMap { $0.stringValue }.joined() }
            let listTabs = numberedParagraphs.flatMap { descendants($0, named: "tab") }
            check(listText == ["Bullet one", "Bullet two", "Number one", "Number nested", "Number two"] && listTabs.isEmpty, "Word list content contains no typed bullet or number prefixes")
            let styleIDs = try semanticXML.nodes(forXPath: "//*[local-name()='pStyle']/@*[local-name()='val']").compactMap { $0.stringValue }
            check(styleIDs == ["Title", "Heading1", "Heading2", "Heading3"], "Title and Heading 1-3 export with semantic w:pStyle IDs")
            let definedIDs = try stylesXML.nodes(forXPath: "//*[local-name()='style']/@*[local-name()='styleId']").compactMap { $0.stringValue }
            let outlineLevels = try stylesXML.nodes(forXPath: "//*[local-name()='outlineLvl']/@*[local-name()='val']").compactMap { $0.stringValue }
            check(["Title", "Heading1", "Heading2", "Heading3"].allSatisfy { definedIDs.contains($0) } && outlineLevels == ["0", "1", "2"], "styles.xml defines actual Word heading outline levels")
            let typesXML = try XMLDocument(data: archiveMember(semanticURL, "[Content_Types].xml"), options: [.nodeLoadExternalEntitiesNever])
            let relsXML = try XMLDocument(data: archiveMember(semanticURL, "word/_rels/document.xml.rels"), options: [.nodeLoadExternalEntitiesNever])
            let parts = try typesXML.nodes(forXPath: "//*[local-name()='Override']/@PartName").compactMap { $0.stringValue }
            let targets = try relsXML.nodes(forXPath: "//*[local-name()='Relationship']/@Target").compactMap { $0.stringValue }
            check(parts.contains("/word/styles.xml") && parts.contains("/word/numbering.xml") && targets.contains("styles.xml") && targets.contains("numbering.xml"), "Semantic parts have package relationships and content types")
            let semantic = try WordDocument(url: semanticURL)
            check((attrs("Heading one", in: semantic.text)[wordParagraphStyleKey] as? String) == "Heading1" && (attrs("Heading one", in: semantic.text)[.font] as! NSFont).pointSize == 18, "Heading identity survives independent font size changes and reopening")
            check((attrs("Number nested", in: semantic.text)[.paragraphStyle] as! NSParagraphStyle).textLists.count == 2 && !(attrs("Bullet one", in: semantic.text)[.paragraphStyle] as! NSParagraphStyle).textLists.isEmpty, "Real Word lists reopen as nested native NSTextList styles")
            reader.display(semantic, url: semanticURL)
            check(!reader.isDirty && !reader.window.isDocumentEdited, "Reopening semantic headings and lists stays clean")
            let semanticAgainURL = fixture.deletingLastPathComponent().appendingPathComponent("semantic-again.docx")
            try reader.writeDocument(to: semanticAgainURL)
            let secondSemantic = try XMLDocument(data: archiveMember(semanticAgainURL, "word/document.xml"), options: [.nodeLoadExternalEntitiesNever])
            let secondLists = try secondSemantic.nodes(forXPath: "//*[local-name()='numPr']")
            let secondHeadings = try secondSemantic.nodes(forXPath: "//*[local-name()='pStyle']")
            check(secondLists.count == 5 && secondHeadings.count == 4, "Repeated save preserves semantic lists and headings")

            _ = select("Heading one"); reader.paintFormat(); _ = select("Heading three")
            check((attrs("Heading three")[wordParagraphStyleKey] as? String) == "Heading1", "Paint format copies semantic heading identity")
            reader.undoEdit()
            check((attrs("Heading three")[wordParagraphStyleKey] as? String) == "Heading3", "Undo paint format restores semantic heading identity")
            reader.display(doc, url: fixture)
            _ = select("Nome"); reader.bold(); reader.setAttributes([.foregroundColor: NSColor.red, .backgroundColor: NSColor.yellow])
            _ = select("Valore"); reader.applyLink("https://example.com/table")
            _ = select("Documento"); reader.applyStyle(2); reader.clearFormat()
            _ = select("Leggerissimo"); reader.textView.setSelectedRange(NSRange(location: reader.textView.selectedRange().location, length: 0)); reader.insertImage(imageData)
            let tableURL = fixture.deletingLastPathComponent().appendingPathComponent("table-formatted.docx")
            try reader.writeDocument(to: tableURL)
            let tableXML = try XMLDocument(data: archiveMember(tableURL, "word/document.xml"), options: [.nodeLoadExternalEntitiesNever])
            let formattedCells = try tableXML.nodes(forXPath: "//*[local-name()='tbl']/*[local-name()='tr']/*[local-name()='tc']")
            let boldCellRuns = formattedCells.first.map { descendants($0, named: "b") } ?? []
            let cellImages = formattedCells.last.map { descendants($0, named: "drawing") } ?? []
            let cellLinks = formattedCells.dropFirst().first.map { descendants($0, named: "hyperlink") } ?? []
            let cellBorders = try tableXML.nodes(forXPath: "//*[local-name()='tcBorders']/*[@*[local-name()='val']='single']")
            check(formattedCells.count == 4 && !boldCellRuns.isEmpty && cellImages.count == 1 && !cellLinks.isEmpty, "Real table cells preserve formatted runs, links and embedded images")
            check(!cellBorders.isEmpty, "Saved table preserves simple cell borders")
            let formattedTable = try WordDocument(url: tableURL)
            check((attrs("Nome", in: formattedTable.text)[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.bold) && !(attrs("Documento", in: formattedTable.text)[.paragraphStyle] as! NSParagraphStyle).textBlocks.isEmpty, "Table formatting survives style application, clear formatting and round trip")
            let outer = NSTextTable(), inner = NSTextTable(); outer.numberOfColumns = 1; inner.numberOfColumns = 1
            let nestedStyle = NSMutableParagraphStyle()
            nestedStyle.textBlocks = [NSTextTableBlock(table: outer, startingRow: 0, rowSpan: 1, startingColumn: 0, columnSpan: 1), NSTextTableBlock(table: inner, startingRow: 0, rowSpan: 1, startingColumn: 0, columnSpan: 1)]
            let blockedURL = fixture.deletingLastPathComponent().appendingPathComponent("unsupported-table.docx")
            let originalBlocked = Data("KEEP ORIGINAL".utf8); try originalBlocked.write(to: blockedURL)
            var rejectedNested = false
            do { try writeDOCX(NSAttributedString(string: "Nested\n", attributes: [.font: NSFont(name: "Arial", size: 12)!, .paragraphStyle: nestedStyle]), to: blockedURL, pageSize: NSSize(width: 595, height: 842), margin: 56) }
            catch { rejectedNested = true }
            let afterBlocked = try Data(contentsOf: blockedURL)
            check(rejectedNested && afterBlocked == originalBlocked, "Unsupported nested table fails safely without overwriting destination")
            let uiFont = NSFont.systemFont(ofSize: 17, weight: .bold)
            let mappedFont = documentFont(uiFont)
            check(!mappedFont.fontName.hasPrefix(".") && !(mappedFont.familyName?.hasPrefix(".") ?? true) && mappedFont.pointSize == 17 && mappedFont.fontDescriptor.symbolicTraits.contains(.bold), "Private system UI fonts map to a real document font with size and traits")
            let savedFonts = try semanticXML.nodes(forXPath: "//*[local-name()='rFonts']/@*").compactMap { $0.stringValue }
            check(!savedFonts.contains { $0.hasPrefix(".") }, "Saved DOCX never contains private system font family names")
            let styleFont = reader.stylePopup.font ?? NSFont.systemFont(ofSize: NSFont.smallSystemFontSize)
            let allStyleNames = ["Testo normale", "Titolo", "Intestazione 1", "Intestazione 2", "Intestazione 3", "Normal text", "Title", "Heading 1", "Heading 2", "Heading 3"]
            let styleWidth = reader.stylePopup.constraints.first { $0.firstAttribute == .width && $0.secondItem == nil }!.constant
            check(allStyleNames.allSatisfy { ($0 as NSString).size(withAttributes: [.font: styleFont]).width + 32 <= styleWidth }, "Every Italian and English paragraph style name fits without truncation")
            // Paint format copies appearance once, without copying text, links or images.
            reader.newBlank()
            reader.textView.insertText("Source\nTarget\nLater", replacementRange: NSRange(location: 0, length: 0))
            _ = select("Source")
            let paintFont = NSFontManager.shared.convert(NSFont(name: "Georgia", size: 20)!, toHaveTrait: .boldFontMask)
            reader.setAttributes([.font: paintFont, .foregroundColor: NSColor.red, .backgroundColor: NSColor.yellow, .underlineStyle: 1, .link: URL(string: "https://example.com/source")!])
            reader.alignCenter(); spacing.tag = 150; reader.spacingChosen(spacing)
            _ = select("Target")
            reader.setAttributes([.link: URL(string: "https://example.com/target")!])
            _ = select("Source")
            reader.textView.undoManager?.removeAllActions()
            let beforePaint = NSAttributedString(attributedString: reader.textView.attributedString())
            reader.paintFormat()
            check(reader.paintAttributes != nil && reader.formatButtons["paint"]?.state == .on, "Paint format arms and reflects toolbar state")
            check(beforePaint.isEqual(to: reader.textView.attributedString()), "Arming paint format does not change the document")
            let targetRange = (reader.textView.string as NSString).range(of: "Target")
            reader.textView.isSelectingWithMouse = true
            reader.textView.setSelectedRange(NSRange(location: targetRange.location, length: 2))
            check(reader.paintAttributes != nil && (attrs("Target")[.font] as! NSFont).pointSize == 12, "Paint format waits for mouse selection to finish")
            reader.textView.setSelectedRange(targetRange)
            reader.textView.isSelectingWithMouse = false
            reader.textView.selectionFinished?()
            check((attrs("Target")[.font] as! NSFont).pointSize == 20 && (attrs("Target")[.font] as! NSFont).fontDescriptor.symbolicTraits.contains(.bold), "Paint format copies font and weight across full selection")
            check((attrs("Target")[.foregroundColor] as! NSColor).isEqual(NSColor.red) && attrs("Target")[.backgroundColor] != nil && (attrs("Target")[.underlineStyle] as? Int) == 1, "Paint format copies text color, highlight and underline")
            check((attrs("Target")[.paragraphStyle] as! NSParagraphStyle).alignment == .center && (attrs("Target")[.paragraphStyle] as! NSParagraphStyle).lineHeightMultiple == 1.5, "Paint format copies paragraph alignment and spacing")
            check((attrs("Target")[.link] as? URL)?.absoluteString == "https://example.com/target" && reader.textView.string == "Source\nTarget\nLater", "Paint format preserves text and target links")
            check(reader.paintAttributes == nil && reader.formatButtons["paint"]?.state == .off, "Paint format disarms after one application")
            reader.undoEdit()
            check(beforePaint.isEqual(to: reader.textView.attributedString()), "Undo restores formatting before paint format")
            reader.redoEdit()
            check((attrs("Target")[.font] as! NSFont).pointSize == 20, "Redo restores paint format")
            _ = select("Later")
            check((attrs("Later")[.font] as! NSFont).pointSize == 12, "Paint format does not affect later selections")
            _ = select("Later"); reader.paintFormat(); _ = select("Target")
            check(attrs("Target")[.backgroundColor] == nil && (attrs("Target")[.underlineStyle] as? Int ?? 0) == 0, "Paint format removes formatting absent from source")
            _ = select("Source"); reader.paintFormat(); reader.textView.cancelOperation(nil)
            check(reader.paintAttributes == nil, "Escape cancels paint format")
            reader.paintFormat(); reader.paintFormat()
            check(reader.paintAttributes == nil, "Clicking paint format again cancels it")
            reader.paintFormat()
            reader.display(reopened, url: editedURL)
            check(reader.paintAttributes == nil, "Opening a document cancels pending paint format")

            let expectedMenus = [appName, L("File", "File"), L("Modifica", "Edit"), L("Visualizza", "View"), L("Inserisci", "Insert"), L("Formato", "Format"), L("Strumenti", "Tools"), L("Finestra", "Window"), L("Guida", "Help")]
            check(NSApp.mainMenu!.items.map { $0.title } == expectedMenus, "Menus follow Google Docs order with native app and Window menus")
            let tools = NSApp.mainMenu!.items[6].submenu!
            check(tools.items.contains { $0.action == #selector(ReaderApp.countDialog) } && tools.items.contains { $0.action == #selector(ReaderApp.spellToggle) }, "Tools menu contains word count and spell check")
            check(NSApp.helpMenu?.items.first?.representedObject as? URL == reader.helpURL && reader.helpURL.absoluteString == "https://github.com/lucasflati12/docx-xlsx-lite", "Help command targets project documentation")
            let expectedControls = ["undo", "redo", "print", "spell", "paint", "zoom", "style", "font", "sizeDown", "size", "sizeUp", "bold", "italic", "underline", "textColor", "highlight", "link", "image", "alignment", "spacing", "check", "bullet", "number", "outdent", "indent", "clear"]
            let controls = reader.toolbar.arrangedSubviews.compactMap { $0.identifier?.rawValue }.filter { expectedControls.contains($0) }
            check(controls == expectedControls && reader.toolbarSeparators.count == 7, "Toolbar follows reference order and grouped separators")
            _ = select("Centered")
            let centerItem = reader.alignmentPopup.menu!.items.first { $0.action == #selector(ReaderApp.alignCenter) }!
            _ = reader.validateMenuItem(centerItem)
            check(reader.alignmentPopup.numberOfItems == 5 && centerItem.state == .on, "Compact alignment menu reflects selection")
            let swatch = NSBitmapImageRep(data: reader.textColorImage(.red).tiffRepresentation!)!
            var hasRedBar = false
            for x in 0..<swatch.pixelsWide { for y in 0..<swatch.pixelsHigh {
                if let color = swatch.colorAt(x: x, y: y)?.usingColorSpace(.sRGB), color.redComponent > 0.8 && color.greenComponent < 0.5 && color.blueComponent < 0.5 { hasRedBar = true }
            } }
            check(hasRedBar && reader.textColorPopup.item(at: 0)?.image?.isTemplate == false, "Text color icon includes a colored bar")
            for width: CGFloat in [900, 1100, 1600] {
                reader.adaptToolbar(width: width)
                reader.window.setContentSize(NSSize(width: width, height: 700))
                reader.window.contentView?.layoutSubtreeIfNeeded()
                let visible = reader.toolbar.arrangedSubviews.filter { $0.superview != nil && !$0.isHidden }
                let rects = visible.map { $0.convert($0.bounds, to: reader.toolbar) }
                check(abs(reader.window.contentView!.bounds.width - width) < 1 && rects.allSatisfy { $0.minX >= -1 && $0.maxX <= reader.toolbar.bounds.width + 1 } && zip(rects, rects.dropFirst()).allSatisfy { $0.maxX <= $1.minX + 1 }, "Grouped toolbar fits without overlap at \(Int(width)) pixels")
            }
            check(reader.search.superview === reader.toolbar && !reader.search.isHidden && reader.formatButtons["fit"]?.superview != nil, "Search and extra zoom controls appear when space permits")
            reader.adaptToolbar(width: 900)
            reader.window.setContentSize(NSSize(width: 900, height: 700))
            check(reader.search.superview == nil || reader.search.isHidden, "Search leaves the compact toolbar at 900 pixels")
            reader.focusSearch()
            reader.window.contentView?.layoutSubtreeIfNeeded()
            check(!reader.findBar.isHidden && reader.search.superview === reader.findBar && reader.findInset.constant == 34, "Command-F reveals an on-demand search field at compact widths")
            reader.search.stringValue = "coffee"
            reader.controlTextDidChange(Notification(name: NSControl.textDidChangeNotification, object: reader.search))
            check(reader.matches.count == 2, "On-demand search preserves existing search behavior")
            check(reader.control(reader.search, textView: reader.textView, doCommandBy: #selector(NSResponder.cancelOperation(_:))) && reader.findBar.isHidden && reader.matches.isEmpty, "Escape closes on-demand search and clears matches")
            reader.display(reopened, url: editedURL)
            reader.window.setContentSize(NSSize(width: 900, height: 700))
            reader.window.contentView?.layoutSubtreeIfNeeded()
            check(reader.window.contentView!.bounds.width <= 901, "Window stays at 900 pixels")
            let attached = reader.toolbar.arrangedSubviews.filter { $0.superview != nil && !$0.isHidden }
            let frames = attached.map { $0.convert($0.bounds, to: reader.toolbar) }
            check(frames.allSatisfy { $0.minX >= -1 && $0.maxX <= reader.toolbar.bounds.width + 1 }, "Toolbar fits at 900 pixels")
            check(zip(frames, frames.dropFirst()).allSatisfy { $0.maxX <= $1.minX + 1 }, "Toolbar controls do not overlap")
            reader.applyZoom(1.15)
            reader.window.displayIfNeeded()
            if let content = reader.window.contentView, let bitmap = content.bitmapImageRepForCachingDisplay(in: content.bounds) {
                content.cacheDisplay(in: content.bounds, to: bitmap)
                try bitmap.representation(using: .png, properties: [:])!.write(to: fixture.deletingLastPathComponent().appendingPathComponent("reader-preview.png"))
            }
            let mutableTable = NSTextTable(); mutableTable.numberOfColumns = 2
            let mutableBlock = NSTextTableBlock(table: mutableTable, startingRow: 0, rowSpan: 1, startingColumn: 0, columnSpan: 1)
            mutableBlock.setWidth(0.5, type: .absoluteValueType, for: .border)
            let mutableList = NSTextList(markerFormat: .disc, options: 0)
            let mutableStyle = NSMutableParagraphStyle(); mutableStyle.textBlocks = [mutableBlock]; mutableStyle.textLists = [mutableList]; mutableStyle.alignment = .center
            let mutableImage = NSImage(data: imageData)!
            let mutableAttachment = NSTextAttachment(); mutableAttachment.fileWrapper = FileWrapper(regularFileWithContents: imageData); mutableAttachment.attachmentCell = NSTextAttachmentCell(imageCell: mutableImage)
            let mutableText = NSMutableAttributedString(string: "x", attributes: [.font: NSFont(name: "Arial", size: 12)!, .paragraphStyle: mutableStyle])
            mutableText.append(NSAttributedString(attachment: mutableAttachment))
            let frozen = saveSnapshot(mutableText)
            let oldImageSize = mutableImage.size
            mutableTable.numberOfColumns = 4; mutableBlock.setWidth(3, type: .absoluteValueType, for: .border); mutableList.startingItemNumber = 7; mutableStyle.alignment = .right; mutableImage.size = NSSize(width: 12, height: 12)
            let frozenStyle = frozen.attribute(.paragraphStyle, at: 0, effectiveRange: nil) as! NSParagraphStyle
            let frozenBlock = frozenStyle.textBlocks[0] as! NSTextTableBlock
            let frozenAttachment = frozen.attribute(.attachment, at: 1, effectiveRange: nil) as! NSTextAttachment
            check(frozenBlock.table.numberOfColumns == 2 && frozenBlock.width(for: .border, edge: .minX) == 0.5 && frozenStyle.alignment == .center && frozenStyle.textLists[0].startingItemNumber == 1, "Save snapshot isolates mutable tables, borders, paragraph styles and lists")
            check((frozenAttachment.attachmentCell as! NSTextAttachmentCell).image!.size == oldImageSize && frozenAttachment.fileWrapper?.regularFileContents == imageData, "Save snapshot isolates image size and preserves encoded bytes")
            let jpeg = NSBitmapImageRep(data: imageData)!.representation(using: .jpeg, properties: [.compressionFactor: 0.7])!
            let jpegAttachment = NSTextAttachment(); jpegAttachment.fileWrapper = FileWrapper(regularFileWithContents: jpeg); jpegAttachment.attachmentCell = NSTextAttachmentCell(imageCell: NSImage(data: jpeg)!)
            let jpegURL = fixture.deletingLastPathComponent().appendingPathComponent("jpeg-preserved.docx")
            try writeDOCX(NSAttributedString(attachment: jpegAttachment), to: jpegURL, pageSize: NSSize(width: 595, height: 842), margin: 56)
            let savedJPEG = try archiveMember(jpegURL, "word/media/docxlite-0.jpg")
            check(savedJPEG == jpeg, "JPEG data is saved without lossy re-encoding")
            passed += try runPreservationTests(reader, folder: fixture.deletingLastPathComponent(), imageData: imageData)
            runAsyncSaveTests(reader, fixture: fixture) { extra in
                print("RESULT: \(passed + extra) passed, 0 failed")
                print("ALL READER AND EDITOR TESTS PASSED")
                reader.cleanText = NSAttributedString(attributedString: reader.textView.attributedString())
                NSApp.terminate(nil)
            }
        } catch {
            fputs("TEST ERROR: \(error)\n", stderr)
            exit(1)
        }
    }
}

func makeSavePerformanceDocument() -> NSAttributedString {
    let text = NSMutableAttributedString(string: "")
    let font = NSFont(name: "Arial", size: 12)!
    for index in 0..<6 {
        let table = NSTextTable(); table.numberOfColumns = 4
        for row in 0..<30 { for column in 0..<4 {
            let block = NSTextTableBlock(table: table, startingRow: row, rowSpan: 1, startingColumn: column, columnSpan: 1)
            block.setWidth(0.5, type: .absoluteValueType, for: .border)
            let style = NSMutableParagraphStyle(); style.textBlocks = [block]
            text.append(NSAttributedString(string: "T\(index) R\(row) C\(column)\n", attributes: [.font: font, .paragraphStyle: style]))
        } }
        text.append(NSAttributedString(string: "\n", attributes: [.font: font]))
    }
    let words = Array(repeating: "document", count: 30).joined(separator: " ") + "\n"
    for _ in 0..<100 { text.append(NSAttributedString(string: words, attributes: [.font: font])) }
    let image = NSImage(size: NSSize(width: 32, height: 24))
    image.lockFocus(); NSColor.systemBlue.setFill(); NSRect(x: 0, y: 0, width: 32, height: 24).fill(); image.unlockFocus()
    let png = NSBitmapImageRep(data: image.tiffRepresentation!)!.representation(using: .png, properties: [:])!
    for _ in 0..<20 {
        let attachment = NSTextAttachment(); attachment.fileWrapper = FileWrapper(regularFileWithContents: png)
        attachment.attachmentCell = NSTextAttachmentCell(imageCell: NSImage(data: png)!)
        text.append(NSAttributedString(attachment: attachment)); text.append(NSAttributedString(string: "\n", attributes: [.font: font]))
    }
    return NSAttributedString(attributedString: text)
}

func runSavePerformanceTest(enforceLimit: Bool) {
    do {
        let root = URL(fileURLWithPath: "/private/tmp/docxlite-tests")
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        let text = makeSavePerformanceDocument()
        let profile = DOCXSaveProfile(), start = ProcessInfo.processInfo.systemUptime
        print("PERF: 6 tables x 30 rows x 4 cells + 3000 words + 20 images, \(text.length) characters")
        fflush(stdout)
        try writeDOCX(text, to: root.appendingPathComponent(enforceLimit ? "performance.docx" : "performance-before.docx"), pageSize: NSSize(width: 595, height: 842), margin: 56, profile: profile)
        let total = ProcessInfo.processInfo.systemUptime - start
        for (label, seconds) in profile.stages { print(String(format: "PROFILE: %@ %.4f s", label, seconds)) }
        print(String(format: "PERF: total %.4f s, AppKit %.4f s, non-AppKit %.4f s", total, profile.appKitSeconds, total - profile.appKitSeconds))
        if enforceLimit {
            guard total < 2 else { fputs("FAIL: release save must complete in < 2 s\n", stderr); exit(1) }
            print("PASS: release save < 2 s")
            let savedURL = root.appendingPathComponent("performance.docx")
            let xml = try XMLDocument(data: archiveMember(savedURL, "word/document.xml"), options: [.nodeLoadExternalEntitiesNever])
            let tables = try xml.nodes(forXPath: "//*[local-name()='tbl']").count
            let cells = try xml.nodes(forXPath: "//*[local-name()='tc']").count
            let images = try xml.nodes(forXPath: "//*[local-name()='drawing']").count
            guard tables == 6 && cells == 720 && images == 20 else { fputs("FAIL: performance package lost tables/cells/images\n", stderr); exit(1) }
            let originalPNG = text.attribute(.attachment, at: text.length - 2, effectiveRange: nil) as! NSTextAttachment
            let savedPNG = try archiveMember(savedURL, "word/media/docxlite-0.png")
            guard savedPNG == originalPNG.fileWrapper?.regularFileContents else { fputs("FAIL: PNG bytes were re-encoded\n", stderr); exit(1) }
            print("PASS: embedded PNG bytes are preserved without re-encoding")
            let snapshotStart = ProcessInfo.processInfo.systemUptime
            _ = saveSnapshot(text)
            print(String(format: "PERF: main-thread snapshot %.4f s", ProcessInfo.processInfo.systemUptime - snapshotStart))
            let reopened = try WordDocument(url: savedURL)
            var words = 0
            text.string.enumerateSubstrings(in: text.string.startIndex..<text.string.endIndex, options: .byWords) { _, _, _, _ in words += 1 }
            guard reopened.wordCount == words else { fputs("FAIL: performance round trip changed word count\n", stderr); exit(1) }
            print("PASS: performance round trip preserves 6 tables, 720 cells, 20 images and \(words) words")
            let plain = NSAttributedString(string: String(repeating: Array(repeating: "word", count: 13).joined(separator: " ") + "\n", count: 100), attributes: [.font: NSFont(name: "Arial", size: 12)!])
            let plainProfile = DOCXSaveProfile(), plainStart = ProcessInfo.processInfo.systemUptime
            try writeDOCX(plain, to: root.appendingPathComponent("performance-plain.docx"), pageSize: NSSize(width: 595, height: 842), margin: 56, profile: plainProfile)
            let plainTotal = ProcessInfo.processInfo.systemUptime - plainStart
            print(String(format: "PERF: plain 6500 chars, total %.4f s, AppKit %.4f s, non-AppKit %.4f s", plainTotal, plainProfile.appKitSeconds, plainTotal - plainProfile.appKitSeconds))
            guard plainTotal - plainProfile.appKitSeconds < 0.3 else { fputs("FAIL: plain save overhead must be < 0.3 s\n", stderr); exit(1) }
            print("PASS: plain save overhead < 0.3 s")
            let normalStyle = NSMutableParagraphStyle(); normalStyle.lineHeightMultiple = 1.15
            let formattedPlain = NSMutableAttributedString(attributedString: plain)
            formattedPlain.addAttributes([wordParagraphStyleKey: "Normal", .paragraphStyle: normalStyle], range: NSRange(location: 0, length: plain.length))
            let normalProfile = DOCXSaveProfile(), normalStart = ProcessInfo.processInfo.systemUptime
            try writeDOCX(formattedPlain, to: root.appendingPathComponent("performance-normal.docx"), pageSize: NSSize(width: 595, height: 842), margin: 56, profile: normalProfile)
            let normalTotal = ProcessInfo.processInfo.systemUptime - normalStart
            print(String(format: "PERF: plain with Normal style and 1.15 spacing, total %.4f s, AppKit %.4f s, non-AppKit %.4f s", normalTotal, normalProfile.appKitSeconds, normalTotal - normalProfile.appKitSeconds))
            guard normalTotal - normalProfile.appKitSeconds < 0.3 else { fputs("FAIL: formatted plain save overhead must be < 0.3 s\n", stderr); exit(1) }
            print("PASS: formatted plain save overhead < 0.3 s")
            let manyLists = NSMutableAttributedString(string: "")
            let bullet = NSTextList(markerFormat: .disc, options: 0)
            let listStyle = NSMutableParagraphStyle(); listStyle.textLists = [bullet]; listStyle.headIndent = 30; listStyle.lineHeightMultiple = 1.15
            for index in 0..<1000 { manyLists.append(NSAttributedString(string: "\t•\tItem \(index)\n", attributes: [.font: NSFont(name: "Arial", size: 12)!, .paragraphStyle: listStyle])) }
            let listProfile = DOCXSaveProfile(), listStart = ProcessInfo.processInfo.systemUptime
            try writeDOCX(manyLists, to: root.appendingPathComponent("performance-lists.docx"), pageSize: NSSize(width: 595, height: 842), margin: 56, profile: listProfile)
            let listTotal = ProcessInfo.processInfo.systemUptime - listStart
            print(String(format: "PERF: 1000 list paragraphs, total %.4f s, non-AppKit %.4f s", listTotal, listTotal - listProfile.appKitSeconds))
            guard listTotal - listProfile.appKitSeconds < 1.5 else { fputs("FAIL: list save overhead must be < 1.5 s\n", stderr); exit(1) }
            print("PASS: list save overhead < 1.5 s")
        }
    } catch { fputs("PERF ERROR: \(error)\n", stderr); exit(1) }
}

private func runAsyncSaveTests(_ reader: ReaderApp, fixture: URL, completion: @escaping (Int) -> Void) {
    var passed = 0
    func check(_ condition: @autoclosure () -> Bool, _ message: String) {
        guard condition() else { fputs("FAIL: \(message)\n", stderr); exit(1) }
        passed += 1; print("PASS: \(message)")
    }
    let watchdog = DispatchWorkItem { fputs("FAIL: async save timed out\n", stderr); exit(1) }
    DispatchQueue.main.asyncAfter(deadline: .now() + 15, execute: watchdog)
    reader.newBlank()
    let undo = reader.textView.undoManager!
    undo.groupsByEvent = false
    func type(_ value: String) {
        undo.beginUndoGrouping()
        reader.textView.insertText(value, replacementRange: reader.textView.selectedRange())
        reader.textView.breakUndoCoalescing()
        undo.endUndoGrouping()
    }
    type("Snapshot version")
    let root = fixture.deletingLastPathComponent(), url = root.appendingPathComponent("async-save.docx")
    let duplicate = root.appendingPathComponent("duplicate-save.docx")
    try? FileManager.default.removeItem(at: duplicate)
    var heartbeat = false
    // Queued before the worker completion: this must execute while save is active.
    DispatchQueue.main.async {
        heartbeat = true
        check(reader.isSaving && reader.textView.isEditable, "UI remains responsive and editable while saving")
        type(" later")
    }
    reader.writeDocument(to: url) { result in
        do {
            if case .failure(let error) = result { throw error }
            check(Thread.isMainThread && reader.lastSaveWasOnMainThread == false && heartbeat, "DOCX writer runs off main thread and completion returns on main thread")
            let saved = try WordDocument(url: url)
            check(saved.text.string.contains("Snapshot version") && !saved.text.string.contains("later"), "Async save writes the captured snapshot")
            check(!reader.isSaving && reader.isDirty && reader.window.isDocumentEdited && reader.textView.string.contains("later"), "Edits typed during save remain unsaved and visible")
            reader.undoEdit()
            check(!reader.isDirty && !reader.window.isDocumentEdited, "Undo to asynchronously saved snapshot restores clean state")
            check(!FileManager.default.fileExists(atPath: duplicate.path), "Overlapping save never creates a second destination")
            type(" unsaved")
            let originalBytes = try Data(contentsOf: url)
            let missing = root.appendingPathComponent(UUID().uuidString).appendingPathComponent("missing/document.docx")
            reader.writeDocument(to: missing) { failure in
                if case .success = failure { check(false, "Async write failure is reported"); return }
                check(!reader.isSaving && reader.isDirty && reader.currentURL == url, "Failed async save preserves dirty state, URL and progress state")
                let after = try? Data(contentsOf: url)
                check(after == originalBytes && !FileManager.default.fileExists(atPath: missing.path), "Failed async save leaves original file intact and no partial destination")
            }
            reader.resolveUnsavedChanges { allowed in
                check(!allowed, "Close/open/quit continuation cannot proceed after failed save")
                let staleURL = root.appendingPathComponent("stale-save.docx")
                reader.writeDocument(to: staleURL) { staleResult in
                    if case .failure(let error) = staleResult { fputs("TEST ERROR: \(error)\n", stderr); exit(1) }
                    check(reader.currentURL == nil && reader.textView.string == "Other document" && !reader.isDirty, "Save completion cannot change a replacement document")
                    undo.groupsByEvent = true
                    watchdog.cancel(); completion(passed)
                }
                reader.display(WordDocument(text: NSAttributedString(string: "Other document", attributes: [.font: NSFont(name: "Arial", size: 12)!])), url: nil)
            }
        } catch { fputs("TEST ERROR: \(error)\n", stderr); exit(1) }
    }
    check(reader.isSaving && reader.saveProgress.superview != nil && reader.status.stringValue.contains(L("Salvataggio", "Saving")), "Async save exposes progress and localized saving status")
    reader.writeDocument(to: duplicate) { result in
        if case .failure = result { check(true, "Concurrent saves are rejected safely") }
        else { check(false, "Concurrent saves are rejected safely") }
    }
}

// Exercise the same attributed-string edits and snapshots used by background save.
func runPreservationTests(_ reader: ReaderApp, folder: URL, imageData: Data) throws -> Int {
    var passed = 0
    func check(_ condition: @autoclosure () -> Bool, _ message: String) {
        guard condition() else { fputs("FAIL: \(message)\n", stderr); exit(1) }
        passed += 1; print("PASS: \(message)")
    }
    let source = folder.appendingPathComponent("retained.docx")
    let original = try Data(contentsOf: source), doc = try WordDocument(url: source)
    check(doc.package?.unsupported.isEmpty == true, "Headers, notes, comments and sections need no overwrite warning")
    reader.display(doc, url: source)
    check(!reader.needsRewriteWarning && reader.headerPreview.stringValue == "First page header" && reader.footerPreview.stringValue.contains("First page footer"), "First-page read-only header/footer previews")
    check(!reader.headerPreview.isEditable && !reader.footerPreview.isEditable && doc.pageSize.width == 841.9 && doc.margin == 36, "Landscape geometry and read-only previews")
    var anchors: [(DOCXAnchor, Int)] = []
    doc.text.enumerateAttribute(docxAnchorKey, in: NSRange(location: 0, length: doc.text.length)) { value, range, _ in
        if let anchor = value as? DOCXAnchor { anchors.append((anchor, range.location)) }
    }
    check(anchors.count == 12, "Import restores both comment ranges, three note references, bookmark and section break")
    let indicators = anchors.filter { $0.0.visible }
    check(indicators.count == 5 && indicators.allSatisfy { doc.text.attribute(.toolTip, at: $0.1, effectiveRange: nil) as? String == $0.0.detail && !$0.0.detail.isEmpty }, "Read-only comment and note indicators have tooltips")
    check(indicators.filter { $0.0.kind != "commentReference" }.allSatisfy { (doc.text.attribute(.baselineOffset, at: $0.1, effectiveRange: nil) as? NSNumber)?.intValue == 3 }, "Footnote/endnote numbers appear as superscripts")
    check(documentPlainText(doc.text).contains("A bold comment span and note.") && !documentPlainText(doc.text).contains("DOCXLITE"), "Plain text and word counts omit invisible anchors")
    let tail = (reader.textView.string as NSString).range(of: "Editable tail.")
    reader.textView.insertText("Edited elsewhere.", replacementRange: tail)
    let image = NSTextAttachment(); image.fileWrapper = FileWrapper(regularFileWithContents: imageData)
    image.attachmentCell = NSTextAttachmentCell(imageCell: NSImage(data: imageData)!)
    reader.textView.textStorage!.append(NSAttributedString(attachment: image))
    reader.textView.textStorage!.append(NSAttributedString(string: "\nNew heading\nNew list item\n", attributes: [.font: NSFont(name: "Arial", size: 12)!]))
    let heading = (reader.textView.string as NSString).range(of: "New heading")
    reader.textView.textStorage!.addAttribute(wordParagraphStyleKey, value: "Heading1", range: heading)
    let list = NSMutableParagraphStyle(); list.textLists = [NSTextList(markerFormat: .decimal, options: 0)]; list.headIndent = 30
    let listRange = (reader.textView.string as NSString).paragraphRange(for: (reader.textView.string as NSString).range(of: "New list item"))
    reader.textView.textStorage!.addAttribute(.paragraphStyle, value: list, range: listRange)
    let saved = folder.appendingPathComponent("retained-edited.docx")
    try reader.writeDocument(to: saved)
    let reopened = try WordDocument(url: saved)
    check(documentPlainText(reopened.text).contains("Edited elsewhere."), "Unrelated edit survives package merge")
    let boldRange = (reopened.text.string as NSString).range(of: "bold comment")
    check((reopened.text.attribute(.font, at: boldRange.location, effectiveRange: nil) as? NSFont)?.fontDescriptor.symbolicTraits.contains(.bold) == true, "Comment spanning bold and plain runs retains formatting")
    let second = folder.appendingPathComponent("retained-second.docx")
    try writeDOCX(saveSnapshot(reopened.text), to: second, pageSize: reopened.pageSize, margin: reopened.margin)
    reader.display(reopened, url: saved)
    let deletion = (reader.textView.string as NSString).paragraphRange(for: (reader.textView.string as NSString).range(of: "Delete annotated paragraph"))
    reader.textView.insertText("", replacementRange: deletion)
    try reader.writeDocument(to: folder.appendingPathComponent("retained-deleted.docx"))
    reader.textView.undoManager?.undo()
    check(reader.textView.string.contains("Delete annotated paragraph"), "Deleting an annotated paragraph remains undoable")
    try reader.writeDocument(to: folder.appendingPathComponent("retained-undone.docx"))
    // Deleting only one boundary must also remove the incomplete range cleanly.
    let partial = NSMutableAttributedString(attributedString: reopened.text)
    var boundary: NSRange?
    partial.enumerateAttribute(docxAnchorKey, in: NSRange(location: 0, length: partial.length)) { value, range, _ in
        if let anchor = value as? DOCXAnchor, anchor.kind == "commentRangeStart", anchor.id == "1" { boundary = range }
    }
    partial.deleteCharacters(in: boundary!)
    try writeDOCX(partial, to: folder.appendingPathComponent("retained-partial.docx"), pageSize: reopened.pageSize, margin: reopened.margin)
    // The app holds the package even when all attributed text has been deleted.
    reader.textView.insertText("", replacementRange: NSRange(location: 0, length: reader.textView.string.utf16.count))
    try reader.writeDocument(to: folder.appendingPathComponent("retained-empty.docx"))
    let empty = try WordDocument(url: folder.appendingPathComponent("retained-empty.docx"))
    check(empty.package?.headerText == "First page header" || empty.package?.headerText.hasPrefix("Default page header") == true, "Deleting all body text retains package and final section")
    let unsupported = try WordDocument(url: folder.appendingPathComponent("unsupported.docx"))
    reader.display(unsupported, url: folder.appendingPathComponent("unsupported.docx"))
    check(reader.needsRewriteWarning && [L("Revisioni e modifiche tracciate", "Tracked changes"), L("Controlli contenuto", "Content controls"), L("Equazioni", "Equations")].allSatisfy { reader.rewriteWarningText.contains($0) }, "Overwrite warning lists the detected unsupported features")
    let unchanged = try Data(contentsOf: source)
    check(unchanged == original, "Preservation tests leave original package unchanged")
    reader.newBlank()
    return passed
}
