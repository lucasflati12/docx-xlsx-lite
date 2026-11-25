import AppKit
import UniformTypeIdentifiers

let appName = "DOCX Lite"
let italian = (Locale.preferredLanguages.first ?? "").lowercased().hasPrefix("it")
func L(_ it: String, _ en: String) -> String { italian ? it : en }
let readableExtensions = ["docx", "doc", "odt", "rtf"]
private let readableTypes = readableExtensions.compactMap { UTType(filenameExtension: $0) }

private struct EmbeddedPicture {
    let marker: String
    let data: Data
    let size: NSSize
}

private struct EmbeddedLink {
    let start: String
    let end: String
    let url: URL
}

private func archiveMember(_ archive: URL, _ member: String) throws -> Data {
    let process = Process()
    process.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
    process.arguments = ["-p", archive.path, member]
    let output = Pipe()
    process.standardOutput = output
    process.standardError = FileHandle.nullDevice
    try process.run()
    let data = output.fileHandleForReading.readDataToEndOfFile()
    process.waitUntilExit()
    guard process.terminationStatus == 0 else {
        throw NSError(domain: "LettoreDOCX", code: 2, userInfo: [NSLocalizedDescriptionKey: L("Archivio DOCX non leggibile.", "Unreadable DOCX archive.")])
    }
    return data
}

// AppKit's DOCX importer omits DrawingML pictures. Insert unique text anchors in a
// temporary copy, then replace them with native attachments after importing styles.
// The original archive is never written to or extracted to the filesystem.
private func preparePictures(_ url: URL, in temporary: URL) throws -> (URL, [EmbeddedPicture], [EmbeddedLink]) {
    let xml = try XMLDocument(data: archiveMember(url, "word/document.xml"), options: [.nodePreserveWhitespace, .nodeLoadExternalEntitiesNever])
    let drawings = try xml.nodes(forXPath: "//*[local-name()='drawing' or local-name()='pict']")
    let hyperlinks = try xml.nodes(forXPath: "//*[local-name()='hyperlink']")
    guard !drawings.isEmpty || !hyperlinks.isEmpty else { return (url, [], []) }
    guard let relData = try? archiveMember(url, "word/_rels/document.xml.rels"),
          let rels = try? XMLDocument(data: relData, options: [.nodeLoadExternalEntitiesNever]) else { return (url, [], []) }
    var targets: [String: String] = [:]
    var linkTargets: [String: URL] = [:]
    for case let rel as XMLElement in try rels.nodes(forXPath: "//*[local-name()='Relationship']") {
        guard let id = rel.attribute(forName: "Id")?.stringValue,
              let target = rel.attribute(forName: "Target")?.stringValue else { continue }
        if rel.attribute(forName: "TargetMode")?.stringValue == "External" {
            if let targetURL = URL(string: target), ["http", "https", "mailto"].contains(targetURL.scheme?.lowercased() ?? "") {
                linkTargets[id] = targetURL
            }
            continue
        }
        let path = target.hasPrefix("/") ? target : "/word/" + target
        targets[id] = String(URL(fileURLWithPath: path).standardizedFileURL.path.dropFirst())
    }
    var pictures: [EmbeddedPicture] = []
    let prefix = "DOCXIMAGE" + UUID().uuidString.replacingOccurrences(of: "-", with: "")
    for drawing in drawings {
        guard let reference = try drawing.nodes(forXPath: ".//*[local-name()='blip']/@*[local-name()='embed'] | .//*[local-name()='imagedata']/@*[local-name()='id']").first?.stringValue,
              let target = targets[reference], let data = try? archiveMember(url, target),
              let image = NSImage(data: data), let parent = drawing.parent as? XMLElement else { continue }
        var size = image.size
        if let extent = try drawing.nodes(forXPath: ".//*[local-name()='extent' and @cx and @cy]").first as? XMLElement,
           let width = Double(extent.attribute(forName: "cx")?.stringValue ?? ""),
           let height = Double(extent.attribute(forName: "cy")?.stringValue ?? ""), width > 0, height > 0 {
            size = NSSize(width: width / 12700, height: height / 12700)
        }
        let marker = prefix + "N\(pictures.count)END"
        pictures.append(EmbeddedPicture(marker: marker, data: data, size: size))
        let text = XMLElement(name: "w:t", stringValue: marker)
        parent.replaceChild(at: drawing.index, with: text)
    }
    var links: [EmbeddedLink] = []
    for case let hyperlink as XMLElement in hyperlinks {
        guard let id = try hyperlink.nodes(forXPath: "./@*[local-name()='id']").first?.stringValue,
              let target = linkTargets[id] else { continue }
        let start = prefix + "LINK\(links.count)START"
        let end = prefix + "LINK\(links.count)END"
        let startRun = XMLElement(name: "w:r")
        startRun.addChild(XMLElement(name: "w:t", stringValue: start))
        let endRun = XMLElement(name: "w:r")
        endRun.addChild(XMLElement(name: "w:t", stringValue: end))
        hyperlink.insertChild(startRun, at: 0)
        hyperlink.addChild(endRun)
        links.append(EmbeddedLink(start: start, end: end, url: target))
    }
    guard !pictures.isEmpty || !links.isEmpty else { return (url, [], []) }
    try FileManager.default.createDirectory(at: temporary.appendingPathComponent("word"), withIntermediateDirectories: true)
    let copy = temporary.appendingPathComponent("document.docx")
    try FileManager.default.copyItem(at: url, to: copy)
    try xml.xmlData.write(to: temporary.appendingPathComponent("word/document.xml"))
    let zip = Process()
    zip.executableURL = URL(fileURLWithPath: "/usr/bin/zip")
    zip.currentDirectoryURL = temporary
    zip.arguments = ["-q", copy.path, "word/document.xml"]
    zip.standardOutput = FileHandle.nullDevice
    zip.standardError = FileHandle.nullDevice
    try zip.run()
    zip.waitUntilExit()
    guard zip.terminationStatus == 0 else { throw NSError(domain: "LettoreDOCX", code: 3) }
    return (copy, pictures, links)
}

struct WordDocument {
    let text: NSAttributedString
    let pageSize: NSSize
    let margin: CGFloat
    let wordCount: Int

    init(url: URL) throws {
        let ext = url.pathExtension.lowercased()
        guard readableExtensions.contains(ext) else {
            throw NSError(domain: "LettoreDOCX", code: 1, userInfo: [
                NSLocalizedDescriptionKey: L("Scegli un documento .docx, .doc, .odt o .rtf.", "Choose a .docx, .doc, .odt or .rtf document.")
            ])
        }
        let temporary = FileManager.default.temporaryDirectory.appendingPathComponent("lettoredocx-" + UUID().uuidString)
        defer { try? FileManager.default.removeItem(at: temporary) }
        // DOCX gets the picture/link pass; .doc, .odt and .rtf use the native importers as they are.
        let (importURL, pictures, links) = ext == "docx" ? try preparePictures(url, in: temporary) : (url, [], [])
        let types: [String: NSAttributedString.DocumentType] = ["docx": .officeOpenXML, "doc": .docFormat, "odt": .openDocument, "rtf": .rtf]
        var attributes: NSDictionary?
        let imported = try NSMutableAttributedString(url: importURL, options: [
            .documentType: types[ext] ?? .officeOpenXML
        ], documentAttributes: &attributes)
        let metadata = attributes as? [NSAttributedString.DocumentAttributeKey: Any] ?? [:]
        let size = (metadata[.paperSize] as? NSValue)?.sizeValue ?? NSSize(width: 595, height: 842)
        pageSize = NSSize(width: min(2400, max(300, size.width)), height: min(3600, max(400, size.height)))
        margin = min(pageSize.width / 4, max(24, (metadata[.leftMargin] as? NSNumber)?.doubleValue ?? 56))
        for picture in pictures {
            let range = (imported.string as NSString).range(of: picture.marker)
            guard range.location != NSNotFound, let image = NSImage(data: picture.data) else { continue }
            let factor = min(1, (pageSize.width - margin * 2) / max(1, picture.size.width))
            image.size = NSSize(width: picture.size.width * factor, height: picture.size.height * factor)
            let attachment = NSTextAttachment()
            attachment.fileWrapper = FileWrapper(regularFileWithContents: picture.data)
            attachment.attachmentCell = NSTextAttachmentCell(imageCell: image)
            let replacement = NSMutableAttributedString(attachment: attachment)
            let runStyle = imported.attributes(at: range.location, effectiveRange: nil)
            if let style = runStyle[.paragraphStyle] { replacement.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: replacement.length)) }
            imported.replaceCharacters(in: range, with: replacement)
        }
        for link in links {
            let start = (imported.string as NSString).range(of: link.start)
            let end = (imported.string as NSString).range(of: link.end)
            guard start.location != NSNotFound, end.location != NSNotFound, end.location >= NSMaxRange(start) else { continue }
            imported.deleteCharacters(in: end)
            imported.deleteCharacters(in: start)
            let range = NSRange(location: start.location, length: end.location - NSMaxRange(start))
            imported.addAttributes([.link: link.url, .foregroundColor: NSColor.systemBlue, .underlineStyle: NSUnderlineStyle.single.rawValue], range: range)
        }
        text = imported
        var count = 0
        text.string.enumerateSubstrings(in: text.string.startIndex..<text.string.endIndex, options: .byWords) { _, _, _, _ in
            count += 1
        }
        wordCount = count
    }
}

// Keep page coordinates independent of zoom: text, images, and tables scale together.
final class PaperCanvas: NSView {
    override var isFlipped: Bool { true }
    var openDocument: ((URL) -> Void)?

    override init(frame: NSRect) {
        super.init(frame: frame)
        registerForDraggedTypes([.fileURL])
    }
    required init?(coder: NSCoder) { fatalError("init(coder:) has not been implemented") }

    private func documentURL(_ sender: NSDraggingInfo) -> URL? {
        let files = sender.draggingPasteboard.readObjects(forClasses: [NSURL.self], options: [
            .urlReadingFileURLsOnly: true
        ]) as? [URL] ?? []
        return files.first { readableExtensions.contains($0.pathExtension.lowercased()) }
    }
    override func draggingEntered(_ sender: NSDraggingInfo) -> NSDragOperation {
        documentURL(sender) == nil ? [] : .copy
    }
    override func prepareForDragOperation(_ sender: NSDraggingInfo) -> Bool { documentURL(sender) != nil }
    override func performDragOperation(_ sender: NSDraggingInfo) -> Bool {
        guard let url = documentURL(sender) else { return false }
        openDocument?(url)
        return true
    }
}

final class ReaderTextView: NSTextView {
    // Let the canvas handle file drops even when the pointer is over the text.
    override func draggingEntered(_ sender: NSDraggingInfo) -> NSDragOperation { [] }
    override func menu(for event: NSEvent) -> NSMenu? {
        let menu = NSMenu()
        menu.addItem(withTitle: L("Copia", "Copy"), action: #selector(copy(_:)), keyEquivalent: "")
        menu.addItem(withTitle: L("Seleziona tutto", "Select All"), action: #selector(selectAll(_:)), keyEquivalent: "")
        return menu
    }
}

final class ReaderApp: NSObject, NSApplicationDelegate, NSWindowDelegate, NSSearchFieldDelegate, NSTextViewDelegate, NSMenuItemValidation {
    var window: NSWindow!
    let scroll = NSScrollView()
    let canvas = PaperCanvas(frame: .zero)
    let textView = ReaderTextView(frame: .zero)
    let search = NSSearchField()
    let zoomLabel = NSButton(title: "100%", target: nil, action: nil)
    let status = NSTextField(labelWithString: L("Sola lettura", "Read-only"))
    let searchStatus = NSTextField(labelWithString: "")
    let progress = NSProgressIndicator()
    let emptyState = NSStackView()
    let recentMenu = NSMenu(title: L("Apri recenti", "Open Recent"))
    var document: WordDocument?
    var currentURL: URL?
    var pendingURL: URL?
    var matches: [NSRange] = []
    var matchIndex = -1
    private var pageWidth: CGFloat = 595
    private var pageHeight: CGFloat = 842
    private var paperHeight: CGFloat = 842
    private var layingOut = false
    private var loadID = UUID()
    private let loader = DispatchQueue(label: "it.local.lettoredocx.import", qos: .userInitiated)

    func applicationDidFinishLaunching(_ notification: Notification) {
        buildMenu()
        buildWindow()
        window.makeKeyAndOrderFront(nil)
        NSApp.activate(ignoringOtherApps: true)
        if let url = pendingURL {
            pendingURL = nil
            load(url)
        }
    }

    func application(_ sender: NSApplication, openFiles filenames: [String]) {
        guard let path = filenames.first else { sender.reply(toOpenOrPrint: .failure); return }
        let url = URL(fileURLWithPath: path)
        if window == nil { pendingURL = url } else { load(url) }
        sender.reply(toOpenOrPrint: .success)
    }
    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { true }

    private func buildMenu() {
        let main = NSMenu()
        func submenu(_ title: String) -> NSMenu {
            let item = NSMenuItem(title: title, action: nil, keyEquivalent: "")
            let sub = NSMenu(title: title)
            item.submenu = sub
            main.addItem(item)
            return sub
        }
        func add(_ menu: NSMenu, _ title: String, _ action: Selector, _ key: String, _ mods: NSEvent.ModifierFlags = .command, target: AnyObject? = nil) {
            let item = menu.addItem(withTitle: title, action: action, keyEquivalent: key)
            item.keyEquivalentModifierMask = mods
            item.target = target
        }
        let app = submenu(appName)
        add(app, L("Informazioni su \(appName)", "About \(appName)"), #selector(NSApplication.orderFrontStandardAboutPanel(_:)), "")
        app.addItem(.separator())
        add(app, L("Nascondi \(appName)", "Hide \(appName)"), #selector(NSApplication.hide(_:)), "h")
        add(app, L("Nascondi altre", "Hide Others"), #selector(NSApplication.hideOtherApplications(_:)), "h", [.command, .option])
        app.addItem(.separator())
        add(app, L("Esci da \(appName)", "Quit \(appName)"), #selector(NSApplication.terminate(_:)), "q")
        let file = submenu("File")
        add(file, L("Apri…", "Open…"), #selector(openFile), "o", target: self)
        let recentItem = NSMenuItem(title: recentMenu.title, action: nil, keyEquivalent: "")
        recentItem.submenu = recentMenu
        file.addItem(recentItem)
        refreshRecentMenu()
        add(file, L("Chiudi", "Close"), #selector(NSWindow.performClose(_:)), "w")
        file.addItem(.separator())
        add(file, L("Esporta come PDF…", "Export as PDF…"), #selector(exportPDF), "e", [.command, .shift], target: self)
        add(file, L("Stampa…", "Print…"), #selector(printDocument), "p", target: self)
        let edit = submenu(L("Modifica", "Edit"))
        add(edit, L("Copia", "Copy"), #selector(NSText.copy(_:)), "c")
        add(edit, L("Seleziona tutto", "Select All"), #selector(NSText.selectAll(_:)), "a")
        edit.addItem(.separator())
        add(edit, L("Cerca nel documento…", "Find…"), #selector(focusSearch), "f", target: self)
        add(edit, L("Trova successivo", "Find Next"), #selector(findNext), "g", target: self)
        add(edit, L("Trova precedente", "Find Previous"), #selector(findPrevious), "g", [.command, .shift], target: self)
        let view = submenu(L("Vista", "View"))
        add(view, L("Ingrandisci", "Zoom In"), #selector(zoomIn), "+", target: self)
        add(view, L("Riduci", "Zoom Out"), #selector(zoomOut), "-", target: self)
        add(view, L("Dimensioni reali", "Actual Size"), #selector(resetZoom), "0", target: self)
        add(view, L("Adatta alla finestra", "Fit Width"), #selector(fitWidth), "9", target: self)
        view.addItem(.separator())
        add(view, L("Schermo intero", "Enter Full Screen"), #selector(NSWindow.toggleFullScreen(_:)), "f", [.command, .control])
        let windowMenu = submenu(L("Finestra", "Window"))
        add(windowMenu, L("Contrai", "Minimize"), #selector(NSWindow.performMiniaturize(_:)), "m")
        add(windowMenu, L("Ridimensiona", "Zoom"), #selector(NSWindow.performZoom(_:)), "")
        NSApp.windowsMenu = windowMenu
        NSApp.mainMenu = main
    }

    private func buildWindow() {
        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 960, height: 780),
                          styleMask: [.titled, .closable, .miniaturizable, .resizable], backing: .buffered, defer: false)
        window.title = appName
        window.minSize = NSSize(width: 680, height: 420)
        window.delegate = self
        window.isReleasedWhenClosed = false
        window.center()
        window.setFrameAutosaveName("ReaderWindow")
        let content = window.contentView!

        let open = NSButton(title: L("Apri…", "Open…"), target: self, action: #selector(openFile))
        open.toolTip = L("Apri un documento (⌘O)", "Open a document (⌘O)")
        search.placeholderString = L("Cerca nel documento", "Search document")
        search.delegate = self
        search.target = self
        search.action = #selector(findNext)
        search.sendsSearchStringImmediately = false
        search.sendsWholeSearchString = true
        search.widthAnchor.constraint(equalToConstant: 210).isActive = true
        search.setAccessibilityLabel(L("Cerca nel documento", "Search document"))
        let minus = NSButton(title: "−", target: self, action: #selector(zoomOut))
        let plus = NSButton(title: "+", target: self, action: #selector(zoomIn))
        minus.toolTip = L("Riduci (⌘−)", "Zoom out (⌘−)")
        plus.toolTip = L("Ingrandisci (⌘+)", "Zoom in (⌘+)")
        zoomLabel.target = self
        zoomLabel.action = #selector(resetZoom)
        zoomLabel.toolTip = L("Dimensioni reali (⌘0)", "Actual size (⌘0)")
        zoomLabel.widthAnchor.constraint(equalToConstant: 58).isActive = true
        let fit = NSButton(image: NSImage(systemSymbolName: "arrow.left.and.right", accessibilityDescription: L("Adatta alla finestra", "Fit width"))!,
                           target: self, action: #selector(fitWidth))
        fit.toolTip = L("Adatta alla finestra (⌘9)", "Fit width (⌘9)")
        progress.style = .spinning
        progress.controlSize = .small
        progress.isDisplayedWhenStopped = false
        progress.widthAnchor.constraint(equalToConstant: 16).isActive = true
        let spacer = NSView()
        spacer.setContentHuggingPriority(.defaultLow, for: .horizontal)
        let bar = NSStackView(views: [open, progress, spacer, search, minus, zoomLabel, plus, fit])
        bar.orientation = .horizontal
        bar.spacing = 8

        scroll.hasVerticalScroller = true
        scroll.hasHorizontalScroller = true
        scroll.autohidesScrollers = true
        scroll.borderType = .noBorder
        scroll.drawsBackground = true
        scroll.backgroundColor = .underPageBackgroundColor
        scroll.allowsMagnification = true
        scroll.minMagnification = 0.5
        scroll.maxMagnification = 3
        scroll.documentView = canvas
        scroll.contentView.postsBoundsChangedNotifications = true
        canvas.openDocument = { [weak self] url in self?.load(url) }

        textView.isEditable = false
        textView.isSelectable = true
        textView.isRichText = true
        textView.importsGraphics = true
        textView.usesFindPanel = false
        textView.usesFontPanel = false
        textView.isAutomaticLinkDetectionEnabled = false
        textView.usesAdaptiveColorMappingForDarkAppearance = false
        textView.appearance = NSAppearance(named: .aqua)
        textView.backgroundColor = .white
        textView.textColor = .black
        textView.isVerticallyResizable = false
        textView.isHorizontallyResizable = false
        textView.textContainer?.widthTracksTextView = false
        textView.textContainer?.heightTracksTextView = false
        textView.textContainer?.lineFragmentPadding = 0
        textView.delegate = self
        textView.unregisterDraggedTypes()
        textView.setAccessibilityLabel(L("Contenuto del documento, sola lettura", "Document content, read-only"))
        canvas.addSubview(textView)
        textView.isHidden = true

        let icon = NSImageView(image: NSImage(systemSymbolName: "doc.text", accessibilityDescription: nil)!)
        icon.contentTintColor = .tertiaryLabelColor
        icon.symbolConfiguration = NSImage.SymbolConfiguration(pointSize: 44, weight: .ultraLight)
        let title = NSTextField(labelWithString: L("Apri un documento", "Open a document"))
        title.font = .systemFont(ofSize: 21, weight: .medium)
        let hint = NSTextField(labelWithString: L("Trascina qui un file .docx, .doc, .odt o .rtf oppure premi ⌘O", "Drop a .docx, .doc, .odt or .rtf file here or press ⌘O"))
        hint.font = .systemFont(ofSize: 13)
        hint.textColor = .secondaryLabelColor
        emptyState.orientation = .vertical
        emptyState.spacing = 14
        for view in [icon, title, hint] { emptyState.addArrangedSubview(view) }

        status.font = .systemFont(ofSize: 11)
        status.textColor = .secondaryLabelColor
        searchStatus.font = .systemFont(ofSize: 11)
        searchStatus.textColor = .secondaryLabelColor
        let footer = NSStackView(views: [status, NSView(), searchStatus])
        footer.orientation = .horizontal
        let separator = NSBox()
        separator.boxType = .separator
        for view in [bar, separator, scroll, footer, emptyState] {
            view.translatesAutoresizingMaskIntoConstraints = false
            content.addSubview(view)
        }
        NSLayoutConstraint.activate([
            bar.topAnchor.constraint(equalTo: content.topAnchor, constant: 10),
            bar.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 14),
            bar.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -14),
            bar.heightAnchor.constraint(equalToConstant: 30),
            separator.topAnchor.constraint(equalTo: bar.bottomAnchor, constant: 10),
            separator.leadingAnchor.constraint(equalTo: content.leadingAnchor),
            separator.trailingAnchor.constraint(equalTo: content.trailingAnchor),
            scroll.topAnchor.constraint(equalTo: separator.bottomAnchor),
            scroll.leadingAnchor.constraint(equalTo: content.leadingAnchor),
            scroll.trailingAnchor.constraint(equalTo: content.trailingAnchor),
            scroll.bottomAnchor.constraint(equalTo: footer.topAnchor, constant: -5),
            footer.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 14),
            footer.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -14),
            footer.bottomAnchor.constraint(equalTo: content.bottomAnchor, constant: -6),
            footer.heightAnchor.constraint(equalToConstant: 18),
            emptyState.centerXAnchor.constraint(equalTo: scroll.centerXAnchor),
            emptyState.centerYAnchor.constraint(equalTo: scroll.centerYAnchor, constant: -16)
        ])
        NotificationCenter.default.addObserver(self, selector: #selector(viewportChanged),
            name: NSView.boundsDidChangeNotification, object: scroll.contentView)
        NotificationCenter.default.addObserver(self, selector: #selector(viewportChanged),
            name: NSScrollView.didEndLiveMagnifyNotification, object: scroll)
        let saved = UserDefaults.standard.double(forKey: "documentZoom")
        scroll.magnification = saved >= 0.5 && saved <= 3 ? saved : 1.25
        viewportChanged()
    }

    func windowDidResize(_ notification: Notification) { layoutPaper() }
    @objc func viewportChanged() {
        zoomLabel.title = "\(Int((scroll.magnification * 100).rounded()))%"
        UserDefaults.standard.set(scroll.magnification, forKey: "documentZoom")
        layoutPaper()
    }
    private func layoutPaper() {
        guard !layingOut else { return }
        layingOut = true
        defer { layingOut = false }
        let viewport = scroll.contentView.bounds.size
        let size = NSSize(width: max(viewport.width, pageWidth + 48), height: max(viewport.height, paperHeight + 48))
        if canvas.frame.size != size { canvas.setFrameSize(size) }
        textView.frame = NSRect(x: floor((size.width - pageWidth) / 2), y: 24, width: pageWidth, height: paperHeight)
    }

    @objc func openFile() {
        let panel = NSOpenPanel()
        panel.title = L("Apri documento", "Open Document")
        panel.allowedContentTypes = readableTypes
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = false
        if let currentURL { panel.directoryURL = currentURL.deletingLastPathComponent() }
        panel.beginSheetModal(for: window) { [weak self] result in
            if result == .OK, let url = panel.url { self?.load(url) }
        }
    }

    private func refreshRecentMenu() {
        recentMenu.removeAllItems()
        for url in NSDocumentController.shared.recentDocumentURLs.prefix(10) {
            let item = recentMenu.addItem(withTitle: url.lastPathComponent, action: #selector(openRecent(_:)), keyEquivalent: "")
            item.target = self
            item.representedObject = url
            item.toolTip = url.path
        }
        if recentMenu.items.isEmpty {
            let item = recentMenu.addItem(withTitle: L("Nessun documento recente", "No Recent Documents"), action: nil, keyEquivalent: "")
            item.isEnabled = false
        } else {
            recentMenu.addItem(.separator())
            recentMenu.addItem(withTitle: L("Cancella elenco", "Clear Menu"), action: #selector(clearRecent), keyEquivalent: "").target = self
        }
    }
    @objc private func openRecent(_ sender: NSMenuItem) {
        if let url = sender.representedObject as? URL { load(url) }
    }
    @objc private func clearRecent() {
        NSDocumentController.shared.clearRecentDocuments(nil)
        refreshRecentMenu()
    }

    func load(_ url: URL) {
        let request = UUID()
        loadID = request
        progress.startAnimation(nil)
        status.stringValue = L("Apertura di \(url.lastPathComponent)…", "Opening \(url.lastPathComponent)…")
        loader.async { [weak self] in
            let result = Result { try WordDocument(url: url) }
            DispatchQueue.main.async {
                guard let self, self.loadID == request else { return }
                self.progress.stopAnimation(nil)
                switch result {
                case .success(let document): self.display(document, url: url)
                case .failure:
                    self.updateStatus()
                    let alert = NSAlert()
                    alert.messageText = L("Impossibile aprire il documento", "Can't open the document")
                    alert.informativeText = L("Il file potrebbe essere danneggiato, protetto da password o in un formato non supportato.", "The file may be damaged, password-protected or in an unsupported format.")
                    alert.addButton(withTitle: "OK")
                    alert.beginSheetModal(for: self.window)
                }
            }
        }
    }

    func display(_ document: WordDocument, url: URL) {
        self.document = document
        currentURL = url
        pageWidth = document.pageSize.width
        pageHeight = document.pageSize.height
        textView.textContainerInset = NSSize(width: document.margin, height: document.margin)
        textView.textContainer?.containerSize = NSSize(width: pageWidth - document.margin * 2, height: .greatestFiniteMagnitude)
        textView.textStorage?.setAttributedString(document.text)
        if let manager = textView.layoutManager, let container = textView.textContainer {
            manager.ensureLayout(for: container)
            paperHeight = max(pageHeight, ceil(manager.usedRect(for: container).maxY) + document.margin * 2)
        }
        emptyState.isHidden = true
        textView.isHidden = false
        textView.setSelectedRange(NSRange(location: 0, length: 0))
        search.stringValue = ""
        updateMatches()
        window.title = url.lastPathComponent
        window.representedURL = url
        NSDocumentController.shared.noteNewRecentDocumentURL(url)
        refreshRecentMenu()
        updateStatus()
        layoutPaper()
        scroll.contentView.scroll(to: NSPoint(x: max(0, (canvas.bounds.width - scroll.contentView.bounds.width) / 2), y: 0))
        scroll.reflectScrolledClipView(scroll.contentView)
        window.makeFirstResponder(textView)
    }

    private func updateStatus() {
        guard let document else { status.stringValue = L("Sola lettura", "Read-only"); return }
        let count = NumberFormatter.localizedString(from: NSNumber(value: document.wordCount), number: .decimal)
        status.stringValue = document.wordCount == 1 ? L("1 parola · Sola lettura", "1 word · Read-only") : L("\(count) parole · Sola lettura", "\(count) words · Read-only")
    }

    @objc func focusSearch() { window.makeFirstResponder(search) }
    func controlTextDidChange(_ obj: Notification) {
        guard (obj.object as? NSSearchField) === search else { return }
        updateMatches()
        if !matches.isEmpty { selectMatch(0) }
    }
    func updateMatches() {
        if let manager = textView.layoutManager, textView.textStorage != nil {
            manager.removeTemporaryAttribute(.backgroundColor, forCharacterRange: NSRange(location: 0, length: textView.string.utf16.count))
        }
        matches.removeAll(keepingCapacity: true)
        matchIndex = -1
        let needle = search.stringValue
        guard !needle.isEmpty, document != nil else { searchStatus.stringValue = ""; return }
        let text = textView.string as NSString
        var remaining = NSRange(location: 0, length: text.length)
        while remaining.length > 0 {
            let match = text.range(of: needle, options: [.caseInsensitive, .diacriticInsensitive], range: remaining)
            guard match.location != NSNotFound, match.length > 0 else { break }
            matches.append(match)
            textView.layoutManager?.addTemporaryAttribute(.backgroundColor,
                value: NSColor.systemYellow.withAlphaComponent(0.4), forCharacterRange: match)
            let next = NSMaxRange(match)
            remaining = NSRange(location: next, length: text.length - next)
        }
        searchStatus.stringValue = matches.isEmpty ? L("Nessun risultato", "No results") : L("\(matches.count) risultati", "\(matches.count) results")
    }
    private func selectMatch(_ index: Int) {
        guard !matches.isEmpty else { return }
        matchIndex = (index + matches.count) % matches.count
        let range = matches[matchIndex]
        textView.setSelectedRange(range)
        textView.scrollRangeToVisible(range)
        textView.showFindIndicator(for: range)
        searchStatus.stringValue = L("\(matchIndex + 1) di \(matches.count)", "\(matchIndex + 1) of \(matches.count)")
    }
    @objc func findNext() {
        if matches.isEmpty { updateMatches() }
        selectMatch(matchIndex + 1)
    }
    @objc func findPrevious() {
        if matches.isEmpty { updateMatches() }
        selectMatch(matchIndex < 0 ? matches.count - 1 : matchIndex - 1)
    }
    @objc func zoomIn() { applyZoom(scroll.magnification + 0.1) }
    @objc func zoomOut() { applyZoom(scroll.magnification - 0.1) }
    @objc func resetZoom() { applyZoom(1) }
    @objc func fitWidth() { applyZoom((scroll.contentSize.width - 48) / pageWidth) }
    func applyZoom(_ value: CGFloat) {
        let viewport = scroll.contentView.bounds
        scroll.setMagnification(min(3, max(0.5, value)), centeredAt: NSPoint(x: viewport.midX, y: viewport.midY))
        viewportChanged()
    }

    @objc func printDocument() { runPrint(to: nil) }
    @objc func exportPDF() {
        guard let currentURL, document != nil else { return }
        let panel = NSSavePanel()
        panel.allowedContentTypes = [.pdf]
        panel.nameFieldStringValue = currentURL.deletingPathExtension().lastPathComponent + ".pdf"
        panel.directoryURL = currentURL.deletingLastPathComponent()
        panel.beginSheetModal(for: window) { [weak self] result in
            if result == .OK, let url = panel.url { self?.runPrint(to: url) }
        }
    }
    func validateMenuItem(_ item: NSMenuItem) -> Bool {
        if item.action == #selector(exportPDF) || item.action == #selector(printDocument) { return document != nil }
        return true
    }
    /// Prints, or writes a PDF when `pdf` is set, from a separate text view at the document's page size.
    func runPrint(to pdf: URL?) {
        guard let document else { return }
        // Print a separate text view so screen zoom and search highlights never reach the output.
        let printable = NSTextView(frame: NSRect(origin: .zero, size: document.pageSize))
        printable.textStorage?.setAttributedString(document.text)
        printable.appearance = NSAppearance(named: .aqua)
        printable.usesAdaptiveColorMappingForDarkAppearance = false
        printable.textContainerInset = .zero
        let info = NSPrintInfo.shared.copy() as! NSPrintInfo
        info.paperSize = document.pageSize
        info.leftMargin = document.margin
        info.rightMargin = document.margin
        info.topMargin = document.margin
        info.bottomMargin = document.margin
        info.isHorizontallyCentered = false
        info.isVerticallyCentered = false
        let operation = NSPrintOperation(view: printable, printInfo: info)
        if let pdf {
            info.jobDisposition = .save
            info.dictionary()[NSPrintInfo.AttributeKey.jobSavingURL] = pdf
            operation.showsPrintPanel = false
        }
        operation.jobTitle = currentURL?.deletingPathExtension().lastPathComponent ?? L("Documento", "Document")
        operation.runModal(for: window, delegate: nil, didRun: nil, contextInfo: nil)
    }

    func textView(_ textView: NSTextView, clickedOnLink link: Any, at charIndex: Int) -> Bool {
        let url = (link as? URL) ?? (link as? String).flatMap(URL.init(string:))
        if let url, ["http", "https", "mailto"].contains(url.scheme?.lowercased() ?? "") {
            NSWorkspace.shared.open(url)
        }
        return true
    }
}

if CommandLine.arguments.count > 2, CommandLine.arguments[1] == "--check" {
    do {
        let document = try WordDocument(url: URL(fileURLWithPath: CommandLine.arguments[2]))
        var attachments = 0
        var tableParagraphs = 0
        let range = NSRange(location: 0, length: document.text.length)
        document.text.enumerateAttribute(.attachment, in: range) { value, _, _ in if value != nil { attachments += 1 } }
        document.text.enumerateAttribute(.paragraphStyle, in: range) { value, _, _ in
            if let style = value as? NSParagraphStyle, !style.textBlocks.isEmpty { tableParagraphs += 1 }
        }
        print("OK: \(document.text.length) characters, \(document.wordCount) words, \(attachments) images, \(tableParagraphs) table paragraphs")
    } catch {
        fputs("Unreadable document: \(error.localizedDescription)\n", stderr)
        exit(1)
    }
} else {
    let app = NSApplication.shared
    let delegate = ReaderApp()
    app.delegate = delegate
    app.setActivationPolicy(.regular)
    if CommandLine.arguments.count > 1, !CommandLine.arguments[1].hasPrefix("-") {
        delegate.pendingURL = URL(fileURLWithPath: CommandLine.arguments[1])
    }
    #if DEBUG
    if CommandLine.arguments.count > 2, CommandLine.arguments[1] == "--smoke-test" {
        scheduleReaderTests(delegate, fixture: URL(fileURLWithPath: CommandLine.arguments[2]))
    }
    #endif
    app.run()
}
