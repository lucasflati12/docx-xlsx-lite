import AppKit
import UniformTypeIdentifiers

let appName = "DOCX Lite"
let italian = (Locale.preferredLanguages.first ?? "").lowercased().hasPrefix("it")
func L(_ it: String, _ en: String) -> String { italian ? it : en }
let readableExtensions = ["docx", "doc", "odt", "rtf"]
private let readableTypes = readableExtensions.compactMap { UTType(filenameExtension: $0) }

private let wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
let wordParagraphStyleKey = NSAttributedString.Key("DOCXLiteParagraphStyle")
private let wordStyleIDs = ["Normal", "Title", "Heading1", "Heading2", "Heading3"]

let docxPackageKey = NSAttributedString.Key("DOCXLitePackage")
let docxAnchorKey = NSAttributedString.Key("DOCXLiteAnchor")
private let relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
private let retainedAnchorNames: Set<String> = ["commentRangeStart", "commentRangeEnd", "commentReference", "footnoteReference", "endnoteReference", "bookmarkStart", "bookmarkEnd"]

private func wordDescendants(_ node: XMLNode, named name: String) -> [XMLElement] {
    var result: [XMLElement] = [], stack = Array((node.children ?? []).reversed())
    while let child = stack.popLast() {
        if child.localName == name, let element = child as? XMLElement { result.append(element) }
        stack.append(contentsOf: (child.children ?? []).reversed())
    }
    return result
}
private func parsePart(_ data: Data) throws -> XMLDocument {
    try XMLDocument(data: data, options: [.nodePreserveWhitespace, .nodeLoadExternalEntitiesNever])
}
private func partPath(_ target: String, relativeTo part: String = "word/document.xml") -> String {
    let path = target.hasPrefix("/") ? target : "/" + (part as NSString).deletingLastPathComponent + "/" + target
    return String(URL(fileURLWithPath: path).standardizedFileURL.path.dropFirst())
}
private func documentTarget(_ path: String) -> String { path.hasPrefix("word/") ? String(path.dropFirst(5)) : "../" + path }
private func relationshipKind(_ node: XMLElement) -> String { node.attribute(forName: "Type")?.stringValue?.split(separator: "/").last.map(String.init) ?? "" }
private func packageEntries(_ url: URL) throws -> Set<String> {
    let task = Process(); task.executableURL = URL(fileURLWithPath: "/usr/bin/unzip"); task.arguments = ["-Z1", url.path]
    let pipe = Pipe(); task.standardOutput = pipe; task.standardError = FileHandle.nullDevice
    let done = try startProcess(task); let bytes = pipe.fileHandleForReading.readDataToEndOfFile(); done.wait()
    guard task.terminationStatus == 0 else { throw structureError() }
    return Set(String(decoding: bytes, as: UTF8.self).split(separator: "\n").map(String.init).filter { !$0.hasSuffix("/") })
}
private func standaloneXML(_ node: XMLElement, root: XMLElement?) -> Data {
    let copy = node.copy() as! XMLElement
    for namespace in root?.namespaces ?? [] { copy.addNamespace(namespace.copy() as! XMLNode) }
    return Data(copy.xmlString.utf8)
}
private func partText(_ data: Data?) -> String {
    guard let data, let xml = try? parsePart(data), let root = xml.rootElement() else { return "" }
    return wordDescendants(root, named: "p").map { wordDescendants($0, named: "t").compactMap { $0.stringValue }.joined() }.joined(separator: "\n")
}

final class DOCXAnchor: NSObject, NSSecureCoding {
    static var supportsSecureCoding: Bool { true }
    let owner: String, kind: String, id: String, xml: Data, properties: Data?, label: String, detail: String
    init(owner: String, kind: String, id: String, xml: Data, properties: Data? = nil, label: String = "", detail: String = "") {
        self.owner = owner; self.kind = kind; self.id = id; self.xml = xml; self.properties = properties; self.label = label; self.detail = detail
    }
    required convenience init?(coder: NSCoder) {
        guard let owner = coder.decodeObject(of: NSString.self, forKey: "owner") as String?, let kind = coder.decodeObject(of: NSString.self, forKey: "kind") as String?, let id = coder.decodeObject(of: NSString.self, forKey: "id") as String?, let xml = coder.decodeObject(of: NSData.self, forKey: "xml") as Data? else { return nil }
        self.init(owner: owner, kind: kind, id: id, xml: xml, properties: coder.decodeObject(of: NSData.self, forKey: "properties") as Data?, label: coder.decodeObject(of: NSString.self, forKey: "label") as String? ?? "", detail: coder.decodeObject(of: NSString.self, forKey: "detail") as String? ?? "")
    }
    func encode(with coder: NSCoder) {
        coder.encode(owner, forKey: "owner"); coder.encode(kind, forKey: "kind"); coder.encode(id, forKey: "id"); coder.encode(xml, forKey: "xml")
        coder.encode(properties, forKey: "properties"); coder.encode(label, forKey: "label"); coder.encode(detail, forKey: "detail")
    }
    var visible: Bool { ["commentReference", "footnoteReference", "endnoteReference"].contains(kind) }
    var key: String { kind + ":" + id }
}

// Original compressed bytes are immutable during editing/saving. Only the media
// lookup is populated during import, before this object enters attributed text.
final class DOCXPackage: NSObject {
    let identity = UUID().uuidString
    let archive: Data
    let entries: Set<String>
    let parts: [String: Data]
    let relationships: [XMLElement]
    let finalSection: Data?
    let sectionSize: NSSize?
    let leftMargin: CGFloat?
    let headerText: String, footerText: String, headerDetails: String, footerDetails: String
    let unsupported: [String]
    let nextDrawingID: Int
    private let notes: [String: [String: String]]
    let expectedComments: [String: Set<String>]
    let originalFootnotes: Set<String>, originalEndnotes: Set<String>
    private(set) var knownMedia: [Data: String] = [:]

    func relatedPart(_ kind: String) -> String? {
        guard let relation = relationships.first(where: { relationshipKind($0) == kind && $0.attribute(forName: "TargetMode")?.stringValue != "External" }), let target = relation.attribute(forName: "Target")?.stringValue else { return nil }
        return partPath(target)
    }
    func noteText(_ kind: String, id: String) -> String { notes[kind]?[id] ?? "" }
    func recordMedia(_ data: Data, path: String) { if knownMedia[data] == nil { knownMedia[data] = path } }

    init(url: URL) throws {
        archive = try Data(contentsOf: url); entries = try packageEntries(url)
        var loaded: [String: Data] = [:]
        for path in ["word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"] where entries.contains(path) { loaded[path] = try archiveMember(url, path) }
        let rels = try loaded["word/_rels/document.xml.rels"].map(parsePart)
        let sourceRelations = rels?.rootElement()?.children?.compactMap { ($0 as? XMLElement)?.copy() as? XMLElement } ?? []
        relationships = sourceRelations
        let metadata: Set<String> = ["styles", "numbering", "comments", "commentsExtended", "commentsIds", "commentsExtensible", "people", "footnotes", "endnotes", "settings"]
        for relation in sourceRelations where metadata.contains(relationshipKind(relation)) && relation.attribute(forName: "TargetMode")?.stringValue != "External" {
            if let target = relation.attribute(forName: "Target")?.stringValue { let path = partPath(target); if entries.contains(path) { loaded[path] = try archiveMember(url, path) } }
        }
        for path in entries where path.hasPrefix("word/header") && path.hasSuffix(".xml") || path.hasPrefix("word/footer") && path.hasSuffix(".xml") { loaded[path] = try archiveMember(url, path) }
        parts = loaded
        nextDrawingID = (loaded.values.flatMap { data -> [Int] in
            guard let xml = try? parsePart(data) else { return [] }
            return wordDescendants(xml, named: "docPr").compactMap { ($0.attribute(forName: "id")?.stringValue).flatMap(Int.init) }
        }.max() ?? 0) + 1
        var noteTexts: [String: [String: String]] = [:]
        for (kind, singular) in [("comments", "comment"), ("footnotes", "footnote"), ("endnotes", "endnote")] {
            if let relation = sourceRelations.first(where: { relationshipKind($0) == kind }), let target = relation.attribute(forName: "Target")?.stringValue, let data = loaded[partPath(target)], let noteXML = try? parsePart(data), let noteRoot = noteXML.rootElement() {
                for note in wordDescendants(noteRoot, named: singular) {
                    if let id = xmlValue(note, "id") { noteTexts[kind, default: [:]][id] = wordDescendants(note, named: "p").map { wordDescendants($0, named: "t").compactMap { $0.stringValue }.joined() }.joined(separator: "\n") }
                }
            }
        }
        notes = noteTexts
        let document = try parsePart(loaded["word/document.xml"] ?? Data())
        let root = document.rootElement()!
        let sections = wordDescendants(root, named: "sectPr").filter { $0.parent?.localName == "pPr" || $0.parent?.localName == "body" }
        finalSection = sections.last(where: { $0.parent?.localName == "body" }).map { standaloneXML($0, root: root) }
        let first = sections.first
        let size = first?.children?.first(where: { $0.localName == "pgSz" }) as? XMLElement
        if let width = Double(xmlValue(size, "w") ?? ""), let height = Double(xmlValue(size, "h") ?? "") { sectionSize = NSSize(width: width / 20, height: height / 20) } else { sectionSize = nil }
        let margin = first?.children?.first(where: { $0.localName == "pgMar" }) as? XMLElement
        leftMargin = Double(xmlValue(margin, "left") ?? "").map { CGFloat($0 / 20) }
        func preview(_ kind: String) -> (String, String) {
            let references = first?.children?.compactMap { $0 as? XMLElement }.filter { $0.localName == kind + "Reference" } ?? []
            let firstPage = first?.children?.contains { $0.localName == "titlePg" } ?? false
            let preferred = references.first { xmlValue($0, "type") == (firstPage ? "first" : "default") } ?? references.first
            func text(_ reference: XMLElement?) -> String {
                guard let id = xmlValue(reference, "id"), let relation = sourceRelations.first(where: { $0.attribute(forName: "Id")?.stringValue == id }), let target = relation.attribute(forName: "Target")?.stringValue else { return "" }
                return partText(loaded[partPath(target)])
            }
            return (text(preferred).replacingOccurrences(of: "\n", with: " · "), references.map { text($0) }.filter { !$0.isEmpty }.joined(separator: "\n"))
        }
        let header = preview("header"), footer = preview("footer")
        headerText = header.0; headerDetails = header.1; footerText = footer.0; footerDetails = footer.1
        var comments: [String: Set<String>] = [:], footnotes = Set<String>(), endnotes = Set<String>(), names = Set<String>()
        var stack: [XMLNode] = [root]
        while let node = stack.popLast() {
            if let name = node.localName { names.insert(name)
                if let element = node as? XMLElement, let id = xmlValue(element, "id") {
                    if name.hasPrefix("commentRange") || name == "commentReference" { comments[id, default: []].insert(name) }
                    if name == "footnoteReference" { footnotes.insert(id) }
                    if name == "endnoteReference" { endnotes.insert(id) }
                }
            }
            stack.append(contentsOf: node.children ?? [])
        }
        expectedComments = comments; originalFootnotes = footnotes; originalEndnotes = endnotes
        var issues: [String] = []
        let settingsPath = sourceRelations.first(where: { relationshipKind($0) == "settings" })?.attribute(forName: "Target")?.stringValue.map { partPath($0) }
        let settings = settingsPath.flatMap { loaded[$0] }.flatMap { try? parsePart($0) }?.rootElement()
        let trackEnabled = settings.flatMap { wordDescendants($0, named: "trackRevisions").first }.map { !["0", "false", "off"].contains(xmlValue($0, "val") ?? "1") } ?? false
        if !names.isDisjoint(with: ["ins", "del", "moveFrom", "moveTo"]) || names.contains(where: { $0.hasSuffix("Change") || $0.hasPrefix("moveFrom") || $0.hasPrefix("moveTo") }) || trackEnabled { issues.append(L("Revisioni e modifiche tracciate", "Tracked changes")) }
        if !names.isDisjoint(with: ["txbxContent", "shape", "wsp", "wgp"]) { issues.append(L("Caselle di testo e forme", "Text boxes and shapes")) }
        if names.contains("anchor") { issues.append(L("Oggetti e immagini flottanti", "Floating objects and images")) }
        if names.contains("relIds") { issues.append(L("SmartArt", "SmartArt")) }
        if names.contains("chart") { issues.append(L("Grafici", "Charts")) }
        if !names.isDisjoint(with: ["oMath", "oMathPara"]) { issues.append(L("Equazioni", "Equations")) }
        if !names.isDisjoint(with: ["object", "OLEObject", "altChunk"]) { issues.append(L("Oggetti incorporati", "Embedded objects")) }
        if names.contains("sdt") { issues.append(L("Controlli contenuto", "Content controls")) }
        if !names.isDisjoint(with: ["fldChar", "fldSimple", "instrText"]) { issues.append(L("Campi automatici nel testo", "Automatic fields in document text")) }
        if wordDescendants(root, named: "tbl").contains(where: { ancestor($0.parent ?? $0, named: "tbl") != nil }) { issues.append(L("Tabelle annidate o complesse", "Nested or complex tables")) }
        unsupported = issues
    }
}
private func packageInText(_ text: NSAttributedString) -> DOCXPackage? {
    var package: DOCXPackage?
    text.enumerateAttribute(docxPackageKey, in: NSRange(location: 0, length: text.length)) { value, _, stop in
        if let found = value as? DOCXPackage { package = found; stop.pointee = true }
    }
    return package
}

private struct ImportedAnchor { let marker: String; let anchor: DOCXAnchor }
private func prepareRetainedAnchors(_ xml: XMLDocument, package: DOCXPackage?) throws -> [ImportedAnchor] {
    guard let package, let root = xml.rootElement() else { return [] }
    var result: [ImportedAnchor] = [], counters: [String: Int] = [:]
    let nodes = try xml.nodes(forXPath: "//*[local-name()='commentRangeStart' or local-name()='commentRangeEnd' or local-name()='commentReference' or local-name()='footnoteReference' or local-name()='endnoteReference' or local-name()='bookmarkStart' or local-name()='bookmarkEnd' or local-name()='sectPr']")
    for case let node as XMLElement in nodes {
        guard let kind = node.localName else { continue }
        if kind == "sectPr", node.parent?.localName == "body" { node.detach(); continue }
        guard retainedAnchorNames.contains(kind) || kind == "sectPr", let parent = node.parent as? XMLElement else { continue }
        let id = xmlValue(node, "id") ?? "section\(result.count)"
        let properties = kind.hasSuffix("Reference") ? parent.children?.first(where: { $0.localName == "rPr" }) as? XMLElement : nil
        var label = "", detail = ""
        if kind == "commentReference" { detail = package.noteText("comments", id: id) }
        if kind == "footnoteReference" || kind == "endnoteReference" {
            counters[kind, default: 0] += 1; label = String(counters[kind]!)
            detail = package.noteText(kind == "footnoteReference" ? "footnotes" : "endnotes", id: id)
        }
        let anchor = DOCXAnchor(owner: package.identity, kind: kind, id: id, xml: standaloneXML(node, root: root), properties: properties.map { standaloneXML($0, root: root) }, label: label, detail: detail)
        let marker = "DOCXLITEKEEP" + UUID().uuidString.replacingOccurrences(of: "-", with: "")
        let text = wordElement("t"); text.stringValue = marker
        if kind == "sectPr" {
            let paragraph = ancestor(parent, named: "p")!
            node.detach(); let run = wordElement("r"); run.addChild(text); paragraph.addChild(run)
        } else if parent.localName == "r" { parent.replaceChild(at: node.index, with: text) }
        else { let run = wordElement("r"); run.addChild(text); parent.replaceChild(at: node.index, with: run) }
        result.append(ImportedAnchor(marker: marker, anchor: anchor))
    }
    return result
}
private func anchorImage(_ anchor: DOCXAnchor) -> NSImage {
    if anchor.kind == "commentReference" {
        let image = NSImage(systemSymbolName: "text.bubble", accessibilityDescription: L("Commento", "Comment"))!.copy() as! NSImage
        image.size = NSSize(width: 12, height: 12); return image
    }
    let label = anchor.label as NSString, font = NSFont(name: "Arial", size: 9)!
    return NSImage(size: NSSize(width: max(8, ceil(label.size(withAttributes: [.font: font]).width) + 2), height: 12), flipped: false) { _ in
        label.draw(at: NSPoint(x: 1, y: 2), withAttributes: [.font: font, .foregroundColor: NSColor.darkGray]); return true
    }
}
private func restoreRetainedAnchors(_ items: [ImportedAnchor], in text: NSMutableAttributedString, package: DOCXPackage?) {
    for item in items {
        let range = (text.string as NSString).range(of: item.marker)
        guard range.location != NSNotFound else { continue }
        var attributes = text.attributes(at: range.location, effectiveRange: nil)
        attributes[docxAnchorKey] = item.anchor
        let replacement: NSMutableAttributedString
        if item.anchor.visible {
            let attachment = NSTextAttachment(); attachment.attachmentCell = NSTextAttachmentCell(imageCell: anchorImage(item.anchor))
            replacement = NSMutableAttributedString(attachment: attachment)
            attributes[.toolTip] = item.anchor.detail
            attributes[.font] = NSFont(name: "Arial", size: 9)!
            attributes[.baselineOffset] = item.anchor.kind == "commentReference" ? 0 : 3
        } else {
            replacement = NSMutableAttributedString(string: "\u{2060}")
            attributes[.font] = NSFont(name: "Arial", size: 0.1)!
            attributes[.foregroundColor] = NSColor.clear
        }
        replacement.addAttributes(attributes, range: NSRange(location: 0, length: replacement.length))
        text.replaceCharacters(in: range, with: replacement)
    }
    if let package, text.length > 0 { text.addAttribute(docxPackageKey, value: package, range: NSRange(location: 0, length: text.length)) }
}

private struct ExportRetainedAnchor { let position: Int; let marker: String; let anchor: DOCXAnchor }
private struct RetainedAnchors {
    let kept: [ExportRetainedAnchor]
    let positions: Set<Int>
    let comments: Set<String>, footnotes: Set<String>, endnotes: Set<String>
}
private func retainedAnchors(in text: NSAttributedString, package: DOCXPackage?) -> RetainedAnchors {
    var records: [ExportRetainedAnchor] = [], positions = Set<Int>(), seen = Set<String>(), comments: [String: Set<String>] = [:], bookmarks: [String: Set<String>] = [:]
    let source = text.string as NSString
    text.enumerateAttribute(docxAnchorKey, in: NSRange(location: 0, length: text.length)) { value, range, _ in
        guard let anchor = value as? DOCXAnchor else { return }
        for position in range.location..<NSMaxRange(range) where source.character(at: position) == (anchor.visible ? 0xfffc : 0x2060) {
            positions.insert(position)
            guard anchor.owner == package?.identity, seen.insert(anchor.key).inserted else { continue }
            records.append(ExportRetainedAnchor(position: position, marker: "DOCXLITEKEEP" + UUID().uuidString.replacingOccurrences(of: "-", with: ""), anchor: anchor))
            if anchor.kind.hasPrefix("comment") { comments[anchor.id, default: []].insert(anchor.kind) }
            if anchor.kind.hasPrefix("bookmark") { bookmarks[anchor.id, default: []].insert(anchor.kind) }
        }
    }
    let commentIDs = Set(comments.compactMap { id, roles in roles == package?.expectedComments[id] && roles.contains("commentReference") ? id : nil })
    let bookmarkIDs = Set(bookmarks.compactMap { id, roles in roles == Set(["bookmarkStart", "bookmarkEnd"]) ? id : nil })
    let kept = records.filter { record in
        if record.anchor.kind.hasPrefix("comment") { return commentIDs.contains(record.anchor.id) }
        if record.anchor.kind.hasPrefix("bookmark") { return bookmarkIDs.contains(record.anchor.id) }
        return true
    }
    return RetainedAnchors(kept: kept, positions: positions, comments: commentIDs, footnotes: Set(kept.filter { $0.anchor.kind == "footnoteReference" }.map { $0.anchor.id }), endnotes: Set(kept.filter { $0.anchor.kind == "endnoteReference" }.map { $0.anchor.id }))
}
func documentPlainText(_ text: NSAttributedString) -> String {
    let result = NSMutableString(string: ""), source = text.string as NSString
    text.enumerateAttribute(docxAnchorKey, in: NSRange(location: 0, length: text.length)) { value, range, _ in
        guard let anchor = value as? DOCXAnchor else { result.append(source.substring(with: range)); return }
        var start = range.location
        for position in range.location..<NSMaxRange(range) where source.character(at: position) == (anchor.visible ? 0xfffc : 0x2060) {
            if position > start { result.append(source.substring(with: NSRange(location: start, length: position - start))) }; start = position + 1
        }
        if start < NSMaxRange(range) { result.append(source.substring(with: NSRange(location: start, length: NSMaxRange(range) - start))) }
    }
    return result as String
}

private func xmlSignature(_ node: XMLNode) -> String {
    if let element = node as? XMLElement {
        let attrs = (element.attributes ?? []).map { ($0.localName ?? $0.name ?? "") + "=" + ($0.stringValue ?? "") }.sorted().joined(separator: "|")
        return "<" + (element.localName ?? "") + attrs + ">" + (element.children ?? []).map(xmlSignature).joined() + "</>"
    }
    let string = node.stringValue ?? ""
    return string.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty ? "" : string
}
private func entrySignature(_ element: XMLElement, id: String) -> String {
    let copy = element.copy() as! XMLElement
    for attribute in copy.attributes ?? [] where attribute.localName == id || attribute.localName == "default" { attribute.detach() }
    return xmlSignature(copy)
}
private func remapWordValues(_ root: XMLNode, names: Set<String>, values: [String: String]) {
    var stack = [root]
    while let node = stack.popLast() {
        if let element = node as? XMLElement, names.contains(element.localName ?? ""), let old = xmlValue(element, "val"), let new = values[old] { setXML(element, ["w:val": new]) }
        stack.append(contentsOf: node.children ?? [])
    }
}
private func putPart(_ data: Data, path: String, temporary: URL) throws {
    let url = temporary.appendingPathComponent(path)
    try FileManager.default.createDirectory(at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
    try data.write(to: url)
}

private func mergeRetainedPackage(_ package: DOCXPackage, xml: XMLDocument, rels: XMLDocument, content: XMLDocument, temporary: URL, archive: URL, members: inout [String], anchors: RetainedAnchors) throws {
    let generatedRelations = rels.rootElement()?.children?.compactMap { $0 as? XMLElement } ?? []
    let generatedTypes = content.rootElement()?.children?.compactMap { $0 as? XMLElement } ?? []
    let originalRels = try package.parts["word/_rels/document.xml.rels"].map(parsePart) ?? XMLDocument(xmlString: "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>")
    let originalTypes = try parsePart(package.parts["[Content_Types].xml"]!)
    rels.setRootElement(originalRels.rootElement()!.copy() as! XMLElement)
    content.setRootElement(originalTypes.rootElement()!.copy() as! XMLElement)
    var output = Set(members), occupied = package.entries, paths: [String: String] = [:]
    let nonce = UUID().uuidString.replacingOccurrences(of: "-", with: "")
    func freePath(_ path: String) -> String {
        if occupied.insert(path).inserted { return path }
        let parent = (path as NSString).deletingLastPathComponent, name = (path as NSString).lastPathComponent
        let renamed = parent + "/docxlite-" + nonce + "-" + name
        occupied.insert(renamed); return renamed
    }
    for path in members where path.hasPrefix("word/media/") {
        let data = try Data(contentsOf: temporary.appendingPathComponent(path))
        output.remove(path)
        if let original = package.knownMedia[data] { paths[path] = original }
        else {
            let target = freePath(path); paths[path] = target
            if target != path { try putPart(data, path: target, temporary: temporary) }
            output.insert(target)
        }
    }
    var styleMap: [String: String] = [:]
    if output.contains("word/styles.xml") {
        let generated = try parsePart(Data(contentsOf: temporary.appendingPathComponent("word/styles.xml")))
        let target = package.relatedPart("styles") ?? freePath("word/styles.xml")
        let original = try package.parts[target].map(parsePart) ?? XMLDocument(rootElement: wordElement("styles"))
        let root = original.rootElement()!
        if root.namespaces?.isEmpty != false { root.addNamespace(XMLNode.namespace(withName: "w", stringValue: wordNamespace) as! XMLNode) }
        var ids = Set(wordDescendants(root, named: "style").compactMap { xmlValue($0, "styleId") })
        var existing: [String: String] = [:]
        for style in wordDescendants(root, named: "style") { if let id = xmlValue(style, "styleId") { existing[entrySignature(style, id: "styleId")] = id } }
        let additions = wordDescendants(generated.rootElement()!, named: "style")
        for style in additions {
            guard let id = xmlValue(style, "styleId") else { continue }
            let comparison = style.copy() as! XMLElement
            remapWordValues(comparison, names: ["basedOn", "next", "link", "pStyle", "rStyle"], values: styleMap)
            if let match = existing[entrySignature(comparison, id: "styleId")] { styleMap[id] = match; continue }
            var next = id, suffix = 0
            while ids.contains(next) { suffix += 1; next = "DOCXLite" + id + (suffix == 1 ? "" : String(suffix)) }
            ids.insert(next); styleMap[id] = next
        }
        for style in additions {
            guard let id = xmlValue(style, "styleId"), let mapped = styleMap[id], !wordDescendants(root, named: "style").contains(where: { xmlValue($0, "styleId") == mapped }) else { continue }
            let copy = style.copy() as! XMLElement
            setXML(copy, ["w:styleId": mapped]); copy.removeAttribute(forName: "w:default")
            remapWordValues(copy, names: ["basedOn", "next", "link", "pStyle", "rStyle"], values: styleMap); root.addChild(copy)
        }
        remapWordValues(xml, names: ["pStyle", "rStyle", "tblStyle"], values: styleMap)
        output.remove("word/styles.xml"); output.insert(target); paths["word/styles.xml"] = target
        try putPart(original.xmlData, path: target, temporary: temporary)
    }
    if output.contains("word/numbering.xml") {
        let generated = try parsePart(Data(contentsOf: temporary.appendingPathComponent("word/numbering.xml")))
        let target = package.relatedPart("numbering") ?? freePath("word/numbering.xml")
        let original = try package.parts[target].map(parsePart) ?? XMLDocument(rootElement: wordElement("numbering"))
        let root = original.rootElement()!
        if root.namespaces?.isEmpty != false { root.addNamespace(XMLNode.namespace(withName: "w", stringValue: wordNamespace) as! XMLNode) }
        remapWordValues(generated, names: ["pStyle", "styleLink", "numStyleLink"], values: styleMap)
        var abstractMap: [String: String] = [:], numberMap: [String: String] = [:]
        for (kind, idName) in [("abstractNum", "abstractNumId"), ("num", "numId")] {
            let oldEntries = wordDescendants(root, named: kind)
            var next = (oldEntries.compactMap { xmlValue($0, idName).flatMap(Int.init) }.max() ?? 0) + 1
            var signatures: [String: String] = [:]
            for entry in oldEntries { if let id = xmlValue(entry, idName) { signatures[entrySignature(entry, id: idName)] = id } }
            for entry in wordDescendants(generated.rootElement()!, named: kind) {
                guard let id = xmlValue(entry, idName) else { continue }
                let copy = entry.copy() as! XMLElement
                if kind == "num" { remapWordValues(copy, names: ["abstractNumId"], values: abstractMap) }
                let signature = entrySignature(copy, id: idName)
                let mapped: String
                if let match = signatures[signature] { mapped = match }
                else {
                    mapped = String(next); next += 1; setXML(copy, ["w:" + idName: mapped])
                    // CT_Numbering orders every abstract definition before num.
                    let following = kind == "abstractNum" ? Set(["num", "numIdMacAtCleanup"]) : Set(["numIdMacAtCleanup"])
                    if let index = root.children?.firstIndex(where: { following.contains($0.localName ?? "") }) { root.insertChild(copy, at: index) }
                    else { root.addChild(copy) }
                    signatures[signature] = mapped
                }
                if kind == "num" { numberMap[id] = mapped } else { abstractMap[id] = mapped }
            }
        }
        remapWordValues(xml, names: ["numId"], values: numberMap)
        output.remove("word/numbering.xml"); output.insert(target); paths["word/numbering.xml"] = target
        try putPart(original.xmlData, path: target, temporary: temporary)
    }
    var usedIDs = Set((rels.rootElement()?.children ?? []).compactMap { ($0 as? XMLElement)?.attribute(forName: "Id")?.stringValue })
    var relationMap: [String: String] = [:]
    for relation in generatedRelations {
        let kind = relationshipKind(relation)
        guard ["image", "hyperlink", "styles", "numbering"].contains(kind), let old = relation.attribute(forName: "Id")?.stringValue, let target = relation.attribute(forName: "Target")?.stringValue else { continue }
        let external = relation.attribute(forName: "TargetMode")?.stringValue == "External"
        let mappedTarget = external ? target : documentTarget(paths[partPath(target)] ?? partPath(target))
        let type = relation.attribute(forName: "Type")?.stringValue
        let match = (rels.rootElement()?.children ?? []).compactMap { $0 as? XMLElement }.first { current in
            guard current.attribute(forName: "Type")?.stringValue == type else { return false }
            if ["styles", "numbering"].contains(kind) { return true }
            let currentExternal = current.attribute(forName: "TargetMode")?.stringValue == "External"
            guard currentExternal == external, let currentTarget = current.attribute(forName: "Target")?.stringValue else { return false }
            return external ? currentTarget == mappedTarget : partPath(currentTarget) == partPath(mappedTarget)
        }
        if let existing = match?.attribute(forName: "Id")?.stringValue { relationMap[old] = existing; continue }
        let id = usedIDs.insert(old).inserted ? old : "docxlite" + UUID().uuidString.replacingOccurrences(of: "-", with: "")
        usedIDs.insert(id); relationMap[old] = id
        let copy = relation.copy() as! XMLElement; setXML(copy, ["Id": id, "Target": mappedTarget]); rels.rootElement()?.addChild(copy)
    }
    var stack: [XMLNode] = [xml]
    while let node = stack.popLast() {
        if let element = node as? XMLElement { for attribute in element.attributes ?? [] where attribute.uri == relationshipNamespace || attribute.prefix == "r" {
            if let old = attribute.stringValue, let mapped = relationMap[old] { attribute.stringValue = mapped }
        } }
        stack.append(contentsOf: node.children ?? [])
    }
    var defaults = Set((content.rootElement()?.children ?? []).compactMap { ($0 as? XMLElement)?.attribute(forName: "Extension")?.stringValue?.lowercased() })
    var overrides = Set((content.rootElement()?.children ?? []).compactMap { ($0 as? XMLElement)?.attribute(forName: "PartName")?.stringValue })
    for item in generatedTypes {
        if item.localName == "Default", let ext = item.attribute(forName: "Extension")?.stringValue, defaults.insert(ext.lowercased()).inserted { content.rootElement()?.addChild(item.copy() as! XMLElement) }
        if item.localName == "Override", let part = item.attribute(forName: "PartName")?.stringValue {
            let old = String(part.dropFirst()), mapped = paths[old] ?? old
            guard output.contains(mapped), overrides.insert("/" + mapped).inserted else { continue }
            let copy = item.copy() as! XMLElement; setXML(copy, ["PartName": "/" + mapped]); content.rootElement()?.addChild(copy)
        }
    }
    try pruneDeletedAnnotations(package, anchors: anchors, temporary: temporary, output: &output)
    // ZIP's update operation replaces only the named edited/generated members.
    // All other original entries, including their local relationship parts, stay.
    try package.archive.write(to: archive)
    members = output.sorted()
}

private func pruneDeletedAnnotations(_ package: DOCXPackage, anchors: RetainedAnchors, temporary: URL, output: inout Set<String>) throws {
    let removedComments = Set(package.expectedComments.keys).subtracting(anchors.comments)
    if !removedComments.isEmpty, let path = package.relatedPart("comments"), let data = package.parts[path] {
        let xml = try parsePart(data), root = xml.rootElement()!
        var removedParagraphs = Set<String>()
        for comment in wordDescendants(root, named: "comment") where removedComments.contains(xmlValue(comment, "id") ?? "") {
            for paragraph in wordDescendants(comment, named: "p") { if let id = xmlValue(paragraph, "paraId") { removedParagraphs.insert(id) } }
            comment.detach()
        }
        try putPart(xml.xmlData, path: path, temporary: temporary); output.insert(path)
        for kind in ["commentsExtended", "commentsIds", "commentsExtensible"] {
            guard let extraPath = package.relatedPart(kind), let extraData = package.parts[extraPath] else { continue }
            let extra = try parsePart(extraData)
            var changed = false
            for node in extra.rootElement()?.children ?? [] {
                guard let element = node as? XMLElement else { continue }
                if let id = xmlValue(element, "paraId"), removedParagraphs.contains(id) { element.detach(); changed = true }
                else if let parent = xmlValue(element, "paraIdParent"), removedParagraphs.contains(parent) { element.attributes?.first(where: { $0.localName == "paraIdParent" })?.detach(); changed = true }
            }
            if changed { try putPart(extra.xmlData, path: extraPath, temporary: temporary); output.insert(extraPath) }
        }
    }
    for (kind, ids) in [("footnotes", package.originalFootnotes.subtracting(anchors.footnotes)), ("endnotes", package.originalEndnotes.subtracting(anchors.endnotes))] where !ids.isEmpty {
        guard let path = package.relatedPart(kind), let data = package.parts[path] else { continue }
        let xml = try parsePart(data)
        for node in xml.rootElement()?.children ?? [] {
            guard let element = node as? XMLElement, ids.contains(xmlValue(element, "id") ?? ""), !["separator", "continuationSeparator", "continuationNotice"].contains(xmlValue(element, "type") ?? "") else { continue }
            element.detach()
        }
        try putPart(xml.xmlData, path: path, temporary: temporary); output.insert(path)
    }
}
private func liftAnchor(_ anchor: XMLElement) {
    while let parent = anchor.parent as? XMLElement, parent.localName != "p", let grand = parent.parent as? XMLElement {
        let children = parent.children ?? []
        guard let position = children.firstIndex(where: { $0 === anchor }), let index = (grand.children ?? []).firstIndex(where: { $0 === parent }) else { return }
        func wrapper(_ values: ArraySlice<XMLNode>) -> XMLElement? {
            guard !values.isEmpty else { return nil }
            let copy = XMLElement(name: parent.name ?? "w:hyperlink")
            for attribute in parent.attributes ?? [] { copy.addAttribute(attribute.copy() as! XMLNode) }
            for child in values { child.detach(); copy.addChild(child) }; return copy
        }
        let before = wrapper(children[..<position]), after = wrapper(children[(position + 1)...])
        anchor.detach(); parent.detach()
        var replacement: [XMLNode] = []
        if let before { replacement.append(before) }; replacement.append(anchor); if let after { replacement.append(after) }
        for (offset, node) in replacement.enumerated() { grand.insertChild(node, at: index + offset) }
    }
}
private func writeRetainedAnchors(_ anchors: RetainedAnchors, package: DOCXPackage?, index: ExportXMLIndex, xml: XMLDocument) throws {
    for record in anchors.kept {
        guard let run = index.anchors[record.marker], let paragraph = ancestor(run, named: "p") else { throw structureError() }
        let original = try XMLElement(xmlString: String(decoding: record.anchor.xml, as: UTF8.self))
        if record.anchor.kind == "sectPr" {
            let properties = xmlChild("pPr", of: paragraph, first: true)
            for child in properties.children ?? [] where child.localName == "sectPr" { child.detach() }
            properties.addChild(original); run.detach()
        } else if record.anchor.visible {
            var children: [XMLNode] = []
            if let properties = record.anchor.properties { children.append(try XMLElement(xmlString: String(decoding: properties, as: UTF8.self))) }
            else if record.anchor.kind != "commentReference" { let props = wordElement("rPr"); props.addChild(wordElement("vertAlign", ["w:val": "superscript"])); children.append(props) }
            children.append(original); replaceXMLChildren(run, with: children)
        } else {
            liftAnchor(run)
            if let parent = run.parent as? XMLElement, let position = (parent.children ?? []).firstIndex(where: { $0 === run }) { run.detach(); parent.insertChild(original, at: position) }
        }
    }
    if let section = package?.finalSection, let body = xml.rootElement()?.children?.first(where: { $0.localName == "body" }) as? XMLElement {
        for child in body.children ?? [] where child.localName == "sectPr" { child.detach() }
        body.addChild(try XMLElement(xmlString: String(decoding: section, as: UTF8.self)))
    }
}

// UI fonts have private PostScript names. They are never valid document families.
func documentFont(_ font: NSFont) -> NSFont {
    guard font.fontName.hasPrefix(".") || (font.familyName?.hasPrefix(".") ?? false) else { return font }
    let base = NSFont(name: "Arial", size: font.pointSize)!
    var traits: NSFontTraitMask = []
    if font.fontDescriptor.symbolicTraits.contains(.bold) { traits.insert(.boldFontMask) }
    if font.fontDescriptor.symbolicTraits.contains(.italic) { traits.insert(.italicFontMask) }
    return NSFontManager.shared.convert(base, toHaveTrait: traits)
}
private func normalizeDocumentFonts(_ text: NSMutableAttributedString) {
    text.enumerateAttribute(.font, in: NSRange(location: 0, length: text.length)) { value, range, _ in
        if let font = value as? NSFont { text.addAttribute(.font, value: documentFont(font), range: range) }
    }
}

// Table/list objects are copied and normalized by TextKit during layout. Compare
// their document values rather than NSObject identities or transient layout data.
private func documentState(_ text: NSAttributedString) -> NSAttributedString {
    let result = NSMutableAttributedString(attributedString: text)
    let keys: Set<NSAttributedString.Key> = [.font, .foregroundColor, .backgroundColor, .paragraphStyle, .underlineStyle, .strikethroughStyle, .link, .attachment, .kern, .baselineOffset, wordParagraphStyleKey, docxAnchorKey]
    var tables: [ObjectIdentifier: Int] = [:]
    text.enumerateAttributes(in: NSRange(location: 0, length: text.length)) { values, range, _ in
        var attributes = values.filter { keys.contains($0.key) }
        if let font = attributes[.font] as? NSFont { attributes[.font] = documentFont(font) }
        if let paragraph = values[.paragraphStyle] as? NSParagraphStyle {
            let style = paragraph.mutableCopy() as! NSMutableParagraphStyle
            var signature: [String] = []
            for case let block as NSTextTableBlock in paragraph.textBlocks {
                let id = ObjectIdentifier(block.table)
                if tables[id] == nil { tables[id] = tables.count }
                signature.append("t\(tables[id]!):\(block.table.numberOfColumns):\(block.startingRow):\(block.startingColumn):\(block.rowSpan):\(block.columnSpan)")
            }
            signature += paragraph.textLists.map { "l\($0.markerFormat.rawValue):\($0.startingItemNumber)" }
            style.textBlocks = []; style.textLists = []
            attributes[.paragraphStyle] = style
            if !signature.isEmpty { attributes[NSAttributedString.Key("DOCXLiteStructure")] = signature.joined(separator: ";") }
        }
        if let attachment = values[.attachment] as? NSTextAttachment {
            let image = (attachment.attachmentCell as? NSTextAttachmentCell)?.image
            attributes[.attachment] = [(attachment.fileWrapper?.regularFileContents ?? Data()) as NSData, NSValue(size: image?.size ?? .zero)] as NSArray
        }
        result.setAttributes(attributes, range: range)
    }
    return result
}

private struct ExportParagraph {
    let range: NSRange
    let marker: String
    let style: NSParagraphStyle
    let heading: String?
    let listPrefix: String
}
private func exportParagraphs(_ text: NSAttributedString) -> [ExportParagraph] {
    let source = text.string as NSString
    var paragraphs: [ExportParagraph] = [], offset = 0
    while offset < source.length {
        let range = source.paragraphRange(for: NSRange(location: offset, length: 0))
        let attrs = text.attributes(at: offset, effectiveRange: nil)
        var paragraphStyle = attrs[.paragraphStyle] as? NSParagraphStyle
        text.enumerateAttribute(.paragraphStyle, in: range) { value, _, _ in
            guard let style = value as? NSParagraphStyle else { return }
            if paragraphStyle == nil || (paragraphStyle!.textBlocks.isEmpty && !style.textBlocks.isEmpty) || (paragraphStyle!.textLists.isEmpty && !style.textLists.isEmpty) { paragraphStyle = style }
        }
        let style = paragraphStyle ?? .default
        let line = source.substring(with: range)
        var prefix = ""
        if !style.textLists.isEmpty, line.hasPrefix("\t"), let end = line.dropFirst().firstIndex(of: "\t") { prefix = String(line[...end]) }
        let heading = (attrs[wordParagraphStyleKey] as? String).flatMap { wordStyleIDs.contains($0) ? $0 : nil }
        paragraphs.append(ExportParagraph(range: range, marker: "DOCXLITEPARA" + UUID().uuidString.replacingOccurrences(of: "-", with: ""), style: style, heading: heading, listPrefix: prefix))
        offset = NSMaxRange(range)
    }
    return paragraphs
}
private func setXML(_ element: XMLElement, _ attributes: [String: String]) {
    for (key, value) in attributes {
        let pieces = key.split(separator: ":"), local = String(pieces.last!)
        let prefix = pieces.count > 1 ? String(pieces[0]) : ""
        let uri = prefix == "w" ? wordNamespace : prefix == "r" ? relationshipNamespace : ""
        // Parsed attributes carry namespace URIs; newly built attributes may
        // carry only a prefix. addAttribute alone can duplicate the parsed one.
        if let existing = element.attributes?.first(where: {
            $0.name == key || ($0.localName == local && (($0.prefix ?? "") == prefix || (!uri.isEmpty && $0.uri == uri)))
        }) { existing.stringValue = value }
        else { element.addAttribute(XMLNode.attribute(withName: key, stringValue: value) as! XMLNode) }
    }
}
private func wordElement(_ name: String, _ attributes: [String: String] = [:]) -> XMLElement {
    let element = XMLElement(name: "w:" + name); setXML(element, attributes); return element
}
private func ancestor(_ node: XMLNode, named name: String) -> XMLElement? {
    var current: XMLNode? = node
    while let value = current { if value.localName == name { return value as? XMLElement }; current = value.parent }
    return nil
}
private func registerPart(_ path: String, type: String, relationship: String, rels: XMLDocument, content: XMLDocument) {
    let uri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + relationship
    if !(rels.rootElement()?.children ?? []).contains(where: { ($0 as? XMLElement)?.attribute(forName: "Type")?.stringValue == uri }) {
        let rel = XMLElement(name: "Relationship")
        setXML(rel, ["Id": "docxlite" + UUID().uuidString.replacingOccurrences(of: "-", with: ""), "Type": uri, "Target": String(path.dropFirst(5))])
        rels.rootElement()?.addChild(rel)
    }
    if !(content.rootElement()?.children ?? []).contains(where: { ($0 as? XMLElement)?.attribute(forName: "PartName")?.stringValue == "/" + path }) {
        let item = XMLElement(name: "Override"); setXML(item, ["PartName": "/" + path, "ContentType": type]); content.rootElement()?.addChild(item)
    }
}
private func stripExportPrefix(_ prefix: String, from paragraph: XMLElement) throws {
    var remaining = prefix.utf16.count
    var stack = Array((paragraph.children ?? []).reversed())
    while remaining > 0, let node = stack.popLast() {
        if node.localName == "tab" { node.detach(); remaining -= 1 }
        else if node.localName == "t" {
            let string = (node.stringValue ?? "") as NSString
            let count = min(remaining, string.length)
            node.stringValue = string.substring(from: count); remaining -= count
        } else { stack.append(contentsOf: (node.children ?? []).reversed()) }
    }
}
private func orderParagraphProperties(_ xml: XMLDocument) throws {
    let names = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
    for case let properties as XMLElement in try xml.nodes(forXPath: "//*[local-name()='pPr']") {
        let children = (properties.children ?? []).enumerated().sorted {
            let left = names.firstIndex(of: $0.element.localName ?? "") ?? names.count
            let right = names.firstIndex(of: $1.element.localName ?? "") ?? names.count
            return left == right ? $0.offset < $1.offset : left < right
        }.map { $0.element }
        for child in children { child.detach() }
        for child in children { properties.addChild(child) }
    }
}
private func structureError() -> NSError {
    NSError(domain: appName, code: 9, userInfo: [NSLocalizedDescriptionKey: L("Questa struttura di tabella non può essere salvata fedelmente. Il file originale non è stato modificato.", "This table structure cannot be saved faithfully. The original file was not changed.")])
}

private func writeStructures(_ paragraphs: [ExportParagraph], xml: XMLDocument, rels: XMLDocument, content: XMLDocument, temporary: URL, members: inout [String], width: CGFloat, index: ExportXMLIndex, profile: DOCXSaveProfile? = nil) throws {
    var mapped: [(ExportParagraph, XMLElement)] = []
    for record in paragraphs {
        guard let anchor = index.anchors[record.marker], let paragraph = ancestor(anchor, named: "p") else { throw structureError() }
        anchor.detach(); mapped.append((record, paragraph))
    }
    profile?.mark("structure anchors")
    if mapped.contains(where: { $0.0.heading != nil && $0.0.heading != "Normal" }) {
        let styles = XMLDocument(rootElement: wordElement("styles")); styles.rootElement()?.addNamespace(XMLNode.namespace(withName: "w", stringValue: wordNamespace) as! XMLNode)
        for (index, id) in wordStyleIDs.enumerated() {
            let style = wordElement("style", ["w:type": "paragraph", "w:styleId": id])
            if index == 0 { setXML(style, ["w:default": "1"]) }
            style.addChild(wordElement("name", ["w:val": index >= 2 ? "heading \(index - 1)" : id]))
            if index > 0 { style.addChild(wordElement("basedOn", ["w:val": "Normal"])); style.addChild(wordElement("next", ["w:val": "Normal"])) }
            style.addChild(wordElement("qFormat"))
            let properties = wordElement("pPr")
            properties.addChild(wordElement("spacing", ["w:before": index == 0 ? "0" : "200", "w:after": index == 0 ? "120" : "240"]))
            if index >= 2 { properties.addChild(wordElement("keepNext")); properties.addChild(wordElement("outlineLvl", ["w:val": String(index - 2)])) }
            style.addChild(properties)
            let run = wordElement("rPr"); run.addChild(wordElement("rFonts", ["w:ascii": "Arial", "w:hAnsi": "Arial"]))
            if index >= 2 { run.addChild(wordElement("b")) }
            run.addChild(wordElement("sz", ["w:val": String([24, 60, 48, 40, 32][index])]))
            style.addChild(run); styles.rootElement()?.addChild(style)
        }
        for (record, paragraph) in mapped { if let heading = record.heading, heading != "Normal" { setXML(xmlChild("pStyle", of: xmlChild("pPr", of: paragraph, first: true), first: true), ["w:val": heading]) } }
        try orderParagraphProperties(styles)
        let path = "word/styles.xml"; try styles.xmlData.write(to: temporary.appendingPathComponent(path)); members.append(path)
        registerPart(path, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", relationship: "styles", rels: rels, content: content)
    }
    var listIDs: [ObjectIdentifier: Int] = [:], formats: [Int: [Int: String]] = [:], starts: [Int: Int] = [:]
    for (record, paragraph) in mapped {
        guard let root = record.style.textLists.first, let list = record.style.textLists.last else { continue }
        let key = ObjectIdentifier(root)
        if listIDs[key] == nil { listIDs[key] = listIDs.count + 1 }
        let id = listIDs[key]!
        let level = min(8, max(record.style.textLists.count - 1, Int(((record.style.headIndent - 30) / 24).rounded())))
        let format = list.markerFormat.rawValue.contains("decimal") ? "decimal" : "bullet"
        formats[id, default: [:]][level] = format; starts[id] = max(1, root.startingItemNumber)
        let properties = xmlChild("pPr", of: paragraph, first: true)
        let number = xmlChild("numPr", of: properties)
        number.addChild(wordElement("ilvl", ["w:val": String(level)])); number.addChild(wordElement("numId", ["w:val": String(id)]))
        let indent = xmlChild("ind", of: properties)
        for name in ["w:first-line", "w:firstLine", "w:left", "w:hanging"] { indent.removeAttribute(forName: name) }
        setXML(indent, ["w:left": String(600 + level * 480), "w:hanging": "300"])
        try stripExportPrefix(record.listPrefix, from: paragraph)
    }
    if !listIDs.isEmpty {
        let numbering = XMLDocument(rootElement: wordElement("numbering")); numbering.rootElement()?.addNamespace(XMLNode.namespace(withName: "w", stringValue: wordNamespace) as! XMLNode)
        for id in formats.keys.sorted() {
            let abstract = wordElement("abstractNum", ["w:abstractNumId": String(id)])
            abstract.addChild(wordElement("multiLevelType", ["w:val": "multilevel"]))
            for level in 0...8 {
                let format = formats[id]?[level] ?? formats[id]?[0] ?? "decimal"
                let definition = wordElement("lvl", ["w:ilvl": String(level)])
                definition.addChild(wordElement("start", ["w:val": String(level == 0 ? starts[id] ?? 1 : 1)]))
                definition.addChild(wordElement("numFmt", ["w:val": format]))
                definition.addChild(wordElement("lvlText", ["w:val": format == "decimal" ? "%\(level + 1)." : "•"]))
                definition.addChild(wordElement("lvlJc", ["w:val": "left"]))
                let properties = wordElement("pPr"), tabs = wordElement("tabs")
                tabs.addChild(wordElement("tab", ["w:val": "num", "w:pos": String(600 + level * 480)])); properties.addChild(tabs)
                properties.addChild(wordElement("ind", ["w:left": String(600 + level * 480), "w:hanging": "300"]))
                definition.addChild(properties); abstract.addChild(definition)
            }
            numbering.rootElement()?.addChild(abstract)
        }
        for id in formats.keys.sorted() {
            let number = wordElement("num", ["w:numId": String(id)]); number.addChild(wordElement("abstractNumId", ["w:val": String(id)])); numbering.rootElement()?.addChild(number)
        }
        let path = "word/numbering.xml"; try numbering.xmlData.write(to: temporary.appendingPathComponent(path)); members.append(path)
        registerPart(path, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", relationship: "numbering", rels: rels, content: content)
    }
    profile?.mark("structure styles/lists")
    var groups: [ObjectIdentifier: [(NSTextTableBlock, XMLElement)]] = [:], order: [ObjectIdentifier] = []
    for (record, paragraph) in mapped where !record.style.textBlocks.isEmpty {
        guard record.style.textBlocks.count == 1, let block = record.style.textBlocks.first as? NSTextTableBlock else { throw structureError() }
        let key = ObjectIdentifier(block.table)
        if groups[key] == nil { order.append(key) }
        groups[key, default: []].append((block, paragraph))
    }
    var parents: [ObjectIdentifier: (XMLElement, [XMLNode])] = [:]
    var tableReplacements: [ObjectIdentifier: XMLElement] = [:]
    var tableParagraphs = Set<ObjectIdentifier>()
    for key in order {
        let records = groups[key]!
        guard let first = records.first, let parent = first.1.parent as? XMLElement, records.allSatisfy({ $0.1.parent === parent }) else { throw structureError() }
        let firstID = ObjectIdentifier(first.1), lastID = ObjectIdentifier(records.last!.1)
        guard let start = index.paragraphPositions[firstID], let end = index.paragraphPositions[lastID], end - start + 1 == records.count else { throw structureError() }
        let parentID = ObjectIdentifier(parent)
        if parents[parentID] == nil { parents[parentID] = (parent, parent.children ?? []) }
        tableParagraphs.formUnion(records.map { ObjectIdentifier($0.1) })
        let columns = first.0.table.numberOfColumns, rows = records.map { $0.0.startingRow + $0.0.rowSpan }.max() ?? 0
        guard columns > 0, columns <= 63, rows > 0, rows * columns <= 100_000 else { throw structureError() }
        var cells: [String: (NSTextTableBlock, [XMLElement])] = [:], occupied: [String: String] = [:]
        for (block, paragraph) in records {
            let cellKey = "\(block.startingRow):\(block.startingColumn)"
            guard block.startingRow >= 0, block.startingColumn >= 0, block.rowSpan > 0, block.columnSpan > 0, block.startingColumn + block.columnSpan <= columns else { throw structureError() }
            if cells[cellKey] == nil {
                cells[cellKey] = (block, [])
                for row in block.startingRow..<(block.startingRow + block.rowSpan) { for column in block.startingColumn..<(block.startingColumn + block.columnSpan) {
                    let slot = "\(row):\(column)"; guard occupied[slot] == nil else { throw structureError() }; occupied[slot] = cellKey
                } }
            }
            cells[cellKey]!.1.append(paragraph)
        }
        let table = wordElement("tbl"), properties = wordElement("tblPr"), grid = wordElement("tblGrid")
        properties.addChild(wordElement("tblW", ["w:w": String(Int(width * 20)), "w:type": "dxa"]))
        table.addChild(properties)
        for _ in 0..<columns { grid.addChild(wordElement("gridCol", ["w:w": String(Int(width * 20 / CGFloat(columns)))])) }
        table.addChild(grid); tableReplacements[firstID] = table
        for row in 0..<rows {
            let tr = wordElement("tr"); table.addChild(tr)
            var column = 0
            while column < columns {
                let cell = wordElement("tc"), cellProperties = wordElement("tcPr")
                cell.addChild(cellProperties); tr.addChild(cell)
                if let cellKey = occupied["\(row):\(column)"], let (block, paragraphs) = cells[cellKey] {
                    guard block.startingColumn == column else { throw structureError() }
                    cellProperties.addChild(wordElement("tcW", ["w:w": String(Int(width * 20 * CGFloat(block.columnSpan) / CGFloat(columns))), "w:type": "dxa"]))
                    if block.columnSpan > 1 { cellProperties.addChild(wordElement("gridSpan", ["w:val": String(block.columnSpan)])) }
                    if block.rowSpan > 1 { cellProperties.addChild(wordElement("vMerge", row == block.startingRow ? ["w:val": "restart"] : [:])) }
                    let borders = wordElement("tcBorders")
                    let edges: [(String, NSRectEdge)] = [("top", .minY), ("left", .minX), ("bottom", .maxY), ("right", .maxX)]
                    for (name, edge) in edges {
                        let size = max(block.width(for: .border, edge: edge), block.table.width(for: .border, edge: edge))
                        borders.addChild(wordElement(name, ["w:val": size > 0 ? "single" : "nil", "w:sz": String(max(2, Int((size * 8).rounded()))), "w:color": "auto"]))
                    }
                    cellProperties.addChild(borders)
                    if row == block.startingRow { for paragraph in paragraphs { paragraph.detach(); cell.addChild(paragraph) } }
                    if cell.children?.last?.localName != "p" { cell.addChild(wordElement("p")) }
                    column += block.columnSpan
                } else { cell.addChild(wordElement("p")); column += 1 }
            }
        }
    }
    for (_, (parent, children)) in parents {
        var output: [XMLNode] = []
        for child in children {
            let id = ObjectIdentifier(child)
            if let table = tableReplacements[id] { output.append(table) }
            else if !tableParagraphs.contains(id) { output.append(child) }
        }
        replaceXMLChildren(parent, with: output)
    }
}


private struct EmbeddedParagraph {
    let start: String
    let end: String
    let heading: String?
    let lists: [NSTextList]
    let level: Int
}
private func xmlValue(_ element: XMLElement?, _ name: String) -> String? {
    element?.attributes?.first(where: { $0.localName == name })?.stringValue
}
private func canonicalHeading(_ value: String) -> String? {
    let key = value.lowercased().replacingOccurrences(of: " ", with: "")
    return wordStyleIDs.first { $0.lowercased() == key }
}
// AppKit's DOCX importer ignores styles.xml (Title, Heading, List Bullet… come in as plain text) and treats
// switched-off toggles such as <w:b w:val="0"/> as switched ON. On the temporary import copy only, resolve
// docDefaults + paragraph/character styles (basedOn chain) into direct properties and drop "off" toggles.
// The user's file and the saved package are untouched by this pass.
private let runPropertyOrder = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
private let paragraphPropertyOrder = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
private let toggleRunProperties: Set<String> = ["b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "vanish", "webHidden", "noProof", "snapToGrid", "specVanish"]
// Paragraph-level properties a style may contribute; tabs/borders/shading stay with AppKit's defaults.
private let inheritedParagraphProperties: Set<String> = ["keepNext", "keepLines", "pageBreakBefore", "numPr", "spacing", "ind", "contextualSpacing", "jc"]

private func isSwitchedOff(_ element: XMLElement) -> Bool {
    guard toggleRunProperties.contains(element.localName ?? "") else { return false }
    return ["0", "false", "off"].contains((xmlValue(element, "val") ?? "1").lowercased())
}

private func sortProperties(_ container: XMLElement, order: [String]) {
    let children = (container.children ?? [])
    let ranked = children.enumerated().sorted { a, b in
        let ra = order.firstIndex(of: a.element.localName ?? "") ?? order.count, rb = order.firstIndex(of: b.element.localName ?? "") ?? order.count
        return ra != rb ? ra < rb : a.offset < b.offset
    }.map { $0.element }
    guard ranked.map({ ObjectIdentifier($0) }) != children.map({ ObjectIdentifier($0) }) else { return }
    for child in children { child.detach() }
    for child in ranked { container.addChild(child) }
}

@discardableResult
private func applyStyleFormatting(_ xml: XMLDocument, url: URL, package: DOCXPackage?) throws -> Bool {
    var changed = false
    var styles: [String: XMLElement] = [:], defaultParagraph: String?
    var defaultRun: [String: XMLElement] = [:], defaultParagraphProps: [String: XMLElement] = [:]
    if let data = try? archiveMember(url, package?.relatedPart("styles") ?? "word/styles.xml"),
       let doc = try? XMLDocument(data: data, options: [.nodeLoadExternalEntitiesNever]) {
        for case let style as XMLElement in try doc.nodes(forXPath: "//*[local-name()='style']") {
            guard let id = xmlValue(style, "styleId") else { continue }
            styles[id] = style
            if xmlValue(style, "type") == "paragraph", ["1", "true", "on"].contains(xmlValue(style, "default") ?? "") { defaultParagraph = id }
        }
        for case let property as XMLElement in try doc.nodes(forXPath: "//*[local-name()='docDefaults']/*[local-name()='rPrDefault']/*[local-name()='rPr']/*") {
            if let name = property.localName { defaultRun[name] = property }
        }
        for case let property as XMLElement in try doc.nodes(forXPath: "//*[local-name()='docDefaults']/*[local-name()='pPrDefault']/*[local-name()='pPr']/*") {
            if let name = property.localName, inheritedParagraphProperties.contains(name) { defaultParagraphProps[name] = property }
        }
    }
    // Resolved (run, paragraph) properties per style id, following basedOn (cycle-safe).
    var cache: [String: ([String: XMLElement], [String: XMLElement])] = [:]
    func resolve(_ id: String, _ seen: Set<String> = []) -> ([String: XMLElement], [String: XMLElement]) {
        if let hit = cache[id] { return hit }
        guard let style = styles[id], !seen.contains(id) else { return ([:], [:]) }
        var run: [String: XMLElement] = [:], paragraph: [String: XMLElement] = [:]
        if let base = (style.children?.first(where: { $0.localName == "basedOn" }) as? XMLElement).flatMap({ xmlValue($0, "val") }) {
            (run, paragraph) = resolve(base, seen.union([id]))
        }
        for case let rPr as XMLElement in style.children ?? [] where rPr.localName == "rPr" {
            for case let property as XMLElement in rPr.children ?? [] { if let name = property.localName { run[name] = property } }
        }
        for case let pPr as XMLElement in style.children ?? [] where pPr.localName == "pPr" {
            for case let property as XMLElement in pPr.children ?? [] {
                if let name = property.localName, inheritedParagraphProperties.contains(name) { paragraph[name] = property }
            }
        }
        cache[id] = (run, paragraph)
        return (run, paragraph)
    }
    for case let paragraph as XMLElement in try xml.nodes(forXPath: "//*[local-name()='body']//*[local-name()='p']") where paragraph.uri == wordNamespace {
        let pPr = paragraph.children?.first(where: { $0.localName == "pPr" }) as? XMLElement
        let styleID = (pPr?.children?.first(where: { $0.localName == "pStyle" }) as? XMLElement).flatMap { xmlValue($0, "val") } ?? defaultParagraph
        var (styleRun, styleParagraph) = styleID.map { resolve($0) } ?? ([:], [:])
        styleRun = defaultRun.merging(styleRun) { $1 }
        styleParagraph = defaultParagraphProps.merging(styleParagraph) { $1 }
        // Paragraph properties: only fill what the paragraph does not set itself.
        let missing = styleParagraph.filter { name, _ in pPr?.children?.contains(where: { $0.localName == name }) != true }
        if !missing.isEmpty {
            let target = pPr ?? { let created = wordElement("pPr"); paragraph.insertChild(created, at: 0); return created }()
            for (_, property) in missing { target.addChild(property.copy() as! XMLElement) }
            sortProperties(target, order: paragraphPropertyOrder); changed = true
        }
        for case let run as XMLElement in try paragraph.nodes(forXPath: ".//*[local-name()='r']") where run.uri == wordNamespace {
            // Runs of nested paragraphs (text boxes) are handled with their own paragraph.
            guard ancestor(run.parent ?? run, named: "p") === paragraph else { continue }
            let rPr = run.children?.first(where: { $0.localName == "rPr" }) as? XMLElement
            var inherited = styleRun
            if let character = (rPr?.children?.first(where: { $0.localName == "rStyle" }) as? XMLElement).flatMap({ xmlValue($0, "val") }) {
                inherited.merge(resolve(character).0) { $1 }
            }
            inherited.removeValue(forKey: "rStyle")
            let own = Set((rPr?.children ?? []).compactMap { $0.localName })
            let additions = inherited.filter { !own.contains($0.key) }
            let offToggles = (rPr?.children ?? []).compactMap { $0 as? XMLElement }.filter(isSwitchedOff)
            guard !additions.isEmpty || !offToggles.isEmpty else { continue }
            let target = rPr ?? { let created = wordElement("rPr"); run.insertChild(created, at: 0); return created }()
            for (_, property) in additions { target.addChild(property.copy() as! XMLElement) }
            for case let element as XMLElement in target.children ?? [] where isSwitchedOff(element) { element.detach() }
            sortProperties(target, order: runPropertyOrder); changed = true
        }
    }
    return changed
}

private func prepareParagraphs(_ url: URL, xml: XMLDocument, prefix: String, package: DOCXPackage?) throws -> [EmbeddedParagraph] {
    var styleIDs: [String: String] = [:]
    if let data = try? archiveMember(url, package?.relatedPart("styles") ?? "word/styles.xml"), let styles = try? XMLDocument(data: data, options: [.nodeLoadExternalEntitiesNever]) {
        for case let style as XMLElement in try styles.nodes(forXPath: "//*[local-name()='style']") {
            guard let id = xmlValue(style, "styleId") else { continue }
            let name = try style.nodes(forXPath: "./*[local-name()='name']").first as? XMLElement
            if let heading = canonicalHeading(id) ?? xmlValue(name, "val").flatMap(canonicalHeading) { styleIDs[id] = heading }
        }
    }
    var definitions: [String: [Int: (String, Int)]] = [:], numbers: [String: String] = [:]
    if let data = try? archiveMember(url, package?.relatedPart("numbering") ?? "word/numbering.xml"), let numbering = try? XMLDocument(data: data, options: [.nodeLoadExternalEntitiesNever]) {
        for case let abstract as XMLElement in try numbering.nodes(forXPath: "//*[local-name()='abstractNum']") {
            guard let id = xmlValue(abstract, "abstractNumId") else { continue }
            for case let level as XMLElement in try abstract.nodes(forXPath: "./*[local-name()='lvl']") {
                let depth = Int(xmlValue(level, "ilvl") ?? "0") ?? 0
                let format = try level.nodes(forXPath: "./*[local-name()='numFmt']").first as? XMLElement
                let start = try level.nodes(forXPath: "./*[local-name()='start']").first as? XMLElement
                definitions[id, default: [:]][depth] = (xmlValue(format, "val") ?? "decimal", Int(xmlValue(start, "val") ?? "1") ?? 1)
            }
        }
        for case let number as XMLElement in try numbering.nodes(forXPath: "//*[local-name()='num']") {
            guard let id = xmlValue(number, "numId"), let abstract = try number.nodes(forXPath: "./*[local-name()='abstractNumId']").first as? XMLElement else { continue }
            numbers[id] = xmlValue(abstract, "val")
        }
    }
    var counters: [String: Int] = [:], listObjects: [String: NSTextList] = [:], result: [EmbeddedParagraph] = []
    for case let paragraph as XMLElement in try xml.nodes(forXPath: "//*[local-name()='p']") {
        let properties = paragraph.children?.first(where: { $0.localName == "pPr" }) as? XMLElement
        let style = properties?.children?.first(where: { $0.localName == "pStyle" }) as? XMLElement
        let heading = xmlValue(style, "val").flatMap { styleIDs[$0] ?? canonicalHeading($0) }
        let num = properties?.children?.first(where: { $0.localName == "numPr" }) as? XMLElement
        let numNode = num?.children?.first(where: { $0.localName == "numId" }) as? XMLElement
        let levelNode = num?.children?.first(where: { $0.localName == "ilvl" }) as? XMLElement
        let id = xmlValue(numNode, "val") ?? "0", depth = min(8, max(0, Int(xmlValue(levelNode, "val") ?? "0") ?? 0))
        var lists: [NSTextList] = [], marker: String?
        if id != "0", let abstract = numbers[id] {
            for level in 0...depth {
                let definition = definitions[abstract]?[level] ?? ("decimal", 1)
                let key = id + ":\(level)"
                if listObjects[key] == nil {
                    let list = NSTextList(markerFormat: definition.0 == "bullet" ? .disc : NSTextList.MarkerFormat(rawValue: "{decimal}."), options: 0)
                    list.startingItemNumber = definition.1; listObjects[key] = list
                }
                lists.append(listObjects[key]!)
            }
            let key = id + ":\(depth)", start = definitions[abstract]?[depth]?.1 ?? 1
            counters[key] = (counters[key] ?? start - 1) + 1
            for deeper in (depth + 1)..<9 { counters.removeValue(forKey: id + ":\(deeper)") }
            marker = lists.last!.marker(forItemNumber: counters[key]!)
            // Import one predictable native marker, independent of AppKit's DOCX
            // numbering support. This prevents duplicate visible list prefixes.
            num?.detach()
        }
        guard heading != nil || !lists.isEmpty else { continue }
        let start = prefix + "PARA\(result.count)START", end = prefix + "PARA\(result.count)END"
        let startRun = wordElement("r"); startRun.addChild(wordElement("t")); startRun.child(at: 0)?.stringValue = start
        let position = properties.map { $0.index + 1 } ?? 0
        paragraph.insertChild(startRun, at: position)
        if let marker {
            let run = wordElement("r"); run.addChild(wordElement("tab"))
            let text = wordElement("t"); text.stringValue = marker; run.addChild(text); run.addChild(wordElement("tab"))
            paragraph.insertChild(run, at: position + 1)
        }
        let endRun = wordElement("r"), endText = wordElement("t"); endText.stringValue = end; endRun.addChild(endText); paragraph.addChild(endRun)
        result.append(EmbeddedParagraph(start: start, end: end, heading: heading, lists: lists, level: depth))
    }
    return result
}

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

private struct EmbeddedStyle {
    let start: String
    let end: String
    let color: NSColor?
    let lineMultiple: CGFloat
}

// Foundation's waitUntilExit pumps a run loop; a termination notification avoids
// its polling delay for the many tiny zip/unzip helper processes.
private func startProcess(_ process: Process) throws -> DispatchSemaphore {
    let finished = DispatchSemaphore(value: 0)
    process.terminationHandler = { _ in finished.signal() }
    try process.run()
    return finished
}

func archiveMember(_ archive: URL, _ member: String) throws -> Data {
    let process = Process()
    process.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
    let literal = member.replacingOccurrences(of: "[", with: "\\[").replacingOccurrences(of: "]", with: "\\]").replacingOccurrences(of: "*", with: "\\*").replacingOccurrences(of: "?", with: "\\?")
    process.arguments = ["-p", archive.path, literal]
    let output = Pipe()
    process.standardOutput = output
    process.standardError = FileHandle.nullDevice
    let finished = try startProcess(process)
    let data = output.fileHandleForReading.readDataToEndOfFile()
    finished.wait()
    guard process.terminationStatus == 0 else {
        throw NSError(domain: "LettoreDOCX", code: 2, userInfo: [NSLocalizedDescriptionKey: L("Archivio DOCX non leggibile.", "Unreadable DOCX archive.")])
    }
    return data
}

// AppKit's DOCX importer omits DrawingML pictures. Insert unique text anchors in a
// temporary copy, then replace them with native attachments after importing styles.
// The original archive is never written to or extracted to the filesystem.
private func preparePictures(_ url: URL, in temporary: URL, package: DOCXPackage?) throws -> (URL, [EmbeddedPicture], [EmbeddedLink], [EmbeddedStyle], [EmbeddedParagraph], [ImportedAnchor]) {
    let xml = try XMLDocument(data: archiveMember(url, "word/document.xml"), options: [.nodePreserveWhitespace, .nodeLoadExternalEntitiesNever])
    let retained = try prepareRetainedAnchors(xml, package: package)
    let styled = try applyStyleFormatting(xml, url: url, package: package)
    let drawings = try xml.nodes(forXPath: "//*[local-name()='drawing' or local-name()='pict']")
    let hyperlinks = try xml.nodes(forXPath: "//*[local-name()='hyperlink']")
    let runs = try xml.nodes(forXPath: "//*[local-name()='r']")
    let rels = (try? archiveMember(url, "word/_rels/document.xml.rels")).flatMap { try? XMLDocument(data: $0, options: [.nodeLoadExternalEntitiesNever]) }
    var targets: [String: String] = [:]
    var linkTargets: [String: URL] = [:]
    for case let rel as XMLElement in (try rels?.nodes(forXPath: "//*[local-name()='Relationship']") ?? []) {
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
        package?.recordMedia(data, path: target)
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
    var styles: [EmbeddedStyle] = []
    let namedColors: [String: NSColor] = ["yellow": .yellow, "green": .green, "cyan": .cyan, "magenta": .magenta, "blue": .blue, "red": .red, "black": .black, "white": .white, "darkYellow": .brown, "darkGreen": .systemGreen, "darkBlue": .systemBlue, "darkRed": .systemRed, "lightGray": .lightGray, "darkGray": .darkGray]
    for case let run as XMLElement in runs {
        let fill = try run.nodes(forXPath: "./*[local-name()='rPr']/*[local-name()='shd']/@*[local-name()='fill']").first?.stringValue
        let named = try run.nodes(forXPath: "./*[local-name()='rPr']/*[local-name()='highlight']/@*[local-name()='val']").first?.stringValue
        var color = named.flatMap { namedColors[$0] }
        if let fill, fill.count == 6, let hex = UInt32(fill, radix: 16) {
            color = NSColor(srgbRed: CGFloat((hex >> 16) & 255) / 255, green: CGFloat((hex >> 8) & 255) / 255, blue: CGFloat(hex & 255) / 255, alpha: 1)
        }
        let spacing = try run.nodes(forXPath: "../*[local-name()='pPr']/*[local-name()='spacing']").first as? XMLElement
        let rule = spacing?.attributes?.first(where: { $0.localName == "lineRule" })?.stringValue
        let line = Double(spacing?.attributes?.first(where: { $0.localName == "line" })?.stringValue ?? "") ?? 0
        let multiple = rule == "auto" || rule == nil ? CGFloat(line / 240) : 0
        guard color != nil || multiple > 0 else { continue }
        let start = prefix + "STYLE\(styles.count)START", end = prefix + "STYLE\(styles.count)END"
        let propertyCount = run.child(at: 0)?.localName == "rPr" ? 1 : 0
        run.insertChild(XMLElement(name: "w:t", stringValue: start), at: propertyCount)
        run.addChild(XMLElement(name: "w:t", stringValue: end))
        styles.append(EmbeddedStyle(start: start, end: end, color: color, lineMultiple: multiple))
    }
    let paragraphs = try prepareParagraphs(url, xml: xml, prefix: prefix, package: package)
    guard styled || !pictures.isEmpty || !links.isEmpty || !styles.isEmpty || !paragraphs.isEmpty || !retained.isEmpty || package?.finalSection != nil else { return (url, [], [], [], [], []) }
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
    let finished = try startProcess(zip)
    finished.wait()
    guard zip.terminationStatus == 0 else { throw NSError(domain: "LettoreDOCX", code: 3) }
    return (copy, pictures, links, styles, paragraphs, retained)
}

struct WordDocument {
    let package: DOCXPackage?
    let text: NSAttributedString
    let pageSize: NSSize
    let margin: CGFloat
    let wordCount: Int

    init(text: NSAttributedString) {
        self.package = packageInText(text)
        self.text = text; pageSize = NSSize(width: 595, height: 842); margin = 56; wordCount = 0
    }

    init(url: URL) throws {
        let ext = url.pathExtension.lowercased()
        guard readableExtensions.contains(ext) else {
            throw NSError(domain: "LettoreDOCX", code: 1, userInfo: [
                NSLocalizedDescriptionKey: L("Scegli un documento .docx, .doc, .odt o .rtf.", "Choose a .docx, .doc, .odt or .rtf document.")
            ])
        }
        package = ext == "docx" ? try DOCXPackage(url: url) : nil
        let temporary = FileManager.default.temporaryDirectory.appendingPathComponent("lettoredocx-" + UUID().uuidString)
        defer { try? FileManager.default.removeItem(at: temporary) }
        // DOCX gets the picture/link pass; .doc, .odt and .rtf use the native importers as they are.
        let (importURL, pictures, links, styles, paragraphs, retained) = ext == "docx" ? try preparePictures(url, in: temporary, package: package) : (url, [], [], [], [], [])
        let types: [String: NSAttributedString.DocumentType] = ["docx": .officeOpenXML, "doc": .docFormat, "odt": .openDocument, "rtf": .rtf]
        var attributes: NSDictionary?
        let imported = try NSMutableAttributedString(url: importURL, options: [
            .documentType: types[ext] ?? .officeOpenXML
        ], documentAttributes: &attributes)
        let metadata = attributes as? [NSAttributedString.DocumentAttributeKey: Any] ?? [:]
        let size = package?.sectionSize ?? (metadata[.paperSize] as? NSValue)?.sizeValue ?? NSSize(width: 595, height: 842)
        pageSize = NSSize(width: min(2400, max(300, size.width)), height: min(3600, max(400, size.height)))
        margin = min(pageSize.width / 4, max(24, package?.leftMargin ?? (metadata[.leftMargin] as? NSNumber)?.doubleValue ?? 56))
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
        for style in styles {
            let start = (imported.string as NSString).range(of: style.start), end = (imported.string as NSString).range(of: style.end)
            guard start.location != NSNotFound, end.location != NSNotFound, end.location >= NSMaxRange(start) else { continue }
            imported.deleteCharacters(in: end); imported.deleteCharacters(in: start)
            let range = NSRange(location: start.location, length: end.location - NSMaxRange(start))
            if let color = style.color { imported.addAttribute(.backgroundColor, value: color, range: range) }
            if style.lineMultiple > 0, range.length > 0 {
                let paragraph = (imported.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
                paragraph.lineHeightMultiple = style.lineMultiple
                imported.addAttribute(.paragraphStyle, value: paragraph, range: range)
            }
        }
        for paragraph in paragraphs {
            let start = (imported.string as NSString).range(of: paragraph.start), end = (imported.string as NSString).range(of: paragraph.end)
            guard start.location != NSNotFound, end.location != NSNotFound, end.location >= NSMaxRange(start) else { continue }
            imported.deleteCharacters(in: end); imported.deleteCharacters(in: start)
            let length = end.location - NSMaxRange(start)
            guard length > 0 else { continue }
            let range = (imported.string as NSString).paragraphRange(for: NSRange(location: start.location, length: length))
            if let heading = paragraph.heading { imported.addAttribute(wordParagraphStyleKey, value: heading, range: range) }
            if !paragraph.lists.isEmpty {
                let style = (imported.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
                style.textLists = paragraph.lists; style.headIndent = 30 + CGFloat(paragraph.level) * 24; style.firstLineHeadIndent = CGFloat(paragraph.level) * 24
                style.tabStops = [NSTextTab(textAlignment: .left, location: style.headIndent - 15), NSTextTab(textAlignment: .left, location: style.headIndent)]
                imported.addAttribute(.paragraphStyle, value: style, range: range)
            }
        }
        // AppKit writes list markers as text. Restore native list editing on reopen.
        let source = imported.string as NSString
        source.enumerateSubstrings(in: NSRange(location: 0, length: source.length), options: [.byParagraphs, .substringNotRequired]) { _, range, _, _ in
            let line = source.substring(with: range)
            if range.location < imported.length, imported.attribute(wordParagraphStyleKey, at: range.location, effectiveRange: nil) == nil {
                imported.addAttribute(wordParagraphStyleKey, value: "Normal", range: source.paragraphRange(for: range))
            }
            let numbered = line.range(of: "^\\t[0-9]+\\.\\t", options: .regularExpression) != nil
            guard numbered || line.hasPrefix("\t•\t"), (imported.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle)?.textLists.isEmpty != false else { return }
            let style = (imported.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
            style.textLists = [NSTextList(markerFormat: numbered ? NSTextList.MarkerFormat(rawValue: "{decimal}.") : .disc, options: 0)]
            imported.addAttribute(.paragraphStyle, value: style, range: source.paragraphRange(for: range))
        }
        restoreRetainedAnchors(retained, in: imported, package: package)
        normalizeDocumentFonts(imported)
        text = imported
        var count = 0
        let plain = documentPlainText(text)
        plain.enumerateSubstrings(in: plain.startIndex..<plain.endIndex, options: .byWords) { _, _, _, _ in
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
    var isSelectingWithMouse = false
    var selectionFinished: (() -> Void)?
    var cancelPendingAction: (() -> Bool)?
    override func mouseDown(with event: NSEvent) {
        isSelectingWithMouse = true
        super.mouseDown(with: event)
        isSelectingWithMouse = false
        selectionFinished?()
    }
    override func insertText(_ insertString: Any, replacementRange: NSRange) {
        if typingAttributes[docxAnchorKey] != nil {
            var attrs = typingAttributes
            for key in [docxAnchorKey, .toolTip, .baselineOffset, .attachment] { attrs.removeValue(forKey: key) }
            let source = textStorage
            let point = min(selectedRange().location, max(0, (source?.length ?? 0) - 1))
            var inherited: [NSAttributedString.Key: Any]?
            if let source, source.length > 0 {
                for distance in 0..<source.length {
                    for position in [point - distance, point + distance] where position >= 0 && position < source.length {
                        if source.attribute(docxAnchorKey, at: position, effectiveRange: nil) == nil { inherited = source.attributes(at: position, effectiveRange: nil); break }
                    }
                    if inherited != nil { break }
                }
            }
            attrs[.font] = inherited?[.font] ?? NSFont(name: "Arial", size: 12)!
            attrs[.foregroundColor] = inherited?[.foregroundColor] ?? NSColor.black
            typingAttributes = attrs
        }
        super.insertText(insertString, replacementRange: replacementRange)
    }
    override func cancelOperation(_ sender: Any?) {
        if cancelPendingAction?() != true { super.cancelOperation(sender) }
    }
}

final class ReaderApp: NSObject, NSApplicationDelegate, NSWindowDelegate, NSSearchFieldDelegate, NSTextViewDelegate, NSMenuItemValidation {
    var window: NSWindow!
    let scroll = NSScrollView()
    let canvas = PaperCanvas(frame: .zero)
    let textView = ReaderTextView(frame: .zero)
    let search = NSSearchField()
    let status = NSTextField(labelWithString: "")
    let searchStatus = NSTextField(labelWithString: "")
    let recentMenu = NSMenu(title: L("Apri recenti", "Open Recent"))
    let stylePopup = NSPopUpButton()
    let fontPopup = NSPopUpButton()
    let zoomPopup = NSPopUpButton()
    let sizeField = NSTextField(string: "12")
    var formatButtons: [String: NSButton] = [:]
    var cleanText = NSAttributedString(string: "")
    var closeApproved = false
    private(set) var isSaving = false
    private(set) var lastSaveWasOnMainThread: Bool?
    let saveProgress = NSProgressIndicator()
    private var documentGeneration = UUID()
    private var saveWaiters: [(Bool) -> Void] = []
    private let saver = DispatchQueue(label: "it.local.lettoredocx.save", qos: .userInitiated)
    var needsRewriteWarning = false
    var isDirty: Bool {
        let live = textView.attributedString()
        if live.string != cleanText.string { return true }
        return !live.isEqual(to: cleanText) && !documentState(live).isEqual(to: documentState(cleanText))
    }
    var margin: CGFloat = 56
    var toolbar: NSStackView!
    var toolbarHeld: [NSView] = []
    var toolbarOptional: [NSView] = []
    var toolbarGroups: [[NSView]] = []
    var toolbarSeparators: [NSView] = []
    var toolbarItems: [String: NSView] = [:]
    var toolbarRight: [NSView] = []
    var alignmentPopup: NSPopUpButton!
    var textColorPopup: NSPopUpButton!
    let findBar = NSStackView()
    var findInset: NSLayoutConstraint!
    var searchWidth: NSLayoutConstraint!
    var findRequested = false
    var paintAttributes: [NSAttributedString.Key: Any]?
    var paintSourceRange = NSRange()
    let paintKeys: [NSAttributedString.Key] = [.font, .foregroundColor, .backgroundColor, .underlineStyle, .strikethroughStyle, .baselineOffset, .kern]
    let helpURL = URL(string: "https://github.com/lucasflati12/docx-xlsx-lite")!
    let styleNames = [L("Testo normale", "Normal text"), L("Titolo", "Title"), L("Intestazione 1", "Heading 1"), L("Intestazione 2", "Heading 2"), L("Intestazione 3", "Heading 3")]
    let styleSizes: [CGFloat] = [12, 30, 24, 20, 16]
    var document: WordDocument?
    private var currentPackage: DOCXPackage?
    let headerPreview = NSTextField(labelWithString: "")
    let footerPreview = NSTextField(labelWithString: "")
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
        guard window == nil else { return }
        buildMenu()
        buildWindow()
        window.makeKeyAndOrderFront(nil)
        NSApp.activate(ignoringOtherApps: true)
        newBlank()
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
        let file = submenu(L("File", "File"))
        add(file, L("Nuovo", "New"), #selector(newDocument), "n", target: self)
        add(file, L("Salva", "Save"), #selector(saveDocument), "s", target: self)
        add(file, L("Salva con nome…", "Save As…"), #selector(saveAs), "s", [.command, .shift], target: self)
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
        add(edit, L("Annulla", "Undo"), #selector(undoEdit), "z", target: self)
        add(edit, L("Ripeti", "Redo"), #selector(redoEdit), "z", [.command, .shift], target: self)
        add(edit, L("Taglia", "Cut"), #selector(NSText.cut(_:)), "x")
        add(edit, L("Incolla", "Paste"), #selector(NSText.paste(_:)), "v")
        add(edit, L("Copia", "Copy"), #selector(NSText.copy(_:)), "c")
        add(edit, L("Seleziona tutto", "Select All"), #selector(NSText.selectAll(_:)), "a")
        edit.addItem(.separator())
        add(edit, L("Cerca nel documento…", "Find…"), #selector(focusSearch), "f", target: self)
        add(edit, L("Trova successivo", "Find Next"), #selector(findNext), "g", target: self)
        add(edit, L("Trova precedente", "Find Previous"), #selector(findPrevious), "g", [.command, .shift], target: self)
        add(edit, L("Trova e sostituisci…", "Find and Replace…"), #selector(replaceDialog), "h", [.command, .shift], target: self)
        let view = submenu(L("Visualizza", "View"))
        add(view, L("Ingrandisci", "Zoom In"), #selector(zoomIn), "+", target: self)
        add(view, L("Riduci", "Zoom Out"), #selector(zoomOut), "-", target: self)
        add(view, L("Dimensioni reali", "Actual Size"), #selector(resetZoom), "0", target: self)
        add(view, L("Adatta alla finestra", "Fit Width"), #selector(fitWidth), "9", target: self)
        view.addItem(.separator())
        add(view, L("Schermo intero", "Enter Full Screen"), #selector(NSWindow.toggleFullScreen(_:)), "f", [.command, .control])
        let insert = submenu(L("Inserisci", "Insert"))
        add(insert, L("Link…", "Link…"), #selector(linkDialog), "k", target: self)
        add(insert, L("Immagine…", "Image…"), #selector(imageDialog), "", target: self)
        add(insert, L("Elenco di controllo", "Checklist"), #selector(checklist), "", target: self)
        let format = submenu(L("Formato", "Format"))
        populateFormatMenu(format)
        let tools = submenu(L("Strumenti", "Tools"))
        add(tools, L("Conteggio parole…", "Word Count…"), #selector(countDialog), "c", [.command, .shift], target: self)
        add(tools, L("Controllo ortografico durante la digitazione", "Check Spelling While Typing"), #selector(spellToggle), "", target: self)
        let windowMenu = submenu(L("Finestra", "Window"))
        add(windowMenu, L("Contrai", "Minimize"), #selector(NSWindow.performMiniaturize(_:)), "m")
        add(windowMenu, L("Ridimensiona", "Zoom"), #selector(NSWindow.performZoom(_:)), "")
        NSApp.windowsMenu = windowMenu
        let help = submenu(L("Guida", "Help"))
        add(help, L("Guida di DOCX Lite", "DOCX Lite Help"), #selector(openHelp), "", target: self)
        help.items.first?.representedObject = helpURL
        NSApp.helpMenu = help
        NSApp.mainMenu = main
    }

    private func buildWindow() {
        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 960, height: 780),
                          styleMask: [.titled, .closable, .miniaturizable, .resizable], backing: .buffered, defer: false)
        window.title = appName
        window.minSize = NSSize(width: 900, height: 420)
        window.delegate = self
        window.isReleasedWhenClosed = false
        window.center()
        window.setFrameAutosaveName("ReaderWindow")
        let content = window.contentView!

        let bar = buildToolbar()
        let closeFind = NSButton(image: NSImage(systemSymbolName: "xmark", accessibilityDescription: L("Chiudi ricerca", "Close find"))!, target: self, action: #selector(closeSearch))
        closeFind.bezelStyle = .recessed
        closeFind.toolTip = L("Chiudi ricerca (Esc)", "Close find (Esc)")
        findBar.orientation = .horizontal; findBar.spacing = 8
        findBar.addArrangedSubview(closeFind); findBar.addArrangedSubview(NSView())
        findBar.isHidden = true
        textView.selectionFinished = { [weak self] in self?.applyPaintToSelection() }
        textView.cancelPendingAction = { [weak self] in
            guard let self, self.paintAttributes != nil else { return false }
            self.cancelPaintFormat(); return true
        }

        scroll.hasVerticalScroller = true
        scroll.hasHorizontalScroller = true
        scroll.autohidesScrollers = true
        scroll.borderType = .noBorder
        scroll.drawsBackground = true
        scroll.backgroundColor = .underPageBackgroundColor
        scroll.allowsMagnification = true
        scroll.minMagnification = 0.5
        scroll.maxMagnification = 2
        scroll.documentView = canvas
        scroll.contentView.postsBoundsChangedNotifications = true
        canvas.openDocument = { [weak self] url in self?.load(url) }

        textView.isEditable = true
        textView.isSelectable = true
        textView.isRichText = true
        textView.importsGraphics = true
        textView.usesFindPanel = false
        textView.font = NSFont(name: "Arial", size: 12)!
        textView.usesFontPanel = true
        textView.allowsUndo = true
        textView.isContinuousSpellCheckingEnabled = true
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
        textView.setAccessibilityLabel(L("Contenuto del documento", "Document content"))
        canvas.addSubview(textView)
        for (field, label) in [(headerPreview, L("Intestazione, sola lettura", "Header, read only")), (footerPreview, L("Piè di pagina, sola lettura", "Footer, read only"))] {
            field.font = .systemFont(ofSize: 10); field.textColor = .secondaryLabelColor
            field.appearance = NSAppearance(named: .aqua); field.drawsBackground = true; field.backgroundColor = .white
            field.lineBreakMode = .byTruncatingTail; field.maximumNumberOfLines = 1
            field.setAccessibilityLabel(label); field.isHidden = true; canvas.addSubview(field)
        }
        textView.isHidden = true

        status.font = .systemFont(ofSize: 11)
        status.textColor = .secondaryLabelColor
        searchStatus.font = .systemFont(ofSize: 11)
        searchStatus.textColor = .secondaryLabelColor
        saveProgress.style = .spinning; saveProgress.controlSize = .small
        saveProgress.isDisplayedWhenStopped = false
        saveProgress.widthAnchor.constraint(equalToConstant: 16).isActive = true
        saveProgress.heightAnchor.constraint(equalToConstant: 16).isActive = true
        saveProgress.setAccessibilityLabel(L("Salvataggio in corso", "Saving document"))
        let footer = NSStackView(views: [saveProgress, status, NSView(), searchStatus])
        footer.orientation = .horizontal
        let separator = NSBox()
        separator.boxType = .separator
        findInset = scroll.topAnchor.constraint(equalTo: separator.bottomAnchor)
        for view in [bar, separator, findBar, scroll, footer] {
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
            findBar.topAnchor.constraint(equalTo: separator.bottomAnchor, constant: 2),
            findBar.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 14),
            findBar.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -14),
            findBar.heightAnchor.constraint(equalToConstant: 30),
            findInset,
            scroll.leadingAnchor.constraint(equalTo: content.leadingAnchor),
            scroll.trailingAnchor.constraint(equalTo: content.trailingAnchor),
            scroll.bottomAnchor.constraint(equalTo: footer.topAnchor, constant: -5),
            footer.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 14),
            footer.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -14),
            footer.bottomAnchor.constraint(equalTo: content.bottomAnchor, constant: -6),
            footer.heightAnchor.constraint(equalToConstant: 18)
        ])
        NotificationCenter.default.addObserver(self, selector: #selector(viewportChanged),
            name: NSView.boundsDidChangeNotification, object: scroll.contentView)
        NotificationCenter.default.addObserver(self, selector: #selector(viewportChanged),
            name: NSScrollView.didEndLiveMagnifyNotification, object: scroll)
        let saved = UserDefaults.standard.double(forKey: "documentZoom")
        scroll.magnification = saved >= 0.5 && saved <= 2 ? saved : 1.25
        viewportChanged()
    }

    func windowWillResize(_ sender: NSWindow, to frameSize: NSSize) -> NSSize {
        adaptToolbar(width: frameSize.width)
        return frameSize
    }
    func windowDidResize(_ notification: Notification) {
        adaptToolbar(width: window.contentView?.bounds.width ?? 900)
        layoutPaper()
    }
    @objc func viewportChanged() {
        if zoomPopup.numberOfItems > 0 { zoomPopup.item(at: 0)?.title = "\(Int((scroll.magnification * 100).rounded()))%"; zoomPopup.selectItem(at: 0) }
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
        headerPreview.frame = NSRect(x: textView.frame.minX + margin, y: 31, width: pageWidth - margin * 2, height: 15)
        footerPreview.frame = NSRect(x: textView.frame.minX + margin, y: textView.frame.maxY - 22, width: pageWidth - margin * 2, height: 15)
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
        resolveUnsavedChanges { [weak self] proceed in if proceed { self?.loadAfterConfirmation(url) } }
    }
    private func loadAfterConfirmation(_ url: URL) {
        textView.isEditable = false
        let request = UUID()
        loadID = request
        status.stringValue = L("Apertura di \(url.lastPathComponent)…", "Opening \(url.lastPathComponent)…")
        loader.async { [weak self] in
            let result = Result { try WordDocument(url: url) }
            DispatchQueue.main.async {
                guard let self, self.loadID == request else { return }
                self.textView.isEditable = true
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

    func display(_ document: WordDocument, url: URL?) {
        documentGeneration = UUID(); closeApproved = false
        paintAttributes = nil
        self.document = document
        currentPackage = document.package
        headerPreview.stringValue = document.package?.headerText ?? ""; headerPreview.toolTip = document.package?.headerDetails
        footerPreview.stringValue = document.package?.footerText ?? ""; footerPreview.toolTip = document.package?.footerDetails
        headerPreview.isHidden = headerPreview.stringValue.isEmpty; footerPreview.isHidden = footerPreview.stringValue.isEmpty
        currentURL = url
        needsRewriteWarning = !(document.package?.unsupported.isEmpty ?? true)
        margin = document.margin
        pageWidth = document.pageSize.width
        pageHeight = document.pageSize.height
        textView.textContainerInset = NSSize(width: document.margin, height: document.margin)
        textView.textContainer?.containerSize = NSSize(width: pageWidth - document.margin * 2, height: .greatestFiniteMagnitude)
        let body = NSMutableAttributedString(attributedString: document.text)
        body.removeAttribute(docxPackageKey, range: NSRange(location: 0, length: body.length))
        textView.textStorage?.setAttributedString(body)
        if let manager = textView.layoutManager, let container = textView.textContainer {
            manager.ensureLayout(for: container)
            paperHeight = max(pageHeight, ceil(manager.usedRect(for: container).maxY) + document.margin * 2)
        }
        cleanText = NSAttributedString(attributedString: textView.attributedString())
        textView.undoManager?.removeAllActions()
        textView.isHidden = false
        textView.setSelectedRange(NSRange(location: 0, length: 0))
        search.stringValue = ""
        updateMatches()
        window.title = url?.lastPathComponent ?? L("Senza titolo", "Untitled")
        window.representedURL = url
        if let url, FileManager.default.fileExists(atPath: url.path) { NSDocumentController.shared.noteNewRecentDocumentURL(url) }
        refreshRecentMenu()
        updateStatus()
        layoutPaper()
        scroll.contentView.scroll(to: NSPoint(x: max(0, (canvas.bounds.width - scroll.contentView.bounds.width) / 2), y: 0))
        scroll.reflectScrolledClipView(scroll.contentView)
        window.makeFirstResponder(textView)
        cleanText = NSAttributedString(attributedString: textView.attributedString())
        updateStatus()
    }

    private func updateStatus() {
        let counts = documentCounts()
        let state = isSaving ? L("Salvataggio…", "Saving…") : (isDirty ? L("Modifiche non salvate", "Unsaved changes") : L("Salvato", "Saved"))
        let words = counts.words == 1 ? L("1 parola", "1 word") : L("\(counts.words) parole", "\(counts.words) words")
        status.stringValue = words + " · " + state
        window.isDocumentEdited = isDirty
    }

    @objc func focusSearch() {
        if search.superview == nil || search.isHidden { findRequested = true; adaptToolbar(width: window.contentView!.bounds.width) }
        window.makeFirstResponder(search)
    }
    @objc func closeSearch() {
        findRequested = false; search.stringValue = ""; updateMatches()
        window.makeFirstResponder(textView)
        adaptToolbar(width: window.contentView!.bounds.width)
    }
    func control(_ control: NSControl, textView: NSTextView, doCommandBy commandSelector: Selector) -> Bool {
        if control === search && commandSelector == #selector(NSResponder.cancelOperation(_:)) { closeSearch(); return true }
        return false
    }
    @objc func openHelp() { NSWorkspace.shared.open(helpURL) }
    func controlTextDidEndEditing(_ notification: Notification) {
        if (notification.object as? NSTextField) === sizeField { sizeChosen() }
    }
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
        scroll.setMagnification(min(2, max(0.5, value)), centeredAt: NSPoint(x: viewport.midX, y: viewport.midY))
        viewportChanged()
    }

    @objc func printDocument() { runPrint(to: nil) }
    @objc func exportPDF() {
        guard document != nil else { return }
        let panel = NSSavePanel()
        panel.allowedContentTypes = [.pdf]
        panel.nameFieldStringValue = (currentURL?.deletingPathExtension().lastPathComponent ?? L("Senza titolo", "Untitled")) + ".pdf"
        panel.directoryURL = currentURL?.deletingLastPathComponent()
        panel.beginSheetModal(for: window) { [weak self] result in
            if result == .OK, let url = panel.url { self?.runPrint(to: url) }
        }
    }
    func validateMenuItem(_ item: NSMenuItem) -> Bool {
        if item.action == #selector(saveDocument) || item.action == #selector(saveAs) { return !isSaving }
        if item.action == #selector(undoEdit) { return textView.undoManager?.canUndo ?? false }
        if item.action == #selector(redoEdit) { return textView.undoManager?.canRedo ?? false }
        let toggles: [Selector: String] = [#selector(bold): "bold", #selector(italicText): "italic", #selector(underline): "underline", #selector(alignLeft): "left", #selector(alignCenter): "center", #selector(alignRight): "right", #selector(alignJustify): "justify", #selector(bulletList): "bullet", #selector(numberedList): "number", #selector(spellToggle): "spell"]
        if let action = item.action, let key = toggles[action] { item.state = formatButtons[key]?.state ?? .off }
        if item.action == #selector(paintFormat) { item.state = paintAttributes == nil ? .off : .on }
        let alignments: [Selector: NSTextAlignment] = [#selector(alignLeft): .left, #selector(alignCenter): .center, #selector(alignRight): .right, #selector(alignJustify): .justified]
        if let action = item.action, let alignment = alignments[action] {
            let current = (selectionAttributes[.paragraphStyle] as? NSParagraphStyle)?.alignment ?? .natural
            item.state = current == alignment || (current == .natural && alignment == .left) ? .on : .off
        }
        if item.action == #selector(spacingChosen(_:)) {
            let current = (selectionAttributes[.paragraphStyle] as? NSParagraphStyle)?.lineHeightMultiple ?? 0
            item.state = abs(max(1, current) * 100 - CGFloat(item.tag)) < 0.1 ? .on : .off
        }
        if item.action == #selector(styleMenuChosen(_:)) { item.state = item.tag == stylePopup.indexOfSelectedItem ? .on : .off }
        if item.action == #selector(exportPDF) || item.action == #selector(printDocument) { return document != nil }
        return true
    }
    /// Prints, or writes a PDF when `pdf` is set, from a separate text view at the document's page size.
    func runPrint(to pdf: URL?) {
        guard document != nil else { return }
        // Print a separate text view so screen zoom and search highlights never reach the output.
        let printable = NSTextView(frame: NSRect(origin: .zero, size: NSSize(width: pageWidth, height: pageHeight)))
        printable.textStorage?.setAttributedString(textView.attributedString())
        printable.appearance = NSAppearance(named: .aqua)
        printable.usesAdaptiveColorMappingForDarkAppearance = false
        printable.textContainerInset = .zero
        let info = NSPrintInfo.shared.copy() as! NSPrintInfo
        info.paperSize = NSSize(width: pageWidth, height: pageHeight)
        info.leftMargin = margin
        info.rightMargin = margin
        info.topMargin = margin
        info.bottomMargin = margin
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

private struct ExportPicture {
    let range: NSRange
    let marker: String
    let data: Data
    let size: NSSize
    let suffix: String
    let contentType: String
}
private struct ExportAnnotation {
    let range: NSRange
    let start: String
    let end: String
    let link: URL?
    let color: NSColor?
    let multiple: CGFloat
}
private func officeImageType(_ data: Data) -> (String, String)? {
    if data.starts(with: [137, 80, 78, 71, 13, 10, 26, 10]) { return ("png", "image/png") }
    if data.starts(with: [255, 216]) { return ("jpg", "image/jpeg") }
    if data.starts(with: Array("GIF8".utf8)) { return ("gif", "image/gif") }
    if data.starts(with: [73, 73, 42, 0]) || data.starts(with: [77, 77, 0, 42]) { return ("tiff", "image/tiff") }
    if data.starts(with: [66, 77]) { return ("bmp", "image/bmp") }
    return nil
}
private func imagePayload(_ attachment: NSTextAttachment, cache: inout [ObjectIdentifier: (Data, String, String)]) -> (Data, String, String)? {
    // Office-compatible encoded bytes never need decoding or re-encoding.
    if let data = attachment.fileWrapper?.regularFileContents, let type = officeImageType(data) { return (data, type.0, type.1) }
    let key = ObjectIdentifier(attachment)
    if let cached = cache[key] { return cached }
    let image = (attachment.attachmentCell as? NSTextAttachmentCell)?.image ?? attachment.fileWrapper?.regularFileContents.flatMap(NSImage.init(data:))
    guard let image, let tiff = image.tiffRepresentation, let bitmap = NSBitmapImageRep(data: tiff), let png = bitmap.representation(using: .png, properties: [:]) else { return nil }
    let value = (png, "png", "image/png"); cache[key] = value; return value
}

// A single forward pass builds the native writer input. Attribute runs, image
// substitutions and boundary events are consumed in source UTF-16 coordinates.
private func prepareExport(_ text: NSAttributedString, paragraphs: [ExportParagraph], pictures: [ExportPicture], annotations: [ExportAnnotation], retained: RetainedAnchors) -> NSAttributedString {
    var events: [Int: [String]] = [:]
    for annotation in annotations {
        events[annotation.range.location, default: []].append(annotation.start)
        events[NSMaxRange(annotation.range), default: []].append(annotation.end)
    }
    for paragraph in paragraphs { events[paragraph.range.location, default: []].append(paragraph.marker) }
    for anchor in retained.kept { events[anchor.position, default: []].append(anchor.marker) }
    var runs: [(NSRange, [NSAttributedString.Key: Any])] = []
    var styles: [ObjectIdentifier: NSParagraphStyle] = [:]
    var boundaries = Set(events.keys)
    boundaries.insert(0); boundaries.insert(text.length)
    text.enumerateAttributes(in: NSRange(location: 0, length: text.length)) { values, range, _ in
        var attrs = values
        for key in [NSAttributedString.Key.attachment, docxAnchorKey, docxPackageKey, .toolTip] { attrs.removeValue(forKey: key) }
        if let font = attrs[.font] as? NSFont { attrs[.font] = documentFont(font) }
        if let paragraph = attrs[.paragraphStyle] as? NSParagraphStyle {
            let id = ObjectIdentifier(paragraph)
            if styles[id] == nil {
                let style = paragraph.mutableCopy() as! NSMutableParagraphStyle
                style.textBlocks = []; style.textLists = []; styles[id] = style.copy() as? NSParagraphStyle
            }
            attrs[.paragraphStyle] = styles[id]
        }
        runs.append((range, attrs)); boundaries.insert(range.location); boundaries.insert(NSMaxRange(range))
    }
    for picture in pictures { boundaries.insert(picture.range.location); boundaries.insert(NSMaxRange(picture.range)) }
    for position in retained.positions { boundaries.insert(position); boundaries.insert(position + 1) }
    let positions = boundaries.sorted(), source = text.string as NSString
    let output = NSMutableAttributedString(string: "")
    let markerFont = NSFont(name: "Arial", size: 1)!
    var run = 0, picture = 0
    output.beginEditing()
    for (index, position) in positions.enumerated() {
        while run + 1 < runs.count && NSMaxRange(runs[run].0) <= position { run += 1 }
        let attributes = runs.isEmpty ? [:] : runs[run].1
        if let markers = events[position] {
            var markerAttributes = attributes
            markerAttributes[.font] = markerFont
            markerAttributes[.foregroundColor] = NSColor.black
            markerAttributes.removeValue(forKey: .link); markerAttributes.removeValue(forKey: .backgroundColor)
            output.append(NSAttributedString(string: markers.joined(), attributes: markerAttributes))
        }
        while picture < pictures.count && NSMaxRange(pictures[picture].range) <= position { picture += 1 }
        if picture < pictures.count && pictures[picture].range.location == position {
            output.append(NSAttributedString(string: pictures[picture].marker, attributes: attributes))
        }
        guard index + 1 < positions.count else { continue }
        if retained.positions.contains(position) || (picture < pictures.count && NSLocationInRange(position, pictures[picture].range)) { continue }
        let range = NSRange(location: position, length: positions[index + 1] - position)
        if range.length > 0 { output.append(NSAttributedString(string: source.substring(with: range), attributes: attributes)) }
    }
    output.endEditing()
    return output
}

private func replaceXMLChildren(_ parent: XMLElement, with children: [XMLNode]) {
    for child in parent.children ?? [] { child.detach() }
    for child in children { if child.parent != nil { child.detach() }; parent.addChild(child) }
}

// Find and split every sentinel in one traversal, never one whole-document XPath
// per sentinel. Run indices also make all disjoint annotation intervals linear.
private final class ExportXMLIndex {
    var anchors: [String: XMLElement] = [:]
    var runs: [XMLElement] = []
    var runPositions: [ObjectIdentifier: Int] = [:]
    var paragraphPositions: [ObjectIdentifier: Int] = [:]
    init(_ xml: XMLDocument, tokens: Set<String>) throws {
        let pattern = try NSRegularExpression(pattern: "DOCXLITE(?:PARA|IMAGE|FORMAT|KEEP)[A-F0-9]{32}(?:START|END)?")
        var replacements: [ObjectIdentifier: [XMLNode]] = [:]
        var parents: [ObjectIdentifier: (XMLElement, [XMLNode])] = [:]
        for case let original as XMLElement in try xml.nodes(forXPath: "//*[local-name()='r']") {
            let children = original.children ?? []
            var matches: [ObjectIdentifier: [(NSRange, String)]] = [:]
            for child in children where child.localName == "t" {
                let string = child.stringValue ?? ""
                guard string.contains("DOCXLITE") else { continue }
                let source = string as NSString
                let found = pattern.matches(in: string, range: NSRange(location: 0, length: source.length)).compactMap { match -> (NSRange, String)? in
                    let token = source.substring(with: match.range)
                    return tokens.contains(token) ? (match.range, token) : nil
                }
                if !found.isEmpty { matches[ObjectIdentifier(child)] = found }
            }
            guard !matches.isEmpty, let parent = original.parent as? XMLElement else { runs.append(original); continue }
            let parentID = ObjectIdentifier(parent)
            if parents[parentID] == nil { parents[parentID] = (parent, parent.children ?? []) }
            let properties = children.first { $0.localName == "rPr" }
            func newRun() -> XMLElement {
                let run = wordElement("r")
                if let copy = properties?.copy() as? XMLNode { run.addChild(copy) }
                return run
            }
            var current = newRun(), pieces: [XMLNode] = []
            func flush() {
                if current.childCount > (properties == nil ? 0 : 1) { pieces.append(current); runs.append(current) }
                current = newRun()
            }
            func appendText(_ string: String) {
                guard !string.isEmpty else { return }
                let node = wordElement("t", ["xml:space": "preserve"]); node.stringValue = string; current.addChild(node)
            }
            for child in children where child.localName != "rPr" {
                if let found = matches[ObjectIdentifier(child)] {
                    let source = (child.stringValue ?? "") as NSString
                    var offset = 0
                    for (range, token) in found {
                        appendText(source.substring(with: NSRange(location: offset, length: range.location - offset))); flush()
                        let anchor = newRun(), text = wordElement("t"); text.stringValue = token; anchor.addChild(text)
                        guard anchors[token] == nil else { throw structureError() }
                        anchors[token] = anchor; pieces.append(anchor); runs.append(anchor); offset = NSMaxRange(range)
                    }
                    appendText(source.substring(from: offset))
                } else { child.detach(); current.addChild(child) }
            }
            flush(); replacements[ObjectIdentifier(original)] = pieces
        }
        for (_, (parent, children)) in parents {
            replaceXMLChildren(parent, with: children.flatMap { replacements[ObjectIdentifier($0)] ?? [$0] })
        }
        for (index, run) in runs.enumerated() { runPositions[ObjectIdentifier(run)] = index }
        for (index, paragraph) in try xml.nodes(forXPath: "//*[local-name()='p']").enumerated() { paragraphPositions[ObjectIdentifier(paragraph)] = index }
        guard anchors.count == tokens.count else { throw structureError() }
    }
}

// Freeze the mutable rich-text objects as well as the string. No text storage,
// live table/list objects or attachment cells are accessed by the save queue.
func saveSnapshot(_ text: NSAttributedString) -> NSAttributedString {
    let snapshot = NSMutableAttributedString(attributedString: text)
    var tables: [ObjectIdentifier: NSTextTable] = [:], blocks: [ObjectIdentifier: NSTextBlock] = [:]
    var lists: [ObjectIdentifier: NSTextList] = [:], attachments: [ObjectIdentifier: NSTextAttachment] = [:]
    text.enumerateAttribute(.paragraphStyle, in: NSRange(location: 0, length: text.length)) { value, range, _ in
        guard let paragraph = value as? NSParagraphStyle else { return }
        let style = paragraph.mutableCopy() as! NSMutableParagraphStyle
        style.textLists = paragraph.textLists.map { list in
            let id = ObjectIdentifier(list)
            if lists[id] == nil { lists[id] = list.copy() as? NSTextList }
            return lists[id]!
        }
        style.textBlocks = paragraph.textBlocks.map { block in
            let id = ObjectIdentifier(block)
            if let cached = blocks[id] { return cached }
            guard let source = block as? NSTextTableBlock else { return block.copy() as! NSTextBlock }
            let tableID = ObjectIdentifier(source.table)
            if tables[tableID] == nil { tables[tableID] = source.table.copy() as? NSTextTable }
            let copy = NSTextTableBlock(table: tables[tableID]!, startingRow: source.startingRow, rowSpan: source.rowSpan, startingColumn: source.startingColumn, columnSpan: source.columnSpan)
            let dimensions: [NSTextBlock.Dimension] = [.width, .minimumWidth, .maximumWidth, .height, .minimumHeight, .maximumHeight]
            for dimension in dimensions { copy.setValue(source.value(for: dimension), type: source.valueType(for: dimension), for: dimension) }
            let layers: [NSTextBlock.Layer] = [.padding, .border, .margin]
            let edges: [NSRectEdge] = [.minX, .minY, .maxX, .maxY]
            for edge in edges {
                for layer in layers { copy.setWidth(source.width(for: layer, edge: edge), type: source.widthValueType(for: layer, edge: edge), for: layer, edge: edge) }
                copy.setBorderColor(source.borderColor(for: edge), for: edge)
            }
            copy.backgroundColor = source.backgroundColor; copy.verticalAlignment = source.verticalAlignment
            blocks[id] = copy; return copy
        }
        snapshot.addAttribute(.paragraphStyle, value: style.copy() as! NSParagraphStyle, range: range)
    }
    text.enumerateAttribute(.attachment, in: NSRange(location: 0, length: text.length)) { value, range, _ in
        guard let attachment = value as? NSTextAttachment else { return }
        let id = ObjectIdentifier(attachment)
        if attachments[id] == nil {
            let copy = NSTextAttachment()
            let data = attachment.fileWrapper?.regularFileContents
            if let data { copy.fileWrapper = FileWrapper(regularFileWithContents: data) }
            if let image = (attachment.attachmentCell as? NSTextAttachmentCell)?.image {
                // The writer needs only logical bounds for encoded pictures. Do
                // not duplicate decoded raster caches on the main thread.
                let frozenImage = data.flatMap(officeImageType) != nil ? NSImage(size: image.size) : image.copy() as! NSImage
                copy.attachmentCell = NSTextAttachmentCell(imageCell: frozenImage)
            }
            attachments[id] = copy
        }
        snapshot.addAttribute(.attachment, value: attachments[id]!, range: range)
    }
    normalizeDocumentFonts(snapshot)
    return NSAttributedString(attributedString: snapshot)
}

final class DOCXSaveProfile {
    private var last = ProcessInfo.processInfo.systemUptime
    private(set) var stages: [(String, Double)] = []
    func mark(_ label: String) {
        let now = ProcessInfo.processInfo.systemUptime
        stages.append((label, now - last)); last = now
    }
    var appKitSeconds: Double { stages.filter { $0.0 == "AppKit writer" }.reduce(0) { $0 + $1.1 } }
}

// AppKit exports text/styles. DrawingML anchors keep native attachments portable
// even on macOS versions whose Office Open XML writer silently omits pictures.
func writeDOCX(_ text: NSAttributedString, to destination: URL, pageSize: NSSize, margin: CGFloat, profile: DOCXSaveProfile? = nil, package explicitPackage: DOCXPackage? = nil) throws {
    let package = explicitPackage ?? packageInText(text)
    let retained = retainedAnchors(in: text, package: package)
    let paragraphs = exportParagraphs(text).filter { !$0.style.textBlocks.isEmpty || !$0.style.textLists.isEmpty || ($0.heading != nil && $0.heading != "Normal") }
    profile?.mark("paragraphs")
    let temporary = FileManager.default.temporaryDirectory.appendingPathComponent("docxlite-export-" + UUID().uuidString)
    try FileManager.default.createDirectory(at: temporary, withIntermediateDirectories: true)
    defer { try? FileManager.default.removeItem(at: temporary) }
    var pictures: [ExportPicture] = [], imageCache: [ObjectIdentifier: (Data, String, String)] = [:]
    let source = text.string as NSString
    text.enumerateAttribute(.attachment, in: NSRange(location: 0, length: text.length)) { value, range, _ in
        guard let attachment = value as? NSTextAttachment, text.attribute(docxAnchorKey, at: range.location, effectiveRange: nil) == nil, let payload = imagePayload(attachment, cache: &imageCache) else { return }
        let image = (attachment.attachmentCell as? NSTextAttachmentCell)?.image
        for position in range.location..<NSMaxRange(range) where source.character(at: position) == 0xfffc {
            pictures.append(ExportPicture(range: NSRange(location: position, length: 1), marker: "DOCXLITEIMAGE" + UUID().uuidString.replacingOccurrences(of: "-", with: ""), data: payload.0, size: image?.size ?? NSSize(width: 100, height: 100), suffix: payload.1, contentType: payload.2))
        }
    }
    guard pictures.count == text.string.utf16.enumerated().filter({ $0.element == 0xfffc && !retained.positions.contains($0.offset) }).count else {
        throw NSError(domain: appName, code: 4, userInfo: [NSLocalizedDescriptionKey: L("Un'immagine non può essere esportata. Il file non è stato salvato.", "An image could not be exported. The file was not saved.")])
    }
    var annotations: [ExportAnnotation] = []
    text.enumerateAttributes(in: NSRange(location: 0, length: text.length)) { attrs, range, _ in
        let link = (attrs[.link] as? URL) ?? (attrs[.link] as? String).flatMap(URL.init(string:))
        let color = attrs[.backgroundColor] as? NSColor
        let multiple = (attrs[.paragraphStyle] as? NSParagraphStyle)?.lineHeightMultiple ?? 0
        guard link != nil || color != nil || multiple > 0 else { return }
        let token = "DOCXLITEFORMAT" + UUID().uuidString.replacingOccurrences(of: "-", with: "")
        annotations.append(ExportAnnotation(range: range, start: token + "START", end: token + "END", link: link, color: color, multiple: multiple))
    }
    let export = prepareExport(text, paragraphs: paragraphs, pictures: pictures, annotations: annotations, retained: retained)
    profile?.mark("prepare")
    let data = try export.data(from: NSRange(location: 0, length: export.length), documentAttributes: [
        .documentType: NSAttributedString.DocumentType.officeOpenXML, .paperSize: NSValue(size: pageSize),
        .leftMargin: margin, .rightMargin: margin, .topMargin: margin, .bottomMargin: margin
    ])
    profile?.mark("AppKit writer")
    if package == nil && pictures.isEmpty && annotations.isEmpty && paragraphs.isEmpty && retained.kept.isEmpty {
        try data.write(to: destination, options: .atomic); profile?.mark("package write"); return
    }
    let archive = temporary.appendingPathComponent("document.docx")
    try data.write(to: archive)
    let xml = try XMLDocument(data: archiveMember(archive, "word/document.xml"), options: [.nodePreserveWhitespace, .nodeLoadExternalEntitiesNever])
    let relData = try? archiveMember(archive, "word/_rels/document.xml.rels")
    let rels = try relData.map { try XMLDocument(data: $0, options: [.nodeLoadExternalEntitiesNever]) } ?? XMLDocument(xmlString: "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>")
    let content = try XMLDocument(data: archiveMember(archive, "[Content_Types].xml"), options: [.nodeLoadExternalEntitiesNever])
    let media = temporary.appendingPathComponent("word/media")
    try FileManager.default.createDirectory(at: media, withIntermediateDirectories: true)
    try FileManager.default.createDirectory(at: temporary.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)
    var members = ["word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"]
    let tokens = Set(paragraphs.map { $0.marker } + pictures.map { $0.marker } + annotations.flatMap { [$0.start, $0.end] } + retained.kept.map { $0.marker })
    profile?.mark("package read")
    let index = try ExportXMLIndex(xml, tokens: tokens)
    profile?.mark("XML index")
    for (imageIndex, picture) in pictures.enumerated() {
        let id = "docxlite" + UUID().uuidString.replacingOccurrences(of: "-", with: "")
        let filename = "docxlite-\(imageIndex)." + picture.suffix
        try picture.data.write(to: media.appendingPathComponent(filename))
        members.append("word/media/" + filename)
        let relationship = XMLElement(name: "Relationship")
        for (key, value) in ["Id": id, "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "Target": "media/" + filename] {
            relationship.addAttribute(XMLNode.attribute(withName: key, stringValue: value) as! XMLNode)
        }
        rels.rootElement()?.addChild(relationship)
        let cx = Int(max(1, picture.size.width) * 12700), cy = Int(max(1, picture.size.height) * 12700)
        let drawingID = (package?.nextDrawingID ?? 1) + imageIndex
        let drawing = try XMLElement(xmlString: """
        <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="\(cx)" cy="\(cy)"/><wp:docPr id="\(drawingID)" name="Image \(imageIndex + 1)"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="\(imageIndex + 1)" name="\(filename)"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="\(id)"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="\(cx)" cy="\(cy)"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>
        """)
        guard let anchor = index.anchors[picture.marker] else { throw structureError() }
        replaceXMLChildren(anchor, with: (anchor.children ?? []).filter { $0.localName == "rPr" } + [drawing])
        let type = XMLElement(name: "Override")
        type.addAttribute(XMLNode.attribute(withName: "PartName", stringValue: "/word/media/" + filename) as! XMLNode)
        type.addAttribute(XMLNode.attribute(withName: "ContentType", stringValue: picture.contentType) as! XMLNode)
        content.rootElement()?.addChild(type)
    }
    profile?.mark("images")
    var linkParents: [ObjectIdentifier: (XMLElement, [XMLNode])] = [:]
    var runLinks: [ObjectIdentifier: String] = [:], removedAnchors = Set<ObjectIdentifier>()
    for annotation in annotations {
        guard let start = index.anchors[annotation.start], let end = index.anchors[annotation.end],
              let first = index.runPositions[ObjectIdentifier(start)], let last = index.runPositions[ObjectIdentifier(end)], last > first else { throw structureError() }
        let id = "docxlite" + UUID().uuidString.replacingOccurrences(of: "-", with: "")
        if let url = annotation.link {
            let relationship = XMLElement(name: "Relationship")
            setXML(relationship, ["Id": id, "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", "Target": url.absoluteString, "TargetMode": "External"])
            rels.rootElement()?.addChild(relationship)
        }
        for run in index.runs[(first + 1)..<last] {
            if let color = annotation.color?.usingColorSpace(.sRGB) {
                setXML(xmlChild("shd", of: xmlChild("rPr", of: run, first: true)), ["w:fill": String(format: "%02X%02X%02X", Int(color.redComponent * 255), Int(color.greenComponent * 255), Int(color.blueComponent * 255)), "w:val": "clear"])
            }
            if annotation.multiple > 0, let paragraph = ancestor(run, named: "p") {
                setXML(xmlChild("spacing", of: xmlChild("pPr", of: paragraph, first: true)), ["w:line": String(Int((annotation.multiple * 240).rounded())), "w:lineRule": "auto"])
            }
            if annotation.link != nil, let parent = run.parent as? XMLElement {
                let key = ObjectIdentifier(parent)
                if linkParents[key] == nil { linkParents[key] = (parent, parent.children ?? []) }
                runLinks[ObjectIdentifier(run)] = id
            }
        }
        removedAnchors.insert(ObjectIdentifier(start)); removedAnchors.insert(ObjectIdentifier(end))
        start.detach(); end.detach()
    }
    for (_, (parent, children)) in linkParents {
        var output: [XMLNode] = []
        for child in children where !removedAnchors.contains(ObjectIdentifier(child)) {
            if let id = runLinks[ObjectIdentifier(child)] {
                let link = wordElement("hyperlink", ["r:id": id]); child.detach(); link.addChild(child); output.append(link)
            } else { output.append(child) }
        }
        replaceXMLChildren(parent, with: output)
    }
    profile?.mark("annotations")
    try writeStructures(paragraphs, xml: xml, rels: rels, content: content, temporary: temporary, members: &members, width: pageSize.width - margin * 2, index: index, profile: profile)
    if let package { try mergeRetainedPackage(package, xml: xml, rels: rels, content: content, temporary: temporary, archive: archive, members: &members, anchors: retained) }
    try writeRetainedAnchors(retained, package: package, index: index, xml: xml)
    profile?.mark("structures")
    try orderParagraphProperties(xml)
    try xml.xmlData.write(to: temporary.appendingPathComponent("word/document.xml"))
    try rels.xmlData.write(to: temporary.appendingPathComponent("word/_rels/document.xml.rels"))
    try content.xmlData.write(to: temporary.appendingPathComponent("[Content_Types].xml"))
    let zip = Process(); zip.executableURL = URL(fileURLWithPath: "/usr/bin/zip"); zip.currentDirectoryURL = temporary
    zip.arguments = ["-q", archive.path] + members
    zip.standardOutput = FileHandle.nullDevice; zip.standardError = FileHandle.nullDevice
    let finished = try startProcess(zip); finished.wait()
    guard zip.terminationStatus == 0 else { throw NSError(domain: appName, code: 6) }
    try Data(contentsOf: archive).write(to: destination, options: .atomic)
    profile?.mark("package write")
}

private func xmlChild(_ name: String, of parent: XMLElement, first: Bool = false) -> XMLElement {
    if let existing = parent.children?.first(where: { $0.localName == name }) as? XMLElement { return existing }
    let child = XMLElement(name: "w:" + name)
    parent.insertChild(child, at: first ? 0 : parent.childCount)
    return child
}

// Formatting changes use one undo transaction, including paragraph and typing attributes.
extension ReaderApp {
    var defaultAttributes: [NSAttributedString.Key: Any] {
        [.font: NSFont(name: "Arial", size: 12) ?? NSFont.systemFont(ofSize: 12), .foregroundColor: NSColor.black, .paragraphStyle: NSParagraphStyle.default]
    }
    var selectionAttributes: [NSAttributedString.Key: Any] {
        let location = textView.selectedRange().location
        if location < textView.string.utf16.count { return textView.textStorage!.attributes(at: location, effectiveRange: nil) }
        return textView.typingAttributes
    }
    func mutate(_ body: () -> Void) {
        let old = NSAttributedString(attributedString: textView.attributedString())
        let selection = textView.selectedRange()
        let typing = textView.typingAttributes
        body()
        if !old.isEqual(to: textView.attributedString()) {
            textView.undoManager?.registerUndo(withTarget: self) { owner in owner.restore(old, selection: selection, typing: typing) }
        }
        textView.breakUndoCoalescing()
        edited()
        window.makeFirstResponder(textView)
    }
    func restore(_ value: NSAttributedString, selection: NSRange, typing: [NSAttributedString.Key: Any]) {
        let old = NSAttributedString(attributedString: textView.attributedString()), oldSelection = textView.selectedRange(), oldTyping = textView.typingAttributes
        textView.undoManager?.registerUndo(withTarget: self) { owner in owner.restore(old, selection: oldSelection, typing: oldTyping) }
        textView.textStorage?.setAttributedString(value)
        textView.setSelectedRange(selection)
        textView.typingAttributes = typing
        edited()
    }
    func edited() {
        if let manager = textView.layoutManager, let container = textView.textContainer {
            manager.ensureLayout(for: container)
            paperHeight = max(pageHeight, ceil(manager.usedRect(for: container).maxY) + margin * 2)
        }
        updateStatus(); updateMatches(); syncToolbar(); layoutPaper()
    }
    func textDidChange(_ notification: Notification) { paintAttributes = nil; edited() }
    func textViewDidChangeSelection(_ notification: Notification) {
        syncToolbar()
        if !textView.isSelectingWithMouse { applyPaintToSelection() }
    }
    @objc func paintFormat() {
        if paintAttributes != nil { cancelPaintFormat(); return }
        var attributes = defaultAttributes
        for key in paintKeys {
            if let value = selectionAttributes[key] { attributes[key] = value }
            else if key != .font && key != .foregroundColor { attributes.removeValue(forKey: key) }
        }
        attributes[wordParagraphStyleKey] = selectionAttributes[wordParagraphStyleKey] ?? "Normal"
        if let style = selectionAttributes[.paragraphStyle] as? NSParagraphStyle { attributes[.paragraphStyle] = style.copy() }
        paintSourceRange = textView.selectedRange()
        paintAttributes = attributes
        syncToolbar()
        window.makeFirstResponder(textView)
    }
    @objc func cancelPaintFormat() { paintAttributes = nil; syncToolbar() }
    func applyPaintToSelection() {
        let range = textView.selectedRange()
        guard !textView.isSelectingWithMouse, range.length > 0, range != paintSourceRange, let sample = paintAttributes else { return }
        paintAttributes = nil
        mutate {
            let characters = sample.filter { paintKeys.contains($0.key) }
            for key in paintKeys { textView.textStorage?.removeAttribute(key, range: range); textView.typingAttributes.removeValue(forKey: key) }
            textView.textStorage?.addAttributes(characters, range: range)
            textView.typingAttributes.merge(characters) { _, new in new }
            if let source = sample[.paragraphStyle] as? NSParagraphStyle {
                let paragraphs = (textView.string as NSString).paragraphRange(for: range)
                textView.textStorage?.addAttribute(wordParagraphStyleKey, value: sample[wordParagraphStyleKey] ?? "Normal", range: paragraphs)
                textView.typingAttributes[wordParagraphStyleKey] = sample[wordParagraphStyleKey] ?? "Normal"
                textView.textStorage?.enumerateAttribute(.paragraphStyle, in: paragraphs) { value, run, _ in
                    let style = source.mutableCopy() as! NSMutableParagraphStyle
                    // Keep target lists/tables and attachment/link contents intact.
                    let target = value as? NSParagraphStyle ?? .default
                    style.textLists = target.textLists; style.textBlocks = target.textBlocks
                    textView.textStorage?.addAttribute(.paragraphStyle, value: style, range: run)
                }
                textView.typingAttributes[.paragraphStyle] = textView.textStorage?.attribute(.paragraphStyle, at: range.location, effectiveRange: nil)
            }
        }
    }
    @objc func undoEdit() { textView.undoManager?.undo(); edited() }
    @objc func redoEdit() { textView.undoManager?.redo(); edited() }
    func newBlank() {
        loadID = UUID()
        display(WordDocument(text: NSAttributedString(string: "", attributes: defaultAttributes)), url: nil)
        currentURL = nil; needsRewriteWarning = false
        window.title = L("Senza titolo", "Untitled"); window.representedURL = nil
        textView.typingAttributes = defaultAttributes; textView.isEditable = true
        syncToolbar()
    }
    @objc func newDocument() { resolveUnsavedChanges { [weak self] proceed in if proceed { self?.newBlank() } } }
    func resolveUnsavedChanges(_ completion: @escaping (Bool) -> Void) {
        if isSaving {
            saveWaiters.append { [weak self] success in
                guard success, let self else { completion(false); return }
                self.resolveUnsavedChanges(completion)
            }
            return
        }
        guard isDirty else { completion(true); return }
        let alert = NSAlert()
        alert.messageText = L("Vuoi salvare le modifiche?", "Save your changes?")
        alert.informativeText = window.title
        alert.addButton(withTitle: L("Salva", "Save"))
        alert.addButton(withTitle: L("Non salvare", "Don't Save"))
        alert.addButton(withTitle: L("Annulla", "Cancel"))
        alert.beginSheetModal(for: window) { [weak self] response in
            guard let self else { completion(false); return }
            switch response {
            case .alertFirstButtonReturn:
                self.save(interactiveAs: false) { success in
                    if success { self.resolveUnsavedChanges(completion) } else { completion(false) }
                }
            case .alertSecondButtonReturn: completion(true)
            default: completion(false)
            }
        }
    }
    func windowShouldClose(_ sender: NSWindow) -> Bool {
        if closeApproved || (!isSaving && !isDirty) { closeApproved = true; return true }
        resolveUnsavedChanges { [weak self] proceed in
            if proceed { self?.closeApproved = true; sender.performClose(nil) }
        }
        return false
    }
    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply {
        if closeApproved || (!isSaving && !isDirty) { return .terminateNow }
        resolveUnsavedChanges { proceed in sender.reply(toApplicationShouldTerminate: proceed) }
        return .terminateLater
    }
    @objc func saveDocument() { save(interactiveAs: false) }
    @objc func saveAs() { save(interactiveAs: true) }
    var rewriteWarningText: String {
        let issues = currentPackage?.unsupported ?? []
        return L("Questi elementi potrebbero non essere conservati durante il salvataggio:", "These elements may not be preserved when saving:") + "\n" + issues.map { "• " + $0 }.joined(separator: "\n")
    }
    func save(interactiveAs: Bool, completion: ((Bool) -> Void)? = nil) {
        guard !isSaving else { completion?(false); return }
        var choose = interactiveAs || currentURL?.pathExtension.lowercased() != "docx"
        if needsRewriteWarning {
            let alert = NSAlert()
            alert.messageText = L("Salvare una copia del documento Word?", "Save a copy of this Word document?")
            alert.informativeText = rewriteWarningText
            alert.addButton(withTitle: L("Salva una copia…", "Save a Copy…"))
            alert.addButton(withTitle: L("Sovrascrivi", "Overwrite"))
            alert.addButton(withTitle: L("Annulla", "Cancel"))
            switch alert.runModal() {
            case .alertFirstButtonReturn: choose = true
            case .alertSecondButtonReturn: choose = false
            default: completion?(false); return
            }
        }
        var destination = currentURL
        if choose {
            let panel = NSSavePanel()
            panel.allowedContentTypes = [UTType(filenameExtension: "docx")!]
            let base = currentURL?.deletingPathExtension().lastPathComponent ?? L("Senza titolo", "Untitled")
            panel.nameFieldStringValue = base + (needsRewriteWarning ? L(" copia", " copy") : "") + ".docx"
            panel.directoryURL = currentURL?.deletingLastPathComponent()
            guard panel.runModal() == .OK, let url = panel.url else { completion?(false); return }
            destination = url
        }
        guard let url = destination else { completion?(false); return }
        writeDocument(to: url) { [weak self] result in
            switch result {
            case .success: completion?(true)
            case .failure(let error):
                guard let self else { completion?(false); return }
                let alert = NSAlert(error: error)
                alert.messageText = L("Impossibile salvare il documento", "Could not save the document")
                alert.beginSheetModal(for: self.window) { _ in completion?(false) }
            }
        }
    }
    private func finishSave(_ snapshot: NSAttributedString, to url: URL) {
        currentURL = url; needsRewriteWarning = false
        // The disk contains the snapshot, not edits typed while saving.
        cleanText = snapshot
        window.title = url.lastPathComponent; window.representedURL = url
        NSDocumentController.shared.noteNewRecentDocumentURL(url); refreshRecentMenu(); updateStatus()
    }
    func writeDocument(to url: URL, completion: @escaping (Result<Void, Error>) -> Void) {
        guard !isSaving else {
            completion(.failure(NSError(domain: appName, code: 10, userInfo: [NSLocalizedDescriptionKey: L("Un salvataggio è già in corso.", "A save is already in progress.")]))); return
        }
        textView.breakUndoCoalescing()
        let snapshot = saveSnapshot(textView.attributedString()), generation = documentGeneration
        let size = NSSize(width: pageWidth, height: pageHeight), inset = margin, package = currentPackage
        isSaving = true; saveProgress.startAnimation(nil); updateStatus()
        saver.async { [weak self] in
            let onMain = Thread.isMainThread
            let result: Result<Void, Error> = autoreleasepool { Result { try writeDOCX(snapshot, to: url, pageSize: size, margin: inset, package: package) } }
            DispatchQueue.main.async {
                guard let self else { return }
                self.lastSaveWasOnMainThread = onMain
                self.isSaving = false; self.saveProgress.stopAnimation(nil)
                if case .success = result, self.documentGeneration == generation { self.finishSave(snapshot, to: url) }
                self.updateStatus()
                let waiters = self.saveWaiters; self.saveWaiters.removeAll()
                completion(result)
                let success: Bool
                if case .success = result { success = true } else { success = false }
                for waiter in waiters { waiter(success) }
            }
        }
    }
    #if DEBUG
    // Synchronous model seam for round-trip tests; UI actions always use the queue.
    func writeDocument(to url: URL) throws {
        let snapshot = saveSnapshot(textView.attributedString())
        try writeDOCX(snapshot, to: url, pageSize: NSSize(width: pageWidth, height: pageHeight), margin: margin, package: currentPackage)
        finishSave(snapshot, to: url)
    }
    #endif
    func setAttributes(_ values: [NSAttributedString.Key: Any], remove: [NSAttributedString.Key] = []) {
        mutate {
            let range = textView.selectedRange()
            if range.length > 0 {
                for key in remove { textView.textStorage?.removeAttribute(key, range: range) }
                textView.textStorage?.addAttributes(values, range: range)
            }
            var typing = textView.typingAttributes
            for key in remove { typing.removeValue(forKey: key) }
            typing.merge(values) { _, new in new }
            textView.typingAttributes = typing
        }
    }
    func fontTrait(_ trait: NSFontTraitMask) {
        let font = documentFont(selectionAttributes[.font] as? NSFont ?? NSFont.systemFont(ofSize: 12))
        let remove = NSFontManager.shared.traits(of: font).contains(trait)
        mutate {
            let range = textView.selectedRange()
            if range.length > 0 {
                textView.textStorage?.enumerateAttribute(.font, in: range) { value, run, _ in
                    let f = documentFont(value as? NSFont ?? font)
                    let converted = remove ? NSFontManager.shared.convert(f, toNotHaveTrait: trait) : NSFontManager.shared.convert(f, toHaveTrait: trait)
                    textView.textStorage?.addAttribute(.font, value: converted, range: run)
                }
            }
            textView.typingAttributes[.font] = remove ? NSFontManager.shared.convert(font, toNotHaveTrait: trait) : NSFontManager.shared.convert(font, toHaveTrait: trait)
        }
    }
    @objc func bold() { fontTrait(.boldFontMask) }
    @objc func italicText() { fontTrait(.italicFontMask) }
    @objc func underline() { setAttributes([.underlineStyle: ((selectionAttributes[.underlineStyle] as? Int ?? 0) == 0 ? 1 : 0)]) }
    func paragraphChange(_ change: (NSMutableParagraphStyle) -> Void) {
        mutate {
            let range = (textView.string as NSString).paragraphRange(for: textView.selectedRange())
            if range.length > 0 {
                textView.textStorage?.enumerateAttribute(.paragraphStyle, in: range) { value, run, _ in
                    let style = (value as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
                    change(style); textView.textStorage?.addAttribute(.paragraphStyle, value: style, range: run)
                }
            }
            let style = (textView.typingAttributes[.paragraphStyle] as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
            change(style); textView.typingAttributes[.paragraphStyle] = style
        }
    }
    func applyStyle(_ index: Int) {
        let selection = textView.selectedRange()
        let range = (textView.string as NSString).paragraphRange(for: selection)
        textView.setSelectedRange(range)
        // Use the same font sizes for recognition after a native DOCX round trip.
        mutate {
            let base = NSFont(name: "Arial", size: styleSizes[index]) ?? NSFont.systemFont(ofSize: styleSizes[index])
            let font = index >= 2 ? NSFontManager.shared.convert(base, toHaveTrait: .boldFontMask) : base
            textView.textStorage?.addAttribute(.font, value: font, range: range)
            textView.textStorage?.addAttribute(wordParagraphStyleKey, value: wordStyleIDs[index], range: range)
            textView.typingAttributes[wordParagraphStyleKey] = wordStyleIDs[index]
            textView.typingAttributes[.font] = font
            textView.textStorage?.enumerateAttribute(.paragraphStyle, in: range) { value, run, _ in
                let style = (value as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
                style.paragraphSpacing = index == 0 ? 6 : 12
                style.paragraphSpacingBefore = index == 0 ? 0 : 10
                textView.textStorage?.addAttribute(.paragraphStyle, value: style, range: run)
            }
            let style = (textView.typingAttributes[.paragraphStyle] as? NSParagraphStyle ?? .default).mutableCopy() as! NSMutableParagraphStyle
            style.paragraphSpacing = index == 0 ? 6 : 12; style.paragraphSpacingBefore = index == 0 ? 0 : 10
            textView.typingAttributes[.paragraphStyle] = style
        }
        textView.setSelectedRange(selection)
    }
    @objc func styleChosen(_ sender: NSPopUpButton) { applyStyle(sender.indexOfSelectedItem) }
    @objc func fontChosen(_ sender: NSPopUpButton) {
        if sender.indexOfSelectedItem == 6 { NSFontManager.shared.orderFrontFontPanel(self); return }
        let font = documentFont(selectionAttributes[.font] as? NSFont ?? .systemFont(ofSize: 12))
        if let replacement = NSFont(name: (sender.titleOfSelectedItem?.hasPrefix(".") == false ? sender.titleOfSelectedItem : nil) ?? "Arial", size: font.pointSize) { setAttributes([.font: NSFontManager.shared.convert(replacement, toHaveTrait: NSFontManager.shared.traits(of: font))]) }
    }
    func setFontSize(_ size: CGFloat) {
        let font = documentFont(selectionAttributes[.font] as? NSFont ?? .systemFont(ofSize: 12))
        setAttributes([.font: NSFontManager.shared.convert(font, toSize: min(200, max(6, size)))])
    }
    @objc func sizeChosen() { setFontSize(CGFloat(sizeField.doubleValue)) }
    @objc func sizeDown() { setFontSize(CGFloat(sizeField.doubleValue) - 1) }
    @objc func sizeUp() { setFontSize(CGFloat(sizeField.doubleValue) + 1) }
    @objc func alignLeft() { paragraphChange { $0.alignment = .left } }
    @objc func alignCenter() { paragraphChange { $0.alignment = .center } }
    @objc func alignRight() { paragraphChange { $0.alignment = .right } }
    @objc func alignJustify() { paragraphChange { $0.alignment = .justified } }
    @objc func indentLess() { paragraphChange { $0.headIndent = max(0, $0.headIndent - 24); $0.firstLineHeadIndent = max(0, $0.firstLineHeadIndent - 24) } }
    @objc func indentMore() { paragraphChange { $0.headIndent += 24; $0.firstLineHeadIndent += 24 } }
    @objc func clearFormat() {
        mutate {
            let range = textView.selectedRange()
            // Preserve image attachments while resetting character and paragraph formatting.
            textView.textStorage?.enumerateAttributes(in: range) { attributes, run, _ in
                var clean = defaultAttributes
                if let attachment = attributes[.attachment] { clean[.attachment] = attachment }
                for key in [docxAnchorKey, docxPackageKey, .toolTip, .baselineOffset] { if let value = attributes[key] { clean[key] = value } }
                if let old = attributes[.paragraphStyle] as? NSParagraphStyle {
                    let style = NSMutableParagraphStyle(); style.textBlocks = old.textBlocks; style.textLists = old.textLists
                    clean[.paragraphStyle] = style
                }
                textView.textStorage?.setAttributes(clean, range: run)
            }
            let paragraphs = (textView.string as NSString).paragraphRange(for: range)
            textView.textStorage?.removeAttribute(wordParagraphStyleKey, range: paragraphs)
            textView.typingAttributes = defaultAttributes
        }
    }
    @objc func spellToggle() { textView.isContinuousSpellCheckingEnabled.toggle(); syncToolbar() }
    @objc func spacingChosen(_ sender: NSMenuItem) { paragraphChange { $0.lineHeightMultiple = CGFloat(sender.tag) / 100 } }
    @objc func bulletList() { toggleList(numbered: false) }
    @objc func numberedList() { toggleList(numbered: true) }
    func toggleList(numbered: Bool) {
        let list = NSTextList(markerFormat: numbered ? NSTextList.MarkerFormat(rawValue: "{decimal}.") : .disc, options: 0)
        let current = (selectionAttributes[.paragraphStyle] as? NSParagraphStyle)?.textLists.last
        let remove = current?.markerFormat == list.markerFormat
        changeLinePrefixes { line, number in
            let stripped = stripListPrefix(line)
            return remove ? stripped : "\t" + list.marker(forItemNumber: number) + "\t" + stripped
        }
        paragraphChange {
            $0.textLists = remove ? [] : [list]
            $0.headIndent = remove ? 0 : 30; $0.firstLineHeadIndent = 0
            $0.tabStops = [NSTextTab(textAlignment: .left, location: 15), NSTextTab(textAlignment: .left, location: 30)]
        }
    }
    func changeLinePrefixes(_ transform: (String, Int) -> String) {
        mutate {
            let range = (textView.string as NSString).paragraphRange(for: textView.selectedRange())
            let source = textView.textStorage!.attributedSubstring(from: range)
            let replacement = NSMutableAttributedString(attributedString: source)
            var lines: [(NSRange, String)] = []
            (source.string as NSString).enumerateSubstrings(in: NSRange(location: 0, length: source.length), options: [.byParagraphs, .substringNotRequired]) { _, line, _, _ in
                lines.append((line, (source.string as NSString).substring(with: line)))
            }
            for (index, pair) in lines.enumerated().reversed() {
                let changed = transform(pair.1, index + 1)
                // Replace only the prefix so styled content keeps its attributes.
                let oldPrefix = pair.1.utf16.count - stripListPrefix(pair.1).utf16.count
                let newPrefix = changed.utf16.count - stripListPrefix(pair.1).utf16.count
                replacement.replaceCharacters(in: NSRange(location: pair.0.location, length: oldPrefix), with: String(changed.prefix(max(0, newPrefix))))
            }
            if range.length == 0 { replacement.append(NSAttributedString(string: transform("", 1), attributes: textView.typingAttributes)) }
            textView.textStorage?.replaceCharacters(in: range, with: replacement)
            textView.setSelectedRange(NSRange(location: range.location, length: replacement.length))
        }
    }
    @objc func checklist() {
        changeLinePrefixes { line, _ in
            if line.hasPrefix("☐ ") { return "☑ " + String(line.dropFirst(2)) }
            if line.hasPrefix("☑ ") { return "☐ " + String(line.dropFirst(2)) }
            return "☐ " + stripListPrefix(line)
        }
    }
    func applyLink(_ value: String) {
        var range = textView.selectedRange()
        if range.length == 0, range.location < textView.string.utf16.count {
            var effective = NSRange()
            if textView.textStorage?.attribute(.link, at: range.location, effectiveRange: &effective) != nil { range = effective; textView.setSelectedRange(range) }
        }
        guard range.length > 0 else { return }
        if value.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty { setAttributes([:], remove: [.link]); return }
        let address = value.contains(":") ? value : "https://" + value
        if let url = URL(string: address) { setAttributes([.link: url, .foregroundColor: NSColor.systemBlue, .underlineStyle: 1]) }
    }
    @objc func linkDialog() {
        let alert = NSAlert()
        alert.messageText = L("Inserisci o modifica link", "Insert or edit link")
        alert.informativeText = L("Seleziona il testo. Lascia vuoto per rimuovere il link.", "Select text. Leave empty to remove the link.")
        let field = NSTextField(frame: NSRect(x: 0, y: 0, width: 330, height: 24))
        field.stringValue = (selectionAttributes[.link] as? URL)?.absoluteString ?? (selectionAttributes[.link] as? String ?? "")
        alert.accessoryView = field; alert.addButton(withTitle: L("Applica", "Apply")); alert.addButton(withTitle: L("Annulla", "Cancel"))
        if alert.runModal() == .alertFirstButtonReturn { applyLink(field.stringValue) }
    }
    func insertImage(_ data: Data) {
        guard let image = NSImage(data: data) else { return }
        let scale = min(1, (pageWidth - margin * 2) / max(1, image.size.width))
        image.size = NSSize(width: image.size.width * scale, height: image.size.height * scale)
        let attachment = NSTextAttachment()
        attachment.fileWrapper = FileWrapper(regularFileWithContents: data)
        attachment.attachmentCell = NSTextAttachmentCell(imageCell: image)
        let replacement = NSMutableAttributedString(attachment: attachment)
        for key in [NSAttributedString.Key.paragraphStyle, wordParagraphStyleKey] {
            if let value = selectionAttributes[key] { replacement.addAttribute(key, value: value, range: NSRange(location: 0, length: 1)) }
        }
        mutate {
            let range = textView.selectedRange()
            textView.textStorage?.replaceCharacters(in: range, with: replacement)
            textView.setSelectedRange(NSRange(location: range.location + 1, length: 0))
        }
    }
    @objc func imageDialog() {
        let panel = NSOpenPanel(); panel.allowedContentTypes = [.image]
        if panel.runModal() == .OK, let url = panel.url, let data = try? Data(contentsOf: url) { insertImage(data) }
    }
    @objc func colorChosen(_ sender: NSMenuItem) {
        let highlight = sender.representedObject as? String == "highlight"
        let colors: [NSColor] = [.black, .systemRed, .systemBlue, .systemGreen, .systemYellow, .white]
        if sender.tag < 0 { setAttributes([:], remove: [.backgroundColor]) }
        else { setAttributes([highlight ? .backgroundColor : .foregroundColor: colors[sender.tag]]) }
    }
    func colorMenu(highlight: Bool) -> NSMenu {
        let menu = NSMenu()
        let names = [L("Nero", "Black"), L("Rosso", "Red"), L("Blu", "Blue"), L("Verde", "Green"), L("Giallo", "Yellow"), L("Bianco", "White")]
        if highlight { let item = menu.addItem(withTitle: L("Nessuno", "None"), action: #selector(colorChosen(_:)), keyEquivalent: ""); item.tag = -1; item.target = self; item.representedObject = "highlight" }
        for (index, name) in names.enumerated() {
            let item = menu.addItem(withTitle: name, action: #selector(colorChosen(_:)), keyEquivalent: "")
            item.target = self; item.tag = index; item.representedObject = highlight ? "highlight" : "text"
        }
        return menu
    }
    @objc func zoomChosen(_ sender: NSPopUpButton) {
        if sender.indexOfSelectedItem == 8 { fitWidth() }
        else if sender.indexOfSelectedItem > 0 { applyZoom(CGFloat([50, 75, 100, 115, 125, 150, 200][sender.indexOfSelectedItem - 1]) / 100) }
    }
    func spacingMenu() -> NSMenu {
        let menu = NSMenu()
        for value in [100, 115, 150, 200] {
            let item = menu.addItem(withTitle: NumberFormatter.localizedString(from: NSNumber(value: Double(value) / 100), number: .decimal), action: #selector(spacingChosen(_:)), keyEquivalent: "")
            item.tag = value; item.target = self
        }
        return menu
    }
    func populateFormatMenu(_ menu: NSMenu) {
        func add(_ title: String, _ action: Selector, _ key: String = "", _ mods: NSEvent.ModifierFlags = .command) {
            let item = menu.addItem(withTitle: title, action: action, keyEquivalent: key); item.target = self; item.keyEquivalentModifierMask = mods
        }
        add(L("Copia formato", "Paint Format"), #selector(paintFormat))
        menu.addItem(.separator())
        add(L("Grassetto", "Bold"), #selector(bold), "b")
        add(L("Corsivo", "Italic"), #selector(italicText), "i")
        add(L("Sottolineato", "Underline"), #selector(underline), "u")
        for (index, name) in styleNames.enumerated() { let item = menu.addItem(withTitle: name, action: #selector(styleMenuChosen(_:)), keyEquivalent: ""); item.tag = index; item.target = self }
        add(L("Carattere…", "Font…"), #selector(showFonts))
        add(L("Riduci dimensione carattere", "Decrease Font Size"), #selector(sizeDown))
        add(L("Aumenta dimensione carattere", "Increase Font Size"), #selector(sizeUp))
        let color = menu.addItem(withTitle: L("Colore testo", "Text Color"), action: nil, keyEquivalent: ""); color.submenu = colorMenu(highlight: false)
        let highlight = menu.addItem(withTitle: L("Colore evidenziazione", "Highlight Color"), action: nil, keyEquivalent: ""); highlight.submenu = colorMenu(highlight: true)
        add(L("Allinea a sinistra", "Align Left"), #selector(alignLeft), "l", [.command, .shift])
        add(L("Allinea al centro", "Align Center"), #selector(alignCenter), "e", [.command, .shift])
        add(L("Allinea a destra", "Align Right"), #selector(alignRight), "r", [.command, .shift])
        add(L("Giustifica", "Justify"), #selector(alignJustify), "j", [.command, .shift])
        let spacing = menu.addItem(withTitle: L("Interlinea", "Line Spacing"), action: nil, keyEquivalent: ""); spacing.submenu = spacingMenu()
        add(L("Elenco puntato", "Bulleted List"), #selector(bulletList), "8", [.command, .shift])
        add(L("Elenco numerato", "Numbered List"), #selector(numberedList), "7", [.command, .shift])
        add(L("Riduci rientro", "Decrease Indent"), #selector(indentLess), "[")
        add(L("Aumenta rientro", "Increase Indent"), #selector(indentMore), "]")
        add(L("Cancella formattazione", "Clear Formatting"), #selector(clearFormat), "\\")
    }
    @objc func showFonts() { NSFontManager.shared.orderFrontFontPanel(self) }
    @objc func styleMenuChosen(_ sender: NSMenuItem) { applyStyle(sender.tag) }
    func textColorImage(_ color: NSColor) -> NSImage {
        let symbol = NSImage(systemSymbolName: "a", accessibilityDescription: nil)?.withSymbolConfiguration(.init(pointSize: 13, weight: .medium))
        let image = NSImage(size: NSSize(width: 18, height: 20), flipped: false) { _ in
            symbol?.draw(in: NSRect(x: 2, y: 5, width: 14, height: 14))
            color.setFill(); NSRect(x: 1, y: 1, width: 16, height: 3).fill()
            return true
        }
        image.isTemplate = false
        return image
    }
    func alignmentMenu() -> NSMenu {
        let menu = NSMenu()
        let choices: [(String, String, Selector, String)] = [
            (L("Allinea a sinistra", "Align Left"), "text.alignleft", #selector(alignLeft), "l"),
            (L("Allinea al centro", "Align Center"), "text.aligncenter", #selector(alignCenter), "e"),
            (L("Allinea a destra", "Align Right"), "text.alignright", #selector(alignRight), "r"),
            (L("Giustifica", "Justify"), "text.justify", #selector(alignJustify), "j")
        ]
        for (title, symbol, action, key) in choices {
            let item = menu.addItem(withTitle: title, action: action, keyEquivalent: key)
            item.target = self; item.keyEquivalentModifierMask = [.command, .shift]
            item.image = NSImage(systemSymbolName: symbol, accessibilityDescription: nil)
        }
        return menu
    }
    func buildToolbar() -> NSStackView {
        var views: [NSView] = [], group: [NSView] = []
        func append(_ view: NSView, _ id: String) {
            view.identifier = NSUserInterfaceItemIdentifier(id)
            toolbarItems[id] = view; views.append(view); group.append(view)
        }
        func button(_ id: String, _ symbol: String, _ tip: String, _ action: Selector) {
            let button = NSButton(image: NSImage(systemSymbolName: symbol, accessibilityDescription: tip) ?? NSImage(), target: self, action: action)
            button.toolTip = tip; button.bezelStyle = .recessed; button.controlSize = .small
            button.showsBorderOnlyWhileMouseInside = true; button.contentTintColor = .labelColor
            button.widthAnchor.constraint(equalToConstant: 26).isActive = true
            button.setButtonType(["bold", "italic", "underline", "spell", "paint", "check", "bullet", "number"].contains(id) ? .pushOnPushOff : .momentaryPushIn)
            formatButtons[id] = button; append(button, id)
        }
        func popup(_ popup: NSPopUpButton, _ id: String, _ titles: [String], _ width: CGFloat, _ tip: String, _ action: Selector) {
            popup.addItems(withTitles: titles); popup.target = self; popup.action = action
            popup.toolTip = tip; popup.controlSize = .small; popup.bezelStyle = .recessed
            popup.widthAnchor.constraint(equalToConstant: width).isActive = true
            append(popup, id)
        }
        @discardableResult func menuButton(_ id: String, _ symbol: String, _ tip: String, _ menu: NSMenu) -> NSPopUpButton {
            let popup = NSPopUpButton(frame: .zero, pullsDown: true)
            menu.insertItem(withTitle: "", action: nil, keyEquivalent: "", at: 0)
            menu.items[0].image = NSImage(systemSymbolName: symbol, accessibilityDescription: tip)
            popup.menu = menu; popup.toolTip = tip; popup.controlSize = .small
            popup.bezelStyle = .recessed; popup.contentTintColor = .labelColor
            popup.widthAnchor.constraint(equalToConstant: 32).isActive = true
            append(popup, id); return popup
        }
        func separator() {
            toolbarGroups.append(group); group = []
            let holder = NSView(), line = NSBox()
            line.boxType = .separator; line.translatesAutoresizingMaskIntoConstraints = false
            holder.addSubview(line)
            NSLayoutConstraint.activate([
                holder.widthAnchor.constraint(equalToConstant: 9), holder.heightAnchor.constraint(equalToConstant: 18),
                line.widthAnchor.constraint(equalToConstant: 1), line.heightAnchor.constraint(equalToConstant: 18),
                line.centerXAnchor.constraint(equalTo: holder.centerXAnchor), line.centerYAnchor.constraint(equalTo: holder.centerYAnchor)
            ])
            views.append(holder); toolbarSeparators.append(holder)
        }
        button("undo", "arrow.uturn.backward", L("Annulla (⌘Z)", "Undo (⌘Z)"), #selector(undoEdit))
        button("redo", "arrow.uturn.forward", L("Ripeti (⇧⌘Z)", "Redo (⇧⌘Z)"), #selector(redoEdit))
        button("print", "printer", L("Stampa (⌘P)", "Print (⌘P)"), #selector(printDocument))
        button("spell", "textformat.abc", L("Controllo ortografico", "Spell check"), #selector(spellToggle))
        button("paint", "paintbrush", L("Copia formato", "Paint format"), #selector(paintFormat))
        separator()
        popup(zoomPopup, "zoom", ["100%", "50%", "75%", "100%", "115%", "125%", "150%", "200%", L("Adatta alla finestra", "Fit Width")], 68, L("Zoom", "Zoom"), #selector(zoomChosen(_:)))
        separator()
        popup(stylePopup, "style", styleNames, 160, L("Stile paragrafo", "Paragraph style"), #selector(styleChosen(_:)))
        separator()
        popup(fontPopup, "font", ["Arial", "Helvetica", "Times New Roman", "Georgia", "Courier New", "Verdana", L("Altro…", "Other…")], 100, L("Carattere", "Font"), #selector(fontChosen(_:)))
        separator()
        button("sizeDown", "minus", L("Riduci dimensione carattere", "Decrease font size"), #selector(sizeDown))
        sizeField.delegate = self; sizeField.alignment = .center; sizeField.font = .systemFont(ofSize: 11)
        sizeField.widthAnchor.constraint(equalToConstant: 32).isActive = true
        sizeField.toolTip = L("Dimensione carattere", "Font size"); append(sizeField, "size")
        button("sizeUp", "plus", L("Aumenta dimensione carattere", "Increase font size"), #selector(sizeUp))
        separator()
        button("bold", "bold", L("Grassetto (⌘B)", "Bold (⌘B)"), #selector(bold))
        button("italic", "italic", L("Corsivo (⌘I)", "Italic (⌘I)"), #selector(italicText))
        button("underline", "underline", L("Sottolineato (⌘U)", "Underline (⌘U)"), #selector(underline))
        textColorPopup = menuButton("textColor", "a", L("Colore testo", "Text color"), colorMenu(highlight: false))
        menuButton("highlight", "highlighter", L("Colore evidenziazione", "Highlight color"), colorMenu(highlight: true))
        separator()
        button("link", "link", L("Link (⌘K)", "Link (⌘K)"), #selector(linkDialog))
        button("image", "photo", L("Inserisci immagine", "Insert image"), #selector(imageDialog))
        separator()
        alignmentPopup = menuButton("alignment", "text.alignleft", L("Allineamento", "Alignment"), alignmentMenu())
        menuButton("spacing", "arrow.up.and.down.text.horizontal", L("Interlinea", "Line spacing"), spacingMenu())
        button("check", "checklist", L("Elenco di controllo", "Checklist"), #selector(checklist))
        button("bullet", "list.bullet", L("Elenco puntato (⇧⌘8)", "Bulleted list (⇧⌘8)"), #selector(bulletList))
        button("number", "list.number", L("Elenco numerato (⇧⌘7)", "Numbered list (⇧⌘7)"), #selector(numberedList))
        button("outdent", "decrease.indent", L("Riduci rientro (⌘[)", "Decrease indent (⌘[)"), #selector(indentLess))
        button("indent", "increase.indent", L("Aumenta rientro (⌘])", "Increase indent (⌘])"), #selector(indentMore))
        button("clear", "textformat", L("Cancella formattazione (⌘\\)", "Clear formatting (⌘\\)"), #selector(clearFormat))
        toolbarGroups.append(group); group = []
        let more = NSMenu(); populateFormatMenu(more); more.addItem(.separator())
        for (title, action) in [(L("Link…", "Link…"), #selector(linkDialog)), (L("Immagine…", "Image…"), #selector(imageDialog)), (L("Elenco di controllo", "Checklist"), #selector(checklist)), (L("Stampa…", "Print…"), #selector(printDocument)), (L("Controllo ortografico", "Spell check"), #selector(spellToggle)), (L("Trova…", "Find…"), #selector(focusSearch))] {
            more.addItem(withTitle: title, action: action, keyEquivalent: "").target = self
        }
        menuButton("more", "ellipsis", L("Altre opzioni", "More options"), more)
        let spacer = NSView(); append(spacer, "spacer")
        search.placeholderString = L("Cerca", "Search"); search.delegate = self; search.target = self; search.action = #selector(findNext)
        search.sendsSearchStringImmediately = true
        searchWidth = search.widthAnchor.constraint(equalToConstant: 120); searchWidth.isActive = true
        search.toolTip = L("Cerca nel documento (⌘F)", "Find (⌘F)"); append(search, "search")
        button("zoomOut", "minus.magnifyingglass", L("Riduci (⌘−)", "Zoom out (⌘−)"), #selector(zoomOut))
        button("zoomIn", "plus.magnifyingglass", L("Ingrandisci (⌘+)", "Zoom in (⌘+)"), #selector(zoomIn))
        button("fit", "arrow.left.and.right", L("Adatta alla finestra (⌘9)", "Fit width (⌘9)"), #selector(fitWidth))
        let bar = NSStackView(views: views); bar.orientation = .horizontal; bar.spacing = 3; bar.detachesHiddenViews = true
        bar.wantsLayer = true; bar.layer?.cornerRadius = 8
        bar.layer?.backgroundColor = NSColor.controlBackgroundColor.cgColor
        let held = ["undo", "redo", "paint", "zoom", "style", "font", "sizeDown", "size", "sizeUp", "bold", "italic", "underline", "textColor", "highlight", "alignment", "more", "spacer"]
        toolbar = bar; toolbarHeld = held.compactMap { toolbarItems[$0] }
        toolbarOptional = ["print", "spell", "link", "image", "spacing", "check", "bullet", "number", "outdent", "indent", "clear"].compactMap { toolbarItems[$0] }
        toolbarRight = ["search", "zoomOut", "zoomIn", "fit"].compactMap { toolbarItems[$0] }
        bar.setClippingResistancePriority(.defaultLow, for: .horizontal)
        adaptToolbar(width: window.contentView?.bounds.width ?? 960)
        return bar
    }
    func adaptToolbar(width: CGFloat) {
        guard let toolbar else { return }
        func cost(_ view: NSView) -> CGFloat {
            if view === search { return 123 }
            if let fixed = view.constraints.first(where: { $0.firstAttribute == .width && $0.secondItem == nil }) { return fixed.constant + 3 }
            return max(0, view.intrinsicContentSize.width) + 3
        }
        // Reserve separators while choosing controls; this never underestimates space.
        var used = toolbarHeld.reduce(CGFloat(toolbarSeparators.count * 12)) { $0 + cost($1) }
        var shown = Set(toolbarHeld.map(ObjectIdentifier.init))
        for view in toolbarOptional where used + cost(view) <= width - 28 {
            shown.insert(ObjectIdentifier(view)); used += cost(view)
        }
        let rightFits = used + toolbarRight.reduce(CGFloat(0)) { $0 + cost($1) } <= width - 28
        let wasSearching = search.currentEditor() != nil
        if !rightFits && wasSearching { findRequested = true }
        for view in toolbarHeld + toolbarOptional { toolbar.setVisibilityPriority(shown.contains(ObjectIdentifier(view)) ? .mustHold : .notVisible, for: view) }
        for view in toolbarRight where view !== search { toolbar.setVisibilityPriority(rightFits ? .mustHold : .notVisible, for: view) }
        // A separator appears only between groups that still contain controls.
        var hasPrevious = false
        for (index, group) in toolbarGroups.enumerated() {
            let visible = group.contains { shown.contains(ObjectIdentifier($0)) }
            if index > 0 { toolbar.setVisibilityPriority(visible && hasPrevious ? .mustHold : .notVisible, for: toolbarSeparators[index - 1]) }
            if visible { hasPrevious = true }
        }
        if findRequested {
            if toolbar.arrangedSubviews.contains(search) { toolbar.removeArrangedSubview(search); search.removeFromSuperview() }
            if !findBar.arrangedSubviews.contains(search) { findBar.insertArrangedSubview(search, at: 0) }
            search.isHidden = false; searchWidth.constant = 220
        } else {
            if findBar.arrangedSubviews.contains(search) { findBar.removeArrangedSubview(search); search.removeFromSuperview() }
            if !toolbar.arrangedSubviews.contains(search) {
                let position = toolbar.arrangedSubviews.firstIndex { $0 === formatButtons["zoomOut"] } ?? toolbar.arrangedSubviews.count
                toolbar.insertArrangedSubview(search, at: position)
            }
            searchWidth.constant = 120
            toolbar.setVisibilityPriority(rightFits ? .mustHold : .notVisible, for: search)
        }
        findBar.isHidden = !findRequested; findInset?.constant = findRequested ? 34 : 0
        if wasSearching && findRequested { window.makeFirstResponder(search) }
    }
    func syncToolbar() {
        guard window != nil else { return }
        let attrs = selectionAttributes
        let font = documentFont(attrs[.font] as? NSFont ?? .systemFont(ofSize: 12))
        let traits = NSFontManager.shared.traits(of: font)
        let paragraph = attrs[.paragraphStyle] as? NSParagraphStyle ?? .default
        for button in formatButtons.values { button.state = .off }
        formatButtons["bold"]?.state = traits.contains(.boldFontMask) ? .on : .off
        formatButtons["italic"]?.state = traits.contains(.italicFontMask) ? .on : .off
        formatButtons["underline"]?.state = (attrs[.underlineStyle] as? Int ?? 0) != 0 ? .on : .off
        formatButtons["spell"]?.state = textView.isContinuousSpellCheckingEnabled ? .on : .off
        formatButtons["paint"]?.state = paintAttributes == nil ? .off : .on
        formatButtons["paint"]?.toolTip = paintAttributes == nil ? L("Copia formato", "Paint format") : L("Seleziona il testo a cui applicare il formato (Esc per annullare)", "Select text to apply formatting (Esc to cancel)")
        textColorPopup?.item(at: 0)?.image = textColorImage(attrs[.foregroundColor] as? NSColor ?? .black)
        let alignment: String = paragraph.alignment == .center ? "center" : paragraph.alignment == .right ? "right" : paragraph.alignment == .justified ? "justify" : "left"
        alignmentPopup?.item(at: 0)?.image = NSImage(systemSymbolName: alignment == "justify" ? "text.justify" : "text.align" + alignment, accessibilityDescription: nil)
        if let list = paragraph.textLists.last { formatButtons[list.markerFormat.rawValue.contains("decimal") ? "number" : "bullet"]?.state = .on }
        let line = (textView.string as NSString).substring(with: (textView.string as NSString).paragraphRange(for: NSRange(location: textView.selectedRange().location, length: 0)))
        formatButtons["check"]?.state = line.hasPrefix("☐ ") || line.hasPrefix("☑ ") ? .on : .off
        formatButtons["undo"]?.isEnabled = textView.undoManager?.canUndo ?? false
        formatButtons["redo"]?.isEnabled = textView.undoManager?.canRedo ?? false
        sizeField.stringValue = String(format: "%g", Double(font.pointSize))
        let family = font.familyName ?? "Arial"
        if fontPopup.item(withTitle: family) == nil {
            if fontPopup.numberOfItems > 7 { fontPopup.removeItem(at: 7) }
            fontPopup.addItem(withTitle: family)
        }
        fontPopup.selectItem(withTitle: family)
        stylePopup.selectItem(at: (attrs[wordParagraphStyleKey] as? String).flatMap { wordStyleIDs.firstIndex(of: $0) } ?? styleSizes.firstIndex(where: { abs($0 - font.pointSize) < 0.2 }) ?? 0)
        zoomPopup.item(at: 0)?.title = "\(Int((scroll.magnification * 100).rounded()))%"
        zoomPopup.selectItem(at: 0)
    }
    func documentCounts() -> (words: Int, characters: Int) {
        let plain = documentPlainText(textView.attributedString())
        var words = 0
        plain.enumerateSubstrings(in: plain.startIndex..<plain.endIndex, options: .byWords) { _, _, _, _ in words += 1 }
        return (words, plain.count)
    }
    @objc func countDialog() {
        let counts = documentCounts(), alert = NSAlert()
        alert.messageText = L("Conteggio parole", "Word Count")
        alert.informativeText = L("\(counts.words) parole\n\(counts.characters) caratteri (spazi inclusi)", "\(counts.words) words\n\(counts.characters) characters (including spaces)")
        alert.runModal()
    }
    @discardableResult func replaceAll(_ needle: String, with replacement: String) -> Int {
        guard !needle.isEmpty else { return 0 }
        search.stringValue = needle; updateMatches(); let ranges = matches
        mutate { for range in ranges.reversed() { textView.textStorage?.replaceCharacters(in: range, with: replacement) } }
        return ranges.count
    }
    @objc func replaceDialog() {
        let alert = NSAlert(); alert.messageText = L("Trova e sostituisci", "Find and Replace")
        let find = NSTextField(string: search.stringValue), replace = NSTextField(string: "")
        find.placeholderString = L("Trova", "Find"); replace.placeholderString = L("Sostituisci con", "Replace with")
        let stack = NSStackView(views: [find, replace]); stack.orientation = .vertical; stack.frame = NSRect(x: 0, y: 0, width: 330, height: 60)
        alert.accessoryView = stack; alert.addButton(withTitle: L("Sostituisci tutto", "Replace All")); alert.addButton(withTitle: L("Annulla", "Cancel"))
        if alert.runModal() == .alertFirstButtonReturn { _ = replaceAll(find.stringValue, with: replace.stringValue) }
    }
}

func stripListPrefix(_ line: String) -> String {
    if line.hasPrefix("☐ ") || line.hasPrefix("☑ ") { return String(line.dropFirst(2)) }
    if line.hasPrefix("\t"), let end = line.dropFirst().firstIndex(of: "\t") { return String(line[line.index(after: end)...]) }
    return line
}

#if DEBUG
if CommandLine.arguments.contains("--performance-test") || CommandLine.arguments.contains("--performance-baseline") {
    runSavePerformanceTest(enforceLimit: !CommandLine.arguments.contains("--performance-baseline"))
    exit(0)
}
#endif

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
        app.finishLaunching()
        delegate.applicationDidFinishLaunching(Notification(name: NSApplication.didFinishLaunchingNotification))
        scheduleReaderTests(delegate, fixture: URL(fileURLWithPath: CommandLine.arguments[2]))
    }
    #endif
    app.run()
}
