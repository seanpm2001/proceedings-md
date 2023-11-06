import * as path from 'path';
import * as fs from 'fs';
import * as JSZip from 'jszip';
import * as pandoc from "./pandoc";
import * as XML from "./xml";
import * as OXML from "./oxml";
import {DocumentMeta} from "./pandoc";
import Relationships from "./wrappers/relationships";
import ContentTypes from "./wrappers/content-types";
import Styles, {DocDefaults, LatentStyles, Style} from "./wrappers/styles";

const pandocFlags = ["--tab-stop=8"]

const properDocXmlns = new Map<string, string>([
    ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
    ["xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["xmlns:o", "urn:schemas-microsoft-com:office:office"],
    ["xmlns:v", "urn:schemas-microsoft-com:vml"],
    ["xmlns:w10", "urn:schemas-microsoft-com:office:word"],
    ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
    ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
    ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
])

export const languages = ["ru", "en"]

function visitStyleCrossReferences(style: Style, callback: (node: XML.Node) => void) {
    let basedOnTag = style.node.getChild("w:basedOn")
    if (basedOnTag) callback(basedOnTag)

    let linkTag = style.node.getChild("w:link")
    if (linkTag) callback(linkTag)

    let nextTag = style.node.getChild("w:next")
    if (nextTag) callback(nextTag)
}

function getStyleCrossReferences(styles: Styles): XML.Node[] {
    let result = []
    for(let style of styles.styles.values()) {
        result.push(style.node.shallowCopy())
        visitStyleCrossReferences(style, (node) => result.push(node))
    }
    return result
}

function getDocStyleUseReferences(doc: XML.Node, result: XML.Node[] = []): XML.Node[] {

    doc.visitSubtree((node: XML.Node) => {
        return node.getTagName() === "w:pStyle" || node.getTagName() == "w:rStyle"
    }, (node: XML.Node) => {
        result.push(node.shallowCopy())
    })

    return result
}

function extractTemplateStyles(styles: Styles): Style[] {
    let result = []
    for(let style of styles.styles.values()) {
        if (style.getId().startsWith("template-")) {
            result.push(new Style().readXml(style.node.deepCopy()))
        }
    }
    return result
}

function patchStyleDefinitions(doc: XML.Node, styles: Styles, map: Map<string, string>) {
    for(let style of styles.styles.values()) {
        if(map.has(style.getId())) {
            styles.removeStyle(style)
            style.setId(map.get(style.getId()))
            styles.addStyle(style)
        }
    }
}

function patchStyleUseReferences(doc: XML.Node, styles: Styles, map: Map<string, string>) {
    let docReferences = getDocStyleUseReferences(doc)
    let crossReferences = getStyleCrossReferences(styles)

    for (let ref of docReferences.concat(crossReferences)) {
        if (ref.getAttr("w:val") && map.has(ref.getAttr("w:val"))) {
            ref.setAttr("w:val", map.get(ref.getAttr("w:val")))
        }
    }
}

function getUsedStyles(doc: XML.Node): Set<string> {
    let references = getDocStyleUseReferences(doc)
    let set = new Set<string>()

    for (let ref of references) {
        set.add(ref.getAttr("w:val"))
    }

    return set
}

function populateStyles(alreadyMet: Set<string>, styles: Styles) {
    for (let styleId of alreadyMet) {
        let style = styles.styles.get(styleId)

        if (!style) {
            throw new Error("Style id " + styleId + " not found")
        }

        if(style.getBaseStyle() !== null) alreadyMet.add(style.getBaseStyle())
        if(style.getLinkedStyle() !== null) alreadyMet.add(style.getLinkedStyle())
        if(style.getNextStyle() !== null) alreadyMet.add(style.getNextStyle())
    }
}

function getUsedStylesDeep(doc: XML.Node, styles: Styles, requiredStyles: string[] = []): Set<string> {
    let usedStyles = getUsedStyles(doc)

    for (let requiredStyle of requiredStyles) {
        usedStyles.add(requiredStyle)
    }

    do {
        let size = usedStyles.size
        populateStyles(usedStyles, styles)
        if (usedStyles.size == size) break;
    } while (true);

    return usedStyles
}

function getStyleIdsByName(styles: Iterable<Style>): Map<string, string> {
    let table = new Map<string, any>()

    for (let style of styles) {
        table.set(style.getName(), style.getId())
    }

    return table
}

function addCollisionPatch(mappingTable: Map<string, string>, styleId: string): string {
    let newId = "template-" + mappingTable.size.toString()
    mappingTable.set(styleId, newId)
    return newId
}

function getMappingTable(usedStyles: Set<string>): Map<string, string> {
    let mappingTable = new Map<string, string>
    for (let style of usedStyles) {
        addCollisionPatch(mappingTable, style);
    }

    return mappingTable
}

function appendStyles(target: XML.Node, defs: XML.Node[]) {
    let styles = target.getChild("w:styles")
    for (let def of defs) {
        styles.pushChild(def)
    }
}

interface ListStyles {
    BulletList: NumIdPatchEntry
    OrderedList: NumIdPatchEntry

    [key: string]: NumIdPatchEntry | undefined
}

interface NumIdPatchEntry {
    styleName: string
    numId: string
}

function applyListStyles(doc: XML.Node, styles: ListStyles): Map<string, string> {

    let stack = []
    let currentState = undefined

    let newStyles = new Map<string, string>()
    let lastId = 10000

    doc.visitSubtree((node) => {
        let tagName = node.getTagName()

        if (tagName === "w:pPr" && currentState) {
            // Remove any old pStyle and add our own
            node.removeChildren("w:pStyle")
            node.pushChild(XML.Node.build("w:pStyle").setAttr("w:val", styles[currentState.listStyle].styleName))
        }

        if (tagName === "w:numId" && currentState) {
            node.setAttr("w:val", String(currentState.numId))
        }

        if (tagName === XML.keys.comment) {
            let commentValue = node.getComment()
            // Switch between ordered list and bullet list
            // if comment is detected

            if (commentValue.indexOf("ListMode OrderedList") != -1) {
                stack.push(currentState)
                currentState = {
                    numId: lastId++,
                    listStyle: "OrderedList"
                }
                newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId)
            }

            if (commentValue.indexOf("ListMode BulletList") != -1) {
                stack.push(currentState)
                currentState = {
                    numId: lastId++,
                    listStyle: "BulletList"
                }
                newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId)
            }

            if (commentValue.indexOf("ListMode None") != -1) {
                currentState = stack[stack.length - 1]
                stack.pop()
            }
        }

        return true
    })

    return newStyles
}

function removeCollidedStyles(styles: Styles, collisions: Set<string>) {
    let ignored = 0

    for(let style of styles.styles.values()) {
        if(collisions.has(style.getId())) {
            styles.removeStyle(style)
        }
    }
}

function copyLatentStyles(source: LatentStyles, target: LatentStyles) {
    source.node.assign(target.node)
}

function copyDocDefaults(source: DocDefaults, target: DocDefaults) {
    source.node.assign(target.node)
}

async function copyFile(source, target, path) {
    target.file(path, await source.file(path).async("arraybuffer"))
}

function addNewNumberings(targetNumberingParsed: XML.Node, newListStyles: Map<string, string>) {
    let numberingTag = targetNumberingParsed.getChild("w:numbering")

    // <w:num w:numId="newNum">
    //   <w:abstractNumId w:val="oldNum"/>
    // </w:num>

    for (let [newNum, oldNum] of newListStyles) {

        let overrides = []
        for (let i = 0; i < 9; i++) {
            overrides.push(
                XML.Node.build("w:lvlOverride")
                    .setAttr("w:ilvl", String(i))
                    .appendChildren([
                        XML.Node.build("w:startOverride").setAttr("w:val", "1")
                    ])
            )
        }

        numberingTag.pushChild(
            XML.Node.build("w:num")
                .setAttr("w:numId", newNum)
                .appendChildren([
                    XML.Node.build("w:abstractNumId").setAttr("w:val", oldNum),
                    ...overrides
                ]));
    }
}

function getParagraphText(paragraph: XML.Node): string {
    let result = ""

    paragraph.visitSubtree("w:t", (node) => {
        result += getRawText(node)
    })

    return result
}

function getRawText(tag: XML.Node): string {
    let result = ""

    tag.visitSubtree(XML.keys.text, (node) => {
        result += node.getText()
    })

    return result
}

function replaceInlineTemplate(node: XML.Node, template: string, value: string) {
    if (value === "@none") {
        let i = findParagraphWithPattern(node, template, 0);
        for (; i !== null; i = findParagraphWithPattern(node, template, i)) {
            node.removeChild([i])
            i = i - 1
        }
    } else {
        replaceStringTemplate(node, template, value)
    }
}

function replaceStringTemplate(tag: XML.Node, template: string, value: string) {
    tag.visitSubtree(XML.keys.text, (node) => {
        node.setText(node.getText().replace(template, value))
    })
}

function findParagraphWithPattern(node: XML.Node, pattern: string, startIndex: number = 0): number | null {
    let found: number | null = null

    node.visitChildren((rel, path) => {
        let text = getParagraphText(rel)
        if (text.indexOf(pattern) !== -1) {
            found = path
            return false
        }
        return true
    }, startIndex)

    return found
}

function findParagraphWithPatternStrict(body: XML.Node, pattern: string, startIndex: number = 0): number | null {
    let paragraphIndex = findParagraphWithPattern(body, pattern, startIndex)
    if (paragraphIndex === null) {
        throw new Error(`The template document should have pattern ${pattern}`)
    }

    let text = getParagraphText(body.getChild([paragraphIndex]))
    if (text != pattern) {
        throw new Error(`The ${pattern} pattern should be the only text of the paragraph`)
    }

    return paragraphIndex
}

function templateReplaceBodyContents(templateBody: XML.Node, body: XML.Node) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{body}}}")

    let children = body.getChildren()

    templateBody.removeChild([paragraphIndex])
    templateBody.insertChildren(children, [paragraphIndex])
}

function clearParagraphContents(paragraph: XML.Node) {
    paragraph.removeChildren("w:r")
}

function templateAuthorList(templateBody: XML.Node, meta: DocumentMeta) {

    let authors = meta.getSection("ispras_templates.authors").asArray()

    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{authors_${language}}}}`)

        let newParagraphs: XML.Node[] = []

        let authorIndex = 1;

        let template = templateBody.getChild([paragraphIndex])

        for (let author of authors) {
            let newParagraph = template.deepCopy()
            clearParagraphContents(newParagraph)

            let name = author.getString("name_" + language)
            let orcid = author.getString("orcid")
            let email = author.getString("email")

            let indexLine = String(authorIndex)
            let authorLine = `${name}, ORCID: ${orcid}, <${email}>`

            newParagraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]))
            newParagraph.pushChild(OXML.buildParagraphTextTag(authorLine))

            newParagraphs.push(newParagraph)

            authorIndex++
        }

        templateBody.removeChild([paragraphIndex])
        templateBody.insertChildren(newParagraphs, [paragraphIndex])
    }

    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{organizations_${language}}}}`)
        let organizations = meta.getSection("ispras_templates.organizations_" + language).asArray()

        let newParagraphs = []
        let orgIndex = 1

        let template = templateBody.getChild([paragraphIndex])

        for (let organization of organizations) {
            let newParagraph = template.deepCopy()
            clearParagraphContents(newParagraph)

            let indexLine = String(orgIndex)

            newParagraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]))
            newParagraph.pushChild(OXML.buildParagraphTextTag(organization.getString()))

            orgIndex++
        }

        templateBody.removeChild([paragraphIndex])
        templateBody.insertChildren(newParagraphs, [paragraphIndex])
    }
}

function templateReplaceLinks(templateBody: XML.Node, meta: DocumentMeta, listRules: any) {
    let litListRule = listRules["LitList"]
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{links}}}")
    let links = meta.getSection("ispras_templates.links").asArray()

    let newParagraphs = []

    for (let link of links) {
        let newParagraph = OXML.buildParagraphWithStyle(litListRule.styleName)
        let style = newParagraph.getChild("w:pPr")
        style.pushChild(OXML.buildNumPr("0", litListRule.numId))

        newParagraph.pushChild(OXML.buildParagraphTextTag(link.getString()))
        newParagraphs.push(newParagraph)
    }

    templateBody.removeChild([paragraphIndex])
    templateBody.insertChildren(newParagraphs, [paragraphIndex])
}

function templateReplaceAuthorsDetail(templateBody: XML.Node, meta: DocumentMeta) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{authors_detail}}}")
    let authors = meta.getSection("ispras_templates.authors").asArray()

    let newParagraphs = []
    let template = templateBody.getChild([paragraphIndex])

    for (let author of authors) {
        for (let language of languages) {
            let newParagraph = template.deepCopy()

            let line = author.getString("details_" + language)

            clearParagraphContents(newParagraph)
            newParagraph.pushChild(OXML.buildParagraphTextTag(line))
            newParagraphs.push(newParagraph)
        }
    }

    templateBody.removeChild([paragraphIndex])
    templateBody.insertChildren(newParagraphs, [paragraphIndex])
}

function replacePageHeaders(headers: XML.Node[], meta: DocumentMeta): any {
    let header_ru = meta.getString("ispras_templates.page_header_ru")
    let header_en = meta.getString("ispras_templates.page_header_en")

    if (header_ru === "@use_citation") {
        header_ru = meta.getString("ispras_templates.for_citation_ru")
    }

    if (header_en === "@use_citation") {
        header_en = meta.getString("ispras_templates.for_citation_en")
    }

    for (let header of headers) {
        replaceInlineTemplate(header, `{{{page_header_ru}}}`, header_ru)
        replaceInlineTemplate(header, `{{{page_header_en}}}`, header_en)
    }
}

function replaceTemplates(template: XML.Node, body: XML.Node, meta: DocumentMeta): XML.Node {
    let templateCopy = template.deepCopy()
    let templateBody = OXML.getDocumentBody(templateCopy)

    templateReplaceBodyContents(templateBody, body)
    templateAuthorList(templateBody, meta)

    let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"]

    for (let template of templates) {
        for (let language of languages) {
            let template_lang = template + "_" + language
            let value = meta.getString("ispras_templates." + template_lang)
            replaceInlineTemplate(templateBody, `{{{${template_lang}}}}`, value)
        }
    }

    templateReplaceAuthorsDetail(templateBody, meta)

    return templateCopy
}

function setXmlns(xml: XML.Node, xmlns: Map<string, string>) {
    const document = xml.getChild("w:document")
    for (let [key, value] of xmlns) {
        document.setAttr(key, value)
    }
}

function patchRelIds(doc: XML.Node, map: Map<string, string>) {
    doc.visitSubtree((node) => {
        for (let attr of ["r:id", "r:embed"]) {
            let relId = node.getAttr(attr)
            if (relId && map.has(relId)) {
                node.setAttr(attr, map.get(relId))
            }
        }

        return true
    })
}

async function fixDocxStyles(sourcePath: string, targetPath: string, meta: DocumentMeta): Promise<void> {
    let resourcesDir = path.dirname(process.argv[1]) + "/../resources"

    // Load the source and target documents
    let target = await JSZip.loadAsync(fs.readFileSync(sourcePath))
    let source = await JSZip.loadAsync(fs.readFileSync(resourcesDir + '/isp-reference.docx'))

    let sourceStylesXML = await source.file("word/styles.xml").async("string");
    let targetStylesXML = await target.file("word/styles.xml").async("string");
    let sourceDocXML = await source.file("word/document.xml").async("string");
    let targetDocXML = await target.file("word/document.xml").async("string");
    let targetContentTypesXML = await target.file("[Content_Types].xml").async("string");
    let targetDocumentRelsXML = await target.file("word/_rels/document.xml.rels").async("string");
    let sourceDocumentRelsXML = await source.file("word/_rels/document.xml.rels").async("string");
    let targetNumberingXML = await source.file("word/numbering.xml").async("string");
    let sourceHeader1 = await source.file("word/header1.xml").async("string");
    let sourceHeader2 = await source.file("word/header2.xml").async("string");
    let sourceHeader3 = await source.file("word/header3.xml").async("string");

    let sourceStyles = new Styles().readXmlString(sourceStylesXML);
    let targetStyles = new Styles().readXmlString(targetStylesXML);

    let sourceDocParsed = XML.Node.fromXmlString(sourceDocXML);
    let targetDocParsed = XML.Node.fromXmlString(targetDocXML);
    let targetNumberingParsed = XML.Node.fromXmlString(targetNumberingXML);
    let sourceHeader1Parsed = XML.Node.fromXmlString(sourceHeader1)
    let sourceHeader2Parsed = XML.Node.fromXmlString(sourceHeader2)
    let sourceHeader3Parsed = XML.Node.fromXmlString(sourceHeader3)

    // Transferring the content types
    let targetContentTypes = new ContentTypes().readXmlString(targetContentTypesXML);
    let sourceContentTypes = new ContentTypes().readXmlString(targetContentTypesXML);
    targetContentTypes.join(sourceContentTypes)
    target.file("[Content_Types].xml", targetContentTypes.toXmlString())

    // Transferring the document relations
    let targetDocumentRels = new Relationships().readXmlString(targetDocumentRelsXML);
    let sourceDocumentRels = new Relationships().readXmlString(sourceDocumentRelsXML);
    let relMap = targetDocumentRels.join(sourceDocumentRels)
    patchRelIds(sourceDocParsed, relMap)
    target.file("word/_rels/document.xml.rels", targetDocumentRels.toXMLString())

    copyLatentStyles(sourceStyles.latentStyles, targetStyles.latentStyles)
    copyDocDefaults(sourceStyles.docDefaults, targetStyles.docDefaults)

    let targetStylesNamesToId = getStyleIdsByName(targetStyles.styles.values());
    let sourceStylesNamesToId = getStyleIdsByName(sourceStyles.styles.values());

    let usedStyles = getUsedStylesDeep(sourceDocParsed, sourceStyles, [
        "ispSubHeader-1 level",
        "ispSubHeader-2 level",
        "ispSubHeader-3 level",
        "ispAuthor",
        "ispAnotation",
        "ispText_main",
        "ispList",
        "ispListing",
        "ispListing Знак",
        "ispLitList",
        "ispPicture_sign",
        "ispNumList",
        "Normal"
    ].map(name => sourceStylesNamesToId.get(name)))
    let mappingTable = getMappingTable(usedStyles)

    patchStyleDefinitions(sourceDocParsed, sourceStyles, mappingTable)
    patchStyleUseReferences(sourceDocParsed, sourceStyles, mappingTable)

    let extractedDefs = extractTemplateStyles(sourceStyles)
    let extractedStyleIdsByName = getStyleIdsByName(extractedDefs)

    let stylePatch = new Map<string, string>([
        ["Heading1", extractedStyleIdsByName.get("ispSubHeader-1 level")],
        ["Heading2", extractedStyleIdsByName.get("ispSubHeader-2 level")],
        ["Heading3", extractedStyleIdsByName.get("ispSubHeader-3 level")],
        ["Author", extractedStyleIdsByName.get("ispAuthor")],
        ["AbstractTitle", extractedStyleIdsByName.get("ispAnotation")],
        ["Abstract", extractedStyleIdsByName.get("ispAnotation")],
        ["BlockText", extractedStyleIdsByName.get("ispText_main")],
        ["BodyText", extractedStyleIdsByName.get("ispText_main")],
        ["FirstParagraph", extractedStyleIdsByName.get("ispText_main")],
        ["Normal", extractedStyleIdsByName.get("Normal")],
        ["SourceCode", extractedStyleIdsByName.get("ispListing")],
        ["VerbatimChar", extractedStyleIdsByName.get("ispListing Знак")],
        ["ImageCaption", extractedStyleIdsByName.get("ispPicture_sign")],
    ])

    let stylesToRemove = new Set<string>([
        "Heading4",
        "Heading5",
        "Heading6",
        "Heading7",
        "Heading8",
        "Heading9",
    ])

    for (let possibleCollision of extractedStyleIdsByName) {
        let sourceStyleName = possibleCollision[0]
        let sourceStyleId = possibleCollision[1]

        if (targetStylesNamesToId.has(sourceStyleName)) {
            let targetStyleId = targetStylesNamesToId.get(sourceStyleName)

            if (!stylePatch.has(targetStyleId)) {
                stylePatch.set(targetStyleId, sourceStyleId)
            }
            stylesToRemove.add(targetStyleId)
        }
    }

    removeCollidedStyles(targetStyles, stylesToRemove)
    for(let style of extractedDefs) {
        targetStyles.addStyle(style)
    }

    patchStyleUseReferences(targetDocParsed, targetStyles, stylePatch)

    let patchRules = {
        "OrderedList": {styleName: extractedStyleIdsByName.get("ispNumList"), numId: "33"},
        "BulletList": {styleName: extractedStyleIdsByName.get("ispList1"), numId: "43"},
        "LitList": {styleName: extractedStyleIdsByName.get("ispLitList"), numId: "80"},
    };

    let newListStyles = applyListStyles(targetDocParsed, patchRules)

    setXmlns(sourceDocParsed, properDocXmlns)

    targetDocParsed = replaceTemplates(sourceDocParsed, OXML.getDocumentBody(targetDocParsed), meta)

    templateReplaceLinks(OXML.getDocumentBody(targetDocParsed), meta, patchRules)

    addNewNumberings(targetNumberingParsed, newListStyles)

    replacePageHeaders([sourceHeader1Parsed, sourceHeader2Parsed, sourceHeader3Parsed], meta)

    await copyFile(source, target, "word/_rels/header1.xml.rels")
    await copyFile(source, target, "word/_rels/header2.xml.rels")
    await copyFile(source, target, "word/_rels/header3.xml.rels")
    await copyFile(source, target, "word/_rels/footer1.xml.rels")
    await copyFile(source, target, "word/_rels/footer2.xml.rels")
    await copyFile(source, target, "word/_rels/footer3.xml.rels")

    await copyFile(source, target, "word/footer1.xml")
    await copyFile(source, target, "word/footer2.xml")
    await copyFile(source, target, "word/footer3.xml")

    await copyFile(source, target, "word/footnotes.xml")
    await copyFile(source, target, "word/theme/theme1.xml")
    await copyFile(source, target, "word/fontTable.xml")
    await copyFile(source, target, "word/settings.xml")
    await copyFile(source, target, "word/webSettings.xml")
    await copyFile(source, target, "word/media/image1.png")

    target.file("word/header1.xml", sourceHeader1Parsed.toXmlString())
    target.file("word/header2.xml", sourceHeader2Parsed.toXmlString())
    target.file("word/header3.xml", sourceHeader3Parsed.toXmlString())

    target.file("word/numbering.xml", targetNumberingParsed.toXmlString())
    target.file("word/styles.xml", targetStyles.toXmlString())
    target.file("word/document.xml", targetDocParsed.toXmlString())

    fs.writeFileSync(targetPath, await target.generateAsync({type: "uint8array"}));
}

function getOpenxmlInjection(xml: string): any {
    return {
        t: "RawBlock",
        c: ["openxml", xml]
    }
}

function fixCompactLists(list): any[] {
    // For compact list, 'para' is replaced with 'plain'.
    // Compact lists were not mentioned in the
    // guidelines, so get rid of them

    for (let i = 0; i < list.c.length; i++) {
        let element = list.c[i]
        if (typeof element[0] === "object" && element[0].t === "Plain") {
            element[0].t = "Para"
        }
        list.c[i] = getPatchedMetaElement(list.c[i])
    }

    return [
        getOpenxmlInjection(`<!-- ListMode ${list.t} -->`),
        list,
        getOpenxmlInjection(`<!-- ListMode None -->`)
    ]
}

function getImageCaption(content): any {
    let paragraph = XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", "ImageCaption"),
            XML.Node.build("w:contextualSpacing").setAttr("w:val", "true"),
        ]),
        OXML.buildParagraphTextTag(pandoc.getMetaString(content))
    ]);

    return getOpenxmlInjection(paragraph.toXmlString());
}

function getListingCaption(content): any {
    let elements = XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", "BodyText"),
            XML.Node.build("w:jc").setAttr("w:val", "left"),
        ]),
        OXML.buildParagraphTextTag(pandoc.getMetaString(content), [
            XML.Node.build("w:i"),
            XML.Node.build("w:iCs"),
            XML.Node.build("w:sz").setAttr("w:val", "18"),
            XML.Node.build("w:szCs").setAttr("w:val", "18"),
        ])
    ])

    return getOpenxmlInjection(elements.toXmlString());
}

function getPatchedMetaElement(element): any {
    if (Array.isArray(element)) {
        let newArray = []

        for (let i = 0; i < element.length; i++) {
            let patched = getPatchedMetaElement(element[i])
            if (Array.isArray(patched) && !Array.isArray(element[i])) {
                newArray.push(...patched)
            } else {
                newArray.push(patched)
            }
        }

        return newArray
    }

    if (typeof element !== "object" || !element) {
        return element
    }

    let type = element.t
    let value = element.c

    if (type === 'Div') {
        let content = value[1];
        let classes = value[0][1];
        if (classes) {
            if (classes.includes("img-caption")) {
                return getImageCaption(content)
            }

            if (classes.includes("table-caption") || classes.includes("listing-caption")) {
                return getListingCaption(content)
            }
        }
    } else if (type === 'BulletList' || type === 'OrderedList') {
        return fixCompactLists(element)
    }

    for (let key of Object.getOwnPropertyNames(element)) {
        // Be safe from prototype pollution
        if (key === "__proto__") continue
        element[key] = getPatchedMetaElement(element[key]);
    }

    return element
}

async function generatePandocDocx(source: string, target: string): Promise<DocumentMeta> {
    let markdown = await fs.promises.readFile(source, "utf-8")

    let meta = await pandoc.pandoc(markdown, ["-f", "markdown", "-t", "json", ...pandocFlags])
    let metaParsed = JSON.parse(meta)

    metaParsed.blocks = getPatchedMetaElement(metaParsed.blocks)

    await pandoc.pandoc(JSON.stringify(metaParsed), ["-f", "json", "-t", "docx", "-o", target])

    return DocumentMeta.fromPandocMeta(metaParsed.meta)
}

async function main(): Promise<void> {
    let argv = process.argv
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>")
        process.exit(1)
    }

    let source = argv[2]
    let target = argv[3]

    let meta = await generatePandocDocx(source, target + ".tmp")
    await fixDocxStyles(target + ".tmp", target, meta).then()
}

main().then()