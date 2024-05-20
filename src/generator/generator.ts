import WordDocument from "src/word/word-document"
import mdast from "mdast"
import {DocumentJsonMeta} from "src/document-json-meta"
import * as XML from "src/xml"
import * as OXML from 'src/word/oxml'
import ParagraphTemplateSubstitution from "src/word-templates/paragraph-template-substitution";
import {languages} from "src/main";
import InlineTemplateSubstitution from "src/word-templates/inline-template-substitution";
import {Num} from "src/word/numbering";
import ImageSize from "image-size";
import fs from "fs"
import path from "path"
import temml from "temml"
import {mml2omml} from "mathml2omml"

export interface NumberingStyle {
    numId: string,
    styleId: string
}

let DecimalNumberingStyle: NumberingStyle = {styleId: "ispNumList", numId: "33"}
let BulletNumberingStyle: NumberingStyle = {styleId: "ispList1", numId: "43"}

export class GenerationContext {
    doc: WordDocument
    meta: DocumentJsonMeta
    nodeStack: XML.Node[] = []
    numberingStack: NumberingStyle[] = []

    visit(element: mdast.Node) {
        visitors[element.type]?.(element, this)
    }

    visitChildren(element: mdast.Parent) {
        for (let child of (element as mdast.Parent).children) {
            this.visit(child)
        }
    }

    getCurrentNode() {
        return this.nodeStack[this.nodeStack.length - 1]
    }

    pushNode(node: XML.Node) {
        this.nodeStack.push(node)
    }

    popNode() {
        this.nodeStack.pop()
    }

    pushNumberingStyle(style: NumberingStyle) {
        this.numberingStack.push(style)
    }

    popNumberingStyle() {
        this.numberingStack.pop()
    }

    getNumberingStyle() {
        return this.numberingStack[this.numberingStack.length - 1]
    }
}

const visitors: {
    [K in keyof mdast.RootContentMap]?: (source: mdast.RootContentMap[K], ctx: GenerationContext) => void;
} = {
    "paragraph": (node, ctx) => {
        let numberingStyle = ctx.getNumberingStyle()
        let styleId = numberingStyle ? numberingStyle.styleId : ctx.doc.styles.resource.getStyleByName("ispText_main").getId()
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        if (numberingStyle) {
            paragraph.getChild("w:pPr").pushChild(
                XML.Node.build("w:numPr").appendChildren([
                    XML.Node.build("w:ilvl").setAttr("w:val", String(ctx.numberingStack.length - 1)),
                    XML.Node.build("w:numId").setAttr("w:val", numberingStyle.numId),
                ])
            )
        }

        ctx.getCurrentNode().pushChild(paragraph)
        ctx.pushNode(paragraph)
        ctx.visitChildren(node)
        ctx.popNode()
    },

    "strong": (node, ctx) => {

    },

    "code": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("ispListing").getId()

        let raw = XML.Node.build("w:r")
        let lines = node.value.split("\n")
        for (let line of lines) {
            if (raw.getChildrenCount() > 0) {
                raw.pushChild(OXML.buildLineBreak())
            }
            raw.pushChild(OXML.buildTextTag(line))
        }

        ctx.getCurrentNode().pushChild(
            OXML.buildParagraphWithStyle(styleId)
                .pushChild(raw)
        )
    },

    "link": (node, ctx) => {

    },

    "delete": (node, ctx) => {

    },

    "emphasis": (node, ctx) => {

    },

    "text": (node, ctx) => {
        ctx.getCurrentNode().pushChild(OXML.buildRawTag(node.value))
    },

    "table": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("Table Grid").getId()

        let table = XML.Node.build("w:tbl")
        let tableProperties = XML.Node.build("w:tPr").appendChildren([
            XML.Node.build("w:tblStyle").setAttr("w:val", styleId)
        ])
        table.pushChild(tableProperties)
        ctx.getCurrentNode().pushChild(table)
        ctx.pushNode(table)
        ctx.visitChildren(node)
        ctx.popNode()
    },

    "tableRow": (node, ctx) => {
        let row = XML.Node.build("w:tr").appendChildren([])
        ctx.getCurrentNode().pushChild(row)
        ctx.pushNode(row)
        ctx.visitChildren(node)
        ctx.popNode()
    },

    "tableCell": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("ispText_main").getId()

        let paragraph = OXML.buildParagraphWithStyle(styleId)
        let cell = XML.Node.build("w:tc").appendChildren([
            paragraph
        ])

        ctx.getCurrentNode().pushChild(cell)
        ctx.pushNode(paragraph)
        ctx.visitChildren(node)
        ctx.popNode()
    },

    "inlineMath": ((node, ctx) => {
        let value = temml.renderToString(node.value)
        let omml = mml2omml(value)
        ctx.getCurrentNode().pushChild(XML.Node.fromXmlString(omml))
    }),

    "listItem": (node, ctx) => {
        ctx.visitChildren(node)
    },

    "image": (node, ctx) => {
        let imageBuffer = fs.readFileSync(node.url)
        let file = ImageSize(imageBuffer)

        let size = getImageSize(node, file.width / file.height)

        // LibreOffice seems to break when filenames contain bad symbols.
        let mediaName = path.basename(node.url).replace(/[^a-zA-Z0-9.]/g, "_")
        let mediaTarget = ctx.doc.getUniqueMediaTarget(mediaName)
        let mediaPath = ctx.doc.getPathForTarget(mediaTarget)

        ctx.doc.saveFile(mediaPath, imageBuffer)

        let unusedId = ctx.doc.document.rels.getUnusedId()
        ctx.doc.document.rels.addRelation({
            id: unusedId,
            target: mediaTarget,
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        })

        let image = createImage(unusedId, size[0], size[1])
        ctx.getCurrentNode().pushChild(XML.Node.build("w:r").pushChild(image))
    },

    "blockquote": (node, ctx) => {

    },

    "break": (node, ctx) => {

    },

    "heading": (node, ctx) => {
        let headings = [
            "ispSubHeader-1 level",
            "ispSubHeader-2 level",
            "ispSubHeader-3 level",
        ]

        let level = node.depth

        if (level > headings.length) {
            level = headings.length as typeof node.depth
        }

        let styleId = ctx.doc.styles.resource.getStyleByName(headings[level - 1]).getId()
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        ctx.getCurrentNode().pushChild(paragraph)
        ctx.pushNode(paragraph)
        ctx.visitChildren(node)
        ctx.popNode()
    },

    "imageReference": (node, ctx) => {

    },

    "list": (node, ctx) => {
        let numberingStyle = node.ordered ? DecimalNumberingStyle : BulletNumberingStyle

        let numId = createNumbering(ctx.doc, numberingStyle.numId, ctx.numberingStack.length, node.start)
        ctx.pushNumberingStyle({
            numId: numId,
            styleId: numberingStyle.styleId
        })
        ctx.visitChildren(node)
        ctx.popNumberingStyle()
    },

    "math": (node, ctx) => {

    },

    "inlineCode": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("ispListing Знак").getId()

        let raw = OXML.buildRawTag(node.value, [
            XML.Node.build("w:rStyle").setAttr("w:val", styleId)
        ])

        ctx.getCurrentNode().pushChild(raw)
    },
}

function parseSizeAttr(attr: string | number | undefined) {
    if (typeof attr !== "string") return null

    if (attr.endsWith("cm")) {
        let centimeters = parseFloat(attr.slice(0, -2))
        if (Number.isFinite(centimeters)) return centimeters
    }

    return null
}

function getImageSize(image: mdast.Image, aspect: number) {
    let width: number | null = parseSizeAttr(image.data?.attrs?.width)
    let height: number | null = parseSizeAttr(image.data?.attrs?.height)

    if (width === null && height === null) {
        width = 10
    }

    if (height === null) {
        height = width / aspect
    }

    if (width === null) {
        width = aspect * height
    }

    return [width, height]
}

function createNumbering(document: WordDocument, abstractId: string, depth: number, start: number) {
    let newNumId = document.numbering.resource.getUnusedNumId()
    let newNumbering = new Num().readXml(XML.Node.build("w:num"))
    newNumbering.setAbstractNumId(abstractId)
    newNumbering.setId(newNumId)

    newNumbering.getLevelOverride(depth).pushChild(
        XML.Node.build("w:startOverride")
            .setAttr("w:val", String(start))
    )

    document.numbering.resource.nums.set(newNumId, newNumbering)
    return newNumId
}

function createImage(id: string, width: number, height: number) {

    let widthEmu = Math.floor(width * 360000)
    let heightEmu = Math.floor(height * 360000)

    return XML.Node.build("w:drawing").appendChildren([
        XML.Node.build("wp:inline").appendChildren([
            XML.Node.build("a:graphic").appendChildren([
                XML.Node.build("a:graphicData").setAttrs({
                    "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                }).appendChildren([
                    XML.Node.build("pic:pic").appendChildren([
                        XML.Node.build("pic:nvPicPr").appendChildren([
                            XML.Node.build("pic:cNvPr").setAttrs({
                                "id": "pic" + id,
                                "name": "Picture"
                            }),
                            XML.Node.build("pic:cNvPicPr")
                        ]),
                        XML.Node.build("pic:blipFill").appendChildren([
                            XML.Node.build("a:blip").setAttr("r:embed", id),
                            XML.Node.build("a:stretch").appendChildren([
                                XML.Node.build("a:fillRect")
                            ])
                        ]),
                        XML.Node.build("pic:spPr").appendChildren([
                            XML.Node.build("a:xfrm").appendChildren([
                                XML.Node.build("a:off").setAttrs({
                                    "x": "0",
                                    "y": "0"
                                }),
                                XML.Node.build("a:ext").setAttrs({
                                    "cx": String(widthEmu),
                                    "cy": String(heightEmu)
                                })
                            ]),
                            XML.Node.build("a:prstGeom").setAttr("prst", "rect").appendChildren([
                                XML.Node.build("a:avLst")
                            ])
                        ])
                    ])
                ])
            ])
        ])
    ]);
}

export function generateDocxBody(source: mdast.Root, target: WordDocument, meta: DocumentJsonMeta) {
    new ParagraphTemplateSubstitution()
        .setDocument(target)
        .setTemplate("{{{body}}}")
        .setReplacement(() => {
            let context = new GenerationContext()

            // Fictive node
            let node = XML.Node.build("w:body")

            context.doc = target
            context.pushNode(node)

            context.visitChildren(source)

            return node.getChildren()
        })
        .perform()
}

function getLinksParagraphs(doc: WordDocument, meta: DocumentJsonMeta) {
    let styleId = doc.styles.resource.getStyleByName("ispLitList").getId()
    let numId = "80"
    let linksSection = meta.getSection("links").asArray()

    let result = []

    for (let link of linksSection) {
        let description: string
        // Backwards compatibility
        if (link.isMap()) {
            description = link.getString("description")
        } else {
            description = link.getString()
        }
        let paragraph = OXML.buildParagraphWithStyle(styleId)
        let style = paragraph.getChild("w:pPr")
        style.pushChild(OXML.buildNumPr("0", numId))

        paragraph.pushChild(OXML.buildRawTag(description))
        result.push(paragraph)
    }

    return result
}

function getAuthors(doc: WordDocument, meta: DocumentJsonMeta, language: string) {
    let styleId = doc.styles.resource.getStyleByName("ispAuthor").getId()
    let authors = meta.getSection("authors").asArray()
    let organizations = meta
        .getSection("organizations")
        .asArray()
        .map(section => section.getString("id"))

    let result = []

    for (let author of authors) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        let name = author.getString("name_" + language)
        let orcid = author.getString("orcid")
        let email = author.getString("email")
        let authorOrgs = author.getSection("organizations")
            .asArray()
            .map(section => section.getString())
            .map(id => organizations.indexOf(id) + 1)
            .join(",")

        let authorLine = `${name}, ORCID: ${orcid}, <${email}>`

        paragraph.pushChild(OXML.buildRawTag(authorOrgs, [OXML.buildSuperscriptTextStyle()]))
        paragraph.pushChild(OXML.buildRawTag(authorLine))

        result.push(paragraph)
    }

    return result
}

function getOrganizations(doc: WordDocument, meta: DocumentJsonMeta, language: string) {
    let styleId = doc.styles.resource.getStyleByName("ispAuthor").getId()
    let organizations = meta.getSection("organizations").asArray()

    let orgIndex = 1
    let result = []

    for (let organization of organizations) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        let indexLine = String(orgIndex)

        paragraph.pushChild(OXML.buildRawTag(indexLine, [OXML.buildSuperscriptTextStyle()]))
        paragraph.pushChild(OXML.buildRawTag(organization.getString("name_" + language)))

        result.push(paragraph)

        orgIndex++
    }

    return result
}

function getAuthorsDetail(doc: WordDocument, meta: DocumentJsonMeta) {
    let styleId = doc.styles.resource.getStyleByName("ispText_main").getId()
    let authors = meta.getSection("authors").asArray()

    let result = []

    for (let author of authors) {
        for (let language of languages) {
            let line = author.getString("details_" + language)
            let newParagraph = OXML.buildParagraphWithStyle(styleId)
            newParagraph.getChild("w:pPr").pushChild(
                XML.Node.build("w:spacing")
                    .setAttr("w:before", "30")
                    .setAttr("w:after", "120")
            )
            newParagraph.pushChild(OXML.buildRawTag(line))
            result.push(newParagraph)
        }
    }

    return result
}

export function substituteTemplates(document: WordDocument, meta: DocumentJsonMeta) {
    let inlineSubstitution = new InlineTemplateSubstitution().setDocument(document)
    let paragraphSubstitution = new ParagraphTemplateSubstitution().setDocument(document)

    for (let language of languages) {
        let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"]
        for (let template of templates) {

            let template_lang = template + "_" + language
            let replacement = meta.getString(template_lang)

            inlineSubstitution
                .setTemplate("{{{" + template_lang + "}}}")
                .setReplacement(replacement)
                .perform()
        }

        let header = meta.getString("page_header_" + language)

        if (header === "@use_citation") {
            header = meta.getString("for_citation_" + language)
        }

        inlineSubstitution
            .setTemplate("{{{page_header_" + language + "}}}")
            .setReplacement(header)
            .perform()

        paragraphSubstitution
            .setTemplate("{{{authors_" + language + "}}}")
            .setReplacement(() => getAuthors(document, meta, language))
            .perform()

        paragraphSubstitution
            .setTemplate("{{{organizations_" + language + "}}}")
            .setReplacement(() => getOrganizations(document, meta, language))
            .perform()
    }

    paragraphSubstitution
        .setTemplate("{{{links}}}")
        .setReplacement(() => getLinksParagraphs(document, meta))
        .perform()

    paragraphSubstitution
        .setTemplate("{{{authors_detail}}}")
        .setReplacement(() => getAuthorsDetail(document, meta))
        .perform()
}