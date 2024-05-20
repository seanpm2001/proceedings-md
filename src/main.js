import * as path from 'path';
import path__default from 'path';
import * as fs from 'fs';
import fs__default from 'fs';
import * as process from 'process';
import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import yaml from 'yaml';
import { fromMarkdown } from 'mdast-util-from-markdown';
import { frontmatterFromMarkdown } from 'mdast-util-frontmatter';
import { frontmatter } from 'micromark-extension-frontmatter';
import { math } from 'micromark-extension-math';
import { mathFromMarkdown } from 'mdast-util-math';
import { gfmTableFromMarkdown } from 'mdast-util-gfm-table';
import { gfmTable } from 'micromark-extension-gfm-table';
import { visit } from 'unist-util-visit';
import ImageSize from 'image-size';
import temml from 'temml';
import { mml2omml } from 'mathml2omml';

const keys = {
    comment: "__comment__",
    text: "__text__",
    attributes: ":@",
    document: "__document__"
};
const parser = new XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: keys.comment,
    textNodeName: keys.text
});
const builder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: keys.comment,
    textNodeName: keys.text
});
function checkFilter(filter, node) {
    if (!filter)
        return true;
    if (typeof filter === "string") {
        return node.getTagName() === filter;
    }
    return filter(node);
}
function getVisitArgs(args) {
    let filter = null;
    let callback = args[0];
    let startPosition = args[1];
    if (typeof args[1] === "function") {
        filter = args[0];
        callback = args[1];
        startPosition = args[2];
    }
    return {
        filter: filter,
        callback: callback,
        startPosition: startPosition
    };
}
class Node {
    element;
    tempDestroyed = false;
    constructor(element) {
        if (Array.isArray(element)) {
            throw new Error("XML.Node must be constructed from the xml object, not its children list");
        }
        this.element = element;
    }
    getTagName() {
        this.checkTemporary();
        for (let key of Object.getOwnPropertyNames(this.element)) {
            // Be safe from prototype pollution
            if (key === "__proto__" || key === keys.attributes)
                continue;
            return key;
        }
        return null;
    }
    pushChild(child) {
        this.checkTemporary();
        let children = this.getRawChildren();
        if (children === null) {
            throw new Error("Cannot call pushChild on " + this.getTagName() + " element");
        }
        children.push(child.raw());
        return this;
    }
    unshiftChild(child) {
        this.checkTemporary();
        let children = this.getRawChildren();
        if (children === null) {
            throw new Error("Cannot call unshiftChild on " + this.getTagName() + " element");
        }
        children.unshift(child.raw());
        return this;
    }
    getChildren(filter = null) {
        this.checkTemporary();
        let result = [];
        this.visitChildren(filter, (child) => {
            result.push(child.shallowCopy());
        });
        return result;
    }
    getChild(arg = null) {
        this.checkTemporary();
        if (Array.isArray(arg)) {
            let path = arg;
            if (path.length === 0) {
                return this;
            }
            let result = new Node(this.element);
            for (let i = 0; i < path.length; i++) {
                if (!result.element)
                    return null;
                let tagName = result.getTagName();
                let pathComponent = path[i];
                let children = result.element[tagName];
                if (pathComponent < 0) {
                    result.element = children[children.length + pathComponent];
                }
                else {
                    result.element = children[pathComponent];
                }
            }
            if (!result.element)
                return null;
            return result;
        }
        else {
            let filter = arg;
            let result = null;
            this.visitChildren(filter, (child) => {
                if (result) {
                    throw new Error("Element have multiple children matching the given filter");
                }
                result = child.shallowCopy();
            });
            return result;
        }
    }
    visitChildren(...args) {
        this.checkTemporary();
        let { filter, callback, startPosition } = getVisitArgs(args);
        let tagName = this.getTagName();
        if (!Array.isArray(this.element[tagName])) {
            return;
        }
        let index = startPosition ?? 0;
        let tmpNode = new Node(null);
        for (let child of this.element[tagName]) {
            tmpNode.element = child;
            if (checkFilter(filter, tmpNode)) {
                if (callback(tmpNode, index) === false) {
                    break;
                }
            }
            index++;
        }
        tmpNode.markDestroyed();
    }
    visitSubtree(...args) {
        this.checkTemporary();
        let { filter, callback, startPosition } = getVisitArgs(args);
        let tmpNode = new Node(null);
        let startPath = startPosition ?? [];
        let startDepth = 0;
        let path = [];
        const walk = (node) => {
            let tagName = node.getTagName();
            let children = node.element[tagName];
            if (!Array.isArray(children)) {
                return;
            }
            let depth = path.length;
            let startIndex = 0;
            if (depth < startDepth && startPath.length) {
                startIndex = startPath[startPath.length];
                startDepth = depth;
            }
            for (let index = startIndex; index < children.length; index++) {
                path.push(index);
                tmpNode.element = children[index];
                let filterPass = checkFilter(filter, tmpNode);
                let goDeeper = true;
                if (filterPass) {
                    goDeeper = callback(tmpNode, path) === true;
                }
                if (goDeeper) {
                    walk(tmpNode);
                }
                // Handle path modification
                index = path[path.length - 1];
                path.pop();
            }
        };
        walk(this);
        tmpNode.markDestroyed();
    }
    removeChild(path) {
        if (path.length === 0) {
            throw new Error("Cannot call removeChild with empty path");
        }
        let topIndex = path.pop();
        let child = this.getChild(path);
        let childChildren = child.getRawChildren();
        if (childChildren === null) {
            throw new Error("Cannot call removeChild for " + child.getTagName() + " element");
        }
        childChildren.splice(topIndex, 1);
        path.push(topIndex);
    }
    removeChildren(filter = null) {
        this.checkTemporary();
        let children = this.getRawChildren();
        if (children === null) {
            throw new Error("Cannot call removeChildren on " + this.getTagName() + " element");
        }
        let node = new Node(null);
        for (let i = 0; i < children.length; i++) {
            node.element = children[i];
            if (checkFilter(filter, node)) {
                children.splice(i, 1);
                i--;
            }
        }
        node.markDestroyed();
    }
    isTextNode() {
        this.checkTemporary();
        return this.getTagName() == keys.text;
    }
    isCommentNode() {
        this.checkTemporary();
        return this.getTagName() == keys.comment;
    }
    getText() {
        this.checkTemporary();
        if (!this.isTextNode()) {
            throw new Error("getText() is called on " + this.getTagName() + " element");
        }
        return String(this.element[keys.text]);
    }
    setText(text) {
        this.checkTemporary();
        if (!this.isTextNode()) {
            throw new Error("setText() is called on " + this.getTagName() + " element");
        }
        this.element[keys.text] = text;
    }
    getComment() {
        this.checkTemporary();
        if (!this.isCommentNode()) {
            throw new Error("getComment() is called on " + this.getTagName() + " element");
        }
        let textChild = this.getChild(keys.text);
        return textChild.getText();
    }
    static build(tagName) {
        let element = {};
        element[tagName] = [];
        return new Node(element);
    }
    static createDocument(args = {}) {
        args = Object.assign({
            version: "1.0",
            encoding: "UTF-8",
            standalone: "yes"
        }, args);
        let document = this.build(keys.document);
        document.appendChildren([
            Node.build("?xml")
                .setAttrs(args)
                .appendChildren([
                Node.buildTextNode("")
            ])
        ]);
        return document;
    }
    static buildTextNode(text) {
        let element = {};
        element[keys.text] = text;
        return new Node(element);
    }
    setAttr(attribute, value) {
        this.checkTemporary();
        if (!this.element[keys.attributes]) {
            this.element[keys.attributes] = {};
        }
        this.element[keys.attributes][attribute] = value;
        return this;
    }
    setAttrs(attributes) {
        this.checkTemporary();
        this.element[keys.attributes] = attributes;
        return this;
    }
    getAttrs() {
        if (!this.element[keys.attributes]) {
            this.element[keys.attributes] = {};
        }
        return this.element[keys.attributes];
    }
    getAttr(attribute) {
        this.checkTemporary();
        let attrs = this.getAttrs();
        let attr = attrs[attribute];
        if (attr === undefined)
            return null;
        return String(attr);
    }
    clearChildren(path = []) {
        this.checkTemporary();
        let parent = this.getChild(path);
        parent.element[parent.getTagName()] = [];
        return this;
    }
    insertChildren(children, path) {
        this.checkTemporary();
        if (path.length === 0) {
            throw new Error("Cannot call insertChildren with empty path");
        }
        let insertIndex = path.pop();
        let parent = this.getChild(path);
        path.push(insertIndex);
        let lastChildren = parent.getRawChildren();
        if (lastChildren === null) {
            throw new Error("Cannot call insertChildren for " + parent.getTagName() + " element");
        }
        if (insertIndex < 0) {
            insertIndex = children.length + insertIndex + 1;
        }
        lastChildren.splice(insertIndex, 0, ...children.map(child => child.raw()));
        return this;
    }
    appendChildren(children, path = []) {
        path.push(-1);
        this.insertChildren(children, path);
        path.pop();
        return this;
    }
    unshiftChildren(children, path = []) {
        path.push(0);
        this.insertChildren(children, path);
        path.pop();
        return this;
    }
    assign(another) {
        this.checkTemporary();
        if (this === another) {
            return;
        }
        if (this.element) {
            this.element[this.getTagName()] = undefined;
        }
        else {
            this.element = {};
        }
        this.element[another.getTagName()] = JSON.parse(JSON.stringify(another.raw()[another.getTagName()]));
        if (another.raw()[keys.attributes]) {
            this.element[keys.attributes] = JSON.parse(JSON.stringify(another.raw()[keys.attributes]));
        }
        else {
            this.element[keys.attributes] = {};
        }
        return this;
    }
    static fromXmlString(str) {
        let object = parser.parse(str);
        let wrapped = {};
        wrapped[keys.document] = object;
        return new Node(wrapped);
    }
    toXmlString() {
        this.checkTemporary();
        if (this.getTagName() === keys.document) {
            return builder.build(this.element[keys.document]);
        }
        else {
            return builder.build([this.element]);
        }
    }
    raw() {
        this.checkTemporary();
        return this.element;
    }
    checkTemporary() {
        if (this.tempDestroyed) {
            throw new Error("Method access to an outdated temporary Node. Make sure to call .shallowCopy() on temporary " +
                "nodes before accessing them outside your visitChildren/visitSubtree body scope");
        }
    }
    markDestroyed() {
        this.checkTemporary();
        // From now on, the checkTemporary method will throw
        this.tempDestroyed = true;
    }
    getRawContents() {
        this.checkTemporary();
        return this.element[this.getTagName()];
    }
    getRawChildren() {
        this.checkTemporary();
        let contents = this.getRawContents();
        if (Array.isArray(contents)) {
            return contents;
        }
        return null;
    }
    shallowCopy() {
        this.checkTemporary();
        return new Node(this.element);
    }
    deepCopy() {
        this.checkTemporary();
        return new Node(null).assign(this);
    }
    isLeaf() {
        this.checkTemporary();
        return this.getRawChildren() === null;
    }
    getChildrenCount() {
        return this.getRawChildren()?.length ?? 0;
    }
}
class Serializable {
    readXmlString(xmlString) {
        this.readXml(Node.fromXmlString(xmlString));
        return this;
    }
    readXml(xml) {
        throw new Error("readXml is not implemented");
    }
    toXmlString() {
        return this.toXml().toXmlString();
    }
    toXml() {
        throw new Error("toXml is not implemented");
    }
}
class Wrapper extends Serializable {
    node = null;
    readXml(xml) {
        this.node = xml;
        return this;
    }
    toXml() {
        return this.node;
    }
}
function getNamespace(name) {
    let parts = name.split(":");
    if (parts.length >= 2) {
        return parts[0];
    }
    return null;
}
function* getUsedNames(tag) {
    let tagName = tag.getTagName();
    yield tagName;
    let attributes = tag.getAttrs();
    for (let key of Object.getOwnPropertyNames(attributes)) {
        // Be safe from prototype pollution
        if (key === "__proto__")
            continue;
        yield key;
    }
}
function getTextContents(tag) {
    let result = "";
    tag.visitSubtree(keys.text, (node) => {
        result += node.getText();
    });
    return result;
}

const wordXmlns = new Map([
    ["wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"],
    ["mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"],
    ["o", "urn:schemas-microsoft-com:office:office"],
    ["r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["v", "urn:schemas-microsoft-com:vml"],
    ["wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"],
    ["wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
    ["w10", "urn:schemas-microsoft-com:office:word"],
    ["w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
    ["w14", "http://schemas.microsoft.com/office/word/2010/wordml"],
    ["w15", "http://schemas.microsoft.com/office/word/2012/wordml"],
    ["wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"],
    ["wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"],
    ["wne", "http://schemas.microsoft.com/office/word/2006/wordml"],
    ["wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"],
    ["pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
    ["a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
]);
const wordXmlnsIgnorable = new Set(["wp14", "w14", "w15"]);
function getProperXmlnsForDocument(document) {
    let result = {};
    let ignorable = new Set();
    document.visitSubtree((child) => {
        for (let name of getUsedNames(child)) {
            let namespace = getNamespace(name);
            if (!namespace || !wordXmlns.has(namespace)) {
                continue;
            }
            result["xmlns:" + namespace] = wordXmlns.get(namespace);
            if (wordXmlnsIgnorable.has(namespace)) {
                ignorable.add(namespace);
            }
        }
        return true;
    });
    if (ignorable.size) {
        result["xmlns:mc"] = wordXmlns.get("mc");
        result["mc:Ignorable"] = Array.from(ignorable).join(" ");
    }
    return result;
}
function buildParagraphWithStyle(style) {
    return Node.build("w:p").appendChildren([
        Node.build("w:pPr").appendChildren([
            Node.build("w:pStyle").setAttr("w:val", style)
        ])
    ]);
}
function buildNumPr(ilvl, numId) {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>
    return Node.build("w:numPr").appendChildren([
        Node.build("w:ilvl").setAttr("w:val", "0"),
        Node.build("w:numId").setAttr("w:val", numId),
    ]);
}
function buildSuperscriptTextStyle() {
    return Node.build("w:vertAlign").setAttr("w:val", "superscript");
}
function buildLineBreak() {
    return Node.build("w:br");
}
function buildTextTag(text) {
    return Node.build("w:t")
        .setAttr("xml:space", "preserve")
        .appendChildren([
        Node.buildTextNode(text)
    ]);
}
function buildRawTag(text, styles) {
    let result = Node.build("w:r");
    if (styles) {
        result.pushChild(Node.build("w:rPr").appendChildren(styles));
    }
    result.pushChild(buildTextTag(text));
    return result;
}
function getChildVal(node, tag) {
    let child = node.getChild(tag);
    if (child)
        return child.getAttr("w:val");
    return null;
}
function setChildVal(node, tag, value) {
    if (value === null) {
        node.removeChildren(tag);
    }
    else {
        let basedOnTag = node.getChild(tag);
        if (basedOnTag)
            basedOnTag.setAttr("w:val", value);
        else
            node.appendChildren([
                Node.build(tag).setAttr("w:val", value)
            ]);
    }
}
function fixXmlns(document, rootTag) {
    document.getChild(rootTag).setAttrs(getProperXmlnsForDocument(document));
}
function normalizePath(pathString) {
    pathString = path__default.posix.normalize(pathString);
    if (!pathString.startsWith("/")) {
        pathString = "/" + pathString;
    }
    return pathString;
}
function getRelsPath(resourcePath) {
    let basename = path__default.basename(resourcePath);
    let dirname = path__default.dirname(resourcePath);
    return normalizePath(dirname + "/_rels/" + basename + ".rels");
}

class Relationships extends Serializable {
    relations = new Map();
    readXml(xml) {
        this.relations = new Map();
        xml.getChild("Relationships")?.visitChildren("Relationship", (child) => {
            let id = child.getAttr("Id");
            let type = child.getAttr("Type");
            let target = child.getAttr("Target");
            if (id !== undefined && type !== undefined && target !== undefined) {
                this.addRelation({
                    id: id,
                    type: type,
                    target: target
                });
            }
        });
        return this;
    }
    toXml() {
        let relations = Array.from(this.relations.values());
        return Node.createDocument().appendChildren([
            Node.build("Relationships")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
                .appendChildren(relations.map((relation) => {
                return Node.build("Relationship")
                    .setAttr("Id", relation.id)
                    .setAttr("Type", relation.type)
                    .setAttr("Target", relation.target);
            }))
        ]);
    }
    getUnusedId() {
        let prefix = "rId";
        let index = 1;
        while (this.relations.has(prefix + index)) {
            index++;
        }
        return prefix + index;
    }
    addRelation(relation) {
        this.relations.set(relation.id, relation);
    }
    getRelForTarget(target) {
        for (let rel of this.relations.values()) {
            if (rel.target === target) {
                return rel;
            }
        }
    }
}

class ContentTypes extends Serializable {
    defaults;
    overrides;
    readXml(xml) {
        this.defaults = [];
        this.overrides = [];
        let types = xml.getChild("Types");
        types?.visitChildren("Default", (child) => {
            let extension = child.getAttr("Extension");
            let contentType = child.getAttr("ContentType");
            if (extension !== undefined && contentType !== undefined) {
                this.defaults.push({
                    extension: extension,
                    contentType: contentType
                });
            }
        });
        types?.visitChildren("Override", (child) => {
            let partName = child.getAttr("PartName");
            let contentType = child.getAttr("ContentType");
            if (partName !== undefined && contentType !== undefined) {
                this.overrides.push({
                    partName: partName,
                    contentType: contentType
                });
            }
        });
        return this;
    }
    toXml() {
        return Node.createDocument().appendChildren([
            Node.build("Types")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
                .appendChildren(this.defaults.map((def) => {
                return Node.build("Default")
                    .setAttr("Extension", def.extension)
                    .setAttr("ContentType", def.contentType);
            }))
                .appendChildren(this.overrides.map((override) => {
                return Node.build("Override")
                    .setAttr("PartName", override.partName)
                    .setAttr("ContentType", override.contentType);
            }))
        ]);
    }
    getContentTypeForExt(ext) {
        for (let def of this.defaults) {
            if (def.extension === ext)
                return def.contentType;
        }
        return null;
    }
    getOverrideForPartName(partName) {
        for (let override of this.overrides) {
            if (override.partName === partName)
                return override.contentType;
        }
        return null;
    }
    getContentTypeForPath(pathString) {
        pathString = normalizePath(pathString);
        let overrideContentType = this.getOverrideForPartName(pathString);
        if (overrideContentType !== null) {
            return overrideContentType;
        }
        const extension = path__default.extname(pathString).slice(1);
        return this.getContentTypeForExt(extension);
    }
    join(other) {
        for (let otherDef of other.defaults) {
            if (this.getContentTypeForExt(otherDef.extension) === null) {
                this.defaults.push({
                    ...otherDef
                });
            }
        }
        for (let otherOverride of other.overrides) {
            if (this.getOverrideForPartName(otherOverride.partName) === null) {
                this.overrides.push({
                    ...otherOverride
                });
            }
        }
    }
}

class Style extends Wrapper {
    getBaseStyle() {
        return getChildVal(this.node, "w:basedOn");
    }
    getLinkedStyle() {
        return getChildVal(this.node, "w:link");
    }
    getNextStyle() {
        return getChildVal(this.node, "w:link");
    }
    getName() {
        return getChildVal(this.node, "w:name");
    }
    getId() {
        return this.node.getAttr("w:styleId");
    }
    setBaseStyle(style) {
        setChildVal(this.node, "w:basedOn", style);
    }
    setLinkedStyle(style) {
        setChildVal(this.node, "w:link", style);
    }
    setNextStyle(style) {
        setChildVal(this.node, "w:next", style);
    }
    setName(name) {
        setChildVal(this.node, "w:name", name);
    }
    setId(id) {
        this.node.setAttr("w:styleId", id);
    }
}
class LatentStyles extends Wrapper {
    readOrCreate(node) {
        if (!node) {
            node = Node.build("w:latentStyles");
        }
        return this.readXml(node);
    }
    getLsdExceptions() {
        let result = new Map();
        this.node.visitChildren("w:lsdException", (child) => {
            let lsdException = new LSDException().readXml(child.shallowCopy());
            result.set(lsdException.name, lsdException);
        });
    }
}
class DocDefaults extends Wrapper {
    readOrCreate(node) {
        if (!node) {
            node = Node.build("w:docDefaults");
        }
        return this.readXml(node);
    }
}
class LSDException extends Wrapper {
    name;
    readXml(node) {
        this.name = node.getAttr("w:name");
        return this;
    }
    setName(name) {
        this.name = name;
        this.node.setAttr("w:name", name);
        return this;
    }
}
class Styles extends Serializable {
    styles = new Map();
    docDefaults = null;
    latentStyles = null;
    rels;
    readXml(xml) {
        this.styles = new Map();
        let styles = xml.getChild("w:styles");
        this.latentStyles = new LatentStyles().readOrCreate(styles.getChild("w:latentStyles"));
        this.docDefaults = new DocDefaults().readOrCreate(styles.getChild("w:docDefaults"));
        styles?.visitChildren("w:style", (child) => {
            let style = new Style().readXml(child.shallowCopy());
            this.styles.set(style.getId(), style);
        });
        return this;
    }
    toXml() {
        let styles = Array.from(this.styles.values());
        let result = Node.createDocument().appendChildren([
            Node.build("w:styles")
                .appendChildren([
                this.docDefaults.node.deepCopy(),
                this.latentStyles.node.deepCopy()
            ])
                .appendChildren(styles.map((style) => {
                return style.node.deepCopy();
            }))
        ]);
        result.getChild("w:styles").setAttrs(getProperXmlnsForDocument(result));
        return result;
    }
    removeStyle(style) {
        this.styles.delete(style.getId());
    }
    addStyle(style) {
        this.styles.set(style.getId(), style);
    }
    getStyleByName(name) {
        for (let [id, style] of this.styles) {
            if (style.getName() === name)
                return style;
        }
        return null;
    }
}

function getNodeLevel(node, tagName, index) {
    let result;
    let strIndex = String(index);
    node.visitChildren(tagName, (child) => {
        if (child.getAttr("w:ilvl") === strIndex) {
            result = child.shallowCopy();
            return false;
        }
    });
    if (!result) {
        result = Node.build(tagName).setAttr("w:ilvl", strIndex);
        node.appendChildren([result]);
    }
    return result;
}
class AbstractNum extends Wrapper {
    getId() {
        return this.node.getAttr("w:abstractNumId");
    }
    getLevel(index) {
        return getNodeLevel(this.node, "w:lvl", index);
    }
}
class Num extends Wrapper {
    getId() {
        return this.node.getAttr("w:numId");
    }
    getAbstractNumId() {
        return getChildVal(this.node, "w:abstractNumId");
    }
    getLevelOverride(index) {
        return getNodeLevel(this.node, "w:lvlOverride", index);
    }
    setId(id) {
        this.node.setAttr("w:numId", id);
    }
    setAbstractNumId(id) {
        setChildVal(this.node, "w:abstractNumId", id);
    }
}
class Numbering extends Serializable {
    abstractNums = new Map();
    nums = new Map();
    readXml(xml) {
        let styles = xml.getChild("w:numbering");
        styles?.visitChildren("w:abstractNum", (child) => {
            let abstractNum = new AbstractNum().readXml(child.shallowCopy());
            this.abstractNums.set(abstractNum.getId(), abstractNum);
        });
        styles?.visitChildren("w:num", (child) => {
            let num = new Num().readXml(child.shallowCopy());
            this.nums.set(num.getId(), num);
        });
        return this;
    }
    toXml() {
        let abstractNums = Array.from(this.abstractNums.values());
        let nums = Array.from(this.nums.values());
        return Node.createDocument().appendChildren([
            Node.build("w:numbering")
                .appendChildren(abstractNums.map((style) => {
                return style.node.deepCopy();
            }))
                .appendChildren(nums.map((style) => {
                return style.node.deepCopy();
            }))
        ]);
    }
    getUnusedNumId() {
        let index = 1;
        while (this.nums.has(String(index))) {
            index += 2;
        }
        return String(index);
    }
}

function getResourceTypeForMimeType(mimeType) {
    for (let key of Object.getOwnPropertyNames(resourceTypes)) {
        if (key === "__proto__")
            continue;
        if (resourceTypes[key].mimeType === mimeType) {
            return key;
        }
    }
}
const resourceTypes = {
    app: {
        mimeType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        xmlnsTag: "Properties"
    },
    core: {
        mimeType: "application/vnd.openxmlformats-package.core-properties+xml",
        xmlnsTag: "cp:coreProperties"
    },
    custom: {
        mimeType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        xmlnsTag: "Properties"
    },
    document: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        xmlnsTag: "w:document"
    },
    relationships: {
        mimeType: "application/vnd.openxmlformats-package.relationships+xml",
        xmlnsTag: "Relationships"
    },
    webSettings: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
        xmlnsTag: "webSettings"
    },
    numbering: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        xmlnsTag: "w:numbering"
    },
    settings: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        xmlnsTag: "w:settings"
    },
    styles: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        xmlnsTag: "w:styles"
    },
    fontTable: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        xmlnsTag: "w:fonts"
    },
    theme: {
        mimeType: "application/vnd.openxmlformats-officedocument.theme+xml",
        xmlnsTag: "a:theme"
    },
    comments: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        xmlnsTag: "w:comments"
    },
    footnotes: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        xmlnsTag: "w:footnotes"
    },
    footer: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        xmlnsTag: "w:ftr"
    },
    header: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        xmlnsTag: "w:hdr"
    },
    png: {
        mimeType: "image/png"
    },
};

function* uniqueNameGenerator(name) {
    let index = 0;
    while (true) {
        let nameCandidate = name;
        if (index > 0)
            nameCandidate += "_" + index;
        yield nameCandidate;
        index++;
    }
}

const contentTypesPath = "/[Content_Types].xml";
const globalRelsPath = "/_rels/.rels";
class WordResource {
    document;
    path;
    resource;
    rels = null;
    constructor(document, path, resource) {
        this.document = document;
        this.path = path;
        this.resource = resource;
    }
    saveRels() {
        if (!this.rels)
            return;
        let relsXml = this.rels.toXml();
        this.document.saveXml(getRelsPath(this.path), relsXml);
    }
    save() {
        let xml = this.resource.toXml();
        let contentType = this.document.contentTypes.resource.getContentTypeForPath(this.path);
        if (contentType) {
            let resourceType = getResourceTypeForMimeType(contentType);
            if (resourceType) {
                fixXmlns(xml, resourceTypes[resourceType].xmlnsTag);
            }
        }
        this.document.saveXml(this.path, xml);
        this.saveRels();
    }
    setRels(rels) {
        this.rels = rels;
        return this;
    }
}
const ResourceFactories = {
    generic: (document, path, xml) => {
        return new WordResource(document, path, new Wrapper().readXml(xml));
    },
    genericWithRel: (document, path, xml, rel) => {
        return new WordResource(document, path, new Wrapper().readXml(xml)).setRels(rel);
    },
    styles: (document, path, xml) => {
        return new WordResource(document, path, new Styles().readXml(xml));
    },
    numbering: (document, path, xml) => {
        return new WordResource(document, path, new Numbering().readXml(xml));
    },
    relationships: (document, path, xml) => {
        return new WordResource(document, path, new Relationships().readXml(xml));
    },
    contentTypes: (document, path, xml) => {
        return new WordResource(document, path, new ContentTypes().readXml(xml));
    },
};
class WordDocument {
    zipContents;
    wrappers = new Map();
    contentTypes;
    globalRels;
    numbering;
    styles;
    document;
    settings;
    fontTable;
    comments;
    headers = [];
    footers = [];
    async load(path) {
        const contents = await fs__default.promises.readFile(path);
        this.zipContents = await JSZip.loadAsync(contents);
        this.contentTypes = await this.createResourceForPath(ResourceFactories.contentTypes, contentTypesPath);
        this.globalRels = await this.createResourceForPath(ResourceFactories.relationships, globalRelsPath);
        this.document = await this.createResourceForType(ResourceFactories.genericWithRel, resourceTypes.document);
        this.styles = await this.createResourceForType(ResourceFactories.styles, resourceTypes.styles);
        this.settings = await this.createResourceForType(ResourceFactories.generic, resourceTypes.settings);
        this.numbering = await this.createResourceForType(ResourceFactories.numbering, resourceTypes.numbering);
        this.fontTable = await this.createResourceForType(ResourceFactories.generic, resourceTypes.fontTable);
        this.comments = await this.createResourceForType(ResourceFactories.generic, resourceTypes.comments);
        this.headers = await this.createResourcesForType(ResourceFactories.generic, resourceTypes.header);
        this.footers = await this.createResourcesForType(ResourceFactories.generic, resourceTypes.footer);
        return this;
    }
    getSinglePathForMimeType(type) {
        let paths = this.getPathsForMimeType(type);
        if (paths.length !== 1)
            return null;
        return paths[0];
    }
    async createResourceForType(factory, type) {
        let path = this.getSinglePathForMimeType(type.mimeType);
        if (!path)
            return null;
        return await this.createResourceForPath(factory, path);
    }
    async createResourcesForType(factory, type) {
        let paths = this.getPathsForMimeType(type.mimeType);
        return await Promise.all(paths.map(path => this.createResourceForPath(factory, path)));
    }
    async createResourceForPath(factory, pathString) {
        pathString = normalizePath(pathString);
        if (this.wrappers.has(pathString)) {
            throw new Error("This resource have already been created");
        }
        let relsPath = getRelsPath(pathString);
        let relationships = null;
        let relationshipsXml = await this.getXml(relsPath);
        if (relationshipsXml) {
            relationships = new Relationships().readXml(relationshipsXml);
        }
        let resource = factory(this, pathString, await this.getXml(pathString), relationships);
        this.wrappers.set(pathString, resource);
        return resource;
    }
    getPathsForMimeType(type) {
        let result = [];
        this.zipContents.forEach((path) => {
            let mimeType = this.contentTypes.resource.getContentTypeForPath(path);
            if (mimeType === type) {
                result.push(path);
            }
        });
        return result;
    }
    hasFile(path) {
        return this.zipContents.file(path) !== null;
    }
    async getFile(path) {
        return await this.zipContents.file(path.slice(1)).async("arraybuffer");
    }
    async getXml(path) {
        let contents = this.zipContents.file(path.slice(1));
        if (!contents)
            return null;
        return Node.fromXmlString(await contents.async("string"));
    }
    getPathForTarget(target) {
        return "/word/" + target;
    }
    getUniqueMediaTarget(name) {
        let targetPath = "media/";
        let extension = path__default.extname(name);
        let basename = path__default.basename(name, extension);
        for (let uniqueBasename of uniqueNameGenerator(basename)) {
            let name = targetPath + uniqueBasename + extension;
            if (!this.zipContents.file(this.getPathForTarget(name))) {
                return name;
            }
        }
    }
    saveFile(path, data) {
        this.zipContents.file(path.slice(1), data);
    }
    saveXml(path, xml) {
        this.zipContents.file(path.slice(1), xml.toXmlString());
    }
    async save(path) {
        for (let [path, resource] of this.wrappers) {
            resource.save();
        }
        const contents = await this.zipContents.generateAsync({ type: "uint8array" });
        await fs__default.writeFileSync(path, contents);
    }
}

function metaType(value) {
    if (Array.isArray(value))
        return "array";
    return typeof value;
}
class DocumentJsonMeta {
    section;
    path;
    constructor(section, path = "") {
        this.section = section;
        this.path = path;
    }
    getSection(path) {
        let any = this.getChild(path);
        return new DocumentJsonMeta(any, this.getAbsPath(path));
    }
    isArray() {
        if (this.section === undefined) {
            return false;
        }
        else
            return Array.isArray(this.section);
    }
    isMap() {
        if (this.section === undefined) {
            return false;
        }
        else
            return (typeof this.section === "object" && !this.isArray());
    }
    asArray() {
        if (this.section === undefined) {
            this.reportNotExistError("", "array");
        }
        else if (!this.isArray()) {
            this.reportWrongTypeError("", "array", metaType(this.section));
        }
        else {
            return this.section.map((element, index) => {
                return new DocumentJsonMeta(element, this.getAbsPath(String(index)));
            });
        }
    }
    getKeys() {
        if (!this.section) {
            this.reportNotExistError("", "object");
        }
        else if (!this.isMap()) {
            this.reportWrongTypeError("", "object", metaType(this.section));
        }
        else {
            return Object.keys(this.section);
        }
    }
    getString(path = "") {
        let child = this.getChild(path);
        if (!child) {
            this.reportNotExistError(path, "MetaInlines");
        }
        else if (typeof child !== "string") {
            this.reportWrongTypeError(path, "MetaInlines", metaType(child));
        }
        else {
            return child;
        }
    }
    reportNotExistError(relPath, expected) {
        let absPath = this.getAbsPath(relPath);
        throw new Error("Failed to parse document metadata: expected to have " + expected + " at path " + absPath);
    }
    reportWrongTypeError(relPath, expected, actual) {
        let absPath = this.getAbsPath(relPath);
        throw new Error("Failed to parse document metadata: expected " + expected + " at path " + absPath + ", got " +
            actual + " instead");
    }
    getAbsPath(relPath) {
        if (this.path.length) {
            if (relPath.length) {
                return this.path + "." + relPath;
            }
            return this.path;
        }
        return relPath;
    }
    getChild(path) {
        if (!path.length)
            return this.section;
        let result = this.section;
        for (let component of path.split(".")) {
            // Be safe from prototype pollution
            if (component === "__proto__")
                return undefined;
            if (!result)
                return undefined;
            if (Array.isArray(result)) {
                let index = Number.parseInt(component);
                if (!Number.isNaN(index)) {
                    result = result[index];
                }
            }
            else if (typeof result === "object") {
                result = result[component];
            }
            else {
                return undefined;
            }
        }
        return result;
    }
    static fromMarkdown(markdown) {
        let child = markdown.children[0];
        if (!child)
            return null;
        if (child.type !== "yaml")
            return null;
        return new DocumentJsonMeta(yaml.parse(child.value));
    }
}

class DocumentReferences {
    stack = [];
    depthThreshold = 1;
    groups = new Map();
    meta;
    constructor(meta) {
        this.meta = meta;
    }
    getPrefixMap(prefix) {
        let map = this.groups.get(prefix);
        if (!map) {
            map = new Map();
            this.groups.set(prefix, map);
        }
        return map;
    }
    getSection(depth, id) {
        depth = Math.max(0, depth - this.depthThreshold);
        while (this.stack.length > depth)
            this.stack.pop();
        if (this.stack.length === depth) {
            this.stack[this.stack.length - 1]++;
        }
        else {
            while (this.stack.length < depth)
                this.stack.push(1);
        }
        if (this.stack.length === 0) {
            return;
        }
        let result = this.stack.join(".");
        if (id) {
            let prefix = id.split(":")[0];
            let map = this.getPrefixMap(prefix);
            if (map.has(id)) {
                console.warn("Multiple definitions of section " + id);
            }
            map.set(id, result);
        }
        if (this.stack.length == 1) {
            result += ".";
        }
        return result;
    }
    getReference(reference) {
        let prefix = reference.split(":")[0];
        let map = this.getPrefixMap(prefix);
        let index = map.get(reference);
        if (index === undefined) {
            index = (map.size + 1).toString();
            map.set(reference, index);
        }
        return index.toString();
    }
    getCite(reference) {
        let links = this.meta.getSection("links").asArray();
        for (let i = 0; i < links.length; i++) {
            let link = links[i];
            if (link.isMap() && link.getString("id") === reference) {
                return "[" + (i + 1).toString() + "]";
            }
        }
        console.warn("Undefined citation: " + reference);
        return "[?]";
    }
}

function addHeaderReferences(tree) {
    visit(tree, 'heading', (node) => {
        let lastChild = node.children[node.children.length - 1];
        if (!lastChild || lastChild.type !== "text")
            return;
        lastChild = lastChild;
        const match = lastChild.value.match(/\s*{#(.*)}/);
        if (match) {
            node.data = node.data || {};
            node.data.id = match[1];
            lastChild.value = lastChild.value.replace(/\s*{#.*}/, '').trim();
        }
    });
}
function removePositions(tree) {
    visit(tree, (node) => {
        node.position = undefined;
    });
}
function addReferences(tree) {
    let meta = DocumentJsonMeta.fromMarkdown(tree);
    let references = new DocumentReferences(meta);
    visit(tree, "heading", (node) => {
        let reference = references.getSection(node.depth, node.data?.id);
        node.children.unshift({
            type: "text",
            value: reference + " "
        });
    });
}
function parseAttributes(attributeString) {
    if (!attributeString.startsWith("{") || !attributeString.endsWith("}"))
        return;
    attributeString = attributeString.slice(1, -1);
    let regex = /(\w[\w0-9_-]*)=(?:(?:(?<quote>["'])((?:\\.|[^\\])*?)\k<quote>)|((?:[^'"]\S*)?))/g;
    let match;
    let result = Object.create(null);
    while ((match = regex.exec(attributeString)) !== null) {
        let key = match[1];
        let value;
        if (match[4]) {
            try {
                value = JSON.parse('"' + match[4] + '"');
            }
            catch (e) {
                value = match[4];
            }
        }
        else {
            value = match[3];
        }
        result[key] = value;
    }
    return result;
}
function parseImageAttrs(tree) {
    visit(tree, (node) => {
        if ('children' in node) {
            for (let i = 1; i < node.children.length; i++) {
                let previous = node.children[i - 1];
                let current = node.children[i];
                if (previous.type !== "image" || current.type !== "text") {
                    continue;
                }
                let text = current;
                let image = previous;
                let attrs = parseAttributes(text.value);
                if (attrs) {
                    image.data = image.data || {};
                    image.data = { attrs: attrs };
                    node.children.splice(i, 1);
                }
            }
        }
    });
}
function parseMarkdown(src) {
    let tree = fromMarkdown(src, {
        extensions: [math(), gfmTable(), frontmatter(['yaml'])],
        mdastExtensions: [mathFromMarkdown(), gfmTableFromMarkdown(), frontmatterFromMarkdown(['yaml'])]
    });
    addHeaderReferences(tree);
    removePositions(tree);
    addReferences(tree);
    parseImageAttrs(tree);
    return tree;
}

class ParagraphTemplateSubstitution {
    document;
    template;
    replacement;
    setDocument(document) {
        this.document = document;
        return this;
    }
    setTemplate(template) {
        this.template = template;
        return this;
    }
    setReplacement(replacement) {
        this.replacement = replacement;
        return this;
    }
    perform() {
        const body = this.document.document.resource.toXml().getChild("w:document").getChild("w:body");
        this.replaceParagraphsWithTemplate(body);
        return this;
    }
    replaceParagraphsWithTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            let paragraphText = "";
            child.visitSubtree("w:t", (textNode) => {
                paragraphText += getTextContents(textNode);
            });
            if (paragraphText.indexOf(this.template) === -1) {
                continue;
            }
            if (paragraphText !== this.template) {
                throw new Error(`The ${this.template} pattern should be the only text of the paragraph`);
            }
            body.removeChild([i]);
            let replacement = this.replacement();
            body.insertChildren(replacement, [i]);
            i += replacement.length - 1;
        }
    }
}

class InlineTemplateSubstitution {
    document;
    template;
    replacement;
    setDocument(document) {
        this.document = document;
        return this;
    }
    setTemplate(template) {
        this.template = template;
        return this;
    }
    setReplacement(replacement) {
        this.replacement = replacement;
        return this;
    }
    replaceInlineTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(keys.text, (textNode) => {
                    textNode.setText(textNode.getText().replace(this.template, this.replacement));
                });
            });
        }
    }
    removeParagraphsWithTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            let found = false;
            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(keys.text, (textNode) => {
                    let text = textNode.getText();
                    if (text.indexOf(this.template) !== -1) {
                        found = true;
                    }
                });
                return !found;
            });
            if (found) {
                body.removeChild([i]);
                i--;
            }
        }
    }
    performIn(body) {
        if (this.replacement === "@none") {
            this.removeParagraphsWithTemplate(body);
        }
        else {
            this.replaceInlineTemplate(body);
        }
    }
    perform() {
        let document = this.document;
        let documentBody = document.document.resource.toXml().getChild("w:document").getChild("w:body");
        this.performIn(documentBody);
        for (let header of document.headers) {
            this.performIn(header.resource.toXml().getChild("w:hdr"));
        }
        for (let footer of document.footers) {
            this.performIn(footer.resource.toXml().getChild("w:ftr"));
        }
        return this;
    }
}

let DecimalNumberingStyle = { styleId: "ispNumList", numId: "33" };
let BulletNumberingStyle = { styleId: "ispList1", numId: "43" };
class GenerationContext {
    doc;
    meta;
    nodeStack = [];
    numberingStack = [];
    visit(element) {
        visitors[element.type]?.(element, this);
    }
    visitChildren(element) {
        for (let child of element.children) {
            this.visit(child);
        }
    }
    getCurrentNode() {
        return this.nodeStack[this.nodeStack.length - 1];
    }
    pushNode(node) {
        this.nodeStack.push(node);
    }
    popNode() {
        this.nodeStack.pop();
    }
    pushNumberingStyle(style) {
        this.numberingStack.push(style);
    }
    popNumberingStyle() {
        this.numberingStack.pop();
    }
    getNumberingStyle() {
        return this.numberingStack[this.numberingStack.length - 1];
    }
}
const visitors = {
    "paragraph": (node, ctx) => {
        let numberingStyle = ctx.getNumberingStyle();
        let styleId = numberingStyle ? numberingStyle.styleId : ctx.doc.styles.resource.getStyleByName("ispText_main").getId();
        let paragraph = buildParagraphWithStyle(styleId);
        if (numberingStyle) {
            paragraph.getChild("w:pPr").pushChild(Node.build("w:numPr").appendChildren([
                Node.build("w:ilvl").setAttr("w:val", String(ctx.numberingStack.length - 1)),
                Node.build("w:numId").setAttr("w:val", numberingStyle.numId),
            ]));
        }
        ctx.getCurrentNode().pushChild(paragraph);
        ctx.pushNode(paragraph);
        ctx.visitChildren(node);
        ctx.popNode();
    },
    "strong": (node, ctx) => {
    },
    "code": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("ispListing").getId();
        let raw = Node.build("w:r");
        let lines = node.value.split("\n");
        for (let line of lines) {
            if (raw.getChildrenCount() > 0) {
                raw.pushChild(buildLineBreak());
            }
            raw.pushChild(buildTextTag(line));
        }
        ctx.getCurrentNode().pushChild(buildParagraphWithStyle(styleId)
            .pushChild(raw));
    },
    "link": (node, ctx) => {
    },
    "delete": (node, ctx) => {
    },
    "emphasis": (node, ctx) => {
    },
    "text": (node, ctx) => {
        ctx.getCurrentNode().pushChild(buildRawTag(node.value));
    },
    "table": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("Table Grid").getId();
        let table = Node.build("w:tbl");
        let tableProperties = Node.build("w:tPr").appendChildren([
            Node.build("w:tblStyle").setAttr("w:val", styleId)
        ]);
        table.pushChild(tableProperties);
        ctx.getCurrentNode().pushChild(table);
        ctx.pushNode(table);
        ctx.visitChildren(node);
        ctx.popNode();
    },
    "tableRow": (node, ctx) => {
        let row = Node.build("w:tr").appendChildren([]);
        ctx.getCurrentNode().pushChild(row);
        ctx.pushNode(row);
        ctx.visitChildren(node);
        ctx.popNode();
    },
    "tableCell": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("ispText_main").getId();
        let paragraph = buildParagraphWithStyle(styleId);
        let cell = Node.build("w:tc").appendChildren([
            paragraph
        ]);
        ctx.getCurrentNode().pushChild(cell);
        ctx.pushNode(paragraph);
        ctx.visitChildren(node);
        ctx.popNode();
    },
    "inlineMath": ((node, ctx) => {
        let value = temml.renderToString(node.value);
        let omml = mml2omml(value);
        ctx.getCurrentNode().pushChild(Node.fromXmlString(omml));
    }),
    "listItem": (node, ctx) => {
        ctx.visitChildren(node);
    },
    "image": (node, ctx) => {
        let imageBuffer = fs__default.readFileSync(node.url);
        let file = ImageSize(imageBuffer);
        let size = getImageSize(node, file.width / file.height);
        // LibreOffice seems to break when filenames contain bad symbols.
        let mediaName = path__default.basename(node.url).replace(/[^a-zA-Z0-9.]/g, "_");
        let mediaTarget = ctx.doc.getUniqueMediaTarget(mediaName);
        let mediaPath = ctx.doc.getPathForTarget(mediaTarget);
        ctx.doc.saveFile(mediaPath, imageBuffer);
        let unusedId = ctx.doc.document.rels.getUnusedId();
        ctx.doc.document.rels.addRelation({
            id: unusedId,
            target: mediaTarget,
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        });
        let image = createImage(unusedId, size[0], size[1]);
        ctx.getCurrentNode().pushChild(Node.build("w:r").pushChild(image));
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
        ];
        let level = node.depth;
        if (level > headings.length) {
            level = headings.length;
        }
        let styleId = ctx.doc.styles.resource.getStyleByName(headings[level - 1]).getId();
        let paragraph = buildParagraphWithStyle(styleId);
        ctx.getCurrentNode().pushChild(paragraph);
        ctx.pushNode(paragraph);
        ctx.visitChildren(node);
        ctx.popNode();
    },
    "imageReference": (node, ctx) => {
    },
    "list": (node, ctx) => {
        let numberingStyle = node.ordered ? DecimalNumberingStyle : BulletNumberingStyle;
        let numId = createNumbering(ctx.doc, numberingStyle.numId, ctx.numberingStack.length, node.start);
        ctx.pushNumberingStyle({
            numId: numId,
            styleId: numberingStyle.styleId
        });
        ctx.visitChildren(node);
        ctx.popNumberingStyle();
    },
    "math": (node, ctx) => {
    },
    "inlineCode": (node, ctx) => {
        let styleId = ctx.doc.styles.resource.getStyleByName("ispListing Знак").getId();
        let raw = buildRawTag(node.value, [
            Node.build("w:rStyle").setAttr("w:val", styleId)
        ]);
        ctx.getCurrentNode().pushChild(raw);
    },
};
function parseSizeAttr(attr) {
    if (typeof attr !== "string")
        return null;
    if (attr.endsWith("cm")) {
        let centimeters = parseFloat(attr.slice(0, -2));
        if (Number.isFinite(centimeters))
            return centimeters;
    }
    return null;
}
function getImageSize(image, aspect) {
    let width = parseSizeAttr(image.data?.attrs?.width);
    let height = parseSizeAttr(image.data?.attrs?.height);
    if (width === null && height === null) {
        width = 10;
    }
    if (height === null) {
        height = width / aspect;
    }
    if (width === null) {
        width = aspect * height;
    }
    return [width, height];
}
function createNumbering(document, abstractId, depth, start) {
    let newNumId = document.numbering.resource.getUnusedNumId();
    let newNumbering = new Num().readXml(Node.build("w:num"));
    newNumbering.setAbstractNumId(abstractId);
    newNumbering.setId(newNumId);
    newNumbering.getLevelOverride(depth).pushChild(Node.build("w:startOverride")
        .setAttr("w:val", String(start)));
    document.numbering.resource.nums.set(newNumId, newNumbering);
    return newNumId;
}
function createImage(id, width, height) {
    let widthEmu = Math.floor(width * 360000);
    let heightEmu = Math.floor(height * 360000);
    return Node.build("w:drawing").appendChildren([
        Node.build("wp:inline").appendChildren([
            Node.build("a:graphic").appendChildren([
                Node.build("a:graphicData").setAttrs({
                    "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                }).appendChildren([
                    Node.build("pic:pic").appendChildren([
                        Node.build("pic:nvPicPr").appendChildren([
                            Node.build("pic:cNvPr").setAttrs({
                                "id": "pic" + id,
                                "name": "Picture"
                            }),
                            Node.build("pic:cNvPicPr")
                        ]),
                        Node.build("pic:blipFill").appendChildren([
                            Node.build("a:blip").setAttr("r:embed", id),
                            Node.build("a:stretch").appendChildren([
                                Node.build("a:fillRect")
                            ])
                        ]),
                        Node.build("pic:spPr").appendChildren([
                            Node.build("a:xfrm").appendChildren([
                                Node.build("a:off").setAttrs({
                                    "x": "0",
                                    "y": "0"
                                }),
                                Node.build("a:ext").setAttrs({
                                    "cx": String(widthEmu),
                                    "cy": String(heightEmu)
                                })
                            ]),
                            Node.build("a:prstGeom").setAttr("prst", "rect").appendChildren([
                                Node.build("a:avLst")
                            ])
                        ])
                    ])
                ])
            ])
        ])
    ]);
}
function generateDocxBody(source, target, meta) {
    new ParagraphTemplateSubstitution()
        .setDocument(target)
        .setTemplate("{{{body}}}")
        .setReplacement(() => {
        let context = new GenerationContext();
        // Fictive node
        let node = Node.build("w:body");
        context.doc = target;
        context.pushNode(node);
        context.visitChildren(source);
        return node.getChildren();
    })
        .perform();
}
function getLinksParagraphs(doc, meta) {
    let styleId = doc.styles.resource.getStyleByName("ispLitList").getId();
    let numId = "80";
    let linksSection = meta.getSection("links").asArray();
    let result = [];
    for (let link of linksSection) {
        let description;
        // Backwards compatibility
        if (link.isMap()) {
            description = link.getString("description");
        }
        else {
            description = link.getString();
        }
        let paragraph = buildParagraphWithStyle(styleId);
        let style = paragraph.getChild("w:pPr");
        style.pushChild(buildNumPr("0", numId));
        paragraph.pushChild(buildRawTag(description));
        result.push(paragraph);
    }
    return result;
}
function getAuthors(doc, meta, language) {
    let styleId = doc.styles.resource.getStyleByName("ispAuthor").getId();
    let authors = meta.getSection("authors").asArray();
    let organizations = meta
        .getSection("organizations")
        .asArray()
        .map(section => section.getString("id"));
    let result = [];
    for (let author of authors) {
        let paragraph = buildParagraphWithStyle(styleId);
        let name = author.getString("name_" + language);
        let orcid = author.getString("orcid");
        let email = author.getString("email");
        let authorOrgs = author.getSection("organizations")
            .asArray()
            .map(section => section.getString())
            .map(id => organizations.indexOf(id) + 1)
            .join(",");
        let authorLine = `${name}, ORCID: ${orcid}, <${email}>`;
        paragraph.pushChild(buildRawTag(authorOrgs, [buildSuperscriptTextStyle()]));
        paragraph.pushChild(buildRawTag(authorLine));
        result.push(paragraph);
    }
    return result;
}
function getOrganizations(doc, meta, language) {
    let styleId = doc.styles.resource.getStyleByName("ispAuthor").getId();
    let organizations = meta.getSection("organizations").asArray();
    let orgIndex = 1;
    let result = [];
    for (let organization of organizations) {
        let paragraph = buildParagraphWithStyle(styleId);
        let indexLine = String(orgIndex);
        paragraph.pushChild(buildRawTag(indexLine, [buildSuperscriptTextStyle()]));
        paragraph.pushChild(buildRawTag(organization.getString("name_" + language)));
        result.push(paragraph);
        orgIndex++;
    }
    return result;
}
function getAuthorsDetail(doc, meta) {
    let styleId = doc.styles.resource.getStyleByName("ispText_main").getId();
    let authors = meta.getSection("authors").asArray();
    let result = [];
    for (let author of authors) {
        for (let language of languages) {
            let line = author.getString("details_" + language);
            let newParagraph = buildParagraphWithStyle(styleId);
            newParagraph.getChild("w:pPr").pushChild(Node.build("w:spacing")
                .setAttr("w:before", "30")
                .setAttr("w:after", "120"));
            newParagraph.pushChild(buildRawTag(line));
            result.push(newParagraph);
        }
    }
    return result;
}
function substituteTemplates(document, meta) {
    let inlineSubstitution = new InlineTemplateSubstitution().setDocument(document);
    let paragraphSubstitution = new ParagraphTemplateSubstitution().setDocument(document);
    for (let language of languages) {
        let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"];
        for (let template of templates) {
            let template_lang = template + "_" + language;
            let replacement = meta.getString(template_lang);
            inlineSubstitution
                .setTemplate("{{{" + template_lang + "}}}")
                .setReplacement(replacement)
                .perform();
        }
        let header = meta.getString("page_header_" + language);
        if (header === "@use_citation") {
            header = meta.getString("for_citation_" + language);
        }
        inlineSubstitution
            .setTemplate("{{{page_header_" + language + "}}}")
            .setReplacement(header)
            .perform();
        paragraphSubstitution
            .setTemplate("{{{authors_" + language + "}}}")
            .setReplacement(() => getAuthors(document, meta, language))
            .perform();
        paragraphSubstitution
            .setTemplate("{{{organizations_" + language + "}}}")
            .setReplacement(() => getOrganizations(document, meta, language))
            .perform();
    }
    paragraphSubstitution
        .setTemplate("{{{links}}}")
        .setReplacement(() => getLinksParagraphs(document, meta))
        .perform();
    paragraphSubstitution
        .setTemplate("{{{authors_detail}}}")
        .setReplacement(() => getAuthorsDetail(document, meta))
        .perform();
}

const languages = ["ru", "en"];
const resourcesDir = path.dirname(process.argv[1]) + "/../resources";
async function main() {
    let argv = process.argv;
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>");
        process.exit(1);
    }
    let markdownSource = argv[2];
    let targetPath = argv[3];
    let markdown = await fs.promises.readFile(markdownSource, "utf-8");
    let markdownParsed = parseMarkdown(markdown);
    await fs.promises.writeFile(markdownSource + ".json", JSON.stringify(markdownParsed, null, 4), "utf-8");
    let documentMeta = DocumentJsonMeta.fromMarkdown(markdownParsed).getSection("ispras_templates");
    let templateDoc = await new WordDocument().load(resourcesDir + '/isp-reference.docx');
    await generateDocxBody(markdownParsed, templateDoc);
    substituteTemplates(templateDoc, documentMeta);
    await templateDoc.save(targetPath);
}
main().then();

export { languages };
//# sourceMappingURL=main.js.map
