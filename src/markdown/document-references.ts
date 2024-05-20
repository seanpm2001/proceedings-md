import {DocumentJsonMeta} from "src/document-json-meta";

export default class DocumentReferences {
    stack: number[] = []
    depthThreshold: number = 1
    groups = new Map<string, Map<string, string>>()
    meta: DocumentJsonMeta

    constructor(meta: DocumentJsonMeta) {
        this.meta = meta
    }

    getPrefixMap(prefix: string) {
        let map = this.groups.get(prefix)
        if (!map) {
            map = new Map()
            this.groups.set(prefix, map)
        }
        return map
    }

    getSection(depth: number, id?: string): string {
        depth = Math.max(0, depth - this.depthThreshold)

        while (this.stack.length > depth) this.stack.pop()
        if (this.stack.length === depth) {
            this.stack[this.stack.length - 1]++
        } else {
            while (this.stack.length < depth) this.stack.push(1)
        }

        if (this.stack.length === 0) {
            return
        }

        let result = this.stack.join(".")

        if(id) {
            let prefix = id.split(":")[0]
            let map = this.getPrefixMap(prefix)

            if (map.has(id)) {
                console.warn("Multiple definitions of section " + id)
            }

            map.set(id, result)
        }

        if (this.stack.length == 1) {
            result += "."
        }

        return result
    }

    getReference(reference: string): string {
        let prefix = reference.split(":")[0]
        let map = this.getPrefixMap(prefix)

        let index = map.get(reference)
        if (index === undefined) {
            index = (map.size + 1).toString()
            map.set(reference, index)
        }

        return index.toString()
    }

    getCite(reference: string): string {
        let links = this.meta.getSection("links").asArray()
        for (let i = 0; i < links.length; i++) {
            let link = links[i]
            if (link.isMap() && link.getString("id") === reference) {
                return "[" + (i + 1).toString() + "]"
            }
        }

        console.warn("Undefined citation: " + reference)

        return "[?]"
    }
}