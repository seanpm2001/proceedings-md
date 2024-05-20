import * as path from 'path';
import * as fs from 'fs';
import * as process from "process";
import WordDocument from "src/word/word-document";
import {DocumentJsonMeta} from "src/document-json-meta";
import {parseMarkdown} from "src/markdown/markdown";
import {generateDocxBody, substituteTemplates} from "src/generator/generator";
export const languages = ["ru", "en"]
const resourcesDir = path.dirname(process.argv[1]) + "/../resources"

async function main(): Promise<void> {
    let argv = process.argv
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>")
        process.exit(1)
    }

    let markdownSource = argv[2]
    let targetPath = argv[3]

    let markdown = await fs.promises.readFile(markdownSource, "utf-8")
    let markdownParsed = parseMarkdown(markdown)

    await fs.promises.writeFile(markdownSource + ".json", JSON.stringify(markdownParsed, null, 4), "utf-8")

    let documentMeta = DocumentJsonMeta.fromMarkdown(markdownParsed).getSection("ispras_templates")
    let templateDoc = await new WordDocument().load(resourcesDir + '/isp-reference.docx')

    await generateDocxBody(markdownParsed, templateDoc, documentMeta);
    substituteTemplates(templateDoc, documentMeta)

    await templateDoc.save(targetPath)
}

main().then()