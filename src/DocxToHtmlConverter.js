import JSZip from "jszip";
import { XMLParser as DefaultXMLParser } from "fast-xml-parser";

export default class DocxToHtmlConverter {
    /**
     * @param {any} zip Instância do JSZip já carregada (usada no create)
     * @param {Function} ParserClass Classe do parser opcional para override
     */
    constructor(zip, ParserClass) {
        this.zip = zip;

        // estado
        this.relationships = null;
        this.numbering = null;
        this.globalStyles = {};
        this.docDefaults = { pPr: [], rPr: [] };
        this.listState = { stack: [] };
        this.listCounters = {};

        const Parser = ParserClass || DefaultXMLParser;

        this.parser = new Parser({
            preserveOrder: true,
            ignoreDeclaration: true,
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            textNodeName: "##text",
            trimValues: false,
        });
    }

    static async create(arrayBuffer, ParserClass) {
        if (!arrayBuffer) throw new Error("O arquivo está vazio ou corrompido.");
        const zip = await JSZip.loadAsync(arrayBuffer);
        const converter = new DocxToHtmlConverter(zip, ParserClass);
        await converter.loadRelationships();
        await converter.loadNumbering();
        await converter.loadStyles();
        return converter;
    }

    // Map DOCX numFmt → CSS list-style-type
    static NUMFMT_TO_CSS = {
        decimal: 'decimal',
        decimalZero: 'decimal',
        lowerRoman: 'lower-roman',
        upperRoman: 'upper-roman',
        lowerLetter: 'lower-alpha',
        upperLetter: 'upper-alpha',
        bullet: 'disc'
    };

    mergeProperties(baseProps = [], derivedProps = []) {
        const propsMap = new Map();

        baseProps.forEach(prop => {
            const key = Object.keys(prop)[0];
            if (key !== ':@') {
                propsMap.set(key, prop);
            }
        });

        derivedProps.forEach(prop => {
            const key = Object.keys(prop)[0];
            if (key !== ':@') {
                propsMap.set(key, prop);
            }
        });

        return Array.from(propsMap.values());
    }

    mergeStyleObjects(baseStyle = {}, derivedStyle = {}) {
        return {
            pPr: this.mergeProperties(baseStyle.pPr, derivedStyle.pPr),
            rPr: this.mergeProperties(baseStyle.rPr, derivedStyle.rPr),
            tblPr: this.mergeProperties(baseStyle.tblPr, derivedStyle.tblPr),
        };
    }

    async loadStyles() {
        const stylesFile = this.zip.file("word/styles.xml");
        if (!stylesFile) return;

        const xmlContent = await stylesFile.async("string");
        const jsonObj = this.parser.parse(xmlContent);
        const stylesNode = this.findChild(jsonObj, "w:styles");
        if (!stylesNode) return;

        const docDefaultsNode = this.findChild(stylesNode["w:styles"], "w:docDefaults");
        if (docDefaultsNode) {
            const rPrDefaultNode = this.findChild(docDefaultsNode["w:docDefaults"], "w:rPrDefault");
            if (rPrDefaultNode) {
                const rPr = this.findChild(rPrDefaultNode["w:rPrDefault"], "w:rPr");
                if (rPr) this.docDefaults.rPr = rPr["w:rPr"];
            }
            const pPrDefaultNode = this.findChild(docDefaultsNode["w:docDefaults"], "w:pPrDefault");
            if (pPrDefaultNode) {
                const pPr = this.findChild(pPrDefaultNode["w:pPrDefault"], "w:pPr");
                if (pPr) this.docDefaults.pPr = pPr["w:pPr"];
            }
        }

        const styleNodes = this.filterChildren(stylesNode["w:styles"], "w:style");
        const rawStyles = {};

        styleNodes.forEach(styleNode => {
            const attrs = styleNode[":@"];
            if (!attrs || !attrs["@_w:styleId"]) return;

            const styleId = attrs["@_w:styleId"];
            const styleChildren = styleNode["w:style"];

            const pPrNode = this.findChild(styleChildren, "w:pPr");
            const rPrNode = this.findChild(styleChildren, "w:rPr");
            const tblPrNode = this.findChild(styleChildren, "w:tblPr");
            const basedOnNode = this.findChild(styleChildren, "w:basedOn");
            const linkNode = this.findChild(styleChildren, "w:link");

            const basedOnId = basedOnNode ? basedOnNode[':@']['@_w:val'] : null;
            const linkId = linkNode ? linkNode[':@']['@_w:val'] : null;

            rawStyles[styleId] = {
                basedOn: basedOnId,
                link: linkId,
                pPr: pPrNode ? pPrNode["w:pPr"] : [],
                rPr: rPrNode ? rPrNode["w:rPr"] : [],
                tblPr: tblPrNode ? tblPrNode["w:tblPr"] : [],
            };
        });

        const styleIds = Object.keys(rawStyles);
        for (const styleId of styleIds) {
            this.resolveStyle(styleId, rawStyles);
        }
    }

    resolveStyle(styleId, rawStyles) {
        if (this.globalStyles[styleId]) {
            return this.globalStyles[styleId];
        }

        const styleData = rawStyles[styleId];
        if (!styleData) {
            return { pPr: [], rPr: [], tblPr: [] };
        }

        let baseStyle = { pPr: [], rPr: [], tblPr: [] };
        if (styleData.basedOn) {
            baseStyle = this.resolveStyle(styleData.basedOn, rawStyles);
        }

        const resolvedStyle = this.mergeStyleObjects(baseStyle, styleData);
        this.globalStyles[styleId] = resolvedStyle;
        return resolvedStyle;
    }

    extractFileNameFromPath(path) {
        const parts = path.split(/[\\/]/);
        return parts[parts.length - 1];
    }

    extractAnchorFileName(anchorText) {
        let fileName = anchorText;
        try {
            const firstPart = anchorText.split(',')[0];
            const cleaned = firstPart.replace(/[^a-zA-Z0-9._-]/g, '');
            fileName = cleaned;
        } catch (e) { }
        return fileName;
    }

    async convert(options = { extractPageStyles: true }) {
        const file = this.zip.file("word/document.xml");
        if (!file) throw new Error("document.xml não encontrado.");
        const xmlContent = await file.async("string");
        const jsonObj = this.parser.parse(xmlContent);
        const documentNode = jsonObj.find(node => node["w:document"]);
        if (!documentNode) throw new Error("Estrutura inesperada: w:document não encontrado.");
        const bodyNode = this.findChild(documentNode["w:document"], "w:body");
        if (!bodyNode) throw new Error("w:body não encontrado.");

        const bodyChildren = bodyNode["w:body"];

        let pageStyles = null;
        let pageStylesCss = null;
        if (options.extractPageStyles) {
            const lastSectPrNode = this.findChild(bodyChildren, "w:sectPr");
            if (lastSectPrNode) {
                pageStyles = this.getPageStylesAsObject(lastSectPrNode["w:sectPr"]);
                pageStylesCss = DocxToHtmlConverter.formatCssFromStyles(pageStyles);
            }
        }

        const sections = [];
        let currentSectionChildren = [];
        const lastSectPrNode = this.findChild(bodyChildren, "w:sectPr");
        const lastSectPr = lastSectPrNode ? lastSectPrNode["w:sectPr"] : null;

        for (const child of bodyChildren) {
            const nodeName = Object.keys(child)[0];
            if (nodeName === 'w:p') {
                const pPrNode = this.findChild(child['w:p'], 'w:pPr');
                if (pPrNode) {
                    const sectPrNode = this.findChild(pPrNode['w:pPr'], 'w:sectPr');
                    if (sectPrNode) {
                        sections.push({ children: currentSectionChildren, sectPr: sectPrNode['w:sectPr'] });
                        currentSectionChildren = [];
                        continue;
                    }
                }
            }
            if (nodeName === 'w:sectPr') continue;
            currentSectionChildren.push(child);
        }

        if (currentSectionChildren.length > 0 || sections.length === 0) {
            sections.push({ children: currentSectionChildren, sectPr: lastSectPr });
        }

        let htmlContent = "";
        for (const section of sections) {
            const sectionHtml = await this.processChildren(section.children);
            const sectionStyle = this.getSectionStyles(section.sectPr);
            const styleAttr = sectionStyle ? ` style="${sectionStyle}"` : "";
            if (sectionHtml.trim()) {
                htmlContent += `<div class="docx-section"${styleAttr}>${sectionHtml}</div>`;
            }
        }

        htmlContent = htmlContent.replace(/<(h[1-6]|p)>\s*<\/\1>/g, "<$1> </$1>");

        return {
            html: `<div class="docx">${htmlContent}</div>`,
            pageStyles: pageStyles,
            pageStylesCss: pageStylesCss
        };
    }

    getPageStylesAsObject(sectPr) {
        if (!sectPr) return null;

        const styles = {
            size: { width: null, height: null, orientation: 'portrait' },
            margin: { top: null, right: null, bottom: null, left: null },
            units: 'pt'
        };

        const pgSzNode = this.findChild(sectPr, "w:pgSz");
        if (pgSzNode && pgSzNode[":@"]) {
            const attrs = pgSzNode[":@"];
            const width = parseInt(attrs["@_w:w"], 10) / 20;
            const height = parseInt(attrs["@_w:h"], 10) / 20;
            if (!isNaN(width)) styles.size.width = width;
            if (!isNaN(height)) styles.size.height = height;
            if (attrs["@_w:orient"] === "landscape") {
                styles.size.orientation = "landscape";
            }
        }

        const pgMarNode = this.findChild(sectPr, "w:pgMar");
        if (pgMarNode && pgMarNode[":@"]) {
            const attrs = pgMarNode[":@"];
            const sides = ["top", "right", "bottom", "left"];
            sides.forEach(side => {
                if (attrs[`@_w:${side}`]) {
                    const marginValue = parseInt(attrs[`@_w:${side}`], 10) / 20;
                    if (!isNaN(marginValue)) styles.margin[side] = marginValue;
                }
            });
        }
        return styles;
    }

    static formatCssFromStyles(styles) {
        if (!styles) return null;

        let cssRules = [];
        const { size, margin, units } = styles;

        if (size.orientation === "landscape") {
            cssRules.push(`size: landscape;`);
        } else if (size.width && size.height) {
            cssRules.push(`size: ${size.width}${units} ${size.height}${units};`);
        }

        Object.keys(margin).forEach(side => {
            if (margin[side] !== null) {
                cssRules.push(`margin-${side}: ${margin[side]}${units};`);
            }
        });

        if (cssRules.length > 0) {
            return `@page { ${cssRules.join(" ")} }`;
        }
        return null;
    }

    getSectionStyles(sectPr) {
        if (!sectPr) return "";

        let style = "";
        const colsNode = this.findChild(sectPr, "w:cols");
        if (colsNode && colsNode[":@"]) {
            const num = parseInt(colsNode[":@"]["@_w:num"], 10);
            if (!isNaN(num) && num > 1) {
                style += `column-count: ${num};`;
                const space = colsNode[":@"]["@_w:space"];
                if (space) {
                    const spaceInPt = parseInt(space, 10) / 20;
                    style += ` column-gap: ${spaceInPt}pt;`;
                }
            }
        }
        return style;
    }

    findChild(nodeArray, tagName) {
        if (!nodeArray || !Array.isArray(nodeArray)) return null;
        return nodeArray.find(child => Object.keys(child)[0] === tagName);
    }

    filterChildren(nodeArray, tagName) {
        if (!nodeArray || !Array.isArray(nodeArray)) return [];
        return nodeArray.filter(child => Object.keys(child)[0] === tagName);
    }

    getText(node) {
        if (node && node["##text"]) {
            return Array.isArray(node["##text"]) ? node["##text"].join('') : node["##text"];
        }
        return "";
    }

    async loadRelationships() {
        const relsFile = this.zip.file("word/_rels/document.xml.rels");
        if (relsFile) {
            const xmlContent = await relsFile.async("string");
            const jsonObj = this.parser.parse(xmlContent);
            const relationshipsRoot = jsonObj.find(node => node.Relationships);
            if (!relationshipsRoot) return;

            const relationshipsChildren = relationshipsRoot.Relationships;
            this.relationships = {};

            const relationshipNodes = this.filterChildren(relationshipsChildren, "Relationship");

            relationshipNodes.forEach(relNode => {
                const attrs = relNode[":@"];
                if (attrs && attrs["@_Id"]) {
                    this.relationships[attrs["@_Id"]] = {
                        type: attrs["@_Type"],
                        target: attrs["@_Target"]
                    };
                }
            });
        }
    }

    async loadNumbering() {
        const numberingFile = this.zip.file("word/numbering.xml");
        if (!numberingFile) return;
        const xmlContent = await numberingFile.async("string");
        const jsonObj = this.parser.parse(xmlContent);
        const root = jsonObj.find(n => n["w:numbering"]);
        if (!root) return;
        const numberingChildren = root["w:numbering"];
        const abstractNums = {};
        this.filterChildren(numberingChildren, "w:abstractNum").forEach(abstractNumNode => {
            const anId = abstractNumNode[":@"]["@_w:abstractNumId"];
            const anChildren = abstractNumNode["w:abstractNum"];
            const lvlNodes = this.filterChildren(anChildren, "w:lvl");
            const levels = {};
            lvlNodes.forEach(lvlNode => {
                const lvlChildren = lvlNode["w:lvl"];
                const ilvl = lvlNode[":@"]["@_w:ilvl"];
                const numFmtNode = this.findChild(lvlChildren, "w:numFmt");
                const lvlTextNode = this.findChild(lvlChildren, "w:lvlText");
                const startNode = this.findChild(lvlChildren, "w:start");

                const numFmt = numFmtNode?.[":@"]?.["@_w:val"] || "decimal";
                const lvlText = lvlTextNode?.[":@"]?.["@_w:val"] || "%1.";
                const start = startNode?.[":@"]?.["@_w:val"] ? parseInt(startNode[":@"]["@_w:val"], 10) : 1;

                levels[ilvl] = { numFmt, lvlText, start };
            });
            abstractNums[anId] = levels;
        });
        this.filterChildren(numberingChildren, "w:num").forEach(numNode => {
            const numId = numNode[":@"]["@_w:numId"];
            const numChildren = numNode["w:num"];
            const abstractNumIdNode = this.findChild(numChildren, "w:abstractNumId");
            if (abstractNumIdNode && abstractNumIdNode[":@"]) {
                const abstractId = abstractNumIdNode[":@"]["@_w:val"];
                const abstractRef = abstractNums[abstractId];
                if (!this.numbering) this.numbering = {};
                this.numbering[numId] = abstractRef;
            }
        });
    }

    async processChildren(elementArray) {
        let html = "";

        for (let i = 0; i < elementArray.length; i++) {
            const node = elementArray[i];
            const nodeName = Object.keys(node)[0];
            if (nodeName.startsWith(":") || nodeName === '##text') continue;

            if (nodeName === 'w:p') {
                const pChildren = node['w:p'];
                const pPrNode = this.findChild(pChildren, "w:pPr");
                const pPr = pPrNode ? pPrNode["w:pPr"] : null;

                let paragraphStyleDef = null;
                if (pPr) {
                    const pStyleNode = this.findChild(pPr, "w:pStyle");
                    if (pStyleNode && pStyleNode[":@"]) {
                        const styleId = pStyleNode[":@"]["@_w:val"];
                        paragraphStyleDef = this.globalStyles[styleId];
                    }
                }

                const content = await this.processParagraphContent(pChildren, paragraphStyleDef);
                const pBdrNode = pPr ? this.findChild(pPr, "w:pBdr") : null;
                const hasBottomBorder = pBdrNode && this.findChild(pBdrNode["w:pBdr"], "w:bottom");
                if (hasBottomBorder && content.trim() === "") {
                    html += this.closeLists() + "<hr>";
                    continue;
                }

                let numId = null;
                let ilvl = null;

                if (pPr) {
                    let numPrNode = this.findChild(pPr, "w:numPr");

                    if (!numPrNode && paragraphStyleDef && paragraphStyleDef.pPr) {
                        numPrNode = this.findChild(paragraphStyleDef.pPr, "w:numPr");
                    }

                    if (numPrNode) {
                        const numPr = numPrNode["w:numPr"];
                        const numIdNode = this.findChild(numPr, "w:numId");
                        const ilvlNode = this.findChild(numPr, "w:ilvl");
                        if (numIdNode && numIdNode[":@"] && ilvlNode && ilvlNode[":@"]) {
                            numId = numIdNode[":@"]["@_w:val"];
                            ilvl = ilvlNode[":@"]["@_w:val"];
                        }
                    }
                }

                if (numId !== null && ilvl !== null) {
                    html += this.handleListItem(numId, parseInt(ilvl, 10), content);
                } else {
                    html += this.closeLists();
                    if (content.trim() === '' && !content.includes(' ')) {
                        html += '<p> </p>';
                    } else {
                        html += await this.renderNonListParagraph(pChildren, content);
                    }
                }
            } else if (nodeName === 'w:tbl') {
                html += this.closeLists();
                html += await this.processTable(node['w:tbl']);
            } else if (nodeName === 'w:sectPr') { }
        }
        html += this.closeLists();
        return html;
    }

    async processParagraphContent(pChildren, paragraphStyleDef) {
        let content = "";
        for (const childNode of pChildren) {
            const tagName = Object.keys(childNode)[0];
            if (tagName === "w:r") content += await this.processRun(childNode["w:r"], paragraphStyleDef);
            else if (tagName === "w:hyperlink") content += await this.processHyperlink(childNode);
        }
        return content;
    }

    getListMeta(numId, level) {
        const lvl = this.numbering?.[numId]?.[level];
        if (!lvl) return { tag: 'ol', css: 'decimal', start: 1, lvlText: '%1.' };
        if (lvl.numFmt === 'bullet') return { tag: 'ul', css: 'disc', start: lvl.start || 1, lvlText: lvl.lvlText };

        return {
            tag: 'ol',
            css: DocxToHtmlConverter.NUMFMT_TO_CSS[lvl.numFmt] || 'decimal',
            start: lvl.start || 1,
            lvlText: lvl.lvlText
        };
    }

    /**
     * Gera o marcador textual da lista a partir do lvlText (%1, %2 ...) e dos contadores atuais.
     */
    formatListMarker(numId, level) {
        const meta = this.getListMeta(numId, level);
        let tpl = meta.lvlText || '%1.';

        if (!this.listCounters[numId]) this.listCounters[numId] = [];

        // Substitui %1, %2... pelos contadores (convertidos conforme numFmt de cada nível)
        tpl = tpl.replace(/%(\d+)/g, (_, n) => {
            const idx = parseInt(n, 10) - 1;
            const counterVal = this.listCounters[numId][idx] || 1;
            const fmt = this.numbering?.[numId]?.[idx]?.numFmt || 'decimal';
            return this.formatCounter(counterVal, fmt);
        });

        return tpl;
    }

    /**
     * Converte um número para o formato exigido (roman, letter, decimal...).
     */
    formatCounter(value, fmt) {
        switch (fmt) {
            case 'lowerRoman': return this.toRoman(value).toLowerCase();
            case 'upperRoman': return this.toRoman(value).toUpperCase();
            case 'lowerLetter': return this.toAlpha(value).toLowerCase();
            case 'upperLetter': return this.toAlpha(value).toUpperCase();
            default: return String(value);
        }
    }

    toRoman(num) {
        const romans = [
            ['M', 1000], ['CM', 900], ['D', 500], ['CD', 400],
            ['C', 100], ['XC', 90], ['L', 50], ['XL', 40],
            ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]
        ];
        let res = '';
        for (const [r, v] of romans) {
            while (num >= v) { res += r; num -= v; }
        }
        return res;
    }

    toAlpha(num) {
        let s = '';
        while (num > 0) {
            num--;
            s = String.fromCharCode(65 + (num % 26)) + s;
            num = Math.floor(num / 26);
        }
        return s;
    }

    closeLists() {
        let html = "";
        while (this.listState.stack.length) {
            const top = this.listState.stack.pop();
            if (top.openLi) html += "</li>";
            html += `</${top.type}>`;
        }
        return html;
    }

    isDefaultMarker(meta) {
        // usa o contador do navegador se o template é só "%1." ou "%1)"
        // e não há referência a níveis superiores (%2, %3...)
        if (!meta || !meta.lvlText) return true;
        const tpl = meta.lvlText.trim();
        const onlyFirst = /^%1[.)]?$/.test(tpl);
        const hasHigher = /%[2-9]/.test(tpl);
        return onlyFirst && !hasHigher;
    }

    handleListItem(numId, level, content) {
        let html = "";

        const currentListDef = this.numbering?.[numId];

        while (this.listState.stack.length > 0 && this.listState.stack[this.listState.stack.length - 1].level > level) {
            const top = this.listState.stack.pop();
            if (top.openLi) html += "</li>";
            html += `</${top.type}>`;
        }

        const stackTop = this.listState.stack.length > 0 ? this.listState.stack[this.listState.stack.length - 1] : null;
        const stackListDef = stackTop ? this.numbering?.[stackTop.numId] : null;

        if (stackTop && stackTop.level === level && currentListDef !== stackListDef) {
            const top = this.listState.stack.pop();
            if (top.openLi) html += "</li>";
            html += `</${top.type}>`;
        }

        while (this.listState.stack.length <= level) {
            const newLevel = this.listState.stack.length;
            const metaLvl = this.getListMeta(numId, newLevel);

            if (!this.listCounters[numId]) this.listCounters[numId] = [];
            if (typeof this.listCounters[numId][newLevel] !== 'number') {
                this.listCounters[numId][newLevel] = metaLvl.start || 1;
            }

            const defaultMarker = this.isDefaultMarker(metaLvl);
            let startAttr = "";
            if (metaLvl.tag === "ol" && metaLvl.start > 1) startAttr = ` start="${metaLvl.start}"`;

            let styleAttr = "";
            if (defaultMarker) {
                // deixa o browser numerar
                styleAttr = metaLvl.css ? ` style="list-style-type:${metaLvl.css};"` : "";
            } else {
                // vamos imprimir o marcador manual → remove numeração do browser
                styleAttr = ` style="list-style-type:none; padding-left:1.5em;"`;
            }

            html += `<${metaLvl.tag}${startAttr}${styleAttr}>`;

            this.listState.stack.push({ numId, level: newLevel, type: metaLvl.tag, openLi: false });
        }

        const container = this.listState.stack[level];
        if (container.openLi) {
            html += "</li>";
        }

        const meta = this.getListMeta(numId, level);
        const defaultMarker = this.isDefaultMarker(meta);

        if (defaultMarker) {
            html += `<li>${content}`;
        } else {
            const marker = this.formatListMarker(numId, level);
            html += `<li><span class="docx-marker">${marker}</span> ${content}`;
        }
        container.openLi = true;

        if (meta.tag === 'ol') {
            if (!this.listCounters[numId]) this.listCounters[numId] = [];

            this.listCounters[numId][level]++;

            const numDef = this.numbering?.[numId];
            if (numDef) {
                for (let l = level + 1; l < Object.keys(numDef).length; l++) {
                    const deeperMeta = this.getListMeta(numId, l);
                    if (this.listCounters[numId]) {
                        this.listCounters[numId][l] = deeperMeta.start || 1;
                    }
                }
            }
        }

        return html;
    }

    async renderNonListParagraph(pChildren, content) {
        const pPrNode = this.findChild(pChildren, "w:pPr");
        const pPr = pPrNode ? pPrNode["w:pPr"] : null;
        if (content.trim() === '') return '<p> </p>';

        const pStyle = this.getParagraphStyle(pPr, pChildren);

        if (pPr) {
            const pStyleNode = this.findChild(pPr, "w:pStyle");
            if (pStyleNode && pStyleNode[":@"]) {
                const styleId = pStyleNode[":@"]["@_w:val"];

                if (styleId.match(/^Heading[1-6]$/i) || styleId.match(/^Ttulo[1-6]$/i)) {
                    const level = styleId.replace(/(heading|Ttulo)/i, '');
                    return `<h${level}${pStyle}>${content}</h${level}>`;
                } else if (styleId.match(/title/i)) {
                    return `<h1${pStyle}>${content}</h1>`;
                } else if (styleId.match(/quote/i)) {
                    return `<blockquote${pStyle}>${content}</blockquote>`;
                }
            }
        }

        return `<p${pStyle}>${content || ' '}</p>`;
    }

    getParagraphStyle(pPr, pChildren = null) {
        const defaultPPr = this.docDefaults?.pPr || [];
        const pStyleNode = pPr ? this.findChild(pPr, "w:pStyle") : null;
        const styleId = pStyleNode ? pStyleNode[":@"]["@_w:val"] : null;
        const styleDef = styleId ? this.globalStyles[styleId] : null;
        const stylePPr = styleDef ? styleDef.pPr : [];

        const directPPr = pPr || [];

        const mergedPPr = this.mergeProperties(
            this.mergeProperties(defaultPPr, stylePPr),
            directPPr
        );

        let style = "";
        if (mergedPPr) {
            const jcNode = this.findChild(mergedPPr, "w:jc");
            if (jcNode && jcNode[":@"]) {
                const align = jcNode[":@"]["@_w:val"];
                if (["left", "center", "right", "both"].includes(align)) {
                    style += `text-align: ${align === 'both' ? 'justify' : align};`;
                }
            }
            const spacingNode = this.findChild(mergedPPr, "w:spacing");
            if (spacingNode && spacingNode[":@"]) {
                const attrs = spacingNode[":@"];
                if (attrs["@_w:before"]) {
                    const beforeTwips = parseInt(attrs["@_w:before"], 10);
                    if (!isNaN(beforeTwips)) style += `margin-top: ${beforeTwips / 20}pt;`;
                }
                if (attrs["@_w:after"]) {
                    const afterTwips = parseInt(attrs["@_w:after"], 10);
                    if (!isNaN(afterTwips)) style += `margin-bottom: ${afterTwips / 20}pt;`;
                }
            }

            const indNode = this.findChild(mergedPPr, "w:ind");
            if (indNode && indNode[":@"]) {
                const attrs = indNode[":@"];
                if (attrs["@_w:left"]) {
                    const leftTwips = parseInt(attrs["@_w:left"], 10);
                    if (!isNaN(leftTwips)) style += `margin-left: ${leftTwips / 20}pt;`;
                }
                if (attrs["@_w:firstLine"]) {
                    const firstLineTwips = parseInt(attrs["@_w:firstLine"], 10);
                    if (!isNaN(firstLineTwips)) style += `text-indent: ${firstLineTwips / 20}pt;`;
                }
                if (attrs["@_w:hanging"]) {
                    const hangingTwips = parseInt(attrs["@_w:hanging"], 10);
                    if (!isNaN(hangingTwips)) style += `padding-left: ${hangingTwips / 20}pt; text-indent: -${hangingTwips / 20}pt;`;
                }
            }

            const pBdrNode = this.findChild(mergedPPr, "w:pBdr");
            if (pBdrNode) {
                const pBdrChildren = pBdrNode["w:pBdr"];
                const bottomBdrNode = this.findChild(pBdrChildren, "w:bottom");
                if (bottomBdrNode && bottomBdrNode[":@"]) {
                    const attrs = bottomBdrNode[":@"];
                    const size = parseInt(attrs["@_w:sz"], 10) / 8;
                    const space = parseInt(attrs["@_w:space"], 10) / 20;
                    const color = attrs["@_w:color"] && attrs["@_w:color"] !== "auto" ? `#${attrs["@_w:color"]}` : 'black';
                    const val = attrs["@_w:val"];

                    if (!isNaN(size) && val && val !== 'none') {
                        style += `border-bottom: ${size}pt solid ${color}; padding-bottom: ${space}pt;`;
                    }
                }
            }
        }

        if (this.paragraphContainsFloatedImage(pChildren)) {
            style += "overflow: auto;";
        }

        return style ? ` style="${style}"` : "";
    }

    paragraphContainsFloatedImage(pChildren) {
        if (!pChildren) return false;

        for (const childNode of pChildren) {
            if (childNode['w:r']) {
                const rChildren = childNode['w:r'];
                const drawingNode = this.findChild(rChildren, "w:drawing");
                if (drawingNode) {
                    const anchorNode = this.findChild(drawingNode["w:drawing"], "wp:anchor");
                    if (anchorNode) {
                        const anchorChildren = anchorNode["wp:anchor"];
                        if (this.findChild(anchorChildren, 'wp:wrapSquare') ||
                            this.findChild(anchorChildren, 'wp:wrapTight') ||
                            this.findChild(anchorChildren, 'wp:wrapThrough')) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    async processRun(rChildren, paragraphStyleDef) {

        const drawingNode = this.findChild(rChildren, "w:drawing");
        if (drawingNode) {
            return await this.processDrawing(drawingNode["w:drawing"]);
        }

        let contentHtml = "";
        let hasActualText = false;

        for (const childNode of rChildren) {
            const tagName = Object.keys(childNode)[0];

            if (tagName === 'w:t') {
                let text = (childNode["w:t"] || []).map(child => child["##text"] || "").join('');
                if (text) {
                    hasActualText = true;
                }
                const attrs = childNode[":@"];
                if (attrs && attrs["@_xml:space"] === "preserve") {
                    text = text
                        .replace(/^ /, ' ')
                        .replace(/ $/, ' ')
                        .replace(/  /g, '  ');
                }
                contentHtml += text.replace(/&/g, "&").replace(/</g, "<").replace(/>/g, ">");
            } else if (tagName === 'w:br') {
                contentHtml += "<br>";
            }
        }

        if (!hasActualText && !contentHtml.includes('<br>')) {
            return contentHtml.includes(" ") ? ' ' : '';
        }

        const defaultRPr = this.docDefaults?.rPr || [];
        const paragraphRPr = (paragraphStyleDef && paragraphStyleDef.rPr) ? paragraphStyleDef.rPr : [];

        const rPrNode = this.findChild(rChildren, "w:rPr");
        const directRPr = rPrNode ? rPrNode["w:rPr"] : [];

        const rStyleNode = this.findChild(directRPr, "w:rStyle");
        const styleId = rStyleNode ? rStyleNode[":@"]["@_w:val"] : null;
        const charStyleDef = styleId ? this.globalStyles[styleId] : null;

        let linkedStyleRPr = [];
        if (charStyleDef && charStyleDef.link && this.globalStyles[charStyleDef.link]) {
            linkedStyleRPr = this.globalStyles[charStyleDef.link].rPr || [];
        }

        const charStyleRPr = charStyleDef ? charStyleDef.rPr : [];

        let mergedRPr = this.mergeProperties(defaultRPr, paragraphRPr);
        mergedRPr = this.mergeProperties(mergedRPr, linkedStyleRPr);
        mergedRPr = this.mergeProperties(mergedRPr, charStyleRPr);
        mergedRPr = this.mergeProperties(mergedRPr, directRPr);

        let styleStart = "", styleEnd = "";
        let inlineStyles = "";

        if (mergedRPr) {
            if (this.findChild(mergedRPr, "w:b")) { styleStart += "<strong>"; styleEnd = "</strong>" + styleEnd; }
            if (this.findChild(mergedRPr, "w:i")) { styleStart += "<em>"; styleEnd = "</em>" + styleEnd; }

            let decorationParts = [];
            let decorationStyle = "";
            let decorationColor = "";

            const underlineNode = this.findChild(mergedRPr, "w:u");
            if (underlineNode && underlineNode[":@"]) {
                const val = underlineNode[":@"]["@_w:val"];
                if (val && val !== "none") {
                    decorationParts.push("underline");
                    if (val === "double") decorationStyle = "double";
                    else if (val === "wave") decorationStyle = "wavy";

                    const colorAttr = underlineNode[":@"]["@_w:color"];
                    if (colorAttr && colorAttr !== "auto") {
                        decorationColor = `#${colorAttr}`;
                    }
                }
            }

            const strikeNode = this.findChild(mergedRPr, "w:strike");
            const dstrikeNode = this.findChild(mergedRPr, "w:dstrike");
            if (strikeNode || dstrikeNode) {
                decorationParts.push("line-through");
                if (dstrikeNode) {
                    decorationStyle = "double";
                }
            }

            if (decorationParts.length > 0) {
                const fullDecoration = [decorationParts.join(' '), decorationStyle, decorationColor].filter(Boolean).join(' ');
                inlineStyles += `text-decoration: ${fullDecoration};`;
            }

            const vertAlignNode = this.findChild(mergedRPr, "w:vertAlign");
            if (vertAlignNode && vertAlignNode[":@"]) {
                const val = vertAlignNode[":@"]["@_w:val"];
                if (val === "superscript") { styleStart += "<sup>"; styleEnd = "</sup>" + styleEnd; }
                else if (val === "subscript") { styleStart += "<sub>"; styleEnd = "</sub>" + styleEnd; }
            }
            const colorNode = this.findChild(mergedRPr, "w:color");
            if (colorNode && colorNode[":@"]) {
                let colorVal = colorNode[":@"]["@_w:val"];
                if (colorVal && colorVal !== "auto") {
                    inlineStyles += `color:#${colorVal};`;
                }
            }
            const shdNode = this.findChild(mergedRPr, "w:shd");
            if (shdNode && shdNode[":@"]) {
                const fill = shdNode[":@"]["@_w:fill"];
                if (fill && fill !== "auto" && fill !== "clear") {
                    inlineStyles += `background-color:#${fill};`;
                }
            }
            const highlightNode = this.findChild(mergedRPr, "w:highlight");
            if (highlightNode && highlightNode[":@"]) {
                let highlightVal = highlightNode[":@"]["@_w:val"];
                const colorMap = { yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000", darkBlue: "#00008b", darkCyan: "#008b8b", darkMagenta: "#8b008b", darkRed: "#8b0000", darkYellow: "#b5a42e", darkGray: "#a9a9a9", lightGray: "#d3d3d3", black: "#000000", white: "#ffffff" };
                const mapped = colorMap[highlightVal];
                if (mapped) inlineStyles += `background-color:${mapped};`;
            }
            if (styleId === 'Hyperlink' && !inlineStyles.includes('text-decoration')) {
                inlineStyles += "text-decoration:underline;";
            }
            const szNode = this.findChild(mergedRPr, "w:sz");
            if (szNode && szNode[":@"]) {
                const size = parseInt(szNode[":@"]["@_w:val"], 10);
                if (!isNaN(size)) inlineStyles += `font-size:${size / 2}pt;`;
            }

            const positionNode = this.findChild(mergedRPr, "w:position");
            if (positionNode && positionNode[":@"]) {
                const pos = parseInt(positionNode[":@"]["@_w:val"], 10);
                if (!isNaN(pos) && pos > 0) {
                    inlineStyles += `padding-bottom: ${pos / 2}pt; display: inline-block; transform: translateY(${-pos / 2}pt);`;
                }
            }

            if (this.findChild(mergedRPr, "w:caps")) { inlineStyles += `text-transform:uppercase;`; }
            if (this.findChild(mergedRPr, "w:smallCaps")) { inlineStyles += `font-variant:small-caps;`; }

            const spacingNode = this.findChild(mergedRPr, "w:spacing");
            if (spacingNode && spacingNode[":@"]) {
                const attrs = spacingNode[":@"];
                let letterSpacing = "";
                if (attrs["@_w:val"]) {
                    const twips = parseInt(attrs["@_w:val"], 10);
                    if (!isNaN(twips)) letterSpacing = (twips / 20).toFixed(2) + "pt";
                }
                if (letterSpacing) {
                    if (inlineStyles.includes("letter-spacing")) {
                        inlineStyles = inlineStyles.replace(/letter-spacing:[^;]+;/, `letter-spacing:${letterSpacing};`);
                    } else {
                        inlineStyles += `letter-spacing:${letterSpacing};`;
                    }
                }
            }
        }
        if (inlineStyles) {
            styleStart += `<span style="${inlineStyles}">`; styleEnd = "</span>" + styleEnd;
        }

        return styleStart + contentHtml + styleEnd;
    }

    mapHyperlinkAnchor(hyperlinkNode) {
        const hyperlinkChildren = hyperlinkNode["w:hyperlink"];
        const anchorAttr = hyperlinkNode[":@"] ? hyperlinkNode[":@"]["@_w:anchor"] : null;
        if (!anchorAttr) return { hyperlinkChildren, anchorAttr, anchorDisplay: null, isExternal: false };
        let anchorDisplay = null;
        for (const child of hyperlinkChildren) {
            const cName = Object.keys(child)[0];
            if (cName === "w:r") {
                const rChildren = child["w:r"];
                const tNode = this.findChild(rChildren, "w:t");
                if (tNode) {
                    anchorDisplay = this.getText(tNode["w:t"][0]);
                    if (anchorDisplay) break;
                }
            }
        }
        const isExternalLikeFile = anchorDisplay &&
            (anchorDisplay.toLowerCase().includes(".docx") ||
                anchorDisplay.toLowerCase().includes(".pdf") ||
                anchorDisplay.toLowerCase().includes(".doc"));
        return { hyperlinkChildren, anchorAttr, anchorDisplay, isExternal: isExternalLikeFile };
    }

    async processHyperlink(hyperlinkNode) {
        if (!hyperlinkNode["w:hyperlink"]) return "";
        const rId = hyperlinkNode[":@"] ? hyperlinkNode[":@"]["@_r:id"] : null;
        const { hyperlinkChildren, anchorAttr, anchorDisplay, isExternal } = this.mapHyperlinkAnchor(hyperlinkNode);
        let anchorContent = "";
        for (const child of hyperlinkChildren) {
            const cName = Object.keys(child)[0];
            if (cName === "w:r") anchorContent += await this.processRun(child["w:r"]);
        }
        if (rId && this.relationships && this.relationships[rId]) {
            const rel = this.relationships[rId];
            if (rel.type.endsWith("/hyperlink")) {
                let href = rel.target;
                if (href && !href.toLowerCase().startsWith("http")) {
                    if (isExternal) {
                        const cleaned = this.extractAnchorFileName(anchorDisplay || "");
                        href = `https://www.gov.br/seedoc/shared/${cleaned}`;
                        anchorContent = anchorDisplay || cleaned || href;
                    } else {
                        href = `https://www.gov.br/seedoc/shared/${this.extractFileNameFromPath(rel.target)}`;
                    }
                }
                return `<a href="${href}" target="_blank" rel="noopener">${anchorContent}</a>`;
            }
        }
        if (anchorAttr) {
            let href = `#${anchorAttr}`;
            anchorContent = anchorDisplay || anchorContent;
            return `<a href="${href}">${anchorContent}</a>`;
        }
        return anchorContent;
    }

    async processDrawing(drawing) {
        const anchorNode = this.findChild(drawing, "wp:anchor") || this.findChild(drawing, "wp:inline");
        if (!anchorNode) return "";

        const anchorOrInlineAttributes = anchorNode[':@'];
        const isAnchor = !!this.findChild(drawing, "wp:anchor");
        const anchorOrInline = anchorNode["wp:anchor"] || anchorNode["wp:inline"];

        const graphicNode = this.findChild(anchorOrInline, "a:graphic");
        if (!graphicNode) return "";
        const graphicDataNode = this.findChild(graphicNode["a:graphic"], "a:graphicData");
        if (!graphicDataNode) return "";
        const picNode = this.findChild(graphicDataNode["a:graphicData"], "pic:pic");
        if (!picNode) return "";

        const pic = picNode["pic:pic"];
        const blipFillNode = this.findChild(pic, "pic:blipFill");
        const blipNode = blipFillNode ? this.findChild(blipFillNode["pic:blipFill"], "a:blip") : null;
        const relAttributes = blipNode ? blipNode[":@"] : null;
        const rId = relAttributes ? (relAttributes["@_r:embed"] || relAttributes["@_r:link"]) : null;
        if (!rId) return "";

        const nvPicPrNode = this.findChild(pic, "pic:nvPicPr");
        const cNvPrNode = nvPicPrNode ? this.findChild(nvPicPrNode["pic:nvPicPr"], "pic:cNvPr") : null;
        const altText = cNvPrNode && cNvPrNode[":@"] ? (cNvPrNode[":@"]["@_descr"] || cNvPrNode[":@"]["@_title"]) : "";

        const rel = rId ? this.relationships[rId] : null;
        if (!rel || !rel.type.includes("image")) return "";

        const imagePath = `word/${rel.target}`;
        const imageFile = this.zip.file(imagePath);
        if (!imageFile) return "";

        const base64 = await imageFile.async("base64");
        const mimeType = this.getMimeType(rel.target);

        let styles = "max-width:100%;height:auto;";

        if (isAnchor) {
            const wrapSquare = this.findChild(anchorOrInline, 'wp:wrapSquare');
            const wrapTopAndBottom = this.findChild(anchorOrInline, 'wp:wrapTopAndBottom');

            if (wrapSquare) {
                styles += "float:left;";
            } else if (wrapTopAndBottom) {
                styles += "clear:both;";
            }

            if (anchorOrInlineAttributes) {
                const emuToPt = (emu) => emu / 12700;
                const attrs = anchorOrInlineAttributes;

                if (attrs['@_distL']) {
                    styles += `margin-left: ${emuToPt(parseInt(attrs['@_distL'], 10))}pt;`;
                }
                if (attrs['@_distR']) {
                    styles += `margin-right: ${emuToPt(parseInt(attrs['@_distR'], 10))}pt;`;
                }
                if (attrs['@_distT']) {
                    styles += `margin-top: ${emuToPt(parseInt(attrs['@_distT'], 10))}pt;`;
                }
                if (attrs['@_distB']) {
                    styles += `margin-bottom: ${emuToPt(parseInt(attrs['@_distB'], 10))}pt;`;
                }
            }
        }

        return `<img src="data:${mimeType};base64,${base64}" alt="${altText}" style="${styles}" />`;
    }

    parseWidth(widthNode) {
        if (!widthNode || !widthNode[":@"]) return "";
        const attrs = widthNode[":@"];
        const type = attrs["@_w:type"] || "dxa";
        const val = parseInt(attrs["@_w:w"], 10);
        if (isNaN(val)) return "";
        if (type === "pct") return `width: ${val / 50}%;`;
        if (type === "dxa") return `width: ${val / 20}pt;`;
        if (type === "auto") return "width: auto;";
        return "";
    }

    parseBorder(borderDef) {
        if (!borderDef || !borderDef[":@"]) return null;
        const attrs = borderDef[":@"];
        const val = attrs["@_w:val"];
        if (!val || val === "none" || val === "nil") return null;

        const size = (parseInt(attrs["@_w:sz"], 10) || 4) / 8;
        const color = (attrs["@_w:color"] && attrs["@_w:color"] !== "auto") ? `#${attrs["@_w:color"]}` : "black";
        const styleMap = { single: 'solid', dashed: 'dashed', dotted: 'dotted', double: 'double' };
        const style = styleMap[val] || val;
        return `${size}pt ${style} ${color}`;
    }

    cssFromBorders(bordersNode) {
        let css = "";
        const borderTypes = ["top", "left", "bottom", "right"];
        borderTypes.forEach(type => {
            const borderDef = this.findChild(bordersNode["w:tblBorders"], `w:${type}`);
            const borderStyle = this.parseBorder(borderDef);
            if (borderStyle) {
                css += `border-${type}: ${borderStyle};`;
            }
        });
        return css;
    }

    parseCellBorders(tcPr) {
        const styles = {};
        if (!tcPr) return styles;
        const bordersNode = this.findChild(tcPr, "w:tcBorders");
        if (!bordersNode) return styles;

        const borderTypes = ["top", "left", "bottom", "right"];
        borderTypes.forEach(type => {
            const borderDef = this.findChild(bordersNode["w:tcBorders"], `w:${type}`);
            const borderStyle = this.parseBorder(borderDef);
            if (borderStyle) {
                styles[`border-${type}`] = borderStyle;
            }
        });
        return styles;
    }

    async processTable(tblChildren) {
        let directTblPrNode = this.findChild(tblChildren, "w:tblPr");
        let directTblPr = directTblPrNode ? directTblPrNode["w:tblPr"] : [];
        let mergedTblPr = directTblPr;

        const tblStyleNode = this.findChild(directTblPr, "w:tblStyle");
        if (tblStyleNode && tblStyleNode[":@"]) {
            const styleId = tblStyleNode[":@"]["@_w:val"];
            const tblStyleDef = this.globalStyles[styleId];
            if (tblStyleDef && tblStyleDef.tblPr) {
                mergedTblPr = this.mergeProperties(tblStyleDef.tblPr, directTblPr);
            }
        }

        let tableStyles = "border-collapse: collapse;";
        const tblWNode = this.findChild(mergedTblPr, "w:tblW");
        if (tblWNode) tableStyles += this.parseWidth(tblWNode);

        let insideHStyle = null;
        let insideVStyle = null;
        const tblBordersNode = this.findChild(mergedTblPr, "w:tblBorders");
        if (tblBordersNode) {
            tableStyles += this.cssFromBorders(tblBordersNode);
            const insideHDef = this.findChild(tblBordersNode["w:tblBorders"], 'w:insideH');
            insideHStyle = this.parseBorder(insideHDef);
            const insideVDef = this.findChild(tblBordersNode["w:tblBorders"], 'w:insideV');
            insideVStyle = this.parseBorder(insideVDef);
        }

        let colgroupHtml = "";
        const tblGridNode = this.findChild(tblChildren, "w:tblGrid");
        if (tblGridNode) {
            const gridColNodes = this.filterChildren(tblGridNode["w:tblGrid"], "w:gridCol");
            if (gridColNodes.length > 0) {
                colgroupHtml += "<colgroup>";
                for (const colNode of gridColNodes) {
                    const widthStyle = this.parseWidth(colNode);
                    colgroupHtml += `<col${widthStyle ? ` style="${widthStyle}"` : ""}>`;
                }
                colgroupHtml += "</colgroup>";
            }
        }

        let tableHtml = `<table border='1' style='${tableStyles}'>${colgroupHtml}`;
        const trNodes = this.filterChildren(tblChildren, "w:tr");

        for (let i = 0; i < trNodes.length; i++) {
            const trNode = trNodes[i];
            const trChildren = trNode["w:tr"];
            tableHtml += "<tr>";
            const tcNodes = this.filterChildren(trChildren, "w:tc");

            for (let j = 0; j < tcNodes.length; j++) {
                const tcNode = tcNodes[j];
                const tcChildren = tcNode["w:tc"];
                const tcPrNode = this.findChild(tcChildren, "w:tcPr");
                const tcPr = tcPrNode ? tcPrNode["w:tcPr"] : [];

                if (tcPr) {
                    const hMergeNode = this.findChild(tcPr, "w:hMerge");
                    if (hMergeNode && (!hMergeNode[":@"] || hMergeNode[":@"]["@_w:val"] !== "restart")) {
                        continue;
                    }
                    const vMergeNode = this.findChild(tcPr, "w:vMerge");
                    if (vMergeNode && (!vMergeNode[":@"] || (vMergeNode[":@"] && vMergeNode[":@"]["@_w:val"] !== "restart"))) {
                        continue;
                    }
                }

                let attrs = "";
                let cellStylesObj = { "vertical-align": "top", "padding": "4px" };

                if (insideHStyle && i < trNodes.length - 1) {
                    cellStylesObj['border-bottom'] = insideHStyle;
                }
                if (insideVStyle && j < tcNodes.length - 1) {
                    cellStylesObj['border-right'] = insideVStyle;
                }

                if (tcPr) {
                    const gridSpanNode = this.findChild(tcPr, "w:gridSpan");
                    if (gridSpanNode && gridSpanNode[":@"]) {
                        const colspan = parseInt(gridSpanNode[":@"]["@_w:val"], 10);
                        if (!isNaN(colspan) && colspan > 1) attrs += ` colspan="${colspan}"`;
                    }

                    const vMergeNode = this.findChild(tcPr, "w:vMerge");
                    if (vMergeNode && vMergeNode[":@"] && vMergeNode[":@"]["@_w:val"] === "restart") {
                        let rowspanCount = 1;
                        for (let k = i + 1; k < trNodes.length; k++) {
                            const nextTrChildren = trNodes[k]["w:tr"];
                            const nextTcNodes = this.filterChildren(nextTrChildren, "w:tc");
                            const cellInSameColumn = nextTcNodes[j];
                            if (cellInSameColumn) {
                                const nextTcChildren = cellInSameColumn["w:tc"];
                                const nextTcPrNode = this.findChild(nextTcChildren, "w:tcPr");
                                const nextTcPr = nextTcPrNode ? nextTcPrNode["w:tcPr"] : [];
                                const nextVMerge = nextTcPr ? this.findChild(nextTcPr, "w:vMerge") : null;
                                if (nextVMerge && (!nextVMerge[":@"] || (nextVMerge[":@"] && nextVMerge[":@"]["@_w:val"] !== "restart"))) {
                                    rowspanCount++;
                                } else {
                                    break;
                                }
                            } else {
                                break;
                            }
                        }
                        if (rowspanCount > 1) {
                            attrs += ` rowspan="${rowspanCount}"`;
                        }
                    }

                    const tcWNode = this.findChild(tcPr, 'w:tcW');
                    if (tcWNode) {
                        const widthStyle = this.parseWidth(tcWNode);
                        if (widthStyle) cellStylesObj['width'] = widthStyle.replace('width:', '');
                    }

                    const shdNode = this.findChild(tcPr, "w:shd");
                    if (shdNode && shdNode[":@"]) {
                        const fill = shdNode[":@"]["@_w:fill"];
                        if (fill && fill !== "auto" && fill !== "clear") {
                            cellStylesObj['background-color'] = `#${fill}`;
                        }
                    }

                    const explicitBorders = this.parseCellBorders(tcPr);
                    Object.assign(cellStylesObj, explicitBorders);
                }

                const cellContent = await this.processChildren(tcChildren);
                const styleString = Object.entries(cellStylesObj).map(([k, v]) => `${k}:${v}`).join(';');
                tableHtml += `<td${attrs} style="${styleString}">${cellContent || ' '}</td>`;
            }
            tableHtml += "</tr>";
        }
        tableHtml += "</table>";
        return tableHtml;
    }

    getMimeType(fileName) {
        const ext = fileName.split('.').pop().toLowerCase();
        switch (ext) {
            case 'png': return 'image/png';
            case 'jpg': case 'jpeg': return 'image/jpeg';
            case 'gif': return 'image/gif';
            case 'bmp': return 'image/bmp';
            case 'svg': return 'image/svg+xml';
            case 'wmf': return 'image/wmf';
            case 'emf': return 'image/emf';
            default: return 'application/octet-stream';
        }
    }
}
