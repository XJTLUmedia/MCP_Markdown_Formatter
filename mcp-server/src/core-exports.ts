import {
    Paragraph,
    TextRun,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType
} from 'docx';

import * as XLSX from 'xlsx';

// ... (previous code)

/**
 * Export table content only as CSV string
 */
export function generateCSV(content: string): string {
    const tableLines = content.match(/\|.*\|/g);
    if (!tableLines) return "";
    let csv = "";
    tableLines.forEach(line => {
        if (line.includes('---')) return;
        const cells = line.split('|').map(c => c.trim()).filter(c => c !== "").map(c => stripMarkdown(c));
        if (cells.length > 0) csv += cells.map(c => `"${c.replace(/"/g, '""')}"`).join(',') + "\n";
    });
    return csv;
}

/**
 * Generate JSON string representation
 */
export function generateJSON(content: string, title: string = 'document'): string {
    const data = {
        title: title,
        export_timestamp: new Date().toISOString(),
        content: stripMarkdown(content),
        structured_content: content.split('\n\n').map(block => stripMarkdown(block))
    };
    return JSON.stringify(data, null, 2);
}

/**
 * Generate XML string representation
 */
export function generateXML(content: string, title: string = 'document'): string {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>\n<document>\n`;
    xml += `  <title>${title}</title>\n`;
    xml += `  <content><![CDATA[${stripMarkdown(content)}]]></content>\n`;
    xml += `  <metadata>\n    <timestamp>${new Date().toISOString()}</timestamp>\n  </metadata>\n`;
    xml += `</document>`;
    return xml;
}

/**
 * Parse markdown content to a 2D array of strings for table-like representations (CSV, XLSX)
 */
export function parseMarkdownToTableData(content: string): string[][] {
    const tableData: string[][] = [];
    const paragraphs = content.split('\n\n');

    paragraphs.forEach(para => {
        const lines = para.trim().split('\n');
        const isTable = lines.some(l => l.includes('|'));

        if (isTable) {
            lines.forEach(line => {
                if (line.includes('---')) return;
                const cells = line.split('|').map(c => c.trim()).filter(c => c !== "").map(c => stripMarkdown(c));
                if (cells.length > 0) tableData.push(cells);
            });
        } else {
            const cleaned = stripMarkdown(para);
            if (cleaned) tableData.push([cleaned]);
        }
    });
    return tableData;
}

/**
 * Generate XLSX Buffer
 */
export function generateXLSXIndex(content: string): Buffer {
    const tableData = parseMarkdownToTableData(content);
    const ws = XLSX.utils.aoa_to_sheet(tableData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Document");
    // Write to buffer
    return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }) as Buffer;
}

export function stripMarkdown(text: string): string {
    if (!text) return "";

    let clean = text;

    // 1. Block Level: Code Blocks
    clean = clean.replace(/```[\s\S]*?```/g, m => m.replace(/```\w*\n?/g, '').replace(/```/g, '').trim());

    // 2. Block Level: Tables (Separator rows)
    clean = clean.replace(/^\|?[\s-:]+\|[\s-:|]*$/gm, '');

    // 3. Block Level: Horizontal Rules & Alternate Headings
    clean = clean.replace(/^[\s\t]*([*_-])\1{2,}\s*$/gm, '');
    clean = clean.replace(/^[\s\t]*[=-]{3,}\s*$/gm, '');

    // 4. Block Level: Blockquotes & ATX Headings
    clean = clean.replace(/^[\s\t]*>+\s?/gm, '');
    clean = clean.replace(/^#{1,6}\s+/gm, '');

    // 5. Inline: Multi-pass Emphasis (Bold, Italic, Strikethrough)
    for (let i = 0; i < 3; i++) {
        clean = clean.replace(/[*_]{3}([^*_]+)[*_]{3}/g, '$1');
        clean = clean.replace(/[*_]{2}([^*_]+)[*_]{2}/g, '$1');
        clean = clean.replace(/[*_]{1}([^*_]+)[*_]{1}/g, '$1');
        clean = clean.replace(/~~([^~]+)~~/g, '$1');
    }

    // 6. Inline: Math, Links, Images, and Extended Syntax
    clean = clean.replace(/\$\$(.*?)\$\$/gs, '$1');
    clean = clean.replace(/\$(.*?)\$/g, '$1');
    clean = clean.replace(/!\[([^\]]*)\]\([^)]+\)/g, '$1');
    clean = clean.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1');
    clean = clean.replace(/\[([^\]]+)\]\[[^\]]*\]/g, '$1');
    clean = clean.replace(/\[[ xX]\]\s+/g, '');
    clean = clean.replace(/\[\^[^\]]+\]/g, '');
    clean = clean.replace(/\{#[^}]+\}/g, '');
    clean = clean.replace(/[~^]([^~^]+)[~^]/g, '$1');

    // 7. Inline: Code & HTML
    clean = clean.replace(/`([^`]+)`/g, '$1');
    clean = clean.replace(/<[^>]*>/g, '');

    // 8. Final Polish: Pipes & Escaped Chars
    clean = clean.replace(/\|/g, ' ');
    clean = clean.replace(/\\([\\`*_{}[\]()#+\-.!|~^])/g, '$1');

    // 9. Normalization
    return clean
        .split('\n')
        .map(line => line.trim())
        .join('\n')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

/**
 * Parse markdown table to structured data
 */
export function parseMarkdownTable(tableText: string): { headers: string[], rows: string[][] } {
    const lines = tableText.trim().split('\n').filter(line => line.trim());
    if (lines.length < 2) return { headers: [], rows: [] };
    const headers = lines[0].split('|').map(c => c.trim()).filter(c => c);
    const rows: string[][] = [];
    for (let i = 2; i < lines.length; i++) {
        const cells = lines[i].split('|').map(c => c.trim()).filter(c => c);
        if (cells.length > 0) rows.push(cells);
    }
    return { headers, rows };
}

/**
 * Clean text by removing markdown symbols
 */
export function cleanMarkdownText(text: string): string {
    return stripMarkdown(text);
}

/**
 * Parse markdown content to LaTeX
 */
export function parseMarkdownToLaTeX(content: string): string {
    let processed = content
        .replace(/^# (.*)$/gm, '\\section{$1}')
        .replace(/^## (.*)$/gm, '\\subsection{$1}')
        .replace(/^### (.*)$/gm, '\\subsubsection{$1}')
        .replace(/\*\*(.*)\*\*/g, '\\textbf{$1}')
        .replace(/\*(.*)\*/g, '\\textit{$1}')
        .replace(/\$\$(.*?)\$\$/gs, '\\begin{equation}\n$1\n\\end{equation}')
        .replace(/\$(.*?)\$/g, '$ $1 $')
        .replace(/^-\s(.*)$/gm, '\\begin{itemize}\n\\item $1\n\\end{itemize}')
        .replace(/\\end{itemize}\n\\begin{itemize}/g, '');

    // Escape LaTeX special chars but try not to break our commands
    processed = processed.replace(/([_%$&~^\\{}])/g, (m) => m === '\\' ? m : `\\${m}`);

    // Final pass to remove any markdown-only artifacts (hashes, backticks, pipe)
    return processed.replace(/[*#`|]/g, '');
}

/**
 * Parse text with inline formatting to TextRuns
 */
export function parseInlineFormatting(text: string): TextRun[] {
    const runs: TextRun[] = [];
    const regex = /(\$\$.*?\$\$|\$.*?\$|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|___[^_]+___|__[^_]+__|_[^_]+_|`[^`]+`|<br\s*\/?>)/g;
    const parts = text.split(regex);
    for (const part of parts) {
        if (!part) continue;
        if (part.startsWith('$$') && part.endsWith('$$')) {
            runs.push(new TextRun({ text: part.slice(2, -2), italics: true, color: '4F46E5', font: 'Cambria Math' }));
        } else if (part.startsWith('$') && part.endsWith('$')) {
            runs.push(new TextRun({ text: part.slice(1, -1), italics: true, color: '4F46E5', font: 'Cambria Math' }));
        } else if (part.startsWith('***') && part.endsWith('***')) {
            runs.push(new TextRun({ text: part.slice(3, -3), bold: true, italics: true }));
        } else if (part.startsWith('___') && part.endsWith('___')) {
            runs.push(new TextRun({ text: part.slice(3, -3), bold: true, italics: true }));
        } else if (part.startsWith('**') && part.endsWith('**')) {
            runs.push(new TextRun({ text: part.slice(2, -2), bold: true }));
        } else if (part.startsWith('__') && part.endsWith('__')) {
            runs.push(new TextRun({ text: part.slice(2, -2), bold: true }));
        } else if (part.startsWith('*') && part.endsWith('*')) {
            runs.push(new TextRun({ text: part.slice(1, -1), italics: true }));
        } else if (part.startsWith('_') && part.endsWith('_')) {
            runs.push(new TextRun({ text: part.slice(1, -1), italics: true }));
        } else if (part.startsWith('`') && part.endsWith('`')) {
            runs.push(new TextRun({ text: part.slice(1, -1), font: 'Consolas', shading: { fill: 'F0F0F0' } }));
        } else if (part.match(/<br\s*\/?>/i)) {
            runs.push(new TextRun({ text: '', break: 1 }));
        } else {
            runs.push(new TextRun({ text: part }));
        }
    }
    return runs.length > 0 ? runs : [new TextRun({ text })];
}

/**
 * Create a Word table from parsed markdown table data
 */
export function createDocxTable(headers: string[], rows: string[][]): Table {
    const allRows: TableRow[] = [];
    if (headers.length > 0) {
        allRows.push(new TableRow({
            children: headers.map(header => new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: cleanMarkdownText(header), bold: true })], alignment: AlignmentType.LEFT })],
                shading: { fill: 'E5E7EB' }
            }))
        }));
    }
    for (const row of rows) {
        allRows.push(new TableRow({
            children: row.map(cell => new TableCell({ children: [new Paragraph({ children: parseInlineFormatting(cell) })] }))
        }));
    }
    return new Table({
        rows: allRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }
        }
    });
}

/**
 * Parse markdown content to docx elements
 */
export function parseMarkdownToDocx(content: string): (Paragraph | Table)[] {
    const elements: (Paragraph | Table)[] = [];
    const lines = content.split('\n');
    let i = 0;
    let inCodeBlock = false;
    let codeBlockContent: string[] = [];
    while (i < lines.length) {
        const line = lines[i];
        const trimmed = line.trim();
        if (/^(\*\*\*|---|__{3,})\s*$/.test(trimmed)) {
            elements.push(new Paragraph({ border: { bottom: { color: 'CCCCCC', space: 1, style: BorderStyle.SINGLE, size: 6 } }, spacing: { before: 200, after: 200 } }));
            i++; continue;
        }
        if (trimmed.startsWith('```')) {
            if (inCodeBlock) {
                elements.push(new Paragraph({ children: [new TextRun({ text: codeBlockContent.join('\n'), font: 'Consolas', size: 20 })], shading: { fill: 'F3F4F6' }, spacing: { before: 200, after: 200 } }));
                codeBlockContent = []; inCodeBlock = false;
            } else { inCodeBlock = true; }
            i++; continue;
        }
        if (inCodeBlock) { codeBlockContent.push(line); i++; continue; }
        if (i + 1 < lines.length) {
            const nextLine = lines[i + 1].trim();
            if (/^={3,}\s*$/.test(nextLine)) {
                elements.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: parseInlineFormatting(trimmed), spacing: { before: 400, after: 200 } }));
                i += 2; continue;
            } else if (/^-{3,}\s*$/.test(nextLine)) {
                elements.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: parseInlineFormatting(trimmed), spacing: { before: 300, after: 150 } }));
                i += 2; continue;
            }
        }
        if (trimmed.includes('|') && trimmed.startsWith('|')) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].trim().includes('|')) { tableLines.push(lines[i]); i++; }
            if (tableLines.length >= 2) {
                const { headers, rows } = parseMarkdownTable(tableLines.join('\n'));
                if (headers.length > 0) { elements.push(createDocxTable(headers, rows)); elements.push(new Paragraph({ spacing: { after: 200 } })); }
            }
            continue;
        }
        if (!trimmed) { elements.push(new Paragraph({ spacing: { after: 100 } })); i++; continue; }
        if (trimmed.startsWith('# ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: parseInlineFormatting(trimmed.slice(2)), spacing: { before: 400, after: 200 } }));
        } else if (trimmed.startsWith('## ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: parseInlineFormatting(trimmed.slice(3)), spacing: { before: 300, after: 150 } }));
        } else if (trimmed.startsWith('### ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: parseInlineFormatting(trimmed.slice(4)), spacing: { before: 250, after: 100 } }));
        } else if (trimmed.startsWith('#### ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_4, children: parseInlineFormatting(trimmed.slice(5)), spacing: { before: 200, after: 100 } }));
        } else if (trimmed.startsWith('>')) {
            const level = (trimmed.match(/^>+/g) || ['>'])[0].length;
            const text = trimmed.replace(/^>+\s*/, '');
            elements.push(new Paragraph({ indent: { left: 720 * level }, children: [new TextRun({ text: cleanMarkdownText(text), italics: true, color: '666666' })], spacing: { after: 100 }, shading: { fill: 'F9FAFB' } }));
        } else if (/^(\s*)[-*+]\s+/.test(line)) {
            const match = line.match(/^(\s*)([-*+]\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const text = line.replace(/^\s*[-*+]\s+/, '');
            elements.push(new Paragraph({ bullet: { level: indent }, children: parseInlineFormatting(text), spacing: { after: 80 } }));
        } else if (/^(\s*)\d+\.\s+/.test(line)) {
            const match = line.match(/^(\s*)(\d+\.\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const text = line.replace(/^\s*\d+\.\s+/, '');
            elements.push(new Paragraph({ numbering: { reference: 'default-numbering', level: indent }, children: parseInlineFormatting(text), spacing: { after: 80 } }));
        } else {
            const paragraphChildren = parseInlineFormatting(trimmed);
            if (line.endsWith('  ')) paragraphChildren.push(new TextRun({ text: '', break: 1 }));
            elements.push(new Paragraph({ children: paragraphChildren, spacing: { after: 150 } }));
        }
        i++;
    }
    return elements;
}

/**
 * RTF Helper: Encode string with Unicode support and RTF escaping
 */
export function encodeRTFText(str: string): string {
    let res = "";
    for (let i = 0; i < str.length; i++) {
        const charCode = str.charCodeAt(i);
        if (charCode > 127) {
            res += `\\u${charCode}?`;
        } else if (str[i] === '\\' || str[i] === '{' || str[i] === '}') {
            res += '\\' + str[i];
        } else {
            res += str[i];
        }
    }
    return res;
}

/**
 * RTF Helper: Parse inline markdown to RTF codes
 */
export function parseInlineToRTF(text: string): string {
    const regex = /(\$\$.*?\$\$|\$.*?\$|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|___[^_]+___|__[^_]+__|_[^_]+_|`[^`]+`|<br\s*\/?>)/g;
    const parts = text.split(regex);
    let result = "";

    for (const part of parts) {
        if (!part) continue;
        if (part.startsWith('$$') && part.endsWith('$$')) {
            result += `{\\i\\cf4\\f2 ${encodeRTFText(part.slice(2, -2))}}`;
        } else if (part.startsWith('$') && part.endsWith('$')) {
            result += `{\\i\\cf4\\f2 ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (part.startsWith('***') && part.endsWith('***')) {
            result += `{\\b\\i ${encodeRTFText(part.slice(3, -3))}}`;
        } else if (part.startsWith('___') && part.endsWith('___')) {
            result += `{\\b\\i ${encodeRTFText(part.slice(3, -3))}}`;
        } else if (part.startsWith('**') && part.endsWith('**')) {
            result += `{\\b ${encodeRTFText(part.slice(2, -2))}}`;
        } else if (part.startsWith('__') && part.endsWith('__')) {
            result += `{\\b ${encodeRTFText(part.slice(2, -2))}}`;
        } else if (part.startsWith('*') && part.endsWith('*')) {
            result += `{\\i ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (part.startsWith('_') && part.endsWith('_')) {
            result += `{\\i ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (part.startsWith('`') && part.endsWith('`')) {
            result += `{\\f1\\highlight3 ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (part.match(/<br\s*\/?>/i)) {
            result += "\\line ";
        } else {
            result += encodeRTFText(part);
        }
    }
    return result;
}

/**
 * RTF Helper: Main parser for Markdown to RTF conversion
 */
export function parseMarkdownToRTF(content: string): string {
    const lines = content.split('\n');
    let rtf = "";
    let i = 0;
    let inCodeBlock = false;
    let codeBlockContent: string[] = [];

    while (i < lines.length) {
        const line = lines[i];
        const trimmed = line.trim();

        // Horizontal Rule
        if (/^(\*\*\*|---|__{3,})\s*$/.test(trimmed)) {
            rtf += "\\pard\\sb200\\sa200\\brdrb\\brdrs\\brdrw10\\brdrcf6\\par\n";
            i++; continue;
        }

        // Code Block
        if (trimmed.startsWith('```')) {
            if (inCodeBlock) {
                rtf += "{\\pard\\f1\\fs20\\highlight3 " + encodeRTFText(codeBlockContent.join("\\line\n")) + "\\par}\n";
                codeBlockContent = []; inCodeBlock = false;
            } else { inCodeBlock = true; }
            i++; continue;
        }
        if (inCodeBlock) { codeBlockContent.push(line); i++; continue; }

        // Setext Headings
        if (i + 1 < lines.length) {
            const nextLine = lines[i + 1].trim();
            if (/^={3,}\s*$/.test(nextLine)) {
                rtf += "{\\pard\\b\\fs40\\sb400\\sa200 " + parseInlineToRTF(trimmed) + "\\par}\n";
                i += 2; continue;
            } else if (/^-{3,}\s*$/.test(nextLine)) {
                rtf += "{\\pard\\b\\fs32\\sb300\\sa150 " + parseInlineToRTF(trimmed) + "\\par}\n";
                i += 2; continue;
            }
        }

        // Tables
        if (trimmed.includes('|') && trimmed.startsWith('|')) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].trim().includes('|')) {
                tableLines.push(lines[i]);
                i++;
            }
            if (tableLines.length >= 2) {
                const { headers, rows } = parseMarkdownTable(tableLines.join('\n'));
                if (headers.length > 0) {
                    const cellWidth = 3000;
                    // Header Row
                    rtf += "\\trowd\\trgaph108\\trleft-108";
                    for (let j = 0; j < headers.length; j++) {
                        rtf += `\\clcbpat5\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx${(j + 1) * cellWidth}`;
                    }
                    rtf += "\\pard\\intbl\\ql ";
                    for (const h of headers) {
                        rtf += "{\\b " + parseInlineToRTF(h) + "}\\cell ";
                    }
                    rtf += "\\row\n";

                    // Data Rows
                    for (const row of rows) {
                        rtf += "\\trowd\\trgaph108\\trleft-108";
                        for (let j = 0; j < row.length; j++) {
                            rtf += `\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx${(j + 1) * cellWidth}`;
                        }
                        rtf += "\\pard\\intbl\\ql ";
                        for (const cell of row) {
                            rtf += parseInlineToRTF(cell) + "\\cell ";
                        }
                        rtf += "\\row\n";
                    }
                    rtf += "\\pard\\sa200\\par\n";
                }
            }
            continue;
        }

        if (!trimmed) {
            rtf += "\\pard\\sa100\\par\n";
            i++; continue;
        }

        // Headings
        if (trimmed.startsWith('# ')) {
            rtf += "{\\pard\\b\\fs40\\sb400\\sa200 " + parseInlineToRTF(trimmed.slice(2)) + "\\par}\n";
        } else if (trimmed.startsWith('## ')) {
            rtf += "{\\pard\\b\\fs32\\sb300\\sa150 " + parseInlineToRTF(trimmed.slice(3)) + "\\par}\n";
        } else if (trimmed.startsWith('### ')) {
            rtf += "{\\pard\\b\\fs28\\sb250\\sa100 " + parseInlineToRTF(trimmed.slice(4)) + "\\par}\n";
        } else if (trimmed.startsWith('#### ')) {
            rtf += "{\\pard\\b\\fs26\\sb200\\sa100 " + parseInlineToRTF(trimmed.slice(5)) + "\\par}\n";
        }
        else if (trimmed.startsWith('>')) {
            const level = (trimmed.match(/^>+/g) || ['>'])[0].length;
            const text = trimmed.replace(/^>+\s*/, '');
            rtf += `{\\pard\\li${level * 720}\\cf2\\i\\sa100 ` + parseInlineToRTF(text) + "\\par}\n";
        }
        else if (/^(\s*)[-*+]\s+/.test(line)) {
            const match = line.match(/^(\s*)([-*+]\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const text = line.replace(/^\s*[-*+]\s+/, '');
            rtf += `{\\pard\\li${(indent + 1) * 360}\\fi-360\\'b7\\tab ` + parseInlineToRTF(text) + "\\par}\n";
        }
        else if (/^(\s*)\d+\.\s+/.test(line)) {
            const match = line.match(/^(\s*)(\d+\.\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const number = match ? match[2] : "1. ";
            const text = line.replace(/^\s*\d+\.\s+/, '');
            rtf += `{\\pard\\li${(indent + 1) * 360}\\fi-360 ${number}\\tab ` + parseInlineToRTF(text) + "\\par}\n";
        }
        else {
            rtf += "{\\pard\\sa150 " + parseInlineToRTF(trimmed) + "\\par}\n";
        }
        i++;
    }
    return rtf;
}
