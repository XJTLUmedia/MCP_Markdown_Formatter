import {
    Paragraph,
    TextRun,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    ExternalHyperlink,
    FootnoteReferenceRun
} from 'docx';

import * as XLSX from 'xlsx';

// ... (previous code)

/**
 * Export table content only as CSV string
 */
export function generateCSV(content: string): string {
    // Strip fenced code blocks first so pipes inside code don't get treated as table rows
    const stripped = content.replace(/```[\s\S]*?```/g, '').replace(/~~~[\s\S]*?~~~/g, '');
    const tableLines = stripped.match(/^\s*\|.*\|\s*$/gm);
    if (!tableLines) return "";
    let csv = "";
    tableLines.forEach(line => {
        if (/^\s*\|?[\s\-:|]+\|?\s*$/.test(line)) return; // separator row
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

    // 5. Inline: Multi-pass Emphasis (Bold, Italic, Strikethrough, Highlight)
    // Require non-space adjacent to delimiters to avoid stripping literal asterisks in prose (e.g. `4*` or `*args`).
    for (let i = 0; i < 3; i++) {
        clean = clean.replace(/(?<![*_\w])[*_]{3}(\S[^*_\n]*?\S|\S)[*_]{3}(?![*_\w])/g, '$1');
        clean = clean.replace(/(?<![*_\w])[*_]{2}(\S[^*_\n]*?\S|\S)[*_]{2}(?![*_\w])/g, '$1');
        clean = clean.replace(/(?<![*_\w])[*_](\S[^*_\n]*?\S|\S)[*_](?![*_\w])/g, '$1');
        clean = clean.replace(/~~(\S[^~\n]*?\S|\S)~~/g, '$1');
        clean = clean.replace(/==(\S[^=\n]*?\S|\S)==/g, '$1');
    }

    // 5b. Block Level: Footnote definitions
    clean = clean.replace(/^\[\^[^\]]+\]:\s+.*$/gm, '');

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
 * Escape LaTeX special characters in text (not in commands)
 */
function escapeLatex(text: string): string {
    return text
        .replace(/\\/g, '\\textbackslash{}')
        .replace(/([&%$#_{}])/g, '\\$1')
        .replace(/~/g, '\\textasciitilde{}')
        .replace(/\^/g, '\\textasciicircum{}');
}

/**
 * Collect footnote definitions from markdown content.
 * Returns a map of label → text and the content with definitions removed.
 */
function collectFootnoteDefinitions(content: string): { footnoteMap: Record<string, string>, cleaned: string } {
    const footnoteMap: Record<string, string> = {};
    const cleaned = content.replace(/^\[\^([^\]]+)\]:\s+(.+)$/gm, (_m, label, text) => {
        footnoteMap[label] = text;
        return '';
    });
    return { footnoteMap, cleaned };
}

/**
 * Convert inline markdown to LaTeX (with optional footnote resolution)
 */
function convertInlineLatex(text: string, footnoteMap: Record<string, string> = {}): string {
    let out = text;
    // Extract math blocks first to protect them from escaping
    const mathBlocks: string[] = [];
    out = out.replace(/\$\$(.*?)\$\$/gs, (_m, math) => {
        mathBlocks.push(`\\[${math}\\]`);
        return `%%MATH${mathBlocks.length - 1}%%`;
    });
    out = out.replace(/\$(.*?)\$/g, (_m, math) => {
        mathBlocks.push(`$${math}$`);
        return `%%MATH${mathBlocks.length - 1}%%`;
    });

    // Inline code → \texttt{}
    out = out.replace(/`([^`]+)`/g, (_m, code) => `\\texttt{${escapeLatex(code)}}`);

    // Bold+Italic
    out = out.replace(/\*\*\*([^*]+)\*\*\*/g, '\\textbf{\\textit{$1}}');
    out = out.replace(/___([^_]+)___/g, '\\textbf{\\textit{$1}}');
    // Bold
    out = out.replace(/\*\*([^*]+)\*\*/g, '\\textbf{$1}');
    out = out.replace(/__([^_]+)__/g, '\\textbf{$1}');
    // Italic
    out = out.replace(/\*([^*]+)\*/g, '\\textit{$1}');
    out = out.replace(/_([^_]+)_/g, '\\textit{$1}');
    // Strikethrough
    out = out.replace(/~~([^~]+)~~/g, '\\sout{$1}');
    // Highlight
    out = out.replace(/==([^=]+)==/g, '\\hl{$1}');

    // Raw HTML inline tags
    out = out.replace(/<sup>([^<]+)<\/sup>/gi, '\\textsuperscript{$1}');
    out = out.replace(/<sub>([^<]+)<\/sub>/gi, '\\textsubscript{$1}');
    out = out.replace(/<mark>([^<]+)<\/mark>/gi, '\\hl{$1}');
    out = out.replace(/<kbd>([^<]+)<\/kbd>/gi, '\\texttt{$1}');

    // Links: [text](url)
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '\\href{$2}{$1}');
    // Images: ![alt](url)
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '\\includegraphics[width=\\linewidth]{$2}');

    // Footnote references: [^label] → \footnote{text} if definition exists, else superscript label
    out = out.replace(/\[\^([^\]]+)\]/g, (_m, label) => {
        const fnText = footnoteMap[label];
        return fnText ? `\\footnote{${fnText}}` : `\\textsuperscript{${label}}`;
    });

    // Escape remaining special chars that aren't part of LaTeX commands
    out = out.replace(/(?<!\\)&/g, '\\&');
    out = out.replace(/(?<!\\)%/g, '\\%');
    out = out.replace(/(?<!\\)#/g, '\\#');

    // Restore math blocks
    for (let idx = 0; idx < mathBlocks.length; idx++) {
        out = out.replace(`%%MATH${idx}%%`, mathBlocks[idx]);
    }

    return out;
}

/**
 * Convert markdown table to LaTeX tabular
 */
function convertTableToLatex(tableLines: string[], footnoteMap: Record<string, string> = {}): string {
    const rows = tableLines.filter(l => !/^\|[\s-:]+\|/.test(l.trim()));
    if (rows.length === 0) return '';
    const firstRow = rows[0].split('|').slice(1, -1).map(c => c.trim());
    const cols = firstRow.length;
    const colSpec = firstRow.map(() => 'l').join(' | ');

    let latex = `\\begin{tabular}{| ${colSpec} |}\n\\hline\n`;
    for (let r = 0; r < rows.length; r++) {
        const cells = rows[r].split('|').slice(1, -1).map(c => convertInlineLatex(c.trim(), footnoteMap));
        while (cells.length < cols) cells.push('');
        latex += cells.join(' & ') + ' \\\\\n\\hline\n';
    }
    latex += `\\end{tabular}`;
    return latex;
}

/**
 * Parse markdown content to LaTeX
 */
export function parseMarkdownToLaTeX(content: string): string {
    // Pre-process: collect footnote definitions and strip them
    const { footnoteMap, cleaned } = collectFootnoteDefinitions(content);
    const lines = cleaned.split('\n');
    const result: string[] = [];
    let i = 0;
    let inCodeBlock = false;
    let codeBlockContent: string[] = [];
    let inItemize = false;
    let inEnumerate = false;

    while (i < lines.length) {
        const line = lines[i];
        const trimmed = line.trim();

        // Code blocks
        if (trimmed.startsWith('```')) {
            if (inCodeBlock) {
                result.push(`\\begin{verbatim}`);
                result.push(codeBlockContent.join('\n'));
                result.push(`\\end{verbatim}`);
                codeBlockContent = [];
                inCodeBlock = false;
            } else {
                inCodeBlock = true;
            }
            i++;
            continue;
        }
        if (inCodeBlock) {
            codeBlockContent.push(line);
            i++;
            continue;
        }

        // Blank line — close open list environments
        if (!trimmed) {
            if (inItemize) { result.push('\\end{itemize}'); inItemize = false; }
            if (inEnumerate) { result.push('\\end{enumerate}'); inEnumerate = false; }
            result.push('');
            i++;
            continue;
        }

        // Horizontal rules
        if (/^(\*\*\*|---|__{3,})\s*$/.test(trimmed)) {
            result.push('\\begin{center}\\rule{0.5\\linewidth}{0.5pt}\\end{center}');
            i++;
            continue;
        }

        // Headings
        if (trimmed.startsWith('###### ')) {
            result.push(`\\textbf{${convertInlineLatex(trimmed.slice(7), footnoteMap)}}\\\\`);
        } else if (trimmed.startsWith('##### ')) {
            result.push(`\\subparagraph{${convertInlineLatex(trimmed.slice(6), footnoteMap)}}`);
        } else if (trimmed.startsWith('#### ')) {
            result.push(`\\paragraph{${convertInlineLatex(trimmed.slice(5), footnoteMap)}}`);
        } else if (trimmed.startsWith('### ')) {
            result.push(`\\subsubsection{${convertInlineLatex(trimmed.slice(4), footnoteMap)}}`);
        } else if (trimmed.startsWith('## ')) {
            result.push(`\\subsection{${convertInlineLatex(trimmed.slice(3), footnoteMap)}}`);
        } else if (trimmed.startsWith('# ')) {
            result.push(`\\section{${convertInlineLatex(trimmed.slice(2), footnoteMap)}}`);
        }
        // Blockquotes
        else if (trimmed.startsWith('>')) {
            const text = trimmed.replace(/^>+\s*/, '');
            result.push(`\\begin{quote}`);
            result.push(convertInlineLatex(text, footnoteMap));
            result.push(`\\end{quote}`);
        }
        // Task lists
        else if (/^\s*[-*+]\s+\[[ xX]\]\s+/.test(line)) {
            if (!inItemize) { result.push('\\begin{itemize}'); inItemize = true; }
            const checked = /\[x\]/i.test(line);
            const text = line.replace(/^\s*[-*+]\s+\[[ xX]\]\s+/, '');
            result.push(`\\item[${checked ? '$\\boxtimes$' : '$\\square$'}] ${convertInlineLatex(text, footnoteMap)}`);
        }
        // Unordered lists
        else if (/^\s*[-*+]\s+/.test(line)) {
            if (!inItemize) { result.push('\\begin{itemize}'); inItemize = true; }
            const text = line.replace(/^\s*[-*+]\s+/, '');
            result.push(`\\item ${convertInlineLatex(text, footnoteMap)}`);
        }
        // Ordered lists
        else if (/^\s*\d+\.\s+/.test(line)) {
            if (!inEnumerate) { result.push('\\begin{enumerate}'); inEnumerate = true; }
            const text = line.replace(/^\s*\d+\.\s+/, '');
            result.push(`\\item ${convertInlineLatex(text, footnoteMap)}`);
        }
        // Tables
        else if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].trim().startsWith('|')) {
                tableLines.push(lines[i]);
                i++;
            }
            result.push(convertTableToLatex(tableLines, footnoteMap));
            continue;
        }
        // Regular paragraph
        else {
            if (inItemize) { result.push('\\end{itemize}'); inItemize = false; }
            if (inEnumerate) { result.push('\\end{enumerate}'); inEnumerate = false; }
            result.push(convertInlineLatex(trimmed, footnoteMap));
        }
        i++;
    }

    // Close any open environments
    if (inItemize) result.push('\\end{itemize}');
    if (inEnumerate) result.push('\\end{enumerate}');
    if (inCodeBlock) {
        result.push('\\begin{verbatim}');
        result.push(codeBlockContent.join('\n'));
        result.push('\\end{verbatim}');
    }

    return result.join('\n');
}

/**
 * Parse text with inline formatting to TextRuns
 */
export function parseInlineFormatting(text: string, footnoteIdMap?: Record<string, number>): (TextRun | ExternalHyperlink | FootnoteReferenceRun)[] {
    const runs: (TextRun | ExternalHyperlink | FootnoteReferenceRun)[] = [];
    const regex = /(\$\$.*?\$\$|\$.*?\$|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|___[^_]+___|__[^_]+__|_[^_]+_|~~[^~]+~~|==[^=]+=+|`[^`]+`|!\[[^\]]*\]\([^)]+\)|\[[^\]]+\]\([^)]+\)|\[\^[^\]]+\]|<sup>[^<]+<\/sup>|<sub>[^<]+<\/sub>|<mark>[^<]+<\/mark>|<kbd>[^<]+<\/kbd>|<br\s*\/?>)/gi;
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
        } else if (part.startsWith('~~') && part.endsWith('~~')) {
            runs.push(new TextRun({ text: part.slice(2, -2), strike: true }));
        } else if (part.startsWith('==') && part.endsWith('==')) {
            runs.push(new TextRun({ text: part.slice(2, -2), highlight: 'yellow' }));
        } else if (part.startsWith('*') && part.endsWith('*')) {
            runs.push(new TextRun({ text: part.slice(1, -1), italics: true }));
        } else if (part.startsWith('_') && part.endsWith('_')) {
            runs.push(new TextRun({ text: part.slice(1, -1), italics: true }));
        } else if (part.startsWith('`') && part.endsWith('`')) {
            runs.push(new TextRun({ text: part.slice(1, -1), font: 'Consolas', shading: { fill: 'F0F0F0' } }));
        } else if (/^<sup>/i.test(part)) {
            const m = part.match(/<sup>([^<]+)<\/sup>/i);
            if (m) runs.push(new TextRun({ text: m[1], superScript: true }));
        } else if (/^<sub>/i.test(part)) {
            const m = part.match(/<sub>([^<]+)<\/sub>/i);
            if (m) runs.push(new TextRun({ text: m[1], subScript: true }));
        } else if (/^<mark>/i.test(part)) {
            const m = part.match(/<mark>([^<]+)<\/mark>/i);
            if (m) runs.push(new TextRun({ text: m[1], highlight: 'yellow' }));
        } else if (/^<kbd>/i.test(part)) {
            const m = part.match(/<kbd>([^<]+)<\/kbd>/i);
            if (m) runs.push(new TextRun({ text: m[1], font: 'Consolas', shading: { fill: 'E5E7EB' } }));
        } else if (/^\[\^/.test(part)) {
            const m = part.match(/^\[\^([^\]]+)\]$/);
            if (m && footnoteIdMap && footnoteIdMap[m[1]] !== undefined) {
                runs.push(new FootnoteReferenceRun(footnoteIdMap[m[1]]));
            } else if (m) {
                runs.push(new TextRun({ text: m[1], superScript: true }));
            }
        } else if (/^!\[/.test(part)) {
            const m = part.match(/^!\[([^\]]*)\]\(([^)]+)\)/);
            if (m) runs.push(new TextRun({ text: m[1] ? `[Image: ${m[1]}]` : '[Image]', italics: true, color: '6B7280' }));
        } else if (/^\[/.test(part)) {
            const m = part.match(/^\[([^\]]+)\]\(([^)]+)\)/);
            if (m) {
                runs.push(new ExternalHyperlink({ children: [new TextRun({ text: m[1], color: '2563EB', underline: { type: 'single' } })], link: m[2] }));
            }
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
export function createDocxTable(headers: string[], rows: string[][], footnoteIdMap?: Record<string, number>): Table {
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
            children: row.map(cell => new TableCell({ children: [new Paragraph({ children: parseInlineFormatting(cell, footnoteIdMap) })] }))
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
 * Result from parseMarkdownToDocx containing both elements and footnote definitions
 */
export interface DocxParseResult {
    elements: (Paragraph | Table)[];
    footnotes: Record<number, { children: Paragraph[] }>;
}

/**
 * Parse markdown content to docx elements with footnote support
 */
export function parseMarkdownToDocx(content: string): DocxParseResult {
    // Pre-process: collect footnote definitions
    const { footnoteMap, cleaned } = collectFootnoteDefinitions(content);

    // Assign numeric IDs to footnotes (docx requires numeric IDs starting from 1)
    const footnoteIdMap: Record<string, number> = {};
    const docxFootnotes: Record<number, { children: Paragraph[] }> = {};
    let fnId = 1;
    for (const label of Object.keys(footnoteMap)) {
        footnoteIdMap[label] = fnId;
        docxFootnotes[fnId] = {
            children: [new Paragraph({ children: [new TextRun({ text: footnoteMap[label] })] })]
        };
        fnId++;
    }

    const elements: (Paragraph | Table)[] = [];
    const lines = cleaned.split('\n');
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
                elements.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: parseInlineFormatting(trimmed, footnoteIdMap), spacing: { before: 400, after: 200 } }));
                i += 2; continue;
            } else if (/^-{3,}\s*$/.test(nextLine)) {
                elements.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: parseInlineFormatting(trimmed, footnoteIdMap), spacing: { before: 300, after: 150 } }));
                i += 2; continue;
            }
        }
        if (trimmed.includes('|') && trimmed.startsWith('|')) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].trim().includes('|')) { tableLines.push(lines[i]); i++; }
            if (tableLines.length >= 2) {
                const { headers, rows } = parseMarkdownTable(tableLines.join('\n'));
                if (headers.length > 0) { elements.push(createDocxTable(headers, rows, footnoteIdMap)); elements.push(new Paragraph({ spacing: { after: 200 } })); }
            }
            continue;
        }
        if (!trimmed) { elements.push(new Paragraph({ spacing: { after: 100 } })); i++; continue; }
        // Note: footnote definitions were already stripped by collectFootnoteDefinitions upstream,
        // so we don't need a guard here.
        if (trimmed.startsWith('# ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: parseInlineFormatting(trimmed.slice(2), footnoteIdMap), spacing: { before: 400, after: 200 } }));
        } else if (trimmed.startsWith('## ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: parseInlineFormatting(trimmed.slice(3), footnoteIdMap), spacing: { before: 300, after: 150 } }));
        } else if (trimmed.startsWith('### ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: parseInlineFormatting(trimmed.slice(4), footnoteIdMap), spacing: { before: 250, after: 100 } }));
        } else if (trimmed.startsWith('#### ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_4, children: parseInlineFormatting(trimmed.slice(5), footnoteIdMap), spacing: { before: 200, after: 100 } }));
        } else if (trimmed.startsWith('##### ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_5, children: parseInlineFormatting(trimmed.slice(6), footnoteIdMap), spacing: { before: 200, after: 80 } }));
        } else if (trimmed.startsWith('###### ')) {
            elements.push(new Paragraph({ heading: HeadingLevel.HEADING_6, children: parseInlineFormatting(trimmed.slice(7), footnoteIdMap), spacing: { before: 150, after: 80 } }));
        } else if (/^\s*[-*+]\s+\[[ xX]\]\s+/.test(line)) {
            const checked = /\[[xX]\]/.test(line);
            const text = line.replace(/^\s*[-*+]\s+\[[ xX]\]\s+/, '');
            const match = line.match(/^(\s*)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            elements.push(new Paragraph({
                indent: { left: (indent + 1) * 360 },
                children: [
                    new TextRun({ text: checked ? '☑ ' : '☐ ', font: 'Segoe UI Symbol' }),
                    ...parseInlineFormatting(text, footnoteIdMap)
                ],
                spacing: { after: 80 }
            }));
        } else if (trimmed.startsWith('>')) {
            const level = (trimmed.match(/^>+/g) || ['>'])[0].length;
            const text = trimmed.replace(/^>+\s*/, '');
            elements.push(new Paragraph({ indent: { left: 720 * level }, children: [new TextRun({ text: cleanMarkdownText(text), italics: true, color: '666666' })], spacing: { after: 100 }, shading: { fill: 'F9FAFB' } }));
        } else if (/^(\s*)[-*+]\s+/.test(line)) {
            const match = line.match(/^(\s*)([-*+]\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const text = line.replace(/^\s*[-*+]\s+/, '');
            elements.push(new Paragraph({ bullet: { level: indent }, children: parseInlineFormatting(text, footnoteIdMap), spacing: { after: 80 } }));
        } else if (/^(\s*)\d+\.\s+/.test(line)) {
            const match = line.match(/^(\s*)(\d+\.\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const text = line.replace(/^\s*\d+\.\s+/, '');
            elements.push(new Paragraph({ numbering: { reference: 'default-numbering', level: indent }, children: parseInlineFormatting(text, footnoteIdMap), spacing: { after: 80 } }));
        } else {
            const paragraphChildren = parseInlineFormatting(trimmed, footnoteIdMap);
            if (line.endsWith('  ')) paragraphChildren.push(new TextRun({ text: '', break: 1 }));
            elements.push(new Paragraph({ children: paragraphChildren, spacing: { after: 150 } }));
        }
        i++;
    }
    return { elements, footnotes: docxFootnotes };
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
export function parseInlineToRTF(text: string, footnoteMap?: Record<string, string>): string {
    const regex = /(\$\$.*?\$\$|\$.*?\$|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|___[^_]+___|__[^_]+__|_[^_]+_|~~[^~]+~~|==[^=]+=+|`[^`]+`|!\[[^\]]*\]\([^)]+\)|\[[^\]]+\]\([^)]+\)|\[\^[^\]]+\]|<sup>[^<]+<\/sup>|<sub>[^<]+<\/sub>|<mark>[^<]+<\/mark>|<kbd>[^<]+<\/kbd>|<br\s*\/?>)/gi;
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
        } else if (part.startsWith('~~') && part.endsWith('~~')) {
            result += `{\\strike ${encodeRTFText(part.slice(2, -2))}}`;
        } else if (part.startsWith('==') && part.endsWith('==')) {
            result += `{\\highlight7 ${encodeRTFText(part.slice(2, -2))}}`;
        } else if (part.startsWith('*') && part.endsWith('*')) {
            result += `{\\i ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (part.startsWith('_') && part.endsWith('_')) {
            result += `{\\i ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (part.startsWith('`') && part.endsWith('`')) {
            result += `{\\f1\\highlight3 ${encodeRTFText(part.slice(1, -1))}}`;
        } else if (/^<sup>/i.test(part)) {
            const m = part.match(/<sup>([^<]+)<\/sup>/i);
            if (m) result += `{\\super ${encodeRTFText(m[1])}}`;
        } else if (/^<sub>/i.test(part)) {
            const m = part.match(/<sub>([^<]+)<\/sub>/i);
            if (m) result += `{\\sub ${encodeRTFText(m[1])}}`;
        } else if (/^<mark>/i.test(part)) {
            const m = part.match(/<mark>([^<]+)<\/mark>/i);
            if (m) result += `{\\highlight7 ${encodeRTFText(m[1])}}`;
        } else if (/^<kbd>/i.test(part)) {
            const m = part.match(/<kbd>([^<]+)<\/kbd>/i);
            if (m) result += `{\\f1\\highlight3 ${encodeRTFText(m[1])}}`;
        } else if (/^\[\^/.test(part)) {
            const m = part.match(/^\[\^([^\]]+)\]$/);
            if (m && footnoteMap && footnoteMap[m[1]]) {
                result += `{\\super ${encodeRTFText(m[1])}}{\\*\\footnote\\pard\\plain\\s99\\f0\\fs20 {\\super ${encodeRTFText(m[1])}} ${encodeRTFText(footnoteMap[m[1]])}}`;
            } else if (m) {
                result += `{\\super ${encodeRTFText(m[1])}}`;
            }
        } else if (/^!\[/.test(part)) {
            const m = part.match(/^!\[([^\]]*)\]\(([^)]+)\)/);
            if (m) result += `{\\i\\cf2 [Image: ${encodeRTFText(m[1] || 'image')}]}`;
        } else if (/^\[/.test(part)) {
            const m = part.match(/^\[([^\]]+)\]\(([^)]+)\)/);
            if (m) result += `{\\field{\\*\\fldinst HYPERLINK "${m[2]}"}{\\fldrslt\\ul\\cf1 ${encodeRTFText(m[1])}}}`;
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
    // Pre-process: collect footnote definitions
    const { footnoteMap, cleaned } = collectFootnoteDefinitions(content);
    const lines = cleaned.split('\n');
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
                rtf += "{\\pard\\b\\fs40\\sb400\\sa200 " + parseInlineToRTF(trimmed, footnoteMap) + "\\par}\n";
                i += 2; continue;
            } else if (/^-{3,}\s*$/.test(nextLine)) {
                rtf += "{\\pard\\b\\fs32\\sb300\\sa150 " + parseInlineToRTF(trimmed, footnoteMap) + "\\par}\n";
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
                        rtf += "{\\b " + parseInlineToRTF(h, footnoteMap) + "}\\cell ";
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
                            rtf += parseInlineToRTF(cell, footnoteMap) + "\\cell ";
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
            rtf += "{\\pard\\b\\fs40\\sb400\\sa200 " + parseInlineToRTF(trimmed.slice(2), footnoteMap) + "\\par}\n";
        } else if (trimmed.startsWith('## ')) {
            rtf += "{\\pard\\b\\fs32\\sb300\\sa150 " + parseInlineToRTF(trimmed.slice(3), footnoteMap) + "\\par}\n";
        } else if (trimmed.startsWith('### ')) {
            rtf += "{\\pard\\b\\fs28\\sb250\\sa100 " + parseInlineToRTF(trimmed.slice(4), footnoteMap) + "\\par}\n";
        } else if (trimmed.startsWith('#### ')) {
            rtf += "{\\pard\\b\\fs26\\sb200\\sa100 " + parseInlineToRTF(trimmed.slice(5), footnoteMap) + "\\par}\n";
        } else if (trimmed.startsWith('##### ')) {
            rtf += "{\\pard\\b\\fs24\\sb200\\sa80 " + parseInlineToRTF(trimmed.slice(6), footnoteMap) + "\\par}\n";
        } else if (trimmed.startsWith('###### ')) {
            rtf += "{\\pard\\b\\fs22\\sb150\\sa80 " + parseInlineToRTF(trimmed.slice(7), footnoteMap) + "\\par}\n";
        }
        else if (/^\s*[-*+]\s+\[[ xX]\]\s+/.test(line)) {
            const checked = /\[[xX]\]/.test(line);
            const text = line.replace(/^\s*[-*+]\s+\[[ xX]\]\s+/, '');
            const match = line.match(/^(\s*)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            rtf += `{\\pard\\li${(indent + 1) * 360}\\fi-360 ${checked ? '\\u9745?' : '\\u9744?'}\\tab ` + parseInlineToRTF(text, footnoteMap) + "\\par}\n";
        }
        else if (trimmed.startsWith('>')) {
            const level = (trimmed.match(/^>+/g) || ['>'])[0].length;
            const text = trimmed.replace(/^>+\s*/, '');
            rtf += `{\\pard\\li${level * 720}\\cf2\\i\\sa100 ` + parseInlineToRTF(text, footnoteMap) + "\\par}\n";
        }
        else if (/^(\s*)[-*+]\s+/.test(line)) {
            const match = line.match(/^(\s*)([-*+]\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const text = line.replace(/^\s*[-*+]\s+/, '');
            rtf += `{\\pard\\li${(indent + 1) * 360}\\fi-360\\'b7\\tab ` + parseInlineToRTF(text, footnoteMap) + "\\par}\n";
        }
        else if (/^(\s*)\d+\.\s+/.test(line)) {
            const match = line.match(/^(\s*)(\d+\.\s+)/);
            const indent = match ? Math.floor(match[1].length / 4) : 0;
            const number = match ? match[2] : "1. ";
            const text = line.replace(/^\s*\d+\.\s+/, '');
            rtf += `{\\pard\\li${(indent + 1) * 360}\\fi-360 ${number}\\tab ` + parseInlineToRTF(text, footnoteMap) + "\\par}\n";
        }
        else {
            rtf += "{\\pard\\sa150 " + parseInlineToRTF(trimmed, footnoteMap) + "\\par}\n";
        }
        i++;
    }
    return rtf;
}
