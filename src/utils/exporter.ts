import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import { jsPDF } from 'jspdf';
import {
    Document,
    Packer,
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

/**
 * THE ULTIMATE MARKDOWN CLEANING ENGINE
 * A multi-pass aggressive stripper designed for professional data exports.
 */
function stripMarkdown(text: string): string {
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
 * Export content as Plain Text (.txt)
 */
export function exportAsText(content: string, filename: string = 'document'): void {
    const cleanText = stripMarkdown(content);
    const blob = new Blob([cleanText], { type: 'text/plain;charset=utf-8' });
    saveAs(blob, filename.endsWith('.txt') ? filename : `${filename}.txt`);
}

/**
 * Export content as Markdown (.md)
 */
export function exportAsMarkdown(content: string, filename: string = 'document'): void {
    const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
    saveAs(blob, filename.endsWith('.md') ? filename : `${filename}.md`);
}

/**
 * Export content as HTML (.html)
 */
export function exportAsHtml(htmlContent: string, filename: string = 'document'): void {
    const fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${filename}</title>
    <style>
        body { font-family: system-ui, -apple-system, sans-serif; max-width: 800px; margin: 40px auto; padding: 20px; line-height: 1.6; color: #1a1a1a; }
        h1, h2, h3 { color: #111; margin-top: 2em; }
        pre { background: #f4f4f4; padding: 15px; border-radius: 5px; overflow-x: auto; }
        code { font-family: monospace; background: #eee; padding: 2px 4px; border-radius: 3px; }
        table { border-collapse: collapse; width: 100%; margin: 1em 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background: #f8f8f8; }
        blockquote { border-left: 4px solid #ddd; margin: 0; padding-left: 1em; color: #666; }
    </style>
</head>
<body>${htmlContent}</body>
</html>`;
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    saveAs(blob, filename.endsWith('.html') ? filename : `${filename}.html`);
}

/**
 * Export content as LaTeX (.tex)
 */
export function exportAsLaTeX(content: string, filename: string = 'document'): void {
    let tex = `\\documentclass{article}
\\usepackage[utf8]{inputenc}
\\usepackage{amsmath}
\\usepackage{hyperref}
\\title{${filename}}
\\begin{document}
\\maketitle
`;

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
    processed = processed.replace(/[*#`|]/g, '');

    tex += processed + "\n\\end{document}";
    const blob = new Blob([tex], { type: 'application/x-latex;charset=utf-8' });
    saveAs(blob, filename.endsWith('.tex') ? filename : `${filename}.tex`);
}

/**
 * Export tables as CSV (.csv)
 */
export function exportAsCSV(content: string, filename: string = 'data'): void {
    const tableLines = content.match(/\|.*\|/g);
    if (!tableLines) return;
    let csv = "";
    tableLines.forEach(line => {
        if (line.includes('---')) return;
        const cells = line.split('|').map(c => c.trim()).filter(c => c !== "").map(c => stripMarkdown(c));
        if (cells.length > 0) csv += cells.map(c => `"${c.replace(/"/g, '""')}"`).join(',') + "\n";
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    saveAs(blob, filename.endsWith('.csv') ? filename : `${filename}.csv`);
}

/**
 * Export as JSON (.json)
 */
export function exportAsJSON(content: string, filename: string = 'document'): void {
    const data = {
        title: filename,
        export_timestamp: new Date().toISOString(),
        content: stripMarkdown(content),
        structured_content: content.split('\n\n').map(block => stripMarkdown(block))
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8' });
    saveAs(blob, filename.endsWith('.json') ? filename : `${filename}.json`);
}

/**
 * Export as XML (.xml)
 */
export function exportAsXML(content: string, filename: string = 'document'): void {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>\n<document>\n`;
    xml += `  <title>${filename}</title>\n`;
    xml += `  <content><![CDATA[${stripMarkdown(content)}]]></content>\n`;
    xml += `  <metadata>\n    <timestamp>${new Date().toISOString()}</timestamp>\n  </metadata>\n`;
    xml += `</document>`;
    const blob = new Blob([xml], { type: 'application/xml;charset=utf-8' });
    saveAs(blob, filename.endsWith('.xml') ? filename : `${filename}.xml`);
}

/**
 * Export as RTF (.rtf) with robust structure and formatting
 */
export function exportAsRTF(content: string, filename: string = 'document'): void {
    const rtfHeader = "{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\n" +
        "{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}{\\f1\\fnil\\fcharset0 Consolas;}{\\f2\\fnil\\fcharset0 Cambria Math;}}\n" +
        "{\\colortbl ;\\red0\\green0\\blue0;\\red102\\green102\\blue102;\\red240\\green240\\blue240;\\red79\\green70\\blue229;\\red229\\green231\\blue235;\\red204\\green204\\blue204;}\n" +
        "\\viewkind4\\uc1\\f0\\fs24 ";

    const rtfContent = parseMarkdownToRTF(content);
    const finalRtf = rtfHeader + rtfContent + "}";

    const blob = new Blob([finalRtf], { type: 'application/rtf' });
    saveAs(blob, filename.endsWith('.rtf') ? filename : `${filename}.rtf`);
}

/**
 * Export as Excel (.xlsx) preserving document narrative and table structures
 */
export function exportAsXLSX(content: string, filename: string = 'data'): void {
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

    const ws = XLSX.utils.aoa_to_sheet(tableData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Document");
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`);
}

/**
 * Export preview as Image (.png) - Uses high-fidelity Phantom Sandbox
 */
export async function exportAsImage(elementId: string, filename: string = 'preview'): Promise<void> {
    const original = document.getElementById(elementId);
    if (!original) return;

    // Create a clone and prepend to body (behind shield)
    const clone = original.cloneNode(true) as HTMLElement;
    const sandbox = document.createElement('div');
    sandbox.className = 'export-sandbox';
    sandbox.appendChild(clone);
    document.body.prepend(sandbox);

    try {
        // Essential delay for layout calculation
        await new Promise(resolve => setTimeout(resolve, 800));

        const dataUrl = await toPng(sandbox, {
            backgroundColor: '#ffffff',
            quality: 1,
            pixelRatio: 2,
            cacheBust: true
        });
        saveAs(dataUrl, `${filename}.png`);
    } catch (err) {
        console.error('Image Export Failure:', err);
        throw err;
    } finally {
        document.body.removeChild(sandbox);
    }
}

/**
 * Parse markdown table to structured data
 */
function parseMarkdownTable(tableText: string): { headers: string[], rows: string[][] } {
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
function cleanMarkdownText(text: string): string {
    return stripMarkdown(text);
}

/**
 * Parse text with inline formatting to TextRuns
 */
function parseInlineFormatting(text: string): TextRun[] {
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
function createDocxTable(headers: string[], rows: string[][]): Table {
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
function parseMarkdownToDocx(content: string): (Paragraph | Table)[] {
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
 * Export content as Word Document (.docx)
 */
export async function exportAsWord(markdownContent: string, filename: string = 'document'): Promise<void> {
    const elements = parseMarkdownToDocx(markdownContent);
    const doc = new Document({
        styles: { paragraphStyles: [{ id: 'Normal', name: 'Normal', run: { font: 'Calibri', size: 24 }, paragraph: { spacing: { line: 276 } } }] },
        numbering: { config: [{ reference: 'default-numbering', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
        sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: elements }]
    });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, filename.endsWith('.docx') ? filename : `${filename}.docx`);
}

/**
 * Export as PDF (.pdf) - Uses high-fidelity Multi-Page engine
 */
export async function exportAsPdf(_htmlContent: string, filename: string = 'document'): Promise<void> {
    const original = document.getElementById('preview-content');
    if (!original) return;

    // Create a clone and prepend to body (behind shield)
    const clone = original.cloneNode(true) as HTMLElement;
    const sandbox = document.createElement('div');
    sandbox.className = 'export-sandbox';
    sandbox.appendChild(clone);
    document.body.prepend(sandbox);

    try {
        // Essential stabilization delay
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Phase 3: High-Res Capture
        const dataUrl = await toPng(sandbox, {
            backgroundColor: '#ffffff',
            quality: 1,
            pixelRatio: 2,
            cacheBust: true
        });

        // Phase 4: Construct PDF with Multi-Page Support
        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        const imgProps = pdf.getImageProperties(dataUrl);
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        const imgWidth = pageWidth;
        const imgHeight = (imgProps.height * imgWidth) / imgProps.width;

        let heightLeft = imgHeight;
        let position = 0;
        let p = 0;

        // Page 1
        pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;

        // Consecutive Pages Waterfall
        while (heightLeft > 0) {
            p++;
            position = -(pageHeight * p);
            pdf.addPage();
            pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
        }

        pdf.save(`${filename}.pdf`);

    } catch (err) {
        console.error('Multi-Page PDF Failure:', err);
        throw err;
    } finally {
        // Phase 5: Cleanup
        document.body.removeChild(sandbox);
    }
}

/**
 * Copy rich HTML to clipboard for pasting into Word/Google Docs
 */
export async function copyRichText(htmlContent: string): Promise<boolean> {
    try {
        const blob = new Blob([htmlContent], { type: 'text/html' });
        const plainText = htmlContent.replace(/<[^>]*>/g, '');
        await navigator.clipboard.write([new ClipboardItem({ 'text/html': blob, 'text/plain': new Blob([plainText], { type: 'text/plain' }) })]);
        return true;
    } catch (err) {
        console.error('Failed to copy:', err);
        try { await navigator.clipboard.writeText(htmlContent.replace(/<[^>]*>/g, '')); return true; } catch { return false; }
    }
}

/**
 * RTF Helper: Encode string with Unicode support and RTF escaping
 */
function encodeRTFText(str: string): string {
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
function parseInlineToRTF(text: string): string {
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
function parseMarkdownToRTF(content: string): string {
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
