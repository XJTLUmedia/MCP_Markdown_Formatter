import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import { jsPDF } from 'jspdf';
import {
    Document,
    Packer,
    AlignmentType
} from 'docx';
import {
    stripMarkdown,
    parseMarkdownToRTF,
    parseMarkdownToDocx,
    parseMarkdownToLaTeX,
    cleanMarkdownText
} from './core-exports';

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

    const processed = parseMarkdownToLaTeX(content);

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
