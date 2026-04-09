import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import { toPng } from 'html-to-image';
import { Document, Packer, AlignmentType } from 'docx';
import {
  parseMarkdownToRTF,
  parseMarkdownToDocx,
  parseMarkdownToLaTeX,
  cleanMarkdownText,
  generateCSV,
  generateJSON,
  generateXML,
  generateXLSXIndex,
} from './core-exports';
import {
  markdownToSlack,
  markdownToDiscord,
  markdownToJira,
  markdownToConfluence,
  markdownToAsciiDoc,
  markdownToRST,
  markdownToMediaWiki,
  markdownToBBCode,
  markdownToTextile,
  markdownToOrgMode,
} from './platform-converters';
import { markdownToEmailHtml } from './email-html';
import { markdownToRichHtml } from './formatter';

// ── Types ────────────────────────────────────────────────────────────

export interface BatchItem {
  id: string;
  filename: string;
  content: string;
}

export interface BatchFormat {
  id: string;
  label: string;
  category: 'document' | 'data' | 'platform';
  extension: string;
}

export interface BatchResult {
  itemId: string;
  filename: string;
  format: string;
  success: boolean;
  blob?: Blob;
  error?: string;
}

export interface BatchProgress {
  total: number;
  completed: number;
  current: string;
  results: BatchResult[];
}

// ── Available formats ────────────────────────────────────────────────

export const BATCH_FORMATS: BatchFormat[] = [
  // Documents
  { id: 'docx', label: 'Word (.docx)', category: 'document', extension: '.docx' },
  { id: 'pdf', label: 'PDF (.pdf)', category: 'document', extension: '.pdf' },
  { id: 'html', label: 'HTML (.html)', category: 'document', extension: '.html' },
  { id: 'latex', label: 'LaTeX (.tex)', category: 'document', extension: '.tex' },
  { id: 'rtf', label: 'RTF (.rtf)', category: 'document', extension: '.rtf' },
  { id: 'txt', label: 'Plain Text (.txt)', category: 'document', extension: '.txt' },
  { id: 'md', label: 'Markdown (.md)', category: 'document', extension: '.md' },
  { id: 'email', label: 'Email HTML (.html)', category: 'document', extension: '.html' },
  // Data
  { id: 'csv', label: 'CSV (.csv)', category: 'data', extension: '.csv' },
  { id: 'json', label: 'JSON (.json)', category: 'data', extension: '.json' },
  { id: 'xml', label: 'XML (.xml)', category: 'data', extension: '.xml' },
  { id: 'xlsx', label: 'Excel (.xlsx)', category: 'data', extension: '.xlsx' },
  // Platform
  { id: 'slack', label: 'Slack', category: 'platform', extension: '.txt' },
  { id: 'discord', label: 'Discord', category: 'platform', extension: '.md' },
  { id: 'jira', label: 'JIRA', category: 'platform', extension: '.txt' },
  { id: 'confluence', label: 'Confluence', category: 'platform', extension: '.txt' },
  { id: 'asciidoc', label: 'AsciiDoc', category: 'platform', extension: '.adoc' },
  { id: 'rst', label: 'reStructuredText', category: 'platform', extension: '.rst' },
  { id: 'mediawiki', label: 'MediaWiki', category: 'platform', extension: '.wiki' },
  { id: 'bbcode', label: 'BBCode', category: 'platform', extension: '.txt' },
  { id: 'textile', label: 'Textile', category: 'platform', extension: '.textile' },
  { id: 'orgmode', label: 'Org Mode', category: 'platform', extension: '.org' },
];

// ── Helpers ──────────────────────────────────────────────────────────

function getOutputFilename(basename: string, formatId: string): string {
  switch (formatId) {
    case 'email': return `${basename}.email.html`;
    case 'slack': return `${basename}.slack.txt`;
    case 'discord': return `${basename}.discord.md`;
    case 'jira': return `${basename}.jira.txt`;
    case 'confluence': return `${basename}.confluence.txt`;
    case 'mediawiki': return `${basename}.mediawiki.txt`;
    case 'bbcode': return `${basename}.bbcode.txt`;
    default: {
      const fmt = BATCH_FORMATS.find(f => f.id === formatId);
      return `${basename}${fmt?.extension || '.txt'}`;
    }
  }
}

async function convertSingle(content: string, filename: string, formatId: string): Promise<Blob> {
  const basename = filename.replace(/\.[^.]+$/, '') || filename;

  switch (formatId) {
    case 'txt':
      return new Blob([cleanMarkdownText(content)], { type: 'text/plain;charset=utf-8' });

    case 'md':
      return new Blob([content], { type: 'text/markdown;charset=utf-8' });

    case 'csv':
      return new Blob([generateCSV(content) || ''], { type: 'text/csv;charset=utf-8' });

    case 'json':
      return new Blob([generateJSON(content, basename)], { type: 'application/json;charset=utf-8' });

    case 'xml':
      return new Blob([generateXML(content, basename)], { type: 'application/xml;charset=utf-8' });

    case 'xlsx': {
      const buffer = generateXLSXIndex(content);
      const arr = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
      return new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }

    case 'latex': {
      const processed = parseMarkdownToLaTeX(content);
      const tex = `\\documentclass{article}
\\usepackage[utf8]{inputenc}
\\usepackage{amsmath}
\\usepackage{hyperref}
\\title{${basename}}
\\begin{document}
\\maketitle
${processed}
\\end{document}`;
      return new Blob([tex], { type: 'application/x-latex;charset=utf-8' });
    }

    case 'rtf': {
      const rtfHeader = "{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\n" +
        "{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}{\\f1\\fnil\\fcharset0 Consolas;}{\\f2\\fnil\\fcharset0 Cambria Math;}}\n" +
        "{\\colortbl ;\\red0\\green0\\blue0;\\red102\\green102\\blue102;\\red240\\green240\\blue240;\\red79\\green70\\blue229;\\red229\\green231\\blue235;\\red204\\green204\\blue204;}\n" +
        "\\viewkind4\\uc1\\f0\\fs24 ";
      const rtfContent = parseMarkdownToRTF(content);
      return new Blob([rtfHeader + rtfContent + "}"], { type: 'application/rtf' });
    }

    case 'html': {
      const htmlContent = await markdownToRichHtml(content);
      const fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${basename}</title>
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
      return new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    }

    case 'docx': {
      const { elements, footnotes } = parseMarkdownToDocx(content);
      const docOptions: any = {
        styles: { paragraphStyles: [{ id: 'Normal', name: 'Normal', run: { font: 'Calibri', size: 24 }, paragraph: { spacing: { line: 276 } } }] },
        numbering: { config: [{ reference: 'default-numbering', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
        sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: elements }],
      };
      if (Object.keys(footnotes).length > 0) {
        docOptions.footnotes = footnotes;
      }
      const doc = new Document(docOptions);
      return await Packer.toBlob(doc);
    }

    case 'email': {
      const emailHtml = await markdownToEmailHtml(content);
      return new Blob([emailHtml], { type: 'text/html;charset=utf-8' });
    }

    case 'pdf': {
      const html = await markdownToRichHtml(content);
      const container = document.createElement('div');
      container.className = 'export-sandbox';
      const inner = document.createElement('div');
      inner.style.cssText = 'width:800px;background:#fff;padding:40px;font-family:system-ui,-apple-system,sans-serif;line-height:1.6;color:#1a1a1a;';
      inner.innerHTML = html;
      container.appendChild(inner);
      document.body.prepend(container);
      try {
        await new Promise(resolve => setTimeout(resolve, 500));
        const dataUrl = await toPng(container, {
          backgroundColor: '#ffffff',
          quality: 1,
          pixelRatio: 2,
          cacheBust: true,
        });
        const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
        const imgProps = pdf.getImageProperties(dataUrl);
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const imgWidth = pageWidth;
        const imgHeight = (imgProps.height * imgWidth) / imgProps.width;
        let heightLeft = imgHeight;
        let position = 0;
        let p = 0;
        pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
        while (heightLeft > 0) {
          p++;
          position = -(pageHeight * p);
          pdf.addPage();
          pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight);
          heightLeft -= pageHeight;
        }
        return new Blob([pdf.output('arraybuffer')], { type: 'application/pdf' });
      } finally {
        document.body.removeChild(container);
      }
    }

    // Platform formats
    case 'slack':
      return new Blob([markdownToSlack(content)], { type: 'text/plain;charset=utf-8' });
    case 'discord':
      return new Blob([markdownToDiscord(content)], { type: 'text/plain;charset=utf-8' });
    case 'jira':
      return new Blob([markdownToJira(content)], { type: 'text/plain;charset=utf-8' });
    case 'confluence':
      return new Blob([markdownToConfluence(content)], { type: 'text/plain;charset=utf-8' });
    case 'asciidoc':
      return new Blob([markdownToAsciiDoc(content)], { type: 'text/plain;charset=utf-8' });
    case 'rst':
      return new Blob([markdownToRST(content)], { type: 'text/plain;charset=utf-8' });
    case 'mediawiki':
      return new Blob([markdownToMediaWiki(content)], { type: 'text/plain;charset=utf-8' });
    case 'bbcode':
      return new Blob([markdownToBBCode(content)], { type: 'text/plain;charset=utf-8' });
    case 'textile':
      return new Blob([markdownToTextile(content)], { type: 'text/plain;charset=utf-8' });
    case 'orgmode':
      return new Blob([markdownToOrgMode(content)], { type: 'text/plain;charset=utf-8' });

    default:
      throw new Error(`Unsupported format: ${formatId}`);
  }
}

// ── Main batch execution ─────────────────────────────────────────────

export async function executeBatch(
  items: BatchItem[],
  formats: string[],
  onProgress: (progress: BatchProgress) => void,
  isCancelled?: () => boolean
): Promise<BatchResult[]> {
  const results: BatchResult[] = [];
  const total = items.length * formats.length;
  let completed = 0;

  for (const item of items) {
    for (const formatId of formats) {
      if (isCancelled?.()) return results;

      const format = BATCH_FORMATS.find(f => f.id === formatId);
      const formatLabel = format?.label || formatId;

      onProgress({
        total,
        completed,
        current: `Converting ${item.filename} → ${formatLabel}...`,
        results: [...results],
      });

      const basename = item.filename.replace(/\.[^.]+$/, '') || item.filename;
      const outputFilename = getOutputFilename(basename, formatId);

      try {
        const blob = await convertSingle(item.content, item.filename, formatId);
        results.push({
          itemId: item.id,
          filename: outputFilename,
          format: formatId,
          success: true,
          blob,
        });
      } catch (err) {
        results.push({
          itemId: item.id,
          filename: outputFilename,
          format: formatId,
          success: false,
          error: err instanceof Error ? err.message : String(err),
        });
      }

      completed++;
    }
  }

  onProgress({
    total,
    completed,
    current: 'Done!',
    results: [...results],
  });

  return results;
}

// ── ZIP packaging ────────────────────────────────────────────────────

export async function packageAsZip(results: BatchResult[]): Promise<Blob> {
  const zip = new JSZip();
  const successResults = results.filter(r => r.success && r.blob);
  const itemIds = [...new Set(successResults.map(r => r.itemId))];
  const singleFile = itemIds.length <= 1;

  for (const result of successResults) {
    if (!result.blob) continue;
    if (singleFile) {
      zip.file(result.filename, result.blob);
    } else {
      zip.file(`${result.itemId}/${result.filename}`, result.blob);
    }
  }

  return zip.generateAsync({ type: 'blob' });
}
