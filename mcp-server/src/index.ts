#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import remarkStringify from 'remark-stringify';
import remarkRehype from 'remark-rehype';
import rehypeKatex from 'rehype-katex';
import rehypeStringify from 'rehype-stringify';
import puppeteer from 'puppeteer-core';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import {
    stripMarkdown,
    parseMarkdownToRTF,
    parseMarkdownToDocx,
    parseMarkdownToLaTeX,
    generateCSV,
    generateJSON,
    generateXML,
    generateXLSXIndex,
    cleanMarkdownText
} from "./core-exports.js";
import { Packer } from "docx";

// Find Chrome/Chromium executable on the system
async function findChrome(): Promise<string> {
    // Docker / CI: honour explicit env var
    const envPath = process.env['PUPPETEER_EXECUTABLE_PATH'];
    if (envPath) {
        try { await fs.access(envPath); return envPath; } catch { /* fall through */ }
    }

    const platform = os.platform();
    const candidates: string[] = [];
    
    if (platform === 'win32') {
        const programFiles = process.env['PROGRAMFILES'] || 'C:\\Program Files';
        const programFilesX86 = process.env['PROGRAMFILES(X86)'] || 'C:\\Program Files (x86)';
        const localAppData = process.env['LOCALAPPDATA'] || '';
        candidates.push(
            path.join(programFiles, 'Google', 'Chrome', 'Application', 'chrome.exe'),
            path.join(programFilesX86, 'Google', 'Chrome', 'Application', 'chrome.exe'),
            path.join(localAppData, 'Google', 'Chrome', 'Application', 'chrome.exe'),
            path.join(programFiles, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
            path.join(programFilesX86, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
        );
    } else if (platform === 'darwin') {
        candidates.push(
            '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
            '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
            '/Applications/Chromium.app/Contents/MacOS/Chromium',
        );
    } else {
        candidates.push(
            '/usr/bin/google-chrome',
            '/usr/bin/google-chrome-stable',
            '/usr/bin/chromium',
            '/usr/bin/chromium-browser',
            '/snap/bin/chromium',
        );
    }

    for (const candidate of candidates) {
        try {
            await fs.access(candidate);
            return candidate;
        } catch { /* not found, try next */ }
    }
    throw new Error(
        'No Chrome/Chromium/Edge browser found. PDF and PNG export require a Chromium-based browser. ' +
        'Please install Google Chrome, Microsoft Edge, or Chromium.'
    );
}

const server = new Server(
    {
        name: "markdown-formatter-mcp",
        version: "1.0.0",
    },
    {
        capabilities: {
            tools: {},
        },
    }
);

// Binary format types that need special handling
const BINARY_FORMATS = ['docx', 'pdf', 'xlsx', 'png', 'image'] as const;
type BinaryFormat = typeof BINARY_FORMATS[number];

// Helper to handle output (save to file or return content)
async function handleOutput(
    content: string | Buffer,
    outputPath?: string,
    options?: { format?: string; sizeBytes?: number; description?: string }
): Promise<{ content: any[], isError?: boolean }> {
    if (outputPath) {
        try {
            await fs.mkdir(path.dirname(outputPath), { recursive: true });
            await fs.writeFile(outputPath, content);
            const stats = await fs.stat(outputPath);
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Successfully saved to ${outputPath}`,
                        file_path: outputPath,
                        file_size_bytes: stats.size,
                        format: options?.format || 'unknown'
                    }, null, 2)
                }]
            };
        } catch (err: any) {
            return { content: [{ type: "text", text: `Error saving to file: ${err.message}` }], isError: true };
        }
    }

    // For binary content without output_path, return helpful guidance
    if (Buffer.isBuffer(content)) {
        const sizeBytes = content.length;
        const format = options?.format || 'binary';

        // For AI usability, don't dump raw Base64 - provide actionable guidance
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    format: format,
                    file_size_bytes: sizeBytes,
                    description: options?.description || `Generated ${format.toUpperCase()} binary content`,
                    hint: `This is a binary file format. To save the file, call this tool again with the 'output_path' parameter specifying where to save it (e.g., "C:/Documents/output.${format}" or "./output.${format}").`,
                    base64_preview: content.toString('base64').substring(0, 100) + '...',
                    full_base64_length: content.toString('base64').length
                }, null, 2)
            }]
        };
    } else {
        return { content: [{ type: "text", text: content }] };
    }
}

// --- Shared parameter description constants ---
const PARAM_MARKDOWN = "The raw Markdown source text to convert. Supports GitHub-Flavored Markdown (tables, task lists, strikethrough) and KaTeX math expressions. Pass the full document content as a string, not a file path.";
const PARAM_OUTPUT_PATH_TEXT = "Optional. Absolute or relative file path (e.g. './output.txt') where the result will be saved. Parent directories are created automatically. If omitted, the converted text content is returned directly in the response as a string. If provided, the file is written to disk and a JSON summary with { success, file_path, file_size_bytes, format } is returned instead.";
const PARAM_OUTPUT_PATH_BINARY = (fmt: string) =>
    `Optional. Absolute or relative file path (e.g. './output.${fmt}') where the binary file will be saved. Parent directories are created automatically. If provided, the file is written to disk and a JSON summary with { success, file_path, file_size_bytes, format } is returned. If omitted, a JSON object with { format, file_size_bytes, hint, base64_preview } is returned — the hint will instruct you to call the tool again with output_path to save the file. Binary formats (${fmt.toUpperCase()}) should almost always specify output_path.`;
const PARAM_TITLE = "Optional. A document title string. Used as the root element name or document metadata title in the output. Defaults to 'document' if omitted.";

// Text-output tool annotations: no file write when output_path is omitted → read-only; with output_path → side effect
const TEXT_TOOL_ANNOTATIONS = {
    title: undefined as string | undefined,
    readOnlyHint: false,      // can write files when output_path is provided
    destructiveHint: false,    // overwrites files at output_path without warning
    idempotentHint: true,      // same input always produces the same output
    openWorldHint: false,      // does not interact with external services
};

// Binary-output tool annotations (PDF/PNG use Puppeteer which launches a browser)
const BROWSER_TOOL_ANNOTATIONS = {
    title: undefined as string | undefined,
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,      // Puppeteer runs a local headless browser, no network needed for rendering
};

server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: [
            {
                name: "harmonize_markdown",
                description:
                    "Standardize and normalize Markdown syntax without changing the document's meaning. " +
                    "Re-formats headers (ATX-style), normalizes list markers to '-', enforces fenced code blocks with backticks, " +
                    "and applies consistent indentation. " +
                    "Side effects: when output_path is provided, writes the harmonized Markdown to disk (creates parent directories as needed, overwrites existing files). " +
                    "When output_path is omitted, returns the harmonized text as a string with no file I/O. " +
                    "Returns: harmonized Markdown string (if no output_path), or JSON with { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this tool when you need to clean up inconsistent Markdown formatting before further processing. " +
                    "Prefer convert_to_md with harmonize=true if you also need to save the result, as it combines both steps. " +
                    "Not suitable for converting Markdown to other formats — use the convert_to_* tools instead.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Harmonize Markdown" },
            },
            {
                name: "convert_to_txt",
                description:
                    "Convert Markdown to plain text by stripping all formatting — removes headers, bold/italic markers, links, images, code fences, and HTML tags. " +
                    "The result is a human-readable plain-text string with no markup. This is a destructive conversion: formatting information is permanently lost. " +
                    "Side effects: when output_path is provided, writes the plain text to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the plain text string directly. " +
                    "Returns: plain text string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this instead of convert_to_md when you need formatting-free content (e.g. for indexing, search, or clipboard). " +
                    "Use convert_to_html or convert_to_pdf if you need to preserve the document's visual structure.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Plain Text" },
            },
            {
                name: "convert_to_rtf",
                description:
                    "Convert Markdown to Rich Text Format (RTF). Produces an RTF document string preserving basic formatting: " +
                    "bold, italic, headers (as styled paragraphs), lists, and code blocks. " +
                    "Side effects: when output_path is provided, writes the RTF file to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the raw RTF markup as a string. " +
                    "Returns: RTF markup string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this when the target application requires RTF (e.g. legacy word processors, email clients). " +
                    "Prefer convert_to_docx for modern Word documents, or convert_to_html for web display.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to RTF" },
            },
            {
                name: "convert_to_latex",
                description:
                    "Convert Markdown to LaTeX source code. Produces a LaTeX document fragment with \\section, \\textbf, \\textit, " +
                    "\\begin{itemize}/\\begin{enumerate} list environments, verbatim code blocks, and table environments. " +
                    "KaTeX math expressions in the Markdown are passed through as native LaTeX math. " +
                    "Side effects: when output_path is provided, writes the .tex file to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the LaTeX source as a string. " +
                    "Returns: LaTeX source string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this when you need to embed content in a LaTeX workflow or compile to PDF via pdflatex/xelatex externally. " +
                    "For direct PDF output without a LaTeX toolchain, use convert_to_pdf instead.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to LaTeX" },
            },
            {
                name: "convert_to_docx",
                description:
                    "Convert Markdown to a Microsoft Word DOCX file. Produces a binary .docx document with styled headings, " +
                    "bold/italic text, numbered and bulleted lists, and code formatting. " +
                    "This is a binary format — output_path should almost always be provided. " +
                    "Side effects: when output_path is provided, writes the DOCX binary to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns a JSON object with { format: 'docx', file_size_bytes, hint, base64_preview } — " +
                    "the hint will tell you to re-call with output_path to save the file. " +
                    "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted). " +
                    "Use this for Word-compatible documents. " +
                    "Prefer convert_to_rtf for legacy word processors, convert_to_pdf for read-only distribution, or convert_to_html for web.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_BINARY("docx") },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to DOCX" },
            },
            {
                name: "convert_to_pdf",
                description:
                    "Convert Markdown to a PDF document. Renders the Markdown as styled HTML (GFM tables, KaTeX math) and then " +
                    "prints it to PDF via a headless Chromium browser (Puppeteer). Requires a locally installed Chrome, Edge, or Chromium — " +
                    "set PUPPETEER_EXECUTABLE_PATH env var to override auto-detection. " +
                    "This is a binary format — output_path should almost always be provided. " +
                    "Side effects: launches a transient headless browser process for rendering (no network requests are made for the conversion itself, " +
                    "though the HTML references a CDN KaTeX stylesheet which may be fetched). " +
                    "When output_path is provided, writes the PDF to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns JSON { format: 'pdf', file_size_bytes, hint, base64_preview }. " +
                    "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted). " +
                    "Use this for high-fidelity, print-ready document output. " +
                    "Prefer convert_to_html for web-viewable output, convert_to_docx for editable documents, or convert_to_latex for LaTeX toolchains.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_BINARY("pdf") },
                    },
                    required: ["markdown"],
                },
                annotations: { ...BROWSER_TOOL_ANNOTATIONS, title: "Convert to PDF" },
            },
            {
                name: "convert_to_image",
                description:
                    "Convert Markdown to a PNG image. Renders the Markdown as styled HTML (GFM tables, KaTeX math) and takes a " +
                    "full-page screenshot via a headless Chromium browser (Puppeteer). Requires a locally installed Chrome, Edge, or Chromium — " +
                    "set PUPPETEER_EXECUTABLE_PATH env var to override auto-detection. " +
                    "This is a binary format — output_path should almost always be provided. " +
                    "Side effects: launches a transient headless browser process (no persistent state; may fetch KaTeX CDN stylesheet). " +
                    "When output_path is provided, writes the PNG to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns JSON { format: 'png', file_size_bytes, hint, base64_preview }. " +
                    "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted). " +
                    "Use this when you need a visual snapshot of the rendered Markdown (e.g. for embedding in chat, previews, social cards). " +
                    "Prefer convert_to_pdf for paginated print output, or convert_to_html for interactive web content.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_BINARY("png") },
                    },
                    required: ["markdown"],
                },
                annotations: { ...BROWSER_TOOL_ANNOTATIONS, title: "Convert to PNG Image" },
            },
            {
                name: "convert_to_csv",
                description:
                    "Extract tables from Markdown and convert them to CSV format. Parses GFM pipe-tables from the input and outputs " +
                    "comma-separated values. If the Markdown contains multiple tables, they are concatenated with a blank line separator. " +
                    "Non-table content is ignored. If the Markdown contains no tables, returns an empty string. " +
                    "Side effects: when output_path is provided, writes the CSV to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the CSV text directly as a string. " +
                    "Returns: CSV text string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this for lightweight tabular export or when downstream tools expect CSV. " +
                    "Prefer convert_to_xlsx for Excel-compatible spreadsheets with multiple sheets, or convert_to_json for structured data.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to CSV" },
            },
            {
                name: "convert_to_json",
                description:
                    "Convert Markdown to a structured JSON representation. Parses the document into a JSON object with the document title " +
                    "as the root key, containing arrays of section objects with headings, paragraphs, lists, code blocks, and tables. " +
                    "Useful for programmatic analysis or feeding structured content into other systems. " +
                    "Side effects: when output_path is provided, writes the JSON to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the JSON string directly. " +
                    "Returns: JSON string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this when you need a machine-readable AST-like representation of the Markdown content. " +
                    "Prefer convert_to_xml for XML-based interchange, or convert_to_csv/convert_to_xlsx for tabular data extraction.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        title: { type: "string", description: PARAM_TITLE },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to JSON" },
            },
            {
                name: "convert_to_xml",
                description:
                    "Convert Markdown to an XML document. Parses the Markdown into a structured XML tree with a root element named after " +
                    "the title parameter, containing <section>, <heading>, <paragraph>, <list>, <code>, and <table> elements. " +
                    "Produces well-formed XML with an <?xml?> declaration. " +
                    "Side effects: when output_path is provided, writes the XML to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the XML string directly. " +
                    "Returns: XML string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this for XML-based data interchange or when downstream systems require XML input. " +
                    "Prefer convert_to_json for JSON APIs, convert_to_html for XHTML/web content, or convert_to_csv for flat tabular data.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        title: { type: "string", description: "Optional. The root XML element name and document title. Must be a valid XML element name (no spaces or special characters). Defaults to 'document' if omitted." },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to XML" },
            },
            {
                name: "convert_to_xlsx",
                description:
                    "Convert Markdown tables to a Microsoft Excel XLSX spreadsheet. Parses GFM pipe-tables from the input " +
                    "and creates an Excel workbook. Each table becomes a sheet in the workbook. Non-table content is ignored. " +
                    "If the Markdown contains no tables, produces an empty workbook. " +
                    "This is a binary format — output_path should almost always be provided. " +
                    "Side effects: when output_path is provided, writes the XLSX binary to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns JSON { format: 'xlsx', file_size_bytes, hint, base64_preview }. " +
                    "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted). " +
                    "Use this when you need a full Excel file with formatting. " +
                    "Prefer convert_to_csv for lightweight plain-text tabular export, or convert_to_json for structured programmatic access.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_BINARY("xlsx") },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to XLSX" },
            },
            {
                name: "convert_to_html",
                description:
                    "Convert Markdown to a complete, styled HTML document. Renders GFM (tables, task lists, strikethrough) and " +
                    "KaTeX math into semantic HTML with an embedded stylesheet for clean presentation. " +
                    "The output is a full <!DOCTYPE html> document with <head> (charset, KaTeX CSS CDN link, inline styles) and <body>. " +
                    "Side effects: when output_path is provided, writes the HTML file to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the full HTML string directly. " +
                    "Returns: HTML document string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this when you need a file saved to disk or when you need the full document. " +
                    "Prefer generate_html if you only need the HTML string returned directly (no file I/O) and want inline styles without a CDN link. " +
                    "Prefer convert_to_pdf for print-ready output, or convert_to_image for a visual snapshot.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to HTML" },
            },
            {
                name: "convert_to_md",
                description:
                    "Export Markdown content, optionally harmonizing its formatting first. When harmonize=false (default), " +
                    "returns the input Markdown unchanged. When harmonize=true, applies the same normalization as harmonize_markdown " +
                    "(ATX-style headers, '-' list markers, fenced code blocks, consistent indentation) before returning. " +
                    "Side effects: when output_path is provided, writes the Markdown to disk (creates parent directories, overwrites existing files). " +
                    "When output_path is omitted, returns the Markdown string directly. " +
                    "Returns: Markdown string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set). " +
                    "Use this when you want to save Markdown to a file (with or without cleanup). " +
                    "Prefer harmonize_markdown if you only want to normalize formatting without saving to disk. " +
                    "Use the convert_to_* family for other output formats.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        harmonize: { type: "boolean", description: "Optional. When true, normalizes Markdown syntax (ATX headers, '-' list markers, fenced code blocks, consistent indentation) before returning or saving. When false or omitted, the Markdown is passed through unchanged. Defaults to false." },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Export Markdown" },
            },
            {
                name: "generate_html",
                description:
                    "Generate a complete, self-contained HTML document from Markdown with all styles inlined. " +
                    "Renders GFM (tables, task lists, strikethrough) and KaTeX math into a full HTML page with an embedded <style> block " +
                    "and a KaTeX CSS CDN link. Returns the HTML string directly — no file is written to disk. " +
                    "Side effects: none. This tool is read-only and performs no file I/O. " +
                    "Returns: a complete HTML document string (<!DOCTYPE html>…</html>) with inline styles, ready for rendering in a browser. " +
                    "The optional title parameter sets the <title> tag in the HTML <head> section. " +
                    "Use this when you need styled HTML output returned as a string (e.g., for embedding in responses or previewing). " +
                    "Prefer convert_to_html when you need to write the HTML to a file on disk. " +
                    "Prefer convert_to_pdf or convert_to_image for non-HTML visual output formats.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        title: { type: "string", description: "Optional. Sets the <title> tag in the HTML document's <head> section. Displayed in browser tabs and bookmarks. Defaults to 'Document' if omitted." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Generate HTML Document",
                    readOnlyHint: true,    // never writes to disk
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            }
        ],
    };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
        const { name, arguments: args } = request.params;
        const markdown = (args as any).markdown;
        const outputPath = (args as any).output_path;

        if (!markdown && name !== 'list_tools') {
            // Basic validation
            throw new Error("Markdown content is required");
        }

        if (name === "harmonize_markdown") {
            const file = await unified()
                .use(remarkParse)
                .use(remarkGfm)
                .use(remarkMath)
                .use(remarkStringify, {
                    bullet: '-',
                    fence: '`',
                    fences: true,
                    incrementListMarker: true,
                    listItemIndent: 'one',
                })
                .process(markdown);
            return handleOutput(String(file), outputPath);
        }

        if (name === "convert_to_txt") {
            const txt = cleanMarkdownText(markdown);
            return handleOutput(txt, outputPath);
        }

        if (name === "convert_to_rtf") {
            const rtf = parseMarkdownToRTF(markdown);
            return handleOutput(rtf, outputPath);
        }

        if (name === "convert_to_latex") {
            const latex = parseMarkdownToLaTeX(markdown);
            return handleOutput(latex, outputPath);
        }

        if (name === "convert_to_docx") {
            const elements = parseMarkdownToDocx(markdown);
            const doc = new ((await import("docx")).Document)({
                sections: [{ children: elements }]
            });
            const buffer = await Packer.toBuffer(doc);
            return handleOutput(buffer, outputPath, {
                format: 'docx',
                description: 'Microsoft Word document generated from Markdown'
            });
        }

        if (name === "convert_to_csv") {
            const csv = generateCSV(markdown);
            return handleOutput(csv, outputPath);
        }

        if (name === "convert_to_json") {
            const title = (args as any).title || "document";
            const json = generateJSON(markdown, title);
            return handleOutput(json, outputPath);
        }

        if (name === "convert_to_xml") {
            const title = (args as any).title || "document";
            const xml = generateXML(markdown, title);
            return handleOutput(xml, outputPath);
        }

        if (name === "convert_to_xlsx") {
            const buffer = generateXLSXIndex(markdown);
            return handleOutput(buffer, outputPath, {
                format: 'xlsx',
                description: 'Microsoft Excel spreadsheet generated from Markdown tables'
            });
        }

        if (name === "convert_to_html" || name === "convert_to_pdf" || name === "convert_to_image") {
            const htmlProcessor = unified()
                .use(remarkParse)
                .use(remarkGfm)
                // @ts-ignore
                .use(remarkRehype)
                // @ts-ignore
                .use(rehypeKatex)
                // @ts-ignore
                .use(rehypeStringify);

            const htmlFile = await htmlProcessor.process(markdown);

            const htmlDoc = `<!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous">
            <style>
                body { font-family: system-ui, -apple-system, sans-serif; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto; background: white; color: black; }
                img { max-width: 100%; }
                pre { background: #f4f4f4; padding: 15px; border-radius: 5px; overflow-x: auto; }
                table { border-collapse: collapse; width: 100%; margin: 1em 0; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #f2f2f2; }
            </style>
        </head>
        <body>${String(htmlFile)}</body>
        </html>`;

            if (name === "convert_to_html") {
                return handleOutput(htmlDoc, outputPath);
            }

            const browser = await puppeteer.launch({ headless: true, executablePath: await findChrome() });
            const page = await browser.newPage();
            await page.setContent(htmlDoc);

            let resultBuffer: Buffer;

            if (name === "convert_to_pdf") {
                resultBuffer = await page.pdf({ format: 'A4' }) as Buffer;
                await browser.close();
                return handleOutput(resultBuffer, outputPath, {
                    format: 'pdf',
                    description: 'PDF document generated from Markdown via Puppeteer'
                });
            } else {
                const screenshot = await page.screenshot({ fullPage: true, encoding: 'binary' });
                resultBuffer = screenshot as Buffer;
                await browser.close();
                return handleOutput(resultBuffer, outputPath, {
                    format: 'png',
                    description: 'PNG image screenshot of the rendered Markdown'
                });
            }
        }

        // New tools: convert_to_md and generate_html
        if (name === "convert_to_md") {
            const shouldHarmonize = (args as any).harmonize;
            let result = markdown;
            if (shouldHarmonize) {
                const file = await unified()
                    .use(remarkParse)
                    .use(remarkGfm)
                    .use(remarkMath)
                    .use(remarkStringify, {
                        bullet: '-',
                        fence: '`',
                        fences: true,
                        incrementListMarker: true,
                        listItemIndent: 'one',
                    })
                    .process(markdown);
                result = String(file);
            }
            return handleOutput(result, outputPath);
        }

        if (name === "generate_html") {
            const title = (args as any).title || 'Document';
            const htmlProcessor = unified()
                .use(remarkParse)
                .use(remarkGfm)
                // @ts-ignore
                .use(remarkRehype)
                // @ts-ignore
                .use(rehypeKatex)
                // @ts-ignore
                .use(rehypeStringify);

            const htmlFile = await htmlProcessor.process(markdown);

            const htmlDoc = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title}</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
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
<body>${String(htmlFile)}</body>
</html>`;
            return { content: [{ type: "text", text: htmlDoc }] };
        }

        throw new Error(`Unknown tool: ${name}`);
    } catch (error: any) {
        return {
            content: [{ type: "text", text: `Error: ${error.message}` }],
            isError: true,
        };
    }
});

const transport = new StdioServerTransport();
await server.connect(transport);
