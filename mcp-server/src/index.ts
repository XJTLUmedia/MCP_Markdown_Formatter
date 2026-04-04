#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
    CallToolRequestSchema,
    ListToolsRequestSchema,
    ListPromptsRequestSchema,
    GetPromptRequestSchema,
    ListResourcesRequestSchema,
    ListResourceTemplatesRequestSchema,
    ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
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
} from "./platform-converters.js";
import {
    repairMarkdown,
    lintMarkdown,
} from "./markdown-repair.js";
import {
    extractCodeBlocks,
    extractLinks,
    generateTOC,
    analyzeDocument,
    extractStructure,
} from "./document-analysis.js";
import { htmlToMarkdown } from "./html-import.js";
import { markdownToEmailHtml } from "./email-html.js";

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
            prompts: {},
            resources: {},
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
            },
            // ── Platform-specific format tools ──
            {
                name: "convert_to_slack",
                description:
                    "Convert Markdown to Slack mrkdwn format. Transforms bold (**) to single asterisks, italic to underscores, " +
                    "links to Slack <url|text> syntax, and headers to bold text. " +
                    "Use this when pasting formatted content into Slack messages.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Slack mrkdwn" },
            },
            {
                name: "convert_to_discord",
                description:
                    "Convert Markdown to Discord-compatible format. Transforms headers to styled bold/underline text that renders " +
                    "correctly in Discord messages. Code blocks and basic formatting are preserved.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Discord Markdown" },
            },
            {
                name: "convert_to_jira",
                description:
                    "Convert Markdown to JIRA wiki markup. Transforms headers to h1./h2., bold to single asterisks, " +
                    "code blocks to {code} blocks, links to [text|url], and lists to JIRA * and # syntax.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to JIRA Markup" },
            },
            {
                name: "convert_to_confluence",
                description:
                    "Convert Markdown to Confluence wiki markup. Similar to JIRA but includes Confluence-specific {info}, {note} panels " +
                    "and {code:language=x} syntax.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Confluence Markup" },
            },
            {
                name: "convert_to_asciidoc",
                description:
                    "Convert Markdown to AsciiDoc format. Transforms headers to = syntax, code blocks to ---- delimited blocks, " +
                    "links to url[text] syntax, and images to image::url[alt] directives.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to AsciiDoc" },
            },
            {
                name: "convert_to_rst",
                description:
                    "Convert Markdown to reStructuredText (RST) format. Transforms headers to underlined text, " +
                    "code blocks to .. code-block:: directives, and links to RST reference syntax.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to reStructuredText" },
            },
            {
                name: "convert_to_mediawiki",
                description:
                    "Convert Markdown to MediaWiki markup. Transforms headers to == syntax, bold to triple quotes, " +
                    "code to <syntaxhighlight> tags, and tables to {| wikitable format.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to MediaWiki" },
            },
            {
                name: "convert_to_bbcode",
                description:
                    "Convert Markdown to BBCode format. Transforms formatting to [b], [i], [s], [code], [url], [img] tags. " +
                    "Used for forum posts on phpBB, vBulletin, and similar platforms.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to BBCode" },
            },
            {
                name: "convert_to_textile",
                description:
                    "Convert Markdown to Textile markup format. Used by Redmine, older versions of Basecamp, and some CMS platforms.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Textile" },
            },
            {
                name: "convert_to_orgmode",
                description:
                    "Convert Markdown to Emacs Org Mode format. Transforms headers to * syntax, bold to *text*, " +
                    "code blocks to #+BEGIN_SRC/#+END_SRC, and links to [[url][text]] syntax.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Org Mode" },
            },
            {
                name: "convert_to_email_html",
                description:
                    "Convert Markdown to email-optimized HTML with all styles inlined. Produces HTML compatible with " +
                    "Outlook, Gmail, Apple Mail, and other email clients. No external CSS dependencies. " +
                    "Wraps content in a responsive email table layout.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: PARAM_MARKDOWN },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Email HTML" },
            },
            // ── Import tools ──
            {
                name: "html_to_markdown",
                description:
                    "Convert HTML to Markdown. Performs round-trip import of HTML content back to Markdown format. " +
                    "Handles headings, tables, lists, code blocks, links, images, and inline formatting. " +
                    "Useful for importing web content or converting HTML emails to Markdown.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        html: { type: "string", description: "The HTML content to convert to Markdown. Can be a full HTML document or a fragment." },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["html"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Import HTML to Markdown" },
            },
            // ── Repair / Lint tools ──
            {
                name: "repair_markdown",
                description:
                    "Repair broken Markdown from LLM output or copy-paste. Fixes unclosed code fences, broken tables " +
                    "(mismatched columns, missing separators), stray emphasis markers, missing heading spaces, " +
                    "inconsistent list indentation, broken links, and excessive whitespace.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The potentially broken Markdown text to repair." },
                        output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT },
                    },
                    required: ["markdown"],
                },
                annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Repair Broken Markdown" },
            },
            {
                name: "lint_markdown",
                description:
                    "Lint Markdown and report issues. Returns a JSON array of lint issues found in the document, " +
                    "each with line number, column, severity (error/warning/info), rule name, message, and fixable flag. " +
                    "Checks for: missing heading spaces, trailing whitespace, inconsistent list markers, hard tabs, " +
                    "multiple blank lines, bare URLs, unclosed emphasis, and unclosed code fences.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The Markdown text to lint." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Lint Markdown",
                    readOnlyHint: true,
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            },
            // ── Analysis tools ──
            {
                name: "extract_code_blocks",
                description:
                    "Extract all code blocks from a Markdown document. Returns a JSON array of code blocks, " +
                    "each with language, code content, and start/end line numbers. " +
                    "Useful for extracting code snippets from LLM responses or documentation.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The Markdown text to extract code blocks from." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Extract Code Blocks",
                    readOnlyHint: true,
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            },
            {
                name: "extract_links",
                description:
                    "Extract all links and images from a Markdown document. Returns a JSON array with link text, URL, " +
                    "line number, and type (inline, reference, image, autolink). " +
                    "Useful for link checking, SEO analysis, or extracting references.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The Markdown text to extract links from." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Extract Links",
                    readOnlyHint: true,
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            },
            {
                name: "generate_toc",
                description:
                    "Generate a Table of Contents from Markdown headings. Returns a Markdown-formatted TOC with " +
                    "indented links to each heading. Handles duplicate heading slugs. " +
                    "The max_depth parameter controls the deepest heading level to include.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The Markdown text to generate a TOC from." },
                        max_depth: { type: "number", description: "Maximum heading depth to include (1-6, default: 6)." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Generate Table of Contents",
                    readOnlyHint: true,
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            },
            {
                name: "analyze_document",
                description:
                    "Analyze a Markdown document and return comprehensive statistics. Returns JSON with: " +
                    "line/word/character/paragraph/sentence counts, heading/code block/table/link/image/list/blockquote counts, " +
                    "and estimated reading time in minutes.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The Markdown text to analyze." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Analyze Document Statistics",
                    readOnlyHint: true,
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            },
            {
                name: "extract_structure",
                description:
                    "Extract the full structure of a Markdown document. Returns JSON with document statistics, " +
                    "heading outline, code block summary (language, line count, positions), and link summary " +
                    "(totals by type, unique URL count). Provides a bird's-eye view of document architecture.",
                inputSchema: {
                    type: "object" as const,
                    properties: {
                        markdown: { type: "string", description: "The Markdown text to extract structure from." },
                    },
                    required: ["markdown"],
                },
                annotations: {
                    title: "Extract Document Structure",
                    readOnlyHint: true,
                    destructiveHint: false,
                    idempotentHint: true,
                    openWorldHint: false,
                },
            },
        ],
    };
});

// ── Prompts ──────────────────────────────────────────────────────────
server.setRequestHandler(ListPromptsRequestSchema, async () => {
    return {
        prompts: [
            {
                name: "convert-document",
                description:
                    "Convert a Markdown document to a specified output format. " +
                    "Supports: PDF, DOCX, HTML, LaTeX, CSV, JSON, XML, XLSX, RTF, PNG, TXT, MD.",
                arguments: [
                    {
                        name: "format",
                        description:
                            "Target output format: pdf, docx, html, latex, csv, json, xml, xlsx, rtf, png, txt, or md",
                        required: true,
                    },
                    {
                        name: "markdown",
                        description: "The Markdown content to convert",
                        required: true,
                    },
                ],
            },
            {
                name: "extract-tables",
                description:
                    "Extract all tables from a Markdown document and export them as CSV or XLSX spreadsheet format.",
                arguments: [
                    {
                        name: "format",
                        description:
                            "Output format for tables: 'csv' for plain text or 'xlsx' for Excel spreadsheet",
                        required: true,
                    },
                    {
                        name: "markdown",
                        description:
                            "The Markdown content containing tables to extract",
                        required: true,
                    },
                ],
            },
            {
                name: "format-for-sharing",
                description:
                    "Prepare a Markdown document for sharing by harmonizing formatting and converting to " +
                    "portable formats (PDF and HTML) with professional styling.",
                arguments: [
                    {
                        name: "markdown",
                        description: "The Markdown content to format for sharing",
                        required: true,
                    },
                ],
            },
            {
                name: "analyze-and-repair",
                description:
                    "Analyze a Markdown document for issues, repair any problems found, and return " +
                    "both the lint report and the repaired document.",
                arguments: [
                    {
                        name: "markdown",
                        description: "The Markdown content to analyze and repair",
                        required: true,
                    },
                ],
            },
            {
                name: "convert-for-platform",
                description:
                    "Convert Markdown to a platform-specific format. " +
                    "Supports: slack, discord, jira, confluence, asciidoc, rst, mediawiki, bbcode, textile, orgmode.",
                arguments: [
                    {
                        name: "platform",
                        description:
                            "Target platform: slack, discord, jira, confluence, asciidoc, rst, mediawiki, bbcode, textile, or orgmode",
                        required: true,
                    },
                    {
                        name: "markdown",
                        description: "The Markdown content to convert",
                        required: true,
                    },
                ],
            },
            {
                name: "document-overview",
                description:
                    "Get a comprehensive overview of a Markdown document: statistics, structure, " +
                    "table of contents, code blocks, and links.",
                arguments: [
                    {
                        name: "markdown",
                        description: "The Markdown content to analyze",
                        required: true,
                    },
                ],
            },
        ],
    };
});

server.setRequestHandler(GetPromptRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    if (name === "convert-document") {
        const format = args?.format || "pdf";
        const markdown = args?.markdown || "";
        return {
            description: `Convert Markdown to ${format.toUpperCase()}`,
            messages: [
                {
                    role: "user" as const,
                    content: {
                        type: "text" as const,
                        text: `Please convert the following Markdown document to ${format.toUpperCase()} format using the convert_to_${format} tool.\n\n${markdown}`,
                    },
                },
            ],
        };
    }

    if (name === "extract-tables") {
        const format = args?.format || "csv";
        const markdown = args?.markdown || "";
        return {
            description: `Extract tables from Markdown as ${format.toUpperCase()}`,
            messages: [
                {
                    role: "user" as const,
                    content: {
                        type: "text" as const,
                        text: `Please extract all tables from the following Markdown and convert them to ${format.toUpperCase()} format using the convert_to_${format} tool.\n\n${markdown}`,
                    },
                },
            ],
        };
    }

    if (name === "format-for-sharing") {
        const markdown = args?.markdown || "";
        return {
            description: "Format Markdown for professional sharing",
            messages: [
                {
                    role: "user" as const,
                    content: {
                        type: "text" as const,
                        text: `Please format the following Markdown for sharing. First, use harmonize_markdown to clean up the formatting, then convert it to both PDF (using convert_to_pdf) and HTML (using convert_to_html) for distribution.\n\n${markdown}`,
                    },
                },
            ],
        };
    }

    if (name === "analyze-and-repair") {
        const markdown = args?.markdown || "";
        return {
            description: "Analyze and repair Markdown document",
            messages: [
                {
                    role: "user" as const,
                    content: {
                        type: "text" as const,
                        text: `Please analyze and repair the following Markdown document:\n\n1. First, use lint_markdown to identify issues in the document.\n2. Then, use repair_markdown to fix the problems.\n3. Finally, use lint_markdown again on the repaired version to confirm issues are resolved.\n\nReturn both the lint report and the repaired Markdown.\n\n${markdown}`,
                    },
                },
            ],
        };
    }

    if (name === "convert-for-platform") {
        const platform = args?.platform || "slack";
        const markdown = args?.markdown || "";
        return {
            description: `Convert Markdown for ${platform}`,
            messages: [
                {
                    role: "user" as const,
                    content: {
                        type: "text" as const,
                        text: `Please convert the following Markdown to ${platform} format using the convert_to_${platform} tool.\n\n${markdown}`,
                    },
                },
            ],
        };
    }

    if (name === "document-overview") {
        const markdown = args?.markdown || "";
        return {
            description: "Comprehensive document overview",
            messages: [
                {
                    role: "user" as const,
                    content: {
                        type: "text" as const,
                        text: `Please provide a comprehensive overview of the following Markdown document:\n\n1. Use analyze_document to get statistics (word count, reading time, etc.)\n2. Use generate_toc to create a table of contents\n3. Use extract_code_blocks to list all code snippets\n4. Use extract_links to catalog all links\n\nSummarize the findings in a clear report.\n\n${markdown}`,
                    },
                },
            ],
        };
    }

    throw new Error(`Unknown prompt: ${name}`);
});

// ── Resources ────────────────────────────────────────────────────────
server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => {
    return { resourceTemplates: [] };
});

server.setRequestHandler(ListResourcesRequestSchema, async () => {
    return {
        resources: [
            {
                uri: "markdown-formatter://supported-formats",
                name: "Supported Output Formats",
                description:
                    "Complete list of all 22+ supported output formats with tool names, types, and descriptions",
                mimeType: "application/json",
            },
            {
                uri: "markdown-formatter://conversion-guide",
                name: "Conversion Guide",
                description:
                    "Guide for choosing the right output format based on your use case",
                mimeType: "text/plain",
            },
        ],
    };
});

server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    const { uri } = request.params;

    if (uri === "markdown-formatter://supported-formats") {
        return {
            contents: [
                {
                    uri,
                    mimeType: "application/json",
                    text: JSON.stringify(
                        {
                            formats: [
                                { id: "md", name: "Markdown", tool: "convert_to_md", type: "text", description: "Export or harmonize Markdown" },
                                { id: "txt", name: "Plain Text", tool: "convert_to_txt", type: "text", description: "Strip all formatting" },
                                { id: "html", name: "HTML", tool: "convert_to_html", type: "text", description: "Styled HTML document" },
                                { id: "email_html", name: "Email HTML", tool: "convert_to_email_html", type: "text", description: "Email-optimized HTML with inlined styles" },
                                { id: "pdf", name: "PDF", tool: "convert_to_pdf", type: "binary", description: "Print-ready PDF (requires Chromium)" },
                                { id: "docx", name: "Word DOCX", tool: "convert_to_docx", type: "binary", description: "Microsoft Word document" },
                                { id: "rtf", name: "Rich Text", tool: "convert_to_rtf", type: "text", description: "RTF for legacy word processors" },
                                { id: "latex", name: "LaTeX", tool: "convert_to_latex", type: "text", description: "LaTeX source code" },
                                { id: "csv", name: "CSV", tool: "convert_to_csv", type: "text", description: "Tables to comma-separated values" },
                                { id: "json", name: "JSON", tool: "convert_to_json", type: "text", description: "Structured JSON representation" },
                                { id: "xml", name: "XML", tool: "convert_to_xml", type: "text", description: "XML document" },
                                { id: "xlsx", name: "Excel XLSX", tool: "convert_to_xlsx", type: "binary", description: "Excel spreadsheet from tables" },
                                { id: "png", name: "PNG Image", tool: "convert_to_image", type: "binary", description: "Screenshot image (requires Chromium)" },
                                { id: "slack", name: "Slack mrkdwn", tool: "convert_to_slack", type: "text", description: "Slack messaging format" },
                                { id: "discord", name: "Discord Markdown", tool: "convert_to_discord", type: "text", description: "Discord-compatible formatting" },
                                { id: "jira", name: "JIRA Markup", tool: "convert_to_jira", type: "text", description: "Atlassian JIRA wiki markup" },
                                { id: "confluence", name: "Confluence Markup", tool: "convert_to_confluence", type: "text", description: "Atlassian Confluence wiki markup" },
                                { id: "asciidoc", name: "AsciiDoc", tool: "convert_to_asciidoc", type: "text", description: "AsciiDoc lightweight markup" },
                                { id: "rst", name: "reStructuredText", tool: "convert_to_rst", type: "text", description: "Python/Sphinx documentation format" },
                                { id: "mediawiki", name: "MediaWiki", tool: "convert_to_mediawiki", type: "text", description: "Wikipedia/MediaWiki markup" },
                                { id: "bbcode", name: "BBCode", tool: "convert_to_bbcode", type: "text", description: "Forum posting format" },
                                { id: "textile", name: "Textile", tool: "convert_to_textile", type: "text", description: "Textile markup (Redmine)" },
                                { id: "orgmode", name: "Org Mode", tool: "convert_to_orgmode", type: "text", description: "Emacs Org Mode format" },
                            ],
                            total: 23,
                            special_tools: [
                                { name: "harmonize_markdown", description: "Normalize Markdown formatting without changing content" },
                                { name: "generate_html", description: "Generate HTML string without file I/O (read-only)" },
                            ],
                            analysis_tools: [
                                { name: "extract_code_blocks", description: "Extract all code blocks with language and line info" },
                                { name: "extract_links", description: "Extract all links and images with type classification" },
                                { name: "generate_toc", description: "Generate a Markdown Table of Contents from headings" },
                                { name: "analyze_document", description: "Comprehensive document statistics (words, lines, reading time)" },
                                { name: "extract_structure", description: "Full document structure overview (stats, outline, summaries)" },
                            ],
                            repair_tools: [
                                { name: "repair_markdown", description: "Fix broken Markdown from LLM output or copy-paste" },
                                { name: "lint_markdown", description: "Lint Markdown and report issues as JSON" },
                            ],
                            import_tools: [
                                { name: "html_to_markdown", description: "Convert HTML back to Markdown" },
                            ],
                        },
                        null,
                        2
                    ),
                },
            ],
        };
    }

    if (uri === "markdown-formatter://conversion-guide") {
        return {
            contents: [
                {
                    uri,
                    mimeType: "text/plain",
                    text: [
                        "Markdown Formatter — Conversion Guide",
                        "======================================",
                        "",
                        "Choose a format based on your use case:",
                        "",
                        "For sharing documents:",
                        "  - PDF  — Best for print-ready, read-only distribution",
                        "  - DOCX — Best for editable documents in Microsoft Word",
                        "  - HTML — Best for web viewing and embedding",
                        "  - Email HTML — For email campaigns with inlined styles (Outlook, Gmail)",
                        "",
                        "For data extraction:",
                        "  - CSV  — Lightweight tabular data from Markdown tables",
                        "  - XLSX — Full Excel workbook with formatted tables",
                        "  - JSON — Machine-readable structured representation",
                        "  - XML  — XML-based data interchange",
                        "",
                        "For visual output:",
                        "  - PNG  — Screenshot image for chat or social media",
                        "  - PDF  — Paginated visual output for printing",
                        "",
                        "For text processing:",
                        "  - TXT  — Plain text with all formatting stripped",
                        "  - MD   — Clean/harmonized Markdown",
                        "  - LaTeX — For academic papers and typesetting",
                        "  - RTF  — For legacy word processors and email clients",
                        "",
                        "For platform-specific sharing:",
                        "  - Slack      — Paste into Slack messages (*bold*, _italic_, <url|text>)",
                        "  - Discord    — Discord-compatible markdown (styled headers, code blocks)",
                        "  - JIRA       — Atlassian JIRA ticket descriptions and comments",
                        "  - Confluence — Atlassian Confluence wiki pages",
                        "  - MediaWiki  — Wikipedia and MediaWiki-based wikis",
                        "  - BBCode     — Forum posts (phpBB, vBulletin, etc.)",
                        "  - Textile    — Redmine and some CMS platforms",
                        "  - Org Mode   — Emacs Org Mode files",
                        "",
                        "For documentation systems:",
                        "  - AsciiDoc — Alternative lightweight markup (used by Antora, etc.)",
                        "  - RST      — reStructuredText for Sphinx/Python documentation",
                        "",
                        "For document analysis:",
                        "  - extract_code_blocks — Pull code snippets from docs or LLM output",
                        "  - extract_links — Get all URLs for link checking or SEO analysis",
                        "  - generate_toc — Auto-generate Table of Contents from headings",
                        "  - analyze_document — Word count, reading time, element statistics",
                        "  - extract_structure — Full document architecture overview",
                        "",
                        "For repair and quality:",
                        "  - repair_markdown — Fix broken LLM output (unclosed fences, bad tables)",
                        "  - lint_markdown — Find issues with severity/rule/fixable info",
                        "",
                        "For import:",
                        "  - html_to_markdown — Convert HTML content back to Markdown",
                        "",
                        "Tips:",
                        "  - Binary formats (PDF, DOCX, XLSX, PNG) should use output_path",
                        "  - PDF and PNG require a Chromium browser on the system",
                        "  - Set PUPPETEER_EXECUTABLE_PATH to override browser detection",
                        "  - Use harmonize_markdown to clean up formatting before conversion",
                        "  - Use repair_markdown to fix broken LLM output before processing",
                        "  - Use lint_markdown to check quality before sharing documents",
                    ].join("\n"),
                },
            ],
        };
    }

    throw new Error(`Unknown resource: ${uri}`);
});

// ── Tool execution ───────────────────────────────────────────────────
server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
        const { name, arguments: args } = request.params;
        const markdown = (args as any).markdown;
        const outputPath = (args as any).output_path;

        // Tools that don't require the markdown parameter
        const noMarkdownTools = ['html_to_markdown'];
        if (!markdown && !noMarkdownTools.includes(name)) {
            throw new Error("Markdown content is required");
        }

        // Guard against oversized inputs to prevent runaway memory/CPU usage
        const MAX_INPUT_BYTES = 1024 * 1024 * 1024 ; // 1 GB
        const inputToCheck = markdown ?? (args as any).html ?? '';
        if (Buffer.byteLength(inputToCheck, 'utf8') > MAX_INPUT_BYTES) {
            throw new Error('Input too large: content exceeds the 1 GB limit. Please split the document into smaller sections.');
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
            try {
                const page = await browser.newPage();
                await page.setContent(htmlDoc);

                let resultBuffer: Buffer;

                if (name === "convert_to_pdf") {
                    resultBuffer = await page.pdf({ format: 'A4' }) as Buffer;
                    return handleOutput(resultBuffer, outputPath, {
                        format: 'pdf',
                        description: 'PDF document generated from Markdown via Puppeteer'
                    });
                } else {
                    const screenshot = await page.screenshot({ fullPage: true, encoding: 'binary' });
                    resultBuffer = screenshot as Buffer;
                    return handleOutput(resultBuffer, outputPath, {
                        format: 'png',
                        description: 'PNG image screenshot of the rendered Markdown'
                    });
                }
            } finally {
                await browser.close();
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

        // ── Platform converter handlers ──
        if (name === "convert_to_slack") {
            return handleOutput(markdownToSlack(markdown), outputPath);
        }
        if (name === "convert_to_discord") {
            return handleOutput(markdownToDiscord(markdown), outputPath);
        }
        if (name === "convert_to_jira") {
            return handleOutput(markdownToJira(markdown), outputPath);
        }
        if (name === "convert_to_confluence") {
            return handleOutput(markdownToConfluence(markdown), outputPath);
        }
        if (name === "convert_to_asciidoc") {
            return handleOutput(markdownToAsciiDoc(markdown), outputPath);
        }
        if (name === "convert_to_rst") {
            return handleOutput(markdownToRST(markdown), outputPath);
        }
        if (name === "convert_to_mediawiki") {
            return handleOutput(markdownToMediaWiki(markdown), outputPath);
        }
        if (name === "convert_to_bbcode") {
            return handleOutput(markdownToBBCode(markdown), outputPath);
        }
        if (name === "convert_to_textile") {
            return handleOutput(markdownToTextile(markdown), outputPath);
        }
        if (name === "convert_to_orgmode") {
            return handleOutput(markdownToOrgMode(markdown), outputPath);
        }

        // ── Email HTML handler ──
        if (name === "convert_to_email_html") {
            const emailHtml = await markdownToEmailHtml(markdown);
            return handleOutput(emailHtml, outputPath);
        }

        // ── Import handler ──
        if (name === "html_to_markdown") {
            const html = (args as any).html;
            if (!html) throw new Error("HTML content is required");
            const md = htmlToMarkdown(html);
            return handleOutput(md, outputPath);
        }

        // ── Repair / Lint handlers ──
        if (name === "repair_markdown") {
            const repaired = repairMarkdown(markdown);
            return handleOutput(repaired, outputPath);
        }
        if (name === "lint_markdown") {
            const issues = lintMarkdown(markdown);
            return {
                content: [{ type: "text", text: JSON.stringify(issues, null, 2) }],
            };
        }

        // ── Analysis handlers ──
        if (name === "extract_code_blocks") {
            const blocks = extractCodeBlocks(markdown);
            return {
                content: [{ type: "text", text: JSON.stringify(blocks, null, 2) }],
            };
        }
        if (name === "extract_links") {
            const links = extractLinks(markdown);
            return {
                content: [{ type: "text", text: JSON.stringify(links, null, 2) }],
            };
        }
        if (name === "generate_toc") {
            const maxDepth = (args as any).max_depth || 6;
            const toc = generateTOC(markdown, maxDepth);
            return handleOutput(toc, outputPath);
        }
        if (name === "analyze_document") {
            const stats = analyzeDocument(markdown);
            return {
                content: [{ type: "text", text: JSON.stringify(stats, null, 2) }],
            };
        }
        if (name === "extract_structure") {
            const structure = extractStructure(markdown);
            return {
                content: [{ type: "text", text: JSON.stringify(structure, null, 2) }],
            };
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
