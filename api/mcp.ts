import { Server } from "@modelcontextprotocol/sdk/server/index.js";
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
import * as fs from 'fs/promises';
import * as path from 'path';
import {
    parseMarkdownToRTF,
    parseMarkdownToDocx,
    parseMarkdownToLaTeX,
    generateCSV,
    generateJSON,
    generateXML,
    generateXLSXIndex,
    cleanMarkdownText
} from "../src/utils/core-exports.js";
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
} from "../mcp-server/src/platform-converters.js";
import {
    repairMarkdown,
    lintMarkdown,
} from "../mcp-server/src/markdown-repair.js";
import {
    extractCodeBlocks,
    extractLinks,
    generateTOC,
    analyzeDocument,
    extractStructure,
} from "../mcp-server/src/document-analysis.js";
import { htmlToMarkdown } from "../mcp-server/src/html-import.js";
import { markdownToEmailHtml } from "../mcp-server/src/email-html.js";
import { Packer } from "docx";
import type { VercelRequest, VercelResponse } from '@vercel/node';

// Setup browser launch helper for Vercel vs Local
const getBrowser = async () => {
    if (process.env.VERCEL) {
        const chromium = (await import('@sparticuz/chromium-min')) as any;
        const puppeteer = (await import('puppeteer-core')) as any;
        return puppeteer.launch({
            args: chromium.args,
            defaultViewport: chromium.defaultViewport,
            executablePath: await chromium.executablePath('https://github.com/sparticuz/chromium/releases/download/v131.0.1/chromium-v131.0.1-pack.tar'),
            headless: chromium.headless,
        });
    } else {
        const puppeteer = (await import('puppeteer')) as any;
        return puppeteer.launch({ headless: true });
    }
};

// Instance interface
interface McpInstance {
    server: Server;
    transport: any; // StreamableHTTPServerTransport
    isNew: boolean;
    lastUsed: number;
}

// Global registry of active instances in this warm lambda
const instances = new Map<string, McpInstance>();

// Session TTL: evict sessions idle for more than 30 minutes
const SESSION_TTL_MS = 30 * 60 * 1000;
function cleanupExpiredSessions() {
    const now = Date.now();
    for (const [id, inst] of instances.entries()) {
        if (now - inst.lastUsed > SESSION_TTL_MS) {
            instances.delete(id);
        }
    }
}

// Shared setup for all instances
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

    if (Buffer.isBuffer(content)) {
        const sizeBytes = content.length;
        const format = options?.format || 'binary';
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    format: format,
                    file_size_bytes: sizeBytes,
                    description: options?.description || `Generated ${format.toUpperCase()} binary content`,
                    hint: `This is a binary file format. To save the file, call this tool again with the 'output_path' parameter.`,
                    base64_preview: content.toString('base64').substring(0, 100) + '...',
                    full_base64_length: content.toString('base64').length
                }, null, 2)
            }]
        };
    } else {
        return { content: [{ type: "text", text: content }] };
    }
}

function setupServerHandlers(server: Server) {
    // --- Shared parameter description constants ---
    const PARAM_MARKDOWN = "The raw Markdown source text to convert. Supports GitHub-Flavored Markdown (tables, task lists, strikethrough) and KaTeX math expressions. Pass the full document content as a string, not a file path.";
    const PARAM_OUTPUT_PATH_TEXT = "Optional. Absolute or relative file path (e.g. './output.txt') where the result will be saved. Parent directories are created automatically. If omitted, the converted text content is returned directly in the response as a string.";
    const PARAM_OUTPUT_PATH_BINARY = (fmt: string) =>
        `Optional. Absolute or relative file path (e.g. './output.${fmt}') where the binary file will be saved. Parent directories are created automatically. If provided, the file is written to disk and a JSON summary is returned. If omitted, a JSON object with { format, file_size_bytes, hint, base64_preview } is returned. Binary formats (${fmt.toUpperCase()}) should almost always specify output_path.`;
    const PARAM_TITLE = "Optional. A document title string. Used as the root element name or document metadata title in the output. Defaults to 'document' if omitted.";

    const TEXT_TOOL_ANNOTATIONS = {
        title: undefined as string | undefined,
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
    };
    const BROWSER_TOOL_ANNOTATIONS = {
        title: undefined as string | undefined,
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
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
                        "Side effects: when output_path is provided, writes the harmonized Markdown to disk. " +
                        "When output_path is omitted, returns the harmonized text as a string with no file I/O. " +
                        "Returns: harmonized Markdown string (if no output_path), or JSON with { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "The result is a human-readable plain-text string with no markup. " +
                        "Side effects: when output_path is provided, writes the plain text to disk. When output_path is omitted, returns the plain text string directly. " +
                        "Returns: plain text string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "Side effects: when output_path is provided, writes the RTF file to disk. When output_path is omitted, returns the raw RTF markup as a string. " +
                        "Returns: RTF markup string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "list environments, verbatim code blocks, and table environments. KaTeX math expressions are passed through as native LaTeX math. " +
                        "Side effects: when output_path is provided, writes the .tex file to disk. When output_path is omitted, returns the LaTeX source as a string. " +
                        "Returns: LaTeX source string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "Side effects: when output_path is provided, writes the DOCX binary to disk. " +
                        "When output_path is omitted, returns a JSON object with { format, file_size_bytes, hint, base64_preview }. " +
                        "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted).",
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
                        "prints it to PDF via a headless Chromium browser. " +
                        "This is a binary format — output_path should almost always be provided. " +
                        "Side effects: launches a transient headless browser process for rendering. " +
                        "When output_path is provided, writes the PDF to disk. When output_path is omitted, returns JSON { format, file_size_bytes, hint, base64_preview }. " +
                        "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted).",
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
                        "full-page screenshot via a headless Chromium browser. " +
                        "This is a binary format — output_path should almost always be provided. " +
                        "Side effects: launches a transient headless browser process. " +
                        "When output_path is provided, writes the PNG to disk. When output_path is omitted, returns JSON { format, file_size_bytes, hint, base64_preview }. " +
                        "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted).",
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
                        "Non-table content is ignored. " +
                        "Side effects: when output_path is provided, writes the CSV to disk. When output_path is omitted, returns the CSV text directly as a string. " +
                        "Returns: CSV text string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "Side effects: when output_path is provided, writes the JSON to disk. When output_path is omitted, returns the JSON string directly. " +
                        "Returns: JSON string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "Side effects: when output_path is provided, writes the XML to disk. When output_path is omitted, returns the XML string directly. " +
                        "Returns: XML string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
                    inputSchema: {
                        type: "object" as const,
                        properties: {
                            markdown: { type: "string", description: PARAM_MARKDOWN },
                            title: { type: "string", description: "Optional. The root XML element name and document title. Must be a valid XML element name. Defaults to 'document' if omitted." },
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
                        "This is a binary format — output_path should almost always be provided. " +
                        "Side effects: when output_path is provided, writes the XLSX binary to disk. " +
                        "When output_path is omitted, returns JSON { format, file_size_bytes, hint, base64_preview }. " +
                        "Returns: JSON write-confirmation (if output_path set), or JSON binary-guidance object (if omitted).",
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
                        "KaTeX math into semantic HTML with an embedded stylesheet. The output is a full <!DOCTYPE html> document. " +
                        "Side effects: when output_path is provided, writes the HTML file to disk. When output_path is omitted, returns the full HTML string directly. " +
                        "Returns: HTML document string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
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
                        "returns the input Markdown unchanged. When harmonize=true, applies normalization (ATX-style headers, '-' list markers, fenced code blocks). " +
                        "Side effects: when output_path is provided, writes the Markdown to disk. When output_path is omitted, returns the Markdown string directly. " +
                        "Returns: Markdown string (if no output_path), or JSON { success, file_path, file_size_bytes, format } (if output_path set).",
                    inputSchema: {
                        type: "object" as const,
                        properties: {
                            markdown: { type: "string", description: PARAM_MARKDOWN },
                            harmonize: { type: "boolean", description: "Optional. When true, normalizes Markdown syntax before returning or saving. Defaults to false." },
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
                        "Renders GFM and KaTeX math into a full HTML page. Returns the HTML string directly — no file is written to disk. " +
                        "Side effects: none. This tool is read-only and performs no file I/O. " +
                        "Returns: a complete HTML document string (<!DOCTYPE html>…</html>) with inline styles.",
                    inputSchema: {
                        type: "object" as const,
                        properties: {
                            markdown: { type: "string", description: PARAM_MARKDOWN },
                            title: { type: "string", description: "Optional. Sets the <title> tag in the HTML document's <head> section. Defaults to 'Document' if omitted." },
                        },
                        required: ["markdown"],
                    },
                    annotations: {
                        title: "Generate HTML Document",
                        readOnlyHint: true,
                        destructiveHint: false,
                        idempotentHint: true,
                        openWorldHint: false,
                    },
                },
                // ── Platform-specific format tools ──
                {
                    name: "convert_to_slack",
                    description: "Convert Markdown to Slack mrkdwn format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Slack mrkdwn" },
                },
                {
                    name: "convert_to_discord",
                    description: "Convert Markdown to Discord-compatible format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Discord Markdown" },
                },
                {
                    name: "convert_to_jira",
                    description: "Convert Markdown to JIRA wiki markup.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to JIRA Markup" },
                },
                {
                    name: "convert_to_confluence",
                    description: "Convert Markdown to Confluence wiki markup.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Confluence Markup" },
                },
                {
                    name: "convert_to_asciidoc",
                    description: "Convert Markdown to AsciiDoc format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to AsciiDoc" },
                },
                {
                    name: "convert_to_rst",
                    description: "Convert Markdown to reStructuredText (RST) format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to reStructuredText" },
                },
                {
                    name: "convert_to_mediawiki",
                    description: "Convert Markdown to MediaWiki markup.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to MediaWiki" },
                },
                {
                    name: "convert_to_bbcode",
                    description: "Convert Markdown to BBCode format for forum posts.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to BBCode" },
                },
                {
                    name: "convert_to_textile",
                    description: "Convert Markdown to Textile markup format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Textile" },
                },
                {
                    name: "convert_to_orgmode",
                    description: "Convert Markdown to Emacs Org Mode format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Org Mode" },
                },
                {
                    name: "convert_to_email_html",
                    description: "Convert Markdown to email-optimized HTML with all styles inlined for Outlook, Gmail, etc.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Email HTML" },
                },
                // ── Import tools ──
                {
                    name: "html_to_markdown",
                    description: "Convert HTML to Markdown. Handles headings, tables, lists, code blocks, links, images, and inline formatting.",
                    inputSchema: { type: "object" as const, properties: { html: { type: "string", description: "The HTML content to convert to Markdown." }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["html"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Import HTML to Markdown" },
                },
                // ── Repair / Lint tools ──
                {
                    name: "repair_markdown",
                    description: "Repair broken Markdown from LLM output or copy-paste. Fixes unclosed code fences, broken tables, stray emphasis markers, etc.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The potentially broken Markdown text to repair." }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Repair Broken Markdown" },
                },
                {
                    name: "lint_markdown",
                    description: "Lint Markdown and report issues as a JSON array with line, severity, rule, message, and fixable flag.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to lint." } }, required: ["markdown"] },
                    annotations: { title: "Lint Markdown", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                // ── Analysis tools ──
                {
                    name: "extract_code_blocks",
                    description: "Extract all code blocks from Markdown. Returns JSON array with language, code, start/end line numbers.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to extract code blocks from." } }, required: ["markdown"] },
                    annotations: { title: "Extract Code Blocks", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "extract_links",
                    description: "Extract all links and images from Markdown. Returns JSON with text, URL, line, and type.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to extract links from." } }, required: ["markdown"] },
                    annotations: { title: "Extract Links", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "generate_toc",
                    description: "Generate a Table of Contents from Markdown headings with indented links.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to generate a TOC from." }, max_depth: { type: "number", description: "Maximum heading depth (1-6, default: 6)." } }, required: ["markdown"] },
                    annotations: { title: "Generate Table of Contents", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "analyze_document",
                    description: "Analyze Markdown and return comprehensive statistics: word/line/character counts, element counts, reading time.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to analyze." } }, required: ["markdown"] },
                    annotations: { title: "Analyze Document Statistics", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "extract_structure",
                    description: "Extract full document structure: stats, heading outline, code block summary, and link summary.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to extract structure from." } }, required: ["markdown"] },
                    annotations: { title: "Extract Document Structure", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
            ],
        };
    });

    server.setRequestHandler(CallToolRequestSchema, async (request) => {
        try {
            const { name, arguments: args } = request.params;
            const markdown = (args as any).markdown;
            const outputPath = (args as any).output_path;

            const noMarkdownTools = ['html_to_markdown'];
            if (!markdown && !noMarkdownTools.includes(name)) throw new Error("Markdown content is required");

            // Guard against oversized inputs to prevent timeouts and memory pressure
            const MAX_INPUT_BYTES = 1024 * 1024 * 1024; // 1 GB
            const inputToCheck = markdown ?? (args as any).html ?? '';
            if (Buffer.byteLength(inputToCheck, 'utf8') > MAX_INPUT_BYTES) {
                throw new Error('Input too large: content exceeds the 1 GB limit. Please split the document into smaller sections.');
            }

            if (name === "harmonize_markdown") {
                const file = await unified().use(remarkParse).use(remarkGfm).use(remarkMath).use(remarkStringify, { bullet: '-', fence: '`', fences: true, incrementListMarker: true, listItemIndent: 'one' }).process(markdown);
                return handleOutput(String(file), outputPath);
            }

            if (name === "convert_to_txt") return handleOutput(cleanMarkdownText(markdown), outputPath);
            if (name === "convert_to_rtf") return handleOutput(parseMarkdownToRTF(markdown), outputPath);
            if (name === "convert_to_latex") return handleOutput(parseMarkdownToLaTeX(markdown), outputPath);

            if (name === "convert_to_docx") {
                const elements = parseMarkdownToDocx(markdown);
                const doc = new (await import("docx")).Document({ sections: [{ children: elements }] });
                const buffer = await Packer.toBuffer(doc);
                return handleOutput(buffer, outputPath, { format: 'docx', description: 'Word document' });
            }

            if (name === "convert_to_csv") return handleOutput(generateCSV(markdown), outputPath);
            if (name === "convert_to_json") return handleOutput(generateJSON(markdown, (args as any).title), outputPath);
            if (name === "convert_to_xml") return handleOutput(generateXML(markdown, (args as any).title), outputPath);
            if (name === "convert_to_xlsx") return handleOutput(generateXLSXIndex(markdown), outputPath, { format: 'xlsx', description: 'Excel spreadsheet' });

            if (name === "convert_to_html" || name === "convert_to_pdf" || name === "convert_to_image") {
                const htmlFile = await unified().use(remarkParse).use(remarkGfm).use(remarkRehype).use(rehypeKatex).use(rehypeStringify).process(markdown);
                const htmlDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous"><style>body { font-family: system-ui; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto; background: white; color: black; } img { max-width: 100%; } pre { background: #f4f4f4; padding: 15px; border-radius: 5px; overflow-x: auto; } table { border-collapse: collapse; width: 100%; margin: 1em 0; } th, td { border: 1px solid #ddd; padding: 8px; text-align: left; } th { background-color: #f2f2f2; } blockquote { border-left: 4px solid #ddd; margin: 0; padding-left: 1em; color: #666; }</style></head><body>${String(htmlFile)}</body></html>`;

                if (name === "convert_to_html") return handleOutput(htmlDoc, outputPath);

                const browser = await getBrowser();
                try {
                    const page = await browser.newPage();
                    await page.setContent(htmlDoc);
                    let resultBuffer: Buffer;

                    if (name === "convert_to_pdf") {
                        resultBuffer = Buffer.from(await page.pdf({ format: 'A4' }));
                        return handleOutput(resultBuffer, outputPath, { format: 'pdf', description: 'PDF document' });
                    } else {
                        resultBuffer = Buffer.from(await page.screenshot({ fullPage: true, encoding: 'binary' }));
                        return handleOutput(resultBuffer, outputPath, { format: 'png', description: 'PNG image' });
                    }
                } finally {
                    await browser.close();
                }
            }

            if (name === "convert_to_md") {
                if (!(args as any).harmonize) return handleOutput(markdown, outputPath);
                const file = await unified().use(remarkParse).use(remarkGfm).use(remarkMath).use(remarkStringify, { bullet: '-', fence: '`', fences: true, incrementListMarker: true, listItemIndent: 'one' }).process(markdown);
                return handleOutput(String(file), outputPath);
            }

            if (name === "generate_html") {
                const htmlFile = await unified().use(remarkParse).use(remarkGfm).use(remarkRehype).use(rehypeKatex).use(rehypeStringify).process(markdown);
                const htmlDoc = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${(args as any).title || 'Document'}</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous">
    <style>
        body { font-family: system-ui, -apple-system, sans-serif; max-width: 800px; margin: 40px auto; padding: 20px; line-height: 1.6; color: #1a1a1a; }
        h1, h2, h3 { color: #111; margin-top: 2em; }
        pre { background: #f4f4f4; padding: 15px; border-radius: 5px; overflow-x: auto; }
        code { font-family: monospace; background: #eee; padding: 2px 4px; border-radius: 3px; }
        table { border-collapse: collapse; width: 100%; margin: 1em 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background: #f8f8f8; }
        blockquote { border-left: 4px solid #ddd; margin: 0; padding-left: 1em; color: #666; }
        img { max-width: 100%; }
    </style>
</head>
<body>${String(htmlFile)}</body>
</html>`;
                return { content: [{ type: "text", text: htmlDoc }] };
            }

            // ── Platform converters ──
            if (name === "convert_to_slack") return handleOutput(markdownToSlack(markdown), outputPath);
            if (name === "convert_to_discord") return handleOutput(markdownToDiscord(markdown), outputPath);
            if (name === "convert_to_jira") return handleOutput(markdownToJira(markdown), outputPath);
            if (name === "convert_to_confluence") return handleOutput(markdownToConfluence(markdown), outputPath);
            if (name === "convert_to_asciidoc") return handleOutput(markdownToAsciiDoc(markdown), outputPath);
            if (name === "convert_to_rst") return handleOutput(markdownToRST(markdown), outputPath);
            if (name === "convert_to_mediawiki") return handleOutput(markdownToMediaWiki(markdown), outputPath);
            if (name === "convert_to_bbcode") return handleOutput(markdownToBBCode(markdown), outputPath);
            if (name === "convert_to_textile") return handleOutput(markdownToTextile(markdown), outputPath);
            if (name === "convert_to_orgmode") return handleOutput(markdownToOrgMode(markdown), outputPath);

            // ── Email HTML ──
            if (name === "convert_to_email_html") {
                const emailHtml = await markdownToEmailHtml(markdown);
                return handleOutput(emailHtml, outputPath);
            }

            // ── Import ──
            if (name === "html_to_markdown") {
                const html = (args as any).html;
                if (!html) throw new Error("HTML content is required");
                return handleOutput(htmlToMarkdown(html), outputPath);
            }

            // ── Repair / Lint ──
            if (name === "repair_markdown") return handleOutput(repairMarkdown(markdown), outputPath);
            if (name === "lint_markdown") return { content: [{ type: "text", text: JSON.stringify(lintMarkdown(markdown), null, 2) }] };

            // ── Analysis ──
            if (name === "extract_code_blocks") return { content: [{ type: "text", text: JSON.stringify(extractCodeBlocks(markdown), null, 2) }] };
            if (name === "extract_links") return { content: [{ type: "text", text: JSON.stringify(extractLinks(markdown), null, 2) }] };
            if (name === "generate_toc") return handleOutput(generateTOC(markdown, (args as any).max_depth || 6), outputPath);
            if (name === "analyze_document") return { content: [{ type: "text", text: JSON.stringify(analyzeDocument(markdown), null, 2) }] };
            if (name === "extract_structure") return { content: [{ type: "text", text: JSON.stringify(extractStructure(markdown), null, 2) }] };

            throw new Error(`Unknown tool: ${name}`);
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    });

    // ── Prompts ──────────────────────────────────────────────────────
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
                            description: "Target output format: pdf, docx, html, latex, csv, json, xml, xlsx, rtf, png, txt, or md",
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
                            description: "Output format for tables: 'csv' for plain text or 'xlsx' for Excel spreadsheet",
                            required: true,
                        },
                        {
                            name: "markdown",
                            description: "The Markdown content containing tables to extract",
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
                    description: "Analyze a Markdown document for issues, repair problems, and return both the lint report and repaired document.",
                    arguments: [
                        { name: "markdown", description: "The Markdown content to analyze and repair", required: true },
                    ],
                },
                {
                    name: "convert-for-platform",
                    description: "Convert Markdown to a platform-specific format: slack, discord, jira, confluence, asciidoc, rst, mediawiki, bbcode, textile, orgmode.",
                    arguments: [
                        { name: "platform", description: "Target platform: slack, discord, jira, confluence, asciidoc, rst, mediawiki, bbcode, textile, or orgmode", required: true },
                        { name: "markdown", description: "The Markdown content to convert", required: true },
                    ],
                },
                {
                    name: "document-overview",
                    description: "Get a comprehensive overview of a Markdown document: statistics, structure, TOC, code blocks, and links.",
                    arguments: [
                        { name: "markdown", description: "The Markdown content to analyze", required: true },
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
                messages: [{ role: "user" as const, content: { type: "text" as const, text: `Please analyze and repair the following Markdown document:\n1. Use lint_markdown to identify issues.\n2. Use repair_markdown to fix them.\n3. Use lint_markdown again to confirm.\n\n${markdown}` } }],
            };
        }

        if (name === "convert-for-platform") {
            const platform = args?.platform || "slack";
            const markdown = args?.markdown || "";
            return {
                description: `Convert Markdown for ${platform}`,
                messages: [{ role: "user" as const, content: { type: "text" as const, text: `Please convert the following Markdown to ${platform} format using the convert_to_${platform} tool.\n\n${markdown}` } }],
            };
        }

        if (name === "document-overview") {
            const markdown = args?.markdown || "";
            return {
                description: "Comprehensive document overview",
                messages: [{ role: "user" as const, content: { type: "text" as const, text: `Please provide a comprehensive overview of the following Markdown document:\n1. Use analyze_document for statistics\n2. Use generate_toc for table of contents\n3. Use extract_code_blocks for code snippets\n4. Use extract_links for all links\n\n${markdown}` } }],
            };
        }

        throw new Error(`Unknown prompt: ${name}`);
    });

    // ── Resources ────────────────────────────────────────────────────
    server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => {
        return { resourceTemplates: [] };
    });

    server.setRequestHandler(ListResourcesRequestSchema, async () => {
        return {
            resources: [
                {
                    uri: "markdown-formatter://supported-formats",
                    name: "Supported Output Formats",
                    description: "Complete list of all supported output formats with tool names, types, and descriptions",
                    mimeType: "application/json",
                },
                {
                    uri: "markdown-formatter://conversion-guide",
                    name: "Conversion Guide",
                    description: "Guide for choosing the right output format based on your use case",
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

}

import { WebStandardStreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js";

async function getOrCreateInstance(sessionId: string): Promise<McpInstance> {
    if (instances.has(sessionId)) {
        const instance = instances.get(sessionId)!;
        instance.isNew = false;
        instance.lastUsed = Date.now();
        return instance;
    }

    const transport = new WebStandardStreamableHTTPServerTransport({
        sessionIdGenerator: () => sessionId,
    });

    const server = new Server(
        {
            name: "markdown-formatter-mcp",
            version: "2.1.0",
        },
        {
            capabilities: {
                tools: {},
                resources: {},
                prompts: {}
            }
        }
    );

    setupServerHandlers(server);
    await server.connect(transport);

    const instance = { server, transport, isNew: true, lastUsed: Date.now() };
    instances.set(sessionId, instance);

    return instance;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
    // Add CORS and Streaming headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, DELETE');
    res.setHeader('Access-Control-Allow-Headers', '*');
    res.setHeader('Access-Control-Expose-Headers', '*');
    res.setHeader('X-Accel-Buffering', 'no');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('Content-Type', 'application/json');

    // 1. High-priority: Handle server-card.json for Smithery discovery
    if (req.url?.includes('server-card.json') || req.url?.includes('.well-known/mcp')) {
        const serverCard = {
            name: "markdown-formatter-mcp",
            displayName: "AI Answer Copier — Markdown Formatter",
            description: "MCP Server with 33 tools: convert Markdown to 23 formats (PDF, DOCX, HTML, Slack, Discord, JIRA, Confluence, AsciiDoc, RST, MediaWiki, BBCode, Textile, Org Mode, Email HTML, and more), plus document analysis, repair/lint, and HTML import. Built for educators, developers, and AI workflows.",
            homepage: "https://ai-answer-copier.vercel.app",
            repository: "https://github.com/XJTLUmedia/AI_answer_copier",
            icons: {
                light: "https://raw.githubusercontent.com/XJTLUmedia/AI_answer_copier/main/mcp-server/icon.svg",
                dark: "https://raw.githubusercontent.com/XJTLUmedia/AI_answer_copier/main/mcp-server/icon.svg",
            },
            mcpV1: {
                capabilities: { tools: true, prompts: true, resources: true },
                tools: [
                    { name: "harmonize_markdown", description: "Standardize and normalize Markdown syntax without changing the document's meaning" },
                    { name: "convert_to_txt", description: "Convert Markdown to plain text by stripping all formatting" },
                    { name: "convert_to_rtf", description: "Convert Markdown to Rich Text Format (RTF)" },
                    { name: "convert_to_latex", description: "Convert Markdown to LaTeX source code" },
                    { name: "convert_to_docx", description: "Convert Markdown to Microsoft Word DOCX file" },
                    { name: "convert_to_pdf", description: "Convert Markdown to PDF document via headless Chromium" },
                    { name: "convert_to_image", description: "Convert Markdown to PNG image via headless Chromium" },
                    { name: "convert_to_csv", description: "Extract tables from Markdown and convert to CSV" },
                    { name: "convert_to_json", description: "Convert Markdown to structured JSON representation" },
                    { name: "convert_to_xml", description: "Convert Markdown to XML document" },
                    { name: "convert_to_xlsx", description: "Convert Markdown tables to Excel XLSX spreadsheet" },
                    { name: "convert_to_html", description: "Convert Markdown to complete, styled HTML document" },
                    { name: "convert_to_md", description: "Export Markdown content, optionally harmonized" },
                    { name: "generate_html", description: "Generate self-contained HTML document (read-only, no file I/O)" },
                    { name: "convert_to_email_html", description: "Convert Markdown to email-optimized HTML with inlined styles" },
                    { name: "convert_to_slack", description: "Convert Markdown to Slack mrkdwn format" },
                    { name: "convert_to_discord", description: "Convert Markdown to Discord-compatible formatting" },
                    { name: "convert_to_jira", description: "Convert Markdown to Atlassian JIRA markup" },
                    { name: "convert_to_confluence", description: "Convert Markdown to Atlassian Confluence wiki markup" },
                    { name: "convert_to_asciidoc", description: "Convert Markdown to AsciiDoc format" },
                    { name: "convert_to_rst", description: "Convert Markdown to reStructuredText" },
                    { name: "convert_to_mediawiki", description: "Convert Markdown to MediaWiki markup" },
                    { name: "convert_to_bbcode", description: "Convert Markdown to BBCode for forums" },
                    { name: "convert_to_textile", description: "Convert Markdown to Textile markup" },
                    { name: "convert_to_orgmode", description: "Convert Markdown to Emacs Org Mode format" },
                    { name: "html_to_markdown", description: "Convert HTML content back to Markdown" },
                    { name: "repair_markdown", description: "Fix broken Markdown from LLM output or copy-paste" },
                    { name: "lint_markdown", description: "Lint Markdown and report issues with severity and location" },
                    { name: "extract_code_blocks", description: "Extract all code blocks with language and line info" },
                    { name: "extract_links", description: "Extract all links and images with type classification" },
                    { name: "generate_toc", description: "Generate Table of Contents from Markdown headings" },
                    { name: "analyze_document", description: "Comprehensive document statistics (words, lines, reading time)" },
                    { name: "extract_structure", description: "Full document structure overview (stats, outline, summaries)" },
                ],
                prompts: [
                    { name: "convert-document", description: "Convert a Markdown document to a specified output format" },
                    { name: "extract-tables", description: "Extract tables from Markdown as CSV or XLSX" },
                    { name: "format-for-sharing", description: "Harmonize and convert Markdown to PDF + HTML for sharing" },
                    { name: "analyze-and-repair", description: "Lint, repair, and re-lint a Markdown document" },
                    { name: "convert-for-platform", description: "Convert Markdown for a specific platform (Slack, Discord, JIRA, etc.)" },
                    { name: "document-overview", description: "Get a comprehensive overview: stats, TOC, code blocks, links" },
                ],
                resources: [
                    { uri: "markdown-formatter://supported-formats", name: "Supported Output Formats" },
                    { uri: "markdown-formatter://conversion-guide", name: "Conversion Guide" },
                ],
            }
        };
        return res.status(200).json(serverCard);
    }

    const providedSessionId = (req.query.sessionId as string) || (req.headers['mcp-session-id'] as string);
    const sessionId = providedSessionId || `s_${Math.random().toString(36).substring(2, 10)}`;
    res.setHeader('mcp-session-id', sessionId);

    // Evict stale sessions on every request (cheap O(n) scan, n stays small in practice)
    cleanupExpiredSessions();

    // Handle DELETE: terminate and remove the session, don't create a new one
    if (req.method === 'DELETE') {
        if (providedSessionId && instances.has(providedSessionId)) {
            instances.delete(providedSessionId);
        }
        res.status(200).json({ message: 'Session terminated' });
        return;
    }

    const isEventStream =
        req.headers.accept?.includes('text/event-stream') ||
        req.headers['mcp-protocol-version'] ||
        (req.query.sessionId && req.method === 'GET');

    if (req.method === 'GET' && !isEventStream) {
        res.status(200).setHeader('Content-Type', 'text/html').send(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Markdown Formatter MCP</title>
                <style>
                    body { font-family: system-ui, -apple-system, sans-serif; padding: 40px; line-height: 1.6; max-width: 700px; margin: 0 auto; background: #0f172a; color: #f8fafc; }
                    .card { background: #1e293b; padding: 24px; border-radius: 12px; border: 1px solid #334155; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); }
                    pre { background: #0f172a; padding: 16px; border-radius: 8px; overflow-x: auto; color: #38bdf8; font-family: 'JetBrains Mono', monospace; font-size: 0.9rem; }
                    .status { display: inline-flex; align-items: center; gap: 8px; padding: 4px 12px; border-radius: 99px; background: #064e3b; color: #34d399; font-size: 0.8125rem; font-weight: 600; }
                    .dot { width: 8px; height: 8px; background: #34d399; border-radius: 50%; box-shadow: 0 0 8px #34d399; }
                    h1 { margin: 0; font-size: 1.5rem; letter-spacing: -0.025em; }
                    code { color: #f472b6; }
                </style>
            </head>
            <body>
                <div class="card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px;">
                        <h1>Markdown Formatter MCP</h1>
                        <span class="status"><span class="dot"></span> Online</span>
                    </div>
                    <p>This is a Model Context Protocol (MCP) server endpoint running as a Vercel Serverless Function.</p>
                    <p>Active instances in this node: ${instances.size}</p>
                    
                    <h2 style="font-size: 1.125rem; margin-top: 32px; color: #94a3b8;">Setup Instructions</h2>
                    <p>To use this server, add it to your <code>claude_desktop_config.json</code> (or typical MCP client):</p>
                    <pre>https://ai-answer-copier.vercel.app/api/mcp</pre>
                </div>
            </body>
            </html>
        `);
        return;
    }

    try {
        const instance = await getOrCreateInstance(sessionId);

        // Anti-409: Close existing stream if this is a new GET for the same session
        if (req.method === 'GET' && isEventStream) {
            instance.transport.closeStandaloneSSEStream();
        }

        const body = (req.method === 'POST' || req.method === 'PUT') ? req.body : undefined;

        // Force initialization for all methods if the instance is new, 
        // because in a serverless environment we can't rely on the client
        // hitting the same instance for the 'initialize' message.
        if (instance.isNew) {
            console.log(`[MCP] Cold Start/Instance Migration detected for session ${sessionId}. Forcing initialization.`);
            // Access internal SDK property to bypass the initialize handshake requirement
            // in stateless serverless environments. Guarded to survive future SDK changes.
            const srv = instance.server as any;
            if (typeof srv._initialized !== 'undefined') {
                srv._initialized = true;
            }
        }

        // Build absolute URL for the Web Request
        const protocol = req.headers['x-forwarded-proto'] || 'http';
        const host = req.headers.host || 'localhost';
        const url = new URL(req.url!, `${protocol}://${host}`);

        // Construct headers correctly
        const headers = new Headers();
        Object.entries(req.headers).forEach(([k, v]) => {
            if (v) {
                if (Array.isArray(v)) v.forEach(val => headers.append(k, val));
                else headers.set(k, v as string);
            }
        });

        const webRequest = new Request(url, {
            method: req.method,
            headers: headers,
            body: body ? (typeof body === 'string' ? body : JSON.stringify(body)) : undefined
        });

        const webResponse = await instance.transport.handleRequest(webRequest);

        // Handle stream piping
        if (webResponse.body) {
            // For SSE, we MUST send headers immediately to satisfy Vercel/proxies
            const contentType = webResponse.headers.get('Content-Type') || '';
            const isSseResponse = contentType.includes('text/event-stream');

            if (isSseResponse) {
                res.status(200);
                // Copy all headers from the SDK response (including mcp-session-id)
                webResponse.headers.forEach((v: string, k: string) => {
                    res.setHeader(k, v);
                });
                // Force headers that are critical for Vercel/proxies
                res.setHeader('Content-Type', 'text/event-stream');
                res.setHeader('Cache-Control', 'no-cache, no-transform');
                res.setHeader('Connection', 'keep-alive');
                res.setHeader('X-Accel-Buffering', 'no');
                res.write(': heartbeat\n\n'); // Initial handshake
            } else {
                res.status(webResponse.status);
                webResponse.headers.forEach((v: string, k: string) => {
                    if (!res.getHeader(k)) res.setHeader(k, v);
                });
            }

            const reader = webResponse.body.getReader();
            try {
                // Keep-alive timer for long SSE connections on serverless
                let keepAlive: NodeJS.Timeout | undefined;
                if (isSseResponse) {
                    keepAlive = setInterval(() => {
                        res.write(': keep-alive\n\n');
                    }, 15000);
                    res.on('close', () => clearInterval(keepAlive));
                }

                try {
                    while (true) {
                        const { done, value } = await reader.read();
                        if (done) break;
                        res.write(value);
                    }
                } finally {
                    // Clear timer whether the stream ended normally or errored
                    if (keepAlive !== undefined) clearInterval(keepAlive);
                }
            } finally {
                reader.releaseLock();
            }
        } else {
            // Standard JSON/Error response
            res.status(webResponse.status);
            webResponse.headers.forEach((v: string, k: string) => {
                if (!res.getHeader(k)) res.setHeader(k, v);
            });
            const text = await webResponse.text();
            res.send(text);
        }
        res.end();
    } catch (error: any) {
        console.error("[MCP] Execution error:", error);
        if (!res.headersSent) {
            res.status(500).json({ error: error.message });
        } else {
            res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
            res.end();
        }
    }
}
