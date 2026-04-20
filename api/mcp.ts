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
} from "../mcp-server/src/core-exports.js";
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
    config: ServerConfig;
}

// Session-level configuration (from Smithery Gateway query params)
interface ServerConfig {
    pdf_page_format: string;
    pdf_margin: string;
    html_theme: string;
    default_title: string;
    max_input_bytes: number;
}

function getDefaultConfig(): ServerConfig {
    return {
        pdf_page_format: 'A4',
        pdf_margin: '20mm',
        html_theme: 'light',
        default_title: 'document',
        // 10 MB default: enough for book-length docs, small enough to prevent
        // accidental DoS via multi-GB payloads.
        max_input_bytes: 10 * 1024 * 1024,
    };
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

function setupServerHandlers(server: Server, config: ServerConfig) {
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
                        readOnlyHint: true,
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
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Slack mrkdwn" },
                },
                {
                    name: "convert_to_discord",
                    description:
                        "Convert Markdown to Discord-compatible format. Transforms headers to styled bold/underline text that renders " +
                        "correctly in Discord messages. Code blocks and basic formatting are preserved.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Discord Markdown" },
                },
                {
                    name: "convert_to_jira",
                    description:
                        "Convert Markdown to JIRA wiki markup. Transforms headers to h1./h2., bold to single asterisks, " +
                        "code blocks to {code} blocks, links to [text|url], and lists to JIRA * and # syntax.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to JIRA Markup" },
                },
                {
                    name: "convert_to_confluence",
                    description:
                        "Convert Markdown to Confluence wiki markup. Similar to JIRA but includes Confluence-specific {info}, {note} panels " +
                        "and {code:language=x} syntax.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Confluence Markup" },
                },
                {
                    name: "convert_to_asciidoc",
                    description:
                        "Convert Markdown to AsciiDoc format. Transforms headers to = syntax, code blocks to ---- delimited blocks, " +
                        "links to url[text] syntax, and images to image::url[alt] directives.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to AsciiDoc" },
                },
                {
                    name: "convert_to_rst",
                    description:
                        "Convert Markdown to reStructuredText (RST) format. Transforms headers to underlined text, " +
                        "code blocks to .. code-block:: directives, and links to RST reference syntax.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to reStructuredText" },
                },
                {
                    name: "convert_to_mediawiki",
                    description:
                        "Convert Markdown to MediaWiki markup. Transforms headers to == syntax, bold to triple quotes, " +
                        "code to <syntaxhighlight> tags, and tables to {| wikitable format.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to MediaWiki" },
                },
                {
                    name: "convert_to_bbcode",
                    description:
                        "Convert Markdown to BBCode format. Transforms formatting to [b], [i], [s], [code], [url], [img] tags. " +
                        "Used for forum posts on phpBB, vBulletin, and similar platforms.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to BBCode" },
                },
                {
                    name: "convert_to_textile",
                    description:
                        "Convert Markdown to Textile markup format. Used by Redmine, older versions of Basecamp, and some CMS platforms.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Textile" },
                },
                {
                    name: "convert_to_orgmode",
                    description:
                        "Convert Markdown to Emacs Org Mode format. Transforms headers to * syntax, bold to *text*, " +
                        "code blocks to #+BEGIN_SRC/#+END_SRC, and links to [[url][text]] syntax.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Org Mode" },
                },
                {
                    name: "convert_to_email_html",
                    description:
                        "Convert Markdown to email-optimized HTML with all styles inlined. Produces HTML compatible with " +
                        "Outlook, Gmail, Apple Mail, and other email clients. No external CSS dependencies. " +
                        "Wraps content in a responsive email table layout.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: PARAM_MARKDOWN }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Convert to Email HTML" },
                },
                // ── Import tools ──
                {
                    name: "html_to_markdown",
                    description:
                        "Convert HTML to Markdown. Performs round-trip import of HTML content back to Markdown format. " +
                        "Handles headings, tables, lists, code blocks, links, images, and inline formatting. " +
                        "Useful for importing web content or converting HTML emails to Markdown.",
                    inputSchema: { type: "object" as const, properties: { html: { type: "string", description: "The HTML content to convert to Markdown. Can be a full HTML document or a fragment." }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["html"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Import HTML to Markdown" },
                },
                // ── Repair / Lint tools ──
                {
                    name: "repair_markdown",
                    description:
                        "Repair broken Markdown from LLM output or copy-paste. Fixes unclosed code fences, broken tables " +
                        "(mismatched columns, missing separators), stray emphasis markers, missing heading spaces, " +
                        "inconsistent list indentation, broken links, and excessive whitespace.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The potentially broken Markdown text to repair." }, output_path: { type: "string", description: PARAM_OUTPUT_PATH_TEXT } }, required: ["markdown"] },
                    annotations: { ...TEXT_TOOL_ANNOTATIONS, title: "Repair Broken Markdown" },
                },
                {
                    name: "lint_markdown",
                    description:
                        "Lint Markdown and report issues. Returns a JSON array of lint issues found in the document, " +
                        "each with line number, column, severity (error/warning/info), rule name, message, and fixable flag. " +
                        "Checks for: missing heading spaces, trailing whitespace, inconsistent list markers, hard tabs, " +
                        "multiple blank lines, bare URLs, unclosed emphasis, and unclosed code fences.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to lint." } }, required: ["markdown"] },
                    annotations: { title: "Lint Markdown", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                // ── Analysis tools ──
                {
                    name: "extract_code_blocks",
                    description:
                        "Extract all code blocks from a Markdown document. Returns a JSON array of code blocks, " +
                        "each with language, code content, and start/end line numbers. " +
                        "Useful for extracting code snippets from LLM responses or documentation.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to extract code blocks from." } }, required: ["markdown"] },
                    annotations: { title: "Extract Code Blocks", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "extract_links",
                    description:
                        "Extract all links and images from a Markdown document. Returns a JSON array with link text, URL, " +
                        "line number, and type (inline, reference, image, autolink). " +
                        "Useful for link checking, SEO analysis, or extracting references.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to extract links from." } }, required: ["markdown"] },
                    annotations: { title: "Extract Links", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "generate_toc",
                    description:
                        "Generate a Table of Contents from Markdown headings. Returns a Markdown-formatted TOC with " +
                        "indented links to each heading. Handles duplicate heading slugs. " +
                        "The max_depth parameter controls the deepest heading level to include.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to generate a TOC from." }, max_depth: { type: "number", description: "Maximum heading depth (1-6, default: 6)." } }, required: ["markdown"] },
                    annotations: { title: "Generate Table of Contents", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "analyze_document",
                    description:
                        "Analyze a Markdown document and return comprehensive statistics. Returns JSON with: " +
                        "line/word/character/paragraph/sentence counts, heading/code block/table/link/image/list/blockquote counts, " +
                        "and estimated reading time in minutes.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to analyze." } }, required: ["markdown"] },
                    annotations: { title: "Analyze Document Statistics", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                {
                    name: "extract_structure",
                    description:
                        "Extract the full structure of a Markdown document. Returns JSON with document statistics, " +
                        "heading outline, code block summary (language, line count, positions), and link summary " +
                        "(totals by type, unique URL count). Provides a bird's-eye view of document architecture.",
                    inputSchema: { type: "object" as const, properties: { markdown: { type: "string", description: "The Markdown text to extract structure from." } }, required: ["markdown"] },
                    annotations: { title: "Extract Document Structure", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
                },
                // ── Batch conversion tool ──
                {
                    name: "batch_convert",
                    description:
                        "Batch convert multiple markdown documents to multiple output formats in a single call. " +
                        "Supports all conversion formats: txt, csv, json, xml, xlsx, latex, rtf, docx, pdf, html, md, email_html, " +
                        "slack, discord, jira, confluence, asciidoc, rst, mediawiki, bbcode, textile, orgmode. " +
                        "Each item in the batch is processed independently — one failure does not stop the rest. " +
                        "Results include per-item success/error status. " +
                        "When output_dir is provided, files are saved to disk in subdirectories per input document. " +
                        "Without output_dir, text format results are returned inline; binary formats return base64.",
                    inputSchema: {
                        type: "object" as const,
                        properties: {
                            items: {
                                type: "array",
                                description: "Array of markdown documents to convert. Each item has 'markdown' (required) and 'title' (optional, used for filename and metadata).",
                                items: {
                                    type: "object",
                                    properties: {
                                        markdown: { type: "string", description: "The markdown content to convert" },
                                        title: { type: "string", description: "Document title (used for filename and metadata). Defaults to 'document'" },
                                    },
                                    required: ["markdown"],
                                },
                            },
                            formats: {
                                type: "array",
                                description: "Array of target format IDs. Valid values: txt, csv, json, xml, xlsx, latex, rtf, docx, pdf, html, md, email_html, slack, discord, jira, confluence, asciidoc, rst, mediawiki, bbcode, textile, orgmode",
                                items: { type: "string" },
                            },
                            output_dir: {
                                type: "string",
                                description: "Optional base directory to save converted files. Files saved as: {output_dir}/{title}/{title}.{ext}. If omitted, results are returned inline.",
                            },
                        },
                        required: ["items", "formats"],
                    },
                    annotations: {
                        title: "Batch Convert Markdown",
                        readOnlyHint: false,
                        destructiveHint: false,
                        idempotentHint: true,
                        openWorldHint: false,
                    },
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
            const maxBytes = config.max_input_bytes;
            const inputToCheck = markdown ?? (args as any).html ?? '';
            if (Buffer.byteLength(inputToCheck, 'utf8') > maxBytes) {
                throw new Error(`Input too large: content exceeds the ${Math.round(maxBytes / (1024 * 1024))} MB limit. Please split the document into smaller sections.`);
            }

            if (name === "harmonize_markdown") {
                const file = await unified().use(remarkParse).use(remarkGfm).use(remarkMath).use(remarkStringify, { bullet: '-', fence: '`', fences: true, incrementListMarker: true, listItemIndent: 'one' }).process(markdown);
                return handleOutput(String(file), outputPath);
            }

            if (name === "convert_to_txt") return handleOutput(cleanMarkdownText(markdown), outputPath);
            if (name === "convert_to_rtf") return handleOutput(parseMarkdownToRTF(markdown), outputPath);
            if (name === "convert_to_latex") return handleOutput(parseMarkdownToLaTeX(markdown), outputPath);

            if (name === "convert_to_docx") {
                const { elements, footnotes } = parseMarkdownToDocx(markdown);
                const docOptions: any = {
                    sections: [{ children: elements }]
                };
                if (Object.keys(footnotes).length > 0) {
                    docOptions.footnotes = footnotes;
                }
                const doc = new (await import("docx")).Document(docOptions);
                const buffer = await Packer.toBuffer(doc);
                return handleOutput(buffer, outputPath, {
                    format: 'docx',
                    description: 'Microsoft Word document generated from Markdown'
                });
            }

            if (name === "convert_to_csv") return handleOutput(generateCSV(markdown), outputPath);
            if (name === "convert_to_json") return handleOutput(generateJSON(markdown, (args as any).title || config.default_title), outputPath);
            if (name === "convert_to_xml") return handleOutput(generateXML(markdown, (args as any).title || config.default_title), outputPath);
            if (name === "convert_to_xlsx") return handleOutput(generateXLSXIndex(markdown), outputPath, { format: 'xlsx', description: 'Excel spreadsheet' });

            if (name === "convert_to_html" || name === "convert_to_pdf" || name === "convert_to_image") {
                const htmlFile = await unified().use(remarkParse).use(remarkGfm).use(remarkRehype).use(rehypeKatex).use(rehypeStringify).process(markdown);
                const isDark = config.html_theme === 'dark';
                const themeStyles = isDark
                    ? 'background: #1a1a2e; color: #e0e0e0;'
                    : 'background: white; color: black;';
                const preStyles = isDark
                    ? 'background: #16213e; color: #e0e0e0;'
                    : 'background: #f4f4f4;';
                const thStyles = isDark
                    ? 'background-color: #0f3460;'
                    : 'background-color: #f2f2f2;';
                const borderColor = isDark ? '#334155' : '#ddd';
                const bqColor = isDark ? '#94a3b8' : '#666';
                const htmlDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous"><style>body { font-family: system-ui; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto; ${themeStyles} } img { max-width: 100%; } pre { ${preStyles} padding: 15px; border-radius: 5px; overflow-x: auto; } table { border-collapse: collapse; width: 100%; margin: 1em 0; } th, td { border: 1px solid ${borderColor}; padding: 8px; text-align: left; } th { ${thStyles} } blockquote { border-left: 4px solid ${borderColor}; margin: 0; padding-left: 1em; color: ${bqColor}; }</style></head><body>${String(htmlFile)}</body></html>`;

                if (name === "convert_to_html") return handleOutput(htmlDoc, outputPath);

                const browser = await getBrowser();
                try {
                    const page = await browser.newPage();
                    await page.setContent(htmlDoc);
                    let resultBuffer: Buffer;

                    if (name === "convert_to_pdf") {
                        const m = config.pdf_margin;
                        resultBuffer = Buffer.from(await page.pdf({ format: config.pdf_page_format as any, margin: { top: m, right: m, bottom: m, left: m } }));
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
                const isDark = config.html_theme === 'dark';
                const docTitle = (args as any).title || config.default_title || 'Document';
                const htmlDoc = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${docTitle}</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous">
    <style>
        body { font-family: system-ui, -apple-system, sans-serif; max-width: 800px; margin: 40px auto; padding: 20px; line-height: 1.6; ${isDark ? 'background: #1a1a2e; color: #e0e0e0;' : 'color: #1a1a1a;'} }
        h1, h2, h3 { ${isDark ? 'color: #f0f0f0;' : 'color: #111;'} margin-top: 2em; }
        pre { ${isDark ? 'background: #16213e; color: #e0e0e0;' : 'background: #f4f4f4;'} padding: 15px; border-radius: 5px; overflow-x: auto; }
        code { font-family: monospace; ${isDark ? 'background: #16213e;' : 'background: #eee;'} padding: 2px 4px; border-radius: 3px; }
        table { border-collapse: collapse; width: 100%; margin: 1em 0; }
        th, td { border: 1px solid ${isDark ? '#334155' : '#ddd'}; padding: 12px; text-align: left; }
        th { ${isDark ? 'background: #0f3460;' : 'background: #f8f8f8;'} }
        blockquote { border-left: 4px solid ${isDark ? '#334155' : '#ddd'}; margin: 0; padding-left: 1em; color: ${isDark ? '#94a3b8' : '#666'}; }
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

            // ── Batch conversion handler ──
            if (name === "batch_convert") {
                const items: Array<{ markdown: string; title?: string }> = (args as any).items;
                const formats: string[] = (args as any).formats;
                const outputDir: string | undefined = (args as any).output_dir;

                if (!items || !Array.isArray(items) || items.length === 0) {
                    throw new Error("'items' must be a non-empty array of objects with 'markdown' field");
                }
                if (!formats || !Array.isArray(formats) || formats.length === 0) {
                    throw new Error("'formats' must be a non-empty array of format strings");
                }

                const results: Array<{
                    title: string;
                    format: string;
                    success: boolean;
                    file_path?: string;
                    file_size_bytes?: number;
                    content?: string;
                    error?: string;
                }> = [];
                let successful = 0;
                let failed = 0;

                const FORMAT_EXT: Record<string, string> = {
                    txt: 'txt', csv: 'csv', json: 'json', xml: 'xml', xlsx: 'xlsx',
                    latex: 'tex', rtf: 'rtf', docx: 'docx', pdf: 'pdf', html: 'html',
                    md: 'md', email_html: 'html',
                    slack: 'txt', discord: 'txt', jira: 'txt', confluence: 'txt',
                    asciidoc: 'adoc', rst: 'rst', mediawiki: 'txt', bbcode: 'txt',
                    textile: 'txt', orgmode: 'org',
                };

                // Pre-launch browser if any PDF conversions are needed
                let browser: any = null;
                let browserLaunchError: string | null = null;
                const needsBrowser = formats.includes('pdf');
                if (needsBrowser) {
                    try {
                        browser = await getBrowser();
                    } catch (err: any) {
                        // Capture the root cause so per-item PDF failures surface a helpful message
                        browserLaunchError = err?.message || String(err);
                    }
                }

                try {
                    for (const item of items) {
                        const md = item.markdown;
                        const title = item.title || config.default_title || 'document';

                        if (!md) {
                            for (const fmt of formats) {
                                results.push({ title, format: fmt, success: false, error: 'Missing markdown content' });
                                failed++;
                            }
                            continue;
                        }

                        for (const fmt of formats) {
                            try {
                                let content: string | Buffer;
                                let isBinary = false;
                                const ext = FORMAT_EXT[fmt] || fmt;

                                switch (fmt) {
                                    case 'txt':
                                        content = cleanMarkdownText(md);
                                        break;
                                    case 'csv':
                                        content = generateCSV(md);
                                        break;
                                    case 'json':
                                        content = generateJSON(md, title);
                                        break;
                                    case 'xml':
                                        content = generateXML(md, title);
                                        break;
                                    case 'latex':
                                        content = parseMarkdownToLaTeX(md);
                                        break;
                                    case 'rtf':
                                        content = parseMarkdownToRTF(md);
                                        break;
                                    case 'html': {
                                        const htmlFile = await unified().use(remarkParse).use(remarkGfm).use(remarkRehype).use(rehypeKatex).use(rehypeStringify).process(md);
                                        const isDark = config.html_theme === 'dark';
                                        const themeStyles = isDark ? 'background: #1a1a2e; color: #e0e0e0;' : 'background: white; color: black;';
                                        const preStyles = isDark ? 'background: #16213e; color: #e0e0e0;' : 'background: #f4f4f4;';
                                        const thStyles = isDark ? 'background-color: #0f3460;' : 'background-color: #f2f2f2;';
                                        const borderColor = isDark ? '#334155' : '#ddd';
                                        const bqColor = isDark ? '#94a3b8' : '#666';
                                        content = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous"><style>body { font-family: system-ui; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto; ${themeStyles} } img { max-width: 100%; } pre { ${preStyles} padding: 15px; border-radius: 5px; overflow-x: auto; } table { border-collapse: collapse; width: 100%; margin: 1em 0; } th, td { border: 1px solid ${borderColor}; padding: 8px; text-align: left; } th { ${thStyles} } blockquote { border-left: 4px solid ${borderColor}; margin: 0; padding-left: 1em; color: ${bqColor}; }</style></head><body>${String(htmlFile)}</body></html>`;
                                        break;
                                    }
                                    case 'md': {
                                        content = md;
                                        break;
                                    }
                                    case 'email_html': {
                                        content = await markdownToEmailHtml(md);
                                        break;
                                    }
                                    case 'docx': {
                                        const { elements, footnotes } = parseMarkdownToDocx(md);
                                        const docOptions: any = { sections: [{ children: elements }] };
                                        if (Object.keys(footnotes).length > 0) {
                                            docOptions.footnotes = footnotes;
                                        }
                                        const doc = new (await import("docx")).Document(docOptions);
                                        content = await Packer.toBuffer(doc);
                                        isBinary = true;
                                        break;
                                    }
                                    case 'xlsx': {
                                        content = generateXLSXIndex(md);
                                        isBinary = true;
                                        break;
                                    }
                                    case 'pdf': {
                                        if (!browser) {
                                            throw new Error(
                                                browserLaunchError
                                                    ? `Browser launch failed — cannot generate PDF. Cause: ${browserLaunchError}`
                                                    : 'Browser launch failed — cannot generate PDF. Ensure Chrome/Chromium is available.'
                                            );
                                        }
                                        const htmlFile = await unified().use(remarkParse).use(remarkGfm).use(remarkRehype).use(rehypeKatex).use(rehypeStringify).process(md);
                                        const isDark = config.html_theme === 'dark';
                                        const themeStyles = isDark ? 'background: #1a1a2e; color: #e0e0e0;' : 'background: white; color: black;';
                                        const preStyles = isDark ? 'background: #16213e; color: #e0e0e0;' : 'background: #f4f4f4;';
                                        const thStyles = isDark ? 'background-color: #0f3460;' : 'background-color: #f2f2f2;';
                                        const borderColor = isDark ? '#334155' : '#ddd';
                                        const bqColor = isDark ? '#94a3b8' : '#666';
                                        const htmlDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css" integrity="sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV" crossorigin="anonymous"><style>body { font-family: system-ui; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto; ${themeStyles} } img { max-width: 100%; } pre { ${preStyles} padding: 15px; border-radius: 5px; overflow-x: auto; } table { border-collapse: collapse; width: 100%; margin: 1em 0; } th, td { border: 1px solid ${borderColor}; padding: 8px; text-align: left; } th { ${thStyles} } blockquote { border-left: 4px solid ${borderColor}; margin: 0; padding-left: 1em; color: ${bqColor}; }</style></head><body>${String(htmlFile)}</body></html>`;
                                        const page = await browser.newPage();
                                        try {
                                            await page.setContent(htmlDoc);
                                            const m = config.pdf_margin;
                                            content = Buffer.from(await page.pdf({ format: config.pdf_page_format as any, margin: { top: m, right: m, bottom: m, left: m } }));
                                            isBinary = true;
                                        } finally {
                                            await page.close();
                                        }
                                        break;
                                    }
                                    // Platform converters
                                    case 'slack': content = markdownToSlack(md); break;
                                    case 'discord': content = markdownToDiscord(md); break;
                                    case 'jira': content = markdownToJira(md); break;
                                    case 'confluence': content = markdownToConfluence(md); break;
                                    case 'asciidoc': content = markdownToAsciiDoc(md); break;
                                    case 'rst': content = markdownToRST(md); break;
                                    case 'mediawiki': content = markdownToMediaWiki(md); break;
                                    case 'bbcode': content = markdownToBBCode(md); break;
                                    case 'textile': content = markdownToTextile(md); break;
                                    case 'orgmode': content = markdownToOrgMode(md); break;
                                    default:
                                        throw new Error(`Unsupported format: ${fmt}`);
                                }

                                // Handle output
                                if (outputDir) {
                                    const dirPath = path.join(outputDir, title);
                                    const filePath = path.join(dirPath, `${title}.${ext}`);
                                    await fs.mkdir(dirPath, { recursive: true });
                                    await fs.writeFile(filePath, content);
                                    const stats = await fs.stat(filePath);
                                    results.push({
                                        title, format: fmt, success: true,
                                        file_path: filePath, file_size_bytes: stats.size,
                                    });
                                } else if (isBinary) {
                                    const buf = Buffer.isBuffer(content) ? content : Buffer.from(content);
                                    results.push({
                                        title, format: fmt, success: true,
                                        file_size_bytes: buf.length,
                                        content: buf.toString('base64').substring(0, 1000) + (buf.toString('base64').length > 1000 ? '...' : ''),
                                    });
                                } else {
                                    results.push({
                                        title, format: fmt, success: true,
                                        content: content as string,
                                    });
                                }
                                successful++;
                            } catch (convErr: any) {
                                results.push({ title, format: fmt, success: false, error: convErr.message });
                                failed++;
                            }
                        }
                    }
                } finally {
                    if (browser) {
                        await browser.close();
                    }
                }

                const summary = {
                    total_conversions: successful + failed,
                    successful,
                    failed,
                    items_processed: items.length,
                    formats_requested: formats.length,
                };

                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({ summary, results }, null, 2),
                    }],
                };
            }

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

async function getOrCreateInstance(sessionId: string, config?: ServerConfig): Promise<McpInstance> {
    if (instances.has(sessionId)) {
        const instance = instances.get(sessionId)!;
        instance.isNew = false;
        instance.lastUsed = Date.now();
        return instance;
    }

    const sessionConfig = config || getDefaultConfig();

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

    setupServerHandlers(server, sessionConfig);
    await server.connect(transport);

    const instance = { server, transport, isNew: true, lastUsed: Date.now(), config: sessionConfig };
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

    // Read session configuration from query params (forwarded by Smithery Gateway)
    const sessionConfig: ServerConfig = {
        pdf_page_format: (req.query.pdf_page_format as string) || 'A4',
        pdf_margin: (req.query.pdf_margin as string) || '20mm',
        html_theme: (req.query.html_theme as string) || 'light',
        default_title: (req.query.default_title as string) || 'document',
        max_input_bytes: Number(req.query.max_input_bytes) || 10 * 1024 * 1024,
    };

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
        const instance = await getOrCreateInstance(sessionId, sessionConfig);

        // Anti-409: Close existing stream if this is a new GET for the same session
        if (req.method === 'GET' && isEventStream) {
            instance.transport.closeStandaloneSSEStream();
        }

        const body = (req.method === 'POST' || req.method === 'PUT') ? req.body : undefined;

        // Detect whether this POST carries an actual MCP "initialize" message.
        // If so, let the SDK handle the handshake naturally — forcing _initialized
        // before the transport processes "initialize" causes it to reject with 400.
        const isInitializeMsg = body && typeof body === 'object' && !Array.isArray(body) && body.method === 'initialize';

        // Force initialization for non-initialize requests on a new instance,
        // because in a serverless environment we can't rely on the client
        // hitting the same instance for the 'initialize' message.
        if (instance.isNew && !isInitializeMsg) {
            console.log(`[MCP] Cold Start/Instance Migration detected for session ${sessionId}. Forcing initialization.`);
            // Access internal SDK property to bypass the initialize handshake requirement
            // in stateless serverless environments. Guarded to survive future SDK changes.
            const srv = instance.server as any;
            if (typeof srv._initialized !== 'undefined') {
                srv._initialized = true;
            }
            // CRITICAL: Also bypass the Transport's own initialization gate.
            // WebStandardStreamableHTTPServerTransport.validateSession() checks
            // transport._initialized and transport.sessionId — without these,
            // all non-initialize requests (resources/list, prompts/list, etc.)
            // return 400 "Server not initialized" or 400 "Mcp-Session-Id required".
            const trn = instance.transport as any;
            if (typeof trn._initialized !== 'undefined') {
                trn._initialized = true;
            }
            if (trn.sessionIdGenerator) {
                trn.sessionId = sessionId;
            }
        }
        instance.isNew = false;

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
