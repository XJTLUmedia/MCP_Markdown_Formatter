import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
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



// Helper to handle output
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

function zodSchemaToToolInput(schema: z.ZodType<any>): any {
    const shape = (schema as any).shape;
    const properties: any = {};
    const required: string[] = [];
    for (const key in shape) {
        const field = shape[key];
        const isOptional = field.isOptional?.() || field instanceof z.ZodOptional;
        properties[key] = { type: "string" };
        if (!isOptional) required.push(key);
    }
    return { type: "object", properties, required };
}

server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: [
            {
                name: "harmonize_markdown",
                description: "Standardize markdown syntax",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_txt",
                description: "Convert Markdown to Plain Text",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_rtf",
                description: "Convert Markdown to RTF",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_latex",
                description: "Convert Markdown to LaTeX",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_docx",
                description: "Convert Markdown to DOCX",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_pdf",
                description: "Convert Markdown to PDF",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_image",
                description: "Convert Markdown to PNG Image",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_csv",
                description: "Extract tables from Markdown to CSV",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_json",
                description: "Convert Markdown to JSON structure",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), title: z.string().optional(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_xml",
                description: "Convert Markdown to XML",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), title: z.string().optional(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_xlsx",
                description: "Convert Markdown tables to Excel",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_html",
                description: "Convert Markdown to HTML",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), output_path: z.string().optional() })),
            },
            {
                name: "convert_to_md",
                description: "Export original Markdown content",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), harmonize: z.boolean().optional(), output_path: z.string().optional() })),
            },
            {
                name: "generate_html",
                description: "Generate complete HTML document",
                inputSchema: zodSchemaToToolInput(z.object({ markdown: z.string(), title: z.string().optional() })),
            }
        ],
    };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
        const { name, arguments: args } = request.params;
        const markdown = (args as any).markdown;
        const outputPath = (args as any).output_path;

        if (!markdown) throw new Error("Markdown content is required");

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
            const htmlDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css"><style>body { font-family: system-ui; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto; }</style></head><body>${String(htmlFile)}</body></html>`;

            if (name === "convert_to_html") return handleOutput(htmlDoc, outputPath);

            const browser = await getBrowser();
            const page = await browser.newPage();
            await page.setContent(htmlDoc);
            let resultBuffer: Buffer;

            if (name === "convert_to_pdf") {
                resultBuffer = Buffer.from(await page.pdf({ format: 'A4' }));
                await browser.close();
                return handleOutput(resultBuffer, outputPath, { format: 'pdf', description: 'PDF document' });
            } else {
                resultBuffer = Buffer.from(await page.screenshot({ fullPage: true, encoding: 'binary' }));
                await browser.close();
                return handleOutput(resultBuffer, outputPath, { format: 'png', description: 'PNG image' });
            }
        }

        if (name === "convert_to_md") {
            if (!(args as any).harmonize) return handleOutput(markdown, outputPath);
            const file = await unified().use(remarkParse).use(remarkGfm).use(remarkMath).use(remarkStringify, { bullet: '-', fence: '`', fences: true, incrementListMarker: true, listItemIndent: 'one' }).process(markdown);
            return handleOutput(String(file), outputPath);
        }

        if (name === "generate_html") {
            const htmlFile = await unified().use(remarkParse).use(remarkGfm).use(remarkRehype).use(rehypeKatex).use(rehypeStringify).process(markdown);
            const htmlDoc = `<!DOCTYPE html><html><head><title>${(args as any).title || 'Doc'}</title></head><body>${String(htmlFile)}</body></html>`;
            return { content: [{ type: "text", text: htmlDoc }] };
        }

        throw new Error(`Unknown tool: ${name}`);
    } catch (error: any) {
        return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
    }
});

import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

// Keep server and transport references at the top level to benefit from Vercel's warm starts
let transport: StreamableHTTPServerTransport | null = null;
let isServerConnected = false;

export default async function handler(req: VercelRequest, res: VercelResponse) {
    // Add CORS and Streaming headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept');
    res.setHeader('X-Accel-Buffering', 'no'); // Disable buffering for SSE streaming

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    // Lazy initialization of transport and server connection
    if (!transport) {
        try {
            transport = new StreamableHTTPServerTransport({
                // Generate unique session IDs to avoid "Conflict: Only one SSE stream is allowed per session"
                sessionIdGenerator: () => Math.random().toString(36).substring(2, 15),
            });
            await server.connect(transport);
            isServerConnected = true;
        } catch (error) {
            console.error("Failed to initialize MCP server:", error);
            res.status(500).json({ error: "Internal Server Error during initialization" });
            return;
        }
    }

    // Friendly message for browser visits (standard GET without event-stream header)
    const isEventStream = req.headers.accept?.includes('text/event-stream') || req.query.sessionId;
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
                    
                    <h2 style="font-size: 1.125rem; margin-top: 32px; color: #94a3b8;">Setup Instructions</h2>
                    <p>To use this server, add it to your <code>claude_desktop_config.json</code>:</p>
                    <pre>{
  "mcpServers": {
    "markdown-formatter": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/inspector", "https://ai-answer-copier.vercel.app/api/mcp"]
    }
  }
}</pre>
                </div>
            </body>
            </html>
        `);
        return;
    }

    try {
        // Handle request via SDK transport
        await transport.handleRequest(req, res, req.body);
    } catch (error: any) {
        console.error("MCP Error:", error);
        if (!res.headersSent) {
            res.status(500).json({ error: error.message });
        }
    }
}

