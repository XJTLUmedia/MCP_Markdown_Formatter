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

// Instance interface
interface McpInstance {
    server: Server;
    transport: any; // StreamableHTTPServerTransport
}

// Global registry of active instances in this warm lambda
const instances = new Map<string, McpInstance>();

// Shared setup for all instances
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

function setupServerHandlers(server: Server) {
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

}

import { WebStandardStreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js";

async function getOrCreateInstance(sessionId: string): Promise<McpInstance> {
    if (instances.has(sessionId)) {
        return instances.get(sessionId)!;
    }

    const transport = new WebStandardStreamableHTTPServerTransport({
        sessionIdGenerator: () => sessionId,
    });

    const server = new Server(
        { name: "markdown-formatter-mcp", version: "1.0.0" },
        { capabilities: { tools: {} } }
    );

    setupServerHandlers(server);
    await server.connect(transport);

    const instance = { server, transport };
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

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    const providedSessionId = (req.query.sessionId as string) || (req.headers['mcp-session-id'] as string);
    const sessionId = providedSessionId || `s_${Math.random().toString(36).substring(2, 10)}`;
    res.setHeader('mcp-session-id', sessionId);

    const isEventStream =
        req.headers.accept?.includes('text/event-stream') ||
        req.headers['mcp-protocol-version'] ||
        req.query.sessionId;

    if (req.method === 'GET' && !isEventStream) {
        res.status(200).setHeader('Content-Type', 'text/html').send(`...Help HTML...`);
        return;
    }

    try {
        const instance = await getOrCreateInstance(sessionId);

        if (req.method === 'GET' && isEventStream) {
            instance.transport.closeStandaloneSSEStream();

            // Send immediate 200 OK headers to fix TLS/Handshake timing out on Vercel
            res.status(200).setHeader('Content-Type', 'text/event-stream');
            res.write(': heartbeat\n\n');
        }

        const body = (req.method === 'POST' || req.method === 'PUT') ? req.body : undefined;

        // Manual bridge to WebStandard API since Vercel Request/Response are slightly different
        const url = new URL(req.url!, `http://${req.headers.host}`);
        const webRequest = new Request(url, {
            method: req.method,
            headers: req.headers as any,
            body: body ? JSON.stringify(body) : undefined
        });

        const webResponse = await instance.transport.handleRequest(webRequest);

        // Copy headers and status
        res.status(webResponse.status);
        webResponse.headers.forEach((v, k) => res.setHeader(k, v));

        if (webResponse.body) {
            const reader = webResponse.body.getReader();
            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                res.write(value);
            }
        }
        res.end();
    } catch (error: any) {
        console.error("[MCP] Error:", error);
        if (!res.headersSent) res.status(500).json({ error: error.message });
    }
}

