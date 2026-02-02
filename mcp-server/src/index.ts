import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
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
import puppeteer from 'puppeteer';
import * as fs from 'fs/promises';
import * as path from 'path';
import {
    stripMarkdown,
    parseMarkdownToRTF,
    parseMarkdownToDocx,
    parseMarkdownToLaTeX,
    generateCSV,
    generateJSON,
    generateXML,
    generateXLSXIndex
} from "../../src/utils/core-exports.js";
import { Packer } from "docx";

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

server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: [
            {
                name: "harmonize_markdown",
                description: "Standardize markdown syntax (headers, list markers, etc.)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_txt",
                description: "Convert Markdown to Plain Text (strips formatting)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_rtf",
                description: "Convert Markdown to RTF (Rich Text Format)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_latex",
                description: "Convert Markdown to LaTeX",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_docx",
                description: "Convert Markdown to DOCX (Word)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_pdf",
                description: "Convert Markdown to PDF (uses Puppeteer)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_image",
                description: "Convert Markdown to PNG Image (uses Puppeteer)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_csv",
                description: "Extract tables from Markdown to CSV",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_json",
                description: "Convert Markdown to JSON structure",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    title: z.string().optional(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_xml",
                description: "Convert Markdown to XML",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    title: z.string().optional(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_xlsx",
                description: "Convert Markdown tables to Excel (XLSX)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_html",
                description: "Convert Markdown to HTML",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "convert_to_md",
                description: "Export original Markdown content (with optional harmonization)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    harmonize: z.boolean().optional(),
                    output_path: z.string().optional(),
                })),
            },
            {
                name: "generate_html",
                description: "Generate a complete HTML document from Markdown with inline styles (returns full HTML string)",
                inputSchema: zodSchemaToToolInput(z.object({
                    markdown: z.string(),
                    title: z.string().optional(),
                })),
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
            const txt = stripMarkdown(markdown);
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

            const browser = await puppeteer.launch({ headless: true });
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

function zodSchemaToToolInput(schema: z.ZodType<any>): any {
    const shape = (schema as any).shape;
    const properties: any = {};
    const required: string[] = [];

    for (const key in shape) {
        const field = shape[key];
        const isOptional = field.isOptional?.() || field instanceof z.ZodOptional;

        properties[key] = { type: "string" };

        if (!isOptional) {
            required.push(key);
        }
    }

    return {
        type: "object",
        properties,
        required
    };
}

const transport = new StdioServerTransport();
await server.connect(transport);
