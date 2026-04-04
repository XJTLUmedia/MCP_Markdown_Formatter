import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';
import remarkRehype from 'remark-rehype';
import rehypeStringify from 'rehype-stringify';

export async function markdownToEmailHtml(md: string): Promise<string> {
    const htmlFile = await unified()
        .use(remarkParse)
        .use(remarkGfm)
        // @ts-ignore
        .use(remarkRehype)
        // @ts-ignore
        .use(rehypeStringify)
        .process(md);

    let html = String(htmlFile);
    html = inlineEmailStyles(html);

    return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background-color:#f6f6f6;">
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f6f6f6;">
<tr>
<td align="center" style="padding:20px 0;">
<table role="presentation" cellpadding="0" cellspacing="0" width="600" style="background-color:#ffffff;border-radius:4px;border:1px solid #e0e0e0;">
<tr>
<td style="padding:30px 40px;font-family:Arial,Helvetica,sans-serif;font-size:16px;line-height:1.6;color:#333333;">
${html}
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>`;
}

function inlineEmailStyles(html: string): string {
    let out = html;
    out = out.replace(/<h1([^>]*)>/gi, '<h1$1 style="font-family:Arial,Helvetica,sans-serif;font-size:28px;font-weight:bold;color:#1a1a1a;margin:24px 0 12px 0;line-height:1.3;">');
    out = out.replace(/<h2([^>]*)>/gi, '<h2$1 style="font-family:Arial,Helvetica,sans-serif;font-size:24px;font-weight:bold;color:#1a1a1a;margin:20px 0 10px 0;line-height:1.3;">');
    out = out.replace(/<h3([^>]*)>/gi, '<h3$1 style="font-family:Arial,Helvetica,sans-serif;font-size:20px;font-weight:bold;color:#1a1a1a;margin:18px 0 8px 0;line-height:1.3;">');
    out = out.replace(/<p([^>]*)>/gi, '<p$1 style="margin:0 0 16px 0;font-size:16px;line-height:1.6;color:#333333;">');
    out = out.replace(/<a([^>]*?)>/gi, '<a$1 style="color:#0066cc;text-decoration:underline;">');
    out = out.replace(/<pre([^>]*)>/gi, '<pre$1 style="background-color:#f4f4f4;border:1px solid #dddddd;border-radius:4px;padding:12px;overflow-x:auto;font-family:Consolas,Monaco,monospace;font-size:14px;line-height:1.4;margin:16px 0;">');
    out = out.replace(/<code([^>]*)>/gi, (match, attrs) => {
        if (match.includes('style=')) return match;
        return `<code${attrs} style="background-color:#f0f0f0;border:1px solid #e0e0e0;border-radius:3px;padding:2px 6px;font-family:Consolas,Monaco,monospace;font-size:14px;">`;
    });
    out = out.replace(/<table([^>]*)>/gi, '<table$1 style="border-collapse:collapse;width:100%;margin:16px 0;border:1px solid #dddddd;">');
    out = out.replace(/<th([^>]*)>/gi, '<th$1 style="border:1px solid #dddddd;padding:10px 12px;text-align:left;background-color:#f8f8f8;font-weight:bold;font-size:14px;">');
    out = out.replace(/<td([^>]*)>/gi, '<td$1 style="border:1px solid #dddddd;padding:10px 12px;text-align:left;font-size:14px;">');
    out = out.replace(/<blockquote([^>]*)>/gi, '<blockquote$1 style="border-left:4px solid #cccccc;margin:16px 0;padding:8px 16px;color:#666666;background-color:#fafafa;">');
    out = out.replace(/<ul([^>]*)>/gi, '<ul$1 style="margin:8px 0 16px 0;padding-left:24px;">');
    out = out.replace(/<ol([^>]*)>/gi, '<ol$1 style="margin:8px 0 16px 0;padding-left:24px;">');
    out = out.replace(/<li([^>]*)>/gi, '<li$1 style="margin:4px 0;line-height:1.6;">');
    out = out.replace(/<hr([^>]*)\/?>/gi, '<hr$1 style="border:none;border-top:1px solid #dddddd;margin:24px 0;">');
    out = out.replace(/<img([^>]*)>/gi, '<img$1 style="max-width:100%;height:auto;border:0;">');
    out = out.replace(/<strong([^>]*)>/gi, '<strong$1 style="font-weight:bold;">');
    out = out.replace(/<em([^>]*)>/gi, '<em$1 style="font-style:italic;">');
    return out;
}
