// ── HTML to Markdown (round-trip import) ─────────────────────────────
// Pure TypeScript implementation, no external dependencies

export function htmlToMarkdown(html: string): string {
    let md = html;

    // Remove doctype, html/head/body wrappers
    md = md.replace(/<!DOCTYPE[^>]*>/gi, '');
    md = md.replace(/<html[^>]*>/gi, '');
    md = md.replace(/<\/html>/gi, '');
    md = md.replace(/<head[\s\S]*?<\/head>/gi, '');
    md = md.replace(/<body[^>]*>/gi, '');
    md = md.replace(/<\/body>/gi, '');

    // Remove scripts and styles
    md = md.replace(/<script[\s\S]*?<\/script>/gi, '');
    md = md.replace(/<style[\s\S]*?<\/style>/gi, '');

    // Convert block elements first (order matters)

    // Pre/code blocks
    md = md.replace(/<pre[^>]*><code[^>]*(?:class="[^"]*language-(\w+)[^"]*")?[^>]*>([\s\S]*?)<\/code><\/pre>/gi,
        (_m, lang, code) => `\n\`\`\`${lang || ''}\n${decodeHtmlEntities(code).trim()}\n\`\`\`\n`
    );
    md = md.replace(/<pre[^>]*>([\s\S]*?)<\/pre>/gi,
        (_m, code) => `\n\`\`\`\n${decodeHtmlEntities(code).trim()}\n\`\`\`\n`
    );

    // Headings
    md = md.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '\n# $1\n');
    md = md.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, '\n## $1\n');
    md = md.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, '\n### $1\n');
    md = md.replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, '\n#### $1\n');
    md = md.replace(/<h5[^>]*>([\s\S]*?)<\/h5>/gi, '\n##### $1\n');
    md = md.replace(/<h6[^>]*>([\s\S]*?)<\/h6>/gi, '\n###### $1\n');

    // Blockquotes
    md = md.replace(/<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi, (_m, content) => {
        return content.trim().split('\n').map((l: string) => '> ' + l.trim()).join('\n') + '\n';
    });

    // Tables
    md = convertHtmlTables(md);

    // Lists
    md = convertHtmlLists(md);

    // Horizontal rules
    md = md.replace(/<hr[^>]*\/?>/gi, '\n---\n');

    // Paragraphs
    md = md.replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, '\n$1\n');

    // Line breaks
    md = md.replace(/<br\s*\/?>/gi, '  \n');

    // Now convert inline elements
    // Bold
    md = md.replace(/<(strong|b)[^>]*>([\s\S]*?)<\/\1>/gi, '**$2**');
    // Italic
    md = md.replace(/<(em|i)[^>]*>([\s\S]*?)<\/\1>/gi, '*$2*');
    // Strikethrough
    md = md.replace(/<(del|s|strike)[^>]*>([\s\S]*?)<\/\1>/gi, '~~$2~~');
    // Highlight / mark
    md = md.replace(/<mark[^>]*>([\s\S]*?)<\/mark>/gi, '==$1==');
    // Keyboard
    md = md.replace(/<kbd[^>]*>([\s\S]*?)<\/kbd>/gi, '`$1`');
    // Superscript / subscript
    md = md.replace(/<sup[^>]*>([\s\S]*?)<\/sup>/gi, '^$1^');
    md = md.replace(/<sub[^>]*>([\s\S]*?)<\/sub>/gi, '~$1~');
    // Abbreviations — just keep the text
    md = md.replace(/<abbr[^>]*>([\s\S]*?)<\/abbr>/gi, '$1');
    // Code
    md = md.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, '`$1`');
    // Links
    md = md.replace(/<a[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');
    // Images
    md = md.replace(/<img[^>]*src="([^"]*)"[^>]*alt="([^"]*)"[^>]*\/?>/gi, '![$2]($1)');
    md = md.replace(/<img[^>]*alt="([^"]*)"[^>]*src="([^"]*)"[^>]*\/?>/gi, '![$1]($2)');
    md = md.replace(/<img[^>]*src="([^"]*)"[^>]*\/?>/gi, '![]($1)');

    // Definition lists
    md = md.replace(/<dl[^>]*>([\s\S]*?)<\/dl>/gi, (_m, content) => {
        let result = '\n';
        const dtRegex = /<dt[^>]*>([\s\S]*?)<\/dt>/gi;
        const ddRegex = /<dd[^>]*>([\s\S]*?)<\/dd>/gi;
        let dtMatch;
        const dts: string[] = [];
        const dds: string[] = [];
        while ((dtMatch = dtRegex.exec(content)) !== null) dts.push(dtMatch[1].replace(/<[^>]+>/g, '').trim());
        let ddMatch;
        while ((ddMatch = ddRegex.exec(content)) !== null) dds.push(ddMatch[1].replace(/<[^>]+>/g, '').trim());
        for (let i = 0; i < Math.max(dts.length, dds.length); i++) {
            if (i < dts.length) result += `${dts[i]}\n`;
            if (i < dds.length) result += `: ${dds[i]}\n`;
        }
        return result + '\n';
    });

    // Details/summary → blockquote with bold summary
    md = md.replace(/<details[^>]*>\s*<summary[^>]*>([\s\S]*?)<\/summary>([\s\S]*?)<\/details>/gi,
        (_m, summary, content) => `\n> **${summary.replace(/<[^>]+>/g, '').trim()}**\n> ${content.replace(/<[^>]+>/g, '').trim()}\n`
    );

    // Figure/figcaption
    md = md.replace(/<figure[^>]*>([\s\S]*?)<\/figure>/gi, (_m, content) => {
        const imgMatch = content.match(/<img[^>]*src="([^"]*)"[^>]*\/?>/i);
        const captionMatch = content.match(/<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i);
        const src = imgMatch ? imgMatch[1] : '';
        const caption = captionMatch ? captionMatch[1].replace(/<[^>]+>/g, '').trim() : '';
        return src ? `\n![${caption}](${src})\n` : '';
    });

    // Task list checkboxes inside list items (already converted to - items)
    md = md.replace(/- <input[^>]*checked[^>]*\/?>\s*/gi, '- [x] ');
    md = md.replace(/- <input[^>]*type="checkbox"[^>]*\/?>\s*/gi, '- [ ] ');

    // Remove remaining HTML tags
    md = md.replace(/<[^>]+>/g, '');

    // Decode HTML entities
    md = decodeHtmlEntities(md);

    // Clean up whitespace
    md = md.replace(/\n{3,}/g, '\n\n');
    md = md.trim() + '\n';

    return md;
}

function convertHtmlTables(html: string): string {
    return html.replace(/<table[^>]*>([\s\S]*?)<\/table>/gi, (_m, tableContent) => {
        const rows: string[][] = [];

        // Extract thead rows
        const theadMatch = tableContent.match(/<thead[^>]*>([\s\S]*?)<\/thead>/i);
        if (theadMatch) {
            const headerCells = extractCells(theadMatch[1], 'th');
            if (headerCells.length === 0) {
                const tdCells = extractCells(theadMatch[1], 'td');
                if (tdCells.length > 0) rows.push(tdCells);
            } else {
                rows.push(headerCells);
            }
        }

        // Extract tbody rows
        const tbodyMatch = tableContent.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i);
        const bodyContent = tbodyMatch ? tbodyMatch[1] : tableContent;
        const rowMatches = bodyContent.match(/<tr[^>]*>([\s\S]*?)<\/tr>/gi);
        if (rowMatches) {
            for (const row of rowMatches) {
                if (theadMatch && row === rowMatches[0] && !tbodyMatch) continue;
                let cells = extractCells(row, 'td');
                if (cells.length === 0) cells = extractCells(row, 'th');
                if (cells.length > 0) {
                    rows.push(cells);
                }
            }
        }

        if (rows.length === 0) return '';

        let result = '\n';
        result += '| ' + rows[0].join(' | ') + ' |\n';
        result += '| ' + rows[0].map(() => '---').join(' | ') + ' |\n';
        for (let i = 1; i < rows.length; i++) {
            // Pad to match column count
            while (rows[i].length < rows[0].length) rows[i].push('');
            result += '| ' + rows[i].join(' | ') + ' |\n';
        }
        return result;
    });
}

function extractCells(rowHtml: string, tag: string): string[] {
    const cells: string[] = [];
    const regex = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, 'gi');
    let match;
    while ((match = regex.exec(rowHtml)) !== null) {
        cells.push(match[1].replace(/<[^>]+>/g, '').trim());
    }
    return cells;
}

function convertHtmlLists(html: string): string {
    // Process nested lists from inside out (multiple passes)
    let out = html;
    for (let pass = 0; pass < 5; pass++) {
        // Unordered lists
        out = out.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, (_m, content) => {
            const items = extractListItems(content);
            return '\n' + items.map(item => '- ' + item.trim()).join('\n') + '\n';
        });

        // Ordered lists
        out = out.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, (_m, content) => {
            const items = extractListItems(content);
            return '\n' + items.map((item, idx) => `${idx + 1}. ` + item.trim()).join('\n') + '\n';
        });
    }
    return out;
}

function extractListItems(listHtml: string): string[] {
    const items: string[] = [];
    const regex = /<li[^>]*>([\s\S]*?)<\/li>/gi;
    let match;
    while ((match = regex.exec(listHtml)) !== null) {
        let item = match[1].trim();
        // Indent nested list content
        item = item.replace(/\n([-\d])/g, '\n  $1');
        items.push(item);
    }
    return items;
}

function decodeHtmlEntities(text: string): string {
    return text
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&nbsp;/g, ' ')
        .replace(/&#(\d+);/g, (_m, code) => String.fromCharCode(parseInt(code)))
        .replace(/&#x([0-9a-f]+);/gi, (_m, code) => String.fromCharCode(parseInt(code, 16)));
}
