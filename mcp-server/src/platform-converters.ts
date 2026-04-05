// ── Shared pre-processing helpers ────────────────────────────────────
function collectFootnotes(text: string): { cleaned: string; footnoteMap: Record<string, string> } {
    const footnoteMap: Record<string, string> = {};
    const cleaned = text.replace(/^\[\^(\w+)\]:\s*(.+)$/gm, (_m, label, content) => {
        footnoteMap[label] = content;
        return '';
    });
    return { cleaned, footnoteMap };
}

function appendEndnotes(
    text: string,
    footnoteMap: Record<string, string>,
    formatRef: (n: number) => string = (n) => `[${n}]`,
    separator: string = '\n\n---\n'
): string {
    const labels = Object.keys(footnoteMap);
    if (labels.length === 0) return text;
    let out = text;
    // Replace refs with formatted numbers
    for (let i = 0; i < labels.length; i++) {
        const escaped = labels[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`\\[\\^${escaped}\\]`, 'g'), formatRef(i + 1));
    }
    // Append endnotes section
    out += separator;
    for (let i = 0; i < labels.length; i++) {
        out += `\n${i + 1}. ${footnoteMap[labels[i]]}`;
    }
    return out;
}

// ── Slack mrkdwn ─────────────────────────────────────────────────────
export function markdownToSlack(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → *text* (bold as closest Slack approximation)
    out = out.replace(/==([^=]+)==/g, '*$1*');
    // Bold: **text** → *text*
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    // Italic: *text* or _text_ → _text_
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    // Strikethrough: ~~text~~ → ~text~
    out = out.replace(/~~([^~]+)~~/g, '~$1~');
    // Inline code stays as backtick
    // Code blocks: ```lang\n...\n``` stays the same (Slack supports it)
    // Links: [text](url) → <url|text>
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<$2|$1>');
    // Images: ![alt](url) → <url|alt>
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<$2|$1>');
    // Headers: # text → *text*
    out = out.replace(/^#{1,6}\s+(.+)$/gm, '*$1*');
    // Task lists
    out = out.replace(/^\s*[-*+]\s+\[x\]\s+(.+)$/gim, '☑ $1');
    out = out.replace(/^\s*[-*+]\s+\[ \]\s+(.+)$/gm, '☐ $1');
    // Blockquotes: > text → > text (Slack supports >)
    // Ordered lists: 1. text stays the same
    // Unordered lists: - text stays the same
    // Horizontal rule
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '───────────────');
    return appendEndnotes(out, footnoteMap);
}

// ── Discord markdown ─────────────────────────────────────────────────
export function markdownToDiscord(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → **text** (bold as closest Discord approximation)
    out = out.replace(/==([^=]+)==/g, '**$1**');
    // Discord supports most standard markdown, just a few tweaks:
    // Headers: # text → **__text__** (Discord renders # only in certain contexts)
    out = out.replace(/^# (.+)$/gm, '**__$1__**');
    out = out.replace(/^## (.+)$/gm, '**$1**');
    out = out.replace(/^### (.+)$/gm, '__$1__');
    out = out.replace(/^#{4,6}\s+(.+)$/gm, '*$1*');
    // Task lists
    out = out.replace(/^\s*[-*+]\s+\[x\]\s+(.+)$/gim, '- ☑ $1');
    out = out.replace(/^\s*[-*+]\s+\[ \]\s+(.+)$/gm, '- ☐ $1');
    // Block quotes: > works in Discord
    // Code blocks: ``` works in Discord
    // Horizontal rules are not rendered in Discord
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '');
    return appendEndnotes(out, footnoteMap);
}

// ── JIRA wiki markup ─────────────────────────────────────────────────
export function markdownToJira(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → {color:yellow}text{color}
    out = out.replace(/==([^=]+)==/g, '{color:yellow}$1{color}');
    // Headers: # → h1., ## → h2., etc.
    out = out.replace(/^######\s+(.+)$/gm, 'h6. $1');
    out = out.replace(/^#####\s+(.+)$/gm, 'h5. $1');
    out = out.replace(/^####\s+(.+)$/gm, 'h4. $1');
    out = out.replace(/^###\s+(.+)$/gm, 'h3. $1');
    out = out.replace(/^##\s+(.+)$/gm, 'h2. $1');
    out = out.replace(/^#\s+(.+)$/gm, 'h1. $1');
    // Bold: **text** → *text*
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    // Italic: *text* → _text_
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    // Strikethrough: ~~text~~ → -text-
    out = out.replace(/~~([^~]+)~~/g, '-$1-');
    // Inline code: `code` → {{code}}
    out = out.replace(/`([^`]+)`/g, '{{$1}}');
    // Code blocks: ```lang → {code:lang} / ``` → {code}
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `{code:${lang}}\n${code.trimEnd()}\n{code}` : `{code}\n${code.trimEnd()}\n{code}`
    );
    // Links: [text](url) → [text|url]
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[$1|$2]');
    // Images: ![alt](url) → !url!
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '!$2!');
    // Unordered list: - text → * text  (JIRA uses *)
    out = out.replace(/^(\s*)[-+]\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' ';
    });
    // Ordered list: 1. text → # text
    out = out.replace(/^(\s*)\d+\.\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '#'.repeat(level) + ' ';
    });
    // Task lists
    out = out.replace(/^(\s*)[-*+]\s+\[x\]\s+/gim, (_m: string, indent: string) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' (/) ';
    });
    out = out.replace(/^(\s*)[-*+]\s+\[ \]\s+/gm, (_m: string, indent: string) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' (x) ';
    });
    // Blockquote: > text → {quote}text{quote}
    out = out.replace(/^>\s+(.+)$/gm, '{quote}$1{quote}');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----');
    // Tables: | A | B | → || A || B || for header, | a | b | for body
    const lines = out.split('\n');
    const result: string[] = [];
    let headerDone = false;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
            // Skip separator rows
            if (/^\|[\s-:]+\|/.test(line.trim())) {
                headerDone = true;
                continue;
            }
            if (!headerDone) {
                // Convert header cells: | A | B | → || A || B ||
                result.push(line.replace(/\|/g, '||'));
            } else {
                result.push(line);
            }
        } else {
            headerDone = false;
            result.push(line);
        }
    }
    return appendEndnotes(result.join('\n'), footnoteMap, (n) => `^[${n}]^`, '\n\n----\n');
}

// ── Confluence wiki markup ───────────────────────────────────────────
export function markdownToConfluence(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → {color:yellow}text{color}
    out = out.replace(/==([^=]+)==/g, '{color:yellow}$1{color}');
    // Headers
    out = out.replace(/^######\s+(.+)$/gm, 'h6. $1');
    out = out.replace(/^#####\s+(.+)$/gm, 'h5. $1');
    out = out.replace(/^####\s+(.+)$/gm, 'h4. $1');
    out = out.replace(/^###\s+(.+)$/gm, 'h3. $1');
    out = out.replace(/^##\s+(.+)$/gm, 'h2. $1');
    out = out.replace(/^#\s+(.+)$/gm, 'h1. $1');
    // Bold, Italic, Strikethrough (same as JIRA)
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    out = out.replace(/~~([^~]+)~~/g, '-$1-');
    // Inline code
    out = out.replace(/`([^`]+)`/g, '{{$1}}');
    // Code blocks
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `{code:language=${lang}}\n${code.trimEnd()}\n{code}` : `{code}\n${code.trimEnd()}\n{code}`
    );
    // Links
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[$1|$2]');
    // Images
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '!$2!');
    // Lists (same as JIRA)
    out = out.replace(/^(\s*)[-+]\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' ';
    });
    out = out.replace(/^(\s*)\d+\.\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '#'.repeat(level) + ' ';
    });
    // Blockquote
    out = out.replace(/^>\s+(.+)$/gm, '{quote}$1{quote}');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----');
    // Task lists
    out = out.replace(/^(\s*)[-*+]\s+\[x\]\s+/gim, (_m: string, indent: string) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' (/) ';
    });
    out = out.replace(/^(\s*)[-*+]\s+\[ \]\s+/gm, (_m: string, indent: string) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' (x) ';
    });
    // Info/note panels from HTML comments
    out = out.replace(/<!--\s*note:\s*([\s\S]*?)-->/gi, '{note}$1{note}');
    out = out.replace(/<!--\s*info:\s*([\s\S]*?)-->/gi, '{info}$1{info}');
    return appendEndnotes(out, footnoteMap, (n) => `^[${n}]^`, '\n\n----\n');
}

// ── AsciiDoc ─────────────────────────────────────────────────────────
export function markdownToAsciiDoc(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → [.mark]#text#
    out = out.replace(/==([^=]+)==/g, '[.mark]#$1#');
    // Footnote refs → native AsciiDoc inline footnotes
    for (const [label, text] of Object.entries(footnoteMap)) {
        const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`\\[\\^${escaped}\\]`, 'g'), `footnote:[${text}]`);
    }
    // Headers: # → =, ## → ==, etc.
    out = out.replace(/^######\s+(.+)$/gm, '====== $1');
    out = out.replace(/^#####\s+(.+)$/gm, '===== $1');
    out = out.replace(/^####\s+(.+)$/gm, '==== $1');
    out = out.replace(/^###\s+(.+)$/gm, '=== $1');
    out = out.replace(/^##\s+(.+)$/gm, '== $1');
    out = out.replace(/^#\s+(.+)$/gm, '= $1');
    // Bold: **text** → *text*
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    // Italic: *text* → _text_
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    // Inline code stays as backtick (AsciiDoc uses + but backtick also works)
    out = out.replace(/`([^`]+)`/g, '`$1`');
    // Code blocks: ```lang → [source,lang]\n----\n...\n----
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) => {
        const attr = lang ? `[source,${lang}]\n` : '';
        return `${attr}----\n${code.trimEnd()}\n----`;
    });
    // Links: [text](url) → url[text]
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$2[$1]');
    // Images: ![alt](url) → image::url[alt]
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, 'image::$2[$1]');
    // Blockquote: > text → [quote]\ntext\n
    const lines = out.split('\n');
    const result: string[] = [];
    let inQuote = false;
    for (const line of lines) {
        if (line.startsWith('> ')) {
            if (!inQuote) {
                result.push('[quote]');
                result.push('____');
                inQuote = true;
            }
            result.push(line.slice(2));
        } else {
            if (inQuote) {
                result.push('____');
                inQuote = false;
            }
            result.push(line);
        }
    }
    if (inQuote) result.push('____');
    out = result.join('\n');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, "'''");
    // Task lists
    out = out.replace(/^(\s*)[-*+]\s+\[x\]\s+/gim, '* [*] ');
    out = out.replace(/^(\s*)[-*+]\s+\[ \]\s+/gm, '* [ ] ');
    // Unordered list marker stays as *
    out = out.replace(/^(\s*)[-+]\s+/gm, '* ');
    // Ordered list: 1. → .
    out = out.replace(/^\d+\.\s+/gm, '. ');
    // Tables
    out = convertMarkdownTableToAsciiDoc(out);
    return out;
}

function convertMarkdownTableToAsciiDoc(text: string): string {
    const lines = text.split('\n');
    const result: string[] = [];
    let inTable = false;
    for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(trimmed)) continue;
            const cells = trimmed.split('|').filter(c => c.trim()).map(c => c.trim());
            if (!inTable) {
                result.push(`[cols="${cells.map(() => '1').join(',')}",options="header"]`);
                result.push('|===');
                inTable = true;
            }
            result.push(cells.map(c => '| ' + c).join(' '));
        } else {
            if (inTable) { result.push('|==='); inTable = false; }
            result.push(line);
        }
    }
    if (inTable) result.push('|===');
    return result.join('\n');
}

// ── reStructuredText ─────────────────────────────────────────────────
export function markdownToRST(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → **text** (RST has no native highlight)
    out = out.replace(/==([^=]+)==/g, '**$1**');
    // Footnote refs → RST native footnotes [#label]_
    const fnLabels = Object.keys(footnoteMap);
    for (const label of fnLabels) {
        const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`\\[\\^${escaped}\\]`, 'g'), ` [#fn_${label}]_`);
    }
    // Code blocks: ```lang → .. code-block:: lang
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) => {
        const directive = lang ? `.. code-block:: ${lang}` : '.. code-block::';
        const indented = code.split('\n').map((l: string) => '   ' + l).join('\n').trimEnd();
        return `${directive}\n\n${indented}`;
    });
    // Bold: **text** → **text** (same in RST)
    // Italic: *text* → *text* (same in RST)
    // Inline code: `code` → ``code``
    out = out.replace(/(?<!`)(`[^`]+`)(?!`)/g, '`$1`');
    // Links: [text](url) → `text <url>`_
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '`$1 <$2>`_');
    // Images: ![alt](url) → .. image:: url\n   :alt: alt
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '.. image:: $2\n   :alt: $1');
    // Headers: use underline characters
    const lines = out.split('\n');
    const result: string[] = [];
    for (const line of lines) {
        const h1 = line.match(/^#\s+(.+)$/);
        const h2 = line.match(/^##\s+(.+)$/);
        const h3 = line.match(/^###\s+(.+)$/);
        const h4 = line.match(/^####\s+(.+)$/);
        if (h1) {
            result.push(h1[1], '='.repeat(h1[1].length));
        } else if (h2) {
            result.push(h2[1], '-'.repeat(h2[1].length));
        } else if (h3) {
            result.push(h3[1], '~'.repeat(h3[1].length));
        } else if (h4) {
            result.push(h4[1], '^'.repeat(h4[1].length));
        } else {
            result.push(line);
        }
    }
    out = result.join('\n');
    // Blockquote: > text → indented text
    out = out.replace(/^>\s+(.+)$/gm, '   $1');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----------');
    // Ordered list: keep as-is (RST uses #.)
    out = out.replace(/^(\d+)\.\s+/gm, '#. ');
    // Task lists
    out = out.replace(/^-\s+\[x\]\s+/gim, '- ☑ ');
    out = out.replace(/^-\s+\[ \]\s+/gm, '- ☐ ');
    // Tables
    out = convertMarkdownTableToRST(out);
    // Append RST footnote definitions
    if (fnLabels.length > 0) {
        out += '\n\n';
        for (const label of fnLabels) {
            out += `.. [#fn_${label}] ${footnoteMap[label]}\n`;
        }
    }
    return out;
}

function convertMarkdownTableToRST(text: string): string {
    const lines = text.split('\n');
    const result: string[] = [];
    const tableRows: string[][] = [];
    let isCollecting = false;
    for (let i = 0; i < lines.length; i++) {
        const trimmed = lines[i].trim();
        if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(trimmed)) continue;
            const cells = trimmed.split('|').filter(c => c.trim()).map(c => c.trim());
            tableRows.push(cells);
            isCollecting = true;
        } else {
            if (isCollecting && tableRows.length > 0) {
                result.push(...renderRSTTable(tableRows));
                tableRows.length = 0;
                isCollecting = false;
            }
            result.push(lines[i]);
        }
    }
    if (tableRows.length > 0) result.push(...renderRSTTable(tableRows));
    return result.join('\n');
}

function renderRSTTable(rows: string[][]): string[] {
    if (rows.length === 0) return [];
    const colCount = Math.max(...rows.map(r => r.length));
    const colWidths = Array(colCount).fill(3);
    for (const row of rows) {
        for (let j = 0; j < row.length; j++) {
            colWidths[j] = Math.max(colWidths[j], (row[j] || '').length + 2);
        }
    }
    const sep = '+' + colWidths.map(w => '-'.repeat(w)).join('+') + '+';
    const headSep = '+' + colWidths.map(w => '='.repeat(w)).join('+') + '+';
    const result: string[] = [sep];
    for (let i = 0; i < rows.length; i++) {
        const line = '|' + colWidths.map((w, j) => (' ' + (rows[i][j] || '') + ' ').padEnd(w)).join('|') + '|';
        result.push(line);
        result.push(i === 0 ? headSep : sep);
    }
    return result;
}

// ── MediaWiki markup ─────────────────────────────────────────────────
export function markdownToMediaWiki(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → <mark>text</mark>
    out = out.replace(/==([^=]+)==/g, '<mark>$1</mark>');
    // Footnote refs → MediaWiki native <ref>text</ref>
    for (const [label, text] of Object.entries(footnoteMap)) {
        const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`\\[\\^${escaped}\\]`, 'g'), `<ref>${text}</ref>`);
    }
    // Headers: # text → == text ==
    out = out.replace(/^######\s+(.+)$/gm, '======= $1 =======');
    out = out.replace(/^#####\s+(.+)$/gm, '====== $1 ======');
    out = out.replace(/^####\s+(.+)$/gm, '===== $1 =====');
    out = out.replace(/^###\s+(.+)$/gm, '==== $1 ====');
    out = out.replace(/^##\s+(.+)$/gm, '=== $1 ===');
    out = out.replace(/^#\s+(.+)$/gm, '== $1 ==');
    // Bold: **text** → '''text'''
    out = out.replace(/\*\*([^*]+)\*\*/g, "'''$1'''");
    // Italic: *text* → ''text''
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, "''$1''");
    // Inline code: `code` → <code>code</code>
    out = out.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Code blocks
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `<syntaxhighlight lang="${lang}">\n${code.trimEnd()}\n</syntaxhighlight>` :
            `<pre>\n${code.trimEnd()}\n</pre>`
    );
    // Links: [text](url) → [url text]
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[$2 $1]');
    // Images: ![alt](url) → [[File:url|alt]]
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '[[File:$2|$1]]');
    // Unordered list: - text → * text
    out = out.replace(/^(\s*)[-+]\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' ';
    });
    // Ordered list
    out = out.replace(/^(\s*)\d+\.\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '#'.repeat(level) + ' ';
    });
    // Blockquote
    out = out.replace(/^>\s+(.+)$/gm, '<blockquote>$1</blockquote>');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----');
    // Tables
    const lines = out.split('\n');
    const result: string[] = [];
    let inTable = false;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line.startsWith('|') && line.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(line)) continue; // skip separator
            const cells = line.split('|').filter(c => c.trim()).map(c => c.trim());
            if (!inTable) {
                result.push('{| class="wikitable"');
                inTable = true;
                // First row as header
                result.push('|-');
                cells.forEach(c => result.push('! ' + c));
            } else {
                result.push('|-');
                cells.forEach(c => result.push('| ' + c));
            }
        } else {
            if (inTable) {
                result.push('|}');
                inTable = false;
            }
            result.push(lines[i]);
        }
    }
    if (inTable) result.push('|}');
    let mwOut = result.join('\n');
    // Append <references/> if there were footnotes
    if (Object.keys(footnoteMap).length > 0) {
        mwOut += '\n\n== References ==\n<references/>';
    }
    return mwOut;
}

// ── BBCode ───────────────────────────────────────────────────────────
export function markdownToBBCode(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → [color=yellow]text[/color]
    out = out.replace(/==([^=]+)==/g, '[color=yellow]$1[/color]');
    // Headers
    out = out.replace(/^#{1,6}\s+(.+)$/gm, '[b][size=5]$1[/size][/b]');
    // Bold
    out = out.replace(/\*\*([^*]+)\*\*/g, '[b]$1[/b]');
    // Italic
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '[i]$1[/i]');
    // Strikethrough
    out = out.replace(/~~([^~]+)~~/g, '[s]$1[/s]');
    // Inline code
    out = out.replace(/`([^`]+)`/g, '[code]$1[/code]');
    // Code blocks
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, _lang, code) =>
        `[code]\n${code.trimEnd()}\n[/code]`
    );
    // Links
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[url=$2]$1[/url]');
    // Images
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '[img]$2[/img]');
    // Blockquote
    out = out.replace(/^>\s+(.+)$/gm, '[quote]$1[/quote]');
    // Task lists
    out = out.replace(/^[-*+]\s+\[x\]\s+(.+)$/gim, '[*]☑ $1');
    out = out.replace(/^[-*+]\s+\[ \]\s+(.+)$/gm, '[*]☐ $1');
    // Unordered list
    out = out.replace(/^[-*+]\s+(.+)$/gm, '[*]$1');
    // Tables
    out = convertMarkdownTableToBBCode(out);
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '[hr]');
    return appendEndnotes(out, footnoteMap, (n) => `[sup][${n}][/sup]`, '\n\n[hr]\n');
}

function convertMarkdownTableToBBCode(text: string): string {
    const lines = text.split('\n');
    const result: string[] = [];
    let inTable = false;
    let isHeader = true;
    for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(trimmed)) { isHeader = false; continue; }
            const cells = trimmed.split('|').filter(c => c.trim()).map(c => c.trim());
            if (!inTable) { result.push('[table]'); inTable = true; }
            if (isHeader) {
                result.push('[tr]' + cells.map(c => `[th]${c}[/th]`).join('') + '[/tr]');
            } else {
                result.push('[tr]' + cells.map(c => `[td]${c}[/td]`).join('') + '[/tr]');
            }
        } else {
            if (inTable) { result.push('[/table]'); inTable = false; isHeader = true; }
            result.push(line);
        }
    }
    if (inTable) result.push('[/table]');
    return result.join('\n');
}

// ── Textile ──────────────────────────────────────────────────────────
export function markdownToTextile(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → %{background:yellow}text%
    out = out.replace(/==([^=]+)==/g, '%{background:yellow}$1%');
    // Footnote refs → Textile native [N]
    const fnLabels = Object.keys(footnoteMap);
    for (let i = 0; i < fnLabels.length; i++) {
        const escaped = fnLabels[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`\\[\\^${escaped}\\]`, 'g'), `[${i + 1}]`);
    }
    // Headers: # → h1., ## → h2.
    out = out.replace(/^######\s+(.+)$/gm, 'h6. $1');
    out = out.replace(/^#####\s+(.+)$/gm, 'h5. $1');
    out = out.replace(/^####\s+(.+)$/gm, 'h4. $1');
    out = out.replace(/^###\s+(.+)$/gm, 'h3. $1');
    out = out.replace(/^##\s+(.+)$/gm, 'h2. $1');
    out = out.replace(/^#\s+(.+)$/gm, 'h1. $1');
    // Bold: **text** → *text*
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    // Italic: *text* → _text_
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    // Strikethrough: ~~text~~ → -text-
    out = out.replace(/~~([^~]+)~~/g, '-$1-');
    // Inline code
    out = out.replace(/`([^`]+)`/g, '@$1@');
    // Code blocks
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, _lang, code) =>
        `bc. ${code.trimEnd()}`
    );
    // Links: [text](url) → "text":url
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '"$1":$2');
    // Images: ![alt](url) → !url(alt)!
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '!$2($1)!');
    // Blockquote
    out = out.replace(/^>\s+(.+)$/gm, 'bq. $1');
    // Task lists
    out = out.replace(/^[-*+]\s+\[x\]\s+/gim, '* ☑ ');
    out = out.replace(/^[-*+]\s+\[ \]\s+/gm, '* ☐ ');
    // Unordered list
    out = out.replace(/^[-+]\s+/gm, '* ');
    // Ordered list
    out = out.replace(/^\d+\.\s+/gm, '# ');
    // Tables
    out = convertMarkdownTableToTextile(out);
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '---');
    // Append Textile footnote definitions
    if (fnLabels.length > 0) {
        out += '\n\n';
        for (let i = 0; i < fnLabels.length; i++) {
            out += `fn${i + 1}. ${footnoteMap[fnLabels[i]]}\n`;
        }
    }
    return out;
}

function convertMarkdownTableToTextile(text: string): string {
    const lines = text.split('\n');
    const result: string[] = [];
    let isHeader = true;
    for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(trimmed)) { isHeader = false; continue; }
            const cells = trimmed.split('|').filter(c => c.trim()).map(c => c.trim());
            if (isHeader) {
                result.push('|_. ' + cells.join(' |_. ') + ' |');
            } else {
                result.push('| ' + cells.join(' | ') + ' |');
            }
        } else {
            isHeader = true;
            result.push(line);
        }
    }
    return result.join('\n');
}

// ── Org Mode ─────────────────────────────────────────────────────────
export function markdownToOrgMode(md: string): string {
    const { cleaned, footnoteMap } = collectFootnotes(md);
    let out = cleaned;
    // Highlight: ==text== → *text* (Org has no native highlight, use bold)
    out = out.replace(/==([^=]+)==/g, '*$1*');
    // Footnote refs → Org Mode native [fn:label]
    for (const [label] of Object.entries(footnoteMap)) {
        const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`\\[\\^${escaped}\\]`, 'g'), `[fn:${label}]`);
    }
    // Headers: # → *, ## → **, etc.
    out = out.replace(/^######\s+(.+)$/gm, '****** $1');
    out = out.replace(/^#####\s+(.+)$/gm, '***** $1');
    out = out.replace(/^####\s+(.+)$/gm, '**** $1');
    out = out.replace(/^###\s+(.+)$/gm, '*** $1');
    out = out.replace(/^##\s+(.+)$/gm, '** $1');
    out = out.replace(/^#\s+(.+)$/gm, '* $1');
    // Bold: **text** → *text*
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    // Italic: *text* → /text/
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '/$1/');
    // Strikethrough: ~~text~~ → +text+
    out = out.replace(/~~([^~]+)~~/g, '+$1+');
    // Inline code: `code` → ~code~
    out = out.replace(/`([^`]+)`/g, '~$1~');
    // Code blocks: ```lang → #+BEGIN_SRC lang
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `#+BEGIN_SRC ${lang}\n${code.trimEnd()}\n#+END_SRC` :
            `#+BEGIN_SRC\n${code.trimEnd()}\n#+END_SRC`
    );
    // Links: [text](url) → [[url][text]]
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[[$2][$1]]');
    // Images: ![alt](url) → [[url]]
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '[[$2]]');
    // Blockquote
    out = out.replace(/^>\s+(.+)$/gm, '#+BEGIN_QUOTE\n$1\n#+END_QUOTE');
    // Task lists: - [ ] → - [ ], - [x] → - [X]
    out = out.replace(/^- \[x\]/gm, '- [X]');
    // Unordered list: - → -  (Org uses -)
    // Ordered list stays similar
    // Tables
    out = convertMarkdownTableToOrg(out);
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '-----');
    // Append Org Mode footnote definitions
    if (Object.keys(footnoteMap).length > 0) {
        out += '\n\n';
        for (const [label, text] of Object.entries(footnoteMap)) {
            out += `[fn:${label}] ${text}\n`;
        }
    }
    return out;
}

function convertMarkdownTableToOrg(text: string): string {
    const lines = text.split('\n');
    const result: string[] = [];
    for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(trimmed)) {
                // Convert separator to org separator
                const cells = trimmed.split('|').filter(c => c.trim());
                result.push('|' + cells.map(c => '-'.repeat(c.trim().length + 2)).join('+') + '|');
                continue;
            }
            result.push(trimmed);
        } else {
            result.push(line);
        }
    }
    return result.join('\n');
}
