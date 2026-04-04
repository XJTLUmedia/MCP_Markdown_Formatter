import { stripMarkdown } from "./core-exports.js";

// ── Slack mrkdwn ─────────────────────────────────────────────────────
export function markdownToSlack(md: string): string {
    let out = md;
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
    // Blockquotes: > text → > text (Slack supports >)
    // Ordered lists: 1. text stays the same
    // Unordered lists: - text stays the same
    // Horizontal rule
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '───────────────');
    return out;
}

// ── Discord markdown ─────────────────────────────────────────────────
export function markdownToDiscord(md: string): string {
    let out = md;
    // Discord supports most standard markdown, just a few tweaks:
    // Headers: # text → **__text__** (Discord renders # only in certain contexts)
    out = out.replace(/^# (.+)$/gm, '**__$1__**');
    out = out.replace(/^## (.+)$/gm, '**$1**');
    out = out.replace(/^### (.+)$/gm, '__$1__');
    out = out.replace(/^#{4,6}\s+(.+)$/gm, '*$1*');
    // Spoiler: no markdown equivalent, keep as-is
    // Block quotes: > works in Discord
    // Code blocks: ``` works in Discord
    // Horizontal rules are not rendered in Discord
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '');
    return out;
}

// ── JIRA wiki markup ─────────────────────────────────────────────────
export function markdownToJira(md: string): string {
    let out = md;
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
    return result.join('\n');
}

// ── Confluence wiki markup ───────────────────────────────────────────
export function markdownToConfluence(md: string): string {
    let out = md;
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
    // Info/note panels from HTML comments
    out = out.replace(/<!--\s*note:\s*([\s\S]*?)-->/gi, '{note}$1{note}');
    out = out.replace(/<!--\s*info:\s*([\s\S]*?)-->/gi, '{info}$1{info}');
    return out;
}

// ── AsciiDoc ─────────────────────────────────────────────────────────
export function markdownToAsciiDoc(md: string): string {
    let out = md;
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
    // Unordered list marker stays as *
    out = out.replace(/^(\s*)[-+]\s+/gm, '* ');
    // Ordered list: 1. → .
    out = out.replace(/^\d+\.\s+/gm, '. ');
    return out;
}

// ── reStructuredText ─────────────────────────────────────────────────
export function markdownToRST(md: string): string {
    let out = md;
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
    return out;
}

// ── MediaWiki markup ─────────────────────────────────────────────────
export function markdownToMediaWiki(md: string): string {
    let out = md;
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
    return result.join('\n');
}

// ── BBCode ───────────────────────────────────────────────────────────
export function markdownToBBCode(md: string): string {
    let out = md;
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
    // Unordered list
    out = out.replace(/^[-*+]\s+(.+)$/gm, '[*]$1');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '[hr]');
    return out;
}

// ── Textile ──────────────────────────────────────────────────────────
export function markdownToTextile(md: string): string {
    let out = md;
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
    // Unordered list
    out = out.replace(/^[-+]\s+/gm, '* ');
    // Ordered list
    out = out.replace(/^\d+\.\s+/gm, '# ');
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '---');
    return out;
}

// ── Org Mode ─────────────────────────────────────────────────────────
export function markdownToOrgMode(md: string): string {
    let out = md;
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
    // Unordered list: - → -  (Org uses -)
    // Ordered list stays similar
    // Horizontal rules
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '-----');
    // Task lists: - [ ] → - [ ], - [x] → - [X]
    out = out.replace(/^- \[x\]/gm, '- [X]');
    return out;
}
