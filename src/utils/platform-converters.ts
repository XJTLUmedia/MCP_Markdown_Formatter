// Platform-specific Markdown converters

// ── Slack mrkdwn ─────────────────────────────────────────────────────
export function markdownToSlack(md: string): string {
    let out = md;
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    out = out.replace(/~~([^~]+)~~/g, '~$1~');
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<$2|$1>');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<$2|$1>');
    out = out.replace(/^#{1,6}\s+(.+)$/gm, '*$1*');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '───────────────');
    return out;
}

// ── Discord markdown ─────────────────────────────────────────────────
export function markdownToDiscord(md: string): string {
    let out = md;
    out = out.replace(/^# (.+)$/gm, '**__$1__**');
    out = out.replace(/^## (.+)$/gm, '**$1**');
    out = out.replace(/^### (.+)$/gm, '__$1__');
    out = out.replace(/^#{4,6}\s+(.+)$/gm, '*$1*');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '');
    return out;
}

// ── JIRA wiki markup ─────────────────────────────────────────────────
export function markdownToJira(md: string): string {
    let out = md;
    out = out.replace(/^######\s+(.+)$/gm, 'h6. $1');
    out = out.replace(/^#####\s+(.+)$/gm, 'h5. $1');
    out = out.replace(/^####\s+(.+)$/gm, 'h4. $1');
    out = out.replace(/^###\s+(.+)$/gm, 'h3. $1');
    out = out.replace(/^##\s+(.+)$/gm, 'h2. $1');
    out = out.replace(/^#\s+(.+)$/gm, 'h1. $1');
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    out = out.replace(/~~([^~]+)~~/g, '-$1-');
    out = out.replace(/`([^`]+)`/g, '{{$1}}');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `{code:${lang}}\n${code.trimEnd()}\n{code}` : `{code}\n${code.trimEnd()}\n{code}`
    );
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[$1|$2]');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '!$2!');
    out = out.replace(/^(\s*)[-+]\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' ';
    });
    out = out.replace(/^(\s*)\d+\.\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '#'.repeat(level) + ' ';
    });
    out = out.replace(/^>\s+(.+)$/gm, '{quote}$1{quote}');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----');
    const lines = out.split('\n');
    const result: string[] = [];
    let headerDone = false;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(line.trim())) { headerDone = true; continue; }
            if (!headerDone) { result.push(line.replace(/\|/g, '||')); } else { result.push(line); }
        } else { headerDone = false; result.push(line); }
    }
    return result.join('\n');
}

// ── Confluence wiki markup ───────────────────────────────────────────
export function markdownToConfluence(md: string): string {
    let out = md;
    out = out.replace(/^######\s+(.+)$/gm, 'h6. $1');
    out = out.replace(/^#####\s+(.+)$/gm, 'h5. $1');
    out = out.replace(/^####\s+(.+)$/gm, 'h4. $1');
    out = out.replace(/^###\s+(.+)$/gm, 'h3. $1');
    out = out.replace(/^##\s+(.+)$/gm, 'h2. $1');
    out = out.replace(/^#\s+(.+)$/gm, 'h1. $1');
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    out = out.replace(/~~([^~]+)~~/g, '-$1-');
    out = out.replace(/`([^`]+)`/g, '{{$1}}');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `{code:language=${lang}}\n${code.trimEnd()}\n{code}` : `{code}\n${code.trimEnd()}\n{code}`
    );
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[$1|$2]');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '!$2!');
    out = out.replace(/^(\s*)[-+]\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' ';
    });
    out = out.replace(/^(\s*)\d+\.\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '#'.repeat(level) + ' ';
    });
    out = out.replace(/^>\s+(.+)$/gm, '{quote}$1{quote}');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----');
    out = out.replace(/<!--\s*note:\s*([\s\S]*?)-->/gi, '{note}$1{note}');
    out = out.replace(/<!--\s*info:\s*([\s\S]*?)-->/gi, '{info}$1{info}');
    return out;
}

// ── AsciiDoc ─────────────────────────────────────────────────────────
export function markdownToAsciiDoc(md: string): string {
    let out = md;
    out = out.replace(/^######\s+(.+)$/gm, '====== $1');
    out = out.replace(/^#####\s+(.+)$/gm, '===== $1');
    out = out.replace(/^####\s+(.+)$/gm, '==== $1');
    out = out.replace(/^###\s+(.+)$/gm, '=== $1');
    out = out.replace(/^##\s+(.+)$/gm, '== $1');
    out = out.replace(/^#\s+(.+)$/gm, '= $1');
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    out = out.replace(/`([^`]+)`/g, '`$1`');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) => {
        const attr = lang ? `[source,${lang}]\n` : '';
        return `${attr}----\n${code.trimEnd()}\n----`;
    });
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$2[$1]');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, 'image::$2[$1]');
    const lines = out.split('\n');
    const result: string[] = [];
    let inQuote = false;
    for (const line of lines) {
        if (line.startsWith('> ')) {
            if (!inQuote) { result.push('[quote]'); result.push('____'); inQuote = true; }
            result.push(line.slice(2));
        } else {
            if (inQuote) { result.push('____'); inQuote = false; }
            result.push(line);
        }
    }
    if (inQuote) result.push('____');
    out = result.join('\n');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, "'''");
    out = out.replace(/^(\s*)[-+]\s+/gm, '* ');
    out = out.replace(/^\d+\.\s+/gm, '. ');
    return out;
}

// ── reStructuredText ─────────────────────────────────────────────────
export function markdownToRST(md: string): string {
    let out = md;
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) => {
        const directive = lang ? `.. code-block:: ${lang}` : '.. code-block::';
        const indented = code.split('\n').map((l: string) => '   ' + l).join('\n').trimEnd();
        return `${directive}\n\n${indented}`;
    });
    out = out.replace(/(?<!`)(`[^`]+`)(?!`)/g, '`$1`');
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '`$1 <$2>`_');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '.. image:: $2\n   :alt: $1');
    const lines = out.split('\n');
    const result: string[] = [];
    for (const line of lines) {
        const h1 = line.match(/^#\s+(.+)$/);
        const h2 = line.match(/^##\s+(.+)$/);
        const h3 = line.match(/^###\s+(.+)$/);
        const h4 = line.match(/^####\s+(.+)$/);
        if (h1) { result.push(h1[1], '='.repeat(h1[1].length)); }
        else if (h2) { result.push(h2[1], '-'.repeat(h2[1].length)); }
        else if (h3) { result.push(h3[1], '~'.repeat(h3[1].length)); }
        else if (h4) { result.push(h4[1], '^'.repeat(h4[1].length)); }
        else { result.push(line); }
    }
    out = result.join('\n');
    out = out.replace(/^>\s+(.+)$/gm, '   $1');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----------');
    out = out.replace(/^(\d+)\.\s+/gm, '#. ');
    return out;
}

// ── MediaWiki markup ─────────────────────────────────────────────────
export function markdownToMediaWiki(md: string): string {
    let out = md;
    out = out.replace(/^######\s+(.+)$/gm, '======= $1 =======');
    out = out.replace(/^#####\s+(.+)$/gm, '====== $1 ======');
    out = out.replace(/^####\s+(.+)$/gm, '===== $1 =====');
    out = out.replace(/^###\s+(.+)$/gm, '==== $1 ====');
    out = out.replace(/^##\s+(.+)$/gm, '=== $1 ===');
    out = out.replace(/^#\s+(.+)$/gm, '== $1 ==');
    out = out.replace(/\*\*([^*]+)\*\*/g, "'''$1'''");
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, "''$1''");
    out = out.replace(/`([^`]+)`/g, '<code>$1</code>');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `<syntaxhighlight lang="${lang}">\n${code.trimEnd()}\n</syntaxhighlight>` :
            `<pre>\n${code.trimEnd()}\n</pre>`
    );
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[$2 $1]');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '[[File:$2|$1]]');
    out = out.replace(/^(\s*)[-+]\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '*'.repeat(level) + ' ';
    });
    out = out.replace(/^(\s*)\d+\.\s+/gm, (_m, indent) => {
        const level = Math.floor(indent.length / 2) + 1;
        return '#'.repeat(level) + ' ';
    });
    out = out.replace(/^>\s+(.+)$/gm, '<blockquote>$1</blockquote>');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '----');
    const lines = out.split('\n');
    const result: string[] = [];
    let inTable = false;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line.startsWith('|') && line.endsWith('|')) {
            if (/^\|[\s-:]+\|/.test(line)) continue;
            const cells = line.split('|').filter(c => c.trim()).map(c => c.trim());
            if (!inTable) {
                result.push('{| class="wikitable"');
                inTable = true;
                result.push('|-');
                cells.forEach(c => result.push('! ' + c));
            } else {
                result.push('|-');
                cells.forEach(c => result.push('| ' + c));
            }
        } else {
            if (inTable) { result.push('|}'); inTable = false; }
            result.push(lines[i]);
        }
    }
    if (inTable) result.push('|}');
    return result.join('\n');
}

// ── BBCode ───────────────────────────────────────────────────────────
export function markdownToBBCode(md: string): string {
    let out = md;
    out = out.replace(/^#{1,6}\s+(.+)$/gm, '[b][size=5]$1[/size][/b]');
    out = out.replace(/\*\*([^*]+)\*\*/g, '[b]$1[/b]');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '[i]$1[/i]');
    out = out.replace(/~~([^~]+)~~/g, '[s]$1[/s]');
    out = out.replace(/`([^`]+)`/g, '[code]$1[/code]');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, _lang, code) =>
        `[code]\n${code.trimEnd()}\n[/code]`
    );
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[url=$2]$1[/url]');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '[img]$2[/img]');
    out = out.replace(/^>\s+(.+)$/gm, '[quote]$1[/quote]');
    out = out.replace(/^[-*+]\s+(.+)$/gm, '[*]$1');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '[hr]');
    return out;
}

// ── Textile ──────────────────────────────────────────────────────────
export function markdownToTextile(md: string): string {
    let out = md;
    out = out.replace(/^######\s+(.+)$/gm, 'h6. $1');
    out = out.replace(/^#####\s+(.+)$/gm, 'h5. $1');
    out = out.replace(/^####\s+(.+)$/gm, 'h4. $1');
    out = out.replace(/^###\s+(.+)$/gm, 'h3. $1');
    out = out.replace(/^##\s+(.+)$/gm, 'h2. $1');
    out = out.replace(/^#\s+(.+)$/gm, 'h1. $1');
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '_$1_');
    out = out.replace(/~~([^~]+)~~/g, '-$1-');
    out = out.replace(/`([^`]+)`/g, '@$1@');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, _lang, code) =>
        `bc. ${code.trimEnd()}`
    );
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '"$1":$2');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '!$2($1)!');
    out = out.replace(/^>\s+(.+)$/gm, 'bq. $1');
    out = out.replace(/^[-+]\s+/gm, '* ');
    out = out.replace(/^\d+\.\s+/gm, '# ');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '---');
    return out;
}

// ── Org Mode ─────────────────────────────────────────────────────────
export function markdownToOrgMode(md: string): string {
    let out = md;
    out = out.replace(/^######\s+(.+)$/gm, '****** $1');
    out = out.replace(/^#####\s+(.+)$/gm, '***** $1');
    out = out.replace(/^####\s+(.+)$/gm, '**** $1');
    out = out.replace(/^###\s+(.+)$/gm, '*** $1');
    out = out.replace(/^##\s+(.+)$/gm, '** $1');
    out = out.replace(/^#\s+(.+)$/gm, '* $1');
    out = out.replace(/\*\*([^*]+)\*\*/g, '*$1*');
    out = out.replace(/(?<!\*)\*(?!\*)([^*]+)(?<!\*)\*(?!\*)/g, '/$1/');
    out = out.replace(/~~([^~]+)~~/g, '+$1+');
    out = out.replace(/`([^`]+)`/g, '~$1~');
    out = out.replace(/```(\w+)?\n([\s\S]*?)```/g, (_m, lang, code) =>
        lang ? `#+BEGIN_SRC ${lang}\n${code.trimEnd()}\n#+END_SRC` :
            `#+BEGIN_SRC\n${code.trimEnd()}\n#+END_SRC`
    );
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '[[$2][$1]]');
    out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '[[$2]]');
    out = out.replace(/^>\s+(.+)$/gm, '#+BEGIN_QUOTE\n$1\n#+END_QUOTE');
    out = out.replace(/^(-{3,}|\*{3,}|_{3,})$/gm, '-----');
    out = out.replace(/^- \[x\]/gm, '- [X]');
    return out;
}
