// ── Markdown Repair / Linting ────────────────────────────────────────
// Fixes common broken markdown from LLM output

export interface LintIssue {
    line: number;
    column: number;
    severity: 'error' | 'warning' | 'info';
    rule: string;
    message: string;
    fixable: boolean;
}

export function repairMarkdown(md: string): string {
    let out = md;
    out = repairCodeFences(out);
    out = repairBrokenTables(out);
    out = repairStrayMarkers(out);
    out = repairHeadings(out);
    out = repairListIndentation(out);
    out = repairBrokenLinks(out);
    out = repairBrokenTaskLists(out);
    out = normalizeWhitespace(out);
    return out;
}

// Fix unclosed or mismatched code fences
export function repairCodeFences(md: string): string {
    const lines = md.split('\n');
    const result: string[] = [];
    let inFence = false;
    let fenceChar = '`';
    let fenceCount = 0;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const backtickMatch = line.match(/^(`{3,})([\w]*)\s*$/);
        const tildeMatch = line.match(/^(~{3,})([\w]*)\s*$/);
        const match = backtickMatch || tildeMatch;

        if (match) {
            const char = match[1][0];
            const count = match[1].length;
            if (!inFence) {
                inFence = true;
                fenceChar = char;
                fenceCount = count;
                result.push(line);
            } else if (char === fenceChar && count >= fenceCount) {
                inFence = false;
                result.push(fenceChar.repeat(fenceCount));
            } else {
                result.push(line);
            }
        } else {
            result.push(line);
        }
    }

    // Close unclosed fence at end
    if (inFence) {
        result.push(fenceChar.repeat(fenceCount));
    }

    return result.join('\n');
}

// Fix broken tables (mismatched columns, missing separators)
export function repairBrokenTables(md: string): string {
    const lines = md.split('\n');
    const result: string[] = [];
    let i = 0;

    while (i < lines.length) {
        const line = lines[i];
        // Detect table start: line with pipes
        if (line.trim().includes('|') && line.trim().startsWith('|')) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].trim().includes('|') && lines[i].trim().startsWith('|')) {
                tableLines.push(lines[i]);
                i++;
            }

            if (tableLines.length >= 1) {
                const repaired = repairTableBlock(tableLines);
                result.push(...repaired);
            }
            continue;
        }
        result.push(line);
        i++;
    }

    return result.join('\n');
}

function repairTableBlock(lines: string[]): string[] {
    // Parse cells and find max column count
    const parsed = lines
        .filter(l => !/^\|[\s-:]+\|[\s-:|]*$/.test(l.trim())) // remove existing separators
        .map(l => l.split('|').slice(1, -1).map(c => c.trim()));

    if (parsed.length === 0) return lines;

    const maxCols = Math.max(...parsed.map(r => r.length));

    // Pad rows to have consistent column count
    const padded = parsed.map(row => {
        while (row.length < maxCols) row.push('');
        return row;
    });

    // Reconstruct table
    const result: string[] = [];
    result.push('| ' + padded[0].join(' | ') + ' |');
    result.push('| ' + padded[0].map(() => '---').join(' | ') + ' |');
    for (let i = 1; i < padded.length; i++) {
        result.push('| ' + padded[i].join(' | ') + ' |');
    }
    return result;
}

// Fix stray emphasis markers (* or _ at start/end without matching)
export function repairStrayMarkers(md: string): string {
    let out = md;
    // Fix unmatched bold markers at line level
    out = out.replace(/^(\*\*[^*]+)$/gm, (m) => {
        if (!m.endsWith('**')) return m + '**';
        return m;
    });
    out = out.replace(/^([^*]+\*\*)$/gm, (m) => {
        if (!m.startsWith('**')) return '**' + m;
        return m;
    });
    // Fix solo backticks (odd count of backticks on a line)
    out = out.replace(/^((?:[^`]*`[^`]*){1,})$/gm, (line) => {
        const count = (line.match(/`/g) || []).length;
        if (count % 2 !== 0) {
            // Find the last backtick and close it
            return line + '`';
        }
        return line;
    });
    return out;
}

// Fix heading spacing and format
export function repairHeadings(md: string): string {
    let out = md;
    // Fix missing space after #: #Heading → # Heading
    out = out.replace(/^(#{1,6})([^\s#])/gm, '$1 $2');
    // Fix trailing hashes: # Heading # → # Heading
    out = out.replace(/^(#{1,6}\s+.+?)\s*#+\s*$/gm, '$1');
    // Ensure blank line before headings (unless at start of doc)
    const lines = out.split('\n');
    const result: string[] = [];
    for (let i = 0; i < lines.length; i++) {
        if (i > 0 && /^#{1,6}\s/.test(lines[i]) && lines[i - 1].trim() !== '') {
            result.push('');
        }
        result.push(lines[i]);
    }
    return result.join('\n');
}

// Fix inconsistent list indentation
export function repairListIndentation(md: string): string {
    const lines = md.split('\n');
    const result: string[] = [];

    for (const line of lines) {
        // Normalize tab-based indentation to spaces
        let fixed = line.replace(/\t/g, '    ');
        // Fix mixed list markers at same level to use consistent marker
        if (/^\s*[+]\s/.test(fixed)) {
            fixed = fixed.replace(/^(\s*)[+]\s/, '$1- ');
        }
        result.push(fixed);
    }

    return result.join('\n');
}

// Fix broken link syntax
export function repairBrokenLinks(md: string): string {
    let out = md;
    // Fix missing closing paren: [text](url → [text](url)
    out = out.replace(/\[([^\]]+)\]\(([^)\s]+)(?=\s|$)/gm, '[$1]($2)');
    // Fix missing closing bracket: [text(url) → [text](url)
    out = out.replace(/\[([^\]]+)\(([^)]+)\)/g, '[$1]($2)');
    return out;
}

// Fix broken task list syntax from LLM output
export function repairBrokenTaskLists(md: string): string {
    let out = md;
    // Fix missing space in checkbox: - [] → - [ ]
    out = out.replace(/^(\s*[-*+])\s+\[\](\s+)/gm, '$1 [ ]$2');
    // Fix uppercase X: - [X] → - [x]
    out = out.replace(/^(\s*[-*+]\s+)\[X\]/gm, '$1[x]');
    // Fix no space after checkbox: - [x]text → - [x] text
    out = out.replace(/^(\s*[-*+]\s+\[[ xX]\])([^\s])/gm, '$1 $2');
    return out;
}

// Normalize excessive whitespace
export function normalizeWhitespace(md: string): string {
    let out = md;
    // Remove trailing whitespace (except intentional line breaks: 2+ spaces)
    out = out.replace(/([^ \n]) +$/gm, '$1');
    // Collapse 3+ blank lines to 2
    out = out.replace(/\n{4,}/g, '\n\n\n');
    // Ensure file ends with single newline
    out = out.replace(/\n*$/, '\n');
    return out;
}

// ── Markdown Linting ─────────────────────────────────────────────────
export function lintMarkdown(md: string): LintIssue[] {
    const issues: LintIssue[] = [];
    const lines = md.split('\n');
    let inCodeBlock = false;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const lineNum = i + 1;

        // Track code blocks
        if (/^```/.test(line.trim())) {
            inCodeBlock = !inCodeBlock;
            continue;
        }
        if (inCodeBlock) continue;

        // Missing space after heading marker
        if (/^#{1,6}[^\s#]/.test(line)) {
            issues.push({
                line: lineNum, column: 1, severity: 'error',
                rule: 'heading-space', message: 'Missing space after heading marker (#)',
                fixable: true
            });
        }

        // Trailing whitespace
        if (/\S +$/.test(line) && !/  $/.test(line)) {
            issues.push({
                line: lineNum, column: line.length, severity: 'warning',
                rule: 'trailing-whitespace', message: 'Trailing whitespace',
                fixable: true
            });
        }

        // Inconsistent list markers
        if (/^\s*[+]\s/.test(line)) {
            issues.push({
                line: lineNum, column: 1, severity: 'warning',
                rule: 'list-marker', message: 'Inconsistent list marker (+), prefer - or *',
                fixable: true
            });
        }

        // Hard tabs
        if (line.includes('\t')) {
            issues.push({
                line: lineNum, column: line.indexOf('\t') + 1, severity: 'warning',
                rule: 'no-hard-tabs', message: 'Hard tab found, prefer spaces',
                fixable: true
            });
        }

        // Multiple blank lines
        if (i > 1 && line.trim() === '' && lines[i - 1].trim() === '' && lines[i - 2]?.trim() === '') {
            issues.push({
                line: lineNum, column: 1, severity: 'info',
                rule: 'no-multiple-blanks', message: 'Multiple consecutive blank lines',
                fixable: true
            });
        }

        // Bare URLs (not inside links or code)
        if (/(?<!\(|<)https?:\/\/[^\s)>]+/.test(line) && !/\[.*\]\(/.test(line) && !/`/.test(line)) {
            issues.push({
                line: lineNum,
                column: line.search(/https?:\/\//) + 1,
                severity: 'info',
                rule: 'bare-url', message: 'Bare URL found, consider wrapping in link syntax',
                fixable: false
            });
        }

        // Unclosed emphasis (simple heuristic)
        const boldCount = (line.match(/\*\*/g) || []).length;
        if (boldCount % 2 !== 0) {
            issues.push({
                line: lineNum, column: 1, severity: 'warning',
                rule: 'unclosed-emphasis', message: 'Possible unclosed bold (**) markers',
                fixable: true
            });
        }

        // Broken task list syntax
        if (/^\s*[-*+]\s+\[\]/.test(line)) {
            issues.push({
                line: lineNum, column: 1, severity: 'warning',
                rule: 'broken-task-list', message: 'Missing space in task list checkbox: [] should be [ ]',
                fixable: true
            });
        }
    }

    // Check for unclosed code fences
    if (inCodeBlock) {
        issues.push({
            line: lines.length, column: 1, severity: 'error',
            rule: 'unclosed-code-fence', message: 'Unclosed code fence (```) at end of document',
            fixable: true
        });
    }

    return issues;
}
