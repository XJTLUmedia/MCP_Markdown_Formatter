import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

async function main() {
    const transport = new StdioClientTransport({
        command: "node",
        args: ["dist/mcp-server/src/index.js"],
    });

    const client = new Client(
        {
            name: "test-client",
            version: "1.0.0",
        },
        {
            capabilities: {},
        }
    );

    console.log("Connecting to server...");
    await client.connect(transport);
    console.log("Connected!\n");

    console.log("=".repeat(60));
    console.log("Listing all available tools...");
    console.log("=".repeat(60));
    const tools = await client.listTools();
    console.log(`Found ${tools.tools.length} tools:\n`);
    tools.tools.forEach((t, i) => {
        console.log(`  ${i + 1}. ${t.name}`);
        console.log(`     Description: ${t.description}`);
    });

    const sampleMarkdown = `# Hello World

This is a **bold** and *italic* test document.

## Features

- Item 1
- Item 2
- Item 3

### Code Block

\`\`\`javascript
console.log("Hello, World!");
\`\`\`

### Table

| Name | Age | City |
|------|-----|------|
| Alice | 30 | NYC |
| Bob | 25 | LA |

### Math

The equation is $E = mc^2$.
`;

    console.log("\n" + "=".repeat(60));
    console.log("Testing individual tools...");
    console.log("=".repeat(60));

    // 1. harmonize_markdown
    console.log("\n1. Testing 'harmonize_markdown'...");
    const harmonizeResult = await client.callTool({
        name: "harmonize_markdown",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("   ✓ Result preview:", ((harmonizeResult.content as any)[0].text as string).substring(0, 80) + "...");

    // 2. convert_to_txt
    console.log("\n2. Testing 'convert_to_txt'...");
    const txtResult = await client.callTool({
        name: "convert_to_txt",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("   ✓ Result preview:", ((txtResult.content as any)[0].text as string).substring(0, 80) + "...");

    // 3. convert_to_rtf
    console.log("\n3. Testing 'convert_to_rtf'...");
    const rtfResult = await client.callTool({
        name: "convert_to_rtf",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("   ✓ RTF starts with:", ((rtfResult.content as any)[0].text as string).substring(0, 50));

    // 4. convert_to_latex
    console.log("\n4. Testing 'convert_to_latex'...");
    const latexResult = await client.callTool({
        name: "convert_to_latex",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("   ✓ LaTeX preview:", ((latexResult.content as any)[0].text as string).substring(0, 80) + "...");

    // 5. convert_to_csv
    console.log("\n5. Testing 'convert_to_csv'...");
    const csvResult = await client.callTool({
        name: "convert_to_csv",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("   ✓ CSV Result:\n", ((csvResult.content as any)[0].text as string));

    // 6. convert_to_json
    console.log("\n6. Testing 'convert_to_json'...");
    const jsonResult = await client.callTool({
        name: "convert_to_json",
        arguments: { markdown: sampleMarkdown, title: "Test Document" }
    });
    const jsonPreview = JSON.parse((jsonResult.content as any)[0].text);
    console.log("   ✓ JSON keys:", Object.keys(jsonPreview));

    // 7. convert_to_xml
    console.log("\n7. Testing 'convert_to_xml'...");
    const xmlResult = await client.callTool({
        name: "convert_to_xml",
        arguments: { markdown: sampleMarkdown, title: "Test Document" }
    });
    console.log("   ✓ XML starts with:", ((xmlResult.content as any)[0].text as string).substring(0, 60));

    // 8. convert_to_html
    console.log("\n8. Testing 'convert_to_html'...");
    const htmlResult = await client.callTool({
        name: "convert_to_html",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("   ✓ HTML starts with:", ((htmlResult.content as any)[0].text as string).substring(0, 60));

    // 9. convert_to_md
    console.log("\n9. Testing 'convert_to_md'...");
    const mdResult = await client.callTool({
        name: "convert_to_md",
        arguments: { markdown: sampleMarkdown, harmonize: true }
    });
    console.log("   ✓ MD preview:", ((mdResult.content as any)[0].text as string).substring(0, 80) + "...");

    // 10. generate_html
    console.log("\n10. Testing 'generate_html'...");
    const genHtmlResult = await client.callTool({
        name: "generate_html",
        arguments: { markdown: sampleMarkdown, title: "Generated Document" }
    });
    console.log("   ✓ Full HTML starts with:", ((genHtmlResult.content as any)[0].text as string).substring(0, 60));

    // 11. convert_to_docx (binary - no output path)
    console.log("\n11. Testing 'convert_to_docx' (without output_path)...");
    const docxResult = await client.callTool({
        name: "convert_to_docx",
        arguments: { markdown: sampleMarkdown }
    });
    const docxInfo = JSON.parse((docxResult.content as any)[0].text);
    console.log("   ✓ DOCX info:", { format: docxInfo.format, size: docxInfo.file_size_bytes, hint: docxInfo.hint?.substring(0, 50) + "..." });

    // 12. convert_to_xlsx (binary - no output path)
    console.log("\n12. Testing 'convert_to_xlsx' (without output_path)...");
    const xlsxResult = await client.callTool({
        name: "convert_to_xlsx",
        arguments: { markdown: sampleMarkdown }
    });
    const xlsxInfo = JSON.parse((xlsxResult.content as any)[0].text);
    console.log("   ✓ XLSX info:", { format: xlsxInfo.format, size: xlsxInfo.file_size_bytes });

    // 13. Test saving DOCX to file
    console.log("\n13. Testing 'convert_to_docx' (WITH output_path)...");
    const docxSaveResult = await client.callTool({
        name: "convert_to_docx",
        arguments: { markdown: sampleMarkdown, output_path: "./test-output.docx" }
    });
    const docxSaveInfo = JSON.parse((docxSaveResult.content as any)[0].text);
    console.log("   ✓ Saved:", docxSaveInfo);

    console.log("\n" + "=".repeat(60));
    console.log("All tests completed successfully!");
    console.log("=".repeat(60));

    await client.close();
}

main().catch(console.error);
