import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { SSEClientTransport } from "@modelcontextprotocol/sdk/client/sse.js";

async function main() {
    // Current server URL (Vercel or local via proxy)
    const url = new URL("http://localhost:3000/mcp");
    const transport = new SSEClientTransport(url);

    const client = new Client(
        {
            name: "test-client-http",
            version: "1.0.0",
        },
        {
            capabilities: {},
        }
    );

    console.log(`Connecting to server at ${url}...`);
    await client.connect(transport);
    console.log("Connected!\n");

    console.log("Listing all available tools...");
    const tools = await client.listTools();
    console.log(`Found ${tools.tools.length} tools:\n`);

    const sampleMarkdown = `# Hello Merged MCP
This is a test document processed via the consolidated Vercel server.
| Feature | Status |
|---------|--------|
| Merged  | Yes    |
| Vercel  | Ready  |
`;

    console.log("Testing 'convert_to_csv'...");
    const csvResult = await client.callTool({
        name: "convert_to_csv",
        arguments: { markdown: sampleMarkdown }
    });
    console.log("CSV Result:\n", ((csvResult.content as any)[0].text as string));

    await client.close();
}

main().catch(console.error);
