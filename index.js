#!/usr/bin/env node

const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { CallToolRequestSchema, ListToolsRequestSchema } = require("@modelcontextprotocol/sdk/types.js");
const { generateDocument } = require('./scripts/generate-application.js');
const path = require('path');

const server = new Server(
  {
    name: "dream-big-application",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "generate_application",
        description: "生成 Dream Big 元大公益圓夢計畫申請書 (Word/Docx)。Generate the Dream Big application document.",
        inputSchema: {
          type: "object",
          properties: {
            data: {
              type: "object",
              description: "申請書資料物件 (JSON)。結構請參考 references/fields-guide.md。",
            },
            output_path: {
              type: "string",
              description: "生成檔案的絕對路徑 (Absolute path for the output .docx file).",
            },
          },
          required: ["data", "output_path"],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  if (request.params.name === "generate_application") {
    const { data, output_path } = request.params.arguments;
    
    try {
      console.error(`Generating application to ${output_path}...`);
      await generateDocument(data, output_path);
      return {
        content: [
          {
            type: "text",
            text: `Successfully generated application document at: ${output_path}`,
          },
        ],
      };
    } catch (error) {
      console.error(`Error: ${error.message}`);
      return {
        content: [
          {
            type: "text",
            text: `Error generating document: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
  throw new Error("Tool not found");
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Dream Big Application MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});
