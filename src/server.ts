import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerPresentationTools } from "./tools/presentation.js";

export function createServer(): McpServer {
  const server = new McpServer({
    name: "apple-powerpoint-mcp",
    version: "0.1.0",
  });

  registerPresentationTools(server);

  return server;
}
