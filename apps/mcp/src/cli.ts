#!/usr/bin/env node
// mroffice-mcp --http   the MCP endpoint at <public url>/mcp, with the editor pages
// mroffice-mcp --stdio  MCP over stdin/stdout (a host that starts the server
//                       itself), with the editor pages on the HTTP side
//
// Settings come from the environment (src/config.ts): MROFFICE_MCP_PORT,
// MROFFICE_MCP_PUBLIC_URL, MROFFICE_MCP_EDITOR_DIST, MROFFICE_MCP_EDITOR_ORIGIN,
// MROFFICE_MCP_ENGINE (none | local-dev), MROFFICE_MCP_DEV_EMAIL, MROFFICE_MCP_DATA_DIR ...

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { Account } from "./auth.ts";
import { loadConfig } from "./config.ts";
import { createEngine } from "./engine.ts";
import { startHttp } from "./http.ts";
import { HeadlessEditor } from "./render.ts";
import { createServer, VERSION, type Services } from "./server.ts";
import { DocumentStore } from "./store.ts";

if (process.argv.includes("--version") || process.argv.includes("-v")) {
  process.stdout.write(`${VERSION}\n`);
  process.exit(0);
}
if (process.argv.includes("--help") || process.argv.includes("-h")) {
  process.stdout.write(`mroffice-mcp ${VERSION}: MR Office's editors as an MCP server

  mroffice-mcp --stdio   MCP over stdin/stdout, for an AI app that starts the server itself
  mroffice-mcp --http    MCP over Streamable HTTP at http://127.0.0.1:<port>/mcp
  mroffice-mcp login     sign in to your MR Office account (for Word, Excel and PowerPoint files)
  mroffice-mcp logout    sign out and revoke this computer's access
  mroffice-mcp whoami    who is signed in

Both also serve the editor page that links open, on http://127.0.0.1:<port> (default 3410).

Environment:
  MROFFICE_MCP_PORT        port of the HTTP side (3410)
  MROFFICE_MCP_DATA_DIR    where documents are kept (a folder in the system temp directory)
  MROFFICE_MCP_EDITOR_ORIGIN  where the editor's scripts load from (https://demo.mroffice.ai)
  MROFFICE_MCP_PUBLIC_URL  the address people and apps reach the HTTP side at
  MROFFICE_MCP_MAX_FILE_BYTES  largest file a tool takes (20 MB)
  MROFFICE_MCP_API_URL     the MR Office to sign in to (https://mroffice.ai)

https://mroffice.ai
`);
  process.exit(0);
}

const config = loadConfig();
const account = config.engine.kind === "account" ? new Account(config.engine.apiUrl) : undefined;
const command = process.argv.slice(2).find((argument) => !argument.startsWith("-"));
if (command === "login" || command === "logout" || command === "whoami") {
  if (!account) {
    process.stderr.write("Signing in needs the account engine (MROFFICE_MCP_ENGINE=account, the default).\n");
    process.exit(2);
  }
  if (command === "whoami") {
    const current = account.credentials();
    process.stdout.write(current ? `Signed in to ${account.apiUrl}${current.account ? ` as ${current.account.email}` : ""}\n` : `Not signed in to ${account.apiUrl}\n`);
    process.exit(current ? 0 : 1);
  }
  if (command === "logout") {
    process.stdout.write((await account.logout()) ? "Signed out; this computer's access is revoked.\n" : "Not signed in.\n");
    process.exit(0);
  }
  const { url, done } = await account.startLogin(true);
  process.stdout.write(`Opening MR Office's sign-in page. If the browser does not open, visit:\n${url}\n`);
  try {
    const signed = await done;
    process.stdout.write(`Signed in to ${account.apiUrl}${signed.account ? ` as ${signed.account.email}` : ""}.\n`);
    process.exit(0);
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exit(1);
  }
}
if (command) {
  process.stderr.write(`Unknown command ${command}; see mroffice-mcp --help\n`);
  process.exit(2);
}

const services: Services = { config, store: new DocumentStore(config.dataDir), engine: createEngine(config, account), renderer: new HeadlessEditor(config), account: config.auth === "oauth" ? undefined : account };
const log = (line: string) => process.stderr.write(`[mroffice-mcp] ${line}\n`);

if (process.argv.includes("--stdio")) {
  // the editor pages and links; another process over the same data directory may already serve them
  try {
    await startHttp(services);
    log(`editor pages at ${config.publicUrl}`);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EADDRINUSE") throw error;
    log(`port ${config.port} is in use: links point to the server already there`);
  }
  await createServer(services).connect(new StdioServerTransport());
  log("MCP on stdio");
} else {
  await startHttp(services);
  log(`MCP at ${config.publicUrl}/mcp; editor pages at ${config.publicUrl}/_mcp/view/<id>; engine ${config.engine.kind}; editor assets from ${config.editorOrigin}; documents in ${config.dataDir}`);
}

const stop = async () => {
  await services.renderer.close();
  process.exit(0);
};
process.on("SIGINT", stop);
process.on("SIGTERM", stop);
