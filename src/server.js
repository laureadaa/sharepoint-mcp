// Servidor MCP remoto (Streamable HTTP, sin estado) protegido con un token.
// Compatible con conectores personalizados de claude.ai.

import express from "express";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { registerTools } from "./tools.js";
import { forceRefresh } from "./auth.js";
import { MCP_AUTH_TOKEN, PORT, AUTH_MODE, validateStartup } from "./config.js";

const startupErrors = validateStartup();
if (startupErrors.length) {
  console.error("[config] Problemas de configuración:");
  for (const e of startupErrors) console.error("  - " + e);
  console.error("Revisa las variables de entorno (.env). Abortando.");
  process.exit(1);
}

function buildServer() {
  const server = new McpServer({ name: "sharepoint-onedrive-mcp", version: "1.0.0" });
  registerTools(server);
  return server;
}

// Comprueba el token: por header Authorization: Bearer <X> o por ?key=<X>.
function authorized(req) {
  if (!MCP_AUTH_TOKEN) return false;
  const header = req.headers["authorization"] || "";
  const bearer = header.startsWith("Bearer ") ? header.slice(7) : null;
  const key = req.query.key;
  return bearer === MCP_AUTH_TOKEN || key === MCP_AUTH_TOKEN;
}

const app = express();
app.use(express.json({ limit: "12mb" }));

// Renovación periódica del refresh token (sólo modo delegado) para que no caduque por inactividad.
const REFRESH_EVERY_MS = 12 * 60 * 60 * 1000;
// Sólo estado ok/fecha: /health es público, el detalle del error va al log.
const tokenStatus = { ok: null, lastRefresh: null };

async function keepTokenAlive() {
  try {
    await forceRefresh();
    tokenStatus.ok = true;
  } catch (e) {
    tokenStatus.ok = false;
    console.error("[auth] Fallo renovando el token:", e.message);
  }
  tokenStatus.lastRefresh = new Date().toISOString();
}

app.get("/health", (_req, res) => res.json({ status: "ok", mode: AUTH_MODE, token: tokenStatus }));

async function handleMcp(req, res) {
  if (!authorized(req)) {
    return res.status(401).json({
      jsonrpc: "2.0",
      error: { code: -32001, message: "No autorizado: token ausente o incorrecto." },
      id: null,
    });
  }
  try {
    const server = buildServer();
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    res.on("close", () => {
      transport.close();
      server.close();
    });
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error("[mcp] Error manejando la petición:", err);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: { code: -32603, message: "Error interno del servidor." },
        id: null,
      });
    }
  }
}

app.post("/mcp", handleMcp);

// En modo sin estado, GET/DELETE de sesión no aplican.
const methodNotAllowed = (_req, res) =>
  res.status(405).json({
    jsonrpc: "2.0",
    error: { code: -32000, message: "Método no permitido (servidor sin estado)." },
    id: null,
  });
app.get("/mcp", methodNotAllowed);
app.delete("/mcp", methodNotAllowed);

app.listen(PORT, () => {
  console.log(`[sharepoint-mcp] Escuchando en :${PORT} (modo=${AUTH_MODE})`);
  console.log(`[sharepoint-mcp] Endpoint MCP: POST /mcp  ·  salud: GET /health`);
  // Pista para no loguear el token entero.
  console.log(`[sharepoint-mcp] Token de acceso configurado (longitud ${MCP_AUTH_TOKEN.length}).`);
  void randomUUID; // reservado por si se activa modo con estado
  if (AUTH_MODE === "delegated") {
    void keepTokenAlive();
    setInterval(keepTokenAlive, REFRESH_EVERY_MS);
  }
});
