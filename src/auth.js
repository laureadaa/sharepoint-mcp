// Gestor de tokens de acceso a Microsoft Graph.
// Soporta dos modos:
//   - "delegated": usa un refresh token (obtenido con `npm run login`) y lo renueva solo.
//   - "app":       client credentials (app-only), requiere consentimiento de administrador.

import { readFileSync, writeFileSync } from "node:fs";
import {
  AUTH_MODE,
  CLIENT_ID,
  CLIENT_SECRET,
  SEED_REFRESH_TOKEN,
  TOKEN_FILE,
  DELEGATED_SCOPES,
  tokenEndpoint,
} from "./config.js";

let cachedAccessToken = null;
let cachedExpiresAt = 0; // epoch ms
let currentRefreshToken = null;

function loadPersistedRefreshToken() {
  try {
    const raw = readFileSync(TOKEN_FILE, "utf8");
    const data = JSON.parse(raw);
    if (data && typeof data.refresh_token === "string") return data.refresh_token;
  } catch {
    /* no existe todavía: normal */
  }
  return null;
}

function persistRefreshToken(token) {
  try {
    writeFileSync(TOKEN_FILE, JSON.stringify({ refresh_token: token }, null, 2), "utf8");
  } catch (e) {
    // En hosts de sólo-lectura o efímeros esto puede fallar; no es crítico porque
    // mantenemos el token en memoria durante la vida del proceso.
    console.warn("[auth] No se pudo persistir el refresh token:", e.message);
  }
}

async function fetchToken(body) {
  const res = await fetch(tokenEndpoint(), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(body).toString(),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    const desc = json.error_description || json.error || `HTTP ${res.status}`;
    throw new Error(`Fallo al obtener token (${AUTH_MODE}): ${desc}`);
  }
  return json;
}

async function getAppToken() {
  const json = await fetchToken({
    grant_type: "client_credentials",
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });
  cachedAccessToken = json.access_token;
  cachedExpiresAt = Date.now() + (json.expires_in - 120) * 1000;
  return cachedAccessToken;
}

async function getDelegatedToken() {
  if (!currentRefreshToken) {
    currentRefreshToken = loadPersistedRefreshToken() || SEED_REFRESH_TOKEN;
  }
  if (!currentRefreshToken) {
    throw new Error(
      "No hay refresh token disponible. Ejecuta `npm run login` y define GRAPH_REFRESH_TOKEN."
    );
  }
  const json = await fetchToken({
    grant_type: "refresh_token",
    client_id: CLIENT_ID,
    refresh_token: currentRefreshToken,
    scope: DELEGATED_SCOPES,
  });
  cachedAccessToken = json.access_token;
  cachedExpiresAt = Date.now() + (json.expires_in - 120) * 1000;
  // AAD rota el refresh token: guardamos el nuevo.
  if (json.refresh_token && json.refresh_token !== currentRefreshToken) {
    currentRefreshToken = json.refresh_token;
    persistRefreshToken(currentRefreshToken);
  }
  return cachedAccessToken;
}

// Fuerza una renovación con el refresh token aunque el access token siga vigente.
// Microsoft invalida el refresh token tras 90 días sin uso; renovarlo a menudo lo mantiene vivo.
export async function forceRefresh() {
  cachedAccessToken = null;
  cachedExpiresAt = 0;
  return getAccessToken();
}

export async function getAccessToken() {
  if (cachedAccessToken && Date.now() < cachedExpiresAt) return cachedAccessToken;
  return AUTH_MODE === "app" ? getAppToken() : getDelegatedToken();
}
