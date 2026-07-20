// Configuración central. Lee variables de entorno y valida lo mínimo imprescindible.

export const AUTH_MODE = (process.env.AUTH_MODE || "delegated").toLowerCase(); // "delegated" | "app"

export const TENANT_ID = process.env.TENANT_ID || "common";
export const CLIENT_ID = process.env.CLIENT_ID || "";
export const CLIENT_SECRET = process.env.CLIENT_SECRET || ""; // sólo modo "app"

// Refresh token semilla para modo delegado en hosts efímeros (se obtiene con `npm run login`).
export const SEED_REFRESH_TOKEN = process.env.GRAPH_REFRESH_TOKEN || "";

// Dónde persistir el refresh token rotado (best-effort; en hosts efímeros puede perderse en redeploy).
export const TOKEN_FILE = process.env.TOKEN_FILE || "./.token.json";

// Protección del endpoint público. Puede llegar por header Authorization: Bearer <X> o por ?key=<X>.
export const MCP_AUTH_TOKEN = process.env.MCP_AUTH_TOKEN || "";

// Lista blanca de drives (bibliotecas) que el servidor tiene permitido tocar.
// Si está VACÍA -> sin restricción (acceso a todo lo que el usuario puede, incluido su OneDrive).
// Si tiene valores (driveIds separados por comas) -> el servidor SOLO opera sobre esas bibliotecas
// y bloquea el OneDrive personal ('me') y cualquier otro drive. Ideal para compartir con compañeros
// sin exponer tus cosas privadas.
export const ALLOWED_DRIVE_IDS = (process.env.ALLOWED_DRIVE_IDS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

// Si es true: bloquea el OneDrive personal ('me' y los resultados de búsqueda en /personal/)
// pero deja libre TODO SharePoint. Ideal cuando el equipo trabaja en SharePoint y no se quiere
// exponer el OneDrive privado del titular de la cuenta.
export const BLOCK_ONEDRIVE = (process.env.BLOCK_ONEDRIVE || "false").toLowerCase() === "true";

export const PORT = parseInt(process.env.PORT || "3000", 10);

// Scopes delegados. offline_access es imprescindible para obtener refresh token.
export const DELEGATED_SCOPES =
  process.env.DELEGATED_SCOPES ||
  "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All";

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export function authority() {
  return `https://login.microsoftonline.com/${TENANT_ID}`;
}

export function tokenEndpoint() {
  return `${authority()}/oauth2/v2.0/token`;
}

export function deviceCodeEndpoint() {
  return `${authority()}/oauth2/v2.0/devicecode`;
}

export function validateStartup() {
  const errors = [];
  if (!CLIENT_ID) errors.push("Falta CLIENT_ID");
  if (AUTH_MODE === "app") {
    if (!CLIENT_SECRET) errors.push("Modo app: falta CLIENT_SECRET");
    if (TENANT_ID === "common")
      errors.push("Modo app: TENANT_ID debe ser el id concreto del tenant, no 'common'");
  } else {
    if (!SEED_REFRESH_TOKEN)
      errors.push(
        "Modo delegado: falta GRAPH_REFRESH_TOKEN (ejecuta `npm run login` una vez y copia el valor)"
      );
  }
  if (!MCP_AUTH_TOKEN)
    errors.push("Falta MCP_AUTH_TOKEN (protege el endpoint público; genera una cadena aleatoria larga)");
  return errors;
}
