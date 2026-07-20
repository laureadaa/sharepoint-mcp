// Login interactivo por Device Code (una sola vez, sin permisos de admin).
// Obtiene un refresh token que luego pones en la variable GRAPH_REFRESH_TOKEN del servidor.
//
// Uso:  CLIENT_ID=... TENANT_ID=... npm run login
// (TENANT_ID puede ser el id de tu tenant, o 'organizations' para cuentas de trabajo.)

import { CLIENT_ID, TENANT_ID, DELEGATED_SCOPES, deviceCodeEndpoint, tokenEndpoint } from "./config.js";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  if (!CLIENT_ID) {
    console.error("Falta CLIENT_ID. Ejecuta: CLIENT_ID=<tu-app-id> TENANT_ID=<tenant> npm run login");
    process.exit(1);
  }

  // 1) Solicitar device code.
  const dcRes = await fetch(deviceCodeEndpoint(), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT_ID, scope: DELEGATED_SCOPES }).toString(),
  });
  const dc = await dcRes.json();
  if (!dcRes.ok) {
    console.error("Error solicitando device code:", dc.error_description || dc.error);
    console.error("¿Has activado 'Allow public client flows' = Yes en la app de Entra?");
    process.exit(1);
  }

  console.log("\n==================== INICIO DE SESIÓN ====================");
  console.log(dc.message || `Abre ${dc.verification_uri} e introduce el código: ${dc.user_code}`);
  console.log("=========================================================\n");
  console.log("Esperando a que completes el login en el navegador...\n");

  // 2) Sondear el endpoint de token.
  const interval = (dc.interval || 5) * 1000;
  const deadline = Date.now() + (dc.expires_in || 900) * 1000;

  while (Date.now() < deadline) {
    await sleep(interval);
    const tRes = await fetch(tokenEndpoint(), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: CLIENT_ID,
        device_code: dc.device_code,
      }).toString(),
    });
    const t = await tRes.json();
    if (tRes.ok) {
      console.log("✅ ¡Login correcto!\n");
      console.log("Copia este REFRESH TOKEN y ponlo como variable de entorno GRAPH_REFRESH_TOKEN en tu host:\n");
      console.log("------------------------------------------------------------");
      console.log(t.refresh_token);
      console.log("------------------------------------------------------------\n");
      console.log("Tenant usado:", TENANT_ID);
      console.log("Scopes:", DELEGATED_SCOPES);
      return;
    }
    if (t.error === "authorization_pending") continue;
    if (t.error === "slow_down") {
      await sleep(interval);
      continue;
    }
    console.error("Error de autenticación:", t.error_description || t.error);
    process.exit(1);
  }
  console.error("Se agotó el tiempo de espera. Vuelve a intentarlo.");
  process.exit(1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
