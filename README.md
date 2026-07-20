# Servidor MCP: SharePoint + OneDrive (lectura y escritura)

Un servidor MCP remoto que conecta Claude con tu **SharePoint** y tu **OneDrive** de Microsoft 365,
con **lectura y escritura** (crear, subir, editar, mover, borrar), accesible **desde cualquier
dispositivo** porque vive en la nube, no en tu ordenador.

Autenticación por defecto: **flujo delegado (Device Code), sin permisos de administrador**. Inicias
sesión una sola vez con tu cuenta y el servidor renueva el acceso solo. Incluye también un modo
`app-only` (Client Credentials) por si prefieres acceso totalmente desatendido (ese sí requiere
consentimiento de administrador).

## Herramientas incluidas

`whoami`, `search_sites`, `list_drives`, `list_children`, `get_item`, `search_files`, `read_file`,
`upload_file`, `create_folder`, `delete_item`, `move_item`, `copy_item`, `create_share_link`.

---

## Requisitos

- Una cuenta de Microsoft 365 (trabajo/escuela).
- Node.js 20+ instalado en tu ordenador (sólo para el login inicial).
- Una cuenta en un host con URL pública HTTPS (recomiendo **Render**; alternativas: Railway, Azure).

---

## Paso 1 — Registrar la aplicación en Entra ID (sin admin)

1. Entra en el [portal de Entra ID](https://entra.microsoft.com) → **Identidad → Aplicaciones →
   Registros de aplicaciones → Nuevo registro**.
2. Nombre: `Claude SharePoint MCP`. Tipo de cuenta: *Cuentas en este directorio organizativo
   únicamente* (single tenant). No hace falta URI de redirección. Pulsa **Registrar**.
3. En la página **Información general**, copia el **Id. de aplicación (cliente)** y el
   **Id. de directorio (inquilino)**. Los necesitarás como `CLIENT_ID` y `TENANT_ID`.
4. Ve a **Autenticación** → sección *Configuración avanzada* → **Permitir flujos de cliente
   público** = **Sí**. Guarda. (Esto habilita el Device Code sin secreto.)
5. Ve a **Permisos de API → Agregar un permiso → Microsoft Graph → Permisos delegados** y añade:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `offline_access`
   - `User.Read`
   No hace falta pulsar "Conceder consentimiento de administrador": **estos permisos los consientes
   tú al iniciar sesión** en el Paso 2.

> ¿Prefieres app-only (desatendido, requiere admin)? Ve a la sección **Modo app-only** al final.

---

## Paso 2 — Obtener tu refresh token (login único)

En tu ordenador, dentro de la carpeta del proyecto:

```bash
npm install
CLIENT_ID=<tu-app-id> TENANT_ID=<tu-tenant-id> npm run login
```

Te mostrará un enlace y un código. Ábrelo, inicia sesión con tu cuenta de Microsoft 365 y acepta los
permisos. Al terminar, la consola imprime un **REFRESH TOKEN**. Cópialo: será la variable
`GRAPH_REFRESH_TOKEN`.

---

## Paso 3 — Generar el token que protege el endpoint

El servidor es público, así que se protege con un token propio. Genera uno:

```bash
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
```

Guárdalo: será la variable `MCP_AUTH_TOKEN`.

---

## Paso 4 — Desplegar (URL pública HTTPS)

### Opción recomendada: Render

1. Sube esta carpeta a un repositorio de GitHub.
2. En [Render](https://render.com): **New → Web Service** → conecta el repo → runtime **Docker**.
3. En **Environment**, añade las variables:
   - `AUTH_MODE=delegated`
   - `CLIENT_ID=<tu-app-id>`
   - `TENANT_ID=<tu-tenant-id>`
   - `GRAPH_REFRESH_TOKEN=<del paso 2>`
   - `MCP_AUTH_TOKEN=<del paso 3>`
4. Deploy. Render te dará una URL tipo `https://sharepoint-onedrive-mcp.onrender.com`.
5. Comprueba salud: abre `https://TU-URL/health` → debe responder `{"status":"ok","mode":"delegated"}`.

Tu **endpoint MCP** será:

```
https://TU-URL/mcp?key=<MCP_AUTH_TOKEN>
```

### Alternativa rápida para probar: local + túnel

```bash
cp .env.example .env   # rellena las variables
npm start              # arranca en localhost:3000
# en otra terminal, expón con un túnel:
npx cloudflared tunnel --url http://localhost:3000
```

Cloudflared te da una URL pública temporal. Útil para una prueba rápida, pero depende de que tu PC
esté encendido (por eso, para el día a día, mejor Render).

---

## Paso 5 — Conectar el servidor a Claude

1. En claude.ai → **Ajustes → Conectores → Añadir conector personalizado**.
2. Nombre: `SharePoint / OneDrive`.
3. URL del servidor MCP:

   ```
   https://TU-URL/mcp?key=<MCP_AUTH_TOKEN>
   ```

   (El `?key=...` va incrustado en la URL, así no necesitas configurar cabeceras.)
4. Guarda. Claude hará el handshake y cargará las 13 herramientas.

> Como el conector vive en tu cuenta de Claude, a partir de aquí funciona **desde cualquier
> dispositivo** en el que uses Claude.

---

## Paso 6 — Probarlo de verdad, aplicado al día a día

Una vez conectado, pruébalo con peticiones reales en lenguaje natural. Sugerencia de secuencia,
de menos a más "arriesgado":

**1. Confirmar identidad y acceso (sólo lectura):**
> "Con el conector de SharePoint, dime con qué cuenta estás conectado y lista mis bibliotecas de OneDrive."

Debería usar `whoami` y `list_drives` con `target: "me"`.

**2. Explorar SharePoint:**
> "Busca el sitio de SharePoint de 'Marketing' y lista los archivos de su biblioteca de documentos."

Usará `search_sites` → `list_drives` → `list_children`.

**3. Leer un documento real:**
> "Lee el archivo 'Notas.txt' que está en la raíz de mi OneDrive y resúmelo."

Usará `read_file`. (Para .docx/.pdf pídele leerlo en base64 o convertirlo.)

**4. Primera escritura (empieza con algo inocuo):**
> "Crea una carpeta llamada 'Pruebas Claude' en mi OneDrive y dentro sube un archivo
> 'hola.txt' con el texto 'funciona'."

Usará `create_folder` y `upload_file`. Luego **compruébalo tú mismo** abriendo OneDrive en el
navegador: la carpeta y el archivo deben aparecer. Ése es el test definitivo de que la escritura
es real.

**5. Flujo completo del día a día:**
> "Genera un acta de reunión en Markdown con estos puntos [...] y súbela como 'acta-2026-07-20.md'
> a la carpeta 'Actas' del sitio de SharePoint de Operaciones. Luego dame un enlace para compartirla."

Encadena `search_sites` → `list_drives` → `upload_file` → `create_share_link`. Verifica el enlace
resultante desde otro dispositivo o pidiéndoselo a un compañero.

**Consejo de prueba segura:** haz las primeras escrituras en una carpeta de pruebas dedicada. Todo
lo que borres con `delete_item` va a la **papelera de reciclaje** de OneDrive/SharePoint (no se
pierde), así que es recuperable.

---

## Modo app-only (desatendido, requiere admin)

Si prefieres que funcione sin ningún login de usuario y sin caducidad:

1. En la app de Entra: **Certificados y secretos → Nuevo secreto de cliente**. Copia el valor →
   `CLIENT_SECRET`.
2. **Permisos de API → Permisos de aplicación** (no delegados): añade `Files.ReadWrite.All` y
   `Sites.ReadWrite.All`, y pulsa **Conceder consentimiento de administrador**.
3. Variables del host: `AUTH_MODE=app`, `CLIENT_ID`, `TENANT_ID` (el id real, no 'organizations'),
   `CLIENT_SECRET`, `MCP_AUTH_TOKEN`. No se usa `GRAPH_REFRESH_TOKEN`.
4. Nota: en app-only no existe "mi OneDrive" (`me`); accede por `siteId` o por el drive de un
   usuario concreto. Conviene además restringir a qué sitios puede acceder la app (Sites.Selected)
   para no darle acceso a todo el tenant.

---

## Seguridad

- El `MCP_AUTH_TOKEN` es lo único que separa tu SharePoint del mundo: que sea largo y aleatorio, y
  no lo compartas. Si se filtra, genera otro y actualízalo en el host y en el conector de Claude.
- Usa siempre HTTPS (Render y Azure lo dan por defecto).
- El refresh token delegado hereda **tus** permisos: el servidor sólo puede tocar lo que tú puedes
  tocar. Rótalo (repite el login) si sospechas que se ha filtrado.
- Considera desplegar en una región/host acorde a las políticas de datos de tu organización.

## Problemas frecuentes

- **`/health` responde pero el conector falla:** revisa que la URL del conector termine en
  `/mcp?key=...` con el token correcto.
- **401 en Claude:** el `key` de la URL no coincide con `MCP_AUTH_TOKEN` del host.
- **`invalid_grant` / el login caducó:** vuelve a ejecutar `npm run login` y actualiza
  `GRAPH_REFRESH_TOKEN`. Ocurre si el token estuvo inactivo mucho tiempo o cambiaste la contraseña.
- **`AADSTS7000218` al hacer login:** falta activar *Permitir flujos de cliente público = Sí*.
- **Archivos grandes (>4MB):** este servidor sube archivos pequeños por sencillez; para archivos
  grandes habría que añadir sesiones de subida por fragmentos (se puede ampliar).
