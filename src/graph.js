// Wrapper fino sobre Microsoft Graph con el token del gestor de auth.

import { getAccessToken } from "./auth.js";
import { GRAPH_BASE, AUTH_MODE, ALLOWED_DRIVE_IDS, BLOCK_ONEDRIVE } from "./config.js";

// Comprueba que el drive solicitado esté permitido según la configuración de seguridad.
function ensureDriveAllowed(driveId) {
  // Modo "bloquear OneDrive": deja SharePoint libre pero corta el OneDrive personal.
  if (BLOCK_ONEDRIVE && driveId === "me") {
    throw new Error(
      "Acceso restringido: el OneDrive personal no está disponible en este servidor. " +
        "Trabaja sobre las bibliotecas de SharePoint (usa search_all o search_sites)."
    );
  }
  // Modo lista blanca: solo los drives autorizados; bloquea 'me' y todo lo demás.
  if (ALLOWED_DRIVE_IDS.length === 0) return; // sin lista blanca
  if (driveId === "me") {
    throw new Error(
      "Acceso restringido: el OneDrive personal ('me') no está disponible en este servidor. " +
        "Usa una de las bibliotecas autorizadas (list_drives sobre el sitio del equipo)."
    );
  }
  if (!ALLOWED_DRIVE_IDS.includes(driveId)) {
    throw new Error(
      "Acceso restringido: este servidor solo puede operar sobre las bibliotecas autorizadas del equipo."
    );
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const RETRY_STATUS = new Set([429, 500, 502, 503, 504]);
const MAX_RETRIES = 4;
const REQUEST_TIMEOUT_MS = 60_000;

// Límites de tamaño: la subida simple de Graph admite hasta 4 MB; por encima, sesión por trozos.
const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024;
const CHUNK_SIZE = 5 * 1024 * 1024; // múltiplo de 320 KiB, como exige Graph
const READ_MAX_BYTES = 25 * 1024 * 1024;
const TEXT_MAX_CHARS = 200_000;

// Petición a Graph con reintentos ante limitación (429) y errores transitorios (5xx),
// respetando Retry-After. `auth: false` para URLs pre-firmadas (sesiones de subida).
async function graphFetch(path, { method = "GET", query, json, rawBody, contentType, headers: extra, auth = true } = {}) {
  let url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  if (query) {
    const qs = new URLSearchParams(query).toString();
    url += (url.includes("?") ? "&" : "?") + qs;
  }
  let body;
  const baseHeaders = { ...extra };
  if (json !== undefined) {
    baseHeaders["Content-Type"] = "application/json";
    body = JSON.stringify(json);
  } else if (rawBody !== undefined) {
    baseHeaders["Content-Type"] = contentType || "application/octet-stream";
    body = rawBody;
  }

  for (let attempt = 0; ; attempt++) {
    const headers = auth ? { ...baseHeaders, Authorization: `Bearer ${await getAccessToken()}` } : baseHeaders;
    let res;
    try {
      res = await fetch(url, { method, headers, body, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
    } catch (e) {
      if (attempt >= MAX_RETRIES) throw new Error(`Sin respuesta de Microsoft Graph: ${e.message}`);
      await sleep(backoff(attempt));
      continue;
    }
    if (!RETRY_STATUS.has(res.status) || attempt >= MAX_RETRIES) return res;
    const retryAfter = Number(res.headers.get("retry-after"));
    await res.body?.cancel().catch(() => {});
    await sleep(retryAfter > 0 ? Math.min(retryAfter, 60) * 1000 : backoff(attempt));
  }
}

const backoff = (attempt) => 1000 * 2 ** attempt + Math.random() * 500;

// Sigue @odata.nextLink hasta reunir `limit` elementos.
async function graphList(path, opts, limit = 500) {
  const items = [];
  let data = await graphJson(path, opts);
  items.push(...(data.value || []));
  while (data["@odata.nextLink"] && items.length < limit) {
    data = await graphJson(data["@odata.nextLink"]);
    items.push(...(data.value || []));
  }
  return { items: items.slice(0, limit), truncated: items.length > limit || !!data["@odata.nextLink"] };
}

async function graphJson(path, opts) {
  const res = await graphFetch(path, opts);
  const text = await res.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }
  if (!res.ok) {
    const msg = data?.error?.message || data?.raw || `HTTP ${res.status}`;
    const err = new Error(`Graph ${res.status}: ${msg}`);
    err.status = res.status;
    throw err;
  }
  return data;
}

// Resuelve el segmento de drive. "me" sólo es válido en modo delegado.
function driveRoot(driveId) {
  ensureDriveAllowed(driveId);
  if (driveId === "me") {
    if (AUTH_MODE !== "delegated") {
      throw new Error("driveId 'me' sólo está disponible en modo delegado. Usa un driveId real o un siteId.");
    }
    return "/me/drive";
  }
  return `/drives/${encodeURIComponent(driveId)}`;
}

// Construye el selector de un item por id o por ruta relativa a la raíz del drive.
// Si `ref` empieza por "id:" se trata como id; en otro caso como ruta ("/Carpeta/archivo.txt").
function itemSelector(driveId, ref) {
  const base = driveRoot(driveId);
  if (!ref || ref === "/" || ref === "root") return `${base}/root`;
  if (ref.startsWith("id:")) return `${base}/items/${encodeURIComponent(ref.slice(3))}`;
  const clean = ref.replace(/^\/+/, "");
  // Forma documentada de Graph: root:/ruta: (los ':' cierran la ruta antes de /children, /content…).
  return `${base}/root:/${clean.replace(/\/+$/, "").split("/").map(encodeURIComponent).join("/")}:`;
}

export const graph = {
  // Identidad del usuario (modo delegado).
  async whoami() {
    return graphJson("/me");
  },

  // Busca sitios de SharePoint por texto.
  async searchSites(queryText) {
    const data = await graphJson("/sites", { query: { search: queryText || "*" } });
    return (data.value || []).map((s) => ({
      id: s.id,
      name: s.name || s.displayName,
      webUrl: s.webUrl,
    }));
  },

  // Lista los drives (bibliotecas de documentos) de un sitio, o del OneDrive del usuario ("me").
  async listDrives(target) {
    if (target === "me") ensureDriveAllowed("me"); // bloquea listar OneDrive si hay lista blanca
    const path = target === "me" ? "/me/drives" : `/sites/${encodeURIComponent(target)}/drives`;
    const data = await graphJson(path);
    return (data.value || []).map((d) => ({
      id: d.id,
      name: d.name,
      driveType: d.driveType,
      webUrl: d.webUrl,
    }));
  },

  // Lista el contenido de una carpeta (todas las páginas, hasta `limit`).
  async listChildren(driveId, ref, limit = 500) {
    const sel = itemSelector(driveId, ref);
    const { items, truncated } = await graphList(
      `${sel}/children`,
      { query: { $select: "id,name,size,folder,file,webUrl,lastModifiedDateTime,parentReference", $top: "200" } },
      limit
    );
    return { count: items.length, truncated, items: items.map(mapItem) };
  },

  // Metadatos de un item.
  async getItem(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    return mapItem(await graphJson(sel));
  },

  // Busca archivos dentro de un drive.
  async searchFiles(driveId, queryText) {
    const base = driveRoot(driveId);
    // En OData las comillas simples se escapan duplicándolas (ej. D'Amico → D''Amico).
    const q = encodeURIComponent(String(queryText).replace(/'/g, "''"));
    const { items } = await graphList(`${base}/root/search(q='${q}')`, undefined, 200);
    return items.map(mapItem);
  },

  // Búsqueda GLOBAL en TODO: OneDrive + todas las bibliotecas de SharePoint a la vez.
  // Usa la API de Microsoft Search. Encuentra archivos y carpetas estén donde estén.
  async searchAll(queryText, size = 25) {
    return microsoftSearch(queryText, size);
  },

  // Lo último modificado: en todo el despacho o dentro de una carpeta, ordenado del más reciente al más antiguo.
  async recentChanges({ days = 7, driveId, ref, query, size = 25 }) {
    const since = new Date(Date.now() - days * 86400_000).toISOString().slice(0, 10);
    const parts = [`LastModifiedTime>=${since}`];
    if (driveId) {
      const folder = await graphJson(itemSelector(driveId, ref), { query: { $select: "webUrl,folder" } });
      if (!folder.folder) throw new Error("`ref` debe ser una carpeta.");
      parts.push(`path:"${decodeURI(folder.webUrl)}"`);
    }
    if (query) parts.push(`(${query})`);
    return microsoftSearch(parts.join(" AND "), size, [{ name: "lastModifiedDateTime", isDescending: true }]);
  },

  // Hojas de un libro Excel (sólo .xlsx/.xlsm, la API de libros de Graph no admite .xls).
  async excelSheets(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    const data = await graphJson(`${sel}/workbook/worksheets`, { query: { $select: "id,name,position,visibility" } });
    return (data.value || []).map((w) => ({ name: w.name, position: w.position, visibility: w.visibility, id: w.id }));
  },

  // Lee una hoja (o un rango A1) de un Excel como tabla: filas y columnas reales.
  async readExcel(driveId, ref, { sheet, range, maxRows = 500, formatted = true } = {}) {
    const sel = itemSelector(driveId, ref);
    const sheets = await this.excelSheets(driveId, ref);
    if (!sheets.length) throw new Error("El libro no tiene hojas.");
    const ws = sheet ? sheets.find((s) => s.name.toLowerCase() === sheet.toLowerCase()) : sheets[0];
    if (!ws) throw new Error(`No existe la hoja '${sheet}'. Hojas disponibles: ${sheets.map((s) => s.name).join(", ")}`);
    const wsPath = `${sel}/workbook/worksheets/${encodeURIComponent(ws.id)}`;
    const rangePath = range
      ? `${wsPath}/range(address='${range.replace(/'/g, "''")}')`
      : `${wsPath}/usedRange(valuesOnly=true)`;
    const data = await graphJson(rangePath, { query: { $select: "address,rowCount,columnCount,values,text" } });
    const rows = (formatted ? data.text : data.values) || [];
    return {
      sheet: ws.name,
      sheets: sheets.map((s) => s.name),
      address: data.address,
      rowCount: data.rowCount,
      columnCount: data.columnCount,
      truncated: rows.length > maxRows,
      rows: rows.slice(0, maxRows),
    };
  },

  // Historial de versiones de un archivo.
  async listVersions(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    const { items } = await graphList(`${sel}/versions`, undefined, 200);
    return items.map((v) => ({
      versionId: v.id,
      lastModified: v.lastModifiedDateTime,
      modifiedBy: v.lastModifiedBy?.user?.displayName || v.lastModifiedBy?.user?.email,
      size: v.size,
    }));
  },

  // Restaura una versión anterior (SharePoint crea una versión nueva con ese contenido; no se pierde nada).
  async restoreVersion(driveId, ref, versionId) {
    const sel = itemSelector(driveId, ref);
    const res = await graphFetch(`${sel}/versions/${encodeURIComponent(versionId)}/restoreVersion`, { method: "POST" });
    if (!res.ok) {
      const t = await res.text().catch(() => "");
      throw new Error(`Graph ${res.status} al restaurar versión: ${t}`);
    }
    return { restored: true, versionId, item: await this.getItem(driveId, ref) };
  },

  // Quién tiene acceso a un archivo o carpeta: usuarios, grupos y enlaces compartidos.
  async listPermissions(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    const { items } = await graphList(`${sel}/permissions`, undefined, 500);
    return items.map((p) => {
      const who = p.grantedToV2 || p.grantedTo;
      const identities = p.grantedToIdentitiesV2 || p.grantedToIdentities || [];
      return {
        permissionId: p.id,
        roles: p.roles,
        kind: p.link ? "enlace" : "directo",
        inherited: !!p.inheritedFrom,
        grantedTo:
          who?.user?.displayName || who?.user?.email || who?.siteGroup?.displayName || who?.group?.displayName ||
          who?.siteUser?.displayName || undefined,
        email: [who?.user?.email, who?.siteUser?.loginName?.split("|").pop()].find((e) => e?.includes("@")),
        link: p.link && {
          type: p.link.type,
          scope: p.link.scope, // anonymous = cualquiera con el enlace
          webUrl: p.link.webUrl,
          expires: p.expirationDateTime,
          hasPassword: p.hasPassword,
          sharedWith: identities.map((i) => i.user?.displayName || i.user?.email).filter(Boolean),
        },
      };
    });
  },

  // Revoca un enlace compartido. Sólo enlaces (no usuarios ni grupos) y sólo los no heredados.
  async revokeShareLink(driveId, ref, permissionId) {
    const sel = itemSelector(driveId, ref);
    const p = await graphJson(`${sel}/permissions/${encodeURIComponent(permissionId)}`);
    if (!p.link) throw new Error("Ese permiso no es un enlace compartido; esta herramienta sólo revoca enlaces.");
    if (p.inheritedFrom) throw new Error("El enlace se hereda de una carpeta superior: revócalo en esa carpeta.");
    const res = await graphFetch(`${sel}/permissions/${encodeURIComponent(permissionId)}`, { method: "DELETE" });
    if (!res.ok) {
      const t = await res.text().catch(() => "");
      throw new Error(`Graph ${res.status} al revocar el enlace: ${t}`);
    }
    return { revoked: true, permissionId, scope: p.link.scope, type: p.link.type };
  },

  // Lee un archivo. format:
  //   "auto"   → texto plano tal cual; Word/Excel/PowerPoint/PDF… se extrae su texto.
  //   "text"   → fuerza decodificar como UTF-8.
  //   "base64" → bytes originales en base64.
  async readFile(driveId, ref, format = "auto") {
    const sel = itemSelector(driveId, ref);
    const meta = await graphJson(sel, { query: { $select: "id,name,size,file,folder" } });
    if (meta.folder) throw new Error("Es una carpeta, no un archivo. Usa list_children.");
    if (meta.size > READ_MAX_BYTES) {
      throw new Error(`Archivo demasiado grande para leerlo (${mb(meta.size)} MB; máximo ${mb(READ_MAX_BYTES)} MB).`);
    }
    const ext = (meta.name.match(/\.([^.]+)$/)?.[1] || "").toLowerCase();
    const base = { name: meta.name, size: meta.size, mimeType: meta.file?.mimeType };

    if (format === "base64") {
      return { ...base, encoding: "base64", content: (await download(`${sel}/content`)).toString("base64") };
    }
    if (format === "auto" && (ext === "pdf" || PDF_CONVERTIBLE.has(ext))) {
      // Office → PDF lo convierte el propio SharePoint; luego extraemos el texto del PDF.
      const pdf = await download(ext === "pdf" ? `${sel}/content` : `${sel}/content?format=pdf`);
      const { extractText } = await import("unpdf");
      const { totalPages, text } = await extractText(new Uint8Array(pdf), { mergePages: false });
      const joined = text.map((t, i) => `--- Página ${i + 1} ---\n${t.trim()}`).join("\n\n");
      return { ...base, encoding: "text", extractedFrom: ext, pages: totalPages, ...clip(joined) };
    }
    if (format === "auto" && BINARY_EXT.has(ext)) {
      throw new Error(`'${meta.name}' es binario y no se puede leer como texto. Usa format="base64".`);
    }
    return { ...base, encoding: "utf8", ...clip((await download(`${sel}/content`)).toString("utf8")) };
  },

  // Crea o sobrescribe un archivo. Hasta 4 MB subida directa; por encima, sesión por trozos.
  async uploadFile(driveId, parentRef, name, content, isBase64 = false) {
    const parentSel = itemSelector(driveId, parentRef);
    // Carpeta por ruta (root:/A/B:) → root:/A/B/nombre: ; raíz o id (items/XYZ) → …:/nombre:
    const target = parentSel.endsWith(":")
      ? `${parentSel.slice(0, -1)}/${encodeURIComponent(name)}:`
      : `${parentSel}:/${encodeURIComponent(name)}:`;
    const body = isBase64 ? Buffer.from(content, "base64") : Buffer.from(content, "utf8");

    if (body.length <= SIMPLE_UPLOAD_MAX) {
      const data = await graphJson(`${target}/content`, {
        method: "PUT",
        rawBody: body,
        contentType: "application/octet-stream",
      });
      return mapItem(data);
    }

    const session = await graphJson(`${target}/createUploadSession`, {
      method: "POST",
      json: { item: { "@microsoft.graph.conflictBehavior": "replace" } },
    });
    let last;
    for (let start = 0; start < body.length; start += CHUNK_SIZE) {
      const chunk = body.subarray(start, Math.min(start + CHUNK_SIZE, body.length));
      // La URL de sesión ya va firmada: no se envía Authorization.
      last = await graphFetch(session.uploadUrl, {
        method: "PUT",
        auth: false,
        rawBody: chunk,
        headers: { "Content-Range": `bytes ${start}-${start + chunk.length - 1}/${body.length}` },
      });
      if (!last.ok) {
        const t = await last.text().catch(() => "");
        await fetch(session.uploadUrl, { method: "DELETE" }).catch(() => {});
        throw new Error(`Graph ${last.status} subiendo el trozo que empieza en ${start}: ${t}`);
      }
    }
    return mapItem(await last.json());
  },

  // Crea una carpeta.
  async createFolder(driveId, parentRef, name) {
    const sel = itemSelector(driveId, parentRef);
    const data = await graphJson(`${sel}/children`, {
      method: "POST",
      json: { name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
    });
    return mapItem(data);
  },

  // Elimina un item (a la papelera de reciclaje).
  async deleteItem(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    const res = await graphFetch(sel, { method: "DELETE" });
    if (!res.ok && res.status !== 204) {
      const t = await res.text().catch(() => "");
      throw new Error(`Graph ${res.status} al borrar: ${t}`);
    }
    return { deleted: true };
  },

  // Mueve (y opcionalmente renombra) un item a otra carpeta del mismo drive.
  async moveItem(driveId, ref, newParentRef, newName) {
    const sel = itemSelector(driveId, ref);
    const parent = await this.getItem(driveId, newParentRef);
    const patch = { parentReference: { id: parent.id } };
    if (newName) patch.name = newName;
    const data = await graphJson(sel, { method: "PATCH", json: patch });
    return mapItem(data);
  },

  // Copia un item a otra carpeta (operación asíncrona en Graph; devolvemos el estado inicial).
  async copyItem(driveId, ref, newParentRef, newName) {
    const sel = itemSelector(driveId, ref);
    const parent = await this.getItem(driveId, newParentRef);
    const body = { parentReference: { driveId: parent.driveId || undefined, id: parent.id } };
    if (newName) body.name = newName;
    const res = await graphFetch(`${sel}/copy`, { method: "POST", json: body });
    if (res.status === 202) {
      return { accepted: true, monitorUrl: res.headers.get("location") };
    }
    const t = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status} al copiar: ${t}`);
  },

  // Crea un enlace para compartir (view o edit).
  async createShareLink(driveId, ref, type = "view") {
    const sel = itemSelector(driveId, ref);
    const data = await graphJson(`${sel}/createLink`, {
      method: "POST",
      json: { type, scope: "organization" },
    });
    return { url: data?.link?.webUrl, type, scope: data?.link?.scope };
  },
};

// Microsoft Search sobre driveItem, aplicando las mismas restricciones de drives que el resto.
async function microsoftSearch(queryString, size = 25, sortProperties) {
  const request = {
    entityTypes: ["driveItem"],
    query: { queryString },
    from: 0,
    size: Math.min(Math.max(size, 1), 50),
  };
  if (sortProperties) request.sortProperties = sortProperties;
  const data = await graphJson("/search/query", { method: "POST", json: { requests: [request] } });
  const hits = data?.value?.[0]?.hitsContainers?.[0]?.hits || [];
  let results = hits.map((h) => {
    const r = h.resource || {};
    return {
      name: r.name,
      // Microsoft Search no devuelve la faceta `folder`: sin `file`, tamaño 0 y sin extensión → carpeta.
      isFolder: !!r.folder || (!r.file && !r.size && !/\.[a-z0-9]{1,5}$/i.test(r.name || "")),
      size: r.size,
      webUrl: r.webUrl,
      driveId: r.parentReference?.driveId,
      itemId: r.id,
      path: r.parentReference?.path,
      lastModified: r.lastModifiedDateTime,
      modifiedBy: r.lastModifiedBy?.user?.displayName,
    };
  });
  // Si hay lista blanca de drives, respétala también en la búsqueda global.
  if (ALLOWED_DRIVE_IDS.length > 0) {
    results = results.filter((r) => r.driveId && ALLOWED_DRIVE_IDS.includes(r.driveId));
  }
  // Si se bloquea el OneDrive personal, oculta sus resultados (URLs de OneDrive: .../personal/...).
  if (BLOCK_ONEDRIVE) {
    results = results.filter((r) => !(r.webUrl || "").toLowerCase().includes("/personal/"));
  }
  return results;
}

// Formatos que SharePoint sabe convertir a PDF (?format=pdf).
const PDF_CONVERTIBLE = new Set([
  "doc", "docx", "docm", "dot", "dotx", "rtf", "odt",
  "xls", "xlsx", "xlsm", "xlsb", "ods",
  "ppt", "pptx", "pptm", "pps", "ppsx", "odp",
]);
const BINARY_EXT = new Set([
  "png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp", "heic",
  "zip", "rar", "7z", "exe", "dll", "msg", "mp3", "mp4", "mov", "avi", "wav",
]);

const mb = (n) => (n / 1024 / 1024).toFixed(1);

function clip(text) {
  if (text.length <= TEXT_MAX_CHARS) return { content: text };
  return { content: text.slice(0, TEXT_MAX_CHARS), truncated: true, totalChars: text.length };
}

async function download(path) {
  const res = await graphFetch(path);
  if (!res.ok) {
    const t = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status} al leer archivo: ${t}`);
  }
  return Buffer.from(await res.arrayBuffer());
}

function mapItem(i) {
  if (!i) return i;
  return {
    id: i.id,
    name: i.name,
    isFolder: !!i.folder,
    size: i.size,
    childCount: i.folder?.childCount,
    webUrl: i.webUrl,
    driveId: i.parentReference?.driveId,
    lastModified: i.lastModifiedDateTime,
  };
}
