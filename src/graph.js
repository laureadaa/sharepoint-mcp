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

async function graphFetch(path, { method = "GET", query, json, rawBody, contentType } = {}) {
  const token = await getAccessToken();
  let url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  if (query) {
    const qs = new URLSearchParams(query).toString();
    url += (url.includes("?") ? "&" : "?") + qs;
  }
  const headers = { Authorization: `Bearer ${token}` };
  let body;
  if (json !== undefined) {
    headers["Content-Type"] = "application/json";
    body = JSON.stringify(json);
  } else if (rawBody !== undefined) {
    headers["Content-Type"] = contentType || "application/octet-stream";
    body = rawBody;
  }
  const res = await fetch(url, { method, headers, body });
  return res;
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
  return `${base}/root:/${clean.split("/").map(encodeURIComponent).join("/")}`;
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

  // Lista el contenido de una carpeta.
  async listChildren(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    const data = await graphJson(`${sel}/children`, {
      query: { $select: "id,name,size,folder,file,webUrl,lastModifiedDateTime" },
    });
    return (data.value || []).map(mapItem);
  },

  // Metadatos de un item.
  async getItem(driveId, ref) {
    const sel = itemSelector(driveId, ref);
    return mapItem(await graphJson(sel));
  },

  // Busca archivos dentro de un drive.
  async searchFiles(driveId, queryText) {
    const base = driveRoot(driveId);
    const data = await graphJson(`${base}/root/search(q='${encodeURIComponent(queryText)}')`);
    return (data.value || []).map(mapItem);
  },

  // Búsqueda GLOBAL en TODO: OneDrive + todas las bibliotecas de SharePoint a la vez.
  // Usa la API de Microsoft Search. Encuentra archivos y carpetas estén donde estén.
  async searchAll(queryText, size = 25) {
    const body = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString: queryText },
          from: 0,
          size: Math.min(Math.max(size, 1), 50),
        },
      ],
    };
    const data = await graphJson("/search/query", { method: "POST", json: body });
    const hits = data?.value?.[0]?.hitsContainers?.[0]?.hits || [];
    let results = hits.map((h) => {
      const r = h.resource || {};
      return {
        name: r.name,
        isFolder: !!r.folder,
        size: r.size,
        webUrl: r.webUrl,
        driveId: r.parentReference?.driveId,
        itemId: r.id,
        path: r.parentReference?.path,
        lastModified: r.lastModifiedDateTime,
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
  },

  // Descarga el contenido de un archivo como texto (o base64 si no es texto).
  async readFile(driveId, ref, asBase64 = false) {
    const sel = itemSelector(driveId, ref);
    const res = await graphFetch(`${sel}/content`);
    if (!res.ok) {
      const t = await res.text().catch(() => "");
      throw new Error(`Graph ${res.status} al leer archivo: ${t}`);
    }
    const buf = Buffer.from(await res.arrayBuffer());
    return asBase64 ? buf.toString("base64") : buf.toString("utf8");
  },

  // Sube/crea o sobrescribe un archivo pequeño (<4MB) por ruta.
  async uploadFile(driveId, parentRef, name, content, isBase64 = false) {
    const base = driveRoot(driveId);
    const parent = !parentRef || parentRef === "/" || parentRef === "root" ? "" : parentRef.replace(/^\/+/, "") + "/";
    const fullPath = `${parent}${name}`;
    const url = `${base}/root:/${fullPath.split("/").map(encodeURIComponent).join("/")}:/content`;
    const body = isBase64 ? Buffer.from(content, "base64") : Buffer.from(content, "utf8");
    const data = await graphJson(url, {
      method: "PUT",
      rawBody: body,
      contentType: "application/octet-stream",
    });
    return mapItem(data);
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
