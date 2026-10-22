// Registro de herramientas MCP sobre el wrapper de Graph.

import { z } from "zod";
import { graph } from "./graph.js";
import { AUTH_MODE } from "./config.js";

const ok = (data) => ({ content: [{ type: "text", text: JSON.stringify(data, null, 2) }] });
const fail = (e) => ({ isError: true, content: [{ type: "text", text: `Error: ${e.message}` }] });

// Registro de auditoría de las operaciones que modifican datos (va a los logs del host).
function audit(action, details, error) {
  console.log(JSON.stringify({ audit: action, at: new Date().toISOString(), ok: !error, ...details, ...(error && { error: error.message }) }));
}

async function write(action, details, fn) {
  try {
    const result = await fn();
    audit(action, details);
    return ok(result);
  } catch (e) {
    audit(action, details, e);
    return fail(e);
  }
}

// `ref` acepta una ruta ("/Documentos/informe.docx") o un id ("id:01ABC...").
const refDesc =
  "Referencia al item: una ruta relativa a la raíz del drive (ej. '/Carpeta/archivo.docx') o un id con prefijo 'id:' (ej. 'id:01ABC...'). Usa '/' o 'root' para la raíz.";
const driveDesc =
  "ID del drive (biblioteca de documentos). Usa 'me' para tu OneDrive (sólo en modo delegado), o un driveId obtenido con list_drives.";

export const GUIDE_URL = "https://claude.ai/artifact/2Tj9xsL7gi5LQShRfhnrPE";

const GUIDE_TEXT = `Guía del conector SharePoint (versión completa con ejemplos: ${GUIDE_URL})

Qué puedes pedir, con palabras normales:
- Buscar en todo SharePoint: "Busca la carpeta del cliente…".
- Ver una carpeta: "¿Qué hay en Fiscal › Impuestos de este cliente?".
- Leer Word, PDF y PowerPoint: "Resume el informe…", "¿qué dice la cláusula 5…?".
- Consultar un Excel como tabla: "Lee la hoja Balance del sumas y saldos de 2025".
- Lo último modificado: "¿Qué se ha movido esta semana en este cliente?".
- Versiones: "¿Quién cambió este Excel?", "vuelve a la versión de ayer" (pide confirmación; no se pierde nada).
- Accesos: "¿Quién tiene acceso a esta carpeta?", "quita el enlace público".
- Subir, crear carpetas, mover, renombrar, copiar, compartir y borrar (el borrado va a la papelera de SharePoint).

Consejos: nombra el cliente y el documento como en el despacho, di qué quieres hacer con él y, en Excel, indica la hoja.

Límites: todo se hace con la cuenta del despacho (la cuenta de administración); el OneDrive personal está bloqueado; los PDF escaneados sin texto no se pueden leer; máximo 25 MB por archivo; los archivos recién subidos tardan unos minutos en salir en búsquedas.

Si falla: con "no encuentro el archivo", prueba otra palabra o di la carpeta; con un error de "token" o de inicio de sesión, avisa a quien administra el conector.`;

export function registerTools(server) {
  server.registerTool(
    "guia",
    {
      title: "Guía de uso del conector",
      description:
        "Devuelve la guía de uso del conector para el equipo (qué se puede pedir, consejos, límites, qué hacer si falla) y el enlace a la guía completa con ejemplos. Úsala cuando el usuario pregunte cómo funciona el conector, qué puede hacer con SharePoint, o no entienda algo del conector. Muéstrale el enlace a la guía completa.",
      inputSchema: {},
    },
    async () => ok({ guia: GUIDE_TEXT, guiaCompleta: GUIDE_URL })
  );

  if (AUTH_MODE === "delegated") {
    server.registerTool(
      "whoami",
      {
        title: "Identidad actual",
        description: "Devuelve la cuenta de usuario con la que está autenticado el servidor.",
        inputSchema: {},
      },
      async () => {
        try {
          const me = await graph.whoami();
          return ok({ displayName: me.displayName, mail: me.mail || me.userPrincipalName, id: me.id });
        } catch (e) {
          return fail(e);
        }
      }
    );
  }

  server.registerTool(
    "search_sites",
    {
      title: "Buscar sitios de SharePoint",
      description: "Busca sitios de SharePoint por texto. Devuelve id, nombre y URL de cada sitio.",
      inputSchema: { query: z.string().describe("Texto a buscar en los nombres de sitio. '*' lista los más relevantes.") },
    },
    async ({ query }) => {
      try {
        return ok(await graph.searchSites(query));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "list_drives",
    {
      title: "Listar bibliotecas/drives",
      description:
        "Lista los drives (bibliotecas de documentos). Pasa 'me' para tu OneDrive, o un siteId de SharePoint (de search_sites) para sus bibliotecas.",
      inputSchema: { target: z.string().describe("'me' para OneDrive, o el id de un sitio SharePoint.") },
    },
    async ({ target }) => {
      try {
        return ok(await graph.listDrives(target));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "list_children",
    {
      title: "Listar contenido de carpeta",
      description: "Lista archivos y subcarpetas dentro de una carpeta de un drive.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().default("root").describe(refDesc),
        limit: z.number().int().min(1).max(2000).default(500).describe("Máximo de elementos a devolver."),
      },
    },
    async ({ driveId, ref, limit }) => {
      try {
        return ok(await graph.listChildren(driveId, ref, limit));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "get_item",
    {
      title: "Metadatos de un item",
      description: "Devuelve metadatos (id, nombre, tamaño, URL, si es carpeta) de un archivo o carpeta.",
      inputSchema: { driveId: z.string().describe(driveDesc), ref: z.string().describe(refDesc) },
    },
    async ({ driveId, ref }) => {
      try {
        return ok(await graph.getItem(driveId, ref));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "search_files",
    {
      title: "Buscar archivos en un drive",
      description: "Busca archivos y carpetas por texto dentro de un drive.",
      inputSchema: { driveId: z.string().describe(driveDesc), query: z.string().describe("Texto a buscar.") },
    },
    async ({ driveId, query }) => {
      try {
        return ok(await graph.searchFiles(driveId, query));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "search_all",
    {
      title: "Buscar en TODO (global)",
      description:
        "Búsqueda global en TODO a la vez: tu OneDrive y todas las bibliotecas de SharePoint. Encuentra archivos y carpetas estén donde estén (ej. una carpeta de cliente). Devuelve driveId + itemId para poder abrir/listar el resultado después. Úsala cuando no sepas en qué sitio o biblioteca está algo.",
      inputSchema: {
        query: z.string().describe("Texto a buscar (nombre de cliente, archivo, carpeta, etc.)."),
        size: z.number().int().min(1).max(50).default(25).describe("Número máximo de resultados."),
      },
    },
    async ({ query, size }) => {
      try {
        return ok(await graph.searchAll(query, size));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "read_file",
    {
      title: "Leer archivo",
      description:
        "Lee un archivo. Con format='auto' (por defecto) devuelve el TEXTO de Word, Excel, PowerPoint y PDF ya extraído, por páginas, y los archivos de texto tal cual. Usa format='base64' para obtener los bytes originales. Máximo 25 MB.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        format: z
          .enum(["auto", "text", "base64"])
          .default("auto")
          .describe("'auto': extrae el texto de documentos; 'text': decodifica como UTF-8; 'base64': bytes originales."),
        as_base64: z.boolean().optional().describe("Obsoleto: equivale a format='base64'."),
      },
    },
    async ({ driveId, ref, format, as_base64 }) => {
      try {
        return ok({ ref, ...(await graph.readFile(driveId, ref, as_base64 ? "base64" : format)) });
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "read_excel",
    {
      title: "Leer Excel como tabla",
      description:
        "Lee una hoja de un Excel (.xlsx/.xlsm) como tabla: filas y columnas reales, con los valores tal como se ven en Excel. Sin `sheet` lee la primera hoja y devuelve también la lista de hojas. Usa `range` (ej. 'A1:F40') para leer sólo una parte. Ideal para sumas y saldos, libros de IVA, modelos…",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        sheet: z.string().optional().describe("Nombre de la hoja (por defecto la primera)."),
        range: z.string().optional().describe("Rango en formato A1, ej. 'A1:F40'. Por defecto, todo lo usado de la hoja."),
        max_rows: z.number().int().min(1).max(5000).default(500).describe("Máximo de filas a devolver."),
        formatted: z
          .boolean()
          .default(true)
          .describe("true: valores con formato de Excel (1.234,56 €, fechas). false: valores en bruto."),
      },
    },
    async ({ driveId, ref, sheet, range, max_rows, formatted }) => {
      try {
        return ok(await graph.readExcel(driveId, ref, { sheet, range, maxRows: max_rows, formatted }));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "recent_changes",
    {
      title: "Lo último modificado",
      description:
        "Archivos y carpetas modificados en los últimos N días, del más reciente al más antiguo, con quién los modificó. En todo el despacho, o sólo dentro de una carpeta (driveId + ref). Ej.: '¿qué se ha movido esta semana en este cliente?'.",
      inputSchema: {
        days: z.number().int().min(1).max(365).default(7).describe("Ventana en días."),
        driveId: z.string().optional().describe("Limitar a una carpeta: drive de esa carpeta."),
        ref: z.string().optional().describe("Limitar a una carpeta: ruta o id de la carpeta."),
        query: z.string().optional().describe("Texto adicional para filtrar (ej. 'IVA', 'contrato')."),
        size: z.number().int().min(1).max(50).default(25).describe("Número máximo de resultados."),
      },
    },
    async ({ days, driveId, ref, query, size }) => {
      try {
        return ok(await graph.recentChanges({ days, driveId, ref, query, size }));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "list_versions",
    {
      title: "Historial de versiones",
      description: "Lista las versiones de un archivo: cuándo se modificó, quién lo hizo y el tamaño de cada versión.",
      inputSchema: { driveId: z.string().describe(driveDesc), ref: z.string().describe(refDesc) },
    },
    async ({ driveId, ref }) => {
      try {
        return ok(await graph.listVersions(driveId, ref));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "restore_version",
    {
      title: "Restaurar versión anterior",
      description:
        "Restaura una versión anterior de un archivo (versionId de list_versions). SharePoint guarda la versión actual en el historial, así que no se pierde nada. Confirma con el usuario antes de usarla.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        versionId: z.string().describe("Id de la versión a restaurar (de list_versions)."),
      },
    },
    async ({ driveId, ref, versionId }) =>
      write("restore_version", { driveId, ref, versionId }, () => graph.restoreVersion(driveId, ref, versionId))
  );

  server.registerTool(
    "list_permissions",
    {
      title: "Quién tiene acceso",
      description:
        "Muestra quién tiene acceso a un archivo o carpeta: usuarios, grupos y enlaces compartidos (con su alcance: 'anonymous' = cualquiera con el enlace, 'organization' = toda la empresa, 'users' = personas concretas) y si el permiso es heredado.",
      inputSchema: { driveId: z.string().describe(driveDesc), ref: z.string().describe(refDesc) },
    },
    async ({ driveId, ref }) => {
      try {
        return ok(await graph.listPermissions(driveId, ref));
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "revoke_share_link",
    {
      title: "Revocar enlace compartido",
      description:
        "Elimina un enlace compartido de un archivo o carpeta (permissionId de list_permissions). Sólo enlaces, no usuarios ni grupos, y sólo si no es heredado. Confirma con el usuario antes de usarla.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        permissionId: z.string().describe("Id del permiso de tipo enlace (de list_permissions)."),
      },
    },
    async ({ driveId, ref, permissionId }) =>
      write("revoke_share_link", { driveId, ref, permissionId }, () => graph.revokeShareLink(driveId, ref, permissionId))
  );

  server.registerTool(
    "upload_file",
    {
      title: "Subir / crear / sobrescribir archivo",
      description:
        "Crea o sobrescribe un archivo en una carpeta (los de más de 4 MB se suben por partes automáticamente). Para texto pasa `content`; para binarios pasa `content` en base64 y is_base64=true.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        parent: z.string().default("root").describe("Carpeta destino (ruta o id). 'root' para la raíz."),
        name: z.string().describe("Nombre del archivo, ej. 'informe.txt'."),
        content: z.string().describe("Contenido del archivo (texto o base64)."),
        is_base64: z.boolean().default(false).describe("El contenido viene en base64."),
      },
    },
    async ({ driveId, parent, name, content, is_base64 }) =>
      write("upload_file", { driveId, parent, name }, () => graph.uploadFile(driveId, parent, name, content, is_base64))
  );

  server.registerTool(
    "create_folder",
    {
      title: "Crear carpeta",
      description: "Crea una nueva carpeta dentro de otra.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        parent: z.string().default("root").describe("Carpeta padre (ruta o id)."),
        name: z.string().describe("Nombre de la nueva carpeta."),
      },
    },
    async ({ driveId, parent, name }) =>
      write("create_folder", { driveId, parent, name }, () => graph.createFolder(driveId, parent, name))
  );

  server.registerTool(
    "delete_item",
    {
      title: "Eliminar item",
      description: "Envía un archivo o carpeta a la papelera de reciclaje.",
      inputSchema: { driveId: z.string().describe(driveDesc), ref: z.string().describe(refDesc) },
    },
    async ({ driveId, ref }) =>
      write("delete_item", { driveId, ref }, () => graph.deleteItem(driveId, ref))
  );

  server.registerTool(
    "move_item",
    {
      title: "Mover / renombrar item",
      description: "Mueve un item a otra carpeta del mismo drive y, opcionalmente, lo renombra.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        new_parent: z.string().describe("Carpeta destino (ruta o id)."),
        new_name: z.string().optional().describe("Nuevo nombre (opcional)."),
      },
    },
    async ({ driveId, ref, new_parent, new_name }) =>
      write("move_item", { driveId, ref, new_parent, new_name }, () => graph.moveItem(driveId, ref, new_parent, new_name))
  );

  server.registerTool(
    "copy_item",
    {
      title: "Copiar item",
      description: "Copia un archivo o carpeta a otra ubicación (operación asíncrona en Graph).",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        new_parent: z.string().describe("Carpeta destino (ruta o id)."),
        new_name: z.string().optional().describe("Nuevo nombre (opcional)."),
      },
    },
    async ({ driveId, ref, new_parent, new_name }) =>
      write("copy_item", { driveId, ref, new_parent, new_name }, () => graph.copyItem(driveId, ref, new_parent, new_name))
  );

  server.registerTool(
    "create_share_link",
    {
      title: "Crear enlace para compartir",
      description: "Genera un enlace para compartir un archivo o carpeta dentro de la organización.",
      inputSchema: {
        driveId: z.string().describe(driveDesc),
        ref: z.string().describe(refDesc),
        type: z.enum(["view", "edit"]).default("view").describe("Tipo de enlace: 'view' o 'edit'."),
      },
    },
    async ({ driveId, ref, type }) =>
      write("create_share_link", { driveId, ref, type }, () => graph.createShareLink(driveId, ref, type))
  );
}
