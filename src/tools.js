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

export function registerTools(server) {
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
