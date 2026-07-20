// Configuración de PM2 para mantener el servidor MCP siempre levantado.
// Arranca con:  pm2 start ecosystem.config.cjs
// Node 22 carga el .env con --env-file (ahí están CLIENT_ID, GRAPH_REFRESH_TOKEN, MCP_AUTH_TOKEN...).
module.exports = {
  apps: [
    {
      name: "sharepoint-mcp",
      script: "src/server.js",
      cwd: "c:/Users/Laura/Downloads/sharepoint-mcp",
      node_args: "--env-file=.env",
      autorestart: true,
      max_restarts: 20,
      watch: false,
      env: {
        NODE_ENV: "production",
      },
    },
    {
      // Túnel ngrok con DIRECCIÓN FIJA (no cambia al reiniciar). Expone localhost:3000.
      // La URL permanente es https://tinfoil-humvee-croak.ngrok-free.dev
      name: "sharepoint-ngrok",
      script: "C:/Users/Laura/Downloads/sharepoint-mcp/ngrok.exe",
      args: "http 3000 --url=https://tinfoil-humvee-croak.ngrok-free.dev --log=stdout",
      interpreter: "none",
      autorestart: true,
      max_restarts: 20,
      watch: false,
    },
  ],
};
