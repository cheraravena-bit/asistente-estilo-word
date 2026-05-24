const fs = require("fs");
const https = require("https");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const root = __dirname;
const port = 3000;
const iconPng = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lG6lXwAAAABJRU5ErkJggg==",
  "base64"
);

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".png": "image/png"
};

devCerts.getHttpsServerOptions().then((options) => {
  https
    .createServer(options, (request, response) => {
      const url = new URL(request.url, `https://localhost:${port}`);

      if (url.pathname.startsWith("/assets/icon-")) {
        response.writeHead(200, { "Content-Type": "image/png" });
        response.end(iconPng);
        return;
      }

      const filePath = resolveFilePath(url.pathname);
      if (!filePath) {
        response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
        response.end("Not found");
        return;
      }

      fs.readFile(filePath, (error, data) => {
        if (error) {
          response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
          response.end("Server error");
          return;
        }

        response.writeHead(200, {
          "Cache-Control": "no-store",
          "Content-Type": contentTypes[path.extname(filePath)] || "application/octet-stream"
        });
        response.end(data);
      });
    })
    .listen(port, () => {
      console.log(`Asistente de Estilo disponible en https://localhost:${port}`);
    });
});

function resolveFilePath(urlPath) {
  const cleanPath = urlPath === "/" ? "/src/taskpane/taskpane.html" : urlPath;
  const filePath = path.normalize(path.join(root, cleanPath));

  if (!filePath.startsWith(root)) {
    return null;
  }

  if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
    return null;
  }

  return filePath;
}
