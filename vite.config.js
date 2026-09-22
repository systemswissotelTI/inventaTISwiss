import { defineConfig } from "vite";

// Rutas relativas para que funcione en GitHub Pages (/inventaTISwiss/)
export default defineConfig({
  base: "./",
  // ExcelJS (~900 kB) va en un archivo aparte que solo se carga al exportar
  build: { chunkSizeWarningLimit: 1000 }
});
