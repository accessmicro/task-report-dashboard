import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const repoBase = "/task-report-dashboard/";

export default defineConfig({
  base: process.env.BASE_PATH || repoBase,
  plugins: [react()],
  resolve: {
    alias: {
      "@": new URL("./src", import.meta.url).pathname
    }
  }
});
