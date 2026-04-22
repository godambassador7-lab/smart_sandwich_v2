import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "/smart_sandwich_v2/",
  plugins: [react()],
});
