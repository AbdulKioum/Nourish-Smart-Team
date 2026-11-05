import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["logo smart team.png"],
      manifest: {
        name: "Nourish Smart Team",
        short_name: "Smart Team",
        description: "Offline-First Poultry & Farm Management App",
        theme_color: "#059669",
        background_color: "#ffffff",
        display: "standalone",
        orientation: "portrait",
        start_url: "/",
        icons: [
          {
            src: "/logo smart team.png",
            sizes: "192x192",
            type: "image/png",
            purpose: "any maskable",
          },
          {
            src: "/logo smart team.png",
            sizes: "512x512",
            type: "image/png",
            purpose: "any maskable",
          },
        ],
      },
      workbox: {
        // Limit barhiye 6 MB (6 * 1024 * 1024 bytes) kore dewa holo
        maximumFileSizeToCacheInBytes: 6 * 1024 * 1024,
        globPatterns: ["**/*.{js,css,html,ico,png,svg,woff2}"],
        cleanupOutdatedCaches: true,
        clientsClaim: true,
        skipWaiting: true,
      },
    }),
  ],
});