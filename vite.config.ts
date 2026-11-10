import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  base: "./",
  server: {
    host: true,
  },
  resolve: {
    alias: {
      // এক্সেল বা অন্যান্য লাইব্রেরির জন্য ব্রাউজার স্ট্রিম কম্প্যাটিবিলিটি ফিক্স
      stream: "stream-browserify",
    },
  },
  define: {
    'process.env': {},
  },
  plugins: [
    react(),
    tailwindcss(), 
    
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
        maximumFileSizeToCacheInBytes: 6 * 1024 * 1024,
        globPatterns: ["**/*.{js,css,html,ico,png,svg,woff2}"],
        cleanupOutdatedCaches: true,
        clientsClaim: true,
        skipWaiting: true,
      },
    }),
  ],
});