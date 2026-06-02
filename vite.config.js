import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',
      includeAssets: ['icons/*.png', 'icons/*.svg'],
      manifest: false, // we supply our own /public/manifest.json
      workbox: {
        // Only cache static icon assets — NOT html/js/css.
        // Vite gives JS/CSS content-hashed names so browser HTTP cache handles them.
        // NOT caching index.html means users always get the latest build references.
        globPatterns: ['**/*.{ico,png,svg}'],
        // Activate new service worker immediately without waiting for tabs to close.
        // This ensures fresh code is served as soon as a new build is deployed.
        skipWaiting: true,
        clientsClaim: true,
        navigateFallback: '/index.html',
        navigateFallbackDenylist: [/^\/api\//],
        runtimeCaching: [
          {
            // Report metadata: network-first so users see updates within 5s
            urlPattern: /\/api\/reports$/,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'api-reports',
              expiration: { maxEntries: 20, maxAgeSeconds: 60 * 60 * 24 },
              networkTimeoutSeconds: 5,
            },
          },
        ],
      },
    }),
  ],
})
