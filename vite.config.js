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
        // Cache the Vite-built JS/CSS bundles
        globPatterns: ['**/*.{js,css,html,ico,png,svg}'],
        // Don't try to cache the Railway API — it's dynamic
        navigateFallback: '/index.html',
        navigateFallbackDenylist: [/^\/api\//],
        runtimeCaching: [
          {
            // Cache API report metadata for offline fallback
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
