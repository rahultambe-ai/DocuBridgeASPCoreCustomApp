import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// `npm run build` -> output goes to ../wwwroot/dms. The ASP.NET Core portal serves index.html at /Dms (only for
//                    signed-in users) and the JS/CSS assets from /dms/assets/*.
// `npm run dev`   -> optional: React dev server on :5173 with /api and /Account forwarded to the portal on :7127.
//                    Sign in at https://localhost:7127 first; for day-to-day use simply run the portal.
export default defineConfig({
  base: '/dms/',
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': { target: 'https://localhost:7127', secure: false, changeOrigin: true },
      '/Account': { target: 'https://localhost:7127', secure: false, changeOrigin: true }
    }
  },
  build: {
    outDir: '../wwwroot/dms',
    emptyOutDir: true,
    sourcemap: false,
    chunkSizeWarningLimit: 2500
  }
});
