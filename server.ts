import 'dotenv/config';
import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { createApiApp } from './server/app.ts';
import { ensureDbReady } from './server/db.ts';

async function startServer() {
  await ensureDbReady();
  const app = createApiApp();

  const distDir = path.resolve(process.cwd(), 'dist');
  const isProd = process.env.NODE_ENV === 'production' && fs.existsSync(distDir);
  const PORT = isProd ? Number(process.env.PORT) || 3000 : 3000;

  if (!isProd) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  const server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`PartnerDesk server running on http://0.0.0.0:${PORT}`);
  });

  server.on('error', (err: NodeJS.ErrnoException) => {
    if (err.code === 'EADDRINUSE' && PORT !== 3000) {
      console.warn(`Port ${PORT} is already in use; retrying on port 3000.`);
      app.listen(3000, '0.0.0.0', () => {
        console.log('PartnerDesk server running on http://0.0.0.0:3000');
      });
      return;
    }
    throw err;
  });
}

startServer().catch((err) => {
  console.error('Failed to start PartnerDesk server:', err);
  process.exit(1);
});
