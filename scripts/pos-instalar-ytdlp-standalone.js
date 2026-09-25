'use strict';

// youtube-dl-exec baixa por padrão o "yt-dlp" que é um script Python (shebang
// #!/usr/bin/env python3) — o ambiente Nixpacks do Railway pra esse serviço é só Node,
// sem Python instalado. yt-dlp também publica um binário "standalone" (yt-dlp_linux,
// PyInstaller, sem depender de nada instalado no sistema) — troca o arquivo já baixado
// por essa versão depois que o npm install/postinstall padrão terminar.
const https = require('https');
const fs = require('fs');
const path = require('path');

const DEST = path.join(__dirname, '..', 'node_modules', 'youtube-dl-exec', 'bin', 'yt-dlp');
const URL_STANDALONE_LINUX = 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp_linux';

function baixar(url, destino) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { 'User-Agent': 'node' } }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        return baixar(res.headers.location, destino).then(resolve, reject);
      }
      if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode} ao baixar yt-dlp standalone`));
      const arquivo = fs.createWriteStream(destino);
      res.pipe(arquivo);
      arquivo.on('finish', () => arquivo.close(() => resolve()));
      arquivo.on('error', reject);
    }).on('error', reject);
  });
}

(async () => {
  if (!fs.existsSync(path.dirname(DEST))) return; // youtube-dl-exec não está instalado (dependência removida)
  try {
    await baixar(URL_STANDALONE_LINUX, DEST);
    fs.chmodSync(DEST, 0o755);
    console.log('[postinstall] yt-dlp standalone (sem depender de python3) instalado com sucesso.');
  } catch (e) {
    console.error('[postinstall] falha ao trocar yt-dlp pela versão standalone:', e.message);
  }
})();
