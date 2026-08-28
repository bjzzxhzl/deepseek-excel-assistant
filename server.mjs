import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { join, normalize, extname } from 'node:path';

const ROOT = normalize(join(import.meta.dirname));
const PORT = Number(process.env.PORT || 8090);
// 空闲自动退出：任务窗格打开期间每 2 分钟心跳一次；窗格关闭后 15 分钟无任何请求即退出
const IDLE_MS = (Number(process.env.IDLE_MINUTES) || 15) * 60 * 1000;
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.md': 'text/plain; charset=utf-8'
};

let lastActivity = Date.now();

const server = http.createServer(async (req, res) => {
  lastActivity = Date.now();
  try {
    const url = new URL(req.url, 'http://x');
    if (url.pathname === '/ping') {
      res.writeHead(200, { 'Content-Type': 'text/plain' });
      res.end('ok');
      return;
    }
    let p = url.pathname === '/' ? '/taskpane.html' : url.pathname;
    const file = normalize(join(ROOT, p));
    if (!file.startsWith(ROOT + '\\') && file !== ROOT) { res.writeHead(403).end(); return; }
    const data = await readFile(file);
    console.log(`[${new Date().toISOString()}] ${req.method} ${p} -> 200`);
    res.writeHead(200, {
      'Content-Type': MIME[extname(file).toLowerCase()] || 'application/octet-stream',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-store'
    });
    res.end(data);
  } catch (e) {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.url} -> 404`);
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('404 Not Found: ' + req.url);
  }
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`ExcelAI server: http://127.0.0.1:${PORT}/taskpane.html (空闲 ${IDLE_MS / 60000} 分钟自动退出)`);
});

// 空闲检测：自适应间隔（正常每分钟一次；超时较短时更频繁）
const CHECK_MS = Math.min(60 * 1000, Math.max(1000, Math.floor(IDLE_MS / 2)));
setInterval(() => {
  if (Date.now() - lastActivity > IDLE_MS) {
    console.log(`[${new Date().toISOString()}] 长时间无请求，自动退出以节省资源`);
    server.close(() => process.exit(0));
  }
}, CHECK_MS);
