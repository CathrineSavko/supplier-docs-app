import http from 'node:http';
import { promises as fs, createWriteStream } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import crypto from 'node:crypto';
import { spawn } from 'node:child_process';
import { deflateRawSync } from 'node:zlib';

const here = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(here, 'public');
const dataDir = path.join(here, 'data');
const templatesDir = path.join(dataDir, 'templates');
const jobsDir = path.join(dataDir, 'jobs');
const port = Number(process.env.PORT || 3213);
let workerProcess;
let stopping = false;

const mime = {
  '.css': 'text/css; charset=utf-8', '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon'
};
const crcTable = Uint32Array.from({ length: 256 }, (_, index) => { let value = index; for (let bit = 0; bit < 8; bit++) value = value & 1 ? (value >>> 1) ^ 0xedb88320 : value >>> 1; return value >>> 0; });
function crc32(buffer) { let value = 0xffffffff; for (const byte of buffer) value = (value >>> 8) ^ crcTable[(value ^ byte) & 255]; return (value ^ 0xffffffff) >>> 0; }
function zip(files) {
  const date = new Date(); const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2); const dosDate = ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  let offset = 0; const local = []; const central = [];
  for (const file of files) { const name = Buffer.from(file.name, 'utf8'); const data = file.data; const packed = deflateRawSync(data); const crc = crc32(data); const header = Buffer.alloc(30); header.writeUInt32LE(0x04034b50, 0); header.writeUInt16LE(20, 4); header.writeUInt16LE(0x800, 6); header.writeUInt16LE(8, 8); header.writeUInt16LE(dosTime, 10); header.writeUInt16LE(dosDate, 12); header.writeUInt32LE(crc, 14); header.writeUInt32LE(packed.length, 18); header.writeUInt32LE(data.length, 22); header.writeUInt16LE(name.length, 26); local.push(header, name, packed); const entry = Buffer.alloc(46); entry.writeUInt32LE(0x02014b50, 0); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt16LE(0x800, 8); entry.writeUInt16LE(8, 10); entry.writeUInt16LE(dosTime, 12); entry.writeUInt16LE(dosDate, 14); entry.writeUInt32LE(crc, 16); entry.writeUInt32LE(packed.length, 20); entry.writeUInt32LE(data.length, 24); entry.writeUInt16LE(name.length, 28); entry.writeUInt32LE(offset, 42); central.push(entry, name); offset += header.length + name.length + packed.length; }
  const centralData = Buffer.concat(central); const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10); end.writeUInt32LE(centralData.length, 12); end.writeUInt32LE(offset, 16); return Buffer.concat([...local, centralData, end]);
}

function safeName(value = 'file') {
  return path.basename(value).replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').slice(0, 180) || 'file';
}

async function ensureFolders() {
  await Promise.all([fs.mkdir(templatesDir, { recursive: true }), fs.mkdir(jobsDir, { recursive: true })]);
}

function startWorker() {
  const output = createWriteStream(path.join(dataDir, 'worker.log'), { flags: 'a' });
  const errors = createWriteStream(path.join(dataDir, 'worker-error.log'), { flags: 'a' });
  const powershell = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  workerProcess = spawn(powershell, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(here, 'scripts', 'worker_v3.ps1'), '-Root', here], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
  workerProcess.stdout.pipe(output);
  workerProcess.stderr.pipe(errors);
  workerProcess.on('error', (error) => {
    errors.write(`${new Date().toISOString()} Cannot start worker: ${error.message}\n`);
  });
  workerProcess.on('close', (code) => {
    output.end(); errors.end();
    if (!stopping) setTimeout(startWorker, 2000);
    console.log(`Supplier Docs worker stopped (code ${code}).`);
  });
}

for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => {
    stopping = true;
    workerProcess?.kill();
    process.exit(0);
  });
}

function parseMultipart(buffer, contentType) {
  const matched = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || '');
  if (!matched) throw new Error('Не найдена граница multipart-формы.');
  const boundary = Buffer.from(`--${matched[1] || matched[2]}`);
  const chunks = [];
  let cursor = buffer.indexOf(boundary) + boundary.length + 2;
  while (cursor > boundary.length + 1 && cursor < buffer.length) {
    const end = buffer.indexOf(boundary, cursor);
    if (end < 0) break;
    const part = buffer.subarray(cursor, end - 2);
    const headerEnd = part.indexOf(Buffer.from('\r\n\r\n'));
    if (headerEnd >= 0) {
      const headers = part.subarray(0, headerEnd).toString('utf8');
      const body = part.subarray(headerEnd + 4);
      const disposition = /name="([^"]+)"/i.exec(headers);
      const fileName = /filename="([^"]*)"/i.exec(headers);
      if (disposition) chunks.push({ name: disposition[1], filename: fileName?.[1], data: body });
    }
    cursor = end + boundary.length + 2;
  }
  return chunks;
}

async function bodyOf(request) {
  const data = [];
  let size = 0;
  for await (const part of request) {
    size += part.length;
    if (size > 250 * 1024 * 1024) throw new Error('Размер загрузки превышает 250 МБ.');
    data.push(part);
  }
  return Buffer.concat(data);
}

async function listFiles(root, prefix = '') {
  let entries = [];
  try { entries = await fs.readdir(root, { withFileTypes: true }); } catch { return []; }
  const result = [];
  for (const entry of entries) {
    const relative = path.join(prefix, entry.name);
    const full = path.join(root, entry.name);
    if (entry.isDirectory()) result.push(...await listFiles(full, relative));
    else {
      const info = await fs.stat(full);
      result.push({ path: relative.replaceAll('\\', '/'), size: info.size, updatedAt: info.mtime.toISOString() });
    }
  }
  return result;
}

function send(response, status, payload, type = 'application/json; charset=utf-8') {
  response.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  response.end(typeof payload === 'string' || Buffer.isBuffer(payload) ? payload : JSON.stringify(payload));
}

async function handleTemplates(request, response) {
  if (request.method === 'GET') return send(response, 200, { files: await listFiles(templatesDir) });
  if (request.method !== 'POST') return send(response, 405, { error: 'Метод не поддерживается.' });
  const parts = parseMultipart(await bodyOf(request), request.headers['content-type']);
  const category = safeName(parts.find(p => p.name === 'category')?.data.toString('utf8') || 'other');
  const files = parts.filter(p => p.filename && p.data.length);
  if (!files.length) return send(response, 400, { error: 'Выберите хотя бы один файл.' });
  const target = path.join(templatesDir, category);
  await fs.mkdir(target, { recursive: true });
  const saved = [];
  for (const file of files) {
    const name = safeName(file.filename);
    await fs.writeFile(path.join(target, name), file.data);
    saved.push(`${category}/${name}`);
  }
  return send(response, 201, { saved });
}

async function handleJobs(request, response) {
  const url = new URL(request.url, `http://${request.headers.host}`);
  const pieces = url.pathname.split('/').filter(Boolean);
  if (request.method === 'GET' && pieces.length === 3) {
    const id = safeName(pieces[2]);
    try { return send(response, 200, JSON.parse(await fs.readFile(path.join(jobsDir, id, 'request.json'), 'utf8'))); }
    catch { return send(response, 404, { error: 'Запуск не найден.' }); }
  }
  if (request.method === 'GET' && pieces.length === 5 && pieces[3] === 'output') {
    const id = safeName(pieces[2]); const fileName = safeName(decodeURIComponent(pieces[4]));
    const outputFile = path.join(jobsDir, id, 'output', fileName);
    try { return send(response, 200, await fs.readFile(outputFile), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); }
    catch { return send(response, 404, { error: 'Готовый файл не найден.' }); }
  }
  if (request.method === 'POST' && pieces.length === 4 && pieces[3] === 'corrections') {
    const sourceId = safeName(pieces[2]);
    const source = JSON.parse(await fs.readFile(path.join(jobsDir, sourceId, 'request.json'), 'utf8'));
    const body = JSON.parse((await bodyOf(request)).toString('utf8'));
    const selected = Array.isArray(body.selected) ? body.selected.filter(item => item?.correction) : [];
    if (!selected.length) return send(response, 400, { error: 'Отметьте хотя бы одно исправление.' });
    const id = `${new Date().toISOString().slice(0, 10)}-${crypto.randomUUID().slice(0, 8)}`;
    const jobDir = path.join(jobsDir, id); const inputDir = path.join(jobDir, 'input'); await fs.mkdir(inputDir, { recursive: true });
    for (const name of source.inputFiles) await fs.copyFile(path.join(jobsDir, sourceId, 'input', name), path.join(inputDir, name));
    const requestInfo = { id, taskId: 'check-final-corrections', status: 'queued', progress: 0, stage: 'Ожидает обработчик…', createdAt: new Date().toISOString(), parameters: { sourceJobId: sourceId, selected }, inputFiles: source.inputFiles, outputFiles: [], errors: [] };
    await fs.writeFile(path.join(jobDir, 'request.json'), JSON.stringify(requestInfo, null, 2), 'utf8');
    return send(response, 201, requestInfo);
  }
  if (request.method === 'POST' && pieces.length === 4 && pieces[3] === 'output-zip') {
    const id = safeName(pieces[2]); const body = JSON.parse((await bodyOf(request)).toString('utf8')); const names = Array.isArray(body.files) ? body.files.map(safeName) : [];
    if (!names.length) return send(response, 400, { error: 'Выберите файлы для скачивания.' });
    const files = await Promise.all(names.map(async name => ({ name, data: await fs.readFile(path.join(jobsDir, id, 'output', name)) })));
    const archive = zip(files); response.writeHead(200, { 'Content-Type': 'application/zip', 'Content-Disposition': 'attachment; filename="corrected-files.zip"', 'Cache-Control': 'no-store' }); response.end(archive); return;
  }
  if (request.method !== 'POST') return send(response, 405, { error: 'Метод не поддерживается.' });
  const parts = parseMultipart(await bodyOf(request), request.headers['content-type']);
  const fields = Object.fromEntries(parts.filter(p => !p.filename).map(p => [p.name, p.data.toString('utf8')]));
  const taskId = safeName(fields.taskId || 'unknown');
  const files = parts.filter(p => p.filename && p.data.length);
  if (!files.length) return send(response, 400, { error: 'Добавьте исходные файлы.' });
  const id = `${new Date().toISOString().slice(0, 10)}-${crypto.randomUUID().slice(0, 8)}`;
  const jobDir = path.join(jobsDir, id);
  const inputDir = path.join(jobDir, 'input');
  await fs.mkdir(inputDir, { recursive: true });
  const savedFiles = [];
  for (const file of files) {
    const name = safeName(file.filename);
    await fs.writeFile(path.join(inputDir, name), file.data);
    savedFiles.push(name);
  }
  const requestInfo = { id, taskId, status: 'queued', progress: 0, stage: 'Ожидает обработчик…', createdAt: new Date().toISOString(), parameters: fields, inputFiles: savedFiles, outputFiles: [], errors: [] };
  await fs.writeFile(path.join(jobDir, 'request.json'), JSON.stringify(requestInfo, null, 2), 'utf8');
  return send(response, 201, requestInfo);
}

async function serveStatic(request, response) {
  const url = new URL(request.url, `http://${request.headers.host}`);
  const requested = url.pathname === '/' ? '/index.html' : url.pathname;
  const file = path.resolve(publicDir, `.${requested}`);
  if (!file.startsWith(publicDir)) return send(response, 403, 'Forbidden', 'text/plain');
  try {
    const data = await fs.readFile(file);
    return send(response, 200, data, mime[path.extname(file)] || 'application/octet-stream');
  } catch { return send(response, 404, 'Не найдено', 'text/plain; charset=utf-8'); }
}

await ensureFolders();
startWorker();
http.createServer(async (request, response) => {
  try {
    const pathname = new URL(request.url, `http://${request.headers.host}`).pathname;
    if (pathname === '/api/templates') return await handleTemplates(request, response);
    if (pathname === '/api/jobs' || pathname.startsWith('/api/jobs/')) return await handleJobs(request, response);
    return await serveStatic(request, response);
  } catch (error) {
    console.error(error);
    return send(response, 500, { error: error.message || 'Внутренняя ошибка сервера.' });
  }
}).listen(port, '127.0.0.1', () => {
  console.log(`Supplier Docs is running at http://127.0.0.1:${port}`);
});
