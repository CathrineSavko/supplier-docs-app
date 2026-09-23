import { promises as fs } from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';

const root = process.argv[2];
if (!root) throw new Error('Application root was not supplied.');
const jobsDir = path.join(root, 'data', 'jobs');
const python = 'C:\\Users\\e.savko\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe';
const powershell = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const scripts = {
  'payment-invoices': ['process_payment_invoices.py'],
  'check-final': ['check_final_documents.py'],
  'standardize': ['standardize_final_documents.py']
};

async function save(file, value) {
  await fs.writeFile(file, JSON.stringify(value, null, 2), 'utf8');
}

function runPython(processor, jobDir, template) {
  return new Promise((resolve, reject) => {
    const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(root, 'scripts', 'run_processor.ps1'), '-Processor', processor, '-JobDir', jobDir];
    if (template) args.push('-Template', template);
    const child = spawn(powershell, args, { windowsHide: true });
    let stdout = '', stderr = '';
    child.stdout.on('data', value => { stdout += value; });
    child.stderr.on('data', value => { stderr += value; });
    child.on('error', reject);
    child.on('close', code => code === 0 ? resolve(stdout) : reject(new Error(stderr.trim() || `Python stopped with code ${code}.`)));
  });
}

async function processJob(folder) {
  const requestPath = path.join(folder, 'request.json');
  let job;
  try { job = JSON.parse(await fs.readFile(requestPath, 'utf8')); } catch { return; }
  if (job.status !== 'queued' || !scripts[job.taskId]) return;
  job.status = 'processing'; job.progress = 5; job.stage = 'Подготавливаю файлы…';
  await save(requestPath, job);
  try {
    const processor = path.join(root, 'scripts', scripts[job.taskId][0]);
    const template = job.taskId === 'payment-invoices' ? path.join(root, 'data', 'templates', 'payment-invoices', '229 WOFENG.xlsx') : null;
    const result = JSON.parse(await runPython(processor, folder, template));
    job.outputFiles = result.outputFiles || [];
    job.notes = result.notes || [];
    if (job.taskId === 'check-final') job.report = result.report || [];
    job.errors = [];
    job.status = 'completed'; job.progress = 100; job.stage = 'Готово';
  } catch (error) {
    job.errors = [error.message || 'Обработка не выполнена.'];
    job.status = 'failed'; job.progress = 100; job.stage = 'Не удалось обработать';
  }
  await save(requestPath, job);
}

let busy = false;
async function tick() {
  if (busy) return;
  busy = true;
  try {
    const entries = await fs.readdir(jobsDir, { withFileTypes: true });
    for (const entry of entries) if (entry.isDirectory()) await processJob(path.join(jobsDir, entry.name));
  } catch (error) {
    console.error(`Worker error: ${error.message}`);
  } finally { busy = false; }
}

await fs.mkdir(jobsDir, { recursive: true });
console.log('Supplier Docs worker is running.');
await tick();
setInterval(tick, 800);
