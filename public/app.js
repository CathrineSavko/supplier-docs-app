const taskDefinitions = [
  { id:'check-final', number:'01', title:'Проверка финальных INV и PL', short:'Отчёт по 11 правилам до любого исправления.', description:'Проверка товарных строк INV и PL с отчётом. Исправления выполняются только после явного подтверждения.', files:'Загрузите один или несколько исходных .xlsx.', fields:[], checks:['отчёт содержит лист, ячейку/строку, проблему и точное предлагаемое исправление','до подтверждения исходные файлы и копии не изменяются','повторный запуск исправлений сохранит только согласованные изменения'] },
  { id:'standardize', number:'02', title:'Единый формат INV и PL', short:'Шрифт, заливка, имя файла и проверка готовых копий.', description:'Приведение товарных таблиц INV и PL к формату поставщика без изменения параметров просмотра и печати.', files:'Загрузите исходные .xlsx для приведения к формату.', fields:[], checks:['Arial Cyr 8 pt bold в товарных таблицах','заливка удалена, а не заменена белой','имя файла сформировано по номеру контейнера'] },
  { id:'commercial-offer', number:'03', title:'Commercial offer', short:'Готово: предложение с товарами и фото.', description:'Создание предложения из ПкЦБ по выбранному поставщику с названием, ценой и фотографией каждой позиции.', files:'Загрузите текущий ПкЦБ .xls. Шаблоны Wofeng уже сохранены. Для GUMI добавьте пустой шаблон в «Шаблоны».', fields:[['pcNumber','ПкЦБ','text','Например, 00612558'],['supplier','Поставщик','select','Wofeng|GUMI'],['documentDate','Дата документа','date','']], checks:['все позиции и итоговая сумма сверены с ПкЦБ','фото пропорционально вписаны в границы PICTURES','печати и форматирование шаблона сохранены'] },
  { id:'product-descriptions', number:'04', title:'Описания товаров Word', short:'Готово: описание с товарами и фото.', description:'Создание Word-описания по шаблону поставщика, с точным порядком товаров и безопасным сопоставлением фото.', files:'Загрузите текущий ПкЦБ .xls. Шаблоны Wofeng и GUMI уже сохранены.', fields:[['pcNumber','ПкЦБ','text','Например, 00612558'],['supplier','Поставщик','select','Wofeng|GUMI'],['documentDate','Дата документа','date',''],['containers','Номера контейнеров','text','Например, 239 240 241']], checks:['нет позиций, которых нет в ПкЦБ','у каждого товара есть описание и точное фото из каталога','печати и исходное оформление шаблона сохранены'] },
  { id:'translate-letters', number:'05', title:'Информационные письма с переводом', short:'Готово: отдельные письма с переводом.', description:'Разделение актуального файла поставщика и добавление второй страницы по сохранённому образцу перевода. Номера контейнеров приложение определяет из инвойсов каждого письма.', files:'Загрузите один актуальный многостраничный документ поставщика .docx. Образец перевода Wofeng уже сохранён.', fields:[['translationDate','Дата перевода','text','Например, 14 сентября 2026 или 21.09.2026']], checks:['первая страница каждого письма остаётся без изменений','во второй странице установлены дата, сумма и номера инвойсов из данного письма','имя каждого файла формируется из контейнеров, указанных в инвойсах письма'] },
  { id:'payment-invoices', number:'06', title:'Инвойсы на оплату', short:'Готово: подготовка INV по образцу 229 WOFENG.', description:'Обработка инвойсов: лист INV, блок после Total, красная печать и заданные объединения/размеры.', files:'Загрузите исходные .xlsx. Образец 229 WOFENG хранится в «Шаблоны».', fields:[], checks:['в готовом файле только INV','нет данных или объектов за A1:I33','печать не пересекает таблицу и находится в D20:G29'] }
];
const categories = [
  ['commercial','Commercial offer: пустые шаблоны и каталоги'],
  ['descriptions','Описания товаров: Word-шаблоны и каталоги'],
  ['translations','Перевод: образцы второй страницы'],
  ['payment-invoices','Инвойсы на оплату: образец 229 WOFENG'],
  ['other','Другой постоянный файл']
];
const byId = id => document.getElementById(id);
let currentTask;

function taskCard(task) { return `<button class="task" data-task="${task.id}"><span class="task-number">${task.number}</span><h3>${task.title}</h3><p>${task.short}</p><small>Открыть задачу →</small></button>`; }
function fieldMarkup([name,label,type,placeholder]) {
  if (type === 'select') return `<label>${label}<select name="${name}" required><option value="" selected disabled>Выберите</option>${placeholder.split('|').map(v=>`<option>${v}</option>`).join('')}</select></label>`;
  if (type === 'textarea') return `<label>${label}<textarea name="${name}" placeholder="${placeholder}" required></textarea></label>`;
  return `<label>${label}<input name="${name}" type="${type}" placeholder="${placeholder}" required></label>`;
}
function openTask(task) {
  currentTask = task;
  byId('workspaceNumber').textContent = `Задача ${task.number}`;
  byId('workspaceTitle').textContent = task.title;
  byId('workspaceDescription').textContent = task.description;
  byId('fields').innerHTML = task.fields.map(fieldMarkup).join('') || '<p class="muted">Параметры не требуются для этого запуска.</p>';
  byId('fileHint').textContent = task.files;
  byId('checks').innerHTML = task.checks.map(item => `<li>${item}</li>`).join('');
  byId('jobResult').textContent = '';
  byId('jobResult').className = 'result';
  byId('downloadAllButton').classList.add('hidden');
  delete byId('downloadAllButton').dataset.job;
  byId('jobForm').reset(); byId('selectedFiles').textContent = '';
  byId('workspace').classList.remove('hidden'); byId('workspace').scrollIntoView({ behavior:'smooth', block:'start' });
}
function formatBytes(bytes) { return bytes < 1024 * 1024 ? `${Math.ceil(bytes/1024)} КБ` : `${(bytes/1024/1024).toFixed(1)} МБ`; }
function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[char])); }
function stageLabel(stage) { return ({'Preparing files...':'Подготавливаю файлы…','Processing failed':'Не удалось обработать','Done':'Готово'})[stage] || stage || 'Выполняется…'; }
function errorLabel(errors) {
  const text = String((errors || [])[0] || 'Обработка не выполнена.');
  const mapping = /No container mapping for payment amount\s+([\d.]+)\s+on page\s+(\d+)/i.exec(text);
  if (mapping) return `Для письма на странице ${mapping[2]} с суммой ${mapping[1]} не заданы номера контейнеров. Добавьте отдельную строку сопоставления и повторите запуск.`;
  if (/Each mapping must have the form/i.test(text)) return 'Проверьте формат сопоставления: «1 страница - 234 235 236» или «519917 → 239 240 241». Для каждого письма нужна отдельная строка.';
  if (/Mapping does not match any supplier letter/i.test(text)) return 'Сопоставление не соответствует ни одному письму. Проверьте номер страницы или сумму.';
  return text.replace(/^python\.exe\s*:\s*/i, '').split(/\r?\n/)[0];
}
function progressMarkup(job) {
  const percent = Math.max(0, Math.min(100, Number(job.progress) || 0));
  return `<div class="progress-status"><div class="progress-label"><span>${escapeHtml(stageLabel(job.stage))}</span><strong>${percent}%</strong></div><div class="progress-track" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${percent}"><span style="width:${percent}%"></span></div></div>`;
}
function reportMarkup(report) {
  return report.filter(Boolean).map(file => {
    if (!file.findings?.length) return `<article class="report-file"><h3>${escapeHtml(file.file)}</h3><p>Ошибок по пунктам 1–11 не обнаружено.</p><p class="muted">${escapeHtml(file.grammarNote || '')}</p></article>`;
    const items = file.findings.map(issue => {
      const selectable = issue.correction ? `<input type="checkbox" class="issue-check" data-issue="${encodeURIComponent(JSON.stringify({...issue, file:file.file}))}">` : '<span class="issue-empty"></span>';
      return `<li class="report-issue"><label>${selectable}<span><strong>[${escapeHtml(issue.point)}; ${escapeHtml(issue.sheet)}, ${escapeHtml(issue.location)}]</strong> ${escapeHtml(issue.found)}<br><span>Предлагаемое исправление: ${escapeHtml(issue.proposal)}</span></span></label></li>`;
    }).join('');
    return `<article class="report-file"><h3>${escapeHtml(file.file)}</h3><ol>${items}</ol><p class="muted">${escapeHtml(file.grammarNote || '')}</p></article>`;
  }).join('');
}
async function followJob(id) {
  const result = byId('jobResult');
  const response = await fetch(`/api/jobs/${encodeURIComponent(id)}`);
  if (!response.ok) return;
  const job = await response.json();
  const downloadAllButton = byId('downloadAllButton');
  const canDownloadAll = job.status === 'completed' && job.outputFiles?.length && (job.taskId === 'standardize' || job.outputFiles.length > 1);
  downloadAllButton.classList.toggle('hidden', !canDownloadAll);
  if (canDownloadAll) downloadAllButton.dataset.job = job.id;
  else delete downloadAllButton.dataset.job;
  if (job.status === 'queued' || job.status === 'processing') { result.className = 'result progress-result'; result.innerHTML = progressMarkup(job); setTimeout(() => followJob(id), 900); return; }
  if (job.status === 'completed') {
    const files = (job.outputFiles || []).map(file => `<a class="download" href="/api/jobs/${encodeURIComponent(id)}/output/${encodeURIComponent(file)}" download>${escapeHtml(file)}</a>`).join('');
    const report = job.report?.some(Boolean) ? `<div class="report"><p>Проверка завершена. Исходные файлы не изменялись.</p><section class="correction-form" data-job="${escapeHtml(job.id)}">${reportMarkup(job.report)}<div class="report-actions"><button class="button accent" type="button" data-correct>Исправить выбранное</button><button class="button secondary" type="button" data-repeat>Подготовить запуск</button><button class="button secondary" type="button" data-download disabled>Получить файлы</button></div><p class="correction-status muted"></p><div class="correction-files"></div></section></div>` : '';
    result.className = 'result progress-result'; result.innerHTML = `${progressMarkup(job)}${report || `<div class="completed-output">Готово. ${files}</div>`}`; return;
  }
  result.className = 'result error progress-result'; result.innerHTML = `${progressMarkup(job)}${escapeHtml(errorLabel(job.errors))}`;
}
async function refreshTemplates() {
  const response = await fetch('/api/templates'); const { files } = await response.json();
  byId('templateList').innerHTML = files.length ? files.map(file => `<div class="template-item"><span>${file.path}</span><small>${formatBytes(file.size)}</small></div>`).join('') : '<p class="muted">Пока нет сохранённых шаблонов.</p>';
}

byId('taskGrid').innerHTML = taskDefinitions.map(taskCard).join('');
document.addEventListener('click', event => { const card = event.target.closest('[data-task]'); if (card) openTask(taskDefinitions.find(t => t.id === card.dataset.task)); });
byId('backButton').addEventListener('click', () => byId('workspace').classList.add('hidden'));
byId('sourceFiles').addEventListener('change', event => { const files = [...event.target.files]; byId('selectedFiles').textContent = files.length ? files.map(f => `${f.name} (${formatBytes(f.size)})`).join(' · ') : ''; });
byId('templatesButton').addEventListener('click', async () => { await refreshTemplates(); byId('templatesDialog').showModal(); });
byId('closeTemplates').addEventListener('click', () => byId('templatesDialog').close());
byId('templateCategory').innerHTML = categories.map(([id,label]) => `<option value="${id}">${label}</option>`).join('');
byId('templateForm').addEventListener('submit', async event => {
  event.preventDefault(); const form = new FormData(event.currentTarget); const result = byId('templateResult'); result.className='result'; result.textContent='Сохраняю…';
  try { const response = await fetch('/api/templates', { method:'POST', body:form }); const data = await response.json(); if (!response.ok) throw new Error(data.error); result.textContent=`Сохранено файлов: ${data.saved.length}.`; event.currentTarget.reset(); await refreshTemplates(); } catch(error) { result.textContent=error.message; result.className='result error'; }
});
byId('jobForm').addEventListener('submit', async event => {
  event.preventDefault(); const files = byId('sourceFiles').files; const result = byId('jobResult');
  if (!files.length) { result.textContent='Добавьте исходные файлы.'; result.className='result error'; return; }
  const form = new FormData(event.currentTarget); form.append('taskId', currentTask.id); byId('downloadAllButton').classList.add('hidden'); result.innerHTML=progressMarkup({progress:0,stage:'Сохраняю локальный запуск…'}); result.className='result progress-result';
  try { const response = await fetch('/api/jobs', { method:'POST', body:form }); const data = await response.json(); if (!response.ok) throw new Error(data.error); sessionStorage.setItem('supplierDocsLastJob', data.id); history.replaceState(null, '', `?job=${encodeURIComponent(data.id)}`); followJob(data.id); } catch(error) { result.textContent=error.message; result.className='result error'; }
});
document.addEventListener('click', async event => {
  const correct = event.target.closest('[data-correct]'); if (!correct) return;
  const form = correct.closest('.correction-form'); const selected = [...form.querySelectorAll('.issue-check:checked')].map(input => JSON.parse(decodeURIComponent(input.dataset.issue)));
  const status = form.querySelector('.correction-status');
  if (!selected.length) { status.textContent = 'Отметьте исправления галочками.'; return; }
  status.textContent = 'Создаю исправленные копии…';
  try { const response = await fetch(`/api/jobs/${encodeURIComponent(form.dataset.job)}/corrections`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({selected})}); const job = await response.json(); if (!response.ok) throw new Error(job.error); status.innerHTML = progressMarkup(job); followCorrection(job.id, form); } catch(error) { status.textContent = error.message; }
});
async function followCorrection(id, form) {
  const response = await fetch(`/api/jobs/${encodeURIComponent(id)}`); if (!response.ok) return;
  const job = await response.json(); const status = form.querySelector('.correction-status'); const files = form.querySelector('.correction-files'); const download = form.querySelector('[data-download]');
  if (job.status === 'queued' || job.status === 'processing') { status.innerHTML = progressMarkup(job); setTimeout(() => followCorrection(id, form), 900); return; }
  if (job.status === 'completed') {
    const links = (job.outputFiles || []).map(file => `<li><label><input class="download-check" type="checkbox" data-name="${escapeHtml(file)}"><a class="download" href="/api/jobs/${encodeURIComponent(id)}/output/${encodeURIComponent(file)}" download>${escapeHtml(file)}</a><strong class="fixed-label">ИСПРАВЛЕН</strong></label></li>`).join('');
    form.querySelectorAll('.issue-check').forEach(input => { input.checked = false; });
    status.textContent = 'Готово. Отметьте исправленные файлы для скачивания.'; files.innerHTML = `<ul>${links}</ul>`; download.disabled = true;
    files.onchange = () => { download.disabled = !files.querySelector('.download-check:checked'); };
    download.onclick = async () => { const chosen = [...files.querySelectorAll('.download-check:checked')].map(input => input.dataset.name); try { const response = await fetch(`/api/jobs/${encodeURIComponent(id)}/output-zip`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({files:chosen})}); if (!response.ok) throw new Error('Не удалось подготовить архив.'); const blob = await response.blob(); const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = 'Исправленные файлы.zip'; link.click(); URL.revokeObjectURL(link.href); } catch(error) { status.textContent = error.message; } }; return;
  }
  status.textContent = (job.errors || ['Не удалось создать исправленные файлы.']).join(' ');
}
async function downloadArchive(jobId, names, fileName) {
  const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/output-zip`, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({files:names}) });
  if (!response.ok) throw new Error('Не удалось подготовить архив.');
  const url = URL.createObjectURL(await response.blob()); const link = document.createElement('a');
  link.href = url; link.download = fileName; document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
document.addEventListener('click', async event => {
  const button = event.target.closest('[data-download-all]'); if (!button) return;
  const original = button.textContent; button.disabled = true; button.textContent = 'Готовлю архив…';
  try {
    const response = await fetch(`/api/jobs/${encodeURIComponent(button.dataset.job)}`); const job = await response.json();
    if (!response.ok || !job.outputFiles?.length) throw new Error('Готовые файлы не найдены.');
    await downloadArchive(job.id, job.outputFiles, 'Готовые файлы.zip');
  } catch (error) { alert(error.message); }
  finally { button.disabled = false; button.textContent = original; }
});
document.addEventListener('click', event => { if (event.target.closest('[data-repeat]')) byId('jobForm').requestSubmit(); });
refreshTemplates();
const savedJobId = new URLSearchParams(location.search).get('job') || sessionStorage.getItem('supplierDocsLastJob');
if (savedJobId) {
  fetch(`/api/jobs/${encodeURIComponent(savedJobId)}`).then(response => response.ok ? response.json() : null).then(job => {
    const task = job && taskDefinitions.find(item => item.id === job.taskId);
    if (task) { openTask(task); followJob(savedJobId); }
  }).catch(() => {});
}
