import fs from 'node:fs/promises';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');

const root = path.resolve('.');
const workDir = path.join(root, '_dashboard_work');
const reportDir = path.join(workDir, 'report');
const outputDir = path.join(root, 'output', 'pdf');
const url = process.env.DASHBOARD_URL || 'http://127.0.0.1:8765/dashboard_dre.html';
const python = process.env.CODEX_PYTHON || 'C:\\Users\\PauloMendonça\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe';
const chromePath = process.env.CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

const views = [
  { key: 'dashboard', title: 'Dashboard Executivo' },
  { key: 'dre', title: 'DRE Consolidado' },
  { key: 'evolution', title: 'Análise de Evolução' },
  { key: 'indicators', title: 'Indicadores' },
];

async function loadNotes() {
  try {
    const raw = await fs.readFile(path.join(workDir, 'dashboard_notes.json'), 'utf8');
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

function notesForView(notes, viewKey) {
  return Object.entries(notes)
    .filter(([key, value]) => key.startsWith(`view-${viewKey}|`) && String(value?.note || '').trim())
    .map(([key, value]) => ({
      key,
      title: String(value.title || 'Nota explicativa'),
      note: String(value.note || '').trim(),
      updatedAt: value.updatedAt || '',
    }));
}

async function waitForDashboard(page) {
  await page.goto(`${url}?report=${Date.now()}`, { waitUntil: 'networkidle' });
  await page.waitForSelector('[data-view="dashboard"]', { timeout: 30000 });
  await page.waitForFunction(() => {
    const title = document.querySelector('#dashboardTitle')?.textContent || '';
    return title && !title.includes('Carregando') && !title.includes('Erro');
  }, null, { timeout: 30000 });
}

async function captureView(page, view) {
  await page.evaluate(() => document.querySelector('#reportCaptureStyle')?.remove());
  await page.click(`[data-view="${view.key}"]`);
  await page.waitForTimeout(550);
  await page.evaluate(() => {
    const style = document.createElement('style');
    style.id = 'reportCaptureStyle';
    style.textContent = `
      .app { display: block !important; }
      .sidebar { display: none !important; }
      .main { padding: 12px 16px !important; overflow: hidden !important; }
      body { background: #eef2f8 !important; }
      .topbar { margin-bottom: 8px !important; }
      .chart-tooltip { display: none !important; }
    `;
    document.head.appendChild(style);
    document.querySelectorAll('.chart-tooltip').forEach(tooltip => {
      tooltip.classList.remove('open');
      tooltip.style.display = 'none';
    });
    document.querySelectorAll('.note-card-btn, .zoom-card-btn').forEach(button => {
      button.style.visibility = 'hidden';
    });
    document.querySelectorAll('.dre-table-wrap').forEach(el => {
      el.scrollTop = 0;
      el.scrollLeft = 0;
    });
    window.scrollTo(0, 0);
  });
  const file = path.join(reportDir, `${view.key}.png`);
  await page.screenshot({ path: file, fullPage: false });
  return { ...view, image: file };
}

async function main() {
  await fs.mkdir(reportDir, { recursive: true });
  await fs.mkdir(outputDir, { recursive: true });
  const notes = await loadNotes();

  const browser = await chromium.launch({ headless: true, executablePath: chromePath });
  const page = await browser.newPage({
    viewport: { width: 1920, height: 1080 },
    deviceScaleFactor: 2,
  });

  await waitForDashboard(page);
  const captured = [];
  for (const view of views) {
    const capturedView = await captureView(page, view);
    captured.push({ ...capturedView, notes: notesForView(notes, view.key) });
  }
  await browser.close();

  const manifest = path.join(reportDir, 'manifest.json');
  await fs.writeFile(manifest, JSON.stringify({
    generatedAt: new Date().toISOString(),
    sourceUrl: url,
    outputDir,
    pages: captured,
  }, null, 2), 'utf8');

  const script = path.join(workDir, 'make_board_report_pdf.py');
  const result = spawnSync(python, [script, manifest], {
    cwd: root,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  if (result.status !== 0) {
    console.error(result.stdout);
    console.error(result.stderr);
    process.exit(result.status || 1);
  }
  process.stdout.write(result.stdout);
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
