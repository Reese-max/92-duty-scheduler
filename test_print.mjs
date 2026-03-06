import { chromium } from 'playwright';
import { resolve } from 'path';
import { readFileSync } from 'fs';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const XLSX = require('xlsx-js-style');
const JSZip = require('jszip');

const DIR = 'C:/Users/User/Desktop/92-duty-scheduler';
const HTML_PATH = 'file:///' + resolve(DIR, 'index.html').replaceAll('\\', '/');
const XLSX_PATH = resolve(DIR, '2-26+.xlsx');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  await page.goto(HTML_PATH);
  await page.waitForSelector('#stepImport', { state: 'visible', timeout: 5000 });

  // 匯入
  await page.locator('#fileInput').setInputFiles(XLSX_PATH);
  await page.waitForSelector('#stepConfig', { state: 'visible', timeout: 10000 });
  const cnt = await page.evaluate(() => S.students.length);
  console.log('學員數:', cnt);

  // 選全部勤務
  await page.evaluate(() => {
    document.querySelectorAll('.duty-card').forEach(c => {
      if (!c.classList.contains('selected')) c.click();
    });
  });

  // 排班
  await page.click('#btnSchedule');
  await page.waitForSelector('#stepResults.active', { timeout: 10000 });

  // 填滿
  await page.click('#btnFillEmpty');
  await page.waitForTimeout(2000);

  // 匯出（async export，等待 download 最多 30 秒）
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 30000 }),
    page.click('#btnExport')
  ]);
  const dlPath = await download.path();
  console.log('匯出完成:', await download.suggestedFilename());

  // === 用 xlsx-js-style 檢查 margins（它能讀 margins）===
  const wb = XLSX.readFile(dlPath);
  console.log('\n=== 邊距檢查 (xlsx-js-style) ===');
  wb.SheetNames.forEach((name, i) => {
    const ws = wb.Sheets[name];
    const mg = ws['!margins'] || {};
    const ok = mg.left === 0.25 && mg.right === 0.25 && mg.top === 0.25 && mg.bottom === 0.25 && mg.header === 0.1 && mg.footer === 0.1;
    console.log(`${i + 1}. [${name}] margins: ${ok ? '✅' : '❌ ' + JSON.stringify(mg)}`);
  });

  // === 用 JSZip 直接解析 XML 驗證 pageSetup ===
  console.log('\n=== pageSetup XML 檢查 (JSZip) ===');
  const buf = readFileSync(dlPath);
  const zip = await JSZip.loadAsync(buf);
  let allOK = true;

  for (let i = 0; i < wb.SheetNames.length; i++) {
    const name = wb.SheetNames[i];
    const xmlPath = 'xl/worksheets/sheet' + (i + 1) + '.xml';
    const f = zip.file(xmlPath);
    if (!f) { console.log(`${i + 1}. [${name}] ❌ XML 不存在`); allOK = false; continue; }
    const xml = await f.async('string');

    // 提取 pageSetup
    const psMatch = xml.match(/<pageSetup([^>]*)\/?>/);
    // 提取 fitToPage
    const ftpMatch = xml.match(/fitToPage="(\d)"/);

    if (!psMatch) {
      console.log(`${i + 1}. [${name}] ❌ pageSetup 缺失`);
      allOK = false;
      continue;
    }

    const attrs = psMatch[1];
    const orient = (attrs.match(/orientation="([^"]+)"/) || [])[1] || '?';
    const ftw = (attrs.match(/fitToWidth="(\d+)"/) || [])[1] || '?';
    const fth = (attrs.match(/fitToHeight="(\d+)"/) || [])[1] || '?';
    const paper = (attrs.match(/paperSize="(\d+)"/) || [])[1] || '?';
    const ftp = ftpMatch ? ftpMatch[1] : '?';

    let issues = [];
    if (paper !== '9') issues.push('非A4');
    if (ftw !== '1') issues.push('fitToWidth!=1');
    if (ftp !== '1') issues.push('fitToPage!=1');
    const status = issues.length === 0 ? '✅' : '❌';
    if (issues.length > 0) allOK = false;

    console.log(`${i + 1}. [${name}] ${orient} | paper:${paper} | fit:${ftw}x${fth} | fitToPage:${ftp} | ${status} ${issues.join(', ')}`);
  }

  // 列印範圍
  if (wb.Workbook && wb.Workbook.Names) {
    console.log('\n=== 列印範圍 ===');
    wb.Workbook.Names.forEach(n => {
      if (n.Name === '_xlnm.Print_Area') console.log(`  Sheet[${n.Sheet}]: ${n.Ref}`);
    });
  }

  console.log('\n' + (allOK ? '✅ 全部通過' : '❌ 有設定需修正'));
  await browser.close();
})();
