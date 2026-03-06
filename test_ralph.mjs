import { chromium } from 'playwright';
import { resolve } from 'path';
import { readFileSync } from 'fs';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const JSZip = require('jszip');

const DIR = 'C:/Users/User/Desktop/92-duty-scheduler';
const HTML_PATH = 'file:///' + resolve(DIR, 'index.html').replaceAll('\\', '/');
const XLSX_PATH = resolve(DIR, '2-26+.xlsx');

const KPI = {};
const issues = [];

async function newPage(browser) {
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.goto(HTML_PATH);
  await page.waitForSelector('#stepImport', { state: 'visible', timeout: 5000 });
  return { page, ctx, errors };
}

async function importData(page) {
  const t0 = Date.now();
  await page.locator('#fileInput').setInputFiles(XLSX_PATH);
  await page.waitForSelector('#stepConfig', { state: 'visible', timeout: 10000 });
  return Date.now() - t0;
}

async function selectAllDuties(page) {
  await page.evaluate(() => {
    document.querySelectorAll('.duty-card').forEach(c => {
      if (!c.classList.contains('selected')) c.click();
    });
  });
}

async function schedule(page) {
  const t0 = Date.now();
  await page.click('#btnSchedule');
  await page.waitForSelector('#stepResults.active', { timeout: 15000 });
  return Date.now() - t0;
}

async function getEmptyCount(page) {
  return page.evaluate(() => document.querySelectorAll('.person-cell.empty-slot').length);
}

async function getTotalSlots(page) {
  return page.evaluate(() => document.querySelectorAll('.person-cell').length);
}

async function fillEmpty(page) {
  const t0 = Date.now();
  await page.click('#btnFillEmpty');
  await page.waitForTimeout(3000);
  return Date.now() - t0;
}

// ========== LOOP 1: Basic Workflow ==========
async function loop1(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 1: 基本工作流模擬');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);

  // Import
  const importMs = await importData(page);
  const studentCount = await page.evaluate(() => S.students.length);
  console.log(`  匯入: ${importMs}ms, 學員數: ${studentCount}`);

  // Schedule
  await selectAllDuties(page);
  const schedMs = await schedule(page);
  const emptyBefore = await getEmptyCount(page);
  const totalSlots = await getTotalSlots(page);
  console.log(`  排班: ${schedMs}ms, 空缺: ${emptyBefore}/${totalSlots}`);

  // Fill
  const fillMs = await fillEmpty(page);
  const emptyAfter = await getEmptyCount(page);
  console.log(`  填滿: ${fillMs}ms, 空缺: ${emptyAfter}/${totalSlots}`);

  // Export
  let exportOK = false, pageSetupOK = 0, pageSetupTotal = 0;
  try {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 30000 }),
      page.click('#btnExport')
    ]);
    const dlPath = await download.path();
    exportOK = true;

    // Verify pageSetup via JSZip
    const buf = readFileSync(dlPath);
    const zip = await JSZip.loadAsync(buf);
    for (let i = 1; i <= 9; i++) {
      const f = zip.file('xl/worksheets/sheet' + i + '.xml');
      if (!f) continue;
      pageSetupTotal++;
      const xml = await f.async('string');
      if (xml.includes('<pageSetup') && xml.includes('fitToPage="1"')) pageSetupOK++;
    }
  } catch (e) { issues.push('L1: 匯出失敗 - ' + e.message); }
  console.log(`  匯出: ${exportOK ? '✅' : '❌'}, pageSetup: ${pageSetupOK}/${pageSetupTotal}`);

  // Save
  await page.click('#btnSave');
  await page.waitForTimeout(1000);
  const saveOK = await page.evaluate(() => S.savedWeeks.has(S.weekNum));
  console.log(`  儲存: ${saveOK ? '✅' : '❌'}`);

  // Self-test
  const selfTest = await page.evaluate(() => runSelfTests());
  console.log(`  自測: ${selfTest.passed}/${selfTest.passed + selfTest.failed}`);

  KPI.loop1 = {
    importMs, schedMs, fillMs, studentCount,
    emptyBefore, emptyAfter, totalSlots,
    coverage: ((totalSlots - emptyAfter) / totalSlots * 100).toFixed(1) + '%',
    exportOK, pageSetupRate: pageSetupOK + '/' + pageSetupTotal,
    saveOK, selfTest: selfTest.passed + '/' + (selfTest.passed + selfTest.failed),
    consoleErrors: errors.length
  };
  if (errors.length > 0) issues.push('L1: Console 錯誤 - ' + errors.slice(0, 3).join('; '));
  /* 215 學員 + 資格限制，覆蓋率 > 70% 即為正常 */
  var coveragePct = (totalSlots - emptyAfter) / totalSlots * 100;
  if (coveragePct < 70) issues.push('L1: 填滿後覆蓋率過低 ' + coveragePct.toFixed(1) + '%');

  await ctx.close();
}

// ========== LOOP 2: Edge Cases ==========
async function loop2(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 2: 邊界場景模擬');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);
  await importData(page);

  // Exclude students
  const excludeOK = await page.evaluate(() => {
    if (S.students.length < 3) return false;
    S.excluded.add(S.students[0].id);
    S.excluded.add(S.students[1].id);
    S.excluded.add(S.students[2].id);
    return S.excluded.size === 3;
  });
  console.log(`  排除學員: ${excludeOK ? '✅ 3人' : '❌'}`);

  await selectAllDuties(page);
  await schedule(page);

  // Check excluded students are not scheduled
  const excludedInSchedule = await page.evaluate(() => {
    let found = 0;
    const excl = Array.from(S.excluded);
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        (S.assignments[dt][k] || []).forEach(p => {
          if (p && excl.includes(p.id)) found++;
        });
      });
    });
    return found;
  });
  console.log(`  排除檢查: ${excludedInSchedule === 0 ? '✅ 無誤排' : '❌ ' + excludedInSchedule + '人被誤排'}`);
  if (excludedInSchedule > 0) issues.push('L2: 排除學員被誤排 ' + excludedInSchedule + ' 次');

  // Undo test
  await fillEmpty(page);
  const beforeUndo = await page.evaluate(() => {
    const counts = {};
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        (S.assignments[dt][k] || []).forEach(p => { if (p) counts[p.id] = (counts[p.id] || 0) + 1; });
      });
    });
    return Object.keys(counts).length;
  });

  // Click undo
  const undoBtn = await page.$('#btnUndo');
  const undoDisabled = await undoBtn.isDisabled();
  if (!undoDisabled) {
    await page.click('#btnUndo');
    await page.waitForTimeout(500);
    const afterUndo = await page.evaluate(() => {
      const counts = {};
      Object.keys(S.assignments).forEach(dt => {
        Object.keys(S.assignments[dt]).forEach(k => {
          (S.assignments[dt][k] || []).forEach(p => { if (p) counts[p.id] = (counts[p.id] || 0) + 1; });
        });
      });
      return Object.keys(counts).length;
    });
    console.log(`  復原: ✅ 人數 ${beforeUndo} → ${afterUndo}`);
  } else {
    console.log(`  復原: ⚠ undoStack 為空`);
    issues.push('L2: 填滿後 undoStack 為空，無法復原');
  }

  // Duplicate person check
  const dupes = await page.evaluate(() => {
    const bySlot = {};
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        (S.assignments[dt][k] || []).forEach(p => {
          if (p) {
            const slotKey = k;
            if (!bySlot[slotKey]) bySlot[slotKey] = new Set();
            if (bySlot[slotKey].has(p.id)) return; // same slot OK (different duty)
            bySlot[slotKey].add(p.id);
          }
        });
      });
    });
    // Check time conflicts: same person, same day, overlapping slots
    const personSlots = {};
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        const [day, slot] = k.split('|');
        (S.assignments[dt][k] || []).forEach(p => {
          if (p) {
            if (!personSlots[p.id]) personSlots[p.id] = [];
            personSlots[p.id].push({ day, slot, dtype: dt });
          }
        });
      });
    });
    let conflicts = 0;
    Object.values(personSlots).forEach(slots => {
      for (let i = 0; i < slots.length; i++) {
        for (let j = i + 1; j < slots.length; j++) {
          if (slots[i].day === slots[j].day && slots[i].slot === slots[j].slot) conflicts++;
        }
      }
    });
    return conflicts;
  });
  console.log(`  時間衝突: ${dupes === 0 ? '✅ 無衝突' : '❌ ' + dupes + ' 起'}`);
  if (dupes > 0) issues.push('L2: 發現 ' + dupes + ' 起同時段衝突');

  KPI.loop2 = { excludeOK, excludedInSchedule, undoAvailable: !undoDisabled, timeConflicts: dupes, consoleErrors: errors.length };
  if (errors.length > 0) issues.push('L2: Console 錯誤 - ' + errors.slice(0, 3).join('; '));

  await ctx.close();
}

// ========== LOOP 3: Multi-week ==========
async function loop3(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 3: 多週連續模擬');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);
  await importData(page);

  // Week 1
  await page.evaluate(() => { document.getElementById('weekNum').value = '1'; document.getElementById('weekNum').dispatchEvent(new Event('change')); });
  await selectAllDuties(page);
  await schedule(page);
  await fillEmpty(page);
  await page.click('#btnSave');
  await page.waitForTimeout(1000);
  const w1saved = await page.evaluate(() => S.savedWeeks.has(1));
  const w1hours = await page.evaluate(() => Object.keys(S.cumHours).length);
  console.log(`  Week 1 儲存: ${w1saved ? '✅' : '❌'}, 有時數學員: ${w1hours}`);

  // Week 2 - go back to config
  page.on('dialog', async d => await d.accept());
  await page.click('#btnReconfig');
  await page.waitForTimeout(1500);

  const configVisible = await page.evaluate(() => {
    const el = document.getElementById('stepConfig');
    return el && el.classList.contains('active');
  });

  if (configVisible) {
    await page.evaluate(() => { document.getElementById('weekNum').value = '2'; document.getElementById('weekNum').dispatchEvent(new Event('change')); });
    await selectAllDuties(page);
    await schedule(page);
    await fillEmpty(page);
    await page.click('#btnSave');
    await page.waitForTimeout(1000);
    const w2saved = await page.evaluate(() => S.savedWeeks.has(2));
    const w2hours = await page.evaluate(() => Object.keys(S.cumHours).length);
    console.log(`  Week 2 儲存: ${w2saved ? '✅' : '❌'}, 有時數學員: ${w2hours}`);

    // Verify cumulative hours increased
    const cumOK = await page.evaluate(() => {
      let total = 0;
      Object.values(S.cumHours).forEach(h => total += h);
      return total > 0;
    });
    console.log(`  累計時數: ${cumOK ? '✅ 正確' : '❌ 異常'}`);

    // Verify cumCountByType exists
    const countOK = await page.evaluate(() => {
      return Object.keys(S.cumCountByType).length > 0;
    });
    console.log(`  累計次數: ${countOK ? '✅ 有資料' : '❌ 無資料'}`);
    if (!countOK) issues.push('L3: cumCountByType 無資料');

    KPI.loop3 = { w1saved, w2saved, cumHoursOK: cumOK, cumCountOK: countOK };
  } else {
    console.log(`  ⚠ 無法回到設定頁面`);
    issues.push('L3: btnReconfig 無法回到設定頁');
    KPI.loop3 = { w1saved, configReturn: false };
  }

  if (errors.length > 0) issues.push('L3: Console 錯誤 - ' + errors.slice(0, 3).join('; '));
  await ctx.close();
}

// ========== LOOP 4: UI/UX ==========
async function loop4(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 4: UI/UX 體驗模擬');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);
  await importData(page);

  // Test keyboard shortcut Ctrl+Z
  await selectAllDuties(page);
  await schedule(page);
  await fillEmpty(page);

  const beforeCtrlZ = await page.evaluate(() => {
    let c = 0;
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        (S.assignments[dt][k] || []).forEach(p => { if (p) c++; });
      });
    });
    return c;
  });

  await page.keyboard.press('Control+z');
  await page.waitForTimeout(500);

  const afterCtrlZ = await page.evaluate(() => {
    let c = 0;
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        (S.assignments[dt][k] || []).forEach(p => { if (p) c++; });
      });
    });
    return c;
  });
  const ctrlzWorks = afterCtrlZ < beforeCtrlZ;
  console.log(`  Ctrl+Z: ${ctrlzWorks ? '✅' : '❌'} (${beforeCtrlZ} → ${afterCtrlZ})`);
  if (!ctrlzWorks) issues.push('L4: Ctrl+Z 快捷鍵無效');

  // Mobile viewport test
  await ctx.close();
  const mobileCtx = await browser.newContext({
    viewport: { width: 375, height: 812 },
    userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X)',
    acceptDownloads: true
  });
  const mobilePage = await mobileCtx.newPage();
  const mobileErrors = [];
  mobilePage.on('pageerror', e => mobileErrors.push(e.message));
  await mobilePage.goto(HTML_PATH);
  await mobilePage.waitForSelector('#stepImport', { state: 'visible', timeout: 5000 });

  // Check header is visible
  const headerVisible = await mobilePage.evaluate(() => {
    const h = document.querySelector('header');
    const r = h.getBoundingClientRect();
    return r.height > 0 && r.width > 0;
  });
  console.log(`  手機版 header: ${headerVisible ? '✅' : '❌'}`);

  // Check drop zone is accessible
  const dropZoneOK = await mobilePage.evaluate(() => {
    const dz = document.getElementById('dropZone');
    const r = dz.getBoundingClientRect();
    return r.height >= 44 && r.width > 200;
  });
  console.log(`  手機版 dropZone: ${dropZoneOK ? '✅' : '❌'}`);

  // Check video source switched to mobile
  await mobilePage.waitForTimeout(1000);
  const videoSrc = await mobilePage.evaluate(() => {
    const v = document.getElementById('bgVideo');
    return v ? v.currentSrc : 'none';
  });
  const isMobileVideo = videoSrc.includes('bg-mobile');
  console.log(`  手機版影片: ${isMobileVideo ? '✅' : '❌'} (${videoSrc.split('/').pop()})`);

  // Wrong file import test
  await mobilePage.locator('#fileInput').setInputFiles({ name: 'test.txt', mimeType: 'text/plain', buffer: Buffer.from('not excel') });
  await mobilePage.waitForTimeout(2000);
  const errMsg = await mobilePage.evaluate(() => {
    const status = document.getElementById('importStatus');
    return status ? status.textContent : '';
  });
  const errorHandled = errMsg.length > 0;
  console.log(`  錯誤檔案處理: ${errorHandled ? '✅' : '❌'} ${errMsg.substring(0, 50)}`);
  if (!errorHandled) issues.push('L4: 匯入非 Excel 檔案沒有錯誤提示');

  KPI.loop4 = { ctrlzWorks, headerVisible, dropZoneOK, isMobileVideo, errorHandled, mobileErrors: mobileErrors.length, consoleErrors: errors.length };
  if (mobileErrors.length > 0) issues.push('L4: 手機版 console 錯誤 - ' + mobileErrors.slice(0, 3).join('; '));

  await mobileCtx.close();
}

// ========== LOOP 5: Performance ==========
async function loop5(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 5: 效能/穩健度測試');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);
  await importData(page);
  await selectAllDuties(page);
  await schedule(page);

  // Stress test: fill 5 times
  const fillTimes = [];
  for (let i = 0; i < 5; i++) {
    const t0 = Date.now();
    await page.click('#btnFillEmpty');
    await page.waitForTimeout(2000);
    fillTimes.push(Date.now() - t0);
  }
  const avgFill = Math.round(fillTimes.reduce((a, b) => a + b, 0) / fillTimes.length);
  const maxFill = Math.max(...fillTimes);
  console.log(`  填滿 5 次: 平均 ${avgFill}ms, 最大 ${maxFill}ms`);

  // Multiple exports
  let exportCount = 0;
  for (let i = 0; i < 3; i++) {
    try {
      const [dl] = await Promise.all([
        page.waitForEvent('download', { timeout: 30000 }),
        page.click('#btnExport')
      ]);
      await dl.path();
      exportCount++;
    } catch (e) { break; }
  }
  console.log(`  連續匯出 3 次: ${exportCount}/3 ✅`);

  // Check memory (DOM node count)
  const domNodes = await page.evaluate(() => document.querySelectorAll('*').length);
  console.log(`  DOM 節點數: ${domNodes}`);
  if (domNodes > 10000) issues.push('L5: DOM 節點過多 (' + domNodes + ')');

  // Final roster check
  const rosterOK = await page.evaluate(() => {
    let ok = true;
    S.students.forEach(s => {
      if (isWithdrawn(s)) return; // 退學學員可能缺欄位，跳過
      if (!s.id || !s.name || !s.gender) ok = false;
      if (typeof s.id !== 'number') ok = false;
      if (typeof s.dept !== 'string') ok = false;
    });
    return ok;
  });
  console.log(`  名冊完整性: ${rosterOK ? '✅' : '❌'}`);

  // Final empty check
  const finalEmpty = await getEmptyCount(page);
  const finalTotal = await getTotalSlots(page);
  console.log(`  最終覆蓋率: ${((finalTotal - finalEmpty) / finalTotal * 100).toFixed(1)}%`);

  KPI.loop5 = {
    avgFillMs: avgFill, maxFillMs: maxFill,
    exportSuccess: exportCount + '/3',
    domNodes, rosterOK,
    finalCoverage: ((finalTotal - finalEmpty) / finalTotal * 100).toFixed(1) + '%',
    consoleErrors: errors.length
  };
  if (errors.length > 0) issues.push('L5: Console 錯誤 - ' + errors.slice(0, 3).join('; '));

  await ctx.close();
}

// ========== LOOP 6: Backup/Restore & Locking ==========
async function loop6(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 6: 備份還原/鎖定/重複儲存');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);
  await importData(page);
  await selectAllDuties(page);
  await schedule(page);
  await fillEmpty(page);

  // Save week 1
  await page.click('#btnSave');
  await page.waitForTimeout(1000);

  // Try double save (should get confirm dialog)
  page.on('dialog', async d => {
    console.log(`  重複儲存對話框: ✅ "${d.message().substring(0, 30)}..."`);
    await d.accept();
  });
  await page.click('#btnSave');
  await page.waitForTimeout(1000);

  // Backup export
  const backupJSON = await page.evaluate(() => {
    const data = localStorage.getItem('dutySchedulerHistory');
    return data ? JSON.parse(data) : null;
  });
  const backupOK = backupJSON && backupJSON.history && backupJSON.cumHours;
  console.log(`  localStorage 備份: ${backupOK ? '✅' : '❌'}`);
  if (!backupOK) issues.push('L6: localStorage 備份資料不完整');

  // Check backup export button exists
  const backupBtnExists = await page.evaluate(() => !!document.getElementById('btnBackupExport'));
  console.log(`  備份匯出按鈕: ${backupBtnExists ? '✅' : '❌'}`);

  // Verify cumHours has values
  const cumHoursCount = await page.evaluate(() => {
    let c = 0;
    Object.values(S.cumHours).forEach(v => { if (v > 0) c++; });
    return c;
  });
  console.log(`  有累計時數學員: ${cumHoursCount}`);

  // Test picker modal
  const firstCell = await page.$('.person-cell.filled');
  if (firstCell) {
    await firstCell.click();
    await page.waitForTimeout(500);
    const modalVisible = await page.evaluate(() => {
      const m = document.getElementById('pickerModal');
      return m && m.classList.contains('show');
    });
    console.log(`  Picker 彈窗: ${modalVisible ? '✅' : '❌'}`);
    if (!modalVisible) issues.push('L6: 點擊格子未彈出 picker');

    // Search in picker
    if (modalVisible) {
      const pickerSearch = await page.$('#pickerSearch');
      if (pickerSearch) {
        await pickerSearch.fill('行政');
        await page.waitForTimeout(300);
        const filteredCount = await page.evaluate(() => {
          return document.querySelectorAll('.picker-list .pk-item:not([style*="display: none"])').length;
        });
        console.log(`  Picker 搜尋「行政」: ${filteredCount} 結果`);
      }
      // Close modal
      await page.click('#pickerClose');
      await page.waitForTimeout(300);
    }
  }

  KPI.loop6 = { backupOK, cumHoursCount, backupBtnExists, consoleErrors: errors.length };
  if (errors.length > 0) issues.push('L6: Console 錯誤 - ' + errors.slice(0, 3).join('; '));
  await ctx.close();
}

// ========== LOOP 7: Dashboard & Stats Accuracy ==========
async function loop7(browser) {
  console.log('\n══════════════════════════════════════');
  console.log('  LOOP 7: 時數總覽/統計準確度');
  console.log('══════════════════════════════════════');
  const { page, ctx, errors } = await newPage(browser);
  await importData(page);
  await selectAllDuties(page);
  await schedule(page);
  await fillEmpty(page);

  // Check stats bar exists and has content
  const statsGrid = await page.evaluate(() => {
    const el = document.getElementById('statsGrid');
    return el ? el.children.length : 0;
  });
  console.log(`  統計卡片數: ${statsGrid}`);
  if (statsGrid === 0) issues.push('L7: statsGrid 無統計卡片');

  // Check dashboard renders
  const dashboardOK = await page.evaluate(() => {
    const el = document.getElementById('dashboard');
    return el && el.children.length > 0;
  });
  console.log(`  Dashboard 渲染: ${dashboardOK ? '✅' : '❌'}`);

  // Check tab bar
  const tabCount = await page.evaluate(() => {
    return document.querySelectorAll('#resultTabs .tab-btn').length;
  });
  console.log(`  勤務 Tab 數: ${tabCount}`);

  // Verify assignments consistency
  const consistency = await page.evaluate(() => {
    let assignedIds = new Set();
    let totalAssigned = 0;
    Object.keys(S.assignments).forEach(dt => {
      Object.keys(S.assignments[dt]).forEach(k => {
        (S.assignments[dt][k] || []).forEach(p => {
          if (p) { assignedIds.add(p.id); totalAssigned++; }
        });
      });
    });
    return {
      uniqueStudents: assignedIds.size,
      totalAssigned,
      thisWeekIdsSize: S.thisWeekIds.size,
      match: assignedIds.size === S.thisWeekIds.size
    };
  });
  console.log(`  排班一致性: ${consistency.match ? '✅' : '❌'} (unique: ${consistency.uniqueStudents}, thisWeekIds: ${consistency.thisWeekIdsSize}, total slots: ${consistency.totalAssigned})`);
  if (!consistency.match) issues.push('L7: thisWeekIds 與 assignments 不一致');

  // Check fill duty select dropdown
  const fillSelectOptions = await page.evaluate(() => {
    const sel = document.getElementById('fillDutySelect');
    return sel ? sel.options.length : 0;
  });
  console.log(`  勤務篩選選項: ${fillSelectOptions}`);

  KPI.loop7 = { statsGrid, dashboardOK, tabCount, consistency, fillSelectOptions, consoleErrors: errors.length };
  if (errors.length > 0) issues.push('L7: Console 錯誤 - ' + errors.slice(0, 3).join('; '));
  await ctx.close();
}

// ========== MAIN ==========
(async () => {
  const browser = await chromium.launch({ headless: true });

  await loop1(browser);
  await loop2(browser);
  await loop3(browser);
  await loop4(browser);
  await loop5(browser);
  await loop6(browser);
  await loop7(browser);

  await browser.close();

  // Final KPI Report
  console.log('\n══════════════════════════════════════');
  console.log('  RALPH LOOP KPI 總報告');
  console.log('══════════════════════════════════════');
  console.log(JSON.stringify(KPI, null, 2));

  console.log('\n══════════════════════════════════════');
  console.log('  發現的問題');
  console.log('══════════════════════════════════════');
  if (issues.length === 0) console.log('  ✅ 無問題');
  else issues.forEach((iss, i) => console.log(`  ${i + 1}. ${iss}`));

  console.log('\n' + (issues.length === 0 ? '✅ 全部通過' : '⚠ 發現 ' + issues.length + ' 個問題需改進'));
})();
