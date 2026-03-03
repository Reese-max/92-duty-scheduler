import{chromium}from'playwright';

(async()=>{
  const b=await chromium.launch({headless:true});
  const p=await b.newPage();
  await p.goto('file:///C:/Users/User/Desktop/92-duty-scheduler/index.html');
  await p.waitForSelector('#stepImport',{state:'visible',timeout:5000});
  const fi=await p.locator('#fileInput');
  await fi.setInputFiles('C:/Users/User/Desktop/92-duty-scheduler/2-26+.xlsx');
  await p.waitForSelector('#stepConfig',{state:'visible',timeout:10000});
  await p.fill('#weekNum','1');
  await p.evaluate(()=>{document.getElementById('weekNum').dispatchEvent(new Event('change'));});
  // Enable 夜巡
  const cards=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards){if((await c.textContent()).includes('夜巡')){await c.click();break;}}
  await p.waitForTimeout(300);
  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:15000});
  console.log('OK: Schedule generated');

  // Mark all Sunday slots as holiday (per-slot API)
  await p.evaluate(()=>{
    var def=DUTY_DEFS['\u591C\u5DE1'];
    def.slots.forEach(function(slot){toggleSlotHoliday('\u591C\u5DE1','\u65E5',slot,true);});
  });
  await p.waitForTimeout(300);

  // Count holiday cells - should equal slots * pp
  const count=await p.evaluate(()=>document.querySelectorAll('.holiday-cell').length);
  const total=await p.evaluate(()=>{
    var cfg=S.active['\u591C\u5DE1'],def=DUTY_DEFS['\u591C\u5DE1'];
    return def.slots.length*cfg.pp;
  });
  console.log('Holiday cells: '+count+'/'+total);
  console.log(count===total?'PASS: All Sunday slot cells are holiday':'FAIL: mismatch '+count+' vs '+total);

  // Verify S.holidays stores day|slot keys
  const hol=await p.evaluate(()=>{var r={};Object.keys(S.holidays).forEach(function(k){r[k]=Array.from(S.holidays[k]);});return r;});
  console.log('S.holidays:',JSON.stringify(hol));
  const keys=hol['\u591C\u5DE1']||[];
  const allHaveSlot=keys.every(k=>k.includes('|'));
  console.log(allHaveSlot?'PASS: All keys are day|slot format':'FAIL: Some keys missing slot');

  // Cancel all Sunday slots
  await p.evaluate(()=>{
    var def=DUTY_DEFS['\u591C\u5DE1'];
    def.slots.forEach(function(slot){toggleSlotHoliday('\u591C\u5DE1','\u65E5',slot,false);});
  });
  await p.waitForTimeout(200);
  const after=await p.evaluate(()=>document.querySelectorAll('.holiday-cell').length);
  console.log('After cancel: '+after+' cells');
  console.log(after===0?'PASS: Cancel works':'FAIL: '+after+' cells remain');

  // Test partial holiday: mark only first slot as holiday
  await p.evaluate(()=>{
    var def=DUTY_DEFS['\u591C\u5DE1'];
    toggleSlotHoliday('\u591C\u5DE1','\u65E5',def.slots[0],true);
  });
  await p.waitForTimeout(200);
  const partial=await p.evaluate(()=>document.querySelectorAll('.holiday-cell').length);
  const pp=await p.evaluate(()=>S.active['\u591C\u5DE1'].pp);
  console.log('Partial holiday cells: '+partial+'/'+pp);
  console.log(partial===pp?'PASS: Only first slot is holiday':'FAIL: expected '+pp+' got '+partial);

  // Test picker button shows up (per-slot text)
  await p.evaluate(()=>{
    var cells=document.querySelectorAll('.person-cell');
    cells[0].click();
  });
  await p.waitForSelector('#pickerModal.show',{timeout:3000});
  const hRow=await p.evaluate(()=>{
    var el=document.getElementById('pickerHolidayRow');
    return{display:el.style.display,text:el.textContent};
  });
  console.log('Holiday row in picker:',JSON.stringify(hRow));
  console.log(hRow.text.includes('\u653E\u5047')||hRow.text.includes('\u6A19\u8A18')?'PASS: Holiday button in picker':'FAIL: missing');

  // Test export with per-slot holidays
  await p.evaluate(()=>{ closePicker(); });
  await p.waitForTimeout(200);

  // Trigger export and intercept download
  const [download]=await Promise.all([
    p.waitForEvent('download',{timeout:10000}),
    p.click('#btnExport')
  ]);
  const path=await download.path();
  console.log('Export downloaded: '+!!path);
  console.log(path?'PASS: Export works with per-slot holidays':'FAIL: Export failed');

  await b.close();
  console.log('\nDone.');
})();
