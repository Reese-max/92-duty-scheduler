import{chromium}from'playwright';
import{rename}from'fs/promises';
import path from'path';
(async()=>{
  const b=await chromium.launch({headless:true});
  const p=await b.newPage();
  await p.goto('file:///C:/Users/User/Desktop/92-duty-scheduler/index.html');
  await p.waitForSelector('#stepImport',{state:'visible',timeout:5000});
  await p.locator('#fileInput').setInputFiles('C:/Users/User/Desktop/92-duty-scheduler/2-26+.xlsx');
  await p.waitForSelector('#stepConfig',{state:'visible',timeout:10000});
  await p.fill('#weekNum','1');
  await p.evaluate(()=>{document.getElementById('weekNum').dispatchEvent(new Event('change'));});
  const cards=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards){if(!(await c.evaluate(el=>el.classList.contains('selected'))))await c.click();}
  await p.waitForTimeout(300);
  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:30000});
  await p.click('#btnFillEmpty');
  await p.waitForTimeout(5000);
  const [dl]=await Promise.all([
    p.waitForEvent('download',{timeout:15000}),
    p.click('#btnExport')
  ]);
  const tmp=await dl.path();
  const dest=path.resolve('C:/Users/User/Desktop/92-duty-scheduler/勤務排班_第1週.xlsx');
  await rename(tmp,dest);
  console.log('Exported to: '+dest);
  await b.close();
})();
