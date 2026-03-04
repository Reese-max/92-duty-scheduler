import{chromium}from'playwright';
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
  for(const c of cards){const t=await c.textContent();if(t.includes('值班')||t.includes('夜巡'))await c.click();}
  await p.waitForTimeout(300);
  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:15000});
  console.log('OK: Generated');

  // Real overlap: assign student to 值班 一|2200-0000, check 夜巡 一|2300-0100
  const r=await p.evaluate(()=>{
    var free=S.students.find(function(s){return s.gender==='\u7537';});
    var fakeKey='\u4E00|2200-0000';
    if(!S.assignments['\u503C\u73ED'])S.assignments['\u503C\u73ED']={};
    S.assignments['\u503C\u73ED'][fakeKey]=[{id:free.id,name:free.name}];
    S.thisWeekIds.add(free.id);

    var avail=getAvailable('\u591C\u5DE1','\u4E00','2300-0100',new Set());
    var inAvail=avail.some(function(a){return a.id===free.id;});
    var all=getAllStudentsForPicker('\u591C\u5DE1','\u4E00','2300-0100',new Set(),null);
    var entry=all.find(function(a){return a.id===free.id;});

    delete S.assignments['\u503C\u73ED'][fakeKey];
    S.thisWeekIds.delete(free.id);
    return{
      student:free.id+' '+free.name,
      inAvailable:inAvail,
      hasConflict:entry?entry.reasons.includes('conflict'):false,
      reasons:entry?entry.reasons:[]
    };
  });
  console.log('Real overlap (值班 2200-0000 vs 夜巡 2300-0100):');
  console.log('  inAvailable='+r.inAvailable+', hasConflict='+r.hasConflict+', reasons='+JSON.stringify(r.reasons));
  console.log(!r.inAvailable&&r.hasConflict?'PASS: Overlap correctly blocked':'FAIL: Overlap not detected');

  // Non-overlap: assign to 值班 一|0600-0800, check 夜巡 一|2300-0100
  const r2=await p.evaluate(()=>{
    var free=S.students.find(function(s){return s.gender==='\u7537';});
    var fakeKey='\u4E00|0600-0800';
    if(!S.assignments['\u503C\u73ED'])S.assignments['\u503C\u73ED']={};
    S.assignments['\u503C\u73ED'][fakeKey]=[{id:free.id,name:free.name}];
    S.thisWeekIds.add(free.id);

    var avail=getAvailable('\u591C\u5DE1','\u4E00','2300-0100',new Set());
    var inAvail=avail.some(function(a){return a.id===free.id;});
    var all=getAllStudentsForPicker('\u591C\u5DE1','\u4E00','2300-0100',new Set(),null);
    var entry=all.find(function(a){return a.id===free.id;});

    delete S.assignments['\u503C\u73ED'][fakeKey];
    S.thisWeekIds.delete(free.id);
    return{
      student:free.id+' '+free.name,
      inAvailable:inAvail,
      hasConflict:entry?entry.reasons.includes('conflict'):false,
      reasons:entry?entry.reasons:[]
    };
  });
  console.log('\nNon-overlap (值班 0600-0800 vs 夜巡 2300-0100):');
  console.log('  inAvailable='+r2.inAvailable+', hasConflict='+r2.hasConflict);
  console.log(!r2.hasConflict?'PASS: No false conflict':'FAIL: False conflict detected');

  await b.close();
  console.log('\nDone.');
})();
