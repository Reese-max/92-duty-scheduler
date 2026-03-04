import{chromium}from'playwright';

(async()=>{
  const b=await chromium.launch({headless:true});
  const p=await b.newPage();
  p.on('console',m=>{if(m.type()==='error')console.log('PAGE ERR:',m.text());});

  await p.goto('file:///C:/Users/User/Desktop/92-duty-scheduler/index.html');
  await p.waitForSelector('#stepImport',{state:'visible',timeout:5000});

  // Upload xlsx
  const fi=await p.locator('#fileInput');
  await fi.setInputFiles('C:/Users/User/Desktop/92-duty-scheduler/2-26+.xlsx');
  await p.waitForSelector('#stepConfig',{state:'visible',timeout:10000});

  // Set week number
  await p.fill('#weekNum','1');
  await p.evaluate(()=>{document.getElementById('weekNum').dispatchEvent(new Event('change'));});

  // Enable only 值班 and 夜巡
  const cards=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards){
    const txt=await c.textContent();
    if(txt.includes('值班')&&!txt.includes('警技館')&&!txt.includes('總隊')){
      if(!(await c.evaluate(el=>el.classList.contains('selected'))))await c.click();
    }else if(txt.includes('夜巡')){
      if(!(await c.evaluate(el=>el.classList.contains('selected'))))await c.click();
    }else{
      if(await c.evaluate(el=>el.classList.contains('selected')))await c.click();
    }
  }
  await p.waitForTimeout(300);

  // Generate + fill schedule
  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:15000});
  await p.click('#btnFillEmpty');
  await p.waitForTimeout(3000);
  console.log('OK: Schedule generated and filled');

  // Find a student assigned to 值班
  const student=await p.evaluate(()=>{
    var keys=Object.keys(S.assignments['值班']||{});
    for(var k of keys){
      var arr=S.assignments['值班'][k];
      for(var a of arr){
        if(a){var parts=k.split('|');return{id:a.id,name:a.name,day:parts[0],slot:parts[1]};}
      }
    }
    return null;
  });
  if(!student){console.log('FAIL: No student in 值班');await b.close();return;}
  console.log('Test student:',JSON.stringify(student));

  // Check each 夜巡 slot on the same day for overlap
  const results=await p.evaluate((s)=>{
    var def=DUTY_DEFS['夜巡'];var out=[];
    for(var slot of def.slots){
      var fakeAssigned=[{day:s.day,slot:s.slot}];
      var overlaps=slotsOverlap(fakeAssigned,s.day,slot);
      var avail=getAvailable('夜巡',s.day,slot,new Set());
      var inAvail=avail.some(a=>a.id===s.id);
      var all=getAllStudentsForPicker('夜巡',s.day,slot,new Set(),null);
      var entry=all.find(a=>a.id===s.id);
      out.push({nightSlot:slot,daySlot:s.slot,overlap:overlaps,inAvail:inAvail,
        reasons:entry?entry.reasons:['not_in_list'],hasConflict:entry?entry.reasons.includes('conflict'):false});
    }
    return out;
  },student);

  console.log('\n=== Cross-Duty Overlap Detection ===');
  let allPass=true;
  for(const r of results){
    console.log('  夜巡 '+r.nightSlot+' vs 值班 '+r.daySlot+': overlap='+r.overlap+', inAvail='+r.inAvail+', reasons=['+r.reasons+']');
    if(r.overlap&&r.inAvail){
      console.log('    FAIL: In available despite overlap');allPass=false;
    }else if(r.overlap){
      console.log('    PASS: Excluded from available');
    }
  }

  // Manual: assign free student to 值班 then check 夜巡
  const manual=await p.evaluate(()=>{
    var assigned=new Set();
    Object.keys(S.assignments).forEach(dt=>{Object.keys(S.assignments[dt]).forEach(k=>{
      S.assignments[dt][k].forEach(p=>{if(p)assigned.add(p.id);});});});
    var free=S.students.find(s=>!assigned.has(s.id)&&s.gender==='男');
    if(!free)return{error:'No free student'};

    var testDay='一',testSlot='2200-2300',key=testDay+'|'+testSlot;
    if(!S.assignments['值班']||!S.assignments['值班'][key])return{error:'Key not found: '+key};
    var oldVal=S.assignments['值班'][key][0];
    S.assignments['值班'][key][0]={id:free.id,name:free.name};
    S.thisWeekIds.add(free.id);

    var aSlots=getAssignedSlots(free.id);
    var nightSlot='2300-0100';
    var overlap=slotsOverlap(aSlots,testDay,nightSlot);
    var avail=getAvailable('夜巡',testDay,nightSlot,new Set());
    var inAvail=avail.some(a=>a.id===free.id);
    var all=getAllStudentsForPicker('夜巡',testDay,nightSlot,new Set(),null);
    var entry=all.find(a=>a.id===free.id);

    // Restore
    S.assignments['值班'][key][0]=oldVal;
    if(!oldVal)S.thisWeekIds.delete(free.id);

    return{student:{id:free.id,name:free.name},day:testDay,dutySlot:testSlot,
      nightSlot:nightSlot,overlap:overlap,inAvail:inAvail,
      reasons:entry?entry.reasons:['not_in_list']};
  });

  console.log('\n=== Manual Test: 值班 2200-2300 vs 夜巡 2300-0100 ===');
  if(manual.error){console.log('  SKIP:',manual.error);}
  else{
    console.log('  Student:',manual.student.id,manual.student.name);
    console.log('  overlap='+manual.overlap+', inAvail='+manual.inAvail+', reasons=['+manual.reasons+']');
    // Adjacent slots: 2200-2300 (end=2300) and 2300-0100 (start=2300) => no overlap (s2<e1 is 1380<1380 = false)
    if(!manual.overlap)console.log('  PASS: Adjacent slots correctly not overlapping');
    else{console.log('  INFO: Adjacent slots do overlap (boundary inclusive)');
      if(!manual.inAvail)console.log('  PASS: Excluded from available');
      else{console.log('  FAIL: In available despite overlap');allPass=false;}
    }
  }

  // slotsOverlap unit tests
  const unit=await p.evaluate(()=>{return{
    'A:2200-0000 vs 2300-0100':slotsOverlap([{day:'一',slot:'2200-0000'}],'一','2300-0100'),
    'B:2000-2200 vs 2300-0100':slotsOverlap([{day:'一',slot:'2000-2200'}],'一','2300-0100'),
    'C:2200-2300 vs 2200-0000':slotsOverlap([{day:'一',slot:'2200-2300'}],'一','2200-0000'),
    'D:0800-1000 vs 0800-1000':slotsOverlap([{day:'一',slot:'0800-1000'}],'一','0800-1000'),
    'E:diff day':slotsOverlap([{day:'一',slot:'0800-1000'}],'二','0800-1000'),
    'F:0600-0800 vs 0800-1000':slotsOverlap([{day:'一',slot:'0600-0800'}],'一','0800-1000')
  };});

  console.log('\n=== slotsOverlap Unit Tests ===');
  const exp={'A:2200-0000 vs 2300-0100':true,'B:2000-2200 vs 2300-0100':false,
    'C:2200-2300 vs 2200-0000':true,'D:0800-1000 vs 0800-1000':true,
    'E:diff day':false,'F:0600-0800 vs 0800-1000':false};
  for(const[k,v]of Object.entries(unit)){
    const e=exp[k],ok=v===e;
    console.log('  '+k+': '+v+' (exp '+e+') '+(ok?'PASS':'FAIL'));
    if(!ok)allPass=false;
  }

  await b.close();
  console.log('\n'+(allPass?'=== ALL CONFLICT TESTS PASSED ===':'=== SOME TESTS FAILED ==='));
})();
