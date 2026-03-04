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

  // Enable 值班 and 夜巡
  const cards=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards){
    const t=await c.textContent();
    if(t.includes('值班')||t.includes('夜巡'))await c.click();
  }
  await p.waitForTimeout(300);
  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:15000});

  console.log('OK: Schedule generated');

  // Check assignment counts
  const counts=await p.evaluate(()=>{
    var r={};
    Object.keys(S.assignments).forEach(function(dt){
      var c=0;Object.keys(S.assignments[dt]).forEach(function(k){
        S.assignments[dt][k].forEach(function(pp){if(pp)c++;});
      });r[dt]=c;
    });return r;
  });
  console.log('Assignments:',JSON.stringify(counts));

  // If auto-fill didn't work, manually assign for testing
  const result=await p.evaluate(()=>{
    // Find a free male student
    var assigned=new Set();
    Object.keys(S.assignments).forEach(function(dt){
      Object.keys(S.assignments[dt]).forEach(function(k){
        S.assignments[dt][k].forEach(function(pp){if(pp)assigned.add(pp.id);});
      });
    });
    var free=S.students.find(function(s){return s.gender==='\u7537'&&!assigned.has(s.id);});
    if(!free)return{error:'No free student'};

    // Manually assign to 值班 2200-0000 on Monday
    var def1=DUTY_DEFS['\u503C\u73ED'];
    var testDay='\u4E00';
    var testSlot=def1.slots[def1.slots.length-1]; // last slot (late night)
    var key=testDay+'|'+testSlot;
    if(!S.assignments['\u503C\u73ED'])S.assignments['\u503C\u73ED']={};
    if(!S.assignments['\u503C\u73ED'][key])S.assignments['\u503C\u73ED'][key]=new Array(S.active['\u503C\u73ED'].pp);
    S.assignments['\u503C\u73ED'][key][0]={id:free.id,name:free.name};
    S.thisWeekIds.add(free.id);

    // Now check: is this student available in 夜巡 for overlapping time?
    var def2=DUTY_DEFS['\u591C\u5DE1'];
    var results=[];
    def2.slots.forEach(function(nslot){
      var overlap=slotsOverlap([{day:testDay,slot:testSlot}],testDay,nslot);
      var avail=getAvailable('\u591C\u5DE1',testDay,nslot,new Set());
      var inAvail=avail.some(function(a){return a.id===free.id;});
      var all=getAllStudentsForPicker('\u591C\u5DE1',testDay,nslot,new Set(),null);
      var entry=all.find(function(a){return a.id===free.id;});
      results.push({
        nightSlot:nslot,
        dutySlot:testSlot,
        overlap:overlap,
        inAvailable:inAvail,
        hasConflict:entry?entry.reasons.includes('conflict'):false,
        reasons:entry?entry.reasons:[]
      });
    });

    // Clean up
    S.assignments['\u503C\u73ED'][key][0]=undefined;
    S.thisWeekIds.delete(free.id);

    return{student:{id:free.id,name:free.name},day:testDay,dutySlot:testSlot,results:results};
  });

  if(result.error){console.log('FAIL:',result.error);await b.close();return;}

  console.log('\nStudent '+result.student.id+' '+result.student.name+' assigned to 值班 '+result.day+'|'+result.dutySlot);
  console.log('\n=== Cross-duty conflict detection ===');
  let allPass=true;
  for(const r of result.results){
    const status=r.overlap?(r.inAvailable?'FAIL':'PASS'):'OK(no overlap)';
    const conflictOk=r.overlap?r.hasConflict:!r.hasConflict;
    console.log('  夜巡 '+r.nightSlot+': overlap='+r.overlap+' inAvail='+r.inAvailable+' conflict='+r.hasConflict+' → '+status);
    if(r.overlap&&r.inAvailable){allPass=false;}
    if(!conflictOk){allPass=false;console.log('    FAIL: conflict badge mismatch');}
  }

  // slotsOverlap unit tests
  const units=await p.evaluate(()=>{
    return[
      {name:'same slot',v:slotsOverlap([{day:'\u4E00',slot:'0800-1000'}],'\u4E00','0800-1000'),exp:true},
      {name:'overlap midnight',v:slotsOverlap([{day:'\u4E00',slot:'2200-0000'}],'\u4E00','2300-0100'),exp:true},
      {name:'adjacent no overlap',v:slotsOverlap([{day:'\u4E00',slot:'0800-1000'}],'\u4E00','1000-1200'),exp:false},
      {name:'different day',v:slotsOverlap([{day:'\u4E00',slot:'0800-1000'}],'\u4E8C','0800-1000'),exp:false},
      {name:'contained',v:slotsOverlap([{day:'\u4E00',slot:'0600-1200'}],'\u4E00','0800-1000'),exp:true},
    ];
  });
  console.log('\n=== slotsOverlap unit tests ===');
  for(const u of units){
    const pass=u.v===u.exp;
    console.log('  '+u.name+': '+u.v+' (exp='+u.exp+') '+(pass?'PASS':'FAIL'));
    if(!pass)allPass=false;
  }

  await b.close();
  console.log('\n'+(allPass?'=== ALL CONFLICT TESTS PASSED ===':'=== SOME TESTS FAILED ==='));
})();
