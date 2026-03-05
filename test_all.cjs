const{chromium}=require('playwright');
(async()=>{
  let allPass=true;
  function check(label,ok){console.log((ok?'  PASS':'  FAIL')+': '+label);if(!ok)allPass=false;}

  const b=await chromium.launch({headless:true});
  const p=await b.newPage();
  p.on('console',m=>{if(m.type()==='error')console.log('PAGE ERR:',m.text());});

  await p.goto('file:///C:/Users/User/Desktop/92-duty-scheduler/index.html',{timeout:60000});
  await p.waitForSelector('#stepImport',{state:'visible',timeout:10000});
  await p.locator('#fileInput').setInputFiles('C:/Users/User/Desktop/92-duty-scheduler/2-26+.xlsx');
  await p.waitForSelector('#stepConfig',{state:'visible',timeout:15000});
  await p.fill('#weekNum','1');
  await p.evaluate(()=>{document.getElementById('weekNum').dispatchEvent(new Event('change'));});

  // Enable all 8 duties
  const cards=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards){if(!(await c.evaluate(el=>el.classList.contains('selected'))))await c.click();}
  await p.waitForTimeout(300);

  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:30000});
  await p.click('#btnFillEmpty');
  await p.waitForTimeout(5000);
  console.log('OK: 排表已產生並填充');

  // === 1. studentMap ===
  console.log('\n=== 1. studentMap 驗證 ===');
  const map=await p.evaluate(()=>{
    var t=S.students.length,m=Object.keys(S.studentMap).length;
    var g=getStudentGender(S.students[0].id);
    return{total:t,mapped:m,gender:g};
  });
  check('studentMap 數量 ('+map.mapped+'/'+map.total+')',map.mapped===map.total);
  check('getStudentGender 非 null',map.gender!==null);

  // === 2. slotsOverlap unit tests ===
  console.log('\n=== 2. slotsOverlap 測試（含跨日）===');
  const unit=await p.evaluate(()=>({
    'A:same-day overlap':slotsOverlap([{day:'一',slot:'2200-0000'}],'一','2300-0100'),
    'B:same-day no overlap':slotsOverlap([{day:'一',slot:'2000-2200'}],'一','2300-0100'),
    'C:same-day partial':slotsOverlap([{day:'一',slot:'2200-2300'}],'一','2200-0000'),
    'D:identical':slotsOverlap([{day:'一',slot:'0800-1000'}],'一','0800-1000'),
    'E:diff day no overlap':slotsOverlap([{day:'一',slot:'0800-1000'}],'二','0800-1000'),
    'F:adjacent':slotsOverlap([{day:'一',slot:'0600-0800'}],'一','0800-1000'),
    'G:cross-day 日→一':slotsOverlap([{day:'日',slot:'2300-0100'}],'一','0000-0200'),
    'H:cross-day adj':slotsOverlap([{day:'六',slot:'2200-0000'}],'日','0000-0200'),
    'I:cross-day no':slotsOverlap([{day:'日',slot:'2300-0100'}],'一','0200-0400'),
  }));
  const exp={A:true,B:false,C:true,D:true,E:false,F:false,G:true,H:false,I:false};
  for(const[k,v]of Object.entries(unit)){
    const key=k.split(':')[0];
    check(k+' = '+v+' (exp '+exp[key]+')',v===exp[key]);
  }

  // === 3. Holiday backup/restore ===
  console.log('\n=== 3. 放假備份與還原 ===');
  const hol=await p.evaluate(()=>{
    var keys=Object.keys(S.assignments['值班']||{});
    for(var i=0;i<keys.length;i++){
      var arr=S.assignments['值班'][keys[i]];
      if(arr&&arr[0]){
        var parts=keys[i].split('|'),person={id:arr[0].id,name:arr[0].name};
        toggleSlotHoliday('值班',parts[0],parts[1],true);
        var cleared=!S.assignments['值班'][keys[i]][0];
        var backed=!!S.holidayBackup['值班|'+keys[i]];
        toggleSlotHoliday('值班',parts[0],parts[1],false);
        var restored=S.assignments['值班'][keys[i]][0];
        return{person:person,cleared:cleared,backed:backed,match:restored&&restored.id===person.id};
      }
    }
    return{error:'no slot'};
  });
  if(hol.error){console.log('  SKIP:',hol.error);}
  else{
    check('放假後清空',hol.cleared);
    check('備份已建立',hol.backed);
    check('還原正確 ('+hol.person.name+')',hol.match);
  }

  // === 4. savedWeeks 防護 ===
  console.log('\n=== 4. savedWeeks 重複儲存防護 ===');
  const sv=await p.evaluate(()=>{
    var b4=JSON.parse(JSON.stringify(S.cumHours));
    saveWeek();
    var a1=JSON.parse(JSON.stringify(S.cumHours));
    saveWeek(); // 2nd should be blocked
    var a2=JSON.parse(JSON.stringify(S.cumHours));
    var doubled=false;
    Object.keys(a2).forEach(function(id){if(a2[id]!==a1[id])doubled=true;});
    return{doubled:doubled};
  });
  check('第二次儲存被阻擋',!sv.doubled);

  // === 5. 全部放假/取消 ===
  console.log('\n=== 5. 全部放假再取消 ===');
  const mass=await p.evaluate(()=>{
    var d='夜巡',def=DUTY_DEFS[d],cfg=S.active[d];
    if(!cfg)return{error:'not active'};
    cfg.days.forEach(function(day){def.slots.forEach(function(slot){toggleSlotHoliday(d,day,slot,true);});});
    var cnt=S.holidays[d]?S.holidays[d].size:0;
    var exp=cfg.days.length*def.slots.length;
    cfg.days.forEach(function(day){def.slots.forEach(function(slot){toggleSlotHoliday(d,day,slot,false);});});
    var aft=S.holidays[d]?S.holidays[d].size:0;
    return{cnt:cnt,exp:exp,aft:aft};
  });
  if(mass.error){console.log('  SKIP:',mass.error);}
  else{
    check('全部放假 ('+mass.cnt+'/'+mass.exp+')',mass.cnt===mass.exp);
    check('全部取消歸零',mass.aft===0);
  }

  // === 6. Export ===
  console.log('\n=== 6. 匯出 Excel ===');
  const[dl]=await Promise.all([
    p.waitForEvent('download',{timeout:15000}),
    p.click('#btnExport')
  ]);
  check('匯出下載成功',!!(await dl.path()));

  // === 7. Gender balance ===
  console.log('\n=== 7. 校巡性別平衡 ===');
  const gen=await p.evaluate(()=>{
    var mixed=0,total=0;
    Object.keys(S.assignments['校巡']||{}).forEach(function(k){
      var arr=S.assignments['校巡'][k];
      if(arr&&arr[0]&&arr[1]){
        total++;
        if(getStudentGender(arr[0].id)!==getStudentGender(arr[1].id))mixed++;
      }
    });
    return{mixed:mixed,total:total};
  });
  console.log('  男女混合組: '+gen.mixed+'/'+gen.total);
  check('性別平衡邏輯生效',gen.mixed>0||gen.total===0);

  // === 8. Per-slot holiday in picker ===
  console.log('\n=== 8. Picker 放假按鈕 ===');
  await p.evaluate(()=>{
    var cells=document.querySelectorAll('.person-cell');
    if(cells.length>0)cells[0].click();
  });
  await p.waitForSelector('#pickerModal.show',{timeout:3000});
  const pickerHol=await p.evaluate(()=>{
    var row=document.getElementById('pickerHolidayRow');
    return{display:row.style.display,hasBtn:row.querySelector('.pk-holiday')!==null};
  });
  check('Picker 有放假按鈕',pickerHol.hasBtn);
  await p.evaluate(()=>{closePicker();});

  await b.close();
  console.log('\n'+(allPass?'=== ALL TESTS PASSED ===':'=== SOME TESTS FAILED ==='));
  process.exit(allPass?0:1);
})();
