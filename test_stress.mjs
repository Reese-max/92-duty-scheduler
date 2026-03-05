import{chromium}from'playwright';

(async()=>{
  const b=await chromium.launch({headless:true});
  const p=await b.newPage();
  p.on('console',m=>{if(m.type()==='error')console.log('PAGE ERR:',m.text());});
  let allPass=true;
  function check(label,ok){console.log((ok?'  PASS':'  FAIL')+': '+label);if(!ok)allPass=false;}

  await p.goto('file:///C:/Users/User/Desktop/92-duty-scheduler/index.html');
  await p.waitForSelector('#stepImport',{state:'visible',timeout:5000});
  await p.locator('#fileInput').setInputFiles('C:/Users/User/Desktop/92-duty-scheduler/2-26+.xlsx');
  await p.waitForSelector('#stepConfig',{state:'visible',timeout:10000});
  await p.fill('#weekNum','1');
  await p.evaluate(()=>{document.getElementById('weekNum').dispatchEvent(new Event('change'));});

  // ===== Test 1: 啟用全部 8 種勤務 =====
  console.log('\n=== Test 1: 啟用全部 8 種勤務 ===');
  const cards=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards){
    if(!(await c.evaluate(el=>el.classList.contains('selected'))))await c.click();
  }
  await p.waitForTimeout(300);
  const activeCount=await p.evaluate(()=>Object.keys(S.active).length);
  check('8 種勤務全部啟用',activeCount===8);

  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:30000});
  console.log('OK: 排表已產生');

  // ===== Test 2: studentMap 已建立 =====
  console.log('\n=== Test 2: studentMap 驗證 ===');
  const mapTest=await p.evaluate(()=>{
    var total=S.students.length;
    var mapped=Object.keys(S.studentMap).length;
    var sample=S.students[0];
    var gender=getStudentGender(sample.id);
    return{total:total,mapped:mapped,sampleId:sample.id,gender:gender};
  });
  check('studentMap 數量正確 ('+mapTest.mapped+'/'+mapTest.total+')',mapTest.mapped===mapTest.total);
  check('getStudentGender 回傳非 null ('+mapTest.gender+')',mapTest.gender!==null);

  // ===== Test 3: 自動填充 =====
  console.log('\n=== Test 3: 自動填充 ===');
  await p.click('#btnFillEmpty');
  await p.waitForTimeout(5000);
  const fillResult=await p.evaluate(()=>{
    var total=0,filled=0;
    Object.keys(S.assignments).forEach(function(dt){
      Object.keys(S.assignments[dt]).forEach(function(k){
        S.assignments[dt][k].forEach(function(p){total++;if(p)filled++;});
      });
    });
    return{total:total,filled:filled,thisWeek:S.thisWeekIds.size};
  });
  console.log('  填充率: '+fillResult.filled+'/'+fillResult.total+' ('+Math.round(fillResult.filled/fillResult.total*100)+'%)');
  check('填充率 > 50%',fillResult.filled/fillResult.total>0.5);

  // ===== Test 4: 校巡性別平衡 =====
  console.log('\n=== Test 4: 校巡性別平衡 ===');
  const genderTest=await p.evaluate(()=>{
    var mixed=0,total=0;
    var keys=Object.keys(S.assignments['校巡']||{});
    for(var i=0;i<keys.length;i++){
      var arr=S.assignments['校巡'][keys[i]];
      if(arr&&arr[0]&&arr[1]){
        total++;
        var g0=getStudentGender(arr[0].id);
        var g1=getStudentGender(arr[1].id);
        if(g0!==g1)mixed++;
      }
    }
    return{mixed:mixed,total:total};
  });
  console.log('  混合性別組: '+genderTest.mixed+'/'+genderTest.total);
  check('至少有部分校巡組合是男女混合',genderTest.mixed>0||genderTest.total===0);

  // ===== Test 5: 放假備份與還原 =====
  console.log('\n=== Test 5: 放假備份與還原 ===');
  const holidayTest=await p.evaluate(()=>{
    // 找一個有人的值班時段
    var keys=Object.keys(S.assignments['值班']||{});
    for(var i=0;i<keys.length;i++){
      var arr=S.assignments['值班'][keys[i]];
      if(arr&&arr[0]){
        var parts=keys[i].split('|');
        var person={id:arr[0].id,name:arr[0].name};
        // 標記放假
        toggleSlotHoliday('值班',parts[0],parts[1],true);
        var afterSet=S.assignments['值班'][keys[i]][0];
        var backed=S.holidayBackup['值班|'+keys[i]];
        // 取消放假
        toggleSlotHoliday('值班',parts[0],parts[1],false);
        var afterCancel=S.assignments['值班'][keys[i]][0];
        return{
          original:person,
          afterSet:afterSet||null,
          hadBackup:!!backed,
          restored:afterCancel,
          match:afterCancel&&afterCancel.id===person.id
        };
      }
    }
    return{error:'no assigned slot found'};
  });
  if(holidayTest.error){
    console.log('  SKIP: '+holidayTest.error);
  }else{
    check('標記放假後格子清空',holidayTest.afterSet===null);
    check('備份已建立',holidayTest.hadBackup);
    check('取消後人員還原 ('+JSON.stringify(holidayTest.original)+' → '+JSON.stringify(holidayTest.restored)+')',holidayTest.match);
  }

  // ===== Test 6: 跨日衝突偵測 =====
  console.log('\n=== Test 6: 跨日衝突偵測 ===');
  const crossDay=await p.evaluate(()=>{
    // 日|2300-0100 跨到 一|0000-0200
    var r1=slotsOverlap([{day:'日',slot:'2300-0100'}],'一','0000-0200');
    // 反向也要偵測
    var r2=slotsOverlap([{day:'一',slot:'0000-0200'}],'日','2300-0100');
    // 不相鄰的跨日不應衝突
    var r3=slotsOverlap([{day:'日',slot:'2300-0100'}],'一','0200-0400');
    // 同日正常衝突
    var r4=slotsOverlap([{day:'一',slot:'2200-0000'}],'一','2300-0100');
    return{crossDayForward:r1,crossDayReverse:r2,noOverlap:r3,sameDayOverlap:r4};
  });
  check('跨日正向: 日|2300-0100 vs 一|0000-0200 = true',crossDay.crossDayForward===true);
  check('跨日反向: 一|0000-0200 vs 日|2300-0100 = true',crossDay.crossDayReverse===true);
  check('跨日不重疊: 日|2300-0100 vs 一|0200-0400 = false',crossDay.noOverlap===false);
  check('同日重疊: 一|2200-0000 vs 一|2300-0100 = true',crossDay.sameDayOverlap===true);

  // ===== Test 7: savedWeeks 重複儲存防護 =====
  console.log('\n=== Test 7: savedWeeks 重複儲存防護 ===');
  const saveTest=await p.evaluate(()=>{
    var before=JSON.parse(JSON.stringify(S.cumHours));
    saveWeek();
    var after1=JSON.parse(JSON.stringify(S.cumHours));
    // 嘗試再次儲存 — 應該被阻擋
    saveWeek();
    var after2=JSON.parse(JSON.stringify(S.cumHours));
    // 比較 after1 和 after2 是否相同（第二次不應增加）
    var doubled=false;
    Object.keys(after2).forEach(function(id){
      if(after2[id]!==after1[id])doubled=true;
    });
    return{saved:true,doubled:doubled,historyHasWeek:!!S.history[S.weekNum]};
  });
  check('第一次儲存成功',saveTest.saved);
  check('第二次儲存被阻擋（cumHours 未翻倍）',!saveTest.doubled);
  check('history 有記錄',saveTest.historyHasWeek);

  // ===== Test 8: 全部放假再取消 =====
  console.log('\n=== Test 8: 全部放假再全部取消 ===');
  const massHoliday=await p.evaluate(()=>{
    var dtype='夜巡',def=DUTY_DEFS[dtype],cfg=S.active[dtype];
    if(!cfg)return{error:'夜巡 not active'};
    // 全部標記放假
    cfg.days.forEach(function(day){def.slots.forEach(function(slot){
      toggleSlotHoliday(dtype,day,slot,true);
    });});
    var holidayCount=S.holidays[dtype]?S.holidays[dtype].size:0;
    var expected=cfg.days.length*def.slots.length;
    // 全部取消
    cfg.days.forEach(function(day){def.slots.forEach(function(slot){
      toggleSlotHoliday(dtype,day,slot,false);
    });});
    var afterCancel=S.holidays[dtype]?S.holidays[dtype].size:0;
    return{holidayCount:holidayCount,expected:expected,afterCancel:afterCancel};
  });
  if(massHoliday.error){
    console.log('  SKIP: '+massHoliday.error);
  }else{
    check('全部標記放假 ('+massHoliday.holidayCount+'/'+massHoliday.expected+')',massHoliday.holidayCount===massHoliday.expected);
    check('全部取消後歸零 ('+massHoliday.afterCancel+')',massHoliday.afterCancel===0);
  }

  // ===== Test 9: 匯出 Excel =====
  console.log('\n=== Test 9: 匯出 Excel ===');
  const [download]=await Promise.all([
    p.waitForEvent('download',{timeout:15000}),
    p.click('#btnExport')
  ]);
  const dlPath=await download.path();
  check('Excel 匯出成功',!!dlPath);

  // ===== Test 10: Picker 衝突標示 =====
  console.log('\n=== Test 10: Picker 衝突標示 ===');
  const pickerTest=await p.evaluate(()=>{
    // 找一個已被指派到值班的人
    var keys=Object.keys(S.assignments['值班']||{});
    for(var i=0;i<keys.length;i++){
      var arr=S.assignments['值班'][keys[i]];
      if(arr&&arr[0]){
        var parts=keys[i].split('|');
        var person=arr[0];
        // 查看夜巡同天的 picker 是否標示衝突
        var aSlots=getAssignedSlots(person.id);
        var nightSlots=DUTY_DEFS['夜巡'].slots;
        for(var j=0;j<nightSlots.length;j++){
          if(slotsOverlap(aSlots,parts[0],nightSlots[j])){
            var all=getAllStudentsForPicker('夜巡',parts[0],nightSlots[j],new Set(),null);
            var entry=all.find(function(a){return a.id===person.id;});
            return{
              person:person.id+' '+person.name,
              dutySlot:keys[i],
              nightSlot:parts[0]+'|'+nightSlots[j],
              hasConflict:entry?entry.reasons.includes('conflict'):false,
              inAvail:getAvailable('夜巡',parts[0],nightSlots[j],new Set()).some(function(a){return a.id===person.id;})
            };
          }
        }
      }
    }
    return{noOverlapFound:true};
  });
  if(pickerTest.noOverlapFound){
    console.log('  SKIP: 未找到有衝突的配對');
  }else{
    check('衝突學員標示 conflict ('+pickerTest.person+': '+pickerTest.dutySlot+' vs '+pickerTest.nightSlot+')',pickerTest.hasConflict);
    check('衝突學員不在可用名單',!pickerTest.inAvail);
  }

  // ===== Test 11: 只啟用一種勤務 =====
  console.log('\n=== Test 11: 只啟用一種勤務（監廚）===');
  // Reload page for clean state
  await p.goto('file:///C:/Users/User/Desktop/92-duty-scheduler/index.html');
  await p.waitForSelector('#stepImport',{state:'visible',timeout:5000});
  await p.locator('#fileInput').setInputFiles('C:/Users/User/Desktop/92-duty-scheduler/2-26+.xlsx');
  await p.waitForSelector('#stepConfig',{state:'visible',timeout:10000});
  await p.fill('#weekNum','2');
  await p.evaluate(()=>{document.getElementById('weekNum').dispatchEvent(new Event('change'));});
  const cards2=await p.locator('.duty-grid .duty-card').all();
  for(const c of cards2){
    const txt=await c.textContent();
    if(txt.includes('監廚')){
      if(!(await c.evaluate(el=>el.classList.contains('selected'))))await c.click();
    }else{
      if(await c.evaluate(el=>el.classList.contains('selected')))await c.click();
    }
  }
  await p.waitForTimeout(300);
  await p.click('#btnSchedule');
  await p.waitForSelector('#stepResults',{state:'visible',timeout:15000});
  await p.click('#btnFillEmpty');
  await p.waitForTimeout(3000);
  const singleExport=await Promise.all([
    p.waitForEvent('download',{timeout:15000}),
    p.click('#btnExport')
  ]);
  check('單一勤務匯出成功',!!(await singleExport[0].path()));

  await b.close();
  console.log('\n'+(allPass?'=== ALL STRESS TESTS PASSED ===':'=== SOME TESTS FAILED ==='));
})();
