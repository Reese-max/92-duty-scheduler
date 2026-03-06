(function(){
  var v=document.getElementById('bgVideo');
  var mob=window.matchMedia('(max-width:768px)');
  function setSrc(){v.src=mob.matches?'bg-mobile.mp4':'bg-desktop.mp4';v.load();}
  setSrc();mob.addEventListener('change',setSrc);
})();

/* ===== Global error handler ===== */
window.onerror=function(msg,src,line){
  console.error('錯誤：'+msg+' (行 '+line+')');
  var t=document.getElementById('toast');
  if(t){t.textContent='系統錯誤：'+msg;t.classList.add('show');setTimeout(function(){t.classList.remove('show');},4000);}
};
window.addEventListener('unhandledrejection',function(e){
  console.error('未處理的 Promise 錯誤：',e.reason);
});
/* ===== Safe DOM helpers ===== */
function h(tag,attrs,children){
  var e=document.createElement(tag);
  if(attrs) Object.keys(attrs).forEach(function(k){
    if(k==='className') e.className=attrs[k];
    else if(k==='textContent') e.textContent=attrs[k];
    else if(k.indexOf('on')===0) e.addEventListener(k.slice(2).toLowerCase(),attrs[k]);
    else e.setAttribute(k,attrs[k]);
  });
  if(children){
    if(typeof children==='string') e.textContent=children;
    else if(Array.isArray(children)) children.forEach(function(c){if(c)e.appendChild(c);});
    else e.appendChild(children);
  }
  return e;
}
function clearEl(el){while(el.firstChild)el.removeChild(el.firstChild);}

/* ===== CONSTANTS ===== */
var SEMESTER_START=new Date(2026,2,2);
function getSemesterStart(){
  var el=document.getElementById('semesterStart');
  if(el&&el.value){var parts=el.value.split('-');return new Date(parseInt(parts[0]),parseInt(parts[1])-1,parseInt(parts[2]));}
  return SEMESTER_START;
}
var DEPT_MAP={
  '行政系':'行政系','犯防系預防組':'犯預','犯防系矯治組':'犯矯',
  '交通系交通組':'交通（交交）','交通系電訊組':'交通（交電）',
  '國境系境管組':'國境管理','國境系移民組':'國境移民',
  '行管系':'行管','安全系情報組':'安全情報','安全系社安組':'安全社安',
  '法律系':'法律','資管系':'資管','鑑識系':'鑑識','外事系':'外事',
  '水上系':'水上','刑事系':'刑事','消防系消防組':'消防','消防系災防組':'消防'
};
var ALL_DAYS=['日','一','二','三','四','五','六'];
var DUTY_DEFS={
  '值班':{label:'隊部值班',maleOnly:true,slots:['0600-0800','0800-1000','1000-1200','1200-1350','1350-1600','1600-1800','1800-2000','2000-2200','2200-2300'],defDays:['一','二','三','四','五'],defPP:1},
  '夜巡':{label:'友誼/警英夜巡',maleOnly:true,slots:['2300-0100','0100-0300','0300-0500'],defDays:['日','一','二','三','四','五'],defPP:2},
  '週間':{label:'週間校巡',maleOnly:true,slots:['2000-2200','2200-2400'],defDays:['一','二','三','四'],defPP:2},
  '校巡':{label:'假日校巡',maleOnly:false,slots:['0600-0800','0800-1000','1000-1200','1200-1400','1400-1600','1600-1800','1800-2000','2000-2200','2200-2400'],defDays:['五','六','日'],defPP:2},
  '寢巡':{label:'寢巡',maleOnly:true,slots:['1000-1200'],defDays:['五'],defPP:3},
  '警技館':{label:'警技館值班',maleOnly:true,slots:['1700-1800','1800-2000','2000-2200'],defDays:['一','二','三','四','五'],defPP:1},
  '監廚':{label:'監廚',maleOnly:true,slots:['0900-1000','1000-1100','1500-1600','1600-1700'],defDays:['一','二','三','四','五'],defPP:1},
  '總隊值星':{label:'總隊值星',maleOnly:true,slots:['0000-0200','0200-0400','0400-0600','0600-0800','0800-1000','1000-1200','1200-1350','1350-1550','1550-1800','1800-2000','2000-2200','2200-0000'],defDays:['日','一','二','三','四','五'],defPP:1}
};

/* ===== EXPORT CONFIG（匯出設定常數） ===== */
var EXPORT_FONT='DFKai-SB';
var EXPORT_FONT_SIZES={'值班':{title:22,hdrLabel:18,dateVal:20,dayVal:24,time:18,data:24,note:20,holiday:28},'夜巡':{title:20,hdrLabel:16,dateVal:20,dayVal:16,time:18,data:24,note:16,holiday:24},'週間':{title:20,hdrLabel:16,dateVal:20,dayVal:16,time:18,data:24,note:16,holiday:24},'校巡':{title:22,hdrLabel:16,dateVal:16,dayVal:16,time:16,data:16,note:16,holiday:16},'寢巡':{title:22,hdrLabel:18,dateVal:20,dayVal:24,time:18,data:24,note:20,holiday:24},'警技館':{title:22,hdrLabel:18,dateVal:20,dayVal:24,time:18,data:24,note:20,holiday:24},'監廚':{title:20,hdrLabel:18,dateVal:20,dayVal:20,time:18,data:24,note:20,holiday:28},'總隊值星':{title:26,hdrLabel:20,dateVal:20,dayVal:20,time:18,data:25,note:20,holiday:25}};
var EXPORT_COL_WIDTHS={'值班':{time:9,id:7,name:14},'夜巡':{time:11,id:7,name:14},'週間':{time:11,id:7,name:14},'校巡':{time:13,id:7,name:11},'寢巡':{time:11,id:7,name:14},'警技館':{time:11,id:7,name:14},'監廚':{time:19,id:8,name:12},'總隊值星':{time:9,id:7,name:21}};
var EXPORT_HDR_HEIGHTS={'夜巡':{date:28,day:49},'校巡':{date:29,day:22},'總隊值星':{date:29,day:48},'_default':{date:27.75,day:33.75}};
var EXPORT_NOTE_HEIGHTS={'值班':[30,49.5,39.75,30,51],'夜巡':[39.75,39.75,39.75,30,69.75],'週間':[39.75,39.75,39.75,30,39.75],'校巡':[37,37,37,37,37],'寢巡':[30,49.5,39.75,30,39.75],'警技館':[30,49.5,39.75,30,39.75],'監廚':[39.75,39.75,39.75,30,39.75],'總隊值星':[42,42,42,42,42]};
var EXPORT_TIME_FONT={'寢巡':'標楷體-繁','警技館':'標楷體-繁'};
var EXPORT_SLOT_FMT={'值班':'raw','監廚':'raw','總隊值星':'raw','校巡':'raw','夜巡':'fullwidth','寢巡':'fullwidth','警技館':'fullwidth','週間':'fullwidth'};
var EXPORT_HORIZ={'校巡':true};
var EXPORT_TITLE_MAP={'值班':'隊部值班表','夜巡':'友誼、警英夜巡勤務表','週間':'週間校巡勤務表','校巡':'總隊值星假日校園巡邏勤務表','寢巡':' 寢巡勤務表','警技館':' 警技館值班勤務表','監廚':'監廚勤務表','總隊值星':' 總隊值星週隊部值班表'};
var EXPORT_SHEET_ORDER=['值班','夜巡','週間','校巡','寢巡','警技館','監廚','總隊值星'];
var EXPORT_DAY_ORDER={'校巡':['五','六','日'],'警技館':['五','一','二','三','四']};
var EXPORT_SLOT_ORDER={'夜巡':['0100-0300','0300-0500','2300-0100']};
var EXPORT_MIN_COLS={'寢巡':11};
/* 每種勤務的列印設定（A4直向為主，校巡水平佈局用橫向，時數總表橫向多頁） */
var EXPORT_PAGE_SETUP={
  '值班':    {orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '夜巡':    {orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '週間':    {orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '校巡':    {orientation:'landscape',fitToWidth:1,fitToHeight:1,paperSize:9},
  '寢巡':    {orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '警技館':  {orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '監廚':    {orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '總隊值星':{orientation:'portrait',fitToWidth:1,fitToHeight:1,paperSize:9},
  '_時數總表':{orientation:'landscape',fitToWidth:1,fitToHeight:0,paperSize:9}
};
var EXPORT_MARGINS={left:0.25,right:0.25,top:0.25,bottom:0.25,header:0.1,footer:0.1};
var EXPORT_NOTE_TEXT='備註：\n1、輪值勤務，請同學務必提早5-10分鐘到勤。\n2、請同學互相提醒值勤時間，特別是夜間值班同學。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，收假前未簽名確認者，將依規定記點處分。';
var EXPORT_NOTE_TEXTS={
  '值班':'備註：\n1、輪值勤務，請同學務必提早5-10分鐘到勤。\n2、請同學互相提醒值勤時間，特別是夜間值班同學。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，收假前未簽名確認者，將依規定記點處分。',
  '夜巡':'備註：\n1、擔服週間巡邏勤務，請著整齊服裝(勤務帽、制服、勤務鞋)，提前5-10分鐘至總隊值星隊簽到，並領取無線電、巡邏表及相關配件。\n2、請同學互相提醒值勤時間，切勿脫勤。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，若收假前未簽名確認者，將依規定記點處分。',
  '週間':'備註：\n1、擔服週間巡邏勤務，請著整齊服裝(勤務帽、制服、勤務鞋)，提前5-10分鐘至總隊值星隊簽到，並領取無線電、巡邏表及相關配件。\n2、請同學互相提醒值勤時間，切勿脫勤。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，若收假前未簽名確認者，將依規定記點處分。',
  '校巡':'備註：\n1、擔服巡邏勤務，請著整齊服裝(勤務帽、制服、勤務鞋)，提前5-10分鐘至隊部簽到，並領取無線電、巡邏表及相關配件。\n2、請同學互相提醒值勤時間，切勿脫勤。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正。',
  '寢巡':'備註：\n1、輪值勤務，請同學務必提早5-10分鐘到勤。\n2、請同學互相提醒值勤時間，特別是夜間值班同學。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，收假前未簽名確認者，將依規定記點處分。',
  '警技館':'備註：\n1、值班同學於警技館三樓重訓室旁之裁判休息室擔服值班勤務(著期服或黑運)，值班時需注意並留心學生、場館及器材的安全狀況。\n2、擔任服最後一班勤務(2000-2200)的同學，須於2150開始進行警技館出入管制 (只出不進)，並於2200進行清場，在確認警技館內(3樓-B1)已無人員逗留後，方能將燈火全數關閉並退勤\n3、警技館值班同學先至隊部簽到，再前往警技館簽到服勤，到勤後撥打電話向隊部值班同學回報。\n4、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n5、本表公布後，請同學於名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，若逾時未簽名者，將依規定記點處分。',
  '監廚':'備註：\n1、輪值勤務，請同學務必提早5-10分鐘到勤。\n2、請同學互相提醒值勤時間。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，若逾時未簽名者，將依規定記點處分。',
  '總隊值星':'備註：\n1、輪值勤務，請同學務必提早5-10分鐘到勤。\n2、請同學互相提醒值勤時間，特別是夜間值班同學。\n3、如有疑問(如衝堂、因故無法值班等)，請向勤務排表員921127江睿承(B218)、921165謝頡譯(B218)反映。\n4、本表公布後，請同學於自己名字下方簽名(押時間)確認，如需更改，請報由區隊長核章修正，收假前未簽名確認者，將依規定記點處分。'
};

/* ===== STATE ===== */
var S={students:[],freePeriods:{},weekNum:1,active:{},assignments:{},thisWeekIds:new Set(),history:{},cumHours:{},cumHoursByType:{},cumCountByType:{},holidays:{},holidayBackup:{},workbook:null,savedWeeks:new Set(),excluded:new Set(),locked:new Set()};
var undoStack=[];

/* ===== UTILS ===== */
function byId(id){return document.getElementById(id);}
/* $ 遮蔽已移除，統一使用 byId() */
var toastTimer=null;
function showToast(msg,type){
  var t=byId('toast');t.textContent=msg;
  t.className='toast '+(type||'success');
  t.classList.add('show');
  if(toastTimer)clearTimeout(toastTimer);
  toastTimer=setTimeout(function(){t.classList.remove('show');toastTimer=null;},2500);
}
function timeToMin(t){return parseInt(t.slice(0,2))*60+parseInt(t.slice(2));}
function slotHours(slot){var p=slot.split('-'),s=timeToMin(p[0]),e=timeToMin(p[1]);if(e<=s)e+=1440;return(e-s)/60;}
function fmtSlot(slot){var p=slot.split('-');return p[0].slice(0,2)+':'+p[0].slice(2)+'-'+p[1].slice(0,2)+':'+p[1].slice(2);}
function dayOffset(day){var m={'一':0,'二':1,'三':2,'四':3,'五':4,'六':5,'日':6};return m[day]!=null?m[day]:0;}
function weekDate(wn,day){var d=new Date(getSemesterStart());d.setDate(d.getDate()+(wn-1)*7+dayOffset(day));return(d.getMonth()+1)+'月'+d.getDate()+'日';}
function weekRange(wn){return weekDate(wn,'一')+' — '+weekDate(wn,'日');}

/* ===== Count-up Animation ===== */
function animateCount(el,target){
  var start=0,duration=600,startTime=null;
  function step(ts){
    if(!startTime)startTime=ts;
    var progress=Math.min((ts-startTime)/duration,1);
    var ease=1-Math.pow(1-progress,3);
    el.textContent=String(Math.round(start+ease*(target-start)));
    if(progress<1)requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

/* ===== PARSING ===== */
function classifySheet(name){
  if(name.indexOf('時數')>=0)return'時數';if(name.indexOf('空堂')>=0)return'空堂';
  if(name.indexOf('值星')>=0)return'值星';if(name.indexOf('警技館')>=0)return'警技館';
  if(name.indexOf('值班')>=0)return'值班';if(name.indexOf('夜巡')>=0)return'夜巡';
  if(name.indexOf('週間')>=0)return'週間';if(name.indexOf('校巡')>=0)return'校巡';
  if(name.indexOf('寢巡')>=0)return'寢巡';if(name.indexOf('監廚')>=0)return'監廚';
  return null;
}
function parseRoster(ws){
  var raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null}),out=[];
  for(var r=1;r<raw.length;r++){
    var row=raw[r];
    if(!row||row[0]==null||typeof row[0]!=='number')continue;
    out.push({id:row[0],name:String(row[1]||''),gender:String(row[2]||''),dept:String(row[3]||''),note:String(row[15]||'')});
  }
  return out;
}
function parseFreePeriods(ws){
  var raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null}),fp={};
  for(var r=1;r<raw.length;r++){
    var row=raw[r];if(!row||!row[0])continue;
    var wk=String(row[0]).trim(),desc=String(row[1]||''),deptStr=String(row[3]||'');
    var tm=desc.match(/(\d{4})-(\d{4})/);if(!tm)continue;
    var timeRange=tm[1]+'-'+tm[2];
    var depts=new Set();
    if(deptStr&&deptStr!=='無'){deptStr.split('、').forEach(function(d){d=d.trim();if(d)depts.add(getDeptShort(d));});}
    if(!fp[wk])fp[wk]=[];
    fp[wk].push({time:timeRange,depts:depts});
  }
  return fp;
}
function loadWorkbook(wb){
  S.workbook=wb;S.students=[];S.studentMap={};S.freePeriods={};
  for(var i=0;i<wb.SheetNames.length;i++){
    var name=wb.SheetNames[i],ws=wb.Sheets[name],type=classifySheet(name);
    if(type==='時數')S.students=parseRoster(ws);
    else if(type==='空堂')S.freePeriods=parseFreePeriods(ws);
  }
  /* Build studentMap for O(1) gender lookup */
  S.students.forEach(function(s){S.studentMap[s.id]=s;});
  try{
    var saved=JSON.parse(localStorage.getItem('dutySchedulerHistory')||'{}');
    /* Schema validation: ensure data types are correct */
    if(typeof saved!=='object'||saved===null) saved={};
    if(typeof saved.history!=='object'||saved.history===null) saved.history={};
    if(typeof saved.cumHours!=='object'||saved.cumHours===null) saved.cumHours={};
    if(typeof saved.cumHoursByType!=='object'||saved.cumHoursByType===null) saved.cumHoursByType={};
    /* Validate cumHours values are numbers */
    Object.keys(saved.cumHours).forEach(function(k){
      if(typeof saved.cumHours[k]!=='number'||isNaN(saved.cumHours[k])) saved.cumHours[k]=0;
    });
    S.history=saved.history;S.cumHours=saved.cumHours;S.cumHoursByType=saved.cumHoursByType;
    if(typeof saved.cumCountByType==='object'&&saved.cumCountByType!==null)S.cumCountByType=saved.cumCountByType;else S.cumCountByType={};
    /* Rebuild savedWeeks from history keys to prevent duplicate saves across reloads */
    S.savedWeeks=new Set();
    Object.keys(S.history).forEach(function(k){S.savedWeeks.add(parseInt(k));});
    if(saved.assignments&&typeof saved.assignments==='object'){
      S.assignments=saved.assignments;S.thisWeekIds=new Set();
      Object.keys(S.assignments).forEach(function(dt){
        if(typeof S.assignments[dt]!=='object') return;
        Object.keys(S.assignments[dt]).forEach(function(k){
          (S.assignments[dt][k]||[]).forEach(function(p){if(p&&p.id)S.thisWeekIds.add(p.id);});
        });
      });
    }
    if(saved.holidays&&typeof saved.holidays==='object'){
      S.holidays={};
      Object.keys(saved.holidays).forEach(function(k){
        if(Array.isArray(saved.holidays[k])) S.holidays[k]=new Set(saved.holidays[k]);
      });
    }
    if(Array.isArray(saved.excluded)){S.excluded=new Set(saved.excluded);}
    if(saved.holidayBackup&&typeof saved.holidayBackup==='object'){S.holidayBackup=saved.holidayBackup;}
  }catch(e){
    console.warn('localStorage 資料損壞，已重置：',e.message);
    S.history={};S.cumHours={};S.cumHoursByType={};S.savedWeeks=new Set();
  }
}

/* ===== SCHEDULING ENGINE ===== */
var _unknownDepts = new Set();
function getDeptShort(full){
  if(DEPT_MAP[full]) return DEPT_MAP[full];
  if(full && !_unknownDepts.has(full)) { _unknownDepts.add(full); console.warn('未知科系對應：' + full); }
  return full;
}
function isStudentFree(dept,weekday,slot){
  // 週末：若空堂表有週六/日資料則檢查，否則視為全天可用
  if(weekday==='六'||weekday==='日'){
    var wkEnd='週'+weekday;
    if(!S.freePeriods[wkEnd]||S.freePeriods[wkEnd].length===0)return true;
  }
  var parts=slot.split('-'),ds=timeToMin(parts[0]),de=timeToMin(parts[1]);
  if(de<=ds)de+=1440;
  if(de<=360||ds>=1080)return true;
  var wk='週'+weekday,entries=S.freePeriods[wk]||[],sd=getDeptShort(dept);
  for(var i=0;i<entries.length;i++){
    var ep=entries[i].time.split('-'),es=timeToMin(ep[0]),ee=timeToMin(ep[1]);
    if(ee<=es)ee+=1440;
    if(ds<ee&&de>es){if(!entries[i].depts.has(sd))return false;}
  }
  return true;
}
function getLastWeekIds(){if(S.weekNum<=1)return new Set();var prev=S.history[S.weekNum-1];return prev?new Set(prev.ids):new Set();}
function parseLastWeekExcel(wb){
  var studentIds=new Set(S.students.map(function(s){return s.id;}));
  var found=new Set();
  wb.SheetNames.forEach(function(name){
    var ws=wb.Sheets[name];
    var range=XLSX.utils.decode_range(ws['!ref']||'A1');
    for(var R=range.s.r;R<=range.e.r;R++){
      for(var C=range.s.c;C<=range.e.c;C++){
        var cell=ws[XLSX.utils.encode_cell({r:R,c:C})];
        if(!cell)continue;
        var v=(cell.t==='n')?cell.v:parseInt(cell.v);
        if(v&&studentIds.has(v))found.add(v);
      }
    }
  });
  return Array.from(found);
}
function isNightSlot(slot){return timeToMin(slot.split('-')[0])>=1200;}
function slotsOverlap(slots,day,slot){
  if(!slots||slots.length===0)return false;
  var DAYS=['一','二','三','四','五','六','日'];
  var di=DAYS.indexOf(day);if(di<0)return false;
  var p=slot.split('-'),s1=timeToMin(p[0]),e1=timeToMin(p[1]);
  if(e1<=s1)e1+=1440;
  var a1=di*1440+s1,b1=di*1440+e1,W=10080;
  for(var i=0;i<slots.length;i++){
    var si=DAYS.indexOf(slots[i].day);if(si<0)continue;
    var q=slots[i].slot.split('-'),s2=timeToMin(q[0]),e2=timeToMin(q[1]);
    if(e2<=s2)e2+=1440;
    var a2=si*1440+s2,b2=si*1440+e2;
    /* Direct overlap or week-wraparound (日→一) overlap */
    if(a1<b2&&a2<b1)return true;
    if(a1<b2+W&&a2+W<b1)return true;
    if(a1+W<b2&&a2<b1+W)return true;
  }
  return false;
}
function getAssignedSlots(sid){
  var slots=[];
  Object.keys(S.assignments).forEach(function(dt){
    Object.keys(S.assignments[dt]).forEach(function(k){
      S.assignments[dt][k].forEach(function(p){
        if(p&&p.id===sid){
          var parts=k.split('|');
          slots.push({day:parts[0],slot:parts[1]});
        }
      });
    });
  });
  return slots;
}
/* Use studentMap for O(1) lookup instead of linear find() (Defect #22) */
function getStudentGender(id){var s=S.studentMap[id];return s?s.gender:null;}
/* Check if student is withdrawn (退學/休學/開除) */
function isWithdrawn(s){return /退學|休學|開除/.test(s.name)||/退學|休學|開除/.test(s.note);}
/* Shared eligibility check — single source of truth for all scheduling filters */
function isEligible(s,dutyType,day,slot,opts){
  if(S.excluded.has(s.id))return false;
  if(isWithdrawn(s))return false;
  var def=DUTY_DEFS[dutyType];
  if(def.maleOnly&&s.gender!=='男')return false;
  if(dutyType==='校巡'&&isNightSlot(slot)&&s.gender==='女')return false;
  var lastWeek=opts.lastWeek||getLastWeekIds();
  if(lastWeek.has(s.id))return false;
  if(!isStudentFree(s.dept,day,slot))return false;
  if(S.thisWeekIds.has(s.id)&&!(opts.tempRelease&&opts.tempRelease.has(s.id)))return false;
  var aSlots=opts.slotsMap?opts.slotsMap.get(s.id)||[]:getAssignedSlots(s.id);
  if(opts.tempRelease&&opts.tempRelease.has(s.id)){aSlots=aSlots.filter(function(sl){return sl.day!==day||sl.slot!==slot;});}
  if(slotsOverlap(aSlots,day,slot))return false;
  return true;
}
function getAvailable(dutyType,day,slot,tempRelease,slotsMap){
  return S.students.filter(function(s){
    return isEligible(s,dutyType,day,slot,{tempRelease:tempRelease,slotsMap:slotsMap});
  });
}
function generateEmptySchedule(){
  undoStack=[];
  var btnU=byId('btnUndo');if(btnU)btnU.disabled=true;
  S.assignments={};S.thisWeekIds=new Set();S.holidays={};S.holidayBackup={};S.locked=new Set();
  Object.keys(S.active).forEach(function(dtype){
    var cfg=S.active[dtype],def=DUTY_DEFS[dtype];
    S.assignments[dtype]={};
    cfg.days.forEach(function(day){
      def.slots.forEach(function(slot){
        var key=day+'|'+slot;
        S.assignments[dtype][key]=[];
        for(var i=0;i<cfg.pp;i++){S.assignments[dtype][key].push(undefined);}
      });
    });
  });
}
/* bucketShuffle: 按 keyFn 複合鍵分桶，桶間升序，桶內 Fisher-Yates 洗牌 */
function bucketShuffle(arr,keyFn){
  var bucketMap={},keys=[];
  for(var i=0;i<arr.length;i++){
    var k=keyFn(arr[i]),tag=k[0]+':'+k[1];
    if(!bucketMap[tag]){bucketMap[tag]={k:k,items:[]};keys.push(tag);}
    bucketMap[tag].items.push(arr[i]);
  }
  keys.sort(function(a,b){
    var ka=bucketMap[a].k,kb=bucketMap[b].k;
    return ka[0]!==kb[0]?ka[0]-kb[0]:ka[1]-kb[1];
  });
  var result=[];
  for(var ki=0;ki<keys.length;ki++){
    var items=bucketMap[keys[ki]].items;
    for(var fi=items.length-1;fi>0;fi--){var fj=Math.floor(Math.random()*(fi+1));var ft=items[fi];items[fi]=items[fj];items[fj]=ft;}
    for(var ri=0;ri<items.length;ri++)result.push(items[ri]);
  }
  return result;
}
/* cloneAssignments: 深拷貝 S.assignments（保留 undefined 空格） */
function cloneAssignments(src){
  var dst={};
  Object.keys(src).forEach(function(dtype){
    dst[dtype]={};
    Object.keys(src[dtype]).forEach(function(key){
      var arr=src[dtype][key];
      if(!arr){dst[dtype][key]=arr;return;}
      var copy=[];
      for(var i=0;i<arr.length;i++){
        copy.push(arr[i]?{id:arr[i].id,name:arr[i].name}:undefined);
      }
      dst[dtype][key]=copy;
    });
  });
  return dst;
}
/* 外層：多次嘗試，保留最佳結果（解決 Issue 2：貪心無回溯） */
function fillRemainingSlots(targetDtype){
  var TRIES=5;
  var snapAssign=cloneAssignments(S.assignments);
  var snapIds=new Set(S.thisWeekIds);
  var bestAssign=null,bestIds=null,bestEmpty=Infinity;
  for(var attempt=0;attempt<TRIES;attempt++){
    S.assignments=cloneAssignments(snapAssign);
    S.thisWeekIds=new Set(snapIds);
    var empty=_fillOnce(targetDtype);
    if(empty<bestEmpty){
      bestEmpty=empty;
      bestAssign=cloneAssignments(S.assignments);
      bestIds=new Set(S.thisWeekIds);
    }
    if(bestEmpty===0)break;
  }
  S.assignments=bestAssign;
  S.thisWeekIds=bestIds;
  S.emptySlotCount=bestEmpty;
}
/* 內層：單次填滿邏輯 */
function _fillOnce(targetDtype){
  S.thisWeekIds=new Set();
  var assignedSlots=new Map();
  Object.keys(S.assignments).forEach(function(dtype){
    Object.keys(S.assignments[dtype]).forEach(function(key){
      var parts=key.split('|'),day=parts[0],slot=parts[1];
      (S.assignments[dtype][key]||[]).forEach(function(p){
        if(p){
          S.thisWeekIds.add(p.id);
          if(!assignedSlots.has(p.id))assignedSlots.set(p.id,[]);
          assignedSlots.get(p.id).push({day:day,slot:slot});
        }
      });
    });
  });
  var lastWeek=getLastWeekIds(),targets=[];
  var dtypes=targetDtype?[targetDtype]:Object.keys(S.active);
  dtypes.forEach(function(dtype){
    if(!S.active[dtype])return;
    var cfg=S.active[dtype],def=DUTY_DEFS[dtype];
    cfg.days.forEach(function(day){
      def.slots.forEach(function(slot){
        if(S.holidays[dtype]&&S.holidays[dtype].has(day+'|'+slot))return;
        var key=day+'|'+slot;
        if(!S.assignments[dtype][key])S.assignments[dtype][key]=[];
        var eligible=S.students.filter(function(s){
          return isEligible(s,dtype,day,slot,{lastWeek:lastWeek,slotsMap:assignedSlots});
        }).length;
        for(var i=0;i<cfg.pp;i++){
          var lockKey=dtype+'|'+day+'|'+slot+'|'+i;
          if(S.locked.has(lockKey)&&S.assignments[dtype][key][i]){
            /* 跳過已鎖定且有人的格子 */
            continue;
          }
          if(!S.assignments[dtype][key][i]){
            targets.push({dtype:dtype,day:day,slot:slot,key:key,idx:i,eligible:eligible});
          }
        }
      });
    });
  });
  targets.sort(function(a,b){return a.eligible-b.eligible;});
  var fillCount=0,RESORT_INTERVAL=10;
  for(var t=0;t<targets.length;t++){
    if(fillCount>0&&fillCount%RESORT_INTERVAL===0){
      for(var rt=t;rt<targets.length;rt++){
        var rtg=targets[rt];
        rtg.eligible=S.students.filter(function(s){
          return isEligible(s,rtg.dtype,rtg.day,rtg.slot,{lastWeek:lastWeek,slotsMap:assignedSlots});
        }).length;
      }
      var remaining=targets.slice(t);
      remaining.sort(function(a,b){return a.eligible-b.eligible;});
      for(var ri=0;ri<remaining.length;ri++){targets[t+ri]=remaining[ri];}
    }
    var tg=targets[t];
    var cands=S.students.filter(function(s){
      return isEligible(s,tg.dtype,tg.day,tg.slot,{lastWeek:lastWeek,slotsMap:assignedSlots});
    });
    /* bucketShuffle：按 [該勤務累計時數, 全域累計時數] 分桶排序，桶內隨機（解決 Issue 1+4） */
    cands=bucketShuffle(cands,function(s){
      var typeHrs=(S.cumHoursByType[s.id]&&S.cumHoursByType[s.id][tg.dtype])||0;
      var globalHrs=S.cumHours[s.id]||0;
      return [typeHrs,globalHrs];
    });
    /* 科系多樣性 tiebreaker（解決 Issue 3）：同桶內偏好不同科系 */
    if(cands.length>1){
      var assigned=S.assignments[tg.dtype][tg.key]||[];
      var usedDepts={};
      assigned.forEach(function(p){if(p){var d=getDeptShort(S.studentMap[p.id]?S.studentMap[p.id].dept:'');usedDepts[d]=true;}});
      var deptPref=[],deptRest=[];
      for(var ci=0;ci<cands.length;ci++){
        var cd=getDeptShort(cands[ci].dept||'');
        if(!usedDepts[cd])deptPref.push(cands[ci]);else deptRest.push(cands[ci]);
      }
      if(deptPref.length>0&&deptRest.length>0)cands=deptPref.concat(deptRest);
    }
    /* 校巡性別平衡邏輯 */
    if(tg.dtype==='校巡'&&tg.idx>0&&cands.length>0){
      var gAssigned=S.assignments[tg.dtype][tg.key];
      var hasM=gAssigned.some(function(p){return p&&getStudentGender(p.id)==='男';});
      var hasF=gAssigned.some(function(p){return p&&getStudentGender(p.id)==='女';});
      if(hasM||hasF){
        var prefer=hasM?'女':'男';
        var preferred=cands.filter(function(s){return s.gender===prefer;});
        var rest=cands.filter(function(s){return s.gender!==prefer;});
        cands=preferred.concat(rest);
      }
    }
    if(cands.length>0){
      var c=cands[0];
      S.assignments[tg.dtype][tg.key][tg.idx]={id:c.id,name:c.name};
      S.thisWeekIds.add(c.id);
      if(!assignedSlots.has(c.id))assignedSlots.set(c.id,[]);
      assignedSlots.get(c.id).push({day:tg.day,slot:tg.slot});
      fillCount++;
    }
  }
  var emptyCount=0;
  for(var t2=0;t2<targets.length;t2++){
    var tg2=targets[t2];
    if(!S.assignments[tg2.dtype][tg2.key]||!S.assignments[tg2.dtype][tg2.key][tg2.idx])emptyCount++;
  }
  return emptyCount;
}

/* ===== UI: Steps with animation ===== */
function showStep(ids){
  document.querySelectorAll('.step').forEach(function(el){
    if(ids.indexOf(el.id)<0){
      el.classList.remove('active');
    }
  });
  ids.forEach(function(id){
    var el=byId(id);
    if(!el.classList.contains('active')){
      el.classList.add('active');
    }
  });
}

/* ===== UI: Config ===== */
function renderConfig(){
  var grid=byId('dutyGrid');clearEl(grid);
  Object.keys(DUTY_DEFS).forEach(function(dtype){
    var def=DUTY_DEFS[dtype];
    var card=h('div',{className:'duty-card','data-duty':dtype});

    var chk=h('div',{className:'dc-check',role:'checkbox','aria-checked':'false',tabindex:'0'});
    var head=h('div',{className:'dc-head'},[
      chk,
      h('div',{className:'dc-label'},def.label)
    ]);
    card.appendChild(head);

    var detail=h('div',{className:'dc-detail'},def.slots.length+' 時段 | 預設 '+def.defDays.join('')+' | '+(def.maleOnly?'僅男':'男女')+' | 每班 '+def.defPP+' 人');
    card.appendChild(detail);

    var cfg=h('div',{className:'dc-config'});
    var ppLabel=h('label',{},'每班人數 ');
    var ppInput=h('input',{type:'number',className:'cfg-pp',min:'1',max:'10',value:String(def.defPP)});
    ppLabel.appendChild(ppInput);
    cfg.appendChild(ppLabel);

    var dayLabel=h('label',{style:'margin-top:6px'},'啟用天數');
    cfg.appendChild(dayLabel);
    var dayDiv=h('div',{className:'day-toggles'});
    ALL_DAYS.forEach(function(d){
      var isOn=def.defDays.indexOf(d)>=0;
      var btn=h('button',{className:'day-tog'+(isOn?' on':''),'data-day':d},d);
      dayDiv.appendChild(btn);
    });
    cfg.appendChild(dayDiv);

    var orderLabel=h('div',{className:'day-order-label'},'排列順序（可拖曳調整）');
    cfg.appendChild(orderLabel);
    var orderZone=h('div',{className:'day-order-zone','data-duty':dtype});
    cfg.appendChild(orderZone);

    card.appendChild(cfg);
    grid.appendChild(card);
    syncDayOrder(card);

    card.addEventListener('click',function(e){
      if(e.target.classList.contains('day-tog')){e.target.classList.toggle('on');syncDayOrder(card);updateConfigSummary();return;}
      if(e.target.closest('.dc-config')&&!e.target.closest('.dc-head')){return;}
      if(e.target.closest('.cfg-pp'))return;
      card.classList.toggle('selected');
      var ck=card.querySelector('.dc-check'),sel=card.classList.contains('selected');
      ck.textContent=sel?'✓':'';
      ck.setAttribute('aria-checked',sel?'true':'false');
      updateConfigSummary();
    });
    ppInput.addEventListener('change',function(){var v=parseInt(this.value)||1;if(v<1)v=1;if(v>10)v=10;this.value=v;updateConfigSummary();});
    ppInput.addEventListener('input',updateConfigSummary);
  });
  updateWeekDates();
  byId('weekNum').addEventListener('change',updateWeekDates);
  /* 即時過濾非數字輸入，強化週數輸入驗證 */
  byId('weekNum').addEventListener('input',function(){
    var v=this.value.replace(/[^0-9]/g,'');
    if(v!==this.value)this.value=v;
  });
}
function updateWeekDates(){
  var startDate=new Date(getSemesterStart());
  if(startDate.getDay()!==1){
    byId('weekDates').textContent='⚠ 學期起始日應為週一';
    return;
  }
  var raw=parseInt(byId('weekNum').value)||1;
  if(raw<1)raw=1;if(raw>20)raw=20;
  byId('weekNum').value=raw;
  S.weekNum=raw;
  byId('weekDates').textContent=weekRange(S.weekNum);
  var prevWeek=S.weekNum-1;
  var st=byId('lastWeekStatus');
  if(st){
    var prev=S.history[prevWeek];
    if(prev&&prev.ids&&prev.ids.length>0){
      st.textContent='已載入：'+prev.ids.length+' 人（第'+prevWeek+'週）';
      st.className='lastweek-status loaded';
    }else{
      st.textContent='未載入';
      st.className='lastweek-status';
    }
  }
  updateConfigSummary();
}
/* ===== Day Order: sync & drag-and-drop ===== */
function syncDayOrder(card){
  var zone=card.querySelector('.day-order-zone');
  if(!zone)return;
  var existing=[];zone.querySelectorAll('.day-chip').forEach(function(c){existing.push(c.getAttribute('data-day'));});
  var active=[];card.querySelectorAll('.day-tog.on').forEach(function(b){active.push(b.getAttribute('data-day'));});
  /* keep existing order for days still active, append newly added ones at end */
  var ordered;
  if(existing.length>0){
    ordered=existing.filter(function(d){return active.indexOf(d)>=0;});
    active.forEach(function(d){if(ordered.indexOf(d)<0)ordered.push(d);});
  }else{
    /* first render: use defDays order (五六日) instead of DOM order (日五六) */
    var dtype=zone.getAttribute('data-duty');
    var defOrder=DUTY_DEFS[dtype]?DUTY_DEFS[dtype].defDays:[];
    ordered=defOrder.filter(function(d){return active.indexOf(d)>=0;});
    active.forEach(function(d){if(ordered.indexOf(d)<0)ordered.push(d);});
  }
  clearEl(zone);
  ordered.forEach(function(d){
    var chip=h('div',{className:'day-chip',draggable:'true','data-day':d});
    var leftArr=h('span',{className:'chip-arrow chip-left',title:'左移'},'◀');
    var label=document.createTextNode(d);
    var rightArr=h('span',{className:'chip-arrow chip-right',title:'右移'},'▶');
    chip.appendChild(leftArr);chip.appendChild(label);chip.appendChild(rightArr);
    leftArr.addEventListener('click',function(e){e.stopPropagation();moveDayChip(zone,chip,-1);updateConfigSummary();});
    rightArr.addEventListener('click',function(e){e.stopPropagation();moveDayChip(zone,chip,1);updateConfigSummary();});
    /* HTML5 drag-and-drop */
    chip.addEventListener('dragstart',function(e){chip.classList.add('dragging');e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain',d);});
    chip.addEventListener('dragend',function(){chip.classList.remove('dragging');});
    zone.appendChild(chip);
  });
  /* drop zone events（僅首次註冊） */
  if(!zone._dragBound){zone._dragBound=true;
  zone.addEventListener('dragover',function(e){e.preventDefault();e.dataTransfer.dropEffect='move';zone.classList.add('drag-over');});
  zone.addEventListener('dragleave',function(){zone.classList.remove('drag-over');});
  zone.addEventListener('drop',function(e){
    e.preventDefault();zone.classList.remove('drag-over');
    var dragging=zone.querySelector('.dragging');if(!dragging)return;
    var target=getDayChipAtPoint(zone,e.clientX);
    if(target&&target!==dragging){
      var rect=target.getBoundingClientRect();
      if(e.clientX<rect.left+rect.width/2){zone.insertBefore(dragging,target);}
      else{zone.insertBefore(dragging,target.nextSibling);}
    }
    updateConfigSummary();
  });
  }
}
function getDayChipAtPoint(zone,x){
  var chips=zone.querySelectorAll('.day-chip:not(.dragging)');
  var closest=null,closestDist=Infinity;
  chips.forEach(function(c){var r=c.getBoundingClientRect();var mid=r.left+r.width/2;var d=Math.abs(x-mid);if(d<closestDist){closestDist=d;closest=c;}});
  return closest;
}
function moveDayChip(zone,chip,dir){
  var chips=Array.from(zone.querySelectorAll('.day-chip'));
  var idx=chips.indexOf(chip);
  if(idx<0)return;
  var newIdx=idx+dir;
  if(newIdx<0||newIdx>=chips.length)return;
  if(dir<0){zone.insertBefore(chip,chips[newIdx]);}
  else{zone.insertBefore(chip,chips[newIdx].nextSibling);}
}
function updateConfigSummary(){
  S.active={};var totalSlots=0,maleSlots=0;
  document.querySelectorAll('.duty-card.selected').forEach(function(card){
    var dtype=card.getAttribute('data-duty'),def=DUTY_DEFS[dtype];
    var pp=parseInt(card.querySelector('.cfg-pp').value)||def.defPP;
    /* read day order from the order zone (user-defined order) */
    var zone=card.querySelector('.day-order-zone');
    var days=[];
    if(zone&&zone.children.length>0){zone.querySelectorAll('.day-chip').forEach(function(c){days.push(c.getAttribute('data-day'));});}
    else{card.querySelectorAll('.day-tog.on').forEach(function(b){days.push(b.getAttribute('data-day'));});}
    if(days.length===0)return;
    S.active[dtype]={days:days,pp:pp};
    var sc=def.slots.length*days.length*pp;totalSlots+=sc;
    if(def.maleOnly)maleSlots+=sc;
  });
  var lastWeek=getLastWeekIds();
  var availM=S.students.filter(function(s){return s.gender==='男'&&!lastWeek.has(s.id);}).length;
  var availF=S.students.filter(function(s){return s.gender==='女'&&!lastWeek.has(s.id);}).length;
  var bar=byId('configSummary');
  if(totalSlots===0){bar.style.display='none';byId('btnSchedule').disabled=true;byId('btnSchedule').title='請先至少選擇一項勤務';return;}
  bar.style.display='flex';byId('btnSchedule').disabled=false;byId('btnSchedule').title='';
  var ok=totalSlots<=(availM+availF)&&maleSlots<=availM;
  clearEl(bar);
  var items=[
    {t:'需求',n:totalSlots,c:''},
    {t:'可用男',n:availM,c:''},
    {t:'可用女',n:availF,c:''},
    {t:ok?'✓ 人數充足 (未含空堂限制)':'⚠ 人數不足 (未含空堂限制)',n:'',c:ok?'sb-ok':'sb-err'}
  ];
  items.forEach(function(it){
    var item=h('div',{className:'sb-item'+(it.c?' '+it.c:'')});
    item.appendChild(document.createTextNode(it.t+' '));
    if(it.n!==''){var sp=h('span',{className:'sb-num'},String(it.n));item.appendChild(sp);}
    bar.appendChild(item);
  });
  var lastInfo=h('div',{className:'sb-item',style:'font-size:.78rem;color:var(--text-d)'},'上週已排 '+lastWeek.size+' 人');
  bar.appendChild(lastInfo);
}

/* ===== UI: Results ===== */
function renderStatsBar(){
  var tf=0,te=0,ts=0;
  Object.keys(S.active).forEach(function(dtype){
    var cfg=S.active[dtype],def=DUTY_DEFS[dtype];
    cfg.days.forEach(function(day){def.slots.forEach(function(slot){
      var key=day+'|'+slot,arr=S.assignments[dtype]&&S.assignments[dtype][key]||[];
      var filled=arr.filter(function(p){return!!p;}).length;
      tf+=filled;te+=Math.max(0,cfg.pp-filled);ts+=cfg.pp;
    });});
  });
  var grid=byId('statsGrid');clearEl(grid);
  [{v:ts,l:'總格數',c:'var(--gold)',k:'total'},{v:tf,l:'已填入',c:'var(--success)',k:'filled'},{v:te,l:'空缺',c:te?'var(--danger)':'var(--text-d)',k:'empty'},{v:S.thisWeekIds.size,l:'本週排勤人數',c:'var(--info)',k:'people'}].forEach(function(d){
    var card=h('div',{className:'stat-card'+(d.k==='empty'&&te>0?' blink-warn':'')});
    var sv=h('div',{className:'sv',style:'color:'+d.c},'0');
    card.appendChild(sv);
    card.appendChild(h('div',{className:'sl'},d.l));
    grid.appendChild(card);
    setTimeout(function(){animateCount(sv,d.v);},100);
  });
}
function renderResults(){
  renderStatsBar();
  var tabBar=byId('resultTabs'),panels=byId('resultPanels');
  clearEl(tabBar);clearEl(panels);
  var first=true;
  Object.keys(S.active).forEach(function(dtype){
    var def=DUTY_DEFS[dtype];
    var btn=h('button',{className:'tab-btn'+(first?' active':''),'data-duty':dtype},def.label);
    tabBar.appendChild(btn);
    var panel=h('div',{className:'tab-panel'+(first?' active':''),id:'panel-'+dtype});
    var rBtn=h('button',{className:'btn btn-outline btn-sm resched-btn','data-dtype':dtype},'↻ 重新排班此勤務');
    (function(dt){rBtn.addEventListener('click',function(){rescheduleDuty(dt);});})(dtype);
    panel.appendChild(rBtn);
    panel.appendChild(buildDutyTable(dtype));
    panels.appendChild(panel);
    first=false;
  });
  /* 動態填充勤務下拉選單 */
  var sel=byId('fillDutySelect');
  if(sel){
    var prev=sel.value;
    while(sel.firstChild)sel.removeChild(sel.firstChild);
    var defOpt=document.createElement('option');
    defOpt.value='';defOpt.textContent='全部勤務';
    sel.appendChild(defOpt);
    Object.keys(S.active).forEach(function(dtype){
      var def=DUTY_DEFS[dtype];
      var opt=document.createElement('option');
      opt.value=dtype;opt.textContent=def.label;
      sel.appendChild(opt);
    });
    sel.value=prev;
  }
}
function buildDutyTable(dtype){
  var cfg=S.active[dtype],def=DUTY_DEFS[dtype];
  var wrap=h('div',{className:'tbl-wrap'});
  var table=h('table',{className:'dtbl'});
  var thead=h('thead');
  var hr=h('tr');
  hr.appendChild(h('th',{},'時段'));
  cfg.days.forEach(function(day){hr.appendChild(h('th',{},weekDate(S.weekNum,day)+'('+day+')'));});
  thead.appendChild(hr);table.appendChild(thead);
  var tbody=h('tbody');
  def.slots.forEach(function(slot){
    for(var p=0;p<cfg.pp;p++){
      var tr=h('tr');
      if(p===0&&cfg.pp>1)tr.classList.add('row-group-first');
      if(p===0){var tc=h('td',{className:'time-cell'},fmtSlot(slot));if(cfg.pp>1)tc.setAttribute('rowspan',String(cfg.pp));tr.appendChild(tc);}
      cfg.days.forEach(function(day){
        var td=h('td',{className:'person-cell'});
        var slotHoliday=S.holidays[dtype]&&S.holidays[dtype].has(day+'|'+slot);
        if(slotHoliday){
          td.classList.add('holiday-cell');
          td.textContent='放假';
        }else{
          var key=day+'|'+slot,arr=S.assignments[dtype]&&S.assignments[dtype][key]||[];
          var person=arr[p];
          var lockKey=dtype+'|'+day+'|'+slot+'|'+p;
          var isLocked=S.locked.has(lockKey);
          if(isLocked)td.classList.add('locked');
          if(person){
            td.classList.add('filled');
            td.appendChild(h('div',{className:'pid'},String(person.id)));
            td.appendChild(h('div',{className:'pname'},person.name));
            if(isLocked)td.appendChild(h('span',{className:'lock-icon'},'🔒'));
          }else{td.classList.add('empty-slot');td.textContent='+';}
        }
        (function(dt,dy,sl,pi,tdEl){td.addEventListener('click',function(){openPicker(dt,dy,sl,pi,tdEl);})})(dtype,day,slot,p,td);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
  });
  table.appendChild(tbody);wrap.appendChild(table);
  /* Mobile scroll hint — outer wrapper so hint is not clipped by overflow */
  var outer=h('div',{className:'tbl-outer'});
  outer.appendChild(wrap);
  var hint=h('div',{className:'tbl-scroll-hint'});
  outer.appendChild(hint);
  wrap.addEventListener('scroll',function(){
    var maxScroll=wrap.scrollWidth-wrap.clientWidth;
    hint.classList.toggle('hidden',wrap.scrollLeft>=maxScroll-10);
  });
  return outer;
}

/* ===== UI: Picker ===== */
var pickerCtx=null,pickerState={},pickerHighlightIdx=-1;
function getAllStudentsForPicker(dtype,day,slot,slotIds,current,slotsMap){
  var def=DUTY_DEFS[dtype],lastWeek=getLastWeekIds();
  var tempRelease=new Set();if(current)tempRelease.add(current.id);
  return S.students.filter(function(s){
    if(slotIds.has(s.id))return false;
    if(current&&s.id===current.id)return false;
    return true;
  }).map(function(s){
    var reasons=[];
    if(def.maleOnly&&s.gender!=='男')reasons.push('gender');
    if(dtype==='校巡'&&isNightSlot(slot)&&s.gender==='女')reasons.push('night');
    if(S.thisWeekIds.has(s.id)&&!(tempRelease&&tempRelease.has(s.id)))reasons.push('thisWeek');
    if(lastWeek.has(s.id))reasons.push('lastWeek');
    if(!isStudentFree(s.dept,day,slot))reasons.push('busy');
    /* 使用預建的 slotsMap 取代逐人呼叫 getAssignedSlots（效能優化） */
    var aSlots=slotsMap?slotsMap.get(s.id)||[]:getAssignedSlots(s.id);
    if(tempRelease&&tempRelease.has(s.id)){aSlots=aSlots.filter(function(sl){return sl.day!==day||sl.slot!==slot;});}
    if(slotsOverlap(aSlots,day,slot))reasons.push('conflict');
    return{id:s.id,name:s.name,dept:s.dept,gender:s.gender,reasons:reasons};
  });
}
function openPicker(dtype,day,slot,idx,tdEl){
  pickerCtx={dtype:dtype,day:day,slot:slot,idx:idx,tdEl:tdEl};
  var def=DUTY_DEFS[dtype],key=day+'|'+slot;
  var arr=S.assignments[dtype]&&S.assignments[dtype][key]||[];
  var current=arr[idx];
  byId('pickerTitle').textContent=def.label+' | '+weekDate(S.weekNum,day)+'('+day+') '+fmtSlot(slot);
  var slotIds=new Set();arr.forEach(function(p,i){if(p&&i!==idx)slotIds.add(p.id);});
  var tempRelease=new Set();if(current)tempRelease.add(current.id);
  /* 預建 assignedSlotsMap：一次遍歷所有指派，避免逐人呼叫 getAssignedSlots（效能優化） */
  var assignedSlotsMap=new Map();
  Object.keys(S.assignments).forEach(function(dt){
    Object.keys(S.assignments[dt]).forEach(function(k){
      var parts=k.split('|'),d=parts[0],sl=parts[1];
      S.assignments[dt][k].forEach(function(p){
        if(p&&p.id){
          if(!assignedSlotsMap.has(p.id))assignedSlotsMap.set(p.id,[]);
          assignedSlotsMap.get(p.id).push({day:d,slot:sl});
        }
      });
    });
  });
  var candidates=getAvailable(dtype,day,slot,tempRelease,assignedSlotsMap).filter(function(s){return!slotIds.has(s.id);});
  candidates.sort(function(a,b){return(S.cumHours[a.id]||0)-(S.cumHours[b.id]||0);});
  var allStudents=getAllStudentsForPicker(dtype,day,slot,slotIds,current,assignedSlotsMap);
  allStudents.sort(function(a,b){return(a.reasons.length-b.reasons.length)||(S.cumHours[a.id]||0)-(S.cumHours[b.id]||0);});
  var autoShowAll=candidates.length===0;
  var cb=byId('pickerShowAll');
  cb.checked=autoShowAll;
  var toggle=byId('pickerToggle');
  toggle.classList.toggle('auto',autoShowAll);
  byId('pickerToggleLabel').textContent=autoShowAll?'已自動開啟：無符合條件人員，顯示全部學員':'顯示全部學員（忽略限制）';
  pickerState={candidates:candidates,allStudents:allStudents,current:current,showAll:autoShowAll,slotsMap:assignedSlotsMap};
  /* Holiday button (per-slot) */
  var hRow=byId('pickerHolidayRow');clearEl(hRow);
  var hkey=day+'|'+slot;
  var isH=S.holidays[dtype]&&S.holidays[dtype].has(hkey);
  hRow.style.display='block';
  var hBtn=h('div',{className:'pk-holiday'+(isH?' cancel':'')},isH?'↩ 取消此時段放假':'🏖 標記此時段放假');
  hBtn.addEventListener('click',function(){toggleSlotHoliday(dtype,day,slot,!isH);closePicker();});
  hRow.appendChild(hBtn);
  /* Lock button */
  var lockKey=dtype+'|'+day+'|'+slot+'|'+idx;
  var isLk=S.locked.has(lockKey);
  var existingLockBtn=document.querySelector('.btn-lock');
  if(existingLockBtn)existingLockBtn.parentNode.removeChild(existingLockBtn);
  var lockBtn=h('button',{className:'btn-lock'+(isLk?' is-locked':'')},isLk?'🔒 已鎖定（點擊解鎖）':'🔓 鎖定此格');
  (function(lk){lockBtn.addEventListener('click',function(){
    if(S.locked.has(lk))S.locked.delete(lk);
    else S.locked.add(lk);
    renderResults();renderDashboard();
    closePicker();
  });})(lockKey);
  hRow.parentNode.insertBefore(lockBtn,hRow.nextSibling);

  /* If holiday, hide student list */
  var isHolidayMode=isH;
  byId('pickerSearch').parentNode.style.display=isHolidayMode?'none':'';
  toggle.style.display=isHolidayMode?'none':'';
  byId('pickerList').style.display=isHolidayMode?'none':'';
  refreshPickerList();
  byId('pickerSearch').value='';
  byId('pickerModal').classList.add('show');
  document.body.style.overflow='hidden';
  if(!isHolidayMode)setTimeout(function(){byId('pickerSearch').focus();},150);
}
function refreshPickerList(){
  pickerHighlightIdx=-1; /* 重新篩選時重設鍵盤高亮索引 */
  var q=byId('pickerSearch').value.trim().toLowerCase().slice(0,50); /* 限制搜尋長度防止 DoS */
  var src=pickerState.showAll?pickerState.allStudents:pickerState.candidates;
  var filtered=q?src.filter(function(s){return String(s.id).indexOf(q)>=0||s.name.toLowerCase().indexOf(q)>=0||s.dept.toLowerCase().indexOf(q)>=0;}):src;
  renderPickerList(filtered,pickerState.current,pickerState.showAll);
}
function renderPickerList(candidates,current,showAll){
  var list=byId('pickerList');clearEl(list);
  list.setAttribute('role','listbox');
  list.setAttribute('aria-label','可選人員清單');
  var clearRow=byId('pickerClearRow');clearEl(clearRow);
  if(current){
    clearRow.style.display='block';
    var clr=h('div',{className:'pk-clear'},'✕ 清除此格（移除 '+current.id+' '+current.name+'）');
    clr.addEventListener('click',function(){assignPerson(null);closePicker();});
    clearRow.appendChild(clr);
  }else{clearRow.style.display='none';}
  if(candidates.length===0){
    var empty=h('div',{className:'pk-empty'},showAll?'搜尋無結果':'目前沒有符合條件的可用人員');
    list.appendChild(empty);
    return;
  }
  if(!showAll){
    var hint=h('div',{className:'pk-hint'},'以下均為上週未排勤、本週未指派、該時段空堂之人員（'+candidates.length+'人）');
    list.appendChild(hint);
  }
  candidates.forEach(function(s){
    var forced=showAll&&s.reasons&&s.reasons.length>0;
    var item=h('div',{className:'pk-item'+(forced?' forced':''),role:'option','aria-selected':'false'});
    item.setAttribute('tabindex','0'); /* 鍵盤導航：允許 pk-item 被聚焦 */
    item.appendChild(h('span',{className:'pk-id'},String(s.id)));
    var nameSpan=h('span',{className:'pk-name'},s.name);
    item.appendChild(nameSpan);
    if(forced){
      s.reasons.forEach(function(r){
        var badge=h('span',{className:'pk-badge '+r});
        if(r==='busy')badge.textContent='有課';
        else if(r==='thisWeek')badge.textContent='本週';
        else if(r==='lastWeek')badge.textContent='上週';
        else if(r==='gender')badge.textContent='性別';
        else if(r==='night')badge.textContent='夜間';
        else if(r==='conflict')badge.textContent='衝突';
        nameSpan.appendChild(badge);
      });
    }
    item.appendChild(h('span',{className:'pk-dept'},s.dept));
    var hrs=S.cumHours[s.id]||0;
    item.appendChild(h('span',{className:'pk-hrs'},hrs?hrs.toFixed(1)+'h':'0h'));
    item.addEventListener('click',function(){assignPerson(s);closePicker();});
    list.appendChild(item);
  });
}
function toggleSlotHoliday(dtype,day,slot,setHoliday){
  if(!S.holidays[dtype])S.holidays[dtype]=new Set();
  var hkey=day+'|'+slot,bkKey=dtype+'|'+hkey;
  if(setHoliday){
    S.holidays[dtype].add(hkey);
    /* Backup assignments before clearing */
    if(S.assignments[dtype]&&S.assignments[dtype][hkey]){
      var backup=S.assignments[dtype][hkey].map(function(p){return p?{id:p.id,name:p.name}:null;});
      if(backup.some(function(p){return p;}))S.holidayBackup[bkKey]=backup;
      S.assignments[dtype][hkey].forEach(function(p){
        if(p){
          var still=false;
          Object.keys(S.assignments).forEach(function(dt){Object.keys(S.assignments[dt]).forEach(function(k){
            if(dt===dtype&&k===hkey)return;
            S.assignments[dt][k].forEach(function(pp){if(pp&&pp.id===p.id)still=true;});
          });});
          if(!still)S.thisWeekIds.delete(p.id);
        }
      });
      for(var i=0;i<S.assignments[dtype][hkey].length;i++)S.assignments[dtype][hkey][i]=undefined;
    }
    showToast(fmtSlot(slot)+' 已標記為放假');
  }else{
    S.holidays[dtype].delete(hkey);
    /* Restore backed-up assignments */
    var backup=S.holidayBackup[bkKey];
    if(backup&&S.assignments[dtype]&&S.assignments[dtype][hkey]){
      for(var i=0;i<backup.length;i++){
        if(backup[i]){S.assignments[dtype][hkey][i]=backup[i];S.thisWeekIds.add(backup[i].id);}
      }
      delete S.holidayBackup[bkKey];
      showToast(fmtSlot(slot)+' 已取消放假（人員已還原）');
    }else{
      showToast(fmtSlot(slot)+' 已取消放假');
    }
  }
  var panel=byId('panel-'+dtype);
  if(panel){clearEl(panel);panel.appendChild(buildDutyTable(dtype));}
  renderStatsBar();renderDashboard();
}
function assignPerson(student){
  if(!pickerCtx)return;
  var ctx=pickerCtx,key=ctx.day+'|'+ctx.slot;
  if(!S.assignments[ctx.dtype])S.assignments[ctx.dtype]={};
  if(!S.assignments[ctx.dtype][key])S.assignments[ctx.dtype][key]=[];
  var arr=S.assignments[ctx.dtype][key],old=arr[ctx.idx];
  /* Push undo entry */
  undoStack.push({dtype:ctx.dtype,key:key,idx:ctx.idx,prev:old?{id:old.id,name:old.name}:undefined,tdEl:ctx.tdEl});
  byId('btnUndo').disabled=false;
  if(old){
    var still=false;
    Object.keys(S.assignments).forEach(function(dt){Object.keys(S.assignments[dt]).forEach(function(k){
      S.assignments[dt][k].forEach(function(p,i){
        if(dt===ctx.dtype&&k===key&&i===ctx.idx)return;
        if(p&&p.id===old.id)still=true;
      });
    });});
    if(!still)S.thisWeekIds.delete(old.id);
  }
  if(student){
    arr[ctx.idx]={id:student.id,name:student.name};
    S.thisWeekIds.add(student.id);
    showToast('已指派 '+student.id+' '+student.name);
  }else{
    arr[ctx.idx]=undefined;
    showToast('已清除');
  }
  /* Update just the clicked cell instead of rebuilding the entire table (Defect #23) */
  if(ctx.tdEl){
    clearEl(ctx.tdEl);ctx.tdEl.className='person-cell';
    var cellLockKey=ctx.dtype+'|'+ctx.day+'|'+ctx.slot+'|'+ctx.idx;
    if(S.locked.has(cellLockKey))ctx.tdEl.classList.add('locked');
    if(student){
      ctx.tdEl.classList.add('filled');
      ctx.tdEl.appendChild(h('div',{className:'pid'},String(student.id)));
      ctx.tdEl.appendChild(h('div',{className:'pname'},student.name));
      if(S.locked.has(cellLockKey))ctx.tdEl.appendChild(h('span',{className:'lock-icon'},'🔒'));
    }else{ctx.tdEl.classList.add('empty-slot');ctx.tdEl.textContent='+';}
  }else{
    var panel=byId('panel-'+ctx.dtype);
    if(panel){clearEl(panel);panel.appendChild(buildDutyTable(ctx.dtype));}
  }
  renderStatsBar();renderDashboard();
}
function closePicker(){byId('pickerModal').classList.remove('show');document.body.style.overflow='';pickerCtx=null;pickerHighlightIdx=-1;}
/* === Picker 鍵盤導航：更新高亮項目並捲動至可見 === */
function updatePickerHighlight(){
  var list=byId('pickerList');
  var items=list.querySelectorAll('.pk-item');
  items.forEach(function(el,i){
    el.classList.toggle('highlighted',i===pickerHighlightIdx);
  });
  if(pickerHighlightIdx>=0&&items[pickerHighlightIdx]){
    items[pickerHighlightIdx].scrollIntoView({block:'nearest',behavior:'smooth'});
  }
}
function undoLastAssign(){
  if(undoStack.length===0){showToast('沒有可復原的操作','warn');return;}
  var u=undoStack.pop();
  /* 整批復原（填滿操作） */
  if(u.type==='fill'){
    S.assignments=u.snapAssign;
    S.thisWeekIds=u.snapIds;
    renderResults();renderDashboard();
    showToast('已復原填滿操作');
    byId('btnUndo').disabled=undoStack.length===0;
    return;
  }
  var arr=S.assignments[u.dtype]&&S.assignments[u.dtype][u.key];
  if(!arr)return;
  var cur=arr[u.idx];
  if(cur){
    var still=false;
    Object.keys(S.assignments).forEach(function(dt){Object.keys(S.assignments[dt]).forEach(function(k){
      S.assignments[dt][k].forEach(function(p,i){if(dt===u.dtype&&k===u.key&&i===u.idx)return;if(p&&p.id===cur.id)still=true;});
    });});
    if(!still)S.thisWeekIds.delete(cur.id);
  }
  arr[u.idx]=u.prev;
  if(u.prev)S.thisWeekIds.add(u.prev.id);
  if(u.tdEl&&document.body.contains(u.tdEl)){
    clearEl(u.tdEl);u.tdEl.className='person-cell';
    if(u.prev){u.tdEl.classList.add('filled');u.tdEl.appendChild(h('div',{className:'pid'},String(u.prev.id)));u.tdEl.appendChild(h('div',{className:'pname'},u.prev.name));}
    else{u.tdEl.classList.add('empty-slot');u.tdEl.textContent='+';}
  }else{
    var panel=byId('panel-'+u.dtype);
    if(panel){clearEl(panel);panel.appendChild(buildDutyTable(u.dtype));}
  }
  renderStatsBar();renderDashboard();
  showToast('已復原'+(u.prev?' → '+u.prev.name:''));
  byId('btnUndo').disabled=undoStack.length===0;
}

/* ===== UI: Dashboard ===== */
function renderDashboard(){
  var weekHrs={};
  Object.keys(S.assignments).forEach(function(dtype){
    Object.keys(S.assignments[dtype]).forEach(function(key){
      var slot=key.split('|')[1],hrs=slotHours(slot);
      (S.assignments[dtype][key]||[]).forEach(function(p){if(p)weekHrs[p.id]=(weekHrs[p.id]||0)+hrs;});
    });
  });
  var dash=byId('dashboard');clearEl(dash);
  var assigned=S.students.filter(function(s){return weekHrs[s.id];});
  assigned.sort(function(a,b){return(weekHrs[b.id]||0)-(weekHrs[a.id]||0);});
  var maxH=1;assigned.forEach(function(s){if(weekHrs[s.id]>maxH)maxH=weekHrs[s.id];});
  var titleH3=h('h3',{style:'margin-bottom:14px;font-size:1.1rem;color:var(--gold);display:flex;align-items:center;gap:8px'});
  titleH3.appendChild(document.createTextNode('本週排勤 '));
  titleH3.appendChild(h('span',{style:'font-size:1.4rem;font-weight:800'},String(assigned.length)));
  titleH3.appendChild(document.createTextNode(' 人'));
  dash.appendChild(titleH3);
  var chart=h('div',{className:'hours-chart'});
  var FOLD_LIMIT=20;
  var foldContainer=null;
  if(assigned.length>FOLD_LIMIT){
    foldContainer=document.createElement('details');
    foldContainer.style.cssText='margin-top:4px';
    var summary=document.createElement('summary');
    summary.style.cssText='cursor:pointer;color:var(--gold);font-size:.88rem;padding:8px 0;user-select:none';
    summary.textContent='展開剩餘 '+(assigned.length-FOLD_LIMIT)+' 人';
    foldContainer.appendChild(summary);
  }
  assigned.forEach(function(s,i){
    var hrs=weekHrs[s.id]||0,pct=Math.round(hrs/maxH*100);
    var ratio=hrs/maxH;
    var hue=Math.round(210-(ratio*174));
    var color='hsl('+hue+',75%,55%)';
    var colorEnd='hsl('+hue+',75%,42%)';
    var row=h('div',{className:'hc-row'});
    row.appendChild(h('span',{className:'hc-name'},s.id+' '+s.name));
    var bw=h('div',{className:'hc-bar-wrap'});
    var bar=h('div',{className:'hc-bar',style:'width:0%;background:linear-gradient(90deg,'+color+','+colorEnd+')'});
    bw.appendChild(bar);
    row.appendChild(bw);
    row.appendChild(h('span',{className:'hc-val'},hrs.toFixed(1)+'h'));
    if(foldContainer&&i>=FOLD_LIMIT){foldContainer.appendChild(row);}
    else{chart.appendChild(row);}
    setTimeout(function(){bar.style.width=pct+'%';},50+i*8);
  });
  if(foldContainer)chart.appendChild(foldContainer);
  dash.appendChild(chart);

  /* === 按勤務類型分類時數長條圖 === */
  var DUTY_COLORS={'值班':'#e74c3c','夜巡':'#3498db','週間':'#2ecc71','校巡':'#f39c12','寢巡':'#9b59b6','警技館':'#1abc9c','監廚':'#e67e22','總隊值星':'#34495e'};
  var weekHrsByType={};
  Object.keys(S.assignments).forEach(function(dtype){
    Object.keys(S.assignments[dtype]).forEach(function(key){
      var slot=key.split('|')[1],hrs=slotHours(slot);
      (S.assignments[dtype][key]||[]).forEach(function(p){
        if(p){
          if(!weekHrsByType[p.id])weekHrsByType[p.id]={};
          weekHrsByType[p.id][dtype]=(weekHrsByType[p.id][dtype]||0)+hrs;
        }
      });
    });
  });
  var activeDtypes=Object.keys(S.active);
  if(activeDtypes.length>0&&assigned.length>0){
    var dtypeWrap=h('div',{className:'dtype-chart-wrap'});
    var dtypeTitle=h('h3',{style:'margin-bottom:12px;font-size:1rem;color:var(--gold);border-left:3px solid var(--gold);padding-left:12px'},'各勤務類型時數分布');
    dtypeWrap.appendChild(dtypeTitle);
    /* 圖例 */
    var legend=h('div',{className:'dtype-legend'});
    activeDtypes.forEach(function(dt){
      var item=h('div',{className:'dtype-legend-item'});
      item.appendChild(h('div',{className:'dtype-legend-swatch',style:'background:'+(DUTY_COLORS[dt]||'#888')}));
      item.appendChild(document.createTextNode(DUTY_DEFS[dt].label));
      legend.appendChild(item);
    });
    dtypeWrap.appendChild(legend);
    /* 排序控制 */
    var controls=h('div',{className:'dtype-chart-controls'});
    controls.appendChild(h('span',{style:'font-size:.78rem;color:var(--text-d)'},'排序：'));
    var sortOptions=[{key:'total',label:'總時數'}];
    activeDtypes.forEach(function(dt){sortOptions.push({key:dt,label:DUTY_DEFS[dt].label});});
    var currentSort='total';
    var dtypeChartBody=h('div',{id:'dtypeChartBody'});
    function renderDtypeRows(sortKey){
      clearEl(dtypeChartBody);
      var sortedStudents=assigned.slice();
      if(sortKey==='total'){
        sortedStudents.sort(function(a,b){return(weekHrs[b.id]||0)-(weekHrs[a.id]||0);});
      }else{
        sortedStudents.sort(function(a,b){
          var ha=weekHrsByType[a.id]&&weekHrsByType[a.id][sortKey]||0;
          var hb=weekHrsByType[b.id]&&weekHrsByType[b.id][sortKey]||0;
          return hb-ha;
        });
      }
      var maxTotal=1;
      sortedStudents.forEach(function(s){if((weekHrs[s.id]||0)>maxTotal)maxTotal=weekHrs[s.id];});
      sortedStudents.forEach(function(s,i){
        var totalH=weekHrs[s.id]||0;
        var row=h('div',{className:'dtype-bar-row'});
        row.appendChild(h('span',{className:'hc-name'},s.id+' '+s.name));
        var bw=h('div',{className:'hc-bar-wrap'});
        activeDtypes.forEach(function(dt){
          var dtHrs=weekHrsByType[s.id]&&weekHrsByType[s.id][dt]||0;
          if(dtHrs>0){
            var pct=(dtHrs/maxTotal*100).toFixed(1);
            var seg=h('div',{className:'dtype-bar-segment',style:'width:0%;background:'+(DUTY_COLORS[dt]||'#888'),title:DUTY_DEFS[dt].label+' '+dtHrs.toFixed(1)+'h'});
            bw.appendChild(seg);
            (function(s2){setTimeout(function(){s2.style.width=pct+'%';},50+i*8);})(seg);
          }
        });
        row.appendChild(bw);
        row.appendChild(h('span',{className:'hc-val'},totalH.toFixed(1)+'h'));
        dtypeChartBody.appendChild(row);
      });
    }
    sortOptions.forEach(function(opt){
      var btn=h('button',{className:'dtype-sort-btn'+(opt.key==='total'?' active':''),'data-sort':opt.key},opt.label);
      btn.addEventListener('click',function(){
        currentSort=opt.key;
        controls.querySelectorAll('.dtype-sort-btn').forEach(function(b){b.classList.remove('active');});
        btn.classList.add('active');
        renderDtypeRows(opt.key);
      });
      controls.appendChild(btn);
    });
    dtypeWrap.appendChild(controls);
    dtypeWrap.appendChild(dtypeChartBody);
    renderDtypeRows('total');
    dash.appendChild(dtypeWrap);
  }
}

/* ===== EXPORT (styled xlsx-js-style, matches original format) ===== */
function downloadBlob(blob,fname){
  /* Mobile: use Web Share API (opens native share sheet on iOS/Android) */
  if(typeof navigator.share==='function'&&typeof File!=='undefined'){
    try{
      var file=new File([blob],fname,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      if(navigator.canShare&&navigator.canShare({files:[file]})){
        navigator.share({files:[file],title:fname}).then(function(){
          showToast('Excel 已分享');
        }).catch(function(err){
          if(err.name!=='AbortError')blobAnchorDownload(blob,fname);
        });
        return;
      }
    }catch(e){/* File constructor not supported, fall through */}
  }
  blobAnchorDownload(blob,fname);
}
function blobAnchorDownload(blob,fname){
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');
  a.href=url;a.download=fname;a.style.display='none';
  document.body.appendChild(a);
  a.click();
  setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url);},500);
  showToast('Excel 已下載');
}
/**
 * 格式化時段字串（匯出用）
 * @param {string} slot - 時段字串，如 '0100-0300'
 * @param {string} fmt - 格式：'raw'=原始, 'fullwidth'=全形冒號+換行, 'halfwidth'=半形冒號
 * @returns {string} 格式化後的時段字串
 */
function fmtSlotEx(slot,fmt){var p=slot.split('-');if(fmt==='raw')return p[0]+'-'+p[1];if(fmt==='fullwidth')return p[0].slice(0,2)+'：'+p[0].slice(2)+'-\n'+p[1].slice(0,2)+'：'+p[1].slice(2);return p[0].slice(0,2)+':'+p[0].slice(2)+'-'+p[1].slice(0,2)+':'+p[1].slice(2);}

/**
 * 建立匯出用的所有儲存格樣式物件
 * @param {Object} fs - 字型大小設定（含 title, hdrLabel, dateVal, dayVal, time, data, note, holiday）
 * @param {string} FN - 主要字型名稱
 * @param {string} timeFN - 時段欄位字型名稱
 * @param {Object} bM - 粗邊框設定
 * @param {Object} bT - 細邊框設定
 * @param {Object} grayFill - 灰色填充樣式
 * @returns {Object} 包含所有樣式物件的集合
 */
function buildExportStyles(fs,FN,timeFN,bM,bT,grayFill){
  var bdrAll={top:bM,bottom:bM,left:bM,right:bM};
  var bdrId={top:bM,bottom:bM,left:bM,right:bT};
  var bdrName={top:bM,bottom:bM,left:bT,right:bM};
  return{
    sTitle:{font:{sz:fs.title,name:FN},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sHdrLabel:{font:{sz:fs.hdrLabel,name:FN},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sDateVal:{font:{sz:fs.dateVal,name:FN},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sDayVal:{font:{sz:fs.dayVal,name:FN},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sTime:{font:{sz:fs.time,name:timeFN},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sNote:{font:{sz:fs.note,name:FN},alignment:{horizontal:'left',vertical:'top',wrapText:true},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sId:{font:{sz:fs.data,name:FN},alignment:{horizontal:'center',vertical:'center'},border:bdrId,fill:{patternType:'solid',fgColor:{rgb:'D8D8D8'}}},
    sName:{font:{sz:fs.data,name:FN},alignment:{horizontal:'center',vertical:'center'},border:bdrName,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sSig:{font:{sz:fs.data,name:FN},alignment:{horizontal:'center',vertical:'center'},border:bdrAll,fill:{patternType:'solid',fgColor:{rgb:'FFFFFF'}}},
    sNameEmpty:{font:{sz:fs.data,name:FN},alignment:{horizontal:'center',vertical:'center'},border:bdrName,fill:grayFill},
    sSigEmpty:{font:{sz:fs.data,name:FN},alignment:{horizontal:'center',vertical:'center'},border:bdrAll,fill:grayFill},
    sHoliday:{font:{sz:fs.holiday||fs.data,name:FN},alignment:{horizontal:'center',vertical:'center'},border:bdrAll,fill:grayFill}
  };
}

/**
 * 建立工作表標頭列（標題、日期、星期，共 3 列）
 * @param {string} dtype - 勤務類型
 * @param {Array} days - 星期陣列
 * @param {number} colsPerDay - 每天佔用的欄數
 * @param {number} totalCols - 總欄數
 * @param {Object} hh - 標頭列高設定 {date, day}
 * @param {string} titleLabel - 表格標題文字（如 '隊部值班表'）
 * @returns {Object} {aoa: 列資料陣列, merges: 合併範圍陣列, rh: 列高陣列}
 */
function buildSheetHeader(dtype,days,colsPerDay,totalCols,hh,titleLabel){
  var aoa=[],merges=[],rh=[];
  /* 第 0 列：標題 */
  var title='中央警察大學 92期1隊  114學年度第2學期'+titleLabel+'  (第'+S.weekNum+'週)';
  var titleRow=new Array(totalCols).fill('');titleRow[0]=title;
  aoa.push(titleRow);merges.push({s:{r:0,c:0},e:{r:0,c:totalCols-1}});
  rh.push({hpt:33});
  /* 第 1 列：日期 */
  var dateRow=new Array(totalCols).fill('');dateRow[0]='日期';
  for(var i=0;i<days.length;i++){var dc=1+i*colsPerDay;dateRow[dc]=weekDate(S.weekNum,days[i]);merges.push({s:{r:1,c:dc},e:{r:1,c:dc+colsPerDay-1}});}
  aoa.push(dateRow);rh.push({hpt:hh.date});
  /* 第 2 列：星期 */
  var dayRow=new Array(totalCols).fill('');dayRow[0]='星期';
  for(var i=0;i<days.length;i++){var dc=1+i*colsPerDay;dayRow[dc]=days[i];merges.push({s:{r:2,c:dc},e:{r:2,c:dc+colsPerDay-1}});}
  aoa.push(dayRow);rh.push({hpt:hh.day});
  return{aoa:aoa,merges:merges,rh:rh};
}

/**
 * 建立工作表資料列（各時段指派人員）
 * @param {string} dtype - 勤務類型
 * @param {Array} days - 星期陣列
 * @param {number} pp - 每時段人數
 * @param {boolean} horiz - 是否為水平排列
 * @param {number} colsPerDay - 每天佔用的欄數
 * @param {Array} exportSlots - 匯出用的時段陣列
 * @param {string} slotFmt - 時段格式化方式
 * @param {Object} holidaySlotMap - 放假時段對照表
 * @param {number} totalCols - 總欄數
 * @param {number} rowsPerSlot - 每時段佔用的列數
 * @returns {Object} {aoa: 列資料陣列, merges: 合併範圍陣列, rh: 列高陣列, endRow: 結束列號}
 */
function buildSheetData(dtype,days,pp,horiz,colsPerDay,exportSlots,slotFmt,holidaySlotMap,totalCols,rowsPerSlot){
  var aoa=[],merges=[],rh=[];
  var r=3;
  exportSlots.forEach(function(slot,exportSlotIdx){
    if(horiz){
      /* 水平排列：每時段 2 列（姓名+簽名），人員並排 */
      var nameRow=new Array(totalCols).fill('');
      var sigRow=new Array(totalCols).fill('');
      nameRow[0]=fmtSlotEx(slot,slotFmt);
      for(var i=0;i<days.length;i++){
        var key=days[i]+'|'+slot,arr=S.assignments[dtype]&&S.assignments[dtype][key]||[];
        for(var p=0;p<pp;p++){
          var dc=1+i*colsPerDay+p*2;
          if(arr[p]){nameRow[dc]=String(arr[p].id).padStart(3,'0');nameRow[dc+1]=arr[p].name;}
        }
        /* 合併簽名格（跳過放假時段） */
        if(!(holidaySlotMap[i]&&holidaySlotMap[i][exportSlotIdx])){for(var p=0;p<pp;p++){
          var dc=1+i*colsPerDay+p*2;
          merges.push({s:{r:r+1,c:dc},e:{r:r+1,c:dc+1}});
        }}
      }
      aoa.push(nameRow);aoa.push(sigRow);
      merges.push({s:{r:r,c:0},e:{r:r+1,c:0}});
      rh.push({hpt:45});rh.push({hpt:60});
      r+=2;
    }else{
      /* 垂直排列：每時段 pp*2 列，每人佔姓名列+簽名列 */
      for(var p=0;p<pp;p++){
        var nameRow=new Array(totalCols).fill('');
        var sigRow=new Array(totalCols).fill('');
        if(p===0)nameRow[0]=fmtSlotEx(slot,slotFmt);
        for(var i=0;i<days.length;i++){
          var key=days[i]+'|'+slot,arr=S.assignments[dtype]&&S.assignments[dtype][key]||[];
          var dc=1+i*2;
          if(arr[p]){nameRow[dc]=String(arr[p].id).padStart(3,'0');nameRow[dc+1]=arr[p].name;}
          /* 合併簽名列（跳過放假時段） */
          if(!(holidaySlotMap[i]&&holidaySlotMap[i][exportSlotIdx]))merges.push({s:{r:r+p*2+1,c:dc},e:{r:r+p*2+1,c:dc+1}});
        }
        aoa.push(nameRow);aoa.push(sigRow);
        rh.push({hpt:45});rh.push({hpt:60});
      }
      /* 合併時段欄跨所有列 */
      if(rowsPerSlot>1)merges.push({s:{r:r,c:0},e:{r:r+rowsPerSlot-1,c:0}});
      r+=rowsPerSlot;
    }
  });
  return{aoa:aoa,merges:merges,rh:rh,endRow:r};
}

/**
 * 建立放假區域的合併範圍，並追蹤放假儲存格位置
 * @param {Object} holidaySlotMap - 放假時段對照表
 * @param {Array} days - 星期陣列
 * @param {boolean} horiz - 是否為水平排列
 * @param {number} colsPerDay - 每天佔用的欄數
 * @param {number} pp - 每時段人數
 * @returns {Object} {merges: 合併範圍陣列, holidayCells: 放假儲存格座標集合, holidayAnchors: 放假錨點陣列}
 */
function buildHolidayMerges(holidaySlotMap,days,horiz,colsPerDay,pp){
  var merges=[];
  var holidayCells=new Set();
  var holidayAnchors=[];
  Object.keys(holidaySlotMap).forEach(function(diStr){
    var di=parseInt(diStr);
    var dc=1+di*(horiz?colsPerDay:2),endC=dc+(horiz?colsPerDay:2)-1;
    var sis=Object.keys(holidaySlotMap[di]).map(Number).sort(function(a,b){return a-b;});
    var gs=[],gs0=sis[0],gs1=sis[0];
    for(var k=1;k<sis.length;k++){if(sis[k]===gs1+1)gs1=sis[k];else{gs.push([gs0,gs1]);gs0=gs1=sis[k];}}
    gs.push([gs0,gs1]);
    gs.forEach(function(g){
      var sr=horiz?(3+g[0]*2):(3+g[0]*pp*2);
      var er=horiz?(3+g[1]*2+1):(3+g[1]*pp*2+pp*2-1);
      merges.push({s:{r:sr,c:dc},e:{r:er,c:endC}});
      holidayAnchors.push({r:sr,c:dc});
      for(var rr=sr;rr<=er;rr++)for(var cc=dc;cc<=endC;cc++)holidayCells.add(rr+','+cc);
    });
  });
  return{merges:merges,holidayCells:holidayCells,holidayAnchors:holidayAnchors};
}

/**
 * 套用所有儲存格樣式至工作表
 * @param {Object} ws - XLSX 工作表物件
 * @param {Object} range - 工作表範圍（由 decode_range 取得）
 * @param {Object} styles - 樣式物件集合（由 buildExportStyles 產生）
 * @param {number} noteR - 備註列起始列號
 * @param {Set} holidayCells - 放假儲存格座標集合
 * @param {Array} holidayAnchors - 放假錨點陣列
 * @param {number} DATA_START_ROW - 資料起始列號（標頭列之後）
 * @param {number} [COLS_PER_PERSON=2] - 每人佔用的欄數（學號+姓名=2）
 */
function applySheetStyles(ws,range,styles,noteR,holidayCells,holidayAnchors,DATA_START_ROW,COLS_PER_PERSON){
  if(!COLS_PER_PERSON)COLS_PER_PERSON=2;
  /* 在合併錨點儲存格寫入「放假」文字 */
  holidayAnchors.forEach(function(a){var addr=XLSX.utils.encode_cell(a);ws[addr]={t:'s',v:'放假'};});
  /* 逐格套用樣式 */
  for(var R=range.s.r;R<=range.e.r;R++){
    for(var C=range.s.c;C<=range.e.c;C++){
      var addr=XLSX.utils.encode_cell({r:R,c:C});
      if(!ws[addr])ws[addr]={t:'s',v:''};
      if(R===0){ws[addr].s=styles.sTitle;}
      else if(R===1){ws[addr].s=(C===0)?styles.sHdrLabel:styles.sDateVal;}
      else if(R===2){ws[addr].s=(C===0)?styles.sHdrLabel:styles.sDayVal;}
      else if(R>=noteR){ws[addr].s=styles.sNote;}
      else if(C===0){ws[addr].s=styles.sTime;}
      else{
        /* 檢查是否為放假合併區域 */
        var inH=holidayCells.has(R+','+C);
        if(inH){ws[addr].s=styles.sHoliday;}
        else{
        /* 判斷此格是否為空（無人指派）*/
        var rowOff=(R-DATA_START_ROW)%2; /* 0=姓名列, 1=簽名列 */
        var nameR=rowOff===1?R-1:R;
        var colOff=(C-1)%COLS_PER_PERSON; /* 0=學號欄, 1=姓名欄 */
        var idC=C-colOff;
        var idAddr=XLSX.utils.encode_cell({r:nameR,c:idC});
        var slotEmpty=!ws[idAddr]||!ws[idAddr].v;
        if(rowOff===1){ws[addr].s=slotEmpty?styles.sSigEmpty:styles.sSig;}
        else if(colOff===0){ws[addr].s=styles.sId;}
        else{ws[addr].s=slotEmpty?styles.sNameEmpty:styles.sName;}
        }
      }
    }
  }
}

function exportExcel(){
  var wb=XLSX.utils.book_new();
  var FN=EXPORT_FONT;
  /* 邊框預設：粗邊=外框/日期分隔（2pt 對應 ODS），細邊=學號姓名分隔 */
  var bM={style:'medium',color:{rgb:'000000'}},bT={style:'thin',color:{rgb:'000000'}};
  /* 空白格灰色填充 */
  var grayFill={patternType:'solid',fgColor:{rgb:'D8D8D8'}};

  EXPORT_SHEET_ORDER.filter(function(dt){return S.active[dt];}).forEach(function(dtype){
    var cfg=S.active[dtype],def=DUTY_DEFS[dtype];
    var days=cfg.days,pp=cfg.pp,horiz=!!EXPORT_HORIZ[dtype];
    var exportOrder=EXPORT_DAY_ORDER[dtype];
    if(exportOrder){days=exportOrder.filter(function(d){return cfg.days.indexOf(d)>=0;});}
    var exportSlots=EXPORT_SLOT_ORDER[dtype]||def.slots;
    var colsPerDay=horiz?pp*2:2;
    var rowsPerSlot=horiz?2:pp*2;
    var totalCols=1+days.length*colsPerDay;
    var minC=EXPORT_MIN_COLS[dtype];
    if(minC&&totalCols<minC)totalCols=minC;

    /* 建立樣式 */
    var fs=EXPORT_FONT_SIZES[dtype]||{title:22,hdrLabel:18,dateVal:20,dayVal:24,time:18,data:24,note:20};
    var timeFN=EXPORT_TIME_FONT[dtype]||FN;
    var slotFmt=EXPORT_SLOT_FMT[dtype]||'halfwidth';
    var hh=EXPORT_HDR_HEIGHTS[dtype]||EXPORT_HDR_HEIGHTS['_default'];
    var styles=buildExportStyles(fs,FN,timeFN,bM,bT,grayFill);

    /* 建立放假時段對照表 */
    var holidaySet=S.holidays[dtype]||new Set();
    var holidaySlotMap={};
    for(var di=0;di<days.length;di++){for(var esi=0;esi<exportSlots.length;esi++){if(holidaySet.has(days[di]+'|'+exportSlots[esi])){if(!holidaySlotMap[di])holidaySlotMap[di]={};holidaySlotMap[di][esi]=true;}}}

    /* 建立標頭列（標題、日期、星期） */
    var titleLabel=EXPORT_TITLE_MAP[dtype]||def.label;
    var hdr=buildSheetHeader(dtype,days,colsPerDay,totalCols,hh,titleLabel);

    /* 建立資料列（時段指派） */
    var data=buildSheetData(dtype,days,pp,horiz,colsPerDay,exportSlots,slotFmt,holidaySlotMap,totalCols,rowsPerSlot);

    /* 合併標頭與資料 */
    var aoa=hdr.aoa.concat(data.aoa);
    var merges=hdr.merges.concat(data.merges);
    var rh=hdr.rh.concat(data.rh);
    var noteR=data.endRow;

    /* 建立放假區域合併 */
    var hm=buildHolidayMerges(holidaySlotMap,days,horiz,colsPerDay,pp);
    merges=merges.concat(hm.merges);

    /* 寢巡：合併空白區域（欄 3~10） */
    if(EXPORT_MIN_COLS[dtype]&&totalCols>1+days.length*colsPerDay){
      var blankStart=1+days.length*colsPerDay;
      var blankEnd=totalCols-1;
      merges.push({s:{r:1,c:blankStart},e:{r:1,c:blankEnd}});
      merges.push({s:{r:2,c:blankStart},e:{r:2,c:blankEnd}});
      if(noteR>3)merges.push({s:{r:3,c:blankStart},e:{r:noteR-1,c:blankEnd}});
    }

    /* 備註列 */
    var nh=EXPORT_NOTE_HEIGHTS[dtype]||[36,36,36,36,36];
    var noteText=EXPORT_NOTE_TEXTS[dtype]||EXPORT_NOTE_TEXT;
    var noteRow=new Array(totalCols).fill('');noteRow[0]=noteText;
    aoa.push(noteRow);merges.push({s:{r:noteR,c:0},e:{r:noteR+4,c:totalCols-1}});
    rh.push({hpt:nh[0]});
    for(var n=1;n<5;n++){aoa.push(new Array(totalCols).fill(''));rh.push({hpt:nh[n]});}

    /* 建立工作表並套用合併 */
    var ws=XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges']=merges;

    /* 套用樣式 */
    var range=XLSX.utils.decode_range(ws['!ref']);
    applySheetStyles(ws,range,styles,noteR,hm.holidayCells,hm.holidayAnchors,3);

    /* 欄寬（對應 ODS） */
    var cwDef=EXPORT_COL_WIDTHS[dtype]||{time:11,id:7,name:14};
    var cw=[{wch:cwDef.time}];
    for(var i=0;i<days.length;i++){
      if(horiz){for(var p=0;p<pp;p++)cw.push({wch:cwDef.id},{wch:cwDef.name});}
      else{cw.push({wch:cwDef.id},{wch:cwDef.name});}
    }
    ws['!cols']=cw;
    ws['!rows']=rh;
    /* 列印設定：依勤務類型個別設定方向、縮放 */
    ws['!pageSetup']=EXPORT_PAGE_SETUP[dtype]||EXPORT_PAGE_SETUP['值班'];
    ws['!margins']=EXPORT_MARGINS;
    /* 列印範圍：精確設定為資料範圍 */
    var printRange=XLSX.utils.decode_range(ws['!ref']);
    var sheetLabel=def.label.replace(/[:\\\/\?\*\[\]]/g,'').slice(0,31);
    XLSX.utils.book_append_sheet(wb,ws,sheetLabel);
    var sheetIdx=wb.SheetNames.length-1;
    if(!wb.Workbook)wb.Workbook={};
    if(!wb.Workbook.Names)wb.Workbook.Names=[];
    wb.Workbook.Names.push({Name:'_xlnm.Print_Area',Ref:sheetLabel+'!$A$1:$'+XLSX.utils.encode_col(printRange.e.c)+'$'+(printRange.e.r+1),Sheet:sheetIdx});
  });
  if(!wb.SheetNames.length){showToast('請先完成排勤再匯出','error');return;}

  /* ===== 勤務時數總表 ===== */
  (function(){
    var HFN='PMingLiU';var HFS=14;
    var dtypeOrder=['值班','夜巡','校巡','警技館','週間','寢巡','監廚','總隊值星'];
    var colHeaders=['學號','姓名','性別','系別','隊部值班','夜巡時數','假日校巡時數','警技館值班時數','週間校巡時數','寢巡','監廚','實習幹部','勤務總時數','寢巡次數','監廚次數','備註'];
    var colWidths=[5,14,5,14,12,16,14,19,17,17,17,17,15,15,12,12];

    /* 計算本週各勤務時數與次數 */
    var weekHrsByType={},weekCountByType={};
    Object.keys(S.assignments).forEach(function(dtype){
      Object.keys(S.assignments[dtype]).forEach(function(key){
        var slot=key.split('|')[1],hrs=slotHours(slot);
        (S.assignments[dtype][key]||[]).forEach(function(p){
          if(p){
            if(!weekHrsByType[p.id])weekHrsByType[p.id]={};
            weekHrsByType[p.id][dtype]=(weekHrsByType[p.id][dtype]||0)+hrs;
            if(!weekCountByType[p.id])weekCountByType[p.id]={};
            weekCountByType[p.id][dtype]=(weekCountByType[p.id][dtype]||0)+1;
          }
        });
      });
    });

    /* 建立資料列 */
    var rows=[];
    S.students.forEach(function(s){
      var id=s.id,cum=S.cumHoursByType[id]||{},wk=weekHrsByType[id]||{},wc=weekCountByType[id]||{};
      var hrsVal=function(dt){return(cum[dt]||0)+(wk[dt]||0);};
      var totalHrs=0;
      dtypeOrder.forEach(function(dt){totalHrs+=hrsVal(dt);});
      var cc=S.cumCountByType[id]||{};
      var bedCount=(cc['寢巡']||0)+(wc['寢巡']||0);
      var kitchenCount=(cc['監廚']||0)+(wc['監廚']||0);
      rows.push({id:id,name:s.name,gender:s.gender,dept:s.dept,
        v:hrsVal('值班'),n:hrsVal('夜巡'),x:hrsVal('校巡'),j:hrsVal('警技館'),
        z:hrsVal('週間'),q:hrsVal('寢巡'),c:hrsVal('監廚'),shi:0,
        total:totalHrs,bedCount:bedCount,kitchenCount:kitchenCount,note:s.note||''});
    });

    /* 按勤務總時數由少到多排序 */
    rows.sort(function(a,b){return a.total-b.total;});

    /* 建立 AOA */
    var aoa=[colHeaders];
    rows.forEach(function(r){
      aoa.push([String(r.id).padStart(3,'0'),r.name,r.gender,r.dept,
        r.v||'',r.n||'',r.x||'',r.j||'',r.z||'',r.q||'',r.c||'',r.shi||'',
        r.total||'',r.bedCount||'',r.kitchenCount||'',r.note]);
    });

    var ws2=XLSX.utils.aoa_to_sheet(aoa);

    /* 樣式 */
    var hBdr={style:'thin',color:{rgb:'000000'}};
    var range2=XLSX.utils.decode_range(ws2['!ref']);
    for(var R=range2.s.r;R<=range2.e.r;R++){
      for(var C=range2.s.c;C<=range2.e.c;C++){
        var addr=XLSX.utils.encode_cell({r:R,c:C});
        if(!ws2[addr])ws2[addr]={v:'',t:'s'};
        ws2[addr].s={
          font:{name:HFN,sz:HFS,bold:R===0},
          alignment:{horizontal:'center',vertical:'center',wrapText:true},
          border:{top:hBdr,bottom:hBdr,left:hBdr,right:hBdr}
        };
      }
    }

    /* 欄寬 */
    ws2['!cols']=colWidths.map(function(w){return{wch:w};});

    /* 列印設定：時數總表用橫向、不限高度（多頁） */
    ws2['!pageSetup']=EXPORT_PAGE_SETUP['_時數總表'];
    ws2['!margins']=EXPORT_MARGINS;
    /* 列印範圍 */
    var printRange2=XLSX.utils.decode_range(ws2['!ref']);
    XLSX.utils.book_append_sheet(wb,ws2,'勤務時數總表');
    var sheetIdx2=wb.SheetNames.length-1;
    if(!wb.Workbook)wb.Workbook={};
    if(!wb.Workbook.Names)wb.Workbook.Names=[];
    wb.Workbook.Names.push({Name:'_xlnm.Print_Area',Ref:'勤務時數總表!$A$1:$'+XLSX.utils.encode_col(printRange2.e.c)+'$'+(printRange2.e.r+1),Sheet:sheetIdx2});
  })();

  var fname='勤務排班_第'+S.weekNum+'週.xlsx';
  try{
    var wbout=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    /* JSZip 後處理：注入 pageSetup XML（xlsx-js-style 不支援寫入 pageSetup） */
    injectPageSetup(wbout,wb.SheetNames).then(function(finalBuf){
      var blob=new Blob([finalBuf],{type:'application/octet-stream'});
      downloadBlob(blob,fname);
    }).catch(function(e){showToast('匯出失敗：'+e.message,'error');});
  }catch(e){showToast('匯出失敗：'+e.message,'error');}
}

/* 用 JSZip 解壓 XLSX，為每個 sheet XML 注入 <pageSetup> 和 fitToPage */
function injectPageSetup(xlsxBuf,sheetNames){
  return JSZip.loadAsync(xlsxBuf).then(function(zip){
    var tasks=[];
    sheetNames.forEach(function(name,idx){
      var xmlPath='xl/worksheets/sheet'+(idx+1)+'.xml';
      var f=zip.file(xmlPath);
      if(!f)return;
      /* 查找該 sheet 對應的列印設定 */
      var ps=null;
      if(name==='勤務時數總表')ps=EXPORT_PAGE_SETUP['_時數總表'];
      else{
        for(var k in EXPORT_PAGE_SETUP){if(k!=='_時數總表'&&name.indexOf(k)!==-1){ps=EXPORT_PAGE_SETUP[k];break;}}
      }
      if(!ps)ps=EXPORT_PAGE_SETUP['值班'];
      tasks.push(f.async('string').then(function(xml){
        /* 構建 pageSetup XML */
        var orient=ps.orientation||'portrait';
        var ftw=ps.fitToWidth!=null?ps.fitToWidth:1;
        var fth=ps.fitToHeight!=null?ps.fitToHeight:1;
        var paper=ps.paperSize||9;
        var psXml='<pageSetup orientation="'+orient+'" fitToWidth="'+ftw+'" fitToHeight="'+fth+'" paperSize="'+paper+'"/>';
        /* 移除既有的 pageSetup（以防萬一） */
        xml=xml.replace(/<pageSetup[^/]*\/>/g,'');
        xml=xml.replace(/<pageSetup[^>]*>[\s\S]*?<\/pageSetup>/g,'');
        /* 在 </worksheet> 前插入 pageSetup */
        xml=xml.replace('</worksheet>',psXml+'</worksheet>');
        /* 注入 fitToPage：需要在 <sheetPr> 中加入 <pageSetUpPr fitToPage="1"/> */
        if(xml.indexOf('fitToPage')===-1){
          if(xml.indexOf('<sheetPr')!==-1){
            /* 已有 sheetPr，在其結尾前插入 */
            if(xml.indexOf('<sheetPr/>') !== -1){
              xml=xml.replace('<sheetPr/>','<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
            }else{
              xml=xml.replace('</sheetPr>','<pageSetUpPr fitToPage="1"/></sheetPr>');
            }
          }else{
            /* 沒有 sheetPr，在 <sheetData> 前插入 */
            xml=xml.replace('<sheetData>','<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData>');
          }
        }
        zip.file(xmlPath,xml);
      }));
    });
    return Promise.all(tasks).then(function(){
      return zip.generateAsync({type:'arraybuffer',compression:'DEFLATE'});
    });
  });
}

/* ===== BACKUP / RESTORE ===== */
function exportBackup(){
  var data=localStorage.getItem('dutySchedulerHistory');
  if(!data||data==='{}'){showToast('沒有可備份的歷史紀錄','warn');return;}
  var blob=new Blob([data],{type:'application/json'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');
  a.href=url;a.download='勤排備份_'+new Date().toISOString().slice(0,10)+'.json';
  a.style.display='none';document.body.appendChild(a);a.click();
  setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url);},500);
  showToast('歷史備份已匯出');
}
function importBackup(file){
  /* 檔案大小限制：防止匯入過大的備份檔 */
  if(file.size>5*1024*1024){showToast('備份檔案過大（上限 5MB）','error');return;}
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      var data=JSON.parse(e.target.result);
      if(typeof data!=='object'||data===null)throw new Error('格式無效');
      if(!data.history&&!data.cumHours)throw new Error('找不到歷史紀錄資料');
      localStorage.setItem('dutySchedulerHistory',JSON.stringify(data));
      S.history=data.history||{};S.cumHours=data.cumHours||{};S.cumHoursByType=data.cumHoursByType||{};
      if(data.assignments)S.assignments=data.assignments;
      if(data.holidays){S.holidays={};Object.keys(data.holidays).forEach(function(k){if(Array.isArray(data.holidays[k]))S.holidays[k]=new Set(data.holidays[k]);});}
      S.savedWeeks=new Set();Object.keys(S.history).forEach(function(k){S.savedWeeks.add(parseInt(k));});
      showToast('歷史備份已匯入（'+Object.keys(S.history).length+' 週）');
    }catch(err){showToast('匯入失敗：'+err.message,'error');}
  };
  reader.readAsText(file);
}

/* ===== PERSISTENCE ===== */
function saveWeek(){
  /* 檢查是否已儲存過 */
  if(S.savedWeeks.has(S.weekNum)){
    if(!confirm('第 '+S.weekNum+' 週已儲存過。重新儲存會覆蓋原有紀錄，確定嗎？'))return;
    /* 扣除舊的 cumHours */
    var oldHrs=S.history[S.weekNum]&&S.history[S.weekNum].hours||{};
    Object.keys(oldHrs).forEach(function(id){
      S.cumHours[id]=(S.cumHours[id]||0)-oldHrs[id];
      if(S.cumHours[id]<0)S.cumHours[id]=0;
    });
    /* 扣除舊的 cumHoursByType */
    var oldByType=S.history[S.weekNum]&&S.history[S.weekNum].hoursByType||{};
    Object.keys(oldByType).forEach(function(id){
      if(!S.cumHoursByType[id])return;
      Object.keys(oldByType[id]).forEach(function(dt){
        S.cumHoursByType[id][dt]=(S.cumHoursByType[id][dt]||0)-oldByType[id][dt];
        if(S.cumHoursByType[id][dt]<0)S.cumHoursByType[id][dt]=0;
      });
    });
    /* 扣除舊的 cumCountByType */
    var oldCount=S.history[S.weekNum]&&S.history[S.weekNum].countByType||{};
    Object.keys(oldCount).forEach(function(id){
      if(!S.cumCountByType[id])return;
      Object.keys(oldCount[id]).forEach(function(dt){
        S.cumCountByType[id][dt]=(S.cumCountByType[id][dt]||0)-oldCount[id][dt];
        if(S.cumCountByType[id][dt]<0)S.cumCountByType[id][dt]=0;
      });
    });
  }
  /* 計算本週時數與次數 */
  var weekHrs={},weekHrsByType={},weekCountByType={};
  Object.keys(S.assignments).forEach(function(dtype){
    Object.keys(S.assignments[dtype]).forEach(function(key){
      var slot=key.split('|')[1],hrs=slotHours(slot);
      (S.assignments[dtype][key]||[]).forEach(function(p){
        if(p){
          weekHrs[p.id]=(weekHrs[p.id]||0)+hrs;
          if(!weekHrsByType[p.id])weekHrsByType[p.id]={};
          weekHrsByType[p.id][dtype]=(weekHrsByType[p.id][dtype]||0)+hrs;
          if(!weekCountByType[p.id])weekCountByType[p.id]={};
          weekCountByType[p.id][dtype]=(weekCountByType[p.id][dtype]||0)+1;
        }
      });
    });
  });
  /* 加入新時數 */
  Object.keys(weekHrs).forEach(function(id){S.cumHours[id]=(S.cumHours[id]||0)+weekHrs[id];});
  Object.keys(weekHrsByType).forEach(function(id){
    if(!S.cumHoursByType[id])S.cumHoursByType[id]={};
    Object.keys(weekHrsByType[id]).forEach(function(dt){
      S.cumHoursByType[id][dt]=(S.cumHoursByType[id][dt]||0)+weekHrsByType[id][dt];
    });
  });
  /* 累積次數 */
  Object.keys(weekCountByType).forEach(function(id){
    if(!S.cumCountByType[id])S.cumCountByType[id]={};
    Object.keys(weekCountByType[id]).forEach(function(dt){
      S.cumCountByType[id][dt]=(S.cumCountByType[id][dt]||0)+weekCountByType[id][dt];
    });
  });
  /* 在 history 中保存 hours、hoursByType、countByType，以便未來覆蓋時扣除 */
  S.history[S.weekNum]={ids:Array.from(S.thisWeekIds),hours:weekHrs,hoursByType:weekHrsByType,countByType:weekCountByType};
  S.savedWeeks.add(S.weekNum);
  var hSer={};Object.keys(S.holidays).forEach(function(k){if(S.holidays[k]&&S.holidays[k].size)hSer[k]=Array.from(S.holidays[k]);});
  try{localStorage.setItem('dutySchedulerHistory',JSON.stringify({history:S.history,cumHours:S.cumHours,cumHoursByType:S.cumHoursByType,cumCountByType:S.cumCountByType,assignments:S.assignments,holidays:hSer,excluded:Array.from(S.excluded),holidayBackup:S.holidayBackup}));}catch(e){showToast('儲存失敗：儲存空間不足','error');return;}
  /* 自動備份：存一份帶時間戳的備份 */
  try{
    var backupKey='dutySchedulerBackup_'+new Date().toISOString().slice(0,10);
    localStorage.setItem(backupKey,JSON.stringify({history:S.history,cumHours:S.cumHours,cumHoursByType:S.cumHoursByType}));
    updateBackupStatus();
    /* 清除超過 5 份的舊備份 */
    var bkeys=[];
    for(var bi=0;bi<localStorage.length;bi++){var bk=localStorage.key(bi);if(bk&&bk.indexOf('dutySchedulerBackup_')===0)bkeys.push(bk);}
    bkeys.sort();while(bkeys.length>5){localStorage.removeItem(bkeys.shift());}
  }catch(e){console.warn('自動備份失敗：',e.message);}
  showToast('第 '+S.weekNum+' 週紀錄已儲存（'+S.thisWeekIds.size+' 人）');
}
/* 更新備份狀態顯示 */
function updateBackupStatus(){
  var el=byId('backupStatusInfo');
  if(!el)return;
  var lastBackup=null;
  for(var i=localStorage.length-1;i>=0;i--){
    var k=localStorage.key(i);
    if(k&&k.indexOf('dutySchedulerBackup_')===0){
      var d=k.replace('dutySchedulerBackup_','');
      if(!lastBackup||d>lastBackup)lastBackup=d;
    }
  }
  if(lastBackup){
    el.textContent='上次自動備份：'+lastBackup;
    el.style.display='inline';
  }else{
    el.style.display='none';
  }
}

/* ===== INIT ===== */
function handleFile(file){
  /* 驗證檔案大小（上限 10MB） */
  if(file.size>10*1024*1024){showToast('檔案過大（上限 10MB）','error');return;}
  /* 驗證 Excel 檔案副檔名 */
  var validTypes=['.xlsx','.xls','.xlsm','.xlsb','.ods'];
  var ext=file.name.toLowerCase().match(/\.[^.]+$/);
  if(!ext||validTypes.indexOf(ext[0])<0){
    var status=byId('importStatus');clearEl(status);
    status.appendChild(h('div',{className:'import-err'},'✕ 不支援的檔案格式「'+file.name+'」，請選擇 Excel 檔案（.xlsx）'));
    showToast('不支援的檔案格式，請選擇 Excel 檔案（.xlsx）','error');return;
  }
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      var wb=XLSX.read(e.target.result);loadWorkbook(wb);
      if(S.students.length===0)throw new Error('找不到勤務時數總表');
      var status=byId('importStatus');clearEl(status);
      var div=h('div',{className:'import-ok'},
        '✓ 匯入成功：'+S.students.length+' 位學員（男 '+S.students.filter(function(s){return s.gender==='男';}).length+' / 女 '+S.students.filter(function(s){return s.gender==='女';}).length+'），'+Object.keys(S.freePeriods).length+' 天空堂資料');
      status.appendChild(div);
      /* 名冊按鈕與面板 */
      var rosterBtn=h('button',{className:'btn btn-outline btn-sm',style:'margin-top:8px'},'查看名冊');
      var rosterPanel=h('div',{id:'rosterPanel',style:'display:none;margin-top:12px'});
      rosterBtn.addEventListener('click',function(){
        rosterPanel.style.display=rosterPanel.style.display==='none'?'block':'none';
        rosterBtn.textContent=rosterPanel.style.display==='none'?'查看名冊':'收合名冊';
      });
      /* 統計 */
      var maleCount=S.students.filter(function(s){return s.gender==='男';}).length;
      var femaleCount=S.students.filter(function(s){return s.gender==='女';}).length;
      var deptSet=new Set();S.students.forEach(function(s){if(s.dept)deptSet.add(s.dept);});
      var statsDiv=h('div',{className:'roster-stats'},'共 '+S.students.length+' 人（男 '+maleCount+'，女 '+femaleCount+'），'+deptSet.size+' 個科系');
      rosterPanel.appendChild(statsDiv);
      /* 搜尋框 */
      var rosterSearch=h('input',{type:'text',className:'roster-search',placeholder:'搜尋學號、姓名或科系...',autocomplete:'off'});
      rosterPanel.appendChild(rosterSearch);
      /* 表格 */
      var rosterWrap=h('div',{className:'roster-wrap'});
      var rosterTable=h('table',{className:'roster-tbl'});
      var rosterThead=h('thead');
      var rosterHr=h('tr');
      ['學號','姓名','性別','科系'].forEach(function(col){rosterHr.appendChild(h('th',{},col));});
      rosterThead.appendChild(rosterHr);rosterTable.appendChild(rosterThead);
      var rosterTbody=h('tbody',{id:'rosterTbody'});
      S.students.forEach(function(s){
        var tr=h('tr');
        tr.appendChild(h('td',{},String(s.id)));
        tr.appendChild(h('td',{},s.name));
        tr.appendChild(h('td',{},s.gender));
        tr.appendChild(h('td',{},s.dept));
        rosterTbody.appendChild(tr);
      });
      rosterTable.appendChild(rosterTbody);rosterWrap.appendChild(rosterTable);
      rosterPanel.appendChild(rosterWrap);
      /* 搜尋篩選 */
      rosterSearch.addEventListener('input',function(){
        var q=rosterSearch.value.trim().toLowerCase();
        var rows=rosterTbody.querySelectorAll('tr');
        rows.forEach(function(row){
          var text=row.textContent.toLowerCase();
          row.style.display=(!q||text.indexOf(q)>=0)?'':'none';
        });
      });
      status.appendChild(rosterBtn);
      status.appendChild(rosterPanel);
      /* Check for withdrawn students */
      var withdrawn=S.students.filter(isWithdrawn);
      if(withdrawn.length>0){
        var wNames=withdrawn.map(function(s){return s.id+' '+s.name;}).join('、');
        status.appendChild(h('div',{className:'import-err',style:'background:rgba(243,156,18,.1);border-color:rgba(243,156,18,.25);color:var(--warning)'},
          '⚠ 發現 '+withdrawn.length+' 位退學/休學學員（'+wNames+'），排班時將自動排除'));
      }
      /* Check for unknown departments */
      _unknownDepts.clear();
      S.students.forEach(function(s){getDeptShort(s.dept);});
      if(_unknownDepts.size>0){
        status.appendChild(h('div',{className:'import-err',style:'background:rgba(243,156,18,.1);border-color:rgba(243,156,18,.25);color:var(--warning)'},
          '⚠ 發現 '+_unknownDepts.size+' 個未知科系：'+Array.from(_unknownDepts).join('、')+'（空堂判斷可能不準確）'));
      }
      showStep(['stepImport','stepConfig']);
      updateConfigSummary();showToast('匯入成功');
    }catch(err){
      var msg = err.message;
      if(msg.indexOf('找不到')>=0) msg = '找不到「勤務時數總表」工作表，請確認 Excel 檔案格式正確';
      else if(msg.indexOf('File is not')>=0 || msg.indexOf('Unsupported')>=0) msg = '不是有效的 Excel 檔案，請選擇 .xlsx 格式';
      var status=byId('importStatus');clearEl(status);
      status.appendChild(h('div',{className:'import-err'},'✕ 匯入失敗：'+msg));
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ===== 內建自測模組 ===== */
function runSelfTests(){
  var passed=0, failed=0, results=[];
  function assert(name, actual, expected){
    if(actual===expected){passed++;results.push('✓ '+name);}
    else{failed++;results.push('✕ '+name+': 預期 '+JSON.stringify(expected)+' 但得到 '+JSON.stringify(actual));}
  }
  function assertTruthy(name, val){assert(name, !!val, true);}
  function assertFalsy(name, val){assert(name, !!val, false);}

  /* --- timeToMin --- */
  assert('timeToMin("0800")', timeToMin('0800'), 480);
  assert('timeToMin("0000")', timeToMin('0000'), 0);
  assert('timeToMin("2359")', timeToMin('2359'), 1439);
  assert('timeToMin("1200")', timeToMin('1200'), 720);

  /* --- slotHours --- */
  assert('slotHours("0800-1000")', slotHours('0800-1000'), 2);
  assert('slotHours("2300-0100") 跨日', slotHours('2300-0100'), 2);
  assert('slotHours("0600-0800")', slotHours('0600-0800'), 2);
  var sh=slotHours('1200-1350');
  assert('slotHours("1200-1350") 近似值', Math.abs(sh-110/60)<0.001, true);

  /* --- slotsOverlap --- */
  assertTruthy('slotsOverlap: 完全重疊', slotsOverlap([{day:'一',slot:'0800-1000'}], '一', '0800-1000'));
  assertTruthy('slotsOverlap: 部分重疊', slotsOverlap([{day:'一',slot:'0800-1000'}], '一', '0900-1100'));
  assertFalsy('slotsOverlap: 不重疊', slotsOverlap([{day:'一',slot:'0800-1000'}], '一', '1000-1200'));
  assertFalsy('slotsOverlap: 不同天', slotsOverlap([{day:'一',slot:'0800-1000'}], '二', '0800-1000'));
  assertFalsy('slotsOverlap: 空陣列', slotsOverlap([], '一', '0800-1000'));
  assertTruthy('slotsOverlap: 跨日重疊', slotsOverlap([{day:'日',slot:'2300-0100'}], '日', '2200-2400'));

  /* --- fmtSlot --- */
  assert('fmtSlot("0800-1000")', fmtSlot('0800-1000'), '08:00-10:00');
  assert('fmtSlot("2300-0100")', fmtSlot('2300-0100'), '23:00-01:00');

  /* --- dayOffset --- */
  assert('dayOffset("一")', dayOffset('一'), 0);
  assert('dayOffset("日")', dayOffset('日'), 6);
  assert('dayOffset("三")', dayOffset('三'), 2);

  /* --- isNightSlot --- */
  assertTruthy('isNightSlot("2000-2200")', isNightSlot('2000-2200'));
  assertFalsy('isNightSlot("0800-1000")', isNightSlot('0800-1000'));
  assertTruthy('isNightSlot("2300-0100")', isNightSlot('2300-0100'));

  /* --- isWithdrawn --- */
  assertTruthy('isWithdrawn: 退學', isWithdrawn({name:'張三(退學)',note:''}));
  assertTruthy('isWithdrawn: 休學 in note', isWithdrawn({name:'李四',note:'休學中'}));
  assertFalsy('isWithdrawn: 正常', isWithdrawn({name:'王五',note:''}));

  /* --- getDeptShort --- */
  assert('getDeptShort: 行政系', getDeptShort('行政系'), '行政系');
  assert('getDeptShort: 犯防系預防組', getDeptShort('犯防系預防組'), '犯預');

  /* --- isStudentFree（週末應為空堂） --- */
  assertTruthy('isStudentFree: 週六視為空堂', isStudentFree('行政系','六','0800-1000'));
  assertTruthy('isStudentFree: 週日視為空堂', isStudentFree('行政系','日','0800-1000'));
  assertTruthy('isStudentFree: 夜間（18點後）視為空堂', isStudentFree('行政系','一','1800-2000'));
  assertTruthy('isStudentFree: 凌晨（6點前）視為空堂', isStudentFree('行政系','一','0300-0500'));

  /* --- isEligible（模擬測試，需暫存狀態） --- */
  var origStudents=S.students,origThisWeek=S.thisWeekIds,origHistory=S.history,origWeekNum=S.weekNum,origAssignments=S.assignments,origStudentMap=S.studentMap,origExcluded=S.excluded;
  S.students=[{id:1,name:'測試甲',gender:'男',dept:'行政系',note:''},{id:2,name:'測試乙',gender:'女',dept:'行政系',note:''},{id:3,name:'測試丙(退學)',gender:'男',dept:'行政系',note:''}];
  S.studentMap={};S.students.forEach(function(s){S.studentMap[s.id]=s;});
  S.thisWeekIds=new Set();S.history={};S.weekNum=2;S.assignments={};S.excluded=new Set();
  assertTruthy('isEligible: 正常男學員應合格', isEligible(S.students[0],'值班','一','0800-1000',{lastWeek:new Set()}));
  assertFalsy('isEligible: 女學員不得排僅男勤務', isEligible(S.students[1],'值班','一','0800-1000',{lastWeek:new Set()}));
  assertFalsy('isEligible: 退學學員不得排班', isEligible(S.students[2],'值班','一','0800-1000',{lastWeek:new Set()}));
  S.thisWeekIds.add(1);
  assertFalsy('isEligible: 本週已排學員不得重複', isEligible(S.students[0],'值班','一','1000-1200',{lastWeek:new Set()}));
  assertFalsy('isEligible: 上週已排學員不得排班', isEligible(S.students[1],'校巡','六','0800-1000',{lastWeek:new Set([2])}));
  S.students=origStudents;S.thisWeekIds=origThisWeek;S.history=origHistory;S.weekNum=origWeekNum;S.assignments=origAssignments;S.studentMap=origStudentMap;S.excluded=origExcluded;

  /* --- 結果摘要 --- */
  var summary='自測結果：'+passed+' 通過 / '+failed+' 失敗（共 '+(passed+failed)+' 項）';
  console.log(summary);
  results.forEach(function(r){console.log(r);});
  if(failed>0){console.error('⚠ 有 '+failed+' 項測試失敗！');}
  else{console.log('✓ 所有自測通過');}
  return {passed:passed, failed:failed, results:results};
}

/* ===== EXCLUDE STUDENTS ===== */
function renderExcludeList(){
  var list=byId('excludeList');clearEl(list);
  S.excluded.forEach(function(id){
    var s=S.studentMap[id];
    if(!s)return;
    var tag=h('span',{className:'exclude-tag'});
    tag.appendChild(document.createTextNode(s.id+' '+s.name));
    var x=h('span',{className:'etag-x'},'×');
    x.addEventListener('click',function(){S.excluded.delete(id);renderExcludeList();updateConfigSummary();});
    tag.appendChild(x);
    list.appendChild(tag);
  });
}
function initExcludeSearch(){
  var input=byId('excludeSearch'),dd=byId('excludeDropdown');
  input.addEventListener('input',function(){
    var q=input.value.trim().toLowerCase().slice(0,50);
    clearEl(dd);
    if(!q||S.students.length===0){dd.classList.remove('show');return;}
    var matches=S.students.filter(function(s){
      if(S.excluded.has(s.id))return false;
      return String(s.id).indexOf(q)>=0||s.name.toLowerCase().indexOf(q)>=0;
    }).slice(0,10);
    if(matches.length===0){dd.classList.remove('show');return;}
    matches.forEach(function(s){
      var item=h('div',{className:'exclude-drop-item'});
      item.appendChild(h('span',{className:'eid'},String(s.id)));
      item.appendChild(h('span',{},s.name));
      item.addEventListener('click',function(){
        S.excluded.add(s.id);
        renderExcludeList();
        input.value='';clearEl(dd);dd.classList.remove('show');
        updateConfigSummary();
        showToast(s.id+' '+s.name+' 已加入排除名單');
      });
      dd.appendChild(item);
    });
    dd.classList.add('show');
  });
  document.addEventListener('click',function(e){
    if(!e.target.closest('.exclude-search-wrap'))dd.classList.remove('show');
  });
}

/* ===== RESCHEDULE SINGLE DUTY ===== */
function rescheduleDuty(dtype){
  if(!confirm('確定要重新排班「'+DUTY_DEFS[dtype].label+'」嗎？已鎖定的格子不會被覆蓋。'))return;
  /* 移除該勤務已排的人員 ID（但保留其他勤務已排的人） */
  var otherIds=new Set();
  Object.keys(S.assignments).forEach(function(dt){
    if(dt===dtype)return;
    Object.keys(S.assignments[dt]).forEach(function(k){
      (S.assignments[dt][k]||[]).forEach(function(p){if(p&&p.id)otherIds.add(p.id);});
    });
  });
  /* 清除該勤務的 assignments（保留 locked 的格子） */
  var cfg=S.active[dtype],def=DUTY_DEFS[dtype];
  if(!cfg)return;
  cfg.days.forEach(function(day){
    def.slots.forEach(function(slot){
      var key=day+'|'+slot;
      var arr=S.assignments[dtype]&&S.assignments[dtype][key];
      if(!arr)return;
      for(var i=0;i<arr.length;i++){
        var lockKey=dtype+'|'+day+'|'+slot+'|'+i;
        if(S.locked.has(lockKey)&&arr[i]){
          /* 保留鎖定格子的人員 */
          continue;
        }
        arr[i]=undefined;
      }
    });
  });
  /* 重建 thisWeekIds（掃描所有勤務的 assignments） */
  S.thisWeekIds=new Set();
  Object.keys(S.assignments).forEach(function(dt){
    Object.keys(S.assignments[dt]).forEach(function(k){
      (S.assignments[dt][k]||[]).forEach(function(p){if(p&&p.id)S.thisWeekIds.add(p.id);});
    });
  });
  fillRemainingSlots(dtype);
  renderResults();renderDashboard();
  showToast('「'+DUTY_DEFS[dtype].label+'」已重新排班');
}

function init(){
  byId('resultTabs').addEventListener('click',function(e){
    var btn=e.target.closest('.tab-btn');if(!btn)return;
    var tabBar=byId('resultTabs'),panels=byId('resultPanels');
    tabBar.querySelectorAll('.tab-btn').forEach(function(b){b.classList.remove('active');});
    panels.querySelectorAll('.tab-panel').forEach(function(p){p.classList.remove('active');});
    btn.classList.add('active');
    byId('panel-'+btn.getAttribute('data-duty')).classList.add('active');
  });
  renderConfig();
  initExcludeSearch();
  var dz=byId('dropZone'),fi=byId('fileInput');
  dz.addEventListener('click',function(e){if(e.target!==fi)fi.click();});
  dz.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();fi.click();}});
  dz.addEventListener('dragover',function(e){e.preventDefault();dz.classList.add('dragover');});
  dz.addEventListener('dragleave',function(){dz.classList.remove('dragover');});
  dz.addEventListener('drop',function(e){e.preventDefault();dz.classList.remove('dragover');if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);});
  fi.addEventListener('change',function(){if(fi.files[0])handleFile(fi.files[0]);fi.value='';});
  byId('btnSchedule').addEventListener('click',function(){
    /* 防止未匯入學員就排勤 */
    if(S.students.length===0){showToast('請先匯入學員名冊','error');return;}
    var btn=byId('btnSchedule');
    btn.disabled=true;
    var spinner=h('span',{className:'spinner'});
    btn.insertBefore(spinner,btn.firstChild);
    setTimeout(function(){
      generateEmptySchedule();renderResults();renderDashboard();
      showStep(['stepImport','stepConfig','stepResults','stepDashboard']);
      btn.removeChild(spinner);
      btn.disabled=false;
      byId('stepResults').scrollIntoView({behavior:'smooth'});
      showToast('已產生空白勤務表，可手動指派後再填滿');
    },300);
  });
  byId('btnFillEmpty').addEventListener('click',function(){
    if(!S.assignments||Object.keys(S.assignments).length===0){showToast('請先產生勤務表','error');return;}
    var btn=byId('btnFillEmpty');
    btn.disabled=true;
    var sel=byId('fillDutySelect');
    var targetDtype=sel?sel.value:'';
    var spinner=h('span',{className:'spinner'});
    btn.insertBefore(spinner,btn.firstChild);
    setTimeout(function(){
      /* 保存填滿前快照，支援 Ctrl+Z 整批復原 */
      var snapBefore=cloneAssignments(S.assignments);
      var snapIds=new Set(S.thisWeekIds);
      fillRemainingSlots(targetDtype||undefined);
      undoStack.push({type:'fill',snapAssign:snapBefore,snapIds:snapIds});
      byId('btnUndo').disabled=false;
      renderResults();renderDashboard();
      btn.removeChild(spinner);
      btn.disabled=false;
      var scopeName=targetDtype?(DUTY_DEFS[targetDtype]?DUTY_DEFS[targetDtype].label:targetDtype):'所有';
      /* 統計各勤務空缺 */
      var emptyByType={};
      Object.keys(S.active).forEach(function(dtype){
        var cfg=S.active[dtype],def=DUTY_DEFS[dtype],empty=0;
        cfg.days.forEach(function(day){def.slots.forEach(function(slot){
          var key=day+'|'+slot,arr=S.assignments[dtype]&&S.assignments[dtype][key]||[];
          for(var i=0;i<cfg.pp;i++) if(!arr[i]) empty++;
        });});
        if(empty>0) emptyByType[dtype]=empty;
      });
      var totalEmpty=Object.values(emptyByType).reduce(function(a,b){return a+b;},0);
      if(totalEmpty>0){
        var details=Object.keys(emptyByType).map(function(dt){return DUTY_DEFS[dt].label+' '+emptyByType[dt]+'個';}).join('、');
        showToast('共有 '+totalEmpty+' 個空缺：'+details,'warn');
      }else{
        showToast('已隨機填滿'+scopeName+'空格');
      }
    },300);
  });
  byId('pickerSearch').addEventListener('input',function(){refreshPickerList();});
  byId('pickerShowAll').addEventListener('change',function(){
    pickerState.showAll=this.checked;
    byId('pickerToggle').classList.toggle('auto',false);
    byId('pickerToggleLabel').textContent='顯示全部學員（忽略限制）';
    refreshPickerList();
  });
  byId('btnReimport').addEventListener('click',function(){if(Object.keys(S.assignments).length>0){if(!confirm('確定要重新匯入？未儲存的排班資料將遺失。'))return;}showStep(['stepImport']);});
  byId('btnReconfig').addEventListener('click',function(){if(Object.keys(S.assignments).length>0){if(!confirm('確定要重新設定？未儲存的排班資料將遺失。'))return;}showStep(['stepImport','stepConfig']);});
  byId('btnUploadLastWeek').addEventListener('click',function(){byId('lastWeekFile').click();});
  byId('lastWeekFile').addEventListener('change',function(){
    var file=byId('lastWeekFile').files[0];
    if(!file)return;
    if(file.size>10*1024*1024){showToast('檔案過大（上限 10MB）','error');byId('lastWeekFile').value='';return;}
    var reader=new FileReader();
    reader.onload=function(e){
      try{
        var wb=XLSX.read(e.target.result);
        var ids=parseLastWeekExcel(wb);
        if(ids.length===0)throw new Error('未找到任何學員');
        var prevWeek=S.weekNum-1;
        S.history[prevWeek]={ids:ids};
        var st=byId('lastWeekStatus');
        st.textContent='已載入：'+ids.length+' 人（第'+prevWeek+'週）';
        st.className='lastweek-status loaded';
        updateConfigSummary();
        showToast('上週排班表已載入（'+ids.length+' 人）');
      }catch(err){
        byId('lastWeekStatus').textContent='載入失敗：'+err.message;
        showToast('載入失敗：'+err.message,'error');
      }
      byId('lastWeekFile').value='';
    };
    reader.readAsArrayBuffer(file);
  });
  window.addEventListener('beforeunload',function(e){
    if(S.assignments&&Object.keys(S.assignments).length>0&&!S.savedWeeks.has(S.weekNum)){
      e.preventDefault();e.returnValue='您有未儲存的排班資料，確定要離開嗎？';
    }
  });
  byId('semesterStart').addEventListener('change',updateWeekDates);
  byId('btnUndo').addEventListener('click',function(){undoLastAssign();});
  byId('btnExport').addEventListener('click',exportExcel);
  byId('btnSave').addEventListener('click',saveWeek);
  byId('btnBackupExport').addEventListener('click',exportBackup);
  byId('btnBackupImport').addEventListener('click',function(){byId('backupFile').click();});
  byId('backupFile').addEventListener('change',function(){if(this.files[0])importBackup(this.files[0]);this.value='';});
  updateBackupStatus();
  byId('pickerClose').addEventListener('click',closePicker);
  byId('pickerModal').addEventListener('click',function(e){if(e.target===byId('pickerModal'))closePicker();});
  /* === Picker 鍵盤導航 + Focus Trap === */
  document.addEventListener('keydown',function(e){
    var modalOpen=byId('pickerModal').classList.contains('show');
    if(e.key==='Escape'&&modalOpen){closePicker();return;}
    if((e.ctrlKey||e.metaKey)&&e.key==='z'&&!e.target.matches('input,textarea')){e.preventDefault();undoLastAssign();return;}
    if(!modalOpen)return;
    var list=byId('pickerList');
    var items=list.querySelectorAll('.pk-item');
    /* 方向鍵上下：在 picker 清單中移動高亮 */
    if(e.key==='ArrowDown'){
      e.preventDefault();
      if(items.length===0)return;
      pickerHighlightIdx=Math.min(pickerHighlightIdx+1,items.length-1);
      updatePickerHighlight();
      return;
    }
    if(e.key==='ArrowUp'){
      e.preventDefault();
      if(items.length===0)return;
      pickerHighlightIdx=Math.max(pickerHighlightIdx-1,0);
      updatePickerHighlight();
      return;
    }
    /* Enter 鍵：選取高亮項目 */
    if(e.key==='Enter'&&pickerHighlightIdx>=0&&items.length>0){
      e.preventDefault();
      if(items[pickerHighlightIdx])items[pickerHighlightIdx].click();
      return;
    }
    /* Focus Trap：Tab 鍵在 modal 內循環 */
    if(e.key==='Tab'){
      var searchInput=byId('pickerSearch');
      var showAllCb=byId('pickerShowAll');
      var closeBtn=byId('pickerClose');
      /* 收集可聚焦元素：搜尋框 → 顯示全部核取方塊 → picker 清單項目 → 關閉按鈕 */
      var focusable=[searchInput,showAllCb];
      items.forEach(function(el){focusable.push(el);});
      focusable.push(closeBtn);
      /* 過濾隱藏元素 */
      focusable=focusable.filter(function(el){return el&&el.offsetParent!==null;});
      if(focusable.length===0)return;
      var currentIdx=focusable.indexOf(document.activeElement);
      if(e.shiftKey){
        /* Shift+Tab：往前循環 */
        e.preventDefault();
        var prevIdx=currentIdx<=0?focusable.length-1:currentIdx-1;
        focusable[prevIdx].focus();
      }else{
        /* Tab：往後循環 */
        e.preventDefault();
        var nextIdx=currentIdx>=focusable.length-1?0:currentIdx+1;
        focusable[nextIdx].focus();
      }
    }
  });
  /* 開發模式自測（URL 含 ?test 參數時自動執行） */
  if(location.search.indexOf('test')>=0){
    setTimeout(function(){
      var r=runSelfTests();
      showToast('自測：'+r.passed+' 通過 / '+r.failed+' 失敗',r.failed>0?'error':'info');
    },500);
  }
}
document.addEventListener('DOMContentLoaded',init);
