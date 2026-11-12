/* ===== Export Data (PPTX) ===== */
const EXP_FONT='FreesiaUPC';
const EXP_AREAS=['A','B','C','D','R'];
const EXP_TOPICS=[
  {id:'situation',label:'Situation (จำนวนบ่อ บ่อมีกุ้ง DOC/Size)'},
  {id:'stocking',label:'Stocking (ลงกุ้ง / จับกุ้ง รายเดือน)'},
  {id:'harvest',label:'Order Harvest (แผน vs จริง รายเดือน)'},
  {id:'performance',label:'Performance (%SR FCR ADG Size)'},
  {id:'prepare',label:'Prepare pond (สถานะ ตามแผน/ล่าช้า)'},
  {id:'financial',label:'Financial / P&L (ยอดขาย ต้นทุน กำไร)'},
  {id:'summary',label:'สรุปข้อสังเกต / บ่อที่ต้องติดตาม'}
];
const EXP_WTOPICS=[
  {id:'w_situation',label:'Situation · แผนผังบ่อ / สถานการณ์ฟาร์ม'},
  {id:'w_prepare',label:'สถานะการเตรียมบ่อ (ตามแผน/ล่าช้า)'},
  {id:'w_preparedetail',label:'สถานะการเตรียมบ่อ (รายละเอียด)'},
  {id:'w_perf',label:'ประสิทธิภาพประจำสัปดาห์ (สรุป)'},
  {id:'w_problem',label:'ประสิทธิภาพ · เฉพาะบ่อที่มีปัญหา'},
  {id:'w_perfdetail',label:'ประสิทธิภาพ · รายละเอียดทุกบ่อ'},
  {id:'w_wq',label:'คุณภาพน้ำเกินเกณฑ์'},
  {id:'w_wqdetail',label:'คุณภาพน้ำ · รายละเอียด'},
  {id:'w_tvc',label:'แบคทีเรียในน้ำ-ในตับ เกินเกณฑ์'},
  {id:'w_tvcdetail',label:'แบคทีเรีย · รายละเอียด'},
  {id:'w_issue',label:'สรุปปัญหาและแนวทางการแก้ไข'},
  {id:'w_stockplan',label:'แผนการลงกุ้ง (สรุป)'},
  {id:'w_stockdetail',label:'แผนการลงกุ้ง (รายละเอียด)'},
  {id:'w_harvest',label:'แผนการจับกุ้ง'},
  {id:'w_pnl',label:'ประมาณการและผลดำเนินการ'},
  {id:'w_loss',label:'สาเหตุการขาดทุน (ประสิทธิภาพการผลิต · No Stk & Damage · Period Cost)'}
];
const EXP_THEMES={green:{teal:'4AC5B1',dark:'1F7A6D',zebra:'EAF8F5'},navy:{teal:'2E75B6',dark:'002060',zebra:'EEF3FA'}};
const EXP_ST={theme:'green',area:'ALL',year:new Date().getFullYear(),wmonths:[new Date().getMonth()+1],style:'exec',months:[1,2,3,4,5,6,7,8,9,10,11,12],topics:['situation','stocking','harvest','performance','prepare','financial','summary'],wtopics:EXP_WTOPICS.map(t=>t.id)};
const EXP_TGT={sr:85,fcr:1.6,harv:90}; /* เกณฑ์อ้างอิง: %SR ขั้นต่ำ, FCR สูงสุด, จับจริง/แผน ขั้นต่ำ(%) */
const EXP_OK='16A34A',EXP_WARN='F59E0B',EXP_BAD='DC2626';
const EXP_BG={};
const EXP_PPTX_URL='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const EXP_C={teal:'4AC5B1',dark:'1F7A6D',yellow:'FFC000',ink:'1F2937',grey:'6B7280',zebra:'EAF8F5',line:'D1D5DB',red:'DC2626'};
const EXP_MONTH_COL=[1,2,3,5,6,7,9,10,11,13,14,15];

/* Size เฉลี่ยที่ถูกต้อง = จำนวนตัวรวม ÷ Biomass รวม (จำนวนตัว = Biomass × Size ต่อบ่อ) */
function expWSize(rows){ let pcs=0,bio=0; rows.forEach(r=>{ const b=num(r[10]),s=num(r[9]); if(!isNaN(b)&&!isNaN(s)&&b>0&&s>0){ pcs+=b*s; bio+=b; } }); return bio?pcs/bio:NaN; }
function expN(x){ const n=num(x); return isNaN(n)?0:n; }
function expMonthOf(s){ const m=String(s||'').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m?{d:+m[1],m:+m[2],y:+m[3]}:null; }
function expAreaList(area){ return area==='ALL'?EXP_AREAS.slice():EXP_AREAS.filter(a=>String(area).split(',').indexOf(a)>=0); }
function expAreaName(area){ return area==='ALL'?'ทั้งฟาร์ม':'พื้นที่ '+expAreaList(area).join(', '); }
function expToggleArea(a){
  if(a==='ALL'){ EXP_ST.area='ALL'; return; }
  let cur=EXP_ST.area==='ALL'?EXP_AREAS.slice():expAreaList(EXP_ST.area);
  const i=cur.indexOf(a); if(i>=0) cur.splice(i,1); else cur.push(a);
  cur=EXP_AREAS.filter(x=>cur.indexOf(x)>=0);
  EXP_ST.area=(cur.length===0||cur.length===EXP_AREAS.length)?'ALL':cur.join(',');
}

async function expLoadData(){
  const [stk,oh,wk,fin,pr,dy]=await Promise.all([getTab('stocking'),getTab('orderharvest'),getTab('weekly'),getTab('financial'),getTab('prepare'),getTab('daily')]);
  let sit=null; try{ sit=await getTab('situation'); }catch(e){}
  return {stk:stk.values,oh:oh.values,wk:wk.values,fin:fin.values,pr:pr.values,dy:dy.values,sit:sit};
}

function expCompute(raw,st){
  const areas=expAreaList(st.area), ms=st.months.slice().sort((a,b)=>a-b);
  const out={};
  /* stocking summary block (cols J..S) */
  const sRows=ms.map(m=>{
    let rec={m,prev:0,pond:0,mil:0,harv:0,end:0,main:0};
    raw.stk.slice(1).forEach(r=>{
      if(areas.indexOf(String(r[10]).trim())<0||+r[17]!==st.year||+r[9]!==m)return;
      rec.prev+=expN(r[11]);rec.pond+=expN(r[12]);rec.mil+=expN(r[13]);rec.harv+=expN(r[14]);rec.end+=expN(r[15]);rec.main+=expN(r[18]);
    });
    rec.pct=rec.main?rec.end/rec.main*100:0; return rec;
  });
  out.stocking=sRows;
  /* order harvest */
  const oh=ms.map(m=>({m,plan:0,act:0,cnt:0}));
  raw.oh.slice(1).forEach(r=>{
    const mod=String(r[10]).trim(); if(areas.indexOf(mod)<0)return;
    const p=expMonthOf(r[12]); if(p&&p.y===st.year){ const o=oh.find(x=>x.m===p.m); if(o){o.plan+=expN(r[19]);o.cnt++;} }
    const a=expMonthOf(r[21]); if(a&&a.y===st.year){ const o=oh.find(x=>x.m===a.m); if(o){o.act+=expN(r[26]);} }
  });
  out.harvest=oh;
  /* weekly */
  const wk=raw.wk.slice(1).filter(r=>areas.indexOf(String(r[0]).trim())>=0&&String(r[1]).trim()&&!(typeof CANCEL_SET!=='undefined'&&CANCEL_SET.has(String(r[1]).trim())));
  out.pondsAll=wk;
  const avg=v=>v.length?v.reduce((a,x)=>a+x,0)/v.length:NaN;
  const cnt=r=>String(r[1]).trim();
  out.situation=areas.map(a=>{
    const rows=wk.filter(r=>String(r[0]).trim()===a);
    const cul=rows.filter(r=>/culture/i.test(r[4])), cl=rows.filter(r=>/close/i.test(r[4]));
    const g=(rs,i)=>rs.map(r=>num(r[i])).filter(x=>!isNaN(x)&&x>0);
    return {area:a,total:rows.length,cul:cul.length,close:cl.length,bio:cul.reduce((x,r)=>x+expN(r[10]),0),doc:avg(g(cul,7)),size:expWSize(cul),sr:avg(g(cul,13)),fcr:avg(g(cul,15)),adg:avg(g(cul,12))};
  });
  const isCx=r=>typeof CANCEL_SET!=='undefined'&&CANCEL_SET.has(String(r[1]).trim());
  out.prep=raw.pr.slice(1).filter(r=>areas.indexOf(String(r[0]).trim())>=0&&String(r[1]).trim()&&/close/i.test(r[4])&&!isCx(r));
  out.prepArea=areas.map(a=>{ const rs=out.prep.filter(r=>String(r[0]).trim()===a); return {area:a,total:rs.length,ok:rs.filter(r=>/ตามแผน/.test(r[10])).length,late:rs.filter(r=>/ล่าช้า/.test(r[10])).length,na:rs.filter(r=>!/ตามแผน|ล่าช้า/.test(r[10]||'')).length}; });
  out.culture=wk.filter(r=>/culture/i.test(r[4]));
  out.close=wk.filter(r=>/close/i.test(r[4]));
  /* financial */
  const starts={A:0,B:54,C:108,D:162,R:216};
  const labels=['Weight','Sales','Production Cost','GrossProfit','Admin','Interest','NetProfit'];
  out.financial=ms.map(m=>{
    const rec={m}; labels.forEach(l=>rec[l]=0); ['NoStk','Period','Direct','Fixed','Variable','AdminInt'].forEach(k=>rec[k]=0);
    const xmap={'nostk&damage':'NoStk','periodcost':'Period','directcost':'Direct','fixedcost':'Fixed','variablecost':'Variable','admin&interest':'AdminInt'};
    areas.forEach(a=>{
      const s=starts[a]; const col=EXP_MONTH_COL[m-1];
      for(let i=s+3;i<=s+9&&i<raw.fin.length;i++){ const lab=String(raw.fin[i][0]).trim(); if(labels.indexOf(lab)>=0) rec[lab]+=expN(raw.fin[i][col]); }
      /* Cost Structure (Baht) block: s+10 .. s+30 */
      for(let i=s+10;i<=s+30&&i<raw.fin.length;i++){ const k=xmap[String(raw.fin[i][0]).replace(/\s+/g,'').toLowerCase()]; if(k) rec[k]+=expN(raw.fin[i][col]); }
    });
    return rec;
  });
  return out;
}

/* ---------- slide helpers ---------- */
function expTitle(pres,slide,text){
  text=String(text).replace(/\s*·?\s*ทั้งฟาร์ม/g,'').replace(/\s+/g,' ').trim();
  const fs=44;
  slide.addText(text,{x:1.5,y:0.1,w:10.4,h:0.95,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:fs,bold:true,color:'FFFFFF',isTextBox:true,margin:0});
}
/* ย้ายเนื้อหาลงใต้แถบหัวข้อ: y<1.0 (หัวข้อ) คงเดิม, y≥1.0 บีบช่วง 1.05–7.0 → 1.45–7.0 */
function expShift(s){
  const K=(7.0-1.45)/(7.0-1.05), my=y=>(typeof y==='number'&&y>=1.0&&y<7.0)?(1.45+(y-1.05)*K):y;
  const fix=o=>{ if(o&&typeof o==='object'){ if(typeof o.y==='number'&&o.y>=1.0&&o.y<7.0){ o.y=Math.max(1.45,my(o.y)); } if(typeof o.h==='number'&&o.h>0.05)o.h=o.h*K; if(typeof o.rowH==='number')o.rowH=o.rowH*K; } return o; };
  ['addText','addShape','addChart','addTable','addImage'].forEach(fn=>{
    const orig=s[fn].bind(s);
    s[fn]=function(){ const a=Array.prototype.slice.call(arguments); const i=a.length-1; if(fn!=='addText'||!a[i]||a[i].y===undefined||a[i].y<7.0||true) a[i]=fix(a[i]); return orig.apply(null,a); };
  });
  return s;
}
function expNote(slide,text){
  slide.addText(text,{x:0.5,y:7.05,w:12.3,h:0.35,fontFace:EXP_FONT,fontSize:14,color:EXP_C.grey,isTextBox:true,margin:0});
}
function expTable(slide,head,rows,o){
  const fs=18, hdr=head.map(h=>({text:String(h),options:{bold:true,color:'FFFFFF',fill:{color:EXP_C.dark},align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:fs}}));
  const body=rows.map((r,i)=>r.map((c0,j)=>{ const ob=(c0&&typeof c0==='object')?c0:{t:c0}; const c=ob.t; return {text:String(c),options:{color:ob.color||(/ล่าช้า|▼/.test(String(c))?EXP_C.red:(/▲/.test(String(c))?EXP_OK:EXP_C.ink)),fill:{color:ob.fill||(i%2?'FFFFFF':EXP_C.zebra)},align:(j===0&&!o.leftFirst)?'center':(ob.align||(/^[\d.,\-–%\s]*$/.test(String(c))?'right':'left')),valign:'middle',fontFace:EXP_FONT,fontSize:fs,bold:!!(ob.bold||(o.boldLast&&i===rows.length-1))}}; }));
  slide.addTable([hdr].concat(body),{x:o.x,y:o.y,w:o.w,colW:o.colW,rowH:Math.max(o.rowH||0.34,0.38),border:{type:'solid',pt:0.5,color:EXP_C.line},margin:[0,0.06,0,0.06]});
}
function expChartBase(title){
  return {showTitle:true,title:title,titleFontFace:EXP_FONT,titleFontSize:20,titleColor:EXP_C.ink,
    showLegend:true,legendPos:'b',legendFontFace:EXP_FONT,legendFontSize:16,
    catAxisLabelFontFace:EXP_FONT,catAxisLabelFontSize:16,catAxisLabelColor:EXP_C.grey,
    valAxisLabelFontFace:EXP_FONT,valAxisLabelFontSize:14,valAxisLabelColor:EXP_C.grey,
    dataLabelFontFace:EXP_FONT,dataLabelFontSize:13,dataLabelColor:EXP_C.ink,
    valGridLine:{color:'E5E7EB',size:0.5},catGridLine:{style:'none'}};
}
function expFmt(n,d){ return fmt(n,d===undefined?0:d); }
function expPaged(pres,bgSlide,title,head,rows,colW,perPage,fs,note){
  perPage=Math.min(perPage,10); const pages=Math.max(1,Math.ceil(rows.length/perPage));
  for(let p=0;p<pages;p++){
    const s=bgSlide(); expTitle(pres,s,title+(pages>1?' ('+(p+1)+'/'+pages+')':''));
    const chunk=rows.slice(p*perPage,(p+1)*perPage);
    expTable(s,head,chunk,{x:0.5,y:1.3,w:12.3,colW:colW,fs:fs,rowH:0.33});
    if(note)expNote(s,note);
  }
}

/* ---------- Executive layout helpers ---------- */
function expHead(s,text){
  s.addText(text,{x:0.5,y:1.08,w:12.3,h:0.55,fontFace:EXP_FONT,fontSize:26,bold:true,color:EXP_C.dark,valign:'middle',isTextBox:true,margin:0});
}
function expKpiRow(pres,s,items,y){
  const n=items.length, gap=0.15, w=(12.3-gap*(n-1))/n;
  items.forEach((k,i)=>{
    const x=0.5+i*(w+gap), h=1.25, col=k.col||EXP_C.dark;
    s.addShape(pres.ShapeType.roundRect,{x:x,y:y,w:w,h:h,fill:{color:'FFFFFF'},line:{color:EXP_C.line,width:0.75},rectRadius:0.08});
    s.addShape(pres.ShapeType.rect,{x:x,y:y+0.12,w:0.08,h:h-0.24,fill:{color:col},line:{color:col,width:0}});
    s.addText(k.label,{x:x+0.1,y:y+0.06,w:w-0.2,h:0.34,align:'center',fontFace:EXP_FONT,fontSize:20,color:EXP_C.grey,isTextBox:true,margin:0});
    s.addText(k.val,{x:x+0.1,y:y+0.36,w:w-0.2,h:0.5,align:'center',fontFace:EXP_FONT,fontSize:n>=5?30:34,bold:true,color:col,isTextBox:true,margin:0});
    s.addText(k.sub||'',{x:x+0.1,y:y+0.86,w:w-0.2,h:0.32,align:'center',fontFace:EXP_FONT,fontSize:18,color:EXP_C.grey,isTextBox:true,margin:0});
  });
}
function expPanel(pres,s,x,y,w,h,fill,head,hc,items,fs,max){
  s.addShape(pres.ShapeType.roundRect,{x:x,y:y,w:w,h:h,fill:{color:fill},line:{color:fill,width:0},rectRadius:0.06});
  const arr=(items&&items.length?items:['ไม่พบประเด็นผิดปกติ']).slice(0,max||3);
  const runs=[{text:head,options:{bold:true,color:hc,fontSize:18,breakLine:true}}];
  arr.forEach((t,i)=>runs.push({text:t,options:{bullet:true,fontSize:fs||15,color:EXP_C.ink,breakLine:i<arr.length-1}}));
  s.addText(runs,{x:x+0.18,y:y+0.08,w:w-0.36,h:h-0.16,fontFace:EXP_FONT,valign:'top',paraSpaceAfter:3,isTextBox:true,margin:0});
}
function expIssues(pres,s,y,h,issues,actions){
  expPanel(pres,s,0.5,y,6.05,h,'FEF2F2','▲ ปัญหา / ประเด็นที่พบ',EXP_BAD,issues,15,3);
  expPanel(pres,s,6.75,y,6.05,h,'F0FDF4','✔ แนวทางแก้ไข / ข้อเสนอ','15803D',actions,15,3);
}
function expExecSlides(pres,bg,st,d,ms,fx){
  const T=id=>st.topics.indexOf(id)>=0, MT=MONTH_TH, an=expAreaName(st.area), labs=ms.map(m=>MT[m-1]);
  const sum=(a,f)=>a.reduce((x,r)=>x+f(r),0);
  const cul=d.culture;
  const avgOf=(rs,i)=>{ const v=rs.map(r=>num(r[i])).filter(x=>!isNaN(x)&&x>0); return v.length?v.reduce((a,x)=>a+x,0)/v.length:NaN; };
  const sit={t:sum(d.situation,r=>r.total),c:sum(d.situation,r=>r.cul),x:sum(d.situation,r=>r.close),b:sum(d.situation,r=>r.bio)};
  const sr=avgOf(cul,13), fcr=avgOf(cul,15), adg=avgOf(cul,12), size=expWSize(cul), doc=avgOf(cul,7);
  const hp=sum(d.harvest,r=>r.plan), ha=sum(d.harvest,r=>r.act), hpct=hp>0?ha/hp*100:NaN;
  const stk={pond:sum(d.stocking,r=>r.pond),mil:sum(d.stocking,r=>r.mil),harv:sum(d.stocking,r=>r.harv)};
  const lastS=d.stocking.slice().reverse().find(r=>r.end>0)||d.stocking[d.stocking.length-1]||{end:0,main:0,pct:0,m:ms[ms.length-1]};
  const pr={t:sum(d.prepArea,r=>r.total),ok:sum(d.prepArea,r=>r.ok),late:sum(d.prepArea,r=>r.late),na:sum(d.prepArea,r=>r.na)};
  const fin={W:sum(d.financial,r=>r.Weight),S:sum(d.financial,r=>r.Sales),C:sum(d.financial,r=>Math.abs(r['Production Cost'])),G:sum(d.financial,r=>r.GrossProfit),N:sum(d.financial,r=>r.NetProfit)};
  const srCol=isNaN(sr)?EXP_C.grey:(sr>=EXP_TGT.sr?EXP_OK:(sr>=EXP_TGT.sr-5?EXP_WARN:EXP_BAD));
  const fcrCol=isNaN(fcr)?EXP_C.grey:(fcr<=EXP_TGT.fcr?EXP_OK:(fcr<=EXP_TGT.fcr+0.2?EXP_WARN:EXP_BAD));
  const hCol=isNaN(hpct)?EXP_C.grey:(hpct>=100?EXP_OK:(hpct>=EXP_TGT.harv?EXP_WARN:EXP_BAD));
  const latePct=pr.t?pr.late/pr.t*100:0, lateCol=pr.late===0?EXP_OK:(latePct<=20?EXP_WARN:EXP_BAD);
  const netCol=fin.N>=0?EXP_OK:EXP_BAD;
  const lowSR=cul.filter(r=>{const x=num(r[13]);return !isNaN(x)&&x>0&&x<EXP_TGT.sr;}).sort((a,b)=>num(a[13])-num(b[13]));
  const hiFCR=cul.filter(r=>{const x=num(r[15]);return !isNaN(x)&&x>EXP_TGT.fcr;}).sort((a,b)=>num(b[15])-num(a[15]));
  const lateL=d.prep.filter(r=>/ล่าช้า/.test(r[10])).sort((a,b)=>expN(b[5])-expN(a[5]));
  const belowM=d.harvest.filter(r=>r.plan>0&&r.act<r.plan*EXP_TGT.harv/100).map(r=>MT[r.m-1]);
  const negM=d.financial.filter(r=>r.NetProfit<0).map(r=>MT[r.m-1]);
  const names=(rs,n)=>rs.slice(0,n).map(r=>r[1]).join(', ')+(rs.length>n?' ฯลฯ':'');
  const mb=v=>fmt(v/1000,2);
  const pct0=(a,b)=>b>0?fmt(a/b*100,0)+'%':'–';

  /* ===== Executive Summary ===== */
  {
    const s=bg(); expTitle(pres,s,'Executive Summary · '+an+' ปี '+st.year);
    expHead(s,'ภาพรวมผลการดำเนินงาน เทียบแผน/เกณฑ์ และประเด็นที่ต้องติดตาม');
    expKpiRow(pres,s,[
      {label:'บ่อมีกุ้ง',val:sit.c+' / '+sit.t,sub:'ใช้งาน '+pct0(sit.c,sit.t)+' · ปิด '+sit.x+' บ่อ',col:EXP_C.dark},
      {label:'จับจริง vs แผน',val:isNaN(hpct)?'–':fmt(hpct,0)+'%',sub:fmt(ha,1)+' / '+fmt(hp,1)+' ตัน',col:hCol},
      {label:'%SR เฉลี่ย',val:isNaN(sr)?'–':fmt(sr,1)+'%',sub:'เกณฑ์ ≥ '+EXP_TGT.sr+'%',col:srCol},
      {label:'FCR เฉลี่ย',val:isNaN(fcr)?'–':fmt(fcr,2),sub:'เกณฑ์ ≤ '+EXP_TGT.fcr,col:fcrCol},
      {label:'Net Profit (ลบ.)',val:mb(fin.N),sub:'ยอดขาย '+mb(fin.S)+' ลบ.',col:netCol},
      {label:'เตรียมบ่อล่าช้า',val:pr.late+' บ่อ',sub:'จาก '+pr.t+' บ่อปิด ('+fmt(latePct,0)+'%)',col:lateCol}
    ],1.75);
    const facts=[
      'บ่อมีกุ้ง '+sit.c+' จาก '+sit.t+' บ่อ · ผลผลิตรวม '+fmt(sit.b/1000,1)+' ตัน',
      'ลงกุ้ง '+stk.pond+' บ่อ ('+fmt(stk.mil,1)+' ล้านตัว) · จับกุ้ง '+stk.harv+' บ่อ',
      'จับจริง '+fmt(ha,1)+' ตัน'+(hp>0?' เทียบแผน '+fmt(hp,1)+' ตัน ('+fmt(hpct,0)+'%)':' (ไม่มีแผนในช่วงที่เลือก)'),
      'Performance: %SR '+fx(sr,1)+'% · FCR '+fx(fcr,2)+' · ADG '+fx(adg,2)+' g/วัน',
      'ยอดขาย '+mb(fin.S)+' ลบ. · ต้นทุน '+mb(fin.C)+' ลบ. · Net Profit '+mb(fin.N)+' ลบ.'
    ];
    const watch=[];
    if(belowM.length) watch.push('จับกุ้งต่ำกว่าแผนเกิน '+(100-EXP_TGT.harv)+'%: '+belowM.join(' '));
    if(lowSR.length) watch.push('%SR ต่ำกว่าเกณฑ์ '+lowSR.length+' บ่อ: '+names(lowSR,4));
    if(hiFCR.length) watch.push('FCR สูงกว่าเกณฑ์ '+hiFCR.length+' บ่อ: '+names(hiFCR,4));
    if(pr.late) watch.push('เตรียมบ่อล่าช้า '+pr.late+' บ่อ: '+names(lateL,4));
    if(negM.length) watch.push('เดือนที่ขาดทุน: '+negM.join(' '));
    expPanel(pres,s,0.5,3.2,6.05,3.75,'EAF8F5','ผลการดำเนินงานที่สำคัญ',EXP_C.dark,facts,19,5);
    expPanel(pres,s,6.75,3.2,6.05,3.75,'FFF7ED','เรื่องที่ต้องติดตาม / ตัดสินใจ','C2410C',watch.length?watch:['ไม่พบประเด็นเกินเกณฑ์ในช่วงที่เลือก'],19,5);
    expNote(s,'เกณฑ์อ้างอิง: %SR ≥ '+EXP_TGT.sr+'% · FCR ≤ '+EXP_TGT.fcr+' · จับจริง ≥ '+EXP_TGT.harv+'% ของแผน (ปรับได้)');
  }
  /* ===== 1 Situation ===== */
  if(T('situation')){
    const s=bg(); expTitle(pres,s,'1 · Situation · สถานการณ์ฟาร์ม '+an);
    expHead(s,'บ่อมีกุ้ง '+sit.c+' จาก '+sit.t+' บ่อ ('+pct0(sit.c,sit.t)+') · ปิด/เตรียมบ่อ '+sit.x+' บ่อ');
    expKpiRow(pres,s,[
      {label:'จำนวนบ่อทั้งหมด',val:String(sit.t),sub:'ไม่รวมบ่อยกเลิก'},
      {label:'บ่อมีกุ้ง (Culture)',val:String(sit.c),sub:'ใช้งาน '+pct0(sit.c,sit.t),col:EXP_OK},
      {label:'บ่อปิด / เตรียมบ่อ',val:String(sit.x),sub:'ยังไม่สร้างรายได้',col:sit.x/Math.max(sit.t,1)>0.3?EXP_WARN:EXP_C.dark},
      {label:'DOC เฉลี่ย · Size เฉลี่ย',val:fx(doc,0)+' · '+fx(size,0),sub:'ผลผลิตรวม '+fmt(sit.b/1000,1)+' ตัน'}
    ],1.7);
    const rows=d.situation.map(r=>[r.area,r.total,r.cul,r.close,pct0(r.cul,r.total),fx(r.doc,0),fx(r.size,0)]);
    if(d.situation.length>1) rows.push(['รวม',sit.t,sit.c,sit.x,pct0(sit.c,sit.t),fx(doc,0),fx(size,0)]);
    expTable(s,['พื้นที่','บ่อทั้งหมด','บ่อมีกุ้ง','บ่อปิด','% ใช้งาน','DOC เฉลี่ย','Size เฉลี่ย'],rows,{x:0.5,y:3.1,w:7.2,colW:[0.9,1.1,1.1,1.0,1.1,1.0,1.0],fs:15,rowH:0.33,boldLast:d.situation.length>1});
    s.addChart(pres.charts.BAR,[{name:'บ่อมีกุ้ง',labels:d.situation.map(r=>r.area),values:d.situation.map(r=>r.cul)},{name:'บ่อปิด',labels:d.situation.map(r=>r.area),values:d.situation.map(r=>r.close)}],
      Object.assign(expChartBase('จำนวนบ่อ แยกตามพื้นที่'),{x:7.9,y:3.0,w:5.0,h:2.6,barDir:'col',barGrouping:'stacked',chartColors:[EXP_C.teal,'F59E0B'],showValue:true,dataLabelPosition:'ctr'}));
    const low=d.situation.filter(r=>r.total>0).sort((a,b)=>a.cul/a.total-b.cul/b.total)[0];
    const iss=[]; if(sit.t&&sit.x/sit.t>0.3) iss.push('บ่อปิด/เตรียมบ่อ '+sit.x+' บ่อ ('+pct0(sit.x,sit.t)+') ยังไม่สร้างรายได้');
    if(low&&d.situation.length>1) iss.push('พื้นที่ '+low.area+' ใช้งานต่ำสุด '+pct0(low.cul,low.total)+' ('+low.cul+'/'+low.total+' บ่อ)');
    if(pr.late) iss.push('มีบ่อเตรียมล่าช้า '+pr.late+' บ่อ ทำให้ลงกุ้งช้ากว่าแผน');
    expIssues(pres,s,5.7,1.35,iss,['เร่งเตรียมบ่อปิดให้พร้อมลงกุ้งตามแผน เริ่มจากพื้นที่ใช้งานต่ำ','จัดลำดับความสำคัญบ่อที่ใกล้เสร็จขั้นตอนเตรียมบ่อ','ติดตามสถานะบ่อทุกสัปดาห์ในที่ประชุมผลิต']);
    expNote(s,'ข้อมูลสถานะบ่อล่าสุดจากหน้า Weekly (ไม่ขึ้นกับตัวกรองเดือน)');
  }
  /* ===== 2 Stocking ===== */
  if(T('stocking')){
    const s=bg(); expTitle(pres,s,'2 · Stocking · ลงกุ้ง / จับกุ้ง '+an+' ปี '+st.year);
    expHead(s,'ลงกุ้ง '+stk.pond+' บ่อ ('+fmt(stk.mil,1)+' ล้านตัว) · จับกุ้ง '+stk.harv+' บ่อ · สิ้นเดือน '+MT[lastS.m-1]+' มีกุ้ง '+fmt(lastS.end,0)+' บ่อ');
    expKpiRow(pres,s,[
      {label:'ลงกุ้ง (บ่อ)',val:fmt(stk.pond,0),sub:fmt(stk.mil,2)+' ล้านตัว'},
      {label:'จับกุ้ง (บ่อ)',val:fmt(stk.harv,0),sub:'ในช่วงที่เลือก'},
      {label:'มีกุ้งสิ้นเดือน '+MT[lastS.m-1],val:fmt(lastS.end,0)+' บ่อ',sub:'จากบ่อหลัก '+fmt(lastS.main,0)+' บ่อ'},
      {label:'% บ่อมีกุ้ง สิ้นเดือน',val:fmt(lastS.pct,0)+'%',sub:'เทียบบ่อหลัก',col:lastS.pct>=80?EXP_OK:(lastS.pct>=60?EXP_WARN:EXP_BAD)}
    ],1.7);
    s.addChart(pres.charts.BAR,[{name:'ลงกุ้ง (บ่อ)',labels:labs,values:d.stocking.map(r=>r.pond)},{name:'จับกุ้ง (บ่อ)',labels:labs,values:d.stocking.map(r=>r.harv)},{name:'มีกุ้งสิ้นเดือน (บ่อ)',labels:labs,values:d.stocking.map(r=>r.end)}],
      Object.assign(expChartBase('ลงกุ้ง / จับกุ้ง / มีกุ้งสิ้นเดือน (บ่อ)'),{x:0.5,y:3.0,w:12.3,h:2.65,barDir:'col',barGrouping:'clustered',chartColors:[EXP_C.teal,'F59E0B',EXP_C.dark],showValue:true,dataLabelPosition:'outEnd',dataLabelFontSize:11}));
    const emptyM=d.stocking.filter(r=>r.pond===0&&r.m<=lastS.m).map(r=>MT[r.m-1]);
    const iss=[]; if(lastS.pct<80) iss.push('บ่อมีกุ้งสิ้นเดือนเพียง '+fmt(lastS.pct,0)+'% ของบ่อหลัก ยังมีกำลังการผลิตว่าง');
    if(emptyM.length) iss.push('เดือนที่ไม่มีการลงกุ้ง: '+emptyM.join(' '));
    if(stk.harv>stk.pond) iss.push('จับกุ้ง ('+stk.harv+' บ่อ) มากกว่าลงกุ้ง ('+stk.pond+' บ่อ) บ่อมีกุ้งลดลง');
    expIssues(pres,s,5.75,1.3,iss,['วางแผนลงกุ้งทดแทนบ่อที่จับ เพื่อรักษาระดับบ่อมีกุ้ง','ประสานโรงเพาะ/ลูกกุ้งล่วงหน้าตามแผนลงกุ้งรายเดือน','ทบทวนแผนลงกุ้งรายสัปดาห์ร่วมกับแผนเตรียมบ่อ']);
  }
  /* ===== 3 Order Harvest ===== */
  if(T('harvest')){
    const s=bg(); expTitle(pres,s,'3 · Order Harvest · แผนจับ vs จับจริง '+an);
    expHead(s,'จับจริง '+fmt(ha,1)+' ตัน'+(hp>0?' = '+fmt(hpct,0)+'% ของแผน '+fmt(hp,1)+' ตัน (ผลต่าง '+(ha-hp>=0?'+':'')+fmt(ha-hp,1)+' ตัน)':''));
    s.addChart(pres.charts.BAR,[{name:'แผนจับ (ตัน)',labels:labs,values:d.harvest.map(r=>+r.plan.toFixed(1))},{name:'จับจริง (ตัน)',labels:labs,values:d.harvest.map(r=>+r.act.toFixed(1))}],
      Object.assign(expChartBase('แผนจับ vs จับจริง (ตัน)'),{x:0.4,y:1.7,w:7.5,h:3.85,barDir:'col',barGrouping:'clustered',chartColors:['9CA3AF',EXP_C.teal],showValue:true,dataLabelPosition:'outEnd',dataLabelFontSize:11}));
    const hv=d.harvest.filter(r=>r.plan>0||r.act>0);
    const rows=hv.map(r=>{ const df=r.act-r.plan; return [MT[r.m-1],fmt(r.plan,1),fmt(r.act,1),(df>=0?'▲ +':'▼ ')+fmt(df,1),r.plan>0?fmt(r.act/r.plan*100,0)+'%':'–']; });
    rows.push(['รวม',fmt(hp,1),fmt(ha,1),(ha-hp>=0?'▲ +':'▼ ')+fmt(ha-hp,1),hp>0?fmt(hpct,0)+'%':'–']);
    expTable(s,['เดือน','แผน','จริง','ผลต่าง (ตัน)','จริง/แผน'],rows,{x:8.1,y:1.75,w:4.8,colW:[0.9,0.9,0.9,1.2,0.9],fs:13,rowH:Math.min(0.3,3.7/(rows.length+1)),boldLast:true});
    const iss=[]; if(belowM.length) iss.push('จับจริงต่ำกว่าแผนเกิน '+(100-EXP_TGT.harv)+'% ในเดือน '+belowM.join(' '));
    if(!isNaN(hpct)&&hpct<100) iss.push('ยอดจับสะสมต่ำกว่าแผน '+fmt(hp-ha,1)+' ตัน กระทบยอดขาย/กระแสเงินสด');
    const over=d.harvest.filter(r=>r.plan>0&&r.act>r.plan*1.1).map(r=>MT[r.m-1]); if(over.length) iss.push('เดือนที่จับเกินแผน >10%: '+over.join(' ')+' (ตรวจสอบการวางแผน)');
    expIssues(pres,s,5.7,1.35,iss,['ทบทวนแผนจับกับตลาด/ลูกค้า และปรับ Order ให้สอดคล้องกำลังผลิต','ติดตามบ่อที่ถึงขนาดจับ (Size/DOC) ล่วงหน้า 2 สัปดาห์','ชดเชยยอดจับที่ขาดในเดือนถัดไป และรายงานสาเหตุรายเดือน']);
    expNote(s,'แผน = วันที่แผนจับ (Order Harvest) · จริง = Ac.Date · หน่วยตัน');
  }
  /* ===== 4 Performance ===== */
  if(T('performance')){
    const s=bg(); expTitle(pres,s,'4 · Performance · %SR FCR ADG Size '+an);
    expHead(s,'%SR เฉลี่ย '+fx(sr,1)+'% (เกณฑ์ ≥ '+EXP_TGT.sr+'%) · FCR เฉลี่ย '+fx(fcr,2)+' (เกณฑ์ ≤ '+EXP_TGT.fcr+')');
    expKpiRow(pres,s,[
      {label:'%SR เฉลี่ย',val:isNaN(sr)?'–':fmt(sr,1)+'%',sub:'ต่ำกว่าเกณฑ์ '+lowSR.length+' บ่อ',col:srCol},
      {label:'FCR เฉลี่ย',val:isNaN(fcr)?'–':fmt(fcr,2),sub:'สูงกว่าเกณฑ์ '+hiFCR.length+' บ่อ',col:fcrCol},
      {label:'ADG เฉลี่ย (g/วัน)',val:fx(adg,2),sub:'บ่อมีกุ้ง '+sit.c+' บ่อ'},
      {label:'Size เฉลี่ย (ตัว/กก.)',val:fx(size,0),sub:'DOC เฉลี่ย '+fx(doc,0)+' วัน'}
    ],1.7);
    const rows=d.situation.filter(r=>r.cul>0).map(r=>[r.area,r.cul,fx(r.sr,1)+(isNaN(r.sr)?'':'%')+(isNaN(r.sr)?'':(r.sr>=EXP_TGT.sr?' ▲':' ▼')),fx(r.fcr,2)+(isNaN(r.fcr)?'':(r.fcr<=EXP_TGT.fcr?' ▲':' ▼')),fx(r.adg,2),fx(r.size,0)]);
    expTable(s,['พื้นที่','บ่อมีกุ้ง','%SR','FCR','ADG','Size'],rows.length?rows:[['–','–','–','–','–','–']],{x:0.5,y:3.1,w:6.0,colW:[0.9,1.0,1.2,1.1,0.9,0.9],fs:15,rowH:0.33});
    const w5=lowSR.slice(0,5).map(r=>[r[1],fmt(num(r[13]),1)+'%',r[15]||'–',r[7]||'–']);
    expTable(s,['บ่อ %SR ต่ำสุด','%SR','FCR','DOC'],w5.length?w5:[['ไม่มีบ่อต่ำกว่าเกณฑ์','','','']],{x:6.7,y:3.1,w:6.1,colW:[2.2,1.3,1.3,1.3],fs:15,rowH:0.33});
    const iss=[]; if(lowSR.length) iss.push('%SR ต่ำกว่า '+EXP_TGT.sr+'% จำนวน '+lowSR.length+' บ่อ: '+names(lowSR,5));
    if(hiFCR.length) iss.push('FCR สูงกว่า '+EXP_TGT.fcr+' จำนวน '+hiFCR.length+' บ่อ: '+names(hiFCR,5));
    expIssues(pres,s,5.7,1.35,iss,['ตรวจคุณภาพน้ำ/โรค (TVC, ALK, TAN, No2) บ่อที่ %SR ต่ำ และกำหนดแผนแก้ไขรายบ่อ','ทบทวนโปรแกรมอาหาร/เช็คยอ เพื่อลด FCR บ่อที่เกินเกณฑ์','นำวิธีจัดการของบ่อที่ผลดีที่สุดมาเป็นมาตรฐาน']);
    expNote(s,'ค่าเฉลี่ยจากบ่อสถานะ Culture (หน้า Weekly) · ▲ ผ่านเกณฑ์ ▼ ต่ำกว่าเกณฑ์');
  }
  /* ===== 5 Prepare pond ===== */
  if(T('prepare')){
    const s=bg(); expTitle(pres,s,'5 · Prepare pond · แผนเตรียมบ่อ '+an);
    expHead(s,'เตรียมบ่อล่าช้า '+pr.late+' จาก '+pr.t+' บ่อปิด ('+fmt(latePct,0)+'%) · ตามแผน '+pr.ok+' บ่อ');
    expKpiRow(pres,s,[
      {label:'บ่อปิดทั้งหมด',val:String(pr.t),sub:'อยู่ระหว่างเตรียมบ่อ'},
      {label:'ตามแผน',val:String(pr.ok),sub:pct0(pr.ok,pr.t),col:EXP_OK},
      {label:'ล่าช้า',val:String(pr.late),sub:pct0(pr.late,pr.t),col:lateCol},
      {label:'ไม่ระบุแผน',val:String(pr.na),sub:pct0(pr.na,pr.t),col:EXP_C.grey}
    ],1.7);
    s.addChart(pres.charts.BAR,[{name:'ตามแผน',labels:d.prepArea.map(r=>r.area),values:d.prepArea.map(r=>r.ok)},{name:'ล่าช้า',labels:d.prepArea.map(r=>r.area),values:d.prepArea.map(r=>r.late)},{name:'ไม่ระบุแผน',labels:d.prepArea.map(r=>r.area),values:d.prepArea.map(r=>r.na)}],
      Object.assign(expChartBase('สถานะเตรียมบ่อ แยกพื้นที่'),{x:0.4,y:3.0,w:5.2,h:2.65,barDir:'col',barGrouping:'stacked',chartColors:[EXP_OK,EXP_BAD,'9CA3AF'],showValue:true,dataLabelPosition:'ctr'}));
    const lr=lateL.slice(0,6).map(r=>[r[1],r[5]||'–',r[6]||'–',r[9]||'–',r[10]||'–']);
    expTable(s,['บ่อล่าช้า','Prepare (วัน)','ขั้นตอนปัจจุบัน','กำหนดเสร็จ','เทียบแผน'],lr.length?lr:[['ไม่มีบ่อล่าช้า','','','','']],{x:5.8,y:3.1,w:7.1,colW:[1.2,1.2,2.4,1.3,1.0],fs:14,rowH:0.33});
    const iss=[]; if(pr.late) iss.push('ล่าช้า '+pr.late+' บ่อ (บ่อที่ล่าช้านานสุด: '+names(lateL,3)+')');
    if(pr.na) iss.push('ไม่ระบุแผน '+pr.na+' บ่อ ติดตามผลไม่ได้');
    expIssues(pres,s,5.75,1.3,iss,['เร่งขั้นตอนที่ค้างของบ่อล่าช้า และกำหนดผู้รับผิดชอบรายบ่อ','บันทึกกำหนดเสร็จให้ครบทุกบ่อ เพื่อเทียบแผนได้','ทบทวนทรัพยากร (คน/เครื่องจักร) ในพื้นที่ที่ล่าช้าสะสม']);
    expNote(s,'บ่อสถานะ Close จากหน้า Prepare pond · เรียงบ่อล่าช้านานสุดก่อน');
  }
  /* ===== 6 Financial ===== */
  if(T('financial')){
    const s=bg(); expTitle(pres,s,'6 · Financial / P&L · '+an+' ปี '+st.year);
    const mg=fin.S?fin.N/fin.S*100:NaN;
    expHead(s,'ยอดขาย '+mb(fin.S)+' ลบ. · ต้นทุนผลิต '+mb(fin.C)+' ลบ. · Net Profit '+mb(fin.N)+' ลบ.'+(isNaN(mg)?'':' (margin '+fmt(mg,1)+'%)'));
    expKpiRow(pres,s,[
      {label:'ผลผลิตขาย (ตัน)',val:fmt(fin.W,1),sub:'ราคาเฉลี่ย '+(fin.W>0?fmt(fin.S/fin.W,0):'–')+' บาท/กก.'},
      {label:'ยอดขาย (ลบ.)',val:mb(fin.S),sub:'พันบาท '+fmt(fin.S,0)},
      {label:'Gross Profit (ลบ.)',val:mb(fin.G),sub:'GP margin '+(fin.S?fmt(fin.G/fin.S*100,1)+'%':'–'),col:fin.G>=0?EXP_OK:EXP_BAD},
      {label:'Net Profit (ลบ.)',val:mb(fin.N),sub:'margin '+(isNaN(mg)?'–':fmt(mg,1)+'%'),col:netCol}
    ],1.7);
    s.addChart(pres.charts.BAR,[{name:'ยอดขาย',labels:labs,values:d.financial.map(r=>r.Sales)},{name:'ต้นทุนผลิต',labels:labs,values:d.financial.map(r=>Math.abs(r['Production Cost']))},{name:'Net Profit',labels:labs,values:d.financial.map(r=>r.NetProfit)}],
      Object.assign(expChartBase('พันบาท'),{x:0.4,y:3.0,w:7.3,h:2.65,barDir:'col',barGrouping:'clustered',chartColors:[EXP_C.teal,'F59E0B',EXP_C.dark],showValue:false}));
    const fm=d.financial.filter(r=>r.Sales!==0||r.NetProfit!==0).slice(-6);
    const rows=fm.map(r=>{ const i=d.financial.indexOf(r), pv=i>0?d.financial[i-1].Sales:0; const ch=pv>0?(r.Sales-pv)/pv*100:NaN; return [MT[r.m-1],mb(r.Sales),mb(Math.abs(r['Production Cost'])),mb(r.NetProfit),isNaN(ch)?'–':(ch>=0?'▲ +':'▼ ')+fmt(ch,0)+'%']; });
    expTable(s,['เดือน','ยอดขาย','ต้นทุน','Net','ยอดขาย vs เดือนก่อน'],rows.length?rows:[['–','–','–','–','–']],{x:7.9,y:3.1,w:5.0,colW:[0.8,0.9,0.9,0.9,1.5],fs:14,rowH:0.33});
    const iss=[]; if(negM.length) iss.push('เดือนที่ขาดทุน (Net Profit < 0): '+negM.join(' '));
    if(!isNaN(mg)&&mg<10) iss.push('Net margin '+fmt(mg,1)+'% ต่ำ ต้นทุนผลิตสูงเมื่อเทียบราคาขาย');
    if(belowM.length) iss.push('ยอดจับต่ำกว่าแผนในเดือน '+belowM.join(' ')+' ส่งผลต่อยอดขาย');
    expIssues(pres,s,5.75,1.3,iss,['ควบคุม FCR/ต้นทุนอาหาร และต้นทุนผันแปรหลัก','จับกุ้งตามแผนและขายที่ Size/ราคาที่เหมาะสม','ติดตาม P&L Forecast รายสัปดาห์เทียบเป้า']);
    expNote(s,'หน่วย: ล้านบาท (ลบ.) ยกเว้นผลผลิต = ตัน และกราฟ = พันบาท · ข้อมูลตามหน้า Financial');
  }
  /* ===== 7 Action plan ===== */
  if(T('summary')){
    const s=bg(); expTitle(pres,s,'7 · สรุปปัญหา และแผนแก้ไข · '+an);
    expHead(s,'ประเด็นที่ต้องติดตาม พร้อมแนวทางแก้ไข (กรอกผู้รับผิดชอบ/กำหนดเสร็จในที่ประชุม)');
    const R=[];
    if(belowM.length) R.push(['จับกุ้งต่ำกว่าแผน ('+belowM.join(' ')+')','ยอดขายและกระแสเงินสดต่ำกว่าเป้า','ทบทวนแผนจับ/Order กับตลาด เร่งจับบ่อที่ถึงขนาด','ระบุ','ระบุ']);
    if(lowSR.length) R.push(['%SR ต่ำกว่า '+EXP_TGT.sr+'% '+lowSR.length+' บ่อ ('+names(lowSR,4)+')','ผลผลิตลด ต้นทุน/กก. สูงขึ้น','ตรวจคุณภาพน้ำ/โรค แก้ไขรายบ่อ ติดตามทุกสัปดาห์','ระบุ','ระบุ']);
    if(hiFCR.length) R.push(['FCR สูงกว่า '+EXP_TGT.fcr+' '+hiFCR.length+' บ่อ ('+names(hiFCR,4)+')','ต้นทุนอาหารสูง กำไรลด','ทบทวนโปรแกรมอาหาร เช็คยอ/ปริมาณให้อาหาร','ระบุ','ระบุ']);
    if(pr.late) R.push(['เตรียมบ่อล่าช้า '+pr.late+' บ่อ ('+names(lateL,4)+')','ลงกุ้งช้า กำลังผลิตว่าง','เร่งขั้นตอนที่ค้าง จัดทรัพยากรเพิ่ม','ระบุ','ระบุ']);
    if(negM.length) R.push(['ขาดทุนเดือน '+negM.join(' '),'กำไรสะสมลดลง','ควบคุมต้นทุน เพิ่มยอดจับ/ราคาขาย ติดตาม Forecast','ระบุ','ระบุ']);
    if(!R.length) R.push(['ไม่พบประเด็นเกินเกณฑ์ในช่วงที่เลือก','–','รักษามาตรฐานและติดตามต่อเนื่อง','–','–']);
    expTable(s,['ปัญหา','ผลกระทบ','แนวทางแก้ไข','ผู้รับผิดชอบ','กำหนดเสร็จ'],R.slice(0,6).map((r,i)=>r),{x:0.5,y:1.8,w:12.3,colW:[3.9,2.5,3.9,1.1,0.9],fs:16,rowH:0.78,leftFirst:true});
    expNote(s,'ประเด็นสร้างอัตโนมัติจากข้อมูล Dashboard เทียบเกณฑ์ %SR ≥ '+EXP_TGT.sr+'% · FCR ≤ '+EXP_TGT.fcr+' · จับจริง ≥ '+EXP_TGT.harv+'% ของแผน');
  }
}

/* ---------- Weekly report helpers ---------- */
function expSnap(key,cur){
  let arr=[]; try{arr=JSON.parse(localStorage.getItem('expWkSnap')||'[]');}catch(e){}
  const now=Date.now(); const prev=arr.filter(x=>x.k===key&&now-x.t>=5*864e5).sort((a,b)=>b.t-a.t)[0]||null;
  arr=arr.filter(x=>!(x.k===key&&now-x.t<864e5)); arr.push(Object.assign({t:now,k:key},cur)); arr=arr.slice(-60);
  try{localStorage.setItem('expWkSnap',JSON.stringify(arr));}catch(e){}
  return prev;
}
function expDelta(cur,prev,lowerBetter){
  if(prev===null||prev===undefined||isNaN(prev)) return 'ยังไม่มีข้อมูลสัปดาห์ก่อน (เริ่มบันทึกสัปดาห์นี้)';
  const df=cur-prev; if(df===0) return 'เท่าเดิม (สัปดาห์ก่อน '+prev+' บ่อ)';
  const better=lowerBetter?df<0:df>0; return (better?'▲ ดีขึ้น ':'▼ แย่ลง ')+Math.abs(df)+' บ่อ (สัปดาห์ก่อน '+prev+' บ่อ)';
}
const EXP_FILL={green:'DCFCE7',yellow:'FEF9C3',red:'FEE2E2'};
function expLet(x){ x=String(x==null?'':x).trim().toLowerCase(); return x==='green'||x==='yellow'||x==='red'?x:''; }
function expTvcLet(v,kind){ const n=Number(v); if(v===''||v==null||isNaN(n))return ''; if(kind==='water')return n<=999?'green':(n<=4999?'yellow':'red'); return n<=50000?'green':(n<=60000?'yellow':'red'); }
function expSci(v){ const n=Number(v); if(v===''||v==null||isNaN(n))return '–'; return n.toExponential(2).replace(/e([+-])(\d)$/,'E$10$2').replace(/e([+-])/,'E$1'); }
function expCell(t,let_,al){ return {t:t,fill:EXP_FILL[let_]||undefined,color:let_==='red'?EXP_BAD:EXP_C.ink,bold:let_==='red',align:al}; }
function expCard(pres,s,x,y,w,h){ s.addShape(pres.ShapeType.roundRect,{x:x,y:y,w:w,h:h,fill:{color:'FFFFFF'},line:{color:EXP_C.line,width:0.75},rectRadius:0.08}); }
function expAlert(pres,s,x,y,w,h,text,sub,bad){
  const c=bad?EXP_BAD:EXP_OK;
  s.addShape(pres.ShapeType.roundRect,{x:x,y:y,w:w,h:h,fill:{color:bad?'FEF2F2':'F0FDF4'},line:{color:c,width:2},rectRadius:0.1});
  s.addText(text,{x:x+0.2,y:y+0.05,w:w-0.4,h:h*0.55,fontFace:EXP_FONT,fontSize:22,bold:true,color:c,align:'center',valign:'middle',isTextBox:true,margin:0});
  if(sub) s.addText(sub,{x:x+0.2,y:y+h*0.58,w:w-0.4,h:h*0.38,fontFace:EXP_FONT,fontSize:16,color:EXP_C.grey,align:'center',valign:'top',isTextBox:true,margin:0});
}
function expDonut(pres,s,x,y,w,h,pass,total,color,label){
  const p=total?Math.round(pass/total*100):0;
  s.addChart(pres.charts.DOUGHNUT,[{name:label,labels:['ผ่าน','ไม่ผ่าน'],values:[pass,Math.max(total-pass,0)]}],{x:x,y:y,w:w,h:h,holeSize:68,chartColors:[color,'E5E7EB'],showLegend:false,showLabel:false,showValue:false,showPercent:false,showTitle:false,dataBorder:{pt:0,color:'FFFFFF'}});
  s.addText(p+'%',{x:x,y:y+h/2-0.3,w:w,h:0.6,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:30,bold:true,color:color,isTextBox:true,margin:0});
  s.addText(total?('ผ่านเกณฑ์ '+pass+' / '+total+' บ่อ'):'ไม่มีข้อมูล',{x:x-0.5,y:y+h-0.05,w:w+1,h:0.32,align:'center',fontFace:EXP_FONT,fontSize:15,color:EXP_C.grey,isTextBox:true,margin:0});
}
function expPagedW(pres,bg,title,head,rows,colW,perPage,fs,note,rowH){
  perPage=Math.min(perPage,10); const pages=Math.max(1,Math.ceil(rows.length/perPage));
  for(let p=0;p<pages;p++){
    const s=bg(); expTitle(pres,s,title+(pages>1?' ('+(p+1)+'/'+pages+')':''));
    expTable(s,head,rows.slice(p*perPage,(p+1)*perPage),{x:0.5,y:1.3,w:12.3,colW:colW,fs:fs,rowH:rowH||0.34});
    if(note)expNote(s,note);
  }
}

/* ---------- Weekly report slides (ตามเทมเพลตรายงานประจำสัปดาห์) ---------- */
function expWeeklySlides(pres,bg,st,raw,d,now){
  const T=id=>st.wtopics.indexOf(id)>=0, MT=MONTH_TH, an=expAreaName(st.area), areas=expAreaList(st.area);
  const isCx=r=>typeof CANCEL_SET!=='undefined'&&CANCEL_SET.has(String(r[1]).trim());
  const wk=d.pondsAll, cul=wk.filter(r=>/culture/i.test(r[4])), clo=wk.filter(r=>/close/i.test(r[4]));
  const docOf={}; wk.forEach(r=>{docOf[String(r[1]).trim()]=r[7];});
  const dy=raw.dy.slice(2).filter(r=>areas.indexOf(String(r[0]).trim())>=0&&String(r[1]).trim()&&!isCx(r)&&/culture/i.test(r[4]));
  const fx=(n,dd)=>isNaN(n)?'–':expFmt(n,dd);
  const nm=x=>{const v=num(x);return isNaN(v)?NaN:v;};
  const nowD=now.getDate(), nowM=now.getMonth()+1, nowY=now.getFullYear();
  const nTot=wk.length, nCul=cul.length, nClo=clo.length, util=nTot?nCul/nTot*100:0;
  const bin=(vals,edges)=>edges.map(e=>vals.filter(e.f).length);
  /* ---- metrics ---- */
  const gainLow=cul.filter(r=>{const g=nm(r[11]);return !isNaN(g)&&g<3;}).sort((a,b)=>expN(b[7])-expN(a[7]));
  const srLow=cul.filter(r=>{const x=nm(r[14]);return !isNaN(x)&&x>0&&x<70;});
  const probSet=cul.filter(r=>{const g=nm(r[11]),x=nm(r[14]);return (!isNaN(g)&&g<3)||(!isNaN(x)&&x>0&&x<70);}).sort((a,b)=>expN(b[7])-expN(a[7]));
  const wqKeys=[{k:'pH',i:20,v:9,dd:2},{k:'T.Alk',i:21,v:10,dd:0},{k:'TAN',i:22,v:11,dd:2},{k:'NO₂-N',i:23,v:12,dd:2}];
  const wqRed=wqKeys.map(q=>dy.filter(r=>expLet(r[q.i])==='red'));
  const wqYel=wqKeys.map(q=>dy.filter(r=>expLet(r[q.i])==='yellow'));
  const wqBad=dy.filter(r=>wqKeys.some(q=>expLet(r[q.i])==='red'));
  const mortRed=dy.filter(r=>expLet(r[24])==='red');
  const tvcW=dy.map(r=>({p:String(r[1]).trim(),w:r[17],l:r[18],cw:expTvcLet(r[17],'water'),cl:expTvcLet(r[18],'liver')}));
  const tvcWdata=tvcW.filter(x=>x.cw), tvcLdata=tvcW.filter(x=>x.cl);
  const tvcWpass=tvcWdata.filter(x=>x.cw==='green').length, tvcLpass=tvcLdata.filter(x=>x.cl==='green').length;
  const tvcWbad=tvcWdata.filter(x=>x.cw!=='green'), tvcLbad=tvcLdata.filter(x=>x.cl!=='green');
  /* ---- snapshot เทียบสัปดาห์ก่อน ---- */
  const prev=expSnap(st.area,{cul:nCul,gainLow:gainLow.length,srLow:srLow.length,wqBad:wqBad.length,tvcW:tvcWbad.length,tvcL:tvcLbad.length});
  const pv=k=>prev?prev[k]:null;

  const STEP_LIST=['ฆ่าเชื้อบ่อด้วยโซดาไฟ (pH 11-13)','ฉีดล้าง ขัดล้างบ่อ','ซ่อมบำรุง อุปกรณ์บ่อเลี้ยง','ซินพื้นบ่อ ซ่อมแซมโครงสร้าง ทาสี','แช่ Subdrain ด่างทับทิม 300 ppm (ตรวจสอบรอยรั่ว)','ฆ่าเชื้อคลอรีน Acidified Chlorine 300 ppm','ตรวจ Swab test (PCR) รอผล PCR Swab test','เติมน้ำ พร้อมปรับคุณภาพน้ำ','ส่งตรวจ PCR น้ำก่อนลงกุ้งรอผล','พร้อมปล่อยกุ้ง'];
  const STEP_RULES=[[9,/พร้อมปล่อยกุ้ง/],[0,/โซดาไฟ|NaOH/],[4,/Subdrain|ด่างทับทิม|ค่างทับทิม|รอยรั่ว|Soak/i],[5,/คลอรีน|Chlorine/i],[6,/Swab/i],[8,/PCR/i],[7,/เติมน้ำ|ปรับคุณภาพน้ำ|ผลออก/],[1,/ฉีด|ขัดล้าง/],[3,/ซิน|ทาสี|ซ่อมแซมโครงสร้าง/],[2,/ซ่อมบำรุง|PM|จัดเตรียมอุปกรณ์/i]];
  const stepIdx=s=>{ for(const q of STEP_RULES){ if(q[1].test(s||'')) return q[0]; } return 99; };
  const polyType=p=>/^RW-/i.test(p)?'R':(/I\d+$/i.test(p)?'I':'G');
  const areaRank=r=>{ const i=EXP_AREAS.indexOf(String(r[0]).trim()); return i<0?99:i; };
  /* ===== 1 Situation (layout ตามหน้า Dashboard · Situation) ===== */
  if(T('w_situation')){
    const s=bg(); expTitle(pres,s,'สถานการณ์ฟาร์มร้อยเพชร'+(st.area==='ALL'?'':' · '+an));
    const lastS=d.stocking.find(r=>r.m===((st.wmonths&&st.wmonths.length)?Math.max.apply(null,st.wmonths):nowM))||d.stocking.slice().reverse().find(r=>r.end>0);
    const eomN=lastS?lastS.end:0, eomP=lastS?lastS.pct:0;
    const nInter=wk.filter(r=>/^inter/i.test(String(r[2]))).length, nGrow=nTot-nInter;
    const nCulI=cul.filter(r=>/^inter/i.test(String(r[2]))).length, nCulG=nCul-nCulI;
    expKpiRow(pres,s,[
      {label:'บ่อหลัก',val:String(nTot),sub:'บ่อ',col:EXP_C.dark},
      {label:'บ่อเลี้ยง (มีกุ้ง)',val:String(nCul),sub:'Inter '+nCulI+' · Growout '+nCulG,col:EXP_OK},
      {label:'เตรียมบ่อ',val:String(nClo),sub:'บ่อ',col:EXP_C.grey},
      {label:'% บ่อมีกุ้ง',val:fmt(util,0)+'%',sub:nCul+'/'+nTot+' บ่อ',col:EXP_WARN},
      {label:'ประมาณการ มีกุ้ง ณ สิ้นเดือน',val:fmt(eomN,0)+' บ่อ',sub:lastS?'('+fmt(eomP,0)+'%) เดือน'+MT[lastS.m-1]:'–',col:'7C3AED'}
    ],1.2);
    const colsC={culture:'CCFFFF',close:'B7B7B7',other:'FE99FD'};
    const PRE0=/^[A-D][1-6][GI][0-9]{2}$/;
    const SV=(raw.sit&&raw.sit.values)||[], SB=(raw.sit&&raw.sit.bg)||[];
    let minC=99,maxC=0,minR=99,maxR=0; for(let i=2;i<=21;i++) for(let c=0;c<=18;c++){ const t0=String((SV[i]&&SV[i][c])||'').trim(); if(PRE0.test(t0)){ minC=Math.min(minC,c); maxC=Math.max(maxC,c); minR=Math.min(minR,i); maxR=Math.max(maxR,i); } }
    if(minC>maxC){ minC=0; maxC=18; minR=2; maxR=21; }
    const gw=Math.min(4.1/(maxC-minC+4), 0.34), gh=Math.min(4.25/(maxR-minR+1), 0.34), gx0=0.4-minC*gw, gy0=2.6-minR*gh;
    const hexOf=c=>{ c=String(c||'').replace('#',''); return /^[0-9a-fA-F]{6}$/.test(c)?c.toUpperCase():'FFFFFF'; };
    let sitWater=0, drewAny=false;
    const PRE=PRE0;
    for(let i=2;i<=21;i++) for(let c=0;c<=18;c++){
      const t=String((SV[i]&&SV[i][c])||'').trim(); if(!t) continue;
      const x=gx0+c*gw, y=gy0+i*gh;
      if(PRE.test(t)){
        drewAny=true; const on=areas.indexOf(t[0])>=0; const fill=hexOf(SB[i]&&SB[i][c]);
        if(on&&fill==='FE99FD') sitWater++;
        s.addShape(pres.ShapeType.rect,{x:x,y:y,w:gw*0.94,h:gh*0.9,fill:{color:on?fill:'F3F4F6'},line:{color:'7F7F7F',width:0.4}});
        s.addText(t,{x:x,y:y,w:gw*0.94,h:gh*0.9,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:Math.max(4,Math.round(gw*17)),bold:true,color:'1F2937',isTextBox:true,margin:0,wrap:false,fit:'none'});
      } else if(/module/i.test(t)){
        const m=t.replace(/module/i,'').trim();
        const lx=Math.min(x,4.85-gw*3.2); s.addShape(pres.ShapeType.roundRect,{x:lx,y:y,w:gw*3.2,h:gh*1.3,fill:{color:'ECFDF5'},line:{color:'6EE7B7',width:0.75},rectRadius:0.05});
        s.addText('โมดูล '+m,{x:lx,y:y,w:gw*3.2,h:gh*1.3,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:10,bold:true,color:'1F2F6B',isTextBox:true,margin:0});
      }
    }
    /* บ่อ RW (New model) ตำแหน่งตามหน้า Dashboard */
    const rwPos={'RW-A501':[15,4],'RW-A201':[21,4],'RW-B301':[21,10],'RW-C501':[9,13],'RW-C502':[9,14],'RW-D201':[8,13],'RW-D202':[8,14],'RW-D203':[8,15]};
    Object.keys(rwPos).forEach(nm=>{ const p=wk.find(r=>String(r[1]).trim()===nm); if(!p) return; const stt=/culture/i.test(p[4])?'culture':(/close/i.test(p[4])?'close':'other'); const pp=rwPos[nm], x=gx0+pp[1]*gw, y=gy0+pp[0]*gh;
      s.addShape(pres.ShapeType.rect,{x:x,y:y,w:gw*0.94,h:gh*0.9,fill:{color:areas.indexOf('R')>=0?colsC[stt]:'F3F4F6'},line:{color:'0F766E',width:1}});
      s.addText(nm.replace('RW-',''),{x:x,y:y,w:gw*0.94,h:gh*0.9,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:Math.max(4,Math.round(gw*17)),bold:true,color:'1F2937',isTextBox:true,margin:0,wrap:false}); });
    const cnt=k=>wk.filter(r=>(k==='culture'?/culture/i.test(r[4]):k==='close'?/close/i.test(r[4]):!/culture|close/i.test(r[4]))).length;
    [['มีกุ้ง (Culture) '+cnt('culture')+' บ่อ','culture'],['เตรียมบ่อ '+cnt('close')+' บ่อ','close'],['พักน้ำ/อื่นๆ '+(sitWater||cnt('other'))+' บ่อ','other']].forEach((l,i)=>{
      s.addShape(pres.ShapeType.roundRect,{x:4.9+i*2.75,y:2.7,w:0.3,h:0.28,fill:{color:colsC[l[1]]},line:{color:EXP_C.line,width:0.5},rectRadius:0.04});
      s.addText(l[0],{x:5.28+i*2.75,y:2.62,w:2.4,h:0.45,fontFace:EXP_FONT,fontSize:18,color:EXP_C.ink,valign:'middle',isTextBox:true,margin:0});
    });
    const nOth=nTot-nCul-nClo;
    const mrows=areas.map(a=>{ const rs=wk.filter(r=>String(r[0]).trim()===a), c=rs.filter(r=>/culture/i.test(r[4])).length, cl=rs.filter(r=>/close/i.test(r[4])).length, o=rs.length-c-cl; return [{t:'โมดูล '+a,align:'center'},rs.length,c,rs.length?fmt(c/rs.length*100,0)+'%':'0%',cl,o?String(o):'']; });
    if(areas.length>1) mrows.push([{t:'รวม',align:'center',bold:true},{t:String(nTot),bold:true},{t:String(nCul),bold:true},{t:fmt(util,0)+'%',bold:true},{t:String(nClo),bold:true},{t:nOth?String(nOth):'',bold:true}]);
    expTable(s,['โมดูล','บ่อหลัก (บ่อ)','มีกุ้ง (บ่อ)','% มีกุ้ง','เตรียม (บ่อ)','อื่นๆ'],mrows,{x:4.9,y:3.3,w:8.0,colW:[1.5,1.6,1.4,1.2,1.4,0.9],fs:18,rowH:0.46});
    expNote(s,'ข้อมูลสถานะบ่อล่าสุดจากหน้า Weekly · ผลผลิตรวมในบ่อ '+fmt(cul.reduce((a,r)=>a+expN(r[10]),0)/1000,1)+' ตัน · DOC เฉลี่ย '+fx(cul.reduce((a,r)=>a+expN(r[7]),0)/Math.max(nCul,1),0)+' วัน');
  }
  /* ===== 1.5 สถานะการเตรียมบ่อ ===== */
  if(T('w_prepare')){
    const s=bg(); expTitle(pres,s,'สถานะการเตรียมบ่อ'+(st.area==='ALL'?'':' · '+an));
    const P=d.prep, pt=P.length;
    const okN=P.filter(r=>/ตามแผน/.test(r[10])).length, lateN=P.filter(r=>/ล่าช้า/.test(r[10])).length, naN=pt-okN-lateN;
    const pc=n=>pt?fmt(n/pt*100,0)+'%':'0%';
    expKpiRow(pres,s,[
      {label:'บ่อปิด / เตรียมบ่อ',val:String(pt),sub:'บ่อ',col:EXP_C.dark},
      {label:'ตามแผน',val:okN+' บ่อ',sub:pc(okN),col:EXP_OK},
      {label:'ล่าช้า',val:lateN+' บ่อ',sub:pc(lateN),col:lateN?EXP_BAD:EXP_OK},
      {label:'ไม่ระบุแผน',val:naN+' บ่อ',sub:pc(naN),col:EXP_C.grey}
    ],1.2);
    const steps={}; P.forEach(r=>{ const i=stepIdx(String(r[6]||'')); const k=i===99?'อื่นๆ / ไม่ระบุขั้นตอน':(i+1)+'. '+STEP_LIST[i]; (steps[k]=steps[k]||[]).push(r); });
    const keys=Object.keys(steps).sort((a,b)=>(parseInt(a)||99)-(parseInt(b)||99)).slice(0,9);
    const rows=keys.map(k=>{ const rs=steps[k], o=rs.filter(r=>/ตามแผน/.test(r[10])).length, l=rs.filter(r=>/ล่าช้า/.test(r[10])).length;
      const stt=l>0?{t:'ล่าช้า '+l+' บ่อ',color:EXP_BAD,bold:true,align:'center'}:{t:'ตามแผน',color:EXP_OK,bold:true,align:'center'};
      return [{t:k,align:'left'},rs.length,pc(rs.length),o,l,stt,{t:rs.slice(0,3).map(r=>String(r[1]).trim()).join(', ')+(rs.length>3?' ฯลฯ':''),align:'left'}]; });
    expTable(s,['ขั้นตอน','จำนวนบ่อ','% ของบ่อปิด','ตามแผน','ล่าช้า','สถานะ','บ่อ'],rows.length?rows:[['ไม่มีบ่อที่อยู่ระหว่างเตรียมบ่อ','','','','','','']],{x:0.5,y:2.65,w:12.3,colW:[3.9,1.1,1.4,1.0,1.0,1.7,2.2],fs:16,rowH:0.44,leftFirst:true});
    expNote(s,'บ่อสถานะ Close จากหน้า Prepare pond · แยกตามขั้นตอนปัจจุบัน · สถานะล่าช้า = มีอย่างน้อย 1 บ่อล่าช้าในขั้นตอนนั้น');
  }
  /* ===== 1.6 สถานะการเตรียมบ่อ (รายละเอียด) ===== */
  if(T('w_preparedetail')){
    const P=d.prep;
    const grp=STEP_LIST.map(()=>[]), oth=[]; P.forEach(r=>{ const i=stepIdx(String(r[6]||'')); (i===99?oth:grp[i]).push(r); });
    const rowsAll=STEP_LIST.map((nmS,i)=>({k:(i+1)+'.'+nmS,rs:grp[i]}));
    if(oth.length) rowsAll.push({k:'อื่นๆ / ไม่ระบุขั้นตอน',rs:oth});
    const mk=o=>{ const rs=o.rs.slice().sort((a,b)=>areaRank(a)-areaRank(b)||String(a[1]).localeCompare(String(b[1]))); const ok=rs.filter(r=>/ตามแผน/.test(r[10])).length, lt=rs.filter(r=>/ล่าช้า/.test(r[10])).length;
      const names=rs.map(r=>String(r[1]).trim()).join(', '), days=rs.map(r=>String(r[5]||'–').trim()).join(', ');
      const lines=Math.max(1,Math.ceil(names.length/38),Math.ceil(o.k.length/34),Math.ceil(days.length/16));
      return {lines:lines,row:[{t:o.k,align:'left'},rs.length,{t:names,align:'left'},{t:days,align:'center'},{t:rs.length?('ตามแผน '+ok+' · ล่าช้า '+lt):'',align:'left',color:lt?EXP_BAD:EXP_C.ink}]}; };
    const items=rowsAll.map(mk), pages=[]; let cur=[],h=0;
    items.forEach(it=>{ const ih=it.lines*0.3+0.1; if(cur.length&&h+ih>4.2){ pages.push(cur); cur=[]; h=0; } cur.push(it.row); h+=ih; }); if(cur.length) pages.push(cur);
    pages.forEach((chunk,pi)=>{ const s=bg(); expTitle(pres,s,'สถานะการเตรียมบ่อ (รายละเอียด)'+(pages.length>1?' ('+(pi+1)+'/'+pages.length+')':''));
      expTable(s,['ขั้นตอน (10 ขั้นตอน)','จำนวนบ่อ','ชื่อบ่อ','วันเตรียม','รายละเอียด'],chunk,{x:0.5,y:1.3,w:12.3,colW:[3.6,1.1,4.1,1.4,2.1],fs:18,rowH:0.4,leftFirst:true});
      if(pi===pages.length-1) expNote(s,'บ่อสถานะ Close จากหน้า Prepare pond · วันเตรียม = จำนวนวันเตรียมบ่อสะสม · เรียงตามพื้นที่ A,B,C,D,R'); });
  }
  /* ===== 2 ประสิทธิภาพ (สรุป) ===== */
  if(T('w_perf')){
    const s=bg(); expTitle(pres,s,'ประสิทธิภาพประจำสัปดาห์ · '+an);
    expCard(pres,s,0.5,1.25,6.05,2.75); s.addText('Utilization (พื้นที่การเลี้ยง)',{x:0.7,y:1.3,w:5.6,h:0.4,fontFace:EXP_FONT,fontSize:20,bold:true,color:EXP_C.ink,isTextBox:true,margin:0});
    expDonut(pres,s,2.3,1.75,2.4,1.9,nCul,nTot,EXP_OK,'Utilization');
    expCard(pres,s,6.75,1.25,6.05,2.75);
    const sz=cul.map(r=>nm(r[9])).filter(x=>!isNaN(x)&&x>0);
    const szb=[{n:'>100',f:x=>x>100},{n:'70-100',f:x=>x>70&&x<=100},{n:'50-70',f:x=>x>50&&x<=70},{n:'40-50',f:x=>x>40&&x<=50},{n:'30-40',f:x=>x>30&&x<=40},{n:'25-30',f:x=>x>25&&x<=30},{n:'20-25',f:x=>x>20&&x<=25},{n:'<20',f:x=>x<=20}];
    s.addChart(pres.charts.BAR,[{name:'จำนวนบ่อ',labels:szb.map(b=>b.n),values:szb.map(b=>sz.filter(b.f).length)}],Object.assign(expChartBase('ขนาดกุ้ง (Size ตัว/กก.) จำนวนบ่อ'),{x:6.85,y:1.3,w:5.85,h:2.65,barDir:'col',chartColors:[EXP_C.teal],showLegend:false,showValue:true,dataLabelPosition:'outEnd',titleFontSize:20}));
    expCard(pres,s,0.5,4.15,6.05,2.8);
    const dc=cul.map(r=>nm(r[7])).filter(x=>!isNaN(x));
    const dcb=[{n:'≤30 วัน',f:x=>x<=30},{n:'31-60 วัน',f:x=>x>30&&x<=60},{n:'61-90 วัน',f:x=>x>60&&x<=90},{n:'>90 วัน',f:x=>x>90}];
    s.addChart(pres.charts.BAR,[{name:'จำนวนบ่อ',labels:dcb.map(b=>b.n),values:dcb.map(b=>dc.filter(b.f).length)}],Object.assign(expChartBase('อายุการเลี้ยง (DOC)'),{x:0.6,y:4.2,w:5.85,h:2.7,barDir:'col',chartColors:['B8860B'],showLegend:false,showValue:true,dataLabelPosition:'outEnd',titleFontSize:20}));
    expCard(pres,s,6.75,4.15,6.05,2.8);
    s.addText('การแจ้งเตือน (Alerts)',{x:6.95,y:4.2,w:5.6,h:0.4,fontFace:EXP_FONT,fontSize:20,bold:true,color:EXP_C.ink,isTextBox:true,margin:0});
    expAlert(pres,s,6.95,4.65,5.65,1.0,'การเติบโตต่ำกว่า 3 g/สัปดาห์: '+gainLow.length+' บ่อ',expDelta(gainLow.length,pv('gainLow'),true),gainLow.length>0);
    expAlert(pres,s,6.95,5.8,5.65,1.0,'อัตรารอดสะสมน้อยกว่า 70%: '+srLow.length+' บ่อ',expDelta(srLow.length,pv('srLow'),true),srLow.length>0);
  }
  /* ===== 3 เฉพาะบ่อที่มีปัญหา ===== */
  if(T('w_problem')){
    const s=bg(); expTitle(pres,s,'ประสิทธิภาพประจำสัปดาห์ · เฉพาะบ่อที่มีปัญหา');
    expAlert(pres,s,0.5,1.12,6.0,0.72,'การเติบโตต่ำกว่า 3 g/สัปดาห์: '+gainLow.length+' บ่อ',null,gainLow.length>0);
    expAlert(pres,s,6.8,1.12,6.0,0.72,'อัตรารอดสะสม < 70%: '+srLow.length+' บ่อ',null,srLow.length>0);
    const tn=(v,dd)=>{ const x=nm(v); return isNaN(x)?'–':fmt(x/1000,dd===undefined?2:dd); };
    const rows=probSet.slice(0,9).map(r=>{const g=nm(r[11]),x=nm(r[14]);return [r[1],r[2]||'–',r[7]||'–',r[8]||'–',expCell(r[11]||'–',(!isNaN(g)&&g<3)?'red':''),r[9]||'–',r[12]||'–',r[13]||'–',expCell(r[14]||'–',(!isNaN(x)&&x>0&&x<70)?'red':((!isNaN(x)&&x<85)?'yellow':'')),tn(r[16]),tn(r[10]),tn(r[17])];});
    expTable(s,['Pond','System','DoC','MBW','MBW Gain','Size','ADG','%SR','Accum %SR','Total Biomass (ton/pond)','Biomass (ton/pond)','Total Yield (Ton/rai)'],rows.length?rows:[['ไม่พบบ่อที่มีปัญหา','','','','','','','','','','','']],{x:0.5,y:2.0,w:12.3,colW:[1.1,1.2,0.75,0.85,1.0,0.75,0.8,0.8,1.0,1.4,1.4,1.25],fs:15,rowH:0.38});
    expNote(s,'เกณฑ์: MBW Gain < 3 g/สัปดาห์ หรือ Accum %SR < 70% (แดง) · < 85% (เหลือง)'+(probSet.length>9?' · แสดง 9 จาก '+probSet.length+' บ่อ (ดูครบในสไลด์รายละเอียด)':''));
  }
  /* ===== 4 ประสิทธิภาพ รายละเอียด ===== */
  if(T('w_perfdetail')){
    const tn2=v=>{ const x=nm(v); return isNaN(x)?'–':fmt(x/1000,2); };
    const rows=cul.slice().sort((a,b)=>areaRank(a)-areaRank(b)||expN(b[7])-expN(a[7])||String(a[1]).localeCompare(String(b[1]))).map(r=>{const g=nm(r[11]),x=nm(r[14]);return [r[1],r[2]||'–',r[7]||'–',r[8]||'–',expCell(r[11]||'–',(!isNaN(g)&&g<3)?'red':''),r[9]||'–',r[12]||'–',r[13]||'–',expCell(r[14]||'–',(!isNaN(x)&&x>0&&x<70)?'red':((!isNaN(x)&&x<85)?'yellow':'')),r[15]||'–',tn2(r[10]),tn2(r[16]),r[18]||'–'];});
    expPagedW(pres,bg,'ประสิทธิภาพประจำสัปดาห์ (รายละเอียด)',['Pond','System','DoC','MBW (g)','MBW Gain','Size','ADG','%SR','Accum %SR','FCR','Biomass (ton/pond)','Total Biomass (ton/pond)','Feed/day (kg)'],rows.length?rows:[['–','','','','','','','','','','','','']],[1.05,1.15,0.75,0.9,0.95,0.75,0.8,0.8,1.05,0.8,1.1,1.2,0.95],10,14,'ข้อมูลล่าสุดจากหน้า Weekly (บ่อสถานะ Culture) · เรียงตามพื้นที่ A,B,C,D,R และ DOC มาก→น้อย',0.38);
  }
  /* ===== 5 คุณภาพน้ำเกินเกณฑ์ ===== */
  if(T('w_wq')){
    const s=bg(); expTitle(pres,s,'คุณภาพน้ำ เกินเกณฑ์มาตรฐาน · '+an);
    expHead(s,'บ่อที่มีค่าเกินเกณฑ์ '+wqBad.length+' จาก '+dy.length+' บ่อ · '+expDelta(wqBad.length,pv('wqBad'),true));
    expKpiRow(pres,s,wqKeys.map((q,i)=>({label:q.k,val:wqRed[i].length+' บ่อ',sub:'เฝ้าระวัง '+wqYel[i].length+' บ่อ',col:wqRed[i].length?EXP_BAD:(wqYel[i].length?EXP_WARN:EXP_OK)})).concat([{label:'ตัวตาย/วัน',val:mortRed.length+' บ่อ',sub:'เกินเกณฑ์',col:mortRed.length?EXP_BAD:EXP_OK}]),1.75);
    const rows=wqBad.slice().sort((a,b)=>wqKeys.filter(q=>expLet(b[q.i])==='red').length-wqKeys.filter(q=>expLet(a[q.i])==='red').length).slice(0,9).map(r=>[String(r[1]).trim(),docOf[String(r[1]).trim()]||'–'].concat(wqKeys.map(q=>expCell(fx(nm(r[q.v]),q.dd),expLet(r[q.i])))).concat([expCell(fx(nm(r[13]),0),expLet(r[24]))]));
    expTable(s,['Pond','DOC','pH','T.Alk (ppm)','TAN (ppm)','NO₂-N (ppm)','ตัวตาย (ตัว/วัน)'],rows.length?rows:[['ไม่มีบ่อเกินเกณฑ์','','','','','','']],{x:0.5,y:3.15,w:12.3,colW:[1.6,1.2,1.6,1.9,1.9,2.0,2.1],fs:15,rowH:0.38});
    expNote(s,'เกณฑ์: pH 7.0–9.0 · T.Alk 120–180 ppm · TAN ≤ 3 ppm · NO₂-N ≤ 3 ppm (แดง = เกินเกณฑ์ · เหลือง = เฝ้าระวัง)'+(wqBad.length>9?' · แสดง 9 จาก '+wqBad.length+' บ่อ':''));
  }
  /* ===== 6 คุณภาพน้ำ รายละเอียด ===== */
  if(T('w_wqdetail')){
    const rows=dy.slice().sort((a,b)=>String(a[1]).localeCompare(String(b[1]))).map(r=>[String(r[1]).trim(),docOf[String(r[1]).trim()]||'–',expCell(fx(nm(r[7]),0),expLet(r[19]))].concat(wqKeys.map(q=>expCell(fx(nm(r[q.v]),q.dd),expLet(r[q.i])))).concat([expCell(fx(nm(r[13]),0),expLet(r[24]))]));
    expPagedW(pres,bg,'คุณภาพน้ำประจำสัปดาห์ (รายละเอียด)',['Pond','DOC','Feed (kg)','pH','T.Alk','TAN','NO₂-N','ตัวตาย/วัน'],rows.length?rows:[['–','','','','','','','']],[1.6,1.1,1.6,1.4,1.4,1.4,1.4,1.5],13,14,'สีพื้น: เขียว = ปกติ · เหลือง = เฝ้าระวัง · แดง = เกินเกณฑ์ (ตามเกณฑ์ในหน้า Daily Report)',0.34);
  }
  /* ===== 7 แบคทีเรียเกินเกณฑ์ ===== */
  if(T('w_tvc')){
    const s=bg(); expTitle(pres,s,'แบคทีเรียในน้ำ-ในตับ เกินเกณฑ์มาตรฐาน · '+an);
    expCard(pres,s,0.5,1.2,6.05,3.0); expCard(pres,s,6.75,1.2,6.05,3.0);
    s.addText('เชื้อในตับ (In Liver)',{x:0.7,y:1.25,w:5.6,h:0.4,fontFace:EXP_FONT,fontSize:20,bold:true,color:EXP_C.ink,align:'center',isTextBox:true,margin:0});
    s.addText('เชื้อในน้ำ (In Water)',{x:6.95,y:1.25,w:5.6,h:0.4,fontFace:EXP_FONT,fontSize:20,bold:true,color:EXP_C.ink,align:'center',isTextBox:true,margin:0});
    expDonut(pres,s,2.3,1.7,2.4,1.9,tvcLpass,tvcLdata.length,EXP_OK,'ตับ');
    expDonut(pres,s,8.55,1.7,2.4,1.9,tvcWpass,tvcWdata.length,tvcWpass/Math.max(tvcWdata.length,1)<0.8?EXP_WARN:EXP_OK,'น้ำ');
    expPanel(pres,s,0.5,4.35,6.05,2.6,'FEF2F2','▲ บ่อที่เกินเกณฑ์',EXP_BAD,[
      'แบคทีเรียในน้ำเกินเกณฑ์ '+tvcWbad.length+' บ่อ: '+(tvcWbad.slice(0,6).map(x=>x.p).join(', ')||'–')+' · '+expDelta(tvcWbad.length,pv('tvcW'),true),
      'แบคทีเรียในตับเกินเกณฑ์ '+tvcLbad.length+' บ่อ: '+(tvcLbad.slice(0,6).map(x=>x.p).join(', ')||'–')+' · '+expDelta(tvcLbad.length,pv('tvcL'),true)],16,3);
    expPanel(pres,s,6.75,4.35,6.05,2.6,'F0FDF4','✔ แนวทางแก้ไข / ข้อเสนอ','15803D',['ลงโปรไบโอติก (PB) ทุกวันในบ่อที่เชื้อในน้ำเกินเกณฑ์ และติดตามผล TVC ซ้ำใน 3–5 วัน','ตรวจสอบตับ-ตับอ่อน (HP) บ่อที่เชื้อในตับสูง ปรับโปรแกรมอาหารเสริม','(แก้ไขข้อความในช่องนี้ได้ใน PowerPoint ก่อนนำเสนอ)'],16,3);
    expNote(s,'เกณฑ์ผ่าน: เชื้อในน้ำ ≤ 999 cfu/ml (เหลือง 1,000–4,999) · เชื้อในตับ ≤ 50,000 cfu/ml (เหลือง 50,001–60,000)');
  }
  /* ===== 8 แบคทีเรีย รายละเอียด ===== */
  if(T('w_tvcdetail')){
    const rows=tvcW.filter(x=>x.cw||x.cl).sort((a,b)=>a.p.localeCompare(b.p)).map(x=>[x.p,docOf[x.p]||'–',expCell(expSci(x.w),x.cw,'center'),expCell(expSci(x.l),x.cl,'center')]);
    expPagedW(pres,bg,'ค่าแบคทีเรียในน้ำ-ในตับ (รายละเอียด)',['Pond','DOC','TVC ในน้ำ (cfu/ml)','TVC ในตับ (cfu/ml)'],rows.length?rows:[['–','','','']],[2.4,1.8,4.0,4.1],13,14,'ค่าผ่านเกณฑ์: น้ำ '+(tvcWdata.length?fmt(tvcWpass/tvcWdata.length*100,0):'–')+'% · ตับ '+(tvcLdata.length?fmt(tvcLpass/tvcLdata.length*100,0):'–')+'%',0.34);
  }
  /* ===== 9 สรุปปัญหาและแนวทางแก้ไข ===== */
  if(T('w_issue')){
    const s=bg(); expTitle(pres,s,'สรุปปัญหาและแนวทางการแก้ไข');
    s.addText('ข้อความสร้างอัตโนมัติจากข้อมูล Dashboard — แก้ไข/เพิ่ม Comment ได้ใน PowerPoint ก่อนนำเสนอ',{x:0.5,y:1.12,w:12.3,h:0.32,align:'center',fontFace:EXP_FONT,fontSize:14,color:EXP_C.grey,isTextBox:true,margin:0});
    const nm5=(rs,n)=>rs.slice(0,n).map(r=>String(r[1]||r.p||'').trim()).join(', ')+(rs.length>n?' ฯลฯ':'');
    const lateP=d.prep.filter(r=>/ล่าช้า/.test(r[10]));
    const cand=[];
    if(gainLow.length) cand.push({n:gainLow.length,p:'กุ้งโตช้ากว่า 3 g/สัปดาห์ '+gainLow.length+' บ่อ\n– บ่อ '+nm5(gainLow,5),s:'ทบทวนปริมาณอาหาร/เช็คยอ ตรวจสุขภาพกุ้งและคุณภาพน้ำ เพิ่มอาหารเสริมตามโปรแกรม'});
    if(srLow.length) cand.push({n:srLow.length,p:'อัตรารอดสะสมต่ำกว่า 70% '+srLow.length+' บ่อ\n– บ่อ '+nm5(srLow,5),s:'ตรวจสาเหตุกุ้งตาย (โรค/คุณภาพน้ำ) วางแผนจับก่อนกำหนดหากเสียหายต่อเนื่อง'});
    if(wqBad.length) cand.push({n:wqBad.length,p:'คุณภาพน้ำเกินเกณฑ์ '+wqBad.length+' บ่อ (TAN/NO₂/pH/Alk)\n– บ่อ '+nm5(wqBad,5),s:'เพิ่มการถ่ายน้ำ/ปรับ Alk และจุลินทรีย์ปรับสภาพน้ำ ลดอาหารชั่วคราวในบ่อที่ TAN/NO₂ สูง'});
    if(tvcWbad.length||tvcLbad.length) cand.push({n:tvcWbad.length+tvcLbad.length,p:'แบคทีเรียเกินเกณฑ์ น้ำ '+tvcWbad.length+' บ่อ · ตับ '+tvcLbad.length+' บ่อ\n– บ่อ '+nm5(tvcWbad.concat(tvcLbad),5),s:'ลงโปรไบโอติกทุกวัน ติดตามผล TVC ซ้ำ และตรวจตับ-ตับอ่อน'});
    if(lateP.length) cand.push({n:lateP.length,p:'เตรียมบ่อล่าช้ากว่าแผน '+lateP.length+' บ่อ\n– บ่อ '+nm5(lateP,5),s:'เร่งขั้นตอนที่ค้าง กำหนดผู้รับผิดชอบรายบ่อ เพื่อให้ลงกุ้งตามแผน'});
    cand.sort((a,b)=>b.n-a.n);
    const top=cand.slice(0,3); while(top.length<3) top.push({p:'(เพิ่มปัญหาที่พบ)',s:'(เพิ่มแนวทางแก้ไข)'});
    const colr=[EXP_C.dark,'B8860B','B91C1C'];
    top.forEach((c,i)=>{
      const y=1.55+i*1.8;
      s.addShape(pres.ShapeType.roundRect,{x:0.5,y:y,w:6.05,h:1.65,fill:{color:'FFFFFF'},line:{color:colr[i],width:2},rectRadius:0.08});
      s.addText([{text:'ปัญหา (Problem)',options:{bold:true,color:colr[i],fontSize:20,breakLine:true}},{text:c.p,options:{color:EXP_C.ink,fontSize:16}}],{x:0.7,y:y+0.05,w:5.7,h:1.55,valign:'top',fontFace:EXP_FONT,isTextBox:true,margin:0});
      s.addShape(pres.ShapeType.roundRect,{x:6.75,y:y,w:6.05,h:1.65,fill:{color:'FFFFFF'},line:{color:colr[i],width:2},rectRadius:0.08});
      s.addText([{text:'แนวทางแก้ไข (Solution)',options:{bold:true,color:colr[i],fontSize:20,breakLine:true}},{text:c.s,options:{color:EXP_C.ink,fontSize:16}}],{x:6.95,y:y+0.05,w:5.7,h:1.55,valign:'top',fontFace:EXP_FONT,isTextBox:true,margin:0});
    });
  }
  /* ===== 10-11 แผนการลงกุ้ง ===== */
  const wmList=(st.wmonths&&st.wmonths.length?st.wmonths:[nowM]).slice().sort((a,b)=>a-b); const ry=st.year||nowY;
  const allStk=raw.stk.slice(1).filter(r=>areas.indexOf(String(r[0]).trim())>=0&&String(r[1]).trim()&&!isCx(r)).map(r=>({m:String(r[0]).trim(),p:String(r[1]).trim(),dt:expMonthOf(r[2]),ds:r[2],mo:+r[3],yr:+r[4],pcs:expN(r[5]),ty:String(r[6]||'').trim()})).filter(x=>x.pcs>0);
  wmList.forEach(rm=>{
  let nmx=rm, nyx=ry;
  let stkRows=allStk.filter(x=>x.mo===nmx&&x.yr===nyx), stkFb=false;
  const isDone=x=>{ if(!x.dt)return false; const t=new Date(x.dt.y,x.dt.m-1,x.dt.d); return t<=new Date(nowY,nowM-1,nowD); };
  const isNur=x=>/PL\s*1[0-9]|IN\d/i.test(x.ty+' '+x.p)&&!/Super/i.test(x.ty);
  const big=stkRows.filter(x=>!isNur(x)).sort((a,b)=>a.p.localeCompare(b.p)), nur=stkRows.filter(isNur).sort((a,b)=>a.p.localeCompare(b.p));
  const lst=a=>a.length?('รวม '+a.length+' บ่อ — '+a.slice(0,12).map(x=>x.p+' ('+fmt(x.pcs/1e6,2)+' ล้านตัว)').join(', ')+(a.length>12?' ฯลฯ':'')):'–';
  if(T('w_stockplan')){
    const s=bg(); expTitle(pres,s,'แผนการลงกุ้งเดือน '+MT[nmx-1]+' '+nyx+' · '+an);
    const tot=stkRows.reduce((a,x)=>a+x.pcs,0), done=stkRows.filter(isDone).reduce((a,x)=>a+x.pcs,0), wait=tot-done;
    expKpiRow(pres,s,[{label:'จำนวนบ่อตามแผน',val:stkRows.length+' บ่อ',sub:'ใหญ่ '+big.length+' · อนุบาล '+nur.length},{label:'ยอดลงตามแผน',val:fmt(tot/1e6,2)+' ล้านตัว',sub:fmt(tot,0)+' ตัว'},{label:'ลงแล้ว',val:fmt(done/1e6,2)+' ล้านตัว',sub:stkRows.filter(isDone).length+' บ่อ',col:EXP_OK},{label:'รอลง',val:fmt(wait/1e6,2)+' ล้านตัว',sub:stkRows.filter(x=>!isDone(x)).length+' บ่อ',col:EXP_WARN}],1.25);
    const pdone=tot?done/tot*100:0;
    const igr=a=>'บ่อ I จำนวน '+a.filter(x=>polyType(x.p)==='I').length+' บ่อ   บ่อ G จำนวน '+a.filter(x=>polyType(x.p)==='G').length+' บ่อ   บ่อ R จำนวน '+a.filter(x=>polyType(x.p)==='R').length+' บ่อ';
    [['สถานะการลงกุ้ง',(tot?'ลงแล้ว '+stkRows.filter(isDone).length+' บ่อ ('+fmt(pdone,0)+'%)':'ไม่มีแผน'),EXP_OK,pdone,'เป้าหมาย '+fmt(tot/1e6,2)+' ล้านตัว / ลงจริง '+fmt(done/1e6,2)+' ล้านตัว ('+stkRows.filter(isDone).length+' บ่อ)   '+igr(stkRows.filter(isDone)),'บ่อที่ลงแล้ว: '+(lst(stkRows.filter(isDone)))],['แผนดำเนินการ',(tot?'รอลง '+stkRows.filter(x=>!isDone(x)).length+' บ่อ ('+fmt(100-pdone,0)+'%)':'ไม่มีแผน'),EXP_WARN,tot?100-pdone:0,'จำนวนที่รอลง '+fmt(wait/1e6,2)+' ล้านตัว ('+stkRows.filter(x=>!isDone(x)).length+' บ่อ)   '+igr(stkRows.filter(x=>!isDone(x))),'แผนลงกุ้ง: '+(lst(stkRows.filter(x=>!isDone(x))))]].forEach((c,i)=>{
      const y=2.75+i*2.1; expCard(pres,s,0.5,y,12.3,1.95);
      s.addText(c[0],{x:0.8,y:y+0.1,w:6,h:0.5,fontFace:EXP_FONT,fontSize:26,bold:true,color:EXP_C.ink,isTextBox:true,margin:0});
      s.addShape(pres.ShapeType.roundRect,{x:10.0,y:y+0.12,w:2.5,h:0.45,fill:{color:c[2]},line:{color:c[2],width:0},rectRadius:0.2});
      s.addText(c[1],{x:10.0,y:y+0.12,w:2.5,h:0.45,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:18,bold:true,color:'FFFFFF',isTextBox:true,margin:0});
      s.addShape(pres.ShapeType.roundRect,{x:0.8,y:y+0.7,w:11.7,h:0.2,fill:{color:'E5E7EB'},line:{color:'E5E7EB',width:0},rectRadius:0.1});
      if(c[3]>0) s.addShape(pres.ShapeType.roundRect,{x:0.8,y:y+0.7,w:Math.max(0.2,11.7*c[3]/100),h:0.2,fill:{color:c[2]},line:{color:c[2],width:0},rectRadius:0.1});
      s.addText(c[4],{x:0.8,y:y+1.0,w:11.7,h:0.35,align:'center',fontFace:EXP_FONT,fontSize:16,color:EXP_C.ink,isTextBox:true,margin:0});
      s.addText(c[5],{x:0.8,y:y+1.38,w:11.7,h:0.5,align:'center',fontFace:EXP_FONT,fontSize:15,color:EXP_C.grey,valign:'top',isTextBox:true,margin:0});
    });
    expNote(s,(stkFb?'ยังไม่มีแผนลงกุ้งเดือนถัดไปในชีต Stocking จึงแสดงแผนเดือนปัจจุบัน · ':'')+'แผนลงกุ้งจากชีต Stocking · ถือว่า "ลงแล้ว" เมื่อวันที่ลงกุ้ง ≤ วันที่ส่งออกรายงาน');
  }
  if(T('w_stockdetail')){
    const head=['Module','Pond name','Stocking Date','Month','Year','Stocking (pcs.)','Type'], cw=[1.2,1.8,1.9,1.1,1.1,2.2,3.0];
    const part=(arr,title)=>{
      const t=arr.reduce((a,x)=>a+x.pcs,0), dn=arr.filter(isDone).reduce((a,x)=>a+x.pcs,0);
      const rows=arr.map(x=>[x.m,x.p,x.ds||'–',x.mo,x.yr,fmt(x.pcs,0),x.ty||'–']);
      const sums=[['รอลงกุ้ง','','','','',fmt(t-dn,0),''],['ลงแล้ว','','','','',fmt(dn,0),''],['รวม','','','','',fmt(t,0),'']];
      const per=14, pages=Math.max(1,Math.ceil(rows.length/per));
      for(let p=0;p<pages;p++){
        const sl=bg(); expTitle(pres,sl,title+' เดือน '+MT[nmx-1]+(pages>1?' ('+(p+1)+'/'+pages+')':''));
        let chunk=rows.slice(p*per,(p+1)*per); if(p===pages-1) chunk=chunk.concat(sums);
        expTable(sl,head,chunk,{x:0.5,y:1.3,w:12.3,colW:cw,fs:14,rowH:0.34,boldLast:p===pages-1});
      }
    };
    if(!big.length&&!nur.length){ const sl=bg(); expTitle(pres,sl,'แผนการลงกุ้ง เดือน '+MT[nmx-1]+' (รายละเอียด)'); sl.addText('ยังไม่มีแผนลงกุ้งในชีต Stocking สำหรับเดือนนี้',{x:0.5,y:3,w:12.3,h:1,align:'center',fontFace:EXP_FONT,fontSize:28,color:EXP_C.grey,isTextBox:true,margin:0}); }
    if(big.length) part(big,'แผนการลงกุ้งใหญ่ (รายละเอียด)');
    if(nur.length) part(nur,'แผนการลงกุ้งอนุบาล (รายละเอียด)');
  }
  /* ===== 12 แผนการจับกุ้ง ===== */
  if(T('w_harvest')){
    const s=bg(); expTitle(pres,s,'แผนการจับกุ้ง เดือน '+MT[rm-1]+' '+ry+' · '+an);
    const items=raw.oh.slice(1).filter(r=>areas.indexOf(String(r[10]).trim())>=0&&String(r[11]).trim()).map(r=>({p:String(r[11]).trim(),pd:expMonthOf(r[12]),ty:String(r[20]||'').trim(),tar:expN(r[19]),tsz:expN(r[18]),ad:expMonthOf(r[21]),asz:expN(r[24]),act:expN(r[26])})).filter(x=>x.pd&&x.pd.m===rm&&x.pd.y===ry&&x.ty);
    const types=[...new Set(items.map(x=>x.ty))].sort();
    const av=a=>{const v=a.filter(x=>x>0);return v.length?v.reduce((p,x)=>p+x,0)/v.length:NaN;};
    const rows=types.map(t=>{const g=items.filter(x=>x.ty===t);const ga=g.filter(x=>x.act>0);return [t,fmt(g.reduce((a,x)=>a+x.tar,0),1),fx(av(g.map(x=>x.tsz)),0),String(g.length),fmt(g.reduce((a,x)=>a+x.act,0),2),fx(av(g.map(x=>x.asz)),0),ga.length?String(ga.length):''];});
    const tT=items.reduce((a,x)=>a+x.tar,0), tA=items.reduce((a,x)=>a+x.act,0), nA=items.filter(x=>x.act>0).length;
    rows.push(['Grand Total',fmt(tT,1),'',String(items.length),fmt(tA,2),'',nA?String(nA):'']);
    expTable(s,['Row Labels','Est.Ton','Est. AVG Size','Est. Pond','Act. (Ton)','Act. AVG Size','Act. pond'],rows.length>1?rows:[['ไม่มีแผนจับเดือนนี้','','','','','','']],{x:0.5,y:1.35,w:7.3,colW:[1.4,0.95,1.1,0.95,1.0,1.1,0.8],fs:16,rowH:0.4,boldLast:true});
    const sT=av(items.map(x=>x.tsz)), sA=av(items.map(x=>x.asz));
    expKpiRow(pres,s,[{label:'Size เฉลี่ย เป้าหมาย',val:fx(sT,0),sub:'ตัว/กก.'},{label:'Size เฉลี่ย เกิดจริง',val:fx(sA,0),sub:'ตัว/กก.',col:(!isNaN(sT)&&!isNaN(sA)&&sA>sT*1.15)?EXP_WARN:EXP_OK}],5.4);
    if(types.length) s.addChart(pres.charts.BAR,[{name:'แผน (ตัน)',labels:types,values:types.map(t=>+items.filter(x=>x.ty===t).reduce((a,x)=>a+x.tar,0).toFixed(1))},{name:'จริง (ตัน)',labels:types,values:types.map(t=>+items.filter(x=>x.ty===t).reduce((a,x)=>a+x.act,0).toFixed(2))}],Object.assign(expChartBase('แผนจับ vs จับจริง แยกประเภท (ตัน)'),{x:8.0,y:1.25,w:4.9,h:3.6,barDir:'col',chartColors:['9CA3AF',EXP_C.teal],showValue:true,dataLabelPosition:'outEnd'}));
    expNote(s,'ข้อมูลจากหน้า Order Harvest (วันที่แผนจับในเดือนปัจจุบัน) · จับจริง = Ac.Ton');
  }
  });
  /* ===== 13 ประมาณการและผลดำเนินการ ===== */
  if(T('w_pnl')){
    const s=bg(); expTitle(pres,s,'ประมาณการณ์และผลดำเนินการ · '+an+' ปี '+st.year);
    const f=d.financial, F={W:f.reduce((a,r)=>a+r.Weight,0),S:f.reduce((a,r)=>a+r.Sales,0),C:f.reduce((a,r)=>a+Math.abs(r['Production Cost']),0),N:f.reduce((a,r)=>a+r.NetProfit,0)};
    const rai=dy.concat(raw.dy.slice(2).filter(r=>areas.indexOf(String(r[0]).trim())>=0&&/close/i.test(r[4])&&!isCx(r))).reduce((a,r)=>a+expN(r[37]),0);
    const szAvg=raw.oh.slice(1).filter(r=>areas.indexOf(String(r[10]).trim())>=0&&expN(r[26])>0&&expN(r[24])>0&&(expMonthOf(r[21])||{}).y===st.year).map(r=>expN(r[24])), szA=szAvg.length?szAvg.reduce((a,x)=>a+x,0)/szAvg.length:NaN;
    const price=F.W>0?F.S/F.W:NaN, cost=F.W>0?F.C/F.W:NaN;
    expKpiRow(pres,s,[
      {label:'ผลผลิต',val:fmt(F.W,0)+' ตัน',sub:rai>0?fmt(F.W*1000/rai/1000,1)+' ตัน/ไร่ (ประมาณ)':'–'},
      {label:'ขนาดเฉลี่ย',val:fx(szA,0),sub:'ตัว/กก. (ไซส์ที่จับจริง)'},
      {label:'ราคาขาย',val:fx(price,0),sub:'บาท/กก.'},
      {label:'ต้นทุนสุทธิ',val:fx(cost,0),sub:'บาท/กก.'},
      {label:'กำไร/กก.',val:fx(price-cost,0),sub:'บาท/กก.',col:(price-cost)>=0?EXP_OK:EXP_BAD},
      {label:'กำไรสุทธิ',val:fmt(F.N/1000,2)+' ลบ.',sub:'margin '+(F.S?fmt(F.N/F.S*100,1)+'%':'–'),col:F.N>=0?EXP_OK:EXP_BAD}
    ],1.15);
    const labs=f.map(r=>MT[r.m-1]);
    s.addChart(pres.charts.BAR,[{name:'กำไร/ขาดทุนสุทธิ (ล้านบาท)',labels:labs,values:f.map(r=>+(r.NetProfit/1000).toFixed(2))}],Object.assign(expChartBase('กำไร/ขาดทุนสุทธิ (ล้านบาท)'),{x:0.4,y:2.45,w:6.3,h:2.8,barDir:'col',chartColors:[EXP_C.dark],invertedColors:['DC2626'],showLegend:false,showValue:true,dataLabelPosition:'outEnd',dataLabelFontSize:11}));
    s.addChart(pres.charts.BAR,[{name:'ผลผลิต (ตัน)',labels:labs,values:f.map(r=>+r.Weight.toFixed(1))}],Object.assign(expChartBase('ผลผลิต (ตัน)'),{x:6.8,y:2.45,w:6.2,h:2.8,barDir:'col',chartColors:[EXP_C.teal],showLegend:false,showValue:true,dataLabelPosition:'outEnd',dataLabelFontSize:11}));
    const uRow=d.stocking.map(r=>fmt(r.pct,0)+'%');
    const ohs=raw.oh.slice(1).filter(r=>areas.indexOf(String(r[10]).trim())>=0&&expN(r[26])>0&&expN(r[24])>0&&(expMonthOf(r[21])||{}).y===st.year).map(r=>({m:expMonthOf(r[21]).m,t:expN(r[26]),sz:expN(r[24])}));
    const szRow=f.map(r=>{ const g=ohs.filter(x=>x.m===r.m), tt=g.reduce((a,x)=>a+x.t,0); return tt>0?fmt(g.reduce((a,x)=>a+x.t*x.sz,0)/tt,0):'–'; });
    const wtRow=f.map(r=>r.Weight>0?fmt(r.Weight,1):'–');
    expTable(s,['เดือน'].concat(labs),[['บ่อมีกุ้ง (%)'].concat(uRow),['Size เฉลี่ย (ตัว/กก.)'].concat(szRow),['ผลผลิต (ตัน)'].concat(wtRow)],{x:0.5,y:5.3,w:12.3,colW:[2.1].concat(labs.map(()=>10.2/labs.length)),fs:13,rowH:0.34,leftFirst:true});
    expNote(s,'หน่วย: ล้านบาท (กำไร) · ตัน (ผลผลิต) · ข้อมูลตามหน้า Financial / Stocking · ตัน/ไร่ = ผลผลิตรวม ÷ พื้นที่ (ไร่) ของบ่อทั้งหมดในพื้นที่ที่เลือก');
  }
  /* ===== 14 สาเหตุการขาดทุน ===== */
  if(T('w_loss')){
    const s=bg(); expTitle(pres,s,'สาเหตุการขาดทุน · '+an+' ปี '+st.year);
    const f=d.financial.filter(r=>r.Sales!==0||r.NetProfit!==0||r.NoStk!==0);
    const M=r=>r/1000, K=(r,k)=>r.Weight>0?r[k]/r.Weight:NaN;
    /* ประสิทธิภาพการผลิต = ยอดขาย − (ต้นทุนตรง+คงที่+แปรผัน) ก่อน No Stk & Damage / Period Cost */
    const eff=r=>r.Sales-(r.Direct+r.Fixed+r.Variable);
    const T0={eff:f.reduce((a,r)=>a+eff(r),0),ns:f.reduce((a,r)=>a+r.NoStk,0),pc:f.reduce((a,r)=>a+r.Period,0),ai:f.reduce((a,r)=>a+Math.abs(r.AdminInt),0),net:f.reduce((a,r)=>a+r.NetProfit,0)};
    expKpiRow(pres,s,[
      {label:'กำไร/ขาดทุนสุทธิ',val:fmt(M(T0.net),2)+' ลบ.',sub:'รวมทุกเดือนที่มีข้อมูล',col:T0.net>=0?EXP_OK:EXP_BAD},
      {label:'ประสิทธิภาพการผลิต',val:fmt(M(T0.eff),2)+' ลบ.',sub:'ยอดขาย − ต้นทุนผลิตปกติ',col:T0.eff>=0?EXP_OK:EXP_BAD},
      {label:'No Stk & Damage',val:fmt(M(T0.ns),2)+' ลบ.',sub:'ลงกุ้งไม่ติด / เสียหาย',col:T0.ns>0?EXP_BAD:EXP_OK},
      {label:'Period Cost',val:fmt(M(T0.pc),2)+' ลบ.',sub:'ต้นทุนช่วงเวลา (ผลต่างตัดจ่าย)',col:T0.pc>0?EXP_BAD:EXP_OK}
    ],1.15);
    const labs=f.map(r=>MT[r.m-1]);
    s.addChart(pres.charts.BAR,[
      {name:'ประสิทธิภาพการผลิต',labels:labs,values:f.map(r=>+M(eff(r)).toFixed(2))},
      {name:'No Stk & Damage',labels:labs,values:f.map(r=>-+M(r.NoStk).toFixed(2))},
      {name:'Period Cost',labels:labs,values:f.map(r=>-+M(r.Period).toFixed(2))},
      {name:'Admin & Interest',labels:labs,values:f.map(r=>-+M(Math.abs(r.AdminInt)).toFixed(2))}
    ],Object.assign(expChartBase('องค์ประกอบกำไร/ขาดทุนรายเดือน (ล้านบาท)'),{x:0.4,y:2.45,w:12.5,h:2.85,barDir:'col',barGrouping:'stacked',chartColors:[EXP_C.teal,'DC2626','F59E0B','94A3B8'],showLegend:true,showValue:false}));
    const row=(name,fn,dg)=>[name].concat(f.map(r=>{ const v=fn(r); return isNaN(v)?'–':fmt(v,dg); }));
    expTable(s,['บาท/กก.'].concat(labs),[
      row('ราคาขาย',r=>K(r,'Sales'),0),
      row('ต้นทุนผลิตปกติ',r=>K(r,'Direct')+K(r,'Fixed')+K(r,'Variable'),0),
      row('No Stk & Damage',r=>K(r,'NoStk'),1),
      row('Period Cost',r=>K(r,'Period'),1)
    ],{x:0.5,y:5.35,w:12.3,colW:[2.4].concat(labs.map(()=>9.9/labs.length)),fs:13,rowH:0.34,leftFirst:true});
    expNote(s,'ประสิทธิภาพการผลิต = ยอดขาย − (ต้นทุนตรง + คงที่ + แปรผัน) · ก่อนหัก No Stk & Damage, Period Cost, Admin & Interest · ข้อมูลหน้า Financial');
  }
}

/* ---------- generator ---------- */
function expGenerate(PptxGenJS,st,raw,bgs){
  const pres=new PptxGenJS(); pres.layout='LAYOUT_WIDE'; pres.title='Roiphet Farm Report'; pres.author='Roiphet Farm Dashboard';
  const weekly=st.style==='weekly', now=new Date();
  const st2=Object.assign({},st,{months:weekly?[1,2,3,4,5,6,7,8,9,10,11,12]:st.months});
  const ms=st2.months.slice().sort((a,b)=>a-b);
  const d=expCompute(raw,st2);
  const navy=st.theme==='navy'; Object.assign(EXP_C,EXP_THEMES[navy?'navy':'green']);
  const bg=()=>{ const s=pres.addSlide(); if(navy){ s.background={color:'FFFFFF'}; s.addShape(pres.ShapeType.rect,{x:0,y:0,w:13.333,h:1.1,fill:{color:'002060'},line:{color:'002060',width:0}}); s.addShape(pres.ShapeType.rect,{x:0,y:1.1,w:13.333,h:0.05,fill:{color:'2E75B6'},line:{color:'2E75B6',width:0}}); } else s.background={data:bgs.content}; return expShift(s); };
  const nowTxt=now.toLocaleDateString('th-TH',{day:'numeric',month:'long',year:'numeric'});
  const c=pres.addSlide(); if(navy) c.background={color:'002060'}; else c.background={data:bgs.cover};
  const MODTH={A:'เอ',B:'บี',C:'ซี',D:'ดี',R:'อาร์'};
  const farmLine='ฟาร์มเลี้ยงกุ้งโรงเรือนระบบปิดร้อยเพชร'+(st.area==='ALL'?'':' (โมดุล '+expAreaList(st.area).map(a=>MODTH[a]||a).join(', ')+')');
  c.addText(weekly?'รายงานความก้าวหน้า':'รายงานสรุปผู้บริหาร',{x:0.4,y:2.8,w:12,h:1.0,fontFace:EXP_FONT,fontSize:60,bold:true,color:'FFFFFF',valign:'middle',isTextBox:true,margin:0});
  c.addText(farmLine,{x:0.4,y:3.95,w:12,h:0.6,fontFace:EXP_FONT,fontSize:28,color:'FFFFFF',valign:'middle',isTextBox:true,margin:0});
  if(!weekly) c.addText('ปี '+st.year+'  |  '+(ms.length===12?'ทุกเดือน':ms.map(m=>MONTH_TH[m-1]).join(' ')),{x:0.4,y:4.5,w:9,h:0.5,fontFace:EXP_FONT,fontSize:24,color:'FFFFFF',valign:'middle',isTextBox:true,margin:0});
  c.addText(nowTxt,{x:0.4,y:6.6,w:6,h:0.5,fontFace:EXP_FONT,fontSize:24,color:'FFFFFF',valign:'middle',isTextBox:true,margin:0});
  if(weekly) expWeeklySlides(pres,bg,st,raw,d,now);
  else expExecSlides(pres,bg,st,d,ms,(n,dd)=>isNaN(n)?'–':expFmt(n,dd));
  const e=pres.addSlide(); if(navy) e.background={color:'002060'}; else e.background={data:bgs.thanks};
  e.addText('ขอบคุณครับ',{x:5.2,y:3.0,w:7.5,h:1.6,fontFace:EXP_FONT,fontSize:72,bold:true,color:'FFFFFF',isTextBox:true,margin:0});
  return pres;
}

/* ---------- UI + actions ---------- */
function expLoadScript(url){ return new Promise((res,rej)=>{ if(window.PptxGenJS)return res(); const s=document.createElement('script'); s.src=url; s.onload=res; s.onerror=()=>rej(new Error('โหลด PptxGenJS ไม่สำเร็จ')); document.head.appendChild(s); }); }
async function expImg(name){
  if(EXP_BG[name])return EXP_BG[name];
  const r=await fetch(name,{cache:'force-cache'}); if(!r.ok)throw new Error('โหลดภาพพื้นหลังไม่สำเร็จ: '+name);
  const b=await r.blob();
  const d=await new Promise((res,rej)=>{ const f=new FileReader(); f.onload=()=>res(f.result); f.onerror=rej; f.readAsDataURL(b); });
  return EXP_BG[name]=String(d).replace(/^data:/,'');
}
async function expRun(){
  const btn=el('expBtn'), msg=el('expMsg'), weekly=EXP_ST.style==='weekly';
  if(!weekly&&!EXP_ST.months.length){ msg.textContent='กรุณาเลือกเดือนอย่างน้อย 1 เดือน'; return; }
  if(!(weekly?EXP_ST.wtopics:EXP_ST.topics).length){ msg.textContent='กรุณาเลือกหัวข้ออย่างน้อย 1 หัวข้อ'; return; }
  btn.disabled=true;
  try{
    msg.textContent='กำลังโหลดข้อมูล…'; const raw=await expLoadData();
    msg.textContent='กำลังโหลดเครื่องมือสร้างไฟล์…'; await expLoadScript(EXP_PPTX_URL);
    const bgs={cover:await expImg('pptx-bg1.jpg'),content:await expImg('pptx-bg2.jpg'),thanks:await expImg('pptx-bg3.jpg')};
    msg.textContent='กำลังสร้างสไลด์…';
    const pres=expGenerate(window.PptxGenJS,EXP_ST,raw,bgs);
    const fn=(weekly?'Roiphet_Weekly_':'Roiphet_Exec_')+(EXP_ST.area==='ALL'?'Farm':EXP_ST.area.replace(/,/g,''))+'_'+new Date().toISOString().slice(0,10)+'.pptx';
    await pres.writeFile({fileName:fn});
    msg.textContent='✓ สร้างไฟล์สำเร็จ: '+fn;
  }catch(e){ msg.textContent='เกิดข้อผิดพลาด: '+(e&&e.message||e); console.error(e); }
  btn.disabled=false;
}
function expToggle(kind,v){
  const a=EXP_ST[kind]; const i=a.indexOf(v); if(i>=0)a.splice(i,1); else a.push(v);
  expPaint();
}
function expPaint(){
  const chip=(on,txt,fn)=>'<button class="expchip'+(on?' on':'')+'" onclick="'+fn+'">'+txt+'</button>';
  const weekly=EXP_ST.style==='weekly';
  el('expStyle').innerHTML=chip(!weekly,'รายงานผู้บริหาร','EXP_ST.style=\'exec\';expPaint()')+chip(weekly,'รายงานประจำสัปดาห์','EXP_ST.style=\'weekly\';expPaint()');
  el('expStyleHint').textContent=weekly?'นำเสนอสรุปความก้าวหน้าประจำสัปดาห์ · ปัญหาและแนวทางแก้ไข · แผนลงกุ้ง/จับกุ้ง · ประมาณการ (ใช้ข้อมูลล่าสุด ณ วันที่ส่งออก)':'สรุปผลการดำเนินงาน เทียบแผน/เกณฑ์ พร้อมปัญหาและแนวทางแก้ไข (ตามช่วงปี/เดือนที่เลือก)';
  el('expTheme').innerHTML=chip(EXP_ST.theme!=='navy','1. สีเขียว','EXP_ST.theme=\'green\';expPaint()')+chip(EXP_ST.theme==='navy','2. สีน้ำเงิน','EXP_ST.theme=\'navy\';expPaint()');
  el('expAreas').innerHTML=['ALL'].concat(EXP_AREAS).map(a=>chip(a==='ALL'?EXP_ST.area==='ALL':(EXP_ST.area!=='ALL'&&expAreaList(EXP_ST.area).indexOf(a)>=0),a==='ALL'?'ทั้งฟาร์ม':a,'expToggleArea(\''+a+'\');expPaint()')).join('');
  el('expYears').innerHTML=[2026,2027].map(y=>chip(EXP_ST.year===y,String(y),'EXP_ST.year='+y+';expPaint()')).join('');
  el('expMonths').innerHTML=MONTH_TH.map((t,i)=>chip(EXP_ST[weekly?'wmonths':'months'].indexOf(i+1)>=0,t,'expToggle(\''+(weekly?'wmonths':'months')+'\','+(i+1)+')')).join('');
  { const k=weekly?'wmonths':'months'; el('expMonthBtns').innerHTML='<button class="expchip" onclick="EXP_ST.'+k+'=[1,2,3,4,5,6,7,8,9,10,11,12];expPaint()">เลือกทั้งหมด</button><button class="expchip" onclick="EXP_ST.'+k+'=[];expPaint()">ล้างการเลือก</button>'; }
  el('expMonthHint').textContent=weekly?'เลือกได้หลายเดือน · ใช้กับแผนการลงกุ้ง / แผนการจับกุ้ง (สร้างสไลด์แยกรายเดือน) และการ์ดมีกุ้งสิ้นเดือน':'เลือกได้หลายเดือน';
  const list=weekly?EXP_WTOPICS:EXP_TOPICS, key=weekly?'wtopics':'topics';
  el('expTopics').innerHTML=list.map(t=>chip(EXP_ST[key].indexOf(t.id)>=0,t.label,'expToggle(\''+key+'\',\''+t.id+'\')')).join('')
    +'<div style="margin-top:6px"><button class="expchip" onclick="EXP_ST.'+key+'=('+JSON.stringify(list.map(t=>t.id))+');expPaint()">เลือกทั้งหมด</button><button class="expchip" onclick="EXP_ST.'+key+'=[];expPaint()">ล้างการเลือก</button></div>';
  el('expTopicTitle').setAttribute('data-t',weekly?'เนื้อหาหัวข้อ (รายงานประจำสัปดาห์)':'เนื้อหาหัวข้อ (รายงานผู้บริหาร)');
  let n=0; document.querySelectorAll('.expbox').forEach(b=>{ if(b.style.display==='none')return; n++; const h=b.querySelector('h3'); h.textContent=n+' · '+h.getAttribute('data-t'); });
}
async function pageExport(){
  el('subtitle').textContent='Export Data · สร้างไฟล์นำเสนอ (.pptx) จากข้อมูล Dashboard';
  el('app').innerHTML='<style>.expchip{border:1px solid var(--line,#d1d5db);background:#fff;border-radius:999px;padding:7px 16px;margin:0 8px 8px 0;font-size:14px;cursor:pointer;font-family:inherit}.expchip.on{background:#4AC5B1;border-color:#4AC5B1;color:#fff;font-weight:700}.expbox{background:#fff;border:1px solid var(--line,#e5e7eb);border-radius:14px;padding:18px 20px;margin-bottom:14px}.expbox h3{margin:0 0 10px;font-size:15px}.expgo{background:#1F7A6D;color:#fff;border:0;border-radius:12px;padding:13px 28px;font-size:16px;font-weight:700;cursor:pointer}.expgo:disabled{opacity:.6;cursor:wait}</style>'
   +'<div class="expbox"><h3 data-t="รูปแบบรายงาน">รูปแบบรายงาน</h3><div id="expStyle"></div><div class="hint" id="expStyleHint" style="margin-top:4px"></div></div>'
   +'<div class="expbox"><h3 data-t="โทนสีนำเสนอ">โทนสีนำเสนอ</h3><div id="expTheme"></div><div class="hint" style="margin-top:4px">สีเขียว = โทนเดิมของฟาร์ม · สีน้ำเงิน = หัวข้อแถบสีกรมท่า ตารางหัวสีน้ำเงิน</div></div>'
   +'<div class="expbox"><h3 data-t="เดือน / ปี">เดือน / ปี</h3><div id="expYears"></div><div id="expMonths"></div><div id="expMonthBtns"></div><div class="hint" id="expMonthHint"></div></div>'
   +'<div class="expbox"><h3 data-t="พื้นที่">พื้นที่</h3><div id="expAreas"></div></div>'
   +'<div class="expbox"><h3 id="expTopicTitle" data-t="เนื้อหาหัวข้อ">เนื้อหาหัวข้อ</h3><div id="expTopics"></div></div>'
   +'<div><button class="expgo" id="expBtn" onclick="expRun()">⭳ Export PowerPoint (.pptx)</button> <span id="expMsg" style="margin-left:12px;font-size:14px"></span></div>'
   +'<div class="hint" style="margin-top:12px">ฟอนต์ในไฟล์ใช้ FreesiaUPC (ต้องติดตั้งฟอนต์นี้ในเครื่องที่เปิดไฟล์) · เทมเพลตตามต้นแบบของฟาร์ม</div>';
  expPaint();
}
