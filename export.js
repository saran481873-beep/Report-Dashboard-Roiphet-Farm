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
const EXP_ST={area:'ALL',year:2026,style:'exec',appendix:true,months:[1,2,3,4,5,6,7,8,9,10,11,12],topics:['situation','stocking','harvest','performance','prepare','financial','summary']};
const EXP_TGT={sr:85,fcr:1.6,harv:90}; /* เกณฑ์อ้างอิง: %SR ขั้นต่ำ, FCR สูงสุด, จับจริง/แผน ขั้นต่ำ(%) */
const EXP_OK='16A34A',EXP_WARN='F59E0B',EXP_BAD='DC2626';
const EXP_BG={};
const EXP_PPTX_URL='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const EXP_C={teal:'4AC5B1',dark:'1F7A6D',yellow:'FFC000',ink:'1F2937',grey:'6B7280',zebra:'EAF8F5',line:'D1D5DB',red:'DC2626'};
const EXP_MONTH_COL=[1,2,3,5,6,7,9,10,11,13,14,15];

function expN(x){ const n=num(x); return isNaN(n)?0:n; }
function expMonthOf(s){ const m=String(s||'').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m?{d:+m[1],m:+m[2],y:+m[3]}:null; }
function expAreaList(area){ return area==='ALL'?EXP_AREAS.slice():[area]; }
function expAreaName(area){ return area==='ALL'?'ทั้งฟาร์ม':'พื้นที่ '+area; }

async function expLoadData(){
  const [stk,oh,wk,fin,pr]=await Promise.all([getTab('stocking'),getTab('orderharvest'),getTab('weekly'),getTab('financial'),getTab('prepare')]);
  return {stk:stk.values,oh:oh.values,wk:wk.values,fin:fin.values,pr:pr.values};
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
    return {area:a,total:rows.length,cul:cul.length,close:cl.length,bio:cul.reduce((x,r)=>x+expN(r[10]),0),doc:avg(g(cul,7)),size:avg(g(cul,9)),sr:avg(g(cul,13)),fcr:avg(g(cul,15)),adg:avg(g(cul,12))};
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
    const rec={m}; labels.forEach(l=>rec[l]=0);
    areas.forEach(a=>{
      const s=starts[a]; const col=EXP_MONTH_COL[m-1];
      for(let i=s+3;i<=s+9&&i<raw.fin.length;i++){ const lab=String(raw.fin[i][0]).trim(); if(labels.indexOf(lab)>=0) rec[lab]+=expN(raw.fin[i][col]); }
    });
    return rec;
  });
  return out;
}

/* ---------- slide helpers ---------- */
function expTitle(pres,slide,text){
  slide.addText(text,{x:1.5,y:0.12,w:10.3,h:0.85,align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:34,bold:true,color:'FFFFFF',isTextBox:true,margin:0});
}
function expNote(slide,text){
  slide.addText(text,{x:0.5,y:7.05,w:12.3,h:0.35,fontFace:EXP_FONT,fontSize:14,color:EXP_C.grey,isTextBox:true,margin:0});
}
function expTable(slide,head,rows,o){
  const fs=o.fs||16, hdr=head.map(h=>({text:String(h),options:{bold:true,color:'FFFFFF',fill:{color:EXP_C.dark},align:'center',valign:'middle',fontFace:EXP_FONT,fontSize:fs}}));
  const body=rows.map((r,i)=>r.map((c,j)=>({text:String(c),options:{color:(/ล่าช้า|▼/.test(String(c))?EXP_C.red:(/▲/.test(String(c))?EXP_OK:EXP_C.ink)),fill:{color:i%2?'FFFFFF':EXP_C.zebra},align:(j===0&&!o.leftFirst)?'center':(/^[\d.,\-–%\s]*$/.test(String(c))?'right':'left'),valign:'middle',fontFace:EXP_FONT,fontSize:fs,bold:!!(o.boldLast&&i===rows.length-1)}})));
  slide.addTable([hdr].concat(body),{x:o.x,y:o.y,w:o.w,colW:o.colW,rowH:o.rowH||0.34,border:{type:'solid',pt:0.5,color:EXP_C.line},margin:[0,0.06,0,0.06]});
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
  const pages=Math.max(1,Math.ceil(rows.length/perPage));
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
    s.addText(k.label,{x:x+0.22,y:y+0.06,w:w-0.3,h:0.32,fontFace:EXP_FONT,fontSize:16,color:EXP_C.grey,isTextBox:true,margin:0});
    s.addText(k.val,{x:x+0.22,y:y+0.36,w:w-0.3,h:0.5,fontFace:EXP_FONT,fontSize:30,bold:true,color:col,isTextBox:true,margin:0});
    s.addText(k.sub||'',{x:x+0.22,y:y+0.86,w:w-0.3,h:0.3,fontFace:EXP_FONT,fontSize:14,color:EXP_C.grey,isTextBox:true,margin:0});
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
  const sr=avgOf(cul,13), fcr=avgOf(cul,15), adg=avgOf(cul,12), size=avgOf(cul,9), doc=avgOf(cul,7);
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
      'บ่อมีกุ้ง '+sit.c+' จาก '+sit.t+' บ่อ · ชีวมวลรวม '+fmt(sit.b/1000,1)+' ตัน',
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
      {label:'DOC เฉลี่ย · Size เฉลี่ย',val:fx(doc,0)+' · '+fx(size,0),sub:'ชีวมวลรวม '+fmt(sit.b/1000,1)+' ตัน'}
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

/* ---------- generator ---------- */
function expInsights(st,d,ms){
  const MT=MONTH_TH, obs=[], watch={};
  const add=(pond,txt)=>{ (watch[pond]=watch[pond]||[]).push(txt); };
  d.culture.forEach(r=>{
    const sr=num(r[13]), fcr=num(r[15]), doc=num(r[7]);
    if(!isNaN(sr)&&sr>0&&sr<85) add(r[1],'%SR ต่ำ '+fmt(sr,1)+'%');
    if(!isNaN(fcr)&&fcr>1.6) add(r[1],'FCR สูง '+fmt(fcr,2));
  });
  d.prep.forEach(r=>{ if(/ล่าช้า/.test(r[10])) add(r[1],'เตรียมบ่อล่าช้า (Prepare '+(r[5]||'–')+' วัน)'); });
  const nCul=d.situation.reduce((a,r)=>a+r.cul,0), nTot=d.situation.reduce((a,r)=>a+r.total,0), nClose=d.situation.reduce((a,r)=>a+r.close,0);
  obs.push('ภาพรวม: มีบ่อทั้งหมด '+nTot+' บ่อ กำลังเลี้ยงกุ้ง '+nCul+' บ่อ ปิดบ่อ/เตรียมบ่อ '+nClose+' บ่อ');
  const late=d.prep.filter(r=>/ล่าช้า/.test(r[10])).length;
  obs.push('การเตรียมบ่อ: ล่าช้ากว่าแผน '+late+' จาก '+d.prep.length+' บ่อ');
  const lowSR=d.culture.filter(r=>{const x=num(r[13]);return !isNaN(x)&&x>0&&x<85;}).length;
  const hiFCR=d.culture.filter(r=>{const x=num(r[15]);return !isNaN(x)&&x>1.6;}).length;
  obs.push('Performance: %SR ต่ำกว่า 85% จำนวน '+lowSR+' บ่อ · FCR สูงกว่า 1.6 จำนวน '+hiFCR+' บ่อ');
  const hp=d.harvest.reduce((a,r)=>a+r.plan,0), ha=d.harvest.reduce((a,r)=>a+r.act,0);
  if(hp>0) obs.push('Order Harvest: จับจริง '+fmt(ha,1)+' ตัน จากแผน '+fmt(hp,1)+' ตัน ('+fmt(ha/hp*100,0)+'% ของแผน)');
  else if(ha>0) obs.push('Order Harvest: จับจริง '+fmt(ha,1)+' ตัน (ไม่มีแผนในช่วงที่เลือก)');
  const lowM=d.harvest.filter(r=>r.plan>0&&r.act<r.plan*0.9).map(r=>MT[r.m-1]);
  if(lowM.length) obs.push('เดือนที่จับจริงต่ำกว่าแผนเกิน 10%: '+lowM.join(' '));
  const neg=d.financial.filter(r=>r.NetProfit<0).map(r=>MT[r.m-1]);
  const tn=d.financial.reduce((a,r)=>a+r.NetProfit,0), ts=d.financial.reduce((a,r)=>a+r.Sales,0);
  obs.push('Financial: ยอดขายรวม '+fmt(ts)+' พันบาท · Net Profit รวม '+fmt(tn)+' พันบาท'+(neg.length?' · เดือนที่ขาดทุน: '+neg.join(' '):''));
  const list=Object.keys(watch).sort().map(k=>[k,watch[k].join(' · ')]);
  return {obs,list};
}

function expGenerate(PptxGenJS,st,d,bgs){
  const pres=new PptxGenJS(); pres.layout='LAYOUT_WIDE'; pres.title='Roiphet Farm Report'; pres.author='Roiphet Farm Dashboard';
  const MT=MONTH_TH, ms=st.months.slice().sort((a,b)=>a-b);
  const msLabel=(ms.length===12)?'ทุกเดือน':ms.map(m=>MT[m-1]).join(' ');
  const now=new Date().toLocaleDateString('th-TH',{day:'numeric',month:'long',year:'numeric'});
  const bg=()=>{ const s=pres.addSlide(); s.background={data:bgs.content}; return s; };
  /* 1. cover */
  const c=pres.addSlide(); c.background={data:bgs.cover};
  c.addText('รายงานสรุปข้อมูลฟาร์ม',{x:0.5,y:1.9,w:7,h:1.0,fontFace:EXP_FONT,fontSize:50,bold:true,color:EXP_C.yellow,valign:'top',isTextBox:true,margin:0});
  c.addText('ฟาร์มเลี้ยงกุ้งโรงเรือนระบบปิดร้อยเพชร',{x:0.5,y:3.0,w:8,h:0.6,fontFace:EXP_FONT,fontSize:26,color:'FFFFFF',isTextBox:true,margin:0});
  c.addText(expAreaName(st.area),{x:0.5,y:3.9,w:7,h:0.9,fontFace:EXP_FONT,fontSize:48,bold:true,color:'FFFFFF',isTextBox:true,margin:0});
  c.addText('ปี '+st.year+'  |  '+msLabel,{x:0.5,y:5.0,w:9,h:0.5,fontFace:EXP_FONT,fontSize:24,color:'FFFFFF',isTextBox:true,margin:0});
  c.addText('ข้อมูล ณ วันที่ '+now,{x:0.5,y:5.6,w:9,h:0.5,fontFace:EXP_FONT,fontSize:22,color:'FFFFFF',isTextBox:true,margin:0});
  const exec=(st.style==='exec'), skipStd=exec&&!st.appendix;
  const T=id=>!skipStd&&st.topics.indexOf(id)>=0&&!(exec&&id==='summary');
  const labs=ms.map(m=>MT[m-1]);
  const fx=(n,dd)=>isNaN(n)?'–':expFmt(n,dd);
  if(exec){
    expExecSlides(pres,bg,st,d,ms,fx);
    if(!skipStd){ const a=pres.addSlide(); a.background={data:bgs.cover};
      a.addText('ภาคผนวก',{x:0.5,y:2.6,w:9,h:1.2,fontFace:EXP_FONT,fontSize:60,bold:true,color:EXP_C.yellow,isTextBox:true,margin:0});
      a.addText('รายละเอียดข้อมูล · '+expAreaName(st.area),{x:0.5,y:3.9,w:9,h:0.7,fontFace:EXP_FONT,fontSize:30,color:'FFFFFF',isTextBox:true,margin:0}); }
  }

  /* 2. Situation */
  if(T('situation')){
    const s=bg(); expTitle(pres,s,'Situation · สถานการณ์ฟาร์ม '+expAreaName(st.area));
    const rows=d.situation.map(r=>[r.area,r.total,r.cul,r.close,fx(r.doc,0),fx(r.size,0),expFmt(r.bio)]);
    if(d.situation.length>1){ const t=d.situation.reduce((a,r)=>({total:a.total+r.total,cul:a.cul+r.cul,close:a.close+r.close,bio:a.bio+r.bio}),{total:0,cul:0,close:0,bio:0}); rows.push(['รวม',t.total,t.cul,t.close,'–','–',expFmt(t.bio)]); }
    expTable(s,['พื้นที่','จำนวนบ่อ','บ่อมีกุ้ง','บ่อปิด','DOC เฉลี่ย','Size เฉลี่ย','ชีวมวล (กก.)'],rows,{x:0.5,y:1.4,w:7.0,colW:[0.9,1,1,0.9,1,1,1.2],fs:16,rowH:0.4,boldLast:d.situation.length>1});
    s.addChart(pres.charts.BAR,[{name:'บ่อมีกุ้ง',labels:d.situation.map(r=>r.area),values:d.situation.map(r=>r.cul)},{name:'บ่อปิด',labels:d.situation.map(r=>r.area),values:d.situation.map(r=>r.close)}],
      Object.assign(expChartBase('จำนวนบ่อ แยกตามพื้นที่'),{x:7.7,y:1.3,w:5.2,h:5.4,barDir:'col',barGrouping:'stacked',chartColors:[EXP_C.teal,'F59E0B'],showValue:true,dataLabelPosition:'ctr'}));
    expNote(s,'ข้อมูลสถานะบ่อล่าสุดจากหน้า Weekly (ไม่ขึ้นกับตัวกรองเดือน) · บ่อมีกุ้ง = สถานะ Culture');
    const cp=d.culture.map(r=>[r[1],r[2],r[7]||'–',r[8]||'–',r[9]||'–',r[10]||'–']);
    if(cp.length) expPaged(pres,bg,'Situation · บ่อที่มีกุ้ง (DOC / Size)',['บ่อ','ประเภท','DOC','MBW (g)','Size (ตัว/กก.)','Biomass (กก.)'],cp,[2,2,2,2,2.2,2.1],14,16,'ข้อมูลล่าสุดจากหน้า Weekly');
  }
  /* 3. Stocking */
  if(T('stocking')){
    const rows=d.stocking.map(r=>[MT[r.m-1],expFmt(r.prev),expFmt(r.pond),expFmt(r.mil,2),expFmt(r.harv),expFmt(r.end),expFmt(r.pct,0)+'%',expFmt(r.main)]);
    const tot=d.stocking.reduce((a,r)=>({pond:a.pond+r.pond,mil:a.mil+r.mil,harv:a.harv+r.harv}),{pond:0,mil:0,harv:0});
    rows.push(['รวม','–',expFmt(tot.pond),expFmt(tot.mil,2),expFmt(tot.harv),'–','–','–']);
    const s=bg(); expTitle(pres,s,'Stocking · '+expAreaName(st.area)+' ปี '+st.year);
    expTable(s,['เดือน','มีกุ้งยกมา (บ่อ)','ลงกุ้ง (บ่อ)','ลงกุ้ง (ล้านตัว)','จับกุ้ง (บ่อ)','มีกุ้งสิ้นเดือน (บ่อ)','% สิ้นเดือน','บ่อหลัก (บ่อ)'],rows,{x:0.5,y:1.25,w:12.3,colW:[1.3,1.6,1.4,1.7,1.4,1.9,1.5,1.5],fs:16,rowH:0.36,boldLast:true});
    const s2=bg(); expTitle(pres,s2,'Stocking · กราฟรายเดือน');
    s2.addChart(pres.charts.BAR,[{name:'ลงกุ้ง (บ่อ)',labels:labs,values:d.stocking.map(r=>r.pond)},{name:'จับกุ้ง (บ่อ)',labels:labs,values:d.stocking.map(r=>r.harv)}],
      Object.assign(expChartBase('ลงกุ้ง / จับกุ้ง (บ่อ)'),{x:0.4,y:1.2,w:6.3,h:5.7,barDir:'col',barGrouping:'clustered',chartColors:[EXP_C.teal,'F59E0B'],showValue:true,dataLabelPosition:'outEnd'}));
    s2.addChart(pres.charts.BAR,[{name:'มีกุ้งยกมา (บ่อ)',labels:labs,values:d.stocking.map(r=>r.prev)},{name:'มีกุ้งสิ้นเดือน (บ่อ)',labels:labs,values:d.stocking.map(r=>r.end)}],
      Object.assign(expChartBase('ยกมา / สิ้นเดือน (บ่อ)'),{x:6.8,y:1.2,w:6.2,h:5.7,barDir:'col',barGrouping:'clustered',chartColors:['9CA3AF',EXP_C.dark],showValue:true,dataLabelPosition:'outEnd'}));
  }
  /* 4. Order Harvest */
  if(T('harvest')){
    const s=bg(); expTitle(pres,s,'Order Harvest · แผนจับ vs จับจริง');
    s.addChart(pres.charts.BAR,[{name:'แผนจับ (ตัน)',labels:labs,values:d.harvest.map(r=>+r.plan.toFixed(1))},{name:'จับจริง (ตัน)',labels:labs,values:d.harvest.map(r=>+r.act.toFixed(1))}],
      Object.assign(expChartBase(expAreaName(st.area)+' ปี '+st.year),{x:0.4,y:1.3,w:7.6,h:5.6,barDir:'col',barGrouping:'clustered',chartColors:['9CA3AF',EXP_C.teal],showValue:true,dataLabelPosition:'outEnd'}));
    const rows=d.harvest.map(r=>[MT[r.m-1],expFmt(r.plan,1),expFmt(r.act,1),r.plan>0?expFmt(r.act/r.plan*100,0)+'%':'–']);
    const t=d.harvest.reduce((a,r)=>({p:a.p+r.plan,a:a.a+r.act}),{p:0,a:0}); rows.push(['รวม',expFmt(t.p,1),expFmt(t.a,1),t.p>0?expFmt(t.a/t.p*100,0)+'%':'–']);
    expTable(s,['เดือน','แผน (ตัน)','จริง (ตัน)','จริง/แผน'],rows,{x:8.2,y:1.4,w:4.7,colW:[1.1,1.2,1.2,1.2],fs:16,rowH:0.36,boldLast:true});
  }
  /* 5. Performance */
  if(T('performance')){
    const s=bg(); expTitle(pres,s,'Performance · ค่าเฉลี่ยรายพื้นที่');
    const rows=d.situation.map(r=>[r.area,r.cul,fx(r.sr,1)+(isNaN(r.sr)?'':'%'),fx(r.fcr,2),fx(r.adg,2),fx(r.size,0)]);
    expTable(s,['พื้นที่','บ่อมีกุ้ง','%SR เฉลี่ย','FCR เฉลี่ย','ADG เฉลี่ย (g/วัน)','Size เฉลี่ย (ตัว/กก.)'],rows,{x:0.5,y:1.4,w:7.4,colW:[1,1.1,1.2,1.2,1.5,1.4],fs:16,rowH:0.4});
    const sl=d.situation.filter(r=>r.cul>0);
    if(sl.length) s.addChart(pres.charts.BAR,[{name:'%SR เฉลี่ย',labels:sl.map(r=>r.area),values:sl.map(r=>isNaN(r.sr)?0:+r.sr.toFixed(1))}],
      Object.assign(expChartBase('%SR เฉลี่ย แยกพื้นที่'),{x:8.1,y:1.3,w:4.8,h:5.4,barDir:'col',chartColors:[EXP_C.teal],showLegend:false,showValue:true,dataLabelPosition:'outEnd'}));
    expNote(s,'ค่าเฉลี่ยจากบ่อสถานะ Culture ตามหน้า Weekly');
    const rows2=d.culture.map(r=>[r[1],r[2],r[7]||'–',r[13]?r[13]+'%':'–',r[15]||'–',r[12]||'–',r[9]||'–',r[10]||'–']);
    if(!rows2.length)rows2.push(['ไม่มีบ่อที่กำลังเลี้ยง','','','','','','','']);
    expPaged(pres,bg,'Performance · รายบ่อ '+expAreaName(st.area),['บ่อ','ประเภท','DOC','%SR','FCR','ADG','Size','Biomass (กก.)'],rows2,[1.6,1.5,1.2,1.5,1.5,1.5,1.7,1.8],14,16,'ข้อมูลล่าสุดจากหน้า Weekly (บ่อสถานะ Culture)');
  }
  /* 6. Prepare pond */
  if(T('prepare')){
    const s=bg(); expTitle(pres,s,'Prepare pond · ตามแผน / ล่าช้า');
    const rows=d.prepArea.map(r=>[r.area,r.total,r.ok,r.late,r.na]);
    if(d.prepArea.length>1){ const t=d.prepArea.reduce((a,r)=>({total:a.total+r.total,ok:a.ok+r.ok,late:a.late+r.late,na:a.na+r.na}),{total:0,ok:0,late:0,na:0}); rows.push(['รวม',t.total,t.ok,t.late,t.na]); }
    expTable(s,['พื้นที่','บ่อปิดทั้งหมด','ตามแผน','ล่าช้า','ไม่ระบุแผน'],rows,{x:0.5,y:1.4,w:6.6,colW:[1.1,1.6,1.3,1.3,1.3],fs:16,rowH:0.4,boldLast:d.prepArea.length>1});
    s.addChart(pres.charts.BAR,[{name:'ตามแผน',labels:d.prepArea.map(r=>r.area),values:d.prepArea.map(r=>r.ok)},{name:'ล่าช้า',labels:d.prepArea.map(r=>r.area),values:d.prepArea.map(r=>r.late)},{name:'ไม่ระบุแผน',labels:d.prepArea.map(r=>r.area),values:d.prepArea.map(r=>r.na)}],
      Object.assign(expChartBase('สถานะเตรียมบ่อ แยกพื้นที่'),{x:7.3,y:1.3,w:5.6,h:5.4,barDir:'col',barGrouping:'stacked',chartColors:[EXP_C.teal,EXP_C.red,'9CA3AF'],showValue:true,dataLabelPosition:'ctr'}));
    expNote(s,'บ่อสถานะ Close จากหน้า Prepare pond');
    const sorted=d.prep.slice().sort((a,b)=>(/ล่าช้า/.test(b[10])?1:0)-(/ล่าช้า/.test(a[10])?1:0)||expN(b[5])-expN(a[5]));
    const rows2=sorted.map(r=>[r[1],r[2],r[5]||'–',r[6]||'–',r[9]||'–',r[10]||'–']);
    if(!rows2.length)rows2.push(['ไม่มีบ่อสถานะ Close','','','','','']);
    expPaged(pres,bg,'Prepare pond · รายบ่อ '+expAreaName(st.area),['บ่อ','ประเภท','Prepare (วัน)','ขั้นตอนปัจจุบัน','กำหนดเสร็จ','เทียบแผน'],rows2,[1.5,1.4,1.5,4.2,1.9,1.8],14,15,'เรียงบ่อล่าช้าก่อน');
  }
  /* 7. Financial */
  if(T('financial')){
    const rows=d.financial.map(r=>[MT[r.m-1],expFmt(r.Weight),expFmt(r.Sales),expFmt(r['Production Cost']),expFmt(r.GrossProfit),expFmt(r.Admin+r.Interest),expFmt(r.NetProfit)]);
    const t=d.financial.reduce((a,r)=>({w:a.w+r.Weight,s:a.s+r.Sales,c:a.c+r['Production Cost'],g:a.g+r.GrossProfit,o:a.o+r.Admin+r.Interest,n:a.n+r.NetProfit}),{w:0,s:0,c:0,g:0,o:0,n:0});
    rows.push(['รวม',expFmt(t.w),expFmt(t.s),expFmt(t.c),expFmt(t.g),expFmt(t.o),expFmt(t.n)]);
    const s=bg(); expTitle(pres,s,'Financial / P&L · '+expAreaName(st.area)+' ปี '+st.year);
    expTable(s,['เดือน','ผลผลิต (ตัน)','ยอดขาย','ต้นทุนผลิต','Gross Profit','Admin+Interest','Net Profit'],rows,{x:0.5,y:1.25,w:12.3,colW:[1.5,1.8,1.8,1.8,1.8,1.8,1.8],fs:16,rowH:0.36,boldLast:true});
    expNote(s,'หน่วย: พันบาท (ยกเว้นผลผลิต = ตัน) · ข้อมูลตามหน้า Financial');
    const s2=bg(); expTitle(pres,s2,'Financial / P&L · ยอดขาย ต้นทุน กำไร');
    s2.addChart(pres.charts.BAR,[{name:'ยอดขาย',labels:labs,values:d.financial.map(r=>r.Sales)},{name:'ต้นทุนผลิต',labels:labs,values:d.financial.map(r=>Math.abs(r['Production Cost']))},{name:'Net Profit',labels:labs,values:d.financial.map(r=>r.NetProfit)}],
      Object.assign(expChartBase('พันบาท'),{x:0.6,y:1.3,w:12.1,h:5.6,barDir:'col',barGrouping:'clustered',chartColors:[EXP_C.teal,'F59E0B',EXP_C.dark],showValue:false}));
  }
  /* 8. Summary */
  if(T('summary')){
    const ins=expInsights(st,d,ms);
    const s=bg(); expTitle(pres,s,'สรุปข้อสังเกต · '+expAreaName(st.area));
    s.addText(ins.obs.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<ins.obs.length-1,paraSpaceAfter:10}})),{x:0.7,y:1.4,w:12,h:5.4,fontFace:EXP_FONT,fontSize:24,color:EXP_C.ink,valign:'top',isTextBox:true,margin:0});
    expNote(s,'เกณฑ์เฝ้าระวัง: %SR < 85% · FCR > 1.6 · เตรียมบ่อล่าช้ากว่าแผน');
    const rows=ins.list.length?ins.list:[['ไม่พบบ่อที่ต้องติดตาม','']];
    expPaged(pres,bg,'บ่อที่ต้องติดตาม',['บ่อ','ประเด็น'],rows,[2,10.3],13,16,'คัดจากบ่อ Culture (%SR, FCR) และบ่อ Close (เตรียมบ่อล่าช้า)');
  }
  /* thank you */
  const e=pres.addSlide(); e.background={data:bgs.thanks};
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
  const btn=el('expBtn'), msg=el('expMsg');
  if(!EXP_ST.months.length){ msg.textContent='กรุณาเลือกเดือนอย่างน้อย 1 เดือน'; return; }
  if(!EXP_ST.topics.length){ msg.textContent='กรุณาเลือกหัวข้ออย่างน้อย 1 หัวข้อ'; return; }
  btn.disabled=true;
  try{
    msg.textContent='กำลังโหลดข้อมูล…'; const raw=await expLoadData();
    msg.textContent='กำลังโหลดเครื่องมือสร้างไฟล์…'; await expLoadScript(EXP_PPTX_URL);
    const bgs={cover:await expImg('pptx-bg1.jpg'),content:await expImg('pptx-bg2.jpg'),thanks:await expImg('pptx-bg3.jpg')};
    msg.textContent='กำลังสร้างสไลด์…';
    const d=expCompute(raw,EXP_ST);
    const pres=expGenerate(window.PptxGenJS,EXP_ST,d,bgs);
    const fn='Roiphet_'+(EXP_ST.area==='ALL'?'Farm':EXP_ST.area)+'_'+EXP_ST.year+'.pptx';
    await pres.writeFile({fileName:fn});
    msg.textContent='✓ สร้างไฟล์สำเร็จ: '+fn;
  }catch(e){ msg.textContent='เกิดข้อผิดพลาด: '+(e&&e.message||e); }
  btn.disabled=false;
}
function expToggle(kind,v){
  const a=EXP_ST[kind]; const i=a.indexOf(v); if(i>=0)a.splice(i,1); else a.push(v);
  expPaint();
}
function expPaint(){
  const chip=(on,txt,fn)=>'<button class="expchip'+(on?' on':'')+'" onclick="'+fn+'">'+txt+'</button>';
  el('expMonths').innerHTML=MONTH_TH.map((t,i)=>chip(EXP_ST.months.indexOf(i+1)>=0,t,'expToggle(\'months\','+(i+1)+')')).join('');
  el('expTopics').innerHTML=EXP_TOPICS.map(t=>chip(EXP_ST.topics.indexOf(t.id)>=0,t.label,'expToggle(\'topics\',\''+t.id+'\')')).join('');
  el('expAreas').innerHTML=['ALL'].concat(EXP_AREAS).map(a=>chip(EXP_ST.area===a,a==='ALL'?'ทั้งฟาร์ม':a,'EXP_ST.area=\''+a+'\';expPaint()')).join('');
  el('expStyle').innerHTML=chip(EXP_ST.style==='exec','รายงานผู้บริหาร (สรุป เทียบแผน ปัญหา/แนวทางแก้ไข)','EXP_ST.style=\'exec\';expPaint()')+chip(EXP_ST.style==='std','แบบตารางข้อมูล (เดิม)','EXP_ST.style=\'std\';expPaint()')+(EXP_ST.style==='exec'?chip(EXP_ST.appendix,'แนบภาคผนวก (ตารางรายละเอียดรายบ่อ/รายเดือน)','EXP_ST.appendix=!EXP_ST.appendix;expPaint()'):'');
  el('expYears').innerHTML=[2026,2027].map(y=>chip(EXP_ST.year===y,String(y),'EXP_ST.year='+y+';expPaint()')).join('');
}
async function pageExport(){
  el('subtitle').textContent='Export Data · สร้างไฟล์นำเสนอ (.pptx) จากข้อมูล Dashboard';
  el('app').innerHTML='<style>.expchip{border:1px solid var(--line,#d1d5db);background:#fff;border-radius:999px;padding:7px 16px;margin:0 8px 8px 0;font-size:14px;cursor:pointer;font-family:inherit}.expchip.on{background:#4AC5B1;border-color:#4AC5B1;color:#fff;font-weight:700}.expbox{background:#fff;border:1px solid var(--line,#e5e7eb);border-radius:14px;padding:18px 20px;margin-bottom:14px}.expbox h3{margin:0 0 10px;font-size:15px}.expgo{background:#1F7A6D;color:#fff;border:0;border-radius:12px;padding:13px 28px;font-size:16px;font-weight:700;cursor:pointer}.expgo:disabled{opacity:.6;cursor:wait}</style>'
   +'<div class="expbox"><h3>0 · รูปแบบรายงาน</h3><div id="expStyle"></div></div>'
   +'<div class="expbox"><h3>1 · พื้นที่</h3><div id="expAreas"></div></div>'
   +'<div class="expbox"><h3>2 · ปี</h3><div id="expYears"></div></div>'
   +'<div class="expbox"><h3>3 · เดือน (เลือกได้หลายเดือน)</h3><div id="expMonths"></div><button class="expchip" onclick="EXP_ST.months=[1,2,3,4,5,6,7,8,9,10,11,12];expPaint()">เลือกทั้งหมด</button><button class="expchip" onclick="EXP_ST.months=[];expPaint()">ล้างการเลือก</button></div>'
   +'<div class="expbox"><h3>4 · เนื้อหาหัวข้อ</h3><div id="expTopics"></div></div>'
   +'<div><button class="expgo" id="expBtn" onclick="expRun()">⭳ Export PowerPoint (.pptx)</button> <span id="expMsg" style="margin-left:12px;font-size:14px"></span></div>'
   +'<div class="hint" style="margin-top:12px">ฟอนต์ในไฟล์ใช้ FreesiaUPC (ต้องติดตั้งฟอนต์นี้ในเครื่องที่เปิดไฟล์) · เทมเพลตตามต้นแบบของฟาร์ม</div>';
  expPaint();
}
