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
const EXP_ST={area:'ALL',year:2026,months:[1,2,3,4,5,6,7,8,9,10,11,12],topics:['situation','stocking','harvest','performance','prepare','financial','summary']};
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
  const body=rows.map((r,i)=>r.map((c,j)=>({text:String(c),options:{color:(/ล่าช้า/.test(String(c))?EXP_C.red:EXP_C.ink),fill:{color:i%2?'FFFFFF':EXP_C.zebra},align:j===0?'center':(/^[\d.,\-–%\s]*$/.test(String(c))?'right':'left'),valign:'middle',fontFace:EXP_FONT,fontSize:fs,bold:!!(o.boldLast&&i===rows.length-1)}})));
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
  const T=id=>st.topics.indexOf(id)>=0;
  const labs=ms.map(m=>MT[m-1]);
  const fx=(n,dd)=>isNaN(n)?'–':expFmt(n,dd);

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
  el('expYears').innerHTML=[2026,2027].map(y=>chip(EXP_ST.year===y,String(y),'EXP_ST.year='+y+';expPaint()')).join('');
}
async function pageExport(){
  el('subtitle').textContent='Export Data · สร้างไฟล์นำเสนอ (.pptx) จากข้อมูล Dashboard';
  el('app').innerHTML='<style>.expchip{border:1px solid var(--line,#d1d5db);background:#fff;border-radius:999px;padding:7px 16px;margin:0 8px 8px 0;font-size:14px;cursor:pointer;font-family:inherit}.expchip.on{background:#4AC5B1;border-color:#4AC5B1;color:#fff;font-weight:700}.expbox{background:#fff;border:1px solid var(--line,#e5e7eb);border-radius:14px;padding:18px 20px;margin-bottom:14px}.expbox h3{margin:0 0 10px;font-size:15px}.expgo{background:#1F7A6D;color:#fff;border:0;border-radius:12px;padding:13px 28px;font-size:16px;font-weight:700;cursor:pointer}.expgo:disabled{opacity:.6;cursor:wait}</style>'
   +'<div class="expbox"><h3>1 · พื้นที่</h3><div id="expAreas"></div></div>'
   +'<div class="expbox"><h3>2 · ปี</h3><div id="expYears"></div></div>'
   +'<div class="expbox"><h3>3 · เดือน (เลือกได้หลายเดือน)</h3><div id="expMonths"></div><button class="expchip" onclick="EXP_ST.months=[1,2,3,4,5,6,7,8,9,10,11,12];expPaint()">เลือกทั้งหมด</button><button class="expchip" onclick="EXP_ST.months=[];expPaint()">ล้างการเลือก</button></div>'
   +'<div class="expbox"><h3>4 · เนื้อหาหัวข้อ</h3><div id="expTopics"></div></div>'
   +'<div><button class="expgo" id="expBtn" onclick="expRun()">⭳ Export PowerPoint (.pptx)</button> <span id="expMsg" style="margin-left:12px;font-size:14px"></span></div>'
   +'<div class="hint" style="margin-top:12px">ฟอนต์ในไฟล์ใช้ FreesiaUPC (ต้องติดตั้งฟอนต์นี้ในเครื่องที่เปิดไฟล์) · เทมเพลตตามต้นแบบของฟาร์ม</div>';
  expPaint();
}
