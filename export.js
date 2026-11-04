/* ===== Export Data (PPTX) ===== */
const EXP_FONT='FreesiaUPC';
const EXP_AREAS=['A','B','C','D','R'];
const EXP_TOPICS=[
  {id:'situation',label:'สถานการณ์ฟาร์ม (บ่อเลี้ยง/บ่อปิด/ชีวมวล)'},
  {id:'stocking',label:'Stocking (ลงกุ้ง / จับกุ้ง รายเดือน)'},
  {id:'harvest',label:'Order Harvest (แผน vs จริง รายเดือน)'},
  {id:'performance',label:'Performance (รายบ่อที่กำลังเลี้ยง)'},
  {id:'prepare',label:'Prepare pond (บ่อปิดรอเตรียมบ่อ)'},
  {id:'financial',label:'Financial (รายได้/ต้นทุน/กำไร รายเดือน)'}
];
const EXP_ST={area:'ALL',year:2026,months:[1,2,3,4,5,6,7,8,9,10,11,12],topics:['situation','stocking','harvest','performance','financial']};
const EXP_BG={};
const EXP_PPTX_URL='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const EXP_C={teal:'4AC5B1',dark:'1F7A6D',yellow:'FFC000',ink:'1F2937',grey:'6B7280',zebra:'EAF8F5',line:'D1D5DB',red:'DC2626'};
const EXP_MONTH_COL=[1,2,3,5,6,7,9,10,11,13,14,15];

function expN(x){ const n=num(x); return isNaN(n)?0:n; }
function expMonthOf(s){ const m=String(s||'').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m?{d:+m[1],m:+m[2],y:+m[3]}:null; }
function expAreaList(area){ return area==='ALL'?EXP_AREAS.slice():[area]; }
function expAreaName(area){ return area==='ALL'?'ทั้งฟาร์ม':'พื้นที่ '+area; }

async function expLoadData(){
  const [stk,oh,wk,fin]=await Promise.all([getTab('stocking'),getTab('orderharvest'),getTab('weekly'),getTab('financial')]);
  return {stk:stk.values,oh:oh.values,wk:wk.values,fin:fin.values};
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
  out.situation=areas.map(a=>{
    const rows=wk.filter(r=>String(r[0]).trim()===a);
    const cul=rows.filter(r=>/culture/i.test(r[4])), cl=rows.filter(r=>/close/i.test(r[4]));
    const bio=cul.reduce((s,r)=>s+expN(r[10]),0);
    const sr=cul.map(r=>num(r[13])).filter(x=>!isNaN(x)&&x>0), fcr=cul.map(r=>num(r[15])).filter(x=>!isNaN(x)&&x>0);
    const avg=v=>v.length?v.reduce((s,x)=>s+x,0)/v.length:NaN;
    return {area:a,total:rows.length,cul:cul.length,close:cl.length,bio,sr:avg(sr),fcr:avg(fcr)};
  });
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
  const body=rows.map((r,i)=>r.map((c,j)=>({text:String(c),options:{color:EXP_C.ink,fill:{color:i%2?'FFFFFF':EXP_C.zebra},align:j===0?'center':'right',valign:'middle',fontFace:EXP_FONT,fontSize:fs,bold:!!(o.boldLast&&i===rows.length-1)}})));
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
function expGenerate(PptxGenJS,st,d,bgs){
  const pres=new PptxGenJS(); pres.layout='LAYOUT_WIDE'; pres.title='Roiphet Farm Report'; pres.author='Roiphet Farm Dashboard';
  const MT=MONTH_TH, ms=st.months.slice().sort((a,b)=>a-b);
  const msLabel=(ms.length===12)?'ทุกเดือน':ms.map(m=>MT[m-1]).join(' ');
  const scope=expAreaName(st.area)+'  |  ปี '+st.year+'  |  '+msLabel;
  const now=new Date().toLocaleDateString('th-TH',{day:'numeric',month:'long',year:'numeric'});
  const bg=()=>{ const s=pres.addSlide(); s.background={data:bgs.content}; return s; };
  /* cover */
  const c=pres.addSlide(); c.background={data:bgs.cover};
  c.addText('รายงานสรุปข้อมูลฟาร์ม',{x:0.5,y:1.9,w:7,h:1.5,fontFace:EXP_FONT,fontSize:54,bold:true,color:EXP_C.yellow,valign:'top',isTextBox:true,margin:0});
  c.addText('ฟาร์มเลี้ยงกุ้งโรงเรือนระบบปิดร้อยเพชร',{x:0.5,y:3.5,w:7,h:0.7,fontFace:EXP_FONT,fontSize:28,color:'FFFFFF',isTextBox:true,margin:0});
  c.addText(scope,{x:0.5,y:5.2,w:9,h:0.5,fontFace:EXP_FONT,fontSize:24,color:'FFFFFF',isTextBox:true,margin:0});
  c.addText('ข้อมูล ณ วันที่ '+now,{x:0.5,y:5.75,w:9,h:0.5,fontFace:EXP_FONT,fontSize:22,color:'FFFFFF',isTextBox:true,margin:0});
  const T=id=>st.topics.indexOf(id)>=0;
  const labs=ms.map(m=>MT[m-1]);

  if(T('situation')){
    const s=bg(); expTitle(pres,s,'สถานการณ์ฟาร์มร้อยเพชร');
    const rows=d.situation.map(r=>[r.area,r.total,r.cul,r.close,expFmt(r.bio),isNaN(r.sr)?'–':expFmt(r.sr,1),isNaN(r.fcr)?'–':expFmt(r.fcr,2)]);
    if(d.situation.length>1){ const t=d.situation.reduce((a,r)=>({total:a.total+r.total,cul:a.cul+r.cul,close:a.close+r.close,bio:a.bio+r.bio}),{total:0,cul:0,close:0,bio:0}); rows.push(['รวม',t.total,t.cul,t.close,expFmt(t.bio),'–','–']); }
    expTable(s,['พื้นที่','บ่อทั้งหมด','กำลังเลี้ยง','ปิดบ่อ','ชีวมวล (กก.)','%SR เฉลี่ย','FCR เฉลี่ย'],rows,{x:0.5,y:1.4,w:6.6,colW:[0.9,1,1,0.9,1.1,0.9,0.8],fs:16,rowH:0.4,boldLast:d.situation.length>1});
    s.addChart(pres.charts.BAR,[{name:'กำลังเลี้ยง',labels:d.situation.map(r=>r.area),values:d.situation.map(r=>r.cul)},{name:'ปิดบ่อ',labels:d.situation.map(r=>r.area),values:d.situation.map(r=>r.close)}],
      Object.assign(expChartBase('จำนวนบ่อ แยกตามพื้นที่'),{x:7.3,y:1.3,w:5.6,h:5.4,barDir:'col',barGrouping:'stacked',chartColors:[EXP_C.teal,'F59E0B'],showValue:true,dataLabelPosition:'ctr'}));
    expNote(s,'ข้อมูลสถานะบ่อล่าสุดจากหน้า Weekly (ไม่ขึ้นกับตัวกรองเดือน)');
  }
  if(T('stocking')){
    const rows=d.stocking.map(r=>[MT[r.m-1],expFmt(r.prev),expFmt(r.pond),expFmt(r.mil,2),expFmt(r.harv),expFmt(r.end),expFmt(r.pct,0)+'%',expFmt(r.main)]);
    const tot=d.stocking.reduce((a,r)=>({pond:a.pond+r.pond,mil:a.mil+r.mil,harv:a.harv+r.harv}),{pond:0,mil:0,harv:0});
    rows.push(['รวม','–',expFmt(tot.pond),expFmt(tot.mil,2),expFmt(tot.harv),'–','–','–']);
    const s=bg(); expTitle(pres,s,'Stocking · '+expAreaName(st.area)+' ปี '+st.year);
    expTable(s,['เดือน','มีกุ้งยกมา (บ่อ)','ลงกุ้ง (บ่อ)','ลงกุ้ง (ล้านตัว)','จับกุ้ง (บ่อ)','มีกุ้งสิ้นเดือน (บ่อ)','% สิ้นเดือน','บ่อหลัก (บ่อ)'],rows,{x:0.5,y:1.25,w:12.3,colW:[1.3,1.6,1.4,1.7,1.4,1.9,1.5,1.5],fs:16,rowH:0.36,boldLast:true});
    const s2=bg(); expTitle(pres,s2,'Stocking · ลงกุ้ง vs จับกุ้ง (บ่อ)');
    s2.addChart(pres.charts.BAR,[{name:'ลงกุ้ง (บ่อ)',labels:labs,values:d.stocking.map(r=>r.pond)},{name:'จับกุ้ง (บ่อ)',labels:labs,values:d.stocking.map(r=>r.harv)}],
      Object.assign(expChartBase('จำนวนบ่อลงกุ้ง / จับกุ้ง รายเดือน'),{x:0.6,y:1.3,w:12.1,h:5.6,barDir:'col',barGrouping:'clustered',chartColors:[EXP_C.teal,'F59E0B'],showValue:true,dataLabelPosition:'outEnd'}));
  }
  if(T('harvest')){
    const s=bg(); expTitle(pres,s,'Order Harvest · แผน vs จริง (ตัน)');
    s.addChart(pres.charts.BAR,[{name:'แผน (ตัน)',labels:labs,values:d.harvest.map(r=>+r.plan.toFixed(1))},{name:'จริง (ตัน)',labels:labs,values:d.harvest.map(r=>+r.act.toFixed(1))}],
      Object.assign(expChartBase(expAreaName(st.area)+' ปี '+st.year),{x:0.4,y:1.3,w:7.6,h:5.6,barDir:'col',barGrouping:'clustered',chartColors:['9CA3AF',EXP_C.teal],showValue:false}));
    const rows=d.harvest.map(r=>[MT[r.m-1],expFmt(r.plan,1),expFmt(r.act,1)]);
    const t=d.harvest.reduce((a,r)=>({p:a.p+r.plan,a:a.a+r.act}),{p:0,a:0}); rows.push(['รวม',expFmt(t.p,1),expFmt(t.a,1)]);
    expTable(s,['เดือน','แผน (ตัน)','จริง (ตัน)'],rows,{x:8.3,y:1.4,w:4.5,colW:[1.3,1.6,1.6],fs:16,rowH:0.36,boldLast:true});
  }
  if(T('performance')){
    const rows=d.culture.map(r=>[r[1],r[2],r[7]||'–',r[8]||'–',r[9]||'–',r[10]||'–',r[12]||'–',r[13]?r[13]+'%':'–',r[15]||'–']);
    if(!rows.length)rows.push(['ไม่มีบ่อที่กำลังเลี้ยง','','','','','','','','']);
    expPaged(pres,bg,'Performance · '+expAreaName(st.area),['บ่อ','ประเภท','DOC','MBW (g)','Size','Biomass (kg)','ADG','%SR','FCR'],rows,[1.6,1.4,1.1,1.3,1.1,1.8,1.2,1.4,1.4],14,16,'ข้อมูลล่าสุดจากหน้า Weekly (บ่อสถานะ Culture)');
  }
  if(T('prepare')){
    const rows=d.close.map(r=>[r[1],r[2],r[5]||'–',r[6]||'–']);
    if(!rows.length)rows.push(['ไม่มีบ่อสถานะ Close','','','']);
    expPaged(pres,bg,'Prepare pond · '+expAreaName(st.area),['บ่อ','ประเภท','วันที่ปิดบ่อ','Prepare (วัน)'],rows,[3,3,3.3,3],14,16,'บ่อสถานะ Close จากหน้า Weekly');
  }
  if(T('financial')){
    const rows=d.financial.map(r=>[MT[r.m-1],expFmt(r.Weight),expFmt(r.Sales),expFmt(r['Production Cost']),expFmt(r.GrossProfit),expFmt(r.Admin+r.Interest),expFmt(r.NetProfit)]);
    const t=d.financial.reduce((a,r)=>({w:a.w+r.Weight,s:a.s+r.Sales,c:a.c+r['Production Cost'],g:a.g+r.GrossProfit,o:a.o+r.Admin+r.Interest,n:a.n+r.NetProfit}),{w:0,s:0,c:0,g:0,o:0,n:0});
    rows.push(['รวม',expFmt(t.w),expFmt(t.s),expFmt(t.c),expFmt(t.g),expFmt(t.o),expFmt(t.n)]);
    const s=bg(); expTitle(pres,s,'Financial · '+expAreaName(st.area)+' ปี '+st.year);
    expTable(s,['เดือน','ผลผลิต (ตัน)','ยอดขาย','ต้นทุนผลิต','Gross Profit','Admin+Interest','Net Profit'],rows,{x:0.5,y:1.25,w:12.3,colW:[1.5,1.8,1.8,1.8,1.8,1.8,1.8],fs:16,rowH:0.36,boldLast:true});
    expNote(s,'หน่วย: พันบาท (ยกเว้นผลผลิต = ตัน)  · ข้อมูลปี 2026 ตามหน้า Financial');
    const s2=bg(); expTitle(pres,s2,'Financial · ยอดขาย vs Net Profit');
    s2.addChart(pres.charts.BAR,[{name:'ยอดขาย',labels:labs,values:d.financial.map(r=>r.Sales)},{name:'Net Profit',labels:labs,values:d.financial.map(r=>r.NetProfit)}],
      Object.assign(expChartBase('พันบาท'),{x:0.6,y:1.3,w:12.1,h:5.6,barDir:'col',barGrouping:'clustered',chartColors:[EXP_C.teal,EXP_C.yellow],showValue:false}));
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
