let trades = [];
let equityChart, dailyChart, winLossChart;
let importedWorkbookActive = false;
let manualTrades = [];

const $ = id => document.getElementById(id);
const REPORT_WIDTH = 1120;
const SAVED_RANGE_KEY = 'qOptionsSelectedReportRange';

function saveSelectedRange(){
  try{
    localStorage.setItem(SAVED_RANGE_KEY,JSON.stringify({
      from:$('fromDate')?.value || '',
      to:$('toDate')?.value || ''
    }));
  }catch(_e){}
}

function restoreSelectedRange(){
  try{
    const saved=JSON.parse(localStorage.getItem(SAVED_RANGE_KEY)||'null');
    if(saved?.from && saved?.to){
      $('fromDate').value=saved.from;
      $('toDate').value=saved.to;
      return true;
    }
  }catch(_e){}
  return false;
}
const money = n => (n < 0 ? '-$' : '$') + Math.abs(Number(n) || 0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const moneyInt = n => (n < 0 ? '-$' : '$') + Math.abs(Math.round(Number(n) || 0)).toLocaleString('en-US');
const pct = n => `${Number(n) >= 0 ? '+' : ''}${(Number(n) || 0).toFixed(2)}%`;

function toISO(date){
  const y = date.getFullYear();
  const m = String(date.getMonth()+1).padStart(2,'0');
  const d = String(date.getDate()).padStart(2,'0');
  return `${y}-${m}-${d}`;
}

function currentWeekRange(){
  const today = new Date();
  today.setHours(12,0,0,0);
  const day = today.getDay();
  const diffToMonday = day === 0 ? -6 : 1-day;
  const from = new Date(today);
  from.setDate(today.getDate()+diffToMonday);
  const to = new Date(from);
  to.setDate(from.getDate()+6);
  return {from:toISO(from),to:toISO(to)};
}

function addDays(iso,days){
  const [y,m,d] = iso.split('-').map(Number);
  const date = new Date(y,m-1,d,12,0,0);
  date.setDate(date.getDate()+days);
  return toISO(date);
}

// البيانات التجريبية من Q-Options-FINAL41(2) فقط — بدون أي تعديل.
function buildDemoTrades(){
  const {from} = currentWeekRange();
  return [
    {date:addDays(from,0),symbol:'NVDA',option:'CALL',strike:'1000',buy:12.50,sell:18.90,profit:640,pct:51.20,notes:'Breakout play'},
    {date:addDays(from,0),symbol:'SPY',option:'CALL',strike:'532',buy:6.30,sell:10.80,profit:450,pct:71.43,notes:'Bull trend follow'},
    {date:addDays(from,1),symbol:'AAPL',option:'CALL',strike:'195',buy:4.60,sell:6.90,profit:230,pct:50.00,notes:'Earnings play'},
    {date:addDays(from,1),symbol:'QQQ',option:'CALL',strike:'445',buy:3.20,sell:4.90,profit:170,pct:53.13,notes:'Momentum'},
    {date:addDays(from,1),symbol:'META',option:'PUT',strike:'510',buy:7.80,sell:7.00,profit:-80,pct:-10.26,notes:'Rejection'},
    {date:addDays(from,2),symbol:'TSLA',option:'PUT',strike:'170',buy:8.20,sell:2.10,profit:-610,pct:-74.39,notes:'Stop hit'},
    {date:addDays(from,2),symbol:'AMD',option:'CALL',strike:'150',buy:5.10,sell:1.20,profit:-390,pct:-76.47,notes:'Weak guidance'},
    {date:addDays(from,2),symbol:'COIN',option:'CALL',strike:'240',buy:3.40,sell:2.80,profit:-60,pct:-17.65,notes:'News impact'},
    {date:addDays(from,3),symbol:'INTC',option:'PUT',strike:'30',buy:1.20,sell:.80,profit:-40,pct:-33.33,notes:'Weak chart'},
    {date:addDays(from,3),symbol:'PLTR',option:'CALL',strike:'20',buy:1.70,sell:1.10,profit:-5.29,pct:-5.29,notes:'Rejection'},
    {date:addDays(from,3),symbol:'RIVN',option:'PUT',strike:'12',buy:.90,sell:.50,profit:-40.49,pct:-44.44,notes:'Stop hit'},
    {date:addDays(from,4),symbol:'SNAP',option:'PUT',strike:'10',buy:.40,sell:.40,profit:0,pct:0,notes:'Flat'},
    {date:addDays(from,4),symbol:'SOFI',option:'CALL',strike:'7',buy:.30,sell:.30,profit:-7.14,pct:-57.14,notes:'Stop hit'},
    {date:addDays(from,4),symbol:'NFLX',option:'CALL',strike:'980',buy:9.20,sell:null,profit:0,pct:0,notes:'مفتوحة'},
    {date:addDays(from,4),symbol:'AMZN',option:'CALL',strike:'190',buy:4.10,sell:6.60,profit:250,pct:60.98,notes:'Breakout'},
    {date:addDays(from,4),symbol:'MSFT',option:'CALL',strike:'420',buy:5.40,sell:7.90,profit:250,pct:46.30,notes:'Trend'},
    {date:addDays(from,4),symbol:'GOOGL',option:'CALL',strike:'175',buy:3.80,sell:5.10,profit:130,pct:34.21,notes:'Continuation'}
  ];
}

function setCurrentWeekRange(){
  const range = currentWeekRange();
  $('fromDate').value = range.from;
  $('toDate').value = range.to;
}

function normalizeDigits(value){
  return String(value ?? '')
    .replace(/[٠-٩]/g,d=>'0123456789'['٠١٢٣٤٥٦٧٨٩'.indexOf(d)])
    .replace(/[۰-۹]/g,d=>'0123456789'['۰۱۲۳۴۵۶۷۸۹'.indexOf(d)]);
}

function parseDate(v){
  if (!v) return '';
  if (typeof v === 'number' && window.XLSX) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }

  const raw = normalizeDigits(v).trim();
  if (!raw) return '';

  // يدعم 2026-08-17 وكذلك 17/08/2026 و 17-08-2026
  let m = raw.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;

  const d = new Date(raw);
  if (!isNaN(d)) return toISO(d);
  return raw;
}

function num(v){
  if (v === null || v === undefined || v === '') return 0;
  const raw = normalizeDigits(v)
    .replace(/٬/g,'')
    .replace(/٫/g,'.')
    .replace(/[\s,$%٪]/g,'')
    .replace(/,/g,'');
  const n = Number(raw);
  return isNaN(n) ? 0 : n;
}

function normalizeKey(value){
  return normalizeDigits(value)
    .trim()
    .toLowerCase()
    .replace(/[أإآ]/g,'ا')
    .replace(/ة/g,'ه')
    .replace(/ى/g,'ي')
    .replace(/[\s_\-–—/\\()\[\].:%٪$]/g,'');
}

function getVal(row,aliases){
  const entries = Object.entries(row);
  const wanted = aliases.map(normalizeKey);
  for (const [k,v] of entries){
    if (wanted.includes(normalizeKey(k))) return v;
  }
  return '';
}

function normalizeRows(rows){
  return rows.map(r => {
    // إذا كانت أسماء الأعمدة مختلفة، نستخدم ترتيب الأعمدة كخيار احتياطي.
    const vals = Object.values(r);
    const pick = (aliases,index) => {
      const byName = getVal(r,aliases);
      return byName !== '' ? byName : (vals[index] ?? '');
    };

    const date = parseDate(pick(['date','trade date','التاريخ','تاريخ','تاريخ الصفقة'],0));
    const symbol = pick(['symbol','ticker','company','stock','اسم الشركة','اسم السهم','الشركة','السهم','الرمز'],1) || '—';
    const option = String(pick(['option','type','contract type','نوع الخيار','الخيار','نوع العقد'],8) || '').trim().toUpperCase();
    const strike = pick(['strike','strike price','الاسترايك','سترايك','سعر الاسترايك'],2) || '—';
    const buy = num(pick(['buy','buy price','entry','entry price','سعر الشراء','سعر الدخول','الدخول'],3));
    const sellRaw = pick(['sell','sell price','exit','exit price','سعر البيع','سعر الخروج','الخروج'],4);
    const sell = sellRaw === '' || sellRaw === null || sellRaw === undefined ? null : num(sellRaw);
    // الحساب دائماً من سعري الشراء والبيع حتى لا تتعارض صيغ Excel مع التقرير.
    const profit = sell === null ? 0 : (sell-buy)*100;
    const p = sell === null || !buy ? 0 : (sell-buy)/buy*100;
    const notes = pick(['notes','note','remarks','الملاحظات','ملاحظات','ملاحظة'],7) || '';

    return {date,symbol:String(symbol).trim(),option,strike:String(strike).trim(),buy,sell,profit,pct:p,notes:String(notes).trim()};
  }).filter(x => x.symbol !== '—' || x.date);
}

function colorize(el,value){
  el.classList.remove('positive','negative','neutral');
  el.classList.add(value > 0 ? 'positive' : value < 0 ? 'negative' : 'neutral');
}

function getFilteredTrades(){
  // بعد رفع Excel نعرض جميع صفوف الملف فوراً؛ التاريخ المختار يبقى عنواناً للتقرير.
  if(importedWorkbookActive) return [...trades];
  const from = $('fromDate').value;
  const to = $('toDate').value;
  return trades.filter(t => (!from || !t.date || t.date >= from) && (!to || !t.date || t.date <= to));
}

function tradeOutcome(trade){
  if(trade.sell===null || trade.sell===undefined || trade.sell==='') return 'open';
  const buy=Number(trade.buy)||0;
  const sell=Number(trade.sell)||0;
  if(sell>buy) return 'win';
  if(sell>0 && sell<buy) return 'stopped';
  if(sell===0 && buy>0) return 'loss';
  return 'flat';
}

function calculateStats(data){
  // مطابق تماماً لطريقة كود 41 للنسخة التجريبية.
  const closed = data.filter(t => tradeOutcome(t) !== 'open');
  const wins = closed.filter(t => tradeOutcome(t) === 'win');
  const stopped = closed.filter(t => tradeOutcome(t) === 'stopped');
  const losses = closed.filter(t => tradeOutcome(t) === 'loss');
  const counted = closed.filter(t => ['win','loss','stopped'].includes(tradeOutcome(t)));
  const negativeTrades = [...losses,...stopped];
  const grossWin = wins.reduce((s,t)=>s+t.profit,0);
  const grossLoss = negativeTrades.reduce((s,t)=>s+t.profit,0);
  const net = closed.reduce((s,t)=>s+t.profit,0);
  const totalCost = closed.reduce((s,t)=>s+(Math.abs(t.buy)*100),0);
  const returnP = totalCost ? net/totalCost*100 : 0;
  const winRate = counted.length ? wins.length/counted.length*100 : 0;
  const avgWin = wins.length ? grossWin/wins.length : 0;
  const avgLoss = negativeTrades.length ? grossLoss/negativeTrades.length : 0;
  const pf = grossLoss ? grossWin/Math.abs(grossLoss) : (grossWin ? Infinity : 0);
  const expectancy = closed.length ? net/closed.length : 0;

  let peak=0,eq=0,maxDD=0;
  const ordered=[...closed].sort((a,b)=>(a.date||'').localeCompare(b.date||''));
  ordered.forEach(t=>{eq+=t.profit;peak=Math.max(peak,eq);maxDD=Math.min(maxDD,eq-peak)});

  const returns=closed.map(t=>t.pct/100);
  const mean=returns.length ? returns.reduce((a,b)=>a+b,0)/returns.length : 0;
  const variance=returns.length>1 ? returns.reduce((s,x)=>s+(x-mean)**2,0)/(returns.length-1) : 0;
  const sharpe=variance ? mean/Math.sqrt(variance)*Math.sqrt(252) : 0;

  const byDay={};
  closed.forEach(t=>{if(t.date) byDay[t.date]=(byDay[t.date]||0)+t.profit});
  const dayVals=Object.values(byDay);
  const bestDay=dayVals.length ? Math.max(...dayVals) : 0;
  const worstDay=dayVals.length ? Math.min(...dayVals) : 0;
  const best=[...closed].sort((a,b)=>b.profit-a.profit)[0] || null;

  return {closed,counted,wins,stopped,losses,grossWin,grossLoss,net,totalCost,returnP,winRate,avgWin,avgLoss,pf,expectancy,maxDD,sharpe,ordered,byDay,bestDay,worstDay,best};
}

function render(){
  const filtered=getFilteredTrades();
  const s=calculateStats(filtered);

  $('totalTrades').textContent=filtered.length;
  $('totalProfit').textContent=money(s.net); colorize($('totalProfit'),s.net);
  $('returnPct').textContent=pct(s.returnP); colorize($('returnPct'),s.returnP);
  $('wins').textContent=s.wins.length;
  $('losses').textContent=s.losses.length;
  $('winRate').textContent=`${s.winRate.toFixed(2)}%`; colorize($('winRate'),s.winRate);

  $('bestSymbol').textContent=s.best ? `${s.best.symbol} ${s.best.strike}` : '—';
  $('bestProfit').textContent=s.best ? money(s.best.profit) : '$0.00'; colorize($('bestProfit'),s.best?.profit || 0);
  $('bestPct').textContent=s.best ? pct(s.best.pct) : '0%'; colorize($('bestPct'),s.best?.pct || 0);

  $('avgWin').textContent=money(s.avgWin); colorize($('avgWin'),s.avgWin);
  $('avgLoss').textContent=money(s.avgLoss); colorize($('avgLoss'),s.avgLoss);
  $('profitFactor').textContent=s.pf===Infinity ? '∞' : s.pf.toFixed(2);
  $('expectancy').textContent=money(s.expectancy); colorize($('expectancy'),s.expectancy);
  $('maxDrawdown').textContent=money(s.maxDD); colorize($('maxDrawdown'),s.maxDD);
  $('sharpe').textContent=s.sharpe.toFixed(2); colorize($('sharpe'),s.sharpe);

  $('summaryGross').textContent=money(s.grossWin); colorize($('summaryGross'),s.grossWin);
  $('summaryLoss').textContent=money(s.grossLoss); colorize($('summaryLoss'),s.grossLoss);
  $('summaryNet').textContent=money(s.net); colorize($('summaryNet'),s.net);
  $('bestDay').textContent=money(s.bestDay); colorize($('bestDay'),s.bestDay);
  $('worstDay').textContent=money(s.worstDay); colorize($('worstDay'),s.worstDay);

  renderTable(filtered);
  renderCharts(s.ordered,s.byDay,s.wins.length,s.losses.length,s.stopped.length);
  $('equityFinal').textContent=money(s.net);
  $('rowsCount').textContent=`${filtered.length} صفقة`;
}

function escapeHtml(value){
  return String(value ?? '').replace(/[&<>"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
}

function renderTable(data){
  $('tradesBody').innerHTML=data.map((t,i)=>`
    <tr>
      <td>${i+1}</td>
      <td dir="ltr">${escapeHtml(t.symbol)}</td>
      <td>${escapeHtml(t.strike)}</td>
      <td dir="ltr">${money(t.buy).replace('$','')}</td>
      <td dir="ltr">${t.sell===null?'—':money(t.sell).replace('$','')}</td>
      <td class="${t.profit>0?'profit':t.profit<0?'loss':''}">${t.sell===null&&t.profit===0?'<span style="color:#b48630">مفتوحة</span>':money(t.profit)}</td>
      <td class="${t.pct>0?'profit':t.pct<0?'loss':''}">${t.sell===null&&t.profit===0?'—':pct(t.pct)}</td>
      <td>${escapeHtml(t.notes || '—')}</td>
    </tr>`).join('');
}

function chartDefaults(){
  if(!window.Chart) return;
  Chart.defaults.color='#6f604f';
  Chart.defaults.font.family='Cairo';
  Chart.defaults.borderColor='rgba(163,137,99,.20)';
  if(!Chart.registry.plugins.get('qOptionsDepth')){
    Chart.register({
      id:'qOptionsDepth',
      beforeDatasetDraw(chart){
        const ctx=chart.ctx;
        ctx.save();
        ctx.shadowColor='rgba(116,72,13,.32)';
        ctx.shadowBlur=14;
        ctx.shadowOffsetX=0;
        ctx.shadowOffsetY=8;
      },
      afterDatasetDraw(chart){chart.ctx.restore()}
    });
  }
}

function chartGradient(context,top,bottom){
  const {chart}=context;
  const area=chart.chartArea;
  if(!area) return top;
  const g=chart.ctx.createLinearGradient(0,area.top,0,area.bottom);
  g.addColorStop(0,top);
  g.addColorStop(.48,bottom);
  g.addColorStop(1,'rgba(255,255,255,.08)');
  return g;
}

function renderCharts(ordered,byDay,winCount,lossCount,stoppedCount=0){
  if(!window.Chart) return;
  chartDefaults();
  const eqLabels=[],eqData=[];let cum=0;
  ordered.forEach((t,i)=>{cum+=t.profit;eqLabels.push(t.date||String(i+1));eqData.push(cum)});
  const days=Object.keys(byDay).sort(),vals=days.map(d=>byDay[d]);

  equityChart?.destroy();dailyChart?.destroy();winLossChart?.destroy();
  equityChart=new Chart($('equityChart'),{
    type:'line',
    data:{labels:eqLabels,datasets:[{data:eqData,borderColor:'#b57a18',backgroundColor:c=>chartGradient(c,'rgba(247,211,126,.72)','rgba(184,122,26,.20)'),fill:true,tension:.34,pointRadius:4,pointHoverRadius:7,pointBackgroundColor:'#fff4c9',pointBorderColor:'#b57a18',pointBorderWidth:3,borderWidth:4}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,plugins:{legend:{display:false}},scales:{x:{ticks:{maxRotation:0,autoSkip:true}},y:{ticks:{callback:v=>'$'+v}}}}
  });
  dailyChart=new Chart($('dailyChart'),{
    type:'bar',
    data:{labels:days,datasets:[{data:vals,backgroundColor:c=>chartGradient(c,c.raw>=0?'#86d6a0':'#ff9b90',c.raw>=0?'#25844b':'#b9312c'),borderColor:vals.map(v=>v>=0?'#247443':'#a92e28'),borderWidth:2,borderRadius:10,borderSkipped:false,barPercentage:.68}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>'$'+v}}}}
  });
  winLossChart=new Chart($('winLossChart'),{
    type:'doughnut',
    data:{labels:['رابحة','خاسرة','موقوفة'],datasets:[{data:[winCount,lossCount,stoppedCount],backgroundColor:c=>chartGradient(c,c.dataIndex===0?'#9ee3af':c.dataIndex===1?'#ffaaa0':'#9fe9f2',c.dataIndex===0?'#27854b':c.dataIndex===1?'#b92e2a':'#159ab1'),borderColor:['#fff5d7','#fff0e8','#e8fbff'],borderWidth:4,hoverOffset:10,spacing:3}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,cutout:'58%',rotation:-105,circumference:360,plugins:{legend:{position:'right',rtl:true,labels:{boxWidth:15,padding:18,font:{weight:'800'}}}}}
  });
}

async function handleFile(file){
  try{
    const ext=file.name.split('.').pop().toLowerCase();
    let normalized=[];

    if(ext==='csv'){
      normalized=normalizeRows(parseCsv(await file.text()));
    }else{
      if(!window.XLSX) throw new Error('مكتبة Excel لم يتم تحميلها');
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array',cellDates:false});

      // نفس قارئ الإكسل المستقر: نستخدم أول ورقة تحتوي على بيانات صفقات.
      for(const sheetName of wb.SheetNames){
        const ws=wb.Sheets[sheetName];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:'',raw:true});
        const candidate=normalizeRows(rows);
        if(candidate.length){normalized=candidate;break}
      }
    }

    if(!normalized.length){
      showToast('لم أجد صفقات. استخدم قالب Q Options الجاهز');
      return;
    }

    trades=[...normalized];
    importedWorkbookActive=true;
    // لا نسمح لملف Excel بتغيير التاريخ الذي اختاره المستخدم.
    saveSelectedRange();
    // تحديث فوري ثم تحديث مؤكد بعد إغلاق منتقي الملفات في Safari/iPhone.
    render();
    document.body.offsetHeight;
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    render();
    setTimeout(render,80);
    setTimeout(render,260);
    showToast(`تم تحميل ${trades.length} صفقة بنجاح`);
  }catch(err){
    console.error(err);
    showToast('تعذر قراءة الملف. استخدم XLSX أو CSV بالقالب المرفق');
  }finally{
    $('excelFile').value='';
  }
}

function parseCsv(text){
  const lines=text.replace(/^\uFEFF/,'').split(/\r?\n/).filter(line=>line.trim()!=='');
  if(!lines.length) return [];

  // Excel قد يحفظ CSV بفاصلة أو فاصلة منقوطة أو Tab حسب الجهاز/اللغة.
  const header=lines[0];
  const counts={',':(header.match(/,/g)||[]).length,';':(header.match(/;/g)||[]).length,'\t':(header.match(/\t/g)||[]).length};
  const delimiter=Object.entries(counts).sort((a,b)=>b[1]-a[1])[0][0];

  const split=line=>{
    const out=[];let cur='',q=false;
    for(let i=0;i<line.length;i++){
      const c=line[i];
      if(c==='"'){if(q&&line[i+1]==='"'){cur+='"';i++}else q=!q}
      else if(c===delimiter&&!q){out.push(cur);cur=''}
      else cur+=c;
    }
    out.push(cur);return out;
  };
  const headers=split(lines[0]).map(h=>h.trim());
  return lines.slice(1).map(line=>{const values=split(line),obj={};headers.forEach((h,i)=>obj[h]=values[i]??'');return obj});
}

function setRangeFromTrades(){
  const dates=trades.map(t=>t.date).filter(Boolean).sort();
  if(dates.length){$('fromDate').value=dates[0];$('toDate').value=dates[dates.length-1]}
  else setCurrentWeekRange();
}

function updateFooterClock(){
  const now=new Date();
  $('footerDate').textContent=now.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})+' - '+now.toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'});
}

function formatRange(){
  return `${$('fromDate').value || '—'}  →  ${$('toDate').value || '—'}`;
}

function safeFileRange(){
  return `${$('fromDate').value || 'from'}_${$('toDate').value || 'to'}`.replace(/[^0-9A-Za-z_-]/g,'-');
}

function topTrades(data,count=5){
  return [...data].filter(t=>t.sell!==null||t.profit!==0).sort((a,b)=>b.profit-a.profit).slice(0,count);
}

function buildPublicTrades(data,count=10){
  const closed=[...data].filter(t=>t.sell!==null||t.profit!==0);
  if(closed.length<=count) return closed;
  const wins=closed.filter(t=>t.profit>0).sort((a,b)=>b.profit-a.profit);
  const losses=closed.filter(t=>t.profit<0).sort((a,b)=>a.profit-b.profit);
  const neutral=closed.filter(t=>t.profit===0);
  const selected=[];
  const pushUnique=item=>{ if(item && !selected.includes(item)) selected.push(item); };
  const lossSlots=Math.min(2, losses.length, Math.max(1, count>=8?2:1));
  const winSlots=Math.min(wins.length, count-lossSlots);
  wins.slice(0,winSlots).forEach(pushUnique);
  losses.slice(0,lossSlots).forEach(pushUnique);
  for(const item of neutral){ if(selected.length>=count) break; pushUnique(item); }
  const rest=[...closed].sort((a,b)=>Math.abs(b.profit)-Math.abs(a.profit));
  for(const item of rest){ if(selected.length>=count) break; pushUnique(item); }
  return selected.slice(0,count).sort((a,b)=>(a.date||'').localeCompare(b.date||'') || String(a.symbol).localeCompare(String(b.symbol)));
}

function logoSrc(){
  const img=$('mainLogo');
  return img?.currentSrc || img?.src || 'QQ.PNG';
}

function chartImage(chart){
  try{return chart?.toBase64Image('image/png',1) || ''}catch{return ''}
}

// ينتظر تحميل الشعار وأي صور داخل قالب التصدير قبل تشغيل html2canvas.
// كانت هذه الدالة مستدعاة في الحفظ وPDF لكنها غير موجودة، فيتوقف الزر فورًا.
async function waitForImages(container){
  if(!container) return;
  const images=[...container.querySelectorAll('img')];
  const jobs=images.map(img=>new Promise(resolve=>{
    if(img.complete && img.naturalWidth>0){resolve();return}
    const done=()=>resolve();
    img.addEventListener('load',done,{once:true});
    img.addEventListener('error',done,{once:true});
    setTimeout(done,8000);
  }));
  if(document.fonts?.ready){
    jobs.push(Promise.race([document.fonts.ready,new Promise(r=>setTimeout(r,3000))]));
  }
  await Promise.all(jobs);
}


function tradeOptionLabel(trade){
  const explicit=String(trade.option||'').toUpperCase();
  if(explicit==='CALL' || explicit==='PUT') return explicit;
  const notes=(trade.notes||'').toLowerCase();
  if(notes.includes('put')) return 'PUT';
  if(notes.includes('call')) return 'CALL';
  return trade.pct >= 0 ? 'CALL' : 'PUT';
}

function tradeStatusMeta(trade){
  const outcome=tradeOutcome(trade);
  if(outcome==='stopped') return {cls:'stopped',labelAr:'موقوفة'};
  if(outcome==='loss') return {cls:'loss',labelAr:'خسارة'};
  if(outcome==='open') return {cls:'open',labelAr:'مفتوحة'};
  if(outcome==='flat') return {cls:'flat',labelAr:'متعادل'};
  return {cls:'win',labelAr:'ربح'};
}

function buildPdfTemplate(){
  return buildShareTemplate(10, "pdfCapture");
}


function buildShareTemplate(maxRows=10, captureId="shareCapture"){
  const filtered=getFilteredTrades();
  const s=calculateStats(filtered);
  // ترتيب التقرير: الرابحة أولاً، ثم الموقوفة، ثم الخاسرة.
  const outcomeOrder={win:0,stopped:1,loss:2};
  const rowsData=[...s.counted].sort((a,b)=>
    (outcomeOrder[tradeOutcome(a)]??9)-(outcomeOrder[tradeOutcome(b)]??9) ||
    (a.date||'').localeCompare(b.date||'') ||
    String(a.symbol).localeCompare(String(b.symbol))
  );
  const periodText=`${$('fromDate').value || '—'}  →  ${$('toDate').value || '—'}`;
  const winPct = s.counted.length ? (s.wins.length / s.counted.length * 100) : 0;
  const stoppedCount = s.stopped.length;
  const lossPct = s.counted.length ? (s.losses.length / s.counted.length * 100) : 0;
  const stoppedPct = s.counted.length ? (stoppedCount / s.counted.length * 100) : 0;
  const distributionPct=n=>s.counted.length?(n/s.counted.length*100):0;
  const chartTrades=rowsData.slice(0,20);
  const denseCharts=chartTrades.length>12;
  const equityTrades=rowsData
    .filter(t=>tradeOutcome(t)==='win')
    .sort((a,b)=>a.profit-b.profit || String(a.symbol).localeCompare(String(b.symbol)))
    .slice(0,20);
  const denseEquity=equityTrades.length>12;
  const maxTrade=Math.max(1,...chartTrades.map(t=>Math.abs(t.profit)));
  const tradeBars=chartTrades.map((t,i)=>{
    const x=18+i*(304/Math.max(1,chartTrades.length));
    const h=Math.max(4,Math.abs(t.profit)/maxTrade*55);
    const y=t.profit>=0?72-h:72;
    const state=tradeOutcome(t);
    const color=state==='win'?`url(#barWin-${captureId})`:state==='stopped'?`url(#barStopped-${captureId})`:`url(#barLoss-${captureId})`;
    const slot=304/Math.max(1,chartTrades.length);
    const w=Math.min(24,Math.max(7,slot*.58));
    const barX=x+(slot-w)/2;
    const cx=barX+w/2;
    const capY=t.profit>=0?y:y+h;
    return `<g class="cylinder-bar">
      <rect x="${barX.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" rx="${Math.min(7,w/2).toFixed(1)}" fill="${color}" stroke="${state==='win'?'#237b43':state==='stopped'?'#14859a':'#a92f29'}" stroke-width="1.2"/>
      <rect x="${(barX+w*.16).toFixed(1)}" y="${(y+2).toFixed(1)}" width="${(w*.20).toFixed(1)}" height="${Math.max(0,h-4).toFixed(1)}" rx="1.5" fill="rgba(255,255,255,.34)"/>
      <text class="bar-symbol" x="${cx.toFixed(1)}" y="146" text-anchor="start" transform="rotate(-90 ${cx.toFixed(1)} 146)">${escapeHtml(t.symbol).slice(0,5)}</text>
    </g>`;
  }).join('');
  let running=0;
  const equityValues=equityTrades.map(t=>(running+=t.profit));
  const eqMin=Math.min(0,...equityValues),eqMax=Math.max(1,...equityValues),eqRange=Math.max(1,eqMax-eqMin);
  const equityPoints=equityValues.map((v,i)=>`${(16+i*(292/Math.max(1,equityValues.length-1))).toFixed(1)},${(122-(v-eqMin)/eqRange*94).toFixed(1)}`).join(' ');
  const rows = rowsData.length ? rowsData.map(t=>{
    const option=tradeOptionLabel(t);
    const st=tradeStatusMeta(t);
    const statusAr = st.labelAr;
    const optionAr = option==='CALL' ? 'Call 📈' : 'Put 📉';
    return `
      <tr>
        <td class="symbol-cell"><div class="symbol-stack"><span>${escapeHtml(t.symbol)}</span></div></td>
        <td dir="ltr"><span class="info-chip ${option==='CALL'?'call':'put'}" style="display:inline;background:transparent;border:0;border-radius:0;box-shadow:none;padding:0;font-size:27px;font-weight:900;color:${option==='CALL'?'#278c43':'#dc3f36'}">${optionAr}</span></td>
        <td>${escapeHtml(t.strike)}</td>
        <td dir="ltr">${money(t.buy)}</td>
        <td dir="ltr">${t.sell===null?'—':money(t.sell)}</td>
        <td class="info-profit ${st.cls==='stopped'?'stopped-value':(t.profit>=0?'pos':'neg')}">${t.sell===null&&t.profit===0?'—':money(t.profit)}</td>
        <td class="info-pct ${st.cls==='stopped'?'stopped-value':(t.pct>=0?'pos':'neg')}">${t.sell===null&&t.profit===0?'—':pct(t.pct)}</td>
        <td><span class="info-status ${st.cls}" style="display:inline;background:transparent;border:0;border-radius:0;box-shadow:none;padding:0;font-size:27px;font-weight:900">${statusAr}</span></td>
      </tr>`;
  }).join('') : `<tr><td colspan="8">لا توجد صفقات ضمن الفترة المحددة</td></tr>`;

  return `
    <div class="infographic-card refined-light" id="${captureId}">
      <div class="info-top-swoosh"></div>
      <div class="info-bottom-swoosh"></div>

      <div class="info-header compact logo-only">
        <div class="info-logo-wrap wide"><img src="${logoSrc()}" class="info-logo bigger" alt="Q Options"></div>
      </div>

      <div class="info-stats refined-order">
        <div class="info-stat loss-stat">
          <div class="info-stat-top"><span class="info-icon loss">−</span><div class="info-label"><b>الصفقات الخاسرة</b></div></div>
          <div class="digital red">${s.losses.length}</div>
        </div>
        <div class="info-stat stopped-stat">
          <div class="info-stat-top"><span class="info-icon stopped">Ⅱ</span><div class="info-label"><b>الصفقات الموقوفة</b></div></div>
          <div class="digital cyan">${stoppedCount}</div>
        </div>
        <div class="info-stat win-stat">
          <div class="info-stat-top"><span class="info-icon win">✓</span><div class="info-label"><b>الصفقات الرابحة</b></div></div>
          <div class="digital green">${s.wins.length}</div>
        </div>
        <div class="info-stat total-trades-stat">
          <div class="info-stat-top"><span class="info-icon layers">≡</span><div class="info-label"><b>إجمالي الصفقات</b></div></div>
          <div class="digital number">${s.counted.length}</div>
        </div>
        <div class="info-stat totalprofit">
          <div class="info-stat-top"><span class="info-icon cash">$</span><div class="info-label"><b>إجمالي الأرباح</b></div></div>
          <div class="digital money bigmoney">${moneyInt(s.net)}</div>
        </div>
      </div>

      <div class="info-bottom">
        <div class="info-box distribution-box">
          <h4>توزيع نتائج الصفقات</h4>
          <div class="distribution-layout">
            <div class="distribution-donut-wrap">
              <svg class="distribution-donut" viewBox="0 0 140 140" aria-label="توزيع نتائج الصفقات">
                <circle class="distribution-depth" cx="70" cy="74" r="50" pathLength="100"/>
                <circle class="distribution-track" cx="70" cy="70" r="50" pathLength="100"/>
                <circle class="distribution-segment win" cx="70" cy="70" r="50" pathLength="100" stroke-dasharray="${distributionPct(s.wins.length).toFixed(2)} ${(100-distributionPct(s.wins.length)).toFixed(2)}" stroke-dashoffset="0"/>
                <circle class="distribution-segment stopped" cx="70" cy="70" r="50" pathLength="100" stroke-dasharray="${distributionPct(stoppedCount).toFixed(2)} ${(100-distributionPct(stoppedCount)).toFixed(2)}" stroke-dashoffset="-${distributionPct(s.wins.length).toFixed(2)}"/>
                <circle class="distribution-segment loss" cx="70" cy="70" r="50" pathLength="100" stroke-dasharray="${distributionPct(s.losses.length).toFixed(2)} ${(100-distributionPct(s.losses.length)).toFixed(2)}" stroke-dashoffset="-${(distributionPct(s.wins.length)+distributionPct(stoppedCount)).toFixed(2)}"/>
              </svg>
              <div class="distribution-total"><b>${s.counted.length}</b><span>إجمالي</span></div>
            </div>
            <div class="distribution-legend">
              <div class="legend-row win"><i></i><span>رابحة</span><b>${s.wins.length} — ${winPct.toFixed(1).replace('.0','')}%</b></div>
              <div class="legend-row stopped"><i></i><span>موقوفة</span><b>${stoppedCount} — ${stoppedPct.toFixed(1).replace('.0','')}%</b></div>
              <div class="legend-row loss"><i></i><span>خاسرة</span><b>${s.losses.length} — ${lossPct.toFixed(1).replace('.0','')}%</b></div>
            </div>
          </div>
        </div>

        <div class="info-box">
          <h4>العائد من كل صفقة</h4>
          <svg class="report-mini-chart cylinder-chart ${denseCharts?'dense-chart':''}" viewBox="0 0 330 150" role="img" aria-label="العائد من كل صفقة">
            <defs>
              <linearGradient id="barWin-${captureId}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#78ce91"/><stop offset=".5" stop-color="#31a65a"/><stop offset="1" stop-color="#247947"/></linearGradient>
              <linearGradient id="barStopped-${captureId}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8ce0ea"/><stop offset=".55" stop-color="#26adc2"/><stop offset="1" stop-color="#168298"/></linearGradient>
              <linearGradient id="barLoss-${captureId}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#ffaaa2"/><stop offset=".52" stop-color="#e65348"/><stop offset="1" stop-color="#ad302b"/></linearGradient>
            </defs>
            <g class="chart-grid"><line x1="12" y1="22" x2="320" y2="22"/><line x1="12" y1="47" x2="320" y2="47"/><line x1="12" y1="72" x2="320" y2="72"/><line x1="12" y1="97" x2="320" y2="97"/><line x1="12" y1="122" x2="320" y2="122"/></g>
            <line class="chart-zero" x1="12" y1="72" x2="320" y2="72"/>
            ${tradeBars}
          </svg>
        </div>

        <div class="info-box">
          <h4 class="equity-title">منحنى الأداء التراكمي</h4>
          <svg class="report-mini-chart equity-chart ${denseEquity?'dense-chart':''}" viewBox="0 0 330 150" role="img" aria-label="منحنى الأداء التراكمي للصفقات الرابحة">
            <defs><linearGradient id="equityFill-${captureId}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#d69b31" stop-opacity=".20"/><stop offset="1" stop-color="#d69b31" stop-opacity=".015"/></linearGradient></defs>
            <g class="chart-grid"><line x1="12" y1="28" x2="320" y2="28"/><line x1="12" y1="52" x2="320" y2="52"/><line x1="12" y1="76" x2="320" y2="76"/><line x1="12" y1="100" x2="320" y2="100"/><line x1="12" y1="123" x2="320" y2="123"/></g>
            <line class="chart-zero" x1="12" y1="123" x2="320" y2="123"/>
            <polygon points="16,123 ${equityPoints} 308,123" fill="url(#equityFill-${captureId})"/>
            <polyline points="${equityPoints}" fill="none" stroke="#a96f18" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>
            ${equityValues.map((v,i)=>{
              const px=16+i*(292/Math.max(1,equityValues.length-1));
              const py=122-(v-eqMin)/eqRange*94;
              const trade=equityTrades[i];
              const labelX=px+(i%2===0?-7:10);
              const labelY=py+(i%2===0?-7:11);
              const label=trade?.profit>0 ? `<text class="equity-symbol" x="${labelX.toFixed(1)}" y="${labelY.toFixed(1)}" text-anchor="start" transform="rotate(-90 ${labelX.toFixed(1)} ${labelY.toFixed(1)})">${escapeHtml(trade?.symbol||'').slice(0,5)}</text>` : '';
              return `<g class="equity-point"><circle cx="${px.toFixed(1)}" cy="${py.toFixed(1)}" r="4" fill="#fff4cf" stroke="#9c6414" stroke-width="3"/>${label}</g>`;
            }).join('')}
          </svg>
        </div>
      </div>

      <div class="infographic-table-panel ${rowsData.length>12?'many-trades':''}">
        <div class="info-table-head table-title-row">
          <div class="table-title-boxes unified-table-boxes">
            <h3 class="table-head-box all-trades-box">جميع الصفقات</h3>
            <div class="best-trade-badge table-head-box">
              <span class="best-cup">🏆</span>
              <span class="best-copy"><small>أفضل صفقة</small><b>${s.best ? `${escapeHtml(s.best.symbol)} ${moneyInt(s.best.profit)}` : '—'}</b></span>
            </div>
          </div>
          <div class="table-week-date table-head-box"><span class="week-date-icon">📅</span><span>${periodText}</span></div>
        </div>
        <table class="info-table roomy">
          <thead>
            <tr>
              <th>الرمز</th>
              <th>الخيار</th>
              <th>الاسترايك</th>
              <th>سعر الشراء</th>
              <th>سعر البيع</th>
              <th>الربح</th>
              <th>النسبة</th>
              <th>الحالة</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>

      <div class="info-contact footer-only">للتواصل عبر تليجرام <b dir="ltr">@Qalshammari</b></div>
    </div>`;
}


function setExportBusy(busy){
  ['pdfBtn','imageBtn'].forEach(id=>{if($(id)) $(id).disabled=busy});
}

function isIOSDevice(){
  return /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}

function clearNativeReportPrint(){
  document.body.classList.remove('print-top-trades-report');
  const stage=$('exportStage');
  if(stage) stage.innerHTML='';
  setExportBusy(false);
}

async function openNativePdfPrint(){
  const stage=$('exportStage');
  if(!stage) return;
  setExportBusy(true);
  showToast('جاري تجهيز صفحة أهم الصفقات PDF…');
  stage.innerHTML=buildShareTemplate(10,'nativePdfCapture');
  const target=$('nativePdfCapture');
  target?.classList.add('capture-mode');
  await waitForImages(target);
  // نحول التقرير إلى صورة واحدة بنفس أبعاده قبل الطباعة؛ هذا يمنع قص الجهة
  // اليسرى في Safari ويحافظ على نسبة التصميم كاملة داخل صفحة PDF.
  if(window.html2canvas && target){
    try{
      const canvas=await html2canvas(target,{
        scale:1.35,useCORS:true,allowTaint:false,
        backgroundColor:'#ffffff',logging:false,imageTimeout:8000
      });
      const reportImage=new Image();
      reportImage.id='nativePdfImage';
      reportImage.alt='Q Options report';
      reportImage.src=canvas.toDataURL('image/png');
      stage.innerHTML='';
      stage.appendChild(reportImage);
      await new Promise(resolve=>{
        if(reportImage.complete) return resolve();
        reportImage.onload=resolve;
        reportImage.onerror=resolve;
      });
    }catch(err){console.warn('Native PDF image fallback',err)}
  }
  document.body.classList.add('print-top-trades-report');

  const cleanup=()=>{
    window.removeEventListener('afterprint',cleanup);
    clearNativeReportPrint();
  };
  window.addEventListener('afterprint',cleanup,{once:true});
  setTimeout(()=>window.print(),260);
  // احتياط لـ Safari إذا لم يرسل afterprint بعد إغلاق نافذة الطباعة.
  setTimeout(()=>{
    if(document.body.classList.contains('print-top-trades-report')) cleanup();
  },120000);
}

async function exportPdf(){
  // على iPhone/iPad: نافذة الطباعة الأصلية أكثر ثباتاً، ومنها يمكن حفظ/مشاركة PDF.
  if(isIOSDevice()){
    await openNativePdfPrint();
    return;
  }

  // إذا لم تحمل مكتبات التصدير لأي سبب، لا يتعطل الزر: ننتقل مباشرة للطباعة.
  if(!window.html2canvas || !window.jspdf){
    await openNativePdfPrint();
    return;
  }

  setExportBusy(true);
  showToast('جاري تجهيز تقرير PDF…');
  let nativeFallback=false;
  try{
    const stage=$('exportStage');
    stage.innerHTML=buildPdfTemplate();
    const target=$('pdfCapture');
    await waitForImages(target);
    await new Promise(r=>setTimeout(r,150));

    const canvas=await html2canvas(target,{
      scale:1.75,
      useCORS:true,
      allowTaint:false,
      backgroundColor:'#f8f1e5',
      logging:false,
      imageTimeout:8000
    });
    const imgData=canvas.toDataURL('image/jpeg',0.94);
    const {jsPDF}=window.jspdf;
    const pageW=210;
    const pageH=pageW*canvas.height/canvas.width;
    const pdf=new jsPDF({orientation:'portrait',unit:'mm',format:[pageW,pageH],compress:true});
    pdf.addImage(imgData,'JPEG',0,0,pageW,pageH,undefined,'FAST');

    pdf.save(`Q-Options-Weekly-Report-${safeFileRange()}.pdf`);
    showToast('تم تجهيز تقرير PDF');
  }catch(err){
    console.error(err);
    nativeFallback=true;
    await openNativePdfPrint();
  }finally{
    if(!nativeFallback){
      $('exportStage').innerHTML='';
      setExportBusy(false);
    }
  }
}

let topTradesExportState = { dataUrl:'', blob:null, filename:'' };

function fitLivePreview(){
  const content=$('previewContent');
  const viewport=$('previewLiveViewport');
  const scaleBox=$('previewLiveScale');
  const card=$('liveShareCapture');
  if(!content || !viewport || !scaleBox || !card) return;

  const available=Math.max(280, content.getBoundingClientRect().width || 320);
  const cardWidth=REPORT_WIDTH;
  const cardHeight=Math.ceil(Math.max(card.scrollHeight,card.offsetHeight,card.getBoundingClientRect().height));
  const scale=Math.min(1, available/REPORT_WIDTH);

  scaleBox.style.transform=`translateX(-50%) scale(${scale})`;
  scaleBox.style.width=`${REPORT_WIDTH}px`;
  scaleBox.style.height=`${cardHeight}px`;
  viewport.style.height=`${Math.ceil(cardHeight*scale)}px`;
}

function showLivePreview(){
  const modal=$('previewModal');
  const content=$('previewContent');
  if(!modal || !content) return;

  topTradesExportState={dataUrl:'',blob:null,filename:''};
  content.innerHTML=`<div class="preview-live-viewport" id="previewLiveViewport"><div class="preview-live-scale" id="previewLiveScale">${buildShareTemplate(10,'liveShareCapture')}</div></div>`;
  $('liveShareCapture')?.classList.add('capture-mode');
  modal.hidden=false;
  document.body.classList.add('preview-open');

  requestAnimationFrame(()=>requestAnimationFrame(fitLivePreview));
  setTimeout(fitLivePreview,250);
  setTimeout(fitLivePreview,800);
}

function hidePreview(){
  const modal=$('previewModal');
  const content=$('previewContent');
  if(modal) modal.hidden=true;
  if(content) content.innerHTML='';
  document.body.classList.remove('preview-open');
  topTradesExportState={dataUrl:'',blob:null,filename:''};
}

function openTopTradesPreview(){
  showLivePreview();
  showToast('تم فتح أهم الصفقات');
}


async function shareOrDownloadBlob(blob,dataUrl,filename,iosWindow=null){
  try{
    if(blob && typeof File!=='undefined' && navigator.share){
      const file=new File([blob],filename,{type:'image/png'});
      if(!navigator.canShare || navigator.canShare({files:[file]})){
        await navigator.share({files:[file],title:'Q Options - أهم الصفقات'});
        if(iosWindow && !iosWindow.closed) iosWindow.close();
        return true;
      }
    }
  }catch(err){
    if(err?.name==='AbortError'){
      if(iosWindow && !iosWindow.closed) iosWindow.close();
      return true;
    }
    console.warn('Native share fallback:',err);
  }

  if(isIOSDevice()){
    const w=iosWindow && !iosWindow.closed ? iosWindow : window.open('', '_blank');
    if(w){
      w.document.open();
      w.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Q Options - أهم الصفقات</title><style>body{margin:0;background:#111;font-family:Arial,sans-serif} .hint{position:sticky;top:0;z-index:2;background:#fff8e8;color:#4b3418;text-align:center;padding:12px;font-weight:700} img{display:block;width:100%;height:auto;margin:0 auto}</style></head><body><div class="hint">اضغطي مطولاً على الصورة ثم اختاري حفظ في الصور</div><img src="${dataUrl}" alt="Q Options Top Trades"></body></html>`);
      w.document.close();
      showToast('فتحت الصورة للحفظ — اضغطي عليها مطولاً');
      return true;
    }
  }

  if(blob){
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    a.download=filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),3000);
  }else{
    const a=document.createElement('a');
    a.href=dataUrl;
    a.download=filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }
  return true;
}

async function saveOrShareTopTrades(e){
  e?.preventDefault?.();
  if(!window.html2canvas){showToast('مكتبة حفظ الصورة لم يتم تحميلها');return}
  // لا نفتح نافذة انتظار على الآيفون؛ يتم الحفظ من المعاينة نفسها.
  const iosWindow=null;

  const btn=$('previewDownload');
  const oldText=btn?.textContent || 'حفظ الصورة';
  if(btn){btn.disabled=true;btn.textContent='جاري تجهيز الصورة…'}
  showToast('جاري تجهيز الصورة عالية الدقة…');

  try{
    const stage=$('exportStage');
    // نلتقط نسخة مستقلة بالحجم الأصلي، وليس المعاينة المصغّرة داخل الجوال.
    // التقاط liveShareCapture كان يرث transform من المعاينة ويضغط التقرير في الزاوية.
    stage.innerHTML=buildShareTemplate(10,'shareCapture');
    const target=$('shareCapture');
    if(!target) throw new Error('لم يتم العثور على صفحة الصفقات');

    // Safari كان يحسب عرض التقرير بعرض شاشة الهاتف رغم أن لوحة الحفظ أكبر.
    // نثبت العرض على العنصر نفسه قبل أن يأخذ html2canvas القياسات.
    stage.style.setProperty('width',`${REPORT_WIDTH}px`,'important');
    stage.style.setProperty('min-width',`${REPORT_WIDTH}px`,'important');
    stage.style.setProperty('max-width',`${REPORT_WIDTH}px`,'important');
    stage.style.setProperty('display','block','important');
    stage.style.setProperty('position','fixed','important');
    stage.style.setProperty('left','0','important');
    stage.style.setProperty('top','0','important');
    target.style.setProperty('width',`${REPORT_WIDTH}px`,'important');
    target.style.setProperty('min-width',`${REPORT_WIDTH}px`,'important');
    target.style.setProperty('max-width',`${REPORT_WIDTH}px`,'important');
    target.style.setProperty('display','block','important');
    target.style.setProperty('position','relative','important');
    target.style.setProperty('margin','0','important');
    target.style.setProperty('box-sizing','border-box','important');
    target.style.setProperty('flex','none','important');
    target.style.setProperty('transform','none','important');
    target.style.setProperty('zoom','1','important');
    target.classList.add('capture-mode');
    await Promise.race([waitForImages(target),new Promise(r=>setTimeout(r,1200))]);
    await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));

    const measuredWidth=Math.round(target.getBoundingClientRect().width);
    if(measuredWidth<REPORT_WIDTH-20){
      throw new Error(`عرض التقرير غير صحيح قبل الحفظ: ${measuredWidth}px`);
    }

    const captureWidth=REPORT_WIDTH;
    const captureHeight=Math.ceil(Math.max(target.scrollHeight,target.getBoundingClientRect().height));
    // دقة أعلى للصورة النهائية مع حد آمن لذاكرة Safari على الآيفون.
    const scale=2;
    const canvas=await html2canvas(target,{
      scale,
      width:captureWidth,
      height:captureHeight,
      useCORS:true,
      allowTaint:false,
      backgroundColor:'#fffefa',
      logging:false,
      imageTimeout:1200,
      windowWidth:captureWidth,
      windowHeight:captureHeight,
      scrollX:0,
      scrollY:0,
      onclone:clonedDocument=>{
        clonedDocument.documentElement.style.setProperty('width',`${REPORT_WIDTH}px`,'important');
        clonedDocument.documentElement.style.setProperty('min-width',`${REPORT_WIDTH}px`,'important');
        clonedDocument.body.style.setProperty('width',`${REPORT_WIDTH}px`,'important');
        clonedDocument.body.style.setProperty('min-width',`${REPORT_WIDTH}px`,'important');
        clonedDocument.body.style.setProperty('margin','0','important');
        const clonedStage=clonedDocument.getElementById('exportStage');
        const clonedTarget=clonedDocument.getElementById('shareCapture');
        if(clonedStage){
          clonedStage.style.setProperty('left','0','important');
          clonedStage.style.setProperty('width',`${REPORT_WIDTH}px`,'important');
          clonedStage.style.setProperty('min-width',`${REPORT_WIDTH}px`,'important');
          clonedStage.style.setProperty('max-width',`${REPORT_WIDTH}px`,'important');
          clonedStage.style.setProperty('transform','none','important');
        }
        if(clonedTarget){
          clonedTarget.classList.add('capture-mode');
          clonedTarget.style.setProperty('width',`${REPORT_WIDTH}px`,'important');
          clonedTarget.style.setProperty('min-width',`${REPORT_WIDTH}px`,'important');
          clonedTarget.style.setProperty('max-width',`${REPORT_WIDTH}px`,'important');
          clonedTarget.style.setProperty('display','block','important');
          clonedTarget.style.setProperty('position','relative','important');
          clonedTarget.style.setProperty('margin','0','important');
          clonedTarget.style.setProperty('box-sizing','border-box','important');
          clonedTarget.style.setProperty('flex','none','important');
          clonedTarget.style.setProperty('transform','none','important');
          clonedTarget.style.setProperty('zoom','1','important');
        }
      }
    });

    if(!canvas.width || !canvas.height) throw new Error('تعذر إنشاء مساحة الصورة');
    const dataUrl=canvas.toDataURL('image/png');
    let blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
    if(!dataUrl) throw new Error('PNG generation failed');
    if(!blob){
      const binary=atob(dataUrl.split(',')[1]);
      const bytes=new Uint8Array(binary.length);
      for(let i=0;i<binary.length;i++) bytes[i]=binary.charCodeAt(i);
      blob=new Blob([bytes],{type:'image/png'});
    }
    const filename=`Q-Options-صفحة-الصفقات-${safeFileRange()}.png`;
    topTradesExportState={dataUrl,blob,filename};

    // نعرض النسخة النهائية داخل المعاينة أولاً؛ وهذا يضمن وجود طريقة
    // للحفظ بالضغط المطوّل حتى إذا منع Safari التنزيل أو المشاركة.
    const previewContent=$('previewContent');
    if(previewContent){
      previewContent.innerHTML=`<div class="saved-image-wrap"><div class="saved-image-hint">اضغط مطولًا على الصورة ثم اختر «حفظ في الصور» إذا لم تظهر نافذة المشاركة</div><img class="saved-trades-image" src="${dataUrl}" alt="صفحة الصفقات"></div>`;
    }
    if(!blob) throw new Error('تعذر إنشاء ملف الصورة');
    if(isIOSDevice()){
      showToast('تم تجهيز الصورة — اضغط عليها مطولًا للحفظ');
    }else{
      await shareOrDownloadBlob(blob,dataUrl,filename,iosWindow);
      showToast('تم حفظ الصورة');
    }
  }catch(err){
    console.error(err);
    showToast('تعذر الحفظ التلقائي — جرّب الضغط مطولًا على المعاينة');
  }finally{
    $('exportStage').innerHTML='';
    if(btn){btn.disabled=false;btn.textContent=oldText}
  }
}


let toastTimer;
function showToast(message){
  const el=$('toast');
  el.textContent=message;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>el.classList.remove('show'),2600);
}

function formatArabicDate(iso){
  if(!iso) return '';
  const [y,m,d]=iso.split('-');
  return `${d}/${m}/${y}`;
}

function updateManualWeekRange(){
  const from=$('fromDate').value || currentWeekRange().from;
  const to=$('toDate').value || currentWeekRange().to;
  $('manualWeekRange').textContent=`الأسبوع الحالي: ${formatArabicDate(from)} إلى ${formatArabicDate(to)}`;
}

function renderManualTrades(){
  const body=$('manualTradesBody');
  body.innerHTML=manualTrades.map((t,i)=>`<tr><td>${escapeHtml(t.symbol)}</td><td>${escapeHtml(t.option)}</td><td>${escapeHtml(t.strike)}</td><td>${Number(t.buy).toFixed(2)}</td><td>${Number(t.sell).toFixed(2)}</td><td>${escapeHtml(t.notes||'—')}</td><td><button type="button" class="manual-delete" data-index="${i}">حذف</button></td></tr>`).join('');
  $('manualCount').textContent=`${manualTrades.length} صفقة`;
  $('manualEmpty').hidden=manualTrades.length>0;
}

function openManualEntry(){
  updateManualWeekRange();
  renderManualTrades();
  $('manualEntryPage').hidden=false;
  document.body.style.overflow='hidden';
}

function closeManualEntry(){
  $('manualEntryPage').hidden=true;
  document.body.style.overflow='';
}

function applyManualTrades(closePage=true){
  if(!manualTrades.length){showToast('أضف صفقة واحدة على الأقل');return false}
  importedWorkbookActive=true;
  trades=manualTrades.map(t=>({...t}));
  render();
  if(closePage) closeManualEntry();
  showToast('تم تحديث التقرير بالصفقات المدخلة');
  return true;
}

$('excelFile').addEventListener('change',e=>{const f=e.target.files[0];if(f) handleFile(f)});
$('demoBtn').addEventListener('click',openManualEntry);
$('manualClose').addEventListener('click',closeManualEntry);
$('manualTradeForm').addEventListener('submit',e=>{
  e.preventDefault();
  const buy=num($('manualBuy').value),sell=num($('manualSell').value);
  const date=$('fromDate').value || currentWeekRange().from;
  manualTrades.push({date,symbol:$('manualSymbol').value.trim().toUpperCase(),option:$('manualOption').value,strike:$('manualStrike').value,buy,sell,profit:(sell-buy)*100,pct:buy?(sell-buy)/buy*100:0,notes:$('manualNotes').value.trim()});
  e.target.reset();
  renderManualTrades();
  $('manualSymbol').focus();
  showToast('تمت إضافة الصفقة');
});
$('manualTradesBody').addEventListener('click',e=>{const btn=e.target.closest('.manual-delete');if(!btn)return;manualTrades.splice(Number(btn.dataset.index),1);renderManualTrades()});
$('manualApply').addEventListener('click',()=>applyManualTrades(true));
$('manualPdf').addEventListener('click',()=>{if(applyManualTrades(true)) exportPdf()});
$('manualImage').addEventListener('click',()=>{if(applyManualTrades(true)) openTopTradesPreview()});
$('pdfBtn').addEventListener('click',exportPdf);
$('imageBtn').addEventListener('click',openTopTradesPreview);
$('previewClose')?.addEventListener('click',hidePreview);
$('previewDownload')?.addEventListener('click',saveOrShareTopTrades);
$('previewModal')?.addEventListener('click',e=>{ if(e.target.id==='previewModal') hidePreview(); });
window.addEventListener('resize',()=>{if(!$('previewModal')?.hidden) fitLivePreview()});
$('fromDate').addEventListener('change',()=>{saveSelectedRange();render()});
$('toDate').addEventListener('change',()=>{saveSelectedRange();render()});

if(!restoreSelectedRange()){
  setCurrentWeekRange();
  saveSelectedRange();
}
trades=buildDemoTrades();
updateFooterClock();
setInterval(updateFooterClock,60000);
render();

// Q OPTIONS FINAL READY — live preview + iOS save
console.log('Q OPTIONS BUILD v7.1 LIVE PREVIEW');
