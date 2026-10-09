// If the Python API is hosted on another domain, set API_BASE to that origin.
const API_BASE = '';
const $ = id => document.getElementById(id);
let latestRows = [];
const asNum = n => n===null||n===undefined||n===''?NaN:Number(n);
const fmtUsd = n => Number.isFinite(Number(n)) ? '$' + Number(n).toLocaleString('en-US',{maximumFractionDigits:0}) : '—';
const fmtPrice = n => Number.isFinite(Number(n)) ? '$' + Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}) : '—';
const fmtDays = n => Number.isFinite(Number(n)) ? Number(n).toLocaleString('en-US',{maximumFractionDigits:1}) : '—';
const fmtPct = n => Number.isFinite(asNum(n)) ? `${asNum(n)>0?'+':''}${asNum(n).toFixed(2)}%` : '—';
const fmtMult = n => Number.isFinite(Number(n)) ? Number(n).toFixed(2) : '—';
function setNotice(message, kind=''){ $('notice').className='notice '+kind; $('notice').textContent=message; }
function renderRows(){
  const query=$('search').value.trim().toUpperCase();
  const rows=latestRows.filter(row=>String(row.ticker||'').toUpperCase().includes(query));
  $('resultCount').textContent=`عرض ${rows.length.toLocaleString('en-US')} من ${latestRows.length.toLocaleString('en-US')} سهم`;
  $('results').innerHTML=rows.length?rows.map(row=>`<tr>
    <td><span class="ticker">${escapeHtml(row.ticker)}</span></td>
    <td>${fmtPct(row.daily_change_pct)}</td><td>${fmtMult(row.relative_volume)}×</td><td>${fmtPrice(row.price)}</td>
    <td>${fmtUsd(row.last_dollar_flow)}</td><td>${fmtUsd(row.avg_dollar_vol)}</td>
    <td>${fmtUsd(row.median_dollar_vol)}</td><td>${fmtDays(row.days_to_liquidate)}</td>
    <td><span class="tier ${escapeHtml(row.tier)}">${tierLabel(row.tier)}</span></td>
  </tr>`).join(''):'<tr><td class="empty" colspan="9">لا توجد أسهم مطابقة للبحث.</td></tr>';
}
function escapeHtml(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function tierLabel(t){return ({high:'مرتفعة',medium:'متوسطة',low:'منخفضة',illiquid:'ضعيفة'})[t]||'—';}
function renderSummary(data){
  $('scanned').textContent=Number(data.scanned||0).toLocaleString('en-US');
  $('flow').textContent=fmtUsd(data.total_last_dollar_flow);
  $('high').textContent=Number(data.tiers?.high||0).toLocaleString('en-US');
  $('medium').textContent=Number(data.tiers?.medium||0).toLocaleString('en-US');
  $('low').textContent=(Number(data.tiers?.low||0)+Number(data.tiers?.illiquid||0)).toLocaleString('en-US');
  $('updated').textContent=data.updated_at?new Date(data.updated_at).toLocaleString('ar-SA',{dateStyle:'short',timeStyle:'short'}):'—';
  for(const ticker of ['SPY','QQQ','DIA','IWM']){
    const node=$(`pulse${ticker}`),move=asNum(data.market_pulse?.[ticker]?.change_pct);
    node.textContent=Number.isFinite(move)?fmtPct(move):'—';node.className=Number.isFinite(move)?(move>0?'up':move<0?'down':''):'';
  }
  latestRows=Array.isArray(data.results)?data.results:[];renderRows();
  $('resultCount').textContent=`اجتازت المعايير الأربعة: ${data.eligible||0} من ${data.screened||0} سهمًا معروفًا · مصدر التصنيف: ${data.screen_source||'Mizan'}${data.screen_as_of?` · آخر تاريخ بيانات ${String(data.screen_as_of).slice(0,10)}`:''}${data.screen_checks_failed?` · تعذر التحقق من ${data.screen_checks_failed} رموز`:''}`;
}
async function refreshScan(){
  const button=$('refreshBtn');button.disabled=true;button.querySelector('span').classList.add('spin');
  setNotice('يجري التحقق من قائمة الأسهم المعروفة وفق أربعة معايير شرعية، ثم تُجلب بيانات التداول للأسهم المستوفية فقط.','loading');
  try{
    const response=await fetch(`${API_BASE}/api/liquidity`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({position_usd:Number($('position').value)||50000,min_dollar_vol:Number($('minDollarVol').value)})});
    const data=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(data.error||`تعذر إكمال الفحص (${response.status})`);
    renderSummary(data);setNotice(`اكتمل الفحص: عُرض ${latestRows.length} سهمًا مستوفيًا بعد فحص ${data.screened||0} سهمًا معروفًا. آخر جلسة أسعار: ${data.latest_session||'—'}؛ التدفق المعروض قيمة التداول اليومية.`);
  }catch(error){setNotice(`${error.message||'تعذر الاتصال بخدمة الفحص.'} تأكدي من تشغيل خدمة السيولة وربط رابطها في liquidity.js.`,'error');}
  finally{button.disabled=false;button.querySelector('span').classList.remove('spin');}
}
$('refreshBtn').addEventListener('click',refreshScan);$('search').addEventListener('input',renderRows);
$('eventsBtn').addEventListener('click',refreshEvents);
async function refreshEvents(){
  const panel=$('eventsPanel'),button=$('eventsBtn');panel.hidden=false;button.disabled=true;
  $('eventsList').innerHTML='<p class="events-empty">جاري جلب الإصدارات الاقتصادية المجدولة لهذا الأسبوع…</p>';
  try{
    const selectedRange=JSON.parse(localStorage.getItem('qOptionsSelectedReportRange')||'null');
    const params=new URLSearchParams();if(/^\d{4}-\d{2}-\d{2}$/.test(selectedRange?.from||''))params.set('anchor',selectedRange.from);
    const response=await fetch(`${API_BASE}/api/events${params.size?`?${params}`:''}`),data=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(data.error||`تعذر جلب الأحداث (${response.status})`);
    $('eventsUpdated').textContent=`الأسبوع ${data.week_start} إلى ${data.week_end} · المصادر: ${(data.sources_checked||[]).join('، ')} · آخر تحقق ${new Date(data.updated_at).toLocaleTimeString('ar-SA',{hour:'2-digit',minute:'2-digit'})}`;
    $('eventsList').innerHTML=data.events?.length?data.events.map(event=>`<article class="event-row">
      <div class="event-date">${new Date(event.datetime).toLocaleString('ar-SA',event.all_day?{weekday:'long',month:'short',day:'numeric'}:{weekday:'long',month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'})}<small>${new Date(event.datetime)<new Date()?'موعد سابق':'موعد قادم'}${event.all_day?' · الوقت غير محدد':''}</small></div>
      <div><div class="event-title">${escapeHtml(event.title)}</div><a class="event-source" href="${escapeHtml(event.source_url)}" target="_blank" rel="noopener">المصدر: ${escapeHtml(event.source)}</a></div>
      <span class="event-impact ${event.impact==='مرتفع'?'high':''}">${escapeHtml(event.impact)}</span>
    </article>`).join(''):'<p class="events-empty">ما فيه إصدارات اقتصادية رئيسية مسجلة في المصادر لهذا الأسبوع.</p>';
  }catch(error){$('eventsList').innerHTML=`<p class="events-empty">${escapeHtml(error.message||'تعذر الاتصال بمصادر الأخبار.')}</p>`;}
  finally{button.disabled=false;}
}
