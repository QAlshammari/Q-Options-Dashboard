let trades = [];
let equityChart, dailyChart, winLossChart;

const $ = id => document.getElementById(id);
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
    const profitRaw = pick(['profit','p/l','pnl','net profit','الربح','الربح والخسارة','صافي الربح','ربح','الخسارة'],5);
    const profit = profitRaw === '' || profitRaw === null || profitRaw === undefined
      ? ((sell !== null ? sell-buy : 0)*100)
      : num(profitRaw);
    const pctRaw = pick(['pct','percent','percentage','profit %','return %','النسبة','النسبة %','نسبة الربح','نسبة العائد'],6);
    const p = pctRaw === '' || pctRaw === null || pctRaw === undefined
      ? (buy ? ((sell ?? buy)-buy)/buy*100 : 0)
      : num(pctRaw);
    const notes = pick(['notes','note','remarks','الملاحظات','ملاحظات','ملاحظة'],7) || '';

    return {date,symbol:String(symbol).trim(),option,strike:String(strike).trim(),buy,sell,profit,pct:p,notes:String(notes).trim()};
  }).filter(x => x.symbol !== '—' || x.date);
}

function colorize(el,value){
  el.classList.remove('positive','negative','neutral');
  el.classList.add(value > 0 ? 'positive' : value < 0 ? 'negative' : 'neutral');
}

function getFilteredTrades(){
  const from = $('fromDate').value;
  const to = $('toDate').value;
  return trades.filter(t => (!from || !t.date || t.date >= from) && (!to || !t.date || t.date <= to));
}

function calculateStats(data){
  const closed = data.filter(t => t.sell !== null || t.profit !== 0);
  const wins = closed.filter(t => t.profit > 0);
  const losses = closed.filter(t => t.profit < 0);
  const grossWin = wins.reduce((s,t)=>s+t.profit,0);
  const grossLoss = losses.reduce((s,t)=>s+t.profit,0);
  const net = closed.reduce((s,t)=>s+t.profit,0);
  const totalCost = closed.reduce((s,t)=>s+(Math.abs(t.buy)*100),0);
  const returnP = totalCost ? net/totalCost*100 : 0;
  const winRate = closed.length ? wins.length/closed.length*100 : 0;
  const avgWin = wins.length ? grossWin/wins.length : 0;
  const avgLoss = losses.length ? grossLoss/losses.length : 0;
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

  return {closed,wins,losses,grossWin,grossLoss,net,totalCost,returnP,winRate,avgWin,avgLoss,pf,expectancy,maxDD,sharpe,ordered,byDay,bestDay,worstDay,best};
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
  renderCharts(s.ordered,s.byDay,s.wins.length,s.losses.length);
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
}

function renderCharts(ordered,byDay,winCount,lossCount){
  if(!window.Chart) return;
  chartDefaults();
  const eqLabels=[],eqData=[];let cum=0;
  ordered.forEach((t,i)=>{cum+=t.profit;eqLabels.push(t.date||String(i+1));eqData.push(cum)});
  const days=Object.keys(byDay).sort(),vals=days.map(d=>byDay[d]);

  equityChart?.destroy();dailyChart?.destroy();winLossChart?.destroy();
  equityChart=new Chart($('equityChart'),{
    type:'line',
    data:{labels:eqLabels,datasets:[{data:eqData,borderColor:'#b88a3b',backgroundColor:'rgba(184,138,59,.14)',fill:true,tension:.28,pointRadius:2,borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,plugins:{legend:{display:false}},scales:{x:{ticks:{maxRotation:0,autoSkip:true}},y:{ticks:{callback:v=>'$'+v}}}}
  });
  dailyChart=new Chart($('dailyChart'),{
    type:'bar',
    data:{labels:days,datasets:[{data:vals,backgroundColor:vals.map(v=>v>=0?'#4d9664':'#c45b52'),borderWidth:0,borderRadius:5}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>'$'+v}}}}
  });
  winLossChart=new Chart($('winLossChart'),{
    type:'doughnut',
    data:{labels:['أرباح','خسائر'],datasets:[{data:[winCount,lossCount],backgroundColor:['#4d9664','#c45b52'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,cutout:'58%',plugins:{legend:{position:'right',rtl:true,labels:{boxWidth:13}}}}
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

      // نجرب جميع الأوراق حتى نجد ورقة تحتوي على صفقات صالحة.
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

    trades=normalized;
    setRangeFromTrades();
    render();
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


function tradeOptionLabel(trade){
  const explicit=String(trade.option||'').toUpperCase();
  if(explicit==='CALL' || explicit==='PUT') return explicit;
  const notes=(trade.notes||'').toLowerCase();
  if(notes.includes('put')) return 'PUT';
  if(notes.includes('call')) return 'CALL';
  return trade.pct >= 0 ? 'CALL' : 'PUT';
}

function tradeStatusMeta(trade){
  if(trade.sell===null && trade.profit===0) return {cls:'open',labelAr:'مفتوحة',labelEn:'OPEN',icon:'◐'};
  if(trade.profit < 0) return {cls:'loss',labelAr:'خاسرة',labelEn:'LOSS',icon:'✕'};
  return {cls:'win',labelAr:'رابح',labelEn:'WIN',icon:'✓'};
}

function buildPdfTemplate(){
  return buildShareTemplate(10, "pdfCapture");
}


function buildShareTemplate(maxRows=10, captureId="shareCapture"){
  const filtered=getFilteredTrades();
  const s=calculateStats(filtered);
  const rowsData=topTrades(filtered,maxRows);
  const periodText=`${$('fromDate').value || '—'}  →  ${$('toDate').value || '—'}`;
  const winPct = s.closed.length ? (s.wins.length / s.closed.length * 100) : 0;
  const neutralCount = Math.max(0, filtered.length - s.wins.length - s.losses.length);
  const rows = rowsData.length ? rowsData.map(t=>{
    const option=tradeOptionLabel(t);
    const st=tradeStatusMeta(t);
    const statusAr = st.cls==='loss' ? 'خاسرة' : st.cls==='open' ? 'مفتوحة' : 'رابح';
    return `
      <tr>
        <td class="symbol-cell"><div class="symbol-stack"><span>${escapeHtml(t.symbol)}</span></div></td>
        <td><span class="info-chip ${option==='CALL'?'call':'put'}">${option}</span></td>
        <td>${escapeHtml(t.strike)}</td>
        <td dir="ltr">${money(t.buy)}</td>
        <td dir="ltr">${t.sell===null?'—':money(t.sell)}</td>
        <td class="info-profit ${t.profit>=0?'pos':'neg'}">${t.sell===null&&t.profit===0?'—':money(t.profit)}</td>
        <td class="info-pct ${t.pct>=0?'pos':'neg'}">${t.sell===null&&t.profit===0?'—':pct(t.pct)}</td>
        <td><span class="info-status ${st.cls}">${statusAr}</span></td>
      </tr>`;
  }).join('') : `<tr><td colspan="8">لا توجد صفقات ضمن الفترة المحددة</td></tr>`;

  return `
    <div class="infographic-card refined-light" id="${captureId}">
      <div class="info-top-swoosh"></div>
      <div class="info-bottom-swoosh"></div>

      <div class="info-header compact logo-only">
        <div class="info-logo-wrap wide"><img src="${logoSrc()}" class="info-logo bigger" alt="Q Options"></div>
      </div>

      <div class="info-dates-under-logo">
        <div class="period-item solo"><span class="period-badge">📅</span><div><div class="digital smallish">${periodText}</div><div class="period-meta">فترة التداول  Trading Period</div></div></div>
      </div>

      <div class="info-stats refined-order">
        <div class="info-stat totalprofit">
          <div class="info-stat-top"><span class="info-icon cash">$</span><div class="info-label"><b>إجمالي الأرباح</b><small>TOTAL PROFIT</small></div></div>
          <div class="digital money bigmoney">${moneyInt(s.net)}</div>
        </div>
        <div class="info-stat">
          <div class="info-stat-top"><span class="info-icon loss">✕</span><div class="info-label"><b>الصفقات الخاسرة</b><small>LOSING TRADES</small></div></div>
          <div class="digital red">${s.losses.length}</div>
        </div>
        <div class="info-stat">
          <div class="info-stat-top"><span class="info-icon win">✓</span><div class="info-label"><b>الصفقات الرابحة</b><small>WINNING TRADES</small></div></div>
          <div class="digital green">${s.wins.length}</div>
        </div>
        <div class="info-stat featured">
          <div class="info-stat-top"><div class="info-label"><b>إجمالي العائد</b><small>TOTAL RETURN</small></div></div>
          <div class="digital green">${pct(s.returnP)}</div>
        </div>
        <div class="info-stat">
          <div class="info-stat-top"><span class="info-icon layers">▤</span><div class="info-label"><b>إجمالي الصفقات</b><small>TOTAL TRADES</small></div></div>
          <div class="digital number">${filtered.length}</div>
        </div>
      </div>

      <div class="infographic-table-panel">
        <div class="info-table-head centered"><h3>أهم الصفقات</h3></div>
        <table class="info-table roomy">
          <thead>
            <tr>
              <th>الرمز<br><small>SYMBOL</small></th>
              <th>الخيار<br><small>OPTION</small></th>
              <th>الضربة<br><small>STRIKE</small></th>
              <th>سعر الشراء<br><small>BUY</small></th>
              <th>سعر البيع<br><small>SELL</small></th>
              <th>الربح<br><small>PROFIT $</small></th>
              <th>الربح %<br><small>PROFIT %</small></th>
              <th>الحالة<br><small>STATUS</small></th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>

      <div class="info-bottom">
        <div class="info-box">
          <h4>إجمالي العائد<small>TOTAL RETURN</small></h4>
          <div class="return-ring-wrap">
            <div class="return-ring">
              <div class="return-ring-arrow">↗</div>
              <div class="return-ring-content">
                <div class="digital">${pct(s.returnP)}</div>
                <div class="ar">إجمالي العائد</div>
              </div>
            </div>
          </div>
        </div>

        <div class="info-box">
          <h4>إحصائيات سريعة<small>QUICK STATISTICS</small></h4>
          <div class="quick-list">
            <div class="quick-row"><span class="name"><span class="mini blue">≣</span> إجمالي الصفقات</span><span class="value">${filtered.length}</span></div>
            <div class="quick-row"><span class="name"><span class="mini green">✓</span> الصفقات الرابحة</span><span class="value green">${s.wins.length}</span></div>
            <div class="quick-row"><span class="name"><span class="mini red">✕</span> الصفقات الخاسرة</span><span class="value red">${s.losses.length}</span></div>
            <div class="quick-row"><span class="name"><span class="mini gray">◐</span> المحايدة</span><span class="value">${neutralCount}</span></div>
            <div class="quick-row"><span class="name"><span class="mini gold">$</span> إجمالي الأرباح</span><span class="value green">${moneyInt(s.net)}</span></div>
          </div>
        </div>

        <div class="info-box">
          <h4>توزيع نتائج الصفقات<small>TRADE DISTRIBUTION</small></h4>
          <div class="donut-layout">
            <div style="position:relative;width:170px;height:170px;display:grid;place-items:center">
              <div class="donut-chart" style="--pct:${winPct.toFixed(2)}"></div>
              <div class="donut-core"><b>${winPct.toFixed(0)}%</b><span>رابح</span></div>
            </div>
            <div class="donut-legend">
              <div class="legend-row"><span style="display:flex;align-items:center;gap:8px"><span class="legend-dot green"></span> رابحة</span><b>${s.wins.length}</b></div>
              <div class="legend-row"><span style="display:flex;align-items:center;gap:8px"><span class="legend-dot red"></span> خاسرة</span><b>${s.losses.length}</b></div>
              <div class="legend-row"><span style="display:flex;align-items:center;gap:8px"><span class="legend-dot gray"></span> محايدة</span><b>${neutralCount}</b></div>
            </div>
          </div>
        </div>
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

function openNativePdfPrint(){
  showToast('فتح نافذة PDF / الطباعة…');
  setTimeout(()=>window.print(),180);
}

async function exportPdf(){
  // على iPhone/iPad: نافذة الطباعة الأصلية أكثر ثباتاً، ومنها يمكن حفظ/مشاركة PDF.
  if(isIOSDevice()){
    openNativePdfPrint();
    return;
  }

  // إذا لم تحمل مكتبات التصدير لأي سبب، لا يتعطل الزر: ننتقل مباشرة للطباعة.
  if(!window.html2canvas || !window.jspdf){
    openNativePdfPrint();
    return;
  }

  setExportBusy(true);
  showToast('جاري تجهيز تقرير PDF…');
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
    const imgData=canvas.toDataURL('image/jpeg',0.92);
    const {jsPDF}=window.jspdf;
    const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4',compress:true});
    const pageW=pdf.internal.pageSize.getWidth();
    const pageH=pdf.internal.pageSize.getHeight();
    const margin=7;
    const contentW=pageW-margin*2;
    const renderedH=canvas.height*contentW/canvas.width;
    let offset=0;

    while(offset<renderedH){
      if(offset>0) pdf.addPage('a4','landscape');
      pdf.addImage(imgData,'JPEG',margin,margin-offset,contentW,renderedH,undefined,'FAST');
      offset+=pageH-margin*2;
    }

    pdf.save(`Q-Options-Weekly-Report-${safeFileRange()}.pdf`);
    showToast('تم تجهيز تقرير PDF');
  }catch(err){
    console.error(err);
    openNativePdfPrint();
  }finally{
    $('exportStage').innerHTML='';
    setExportBusy(false);
  }
}

let topTradesExportState = { dataUrl:'', blob:null, filename:'' };

function fitLivePreview(){
  const content=$('previewContent');
  const viewport=$('previewLiveViewport');
  const scaleBox=$('previewLiveScale');
  const card=$('liveShareCapture');
  if(!content || !viewport || !scaleBox || !card) return;

  const available=Math.max(280, content.clientWidth || 320);
  const cardWidth=card.offsetWidth || 1240;
  const cardHeight=card.offsetHeight || 1740;
  const scale=Math.min(1, available/cardWidth);

  scaleBox.style.transform=`translateX(-50%) scale(${scale})`;
  viewport.style.height=`${Math.ceil(cardHeight*scale)}px`;
}

function showLivePreview(){
  const modal=$('previewModal');
  const content=$('previewContent');
  if(!modal || !content) return;

  topTradesExportState={dataUrl:'',blob:null,filename:''};
  content.innerHTML=`<div class="preview-live-viewport" id="previewLiveViewport"><div class="preview-live-scale" id="previewLiveScale">${buildShareTemplate(10,'liveShareCapture')}</div></div>`;
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

  // فتح نافذة iOS فور ضغطة المستخدم، قبل أي await، حتى لا يمنع Safari الحفظ لاحقاً.
  const iosWindow=isIOSDevice()?window.open('', '_blank'):null;
  if(iosWindow){
    iosWindow.document.write('<!doctype html><html lang="ar" dir="rtl"><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>جاري تجهيز الصورة</title></head><body style="font-family:Arial;text-align:center;padding:50px;background:#fffaf2;color:#4b3418"><h2>جاري تجهيز صورة أهم الصفقات…</h2><p>انتظري لحظات.</p></body></html>');
    iosWindow.document.close();
  }

  const btn=$('previewDownload');
  const oldText=btn?.textContent || 'حفظ / مشاركة الصورة';
  if(btn){btn.disabled=true;btn.textContent='جاري تجهيز الصورة…'}
  showToast('جاري تجهيز الصورة عالية الدقة…');

  try{
    const stage=$('exportStage');
    stage.innerHTML=buildShareTemplate(10,'shareCapture');
    const target=$('shareCapture');
    await waitForImages(target);
    await new Promise(r=>setTimeout(r,180));

    const scale=isIOSDevice()?1:1.55;
    const canvas=await html2canvas(target,{
      scale,
      useCORS:true,
      allowTaint:false,
      backgroundColor:'#fffefa',
      logging:false,
      imageTimeout:15000,
      width:target.scrollWidth || 1120,
      height:target.scrollHeight || 1570,
      windowWidth:target.scrollWidth || 1120,
      scrollX:0,
      scrollY:0
    });

    const dataUrl=canvas.toDataURL('image/png',0.98);
    const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png',0.98));
    if(!dataUrl) throw new Error('PNG generation failed');

    const filename=`Q-Options-Top-Trades-${safeFileRange()}.png`;
    topTradesExportState={dataUrl,blob,filename};
    await shareOrDownloadBlob(blob,dataUrl,filename,iosWindow);
    showToast('تم تجهيز الصورة');
  }catch(err){
    console.error(err);
    if(iosWindow && !iosWindow.closed){
      iosWindow.document.open();
      iosWindow.document.write('<!doctype html><html lang="ar" dir="rtl"><meta name="viewport" content="width=device-width,initial-scale=1"><body style="font-family:Arial;text-align:center;padding:40px"><h2>تعذر حفظ الصورة</h2><p>ارجعي للصفحة وحاولي مرة أخرى.</p></body></html>');
      iosWindow.document.close();
    }
    showToast('تعذر حفظ الصورة');
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

$('excelFile').addEventListener('change',e=>{const f=e.target.files[0];if(f) handleFile(f)});
$('demoBtn').addEventListener('click',()=>{setCurrentWeekRange();trades=buildDemoTrades();render();showToast('تم تحميل بيانات تجريبية للأسبوع الحالي')});
$('pdfBtn').addEventListener('click',exportPdf);
$('imageBtn').addEventListener('click',openTopTradesPreview);
$('previewClose')?.addEventListener('click',hidePreview);
$('previewDownload')?.addEventListener('click',saveOrShareTopTrades);
$('previewModal')?.addEventListener('click',e=>{ if(e.target.id==='previewModal') hidePreview(); });
window.addEventListener('resize',()=>{if(!$('previewModal')?.hidden) fitLivePreview()});
$('fromDate').addEventListener('change',render);
$('toDate').addEventListener('change',render);

setCurrentWeekRange();
trades=buildDemoTrades();
updateFooterClock();
setInterval(updateFooterClock,60000);
render();

// Q OPTIONS FINAL READY — live preview + iOS save
console.log('Q OPTIONS BUILD v7.1 LIVE PREVIEW');
