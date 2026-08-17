let trades = [];
let equityChart, dailyChart, winLossChart;

const demoTrades = [
  {date:'2025-05-18',symbol:'NVDA',strike:'1000',buy:12.50,sell:18.90,profit:640,pct:51.20,notes:'Breakout play'},
  {date:'2025-05-19',symbol:'SPY',strike:'532',buy:6.30,sell:10.80,profit:450,pct:71.43,notes:'Bull trend follow'},
  {date:'2025-05-19',symbol:'AAPL',strike:'195',buy:4.60,sell:6.90,profit:230,pct:50.00,notes:'Earnings play'},
  {date:'2025-05-20',symbol:'QQQ',strike:'445',buy:3.20,sell:4.90,profit:170,pct:53.13,notes:'Momentum'},
  {date:'2025-05-20',symbol:'META',strike:'510',buy:7.80,sell:7.00,profit:-80,pct:-10.26,notes:'Rejection'},
  {date:'2025-05-20',symbol:'TSLA',strike:'170',buy:8.20,sell:2.10,profit:-610,pct:-74.39,notes:'Stop hit'},
  {date:'2025-05-21',symbol:'AMD',strike:'150',buy:5.10,sell:1.20,profit:-390,pct:-76.47,notes:'Weak guidance'},
  {date:'2025-05-21',symbol:'COIN',strike:'240',buy:3.40,sell:2.80,profit:-60,pct:-17.65,notes:'News impact'},
  {date:'2025-05-21',symbol:'INTC',strike:'30',buy:1.20,sell:.80,profit:-40,pct:-33.33,notes:'Weak chart'},
  {date:'2025-05-22',symbol:'PLTR',strike:'20',buy:1.70,sell:1.10,profit:-5.29,pct:-5.29,notes:'Rejection'},
  {date:'2025-05-22',symbol:'RIVN',strike:'12',buy:.90,sell:.50,profit:-40.49,pct:-44.44,notes:'Stop hit'},
  {date:'2025-05-23',symbol:'SNAP',strike:'10',buy:.40,sell:.40,profit:0,pct:0,notes:'Trend down'},
  {date:'2025-05-23',symbol:'SOFI',strike:'7',buy:.30,sell:.30,profit:-7.14,pct:-57.14,notes:'Stop hit'},
  {date:'2025-05-24',symbol:'NFLX',strike:'980',buy:9.20,sell:null,profit:0,pct:0,notes:'مفتوحة'},
  {date:'2025-05-24',symbol:'AMZN',strike:'190',buy:4.10,sell:6.60,profit:250,pct:60.98,notes:'Breakout'},
  {date:'2025-05-24',symbol:'MSFT',strike:'420',buy:5.40,sell:7.90,profit:250,pct:46.30,notes:'Trend'},
  {date:'2025-05-24',symbol:'GOOGL',strike:'175',buy:3.80,sell:5.10,profit:130,pct:34.21,notes:'Continuation'},
];

const $ = id => document.getElementById(id);
const money = n => (n < 0 ? '-$' : '$') + Math.abs(n || 0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const pct = n => `${n >= 0 ? '+' : ''}${(n || 0).toFixed(2)}%`;

function parseDate(v){
  if (!v) return '';
  if (typeof v === 'number' && window.XLSX) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const d = new Date(v);
  if (!isNaN(d)) return d.toISOString().slice(0,10);
  return String(v).trim();
}
function num(v){
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(String(v).replace(/[$,%\s,]/g,''));
  return isNaN(n) ? 0 : n;
}
function getVal(row, aliases){
  const entries = Object.entries(row);
  for (const a of aliases){
    const hit = entries.find(([k]) => String(k).trim().toLowerCase() === a.toLowerCase());
    if (hit) return hit[1];
  }
  return '';
}
function normalizeRows(rows){
  return rows.map(r => {
    const date = parseDate(getVal(r,['date','التاريخ','تاريخ']));
    const symbol = getVal(r,['symbol','ticker','company','اسم الشركة','الشركة','السهم']) || '—';
    const strike = getVal(r,['strike','الاسترايك','سترايك']) || '—';
    const buy = num(getVal(r,['buy','buy price','entry','سعر الشراء','الدخول']));
    let sellRaw = getVal(r,['sell','sell price','exit','سعر البيع','الخروج']);
    const sell = sellRaw === '' ? null : num(sellRaw);
    let profitRaw = getVal(r,['profit','p/l','pnl','الربح','الربح والخسارة','صافي الربح']);
    let profit = profitRaw === '' ? ((sell !== null ? sell - buy : 0) * 100) : num(profitRaw);
    let pctRaw = getVal(r,['pct','percent','percentage','النسبة','النسبة %','profit %']);
    let p = pctRaw === '' ? (buy ? ((sell ?? buy)-buy)/buy*100 : 0) : num(pctRaw);
    const notes = getVal(r,['notes','note','الملاحظات','ملاحظات']) || '';
    return {date,symbol:String(symbol),strike:String(strike),buy,sell,profit,pct:p,notes:String(notes)};
  }).filter(x => x.symbol !== '—' || x.date);
}
function colorize(el, value){
  el.classList.remove('positive','negative','neutral');
  el.classList.add(value > 0 ? 'positive' : value < 0 ? 'negative' : 'neutral');
}

function render(){
  const from = $('fromDate').value;
  const to = $('toDate').value;
  const filtered = trades.filter(t => (!from || t.date >= from) && (!to || t.date <= to));
  const closed = filtered.filter(t => t.sell !== null || t.profit !== 0);
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

  $('totalTrades').textContent = filtered.length;
  $('totalProfit').textContent = money(net);
  colorize($('totalProfit'),net);
  $('returnPct').textContent = pct(returnP);
  colorize($('returnPct'),returnP);
  $('wins').textContent = wins.length;
  $('losses').textContent = losses.length;
  $('winRate').textContent = `${winRate.toFixed(2)}%`;
  colorize($('winRate'),winRate);

  const best = [...closed].sort((a,b)=>b.profit-a.profit)[0];
  $('bestSymbol').textContent = best ? `${best.symbol} ${best.strike}` : '—';
  $('bestProfit').textContent = best ? `+${money(best.profit)}`.replace('+$','$') : '$0.00';
  $('bestPct').textContent = best ? pct(best.pct) : '0%';

  $('avgWin').textContent = money(avgWin); colorize($('avgWin'),avgWin);
  $('avgLoss').textContent = money(avgLoss); colorize($('avgLoss'),avgLoss);
  $('profitFactor').textContent = pf === Infinity ? '∞' : pf.toFixed(2);
  $('expectancy').textContent = money(expectancy); colorize($('expectancy'),expectancy);

  let peak = 0, eq = 0, maxDD = 0;
  const ordered = [...closed].sort((a,b)=>a.date.localeCompare(b.date));
  ordered.forEach(t=>{eq+=t.profit; peak=Math.max(peak,eq); maxDD=Math.min(maxDD,eq-peak);});
  $('maxDrawdown').textContent = money(maxDD); colorize($('maxDrawdown'),maxDD);

  const returns = closed.map(t=>t.pct/100);
  const mean = returns.length ? returns.reduce((a,b)=>a+b,0)/returns.length : 0;
  const variance = returns.length>1 ? returns.reduce((s,x)=>s+(x-mean)**2,0)/(returns.length-1) : 0;
  const sharpe = variance ? mean/Math.sqrt(variance)*Math.sqrt(252) : 0;
  $('sharpe').textContent = sharpe.toFixed(2); colorize($('sharpe'),sharpe);

  $('summaryGross').textContent = money(grossWin); colorize($('summaryGross'),grossWin);
  $('summaryLoss').textContent = money(grossLoss); colorize($('summaryLoss'),grossLoss);
  $('summaryNet').textContent = money(net); colorize($('summaryNet'),net);

  const byDay = {};
  closed.forEach(t=>byDay[t.date]=(byDay[t.date]||0)+t.profit);
  const dayVals = Object.values(byDay);
  const bestDay = dayVals.length ? Math.max(...dayVals) : 0;
  const worstDay = dayVals.length ? Math.min(...dayVals) : 0;
  $('bestDay').textContent = money(bestDay); colorize($('bestDay'),bestDay);
  $('worstDay').textContent = money(worstDay); colorize($('worstDay'),worstDay);

  renderTable(filtered);
  renderCharts(ordered, byDay, wins.length, losses.length);
  $('equityFinal').textContent = money(net);
  $('rowsCount').textContent = `${filtered.length} صفقة`;
}

function renderTable(data){
  $('tradesBody').innerHTML = data.map((t,i)=>`
    <tr>
      <td>${i+1}</td>
      <td dir="ltr">${t.symbol}</td>
      <td>${t.strike}</td>
      <td dir="ltr">${money(t.buy).replace('$','')}</td>
      <td dir="ltr">${t.sell === null ? '—' : money(t.sell).replace('$','')}</td>
      <td class="${t.profit>0?'profit':t.profit<0?'loss':''}">${t.sell===null && t.profit===0 ? '<span style="color:#ffd11a">مفتوحة</span>' : money(t.profit)}</td>
      <td class="${t.pct>0?'profit':t.pct<0?'loss':''}">${t.sell===null && t.profit===0 ? '—' : pct(t.pct)}</td>
      <td>${t.notes || '—'}</td>
    </tr>
  `).join('');
}

function chartDefaults(){
  Chart.defaults.color = '#6f604f';
  Chart.defaults.font.family = 'Cairo';
  Chart.defaults.borderColor = 'rgba(163,137,99,.20)';
}
function renderCharts(ordered, byDay, winCount, lossCount){
  if (!window.Chart) return;
  chartDefaults();
  const eqLabels=[], eqData=[]; let cum=0;
  ordered.forEach((t,i)=>{cum+=t.profit;eqLabels.push(t.date || String(i+1));eqData.push(cum);});
  const days=Object.keys(byDay).sort(), vals=days.map(d=>byDay[d]);

  equityChart?.destroy();
  dailyChart?.destroy();
  winLossChart?.destroy();

  equityChart = new Chart($('equityChart'),{
    type:'line',
    data:{labels:eqLabels,datasets:[{data:eqData,borderColor:'#b88a3b',backgroundColor:'rgba(184,138,59,.14)',fill:true,tension:.28,pointRadius:2,borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{maxRotation:0,autoSkip:true}},y:{ticks:{callback:v=>'$'+v}}}}
  });
  dailyChart = new Chart($('dailyChart'),{
    type:'bar',
    data:{labels:days,datasets:[{data:vals,backgroundColor:vals.map(v=>v>=0?'#4d9664':'#c45b52'),borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>'$'+v}}}}
  });
  winLossChart = new Chart($('winLossChart'),{
    type:'doughnut',
    data:{labels:['أرباح','خسائر'],datasets:[{data:[winCount,lossCount],backgroundColor:['#4d9664','#c45b52'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'58%',plugins:{legend:{position:'right',rtl:true,labels:{boxWidth:13}}}}
  });
}

async function handleFile(file){
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv'){
    const text = await file.text();
    const rows = parseCsv(text);
    trades = normalizeRows(rows);
  } else {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf,{type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws,{defval:''});
    trades = normalizeRows(rows);
  }
  if (!trades.length) {
    alert('لم أجد صفوف صفقات صالحة. تأكد من أسماء الأعمدة ثم جرّب مرة أخرى.');
    return;
  }
  setRangeFromTrades();
  render();
}

function parseCsv(text){
  const lines = text.replace(/^\uFEFF/,'').split(/\r?\n/).filter(Boolean);
  if (!lines.length) return [];
  const split = line => {
    const out=[]; let cur='', q=false;
    for(let i=0;i<line.length;i++){
      const c=line[i];
      if(c === '"'){ if(q && line[i+1]==='"'){cur+='"';i++;} else q=!q; }
      else if(c===',' && !q){out.push(cur);cur='';}
      else cur+=c;
    }
    out.push(cur); return out;
  };
  const headers=split(lines[0]);
  return lines.slice(1).map(line=>{
    const values=split(line), obj={};
    headers.forEach((h,i)=>obj[h]=values[i]??'');
    return obj;
  });
}

function setRangeFromTrades(){
  const dates = trades.map(t=>t.date).filter(Boolean).sort();
  if (dates.length){$('fromDate').value=dates[0];$('toDate').value=dates[dates.length-1];}
}

$('excelFile').addEventListener('change',e=>{ const f=e.target.files[0]; if(f) handleFile(f); });
$('demoBtn').addEventListener('click',()=>{trades=[...demoTrades];setRangeFromTrades();render();});
$('printBtn').addEventListener('click',()=>window.print());
$('fromDate').addEventListener('change',render);
$('toDate').addEventListener('change',render);

const now = new Date();
$('footerDate').textContent = now.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}) + ' - ' + now.toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'});
trades=[...demoTrades];
setRangeFromTrades();
render();
