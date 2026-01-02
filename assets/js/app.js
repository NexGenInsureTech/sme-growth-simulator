/* SME Channel Strategy Simulator & Analyzer — offline local-only */
(function(){
  'use strict';
  const $ = (id) => document.getElementById(id);
  const text = (el, v) => { if (!el) return; el.textContent = v; };
  const fmt = (x) => (Number.isFinite(Number(x||0)) ? Number(x||0) : 0).toLocaleString('en-IN');
  const rupee = (x) => '₹' + fmt(x);
  const toNum = (v) => {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    let s = String(v).trim();
    if (!s) return 0;
    s = s.replace(/₹/g,'').replace(/,/g,'').replace(/\s+/g,'');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
  };
  const normKey = (k) => String(k || '').trim().toUpperCase();
  function normalizeRow(obj){
    const out = {};
    for (const [k,v] of Object.entries(obj || {})) out[normKey(k)] = v;
    return out;
  }

  function findCol(headers, candidates){
    const H = headers.map(h => normKey(h));
    const cand = candidates.map(c => normKey(c));
    for (const c of cand){ const idx = H.indexOf(c); if (idx >= 0) return H[idx]; }
    const strip = (s) => s.replace(/[^A-Z0-9]+/g,'');
    const mapStrip = new Map(H.map(h => [strip(h), h]));
    for (const c of cand){ const s = strip(c); if (mapStrip.has(s)) return mapStrip.get(s); }
    for (const c of cand){
      const tokens = c.split(/\s+/).filter(Boolean);
      for (const h of H){ if (tokens.every(t => h.includes(t))) return h; }
    }
    for (const h of H){ for (const c of cand){ if (h.includes(c) || c.includes(h)) return h; } }
    return null;
  }

  function parseCSVtoAOA(text){
    const rows = [];
    let i = 0, field = '', row = [], inQuotes = false;
    const pushField = () => { row.push(field); field = ''; };
    const pushRow = () => {
      const nonEmpty = row.some(x => String(x ?? '').trim() !== '');
      if (nonEmpty) rows.push(row);
      row = [];
    };
    while (i < text.length){
      const ch = text[i];
      if (inQuotes){
        if (ch === '"'){
          if (text[i+1] === '"'){ field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += ch; i++; continue;
      } else {
        if (ch === '"'){ inQuotes = true; i++; continue; }
        if (ch === ','){ pushField(); i++; continue; }
        if (ch === '\n'){ pushField(); pushRow(); i++; continue; }
        if (ch === '\r'){ if (text[i+1] === '\n') i++; pushField(); pushRow(); i++; continue; }
        field += ch; i++; continue;
      }
    }
    pushField();
    if (row.length) pushRow();
    return rows;
  }

  const REQUIRED_HEADER_PROBES = [
    ['USGI NET PREMIUM','NET PREMIUM','PREMIUM'],
    ['BUSINESS TYPE FRESH RENEWAL','BUSINESS TYPE','FRESH RENEWAL','NEW/RENEW'],
    ['BA NAME','ADVISOR NAME','ADVISOR','BA'],
    ['LINE OF BUSINESS','LOB','PRODUCT LINE'],
    ['INTERMEDIARY CATEGORY','CHANNEL CATEGORY','INTERMEDIARY CAT']
  ];
  function headerScore(cellsUpper){
    let score = 0;
    for (const probe of REQUIRED_HEADER_PROBES){ if (findCol(cellsUpper, probe)) score++; }
    return score;
  }
  function aoaToObjects(aoa){
    if (!aoa || !aoa.length) return { rows: [], headerRowIndex: -1, headers: [] };
    const N = Math.min(30, aoa.length);
    let bestIdx = 0, bestScore = -1;
    for (let r=0;r<N;r++){
      const sc = headerScore((aoa[r]||[]).map(c => normKey(c)));
      if (sc > bestScore){ bestScore = sc; bestIdx = r; }
      if (sc >= REQUIRED_HEADER_PROBES.length) break;
    }
    if (bestScore <= 0) bestIdx = 0;
    const rawHeaders = (aoa[bestIdx]||[]).map(h => normKey(h));
    const headers = []; const seen = new Map();
    for (let i=0;i<rawHeaders.length;i++){
      let h = rawHeaders[i] || ('COL_' + (i+1));
      if (seen.has(h)) { const k = seen.get(h) + 1; seen.set(h,k); h = h + '_' + k; }
      else seen.set(h,1);
      headers.push(h);
    }
    const rows = [];
    for (let r=bestIdx+1;r<aoa.length;r++){
      const row = aoa[r]||[];
      const nonEmpty = row.some(x => String(x ?? '').trim() !== '');
      if (!nonEmpty) continue;
      const obj = {};
      for (let c=0;c<headers.length;c++) obj[headers[c]] = (row[c] ?? '');
      rows.push(obj);
    }
    return { rows, headerRowIndex: bestIdx, headers };
  }

  function quintileThresholds(values){
    const v = values.filter(x => Number.isFinite(x)).slice().sort((a,b)=>a-b);
    const n = v.length;
    if (n < 10) return null;
    const q = (p) => v[Math.floor((n-1)*p)];
    return [q(0.2), q(0.4), q(0.6), q(0.8)];
  }
  function bandLabel(x, thr){
    const [q20,q40,q60,q80] = thr;
    if (x <= q20) return 'Very Low';
    if (x <= q40) return 'Low';
    if (x <= q60) return 'Mid';
    if (x <= q80) return 'High';
    return 'Very High';
  }

  function buildDatasetFromRows(rawRows){
    const rows = rawRows.map(normalizeRow);
    const headers = Array.from(new Set(rows.flatMap(r => Object.keys(r))));

    const colNet = findCol(headers, ['USGI NET PREMIUM', 'NET PREMIUM', 'PREMIUM']);
    const colType = findCol(headers, ['BUSINESS TYPE FRESH RENEWAL', 'BUSINESS TYPE', 'FRESH RENEWAL']);
    const colICat = findCol(headers, ['INTERMEDIARY CATEGORY', 'CHANNEL CATEGORY', 'INTERMEDIARY CAT']);
    const colBA = findCol(headers, ['BA NAME', 'ADVISOR NAME', 'ADVISOR', 'PRIMARY SALES MANAGER NAME']);
    const colLOB = findCol(headers, ['LINE OF BUSINESS', 'LOB', 'PRODUCT LINE']);
    const colProd = findCol(headers, ['PRODUCT NAME', 'PRODUCT']);
    const colPol = findCol(headers, ['POLICY NO', 'POLICY NUMBER', 'USGIPOS POLICY NUMBER']);
    const colInterm = findCol(headers, ['INTERMEDIARY', 'CHANNEL']);
    const colMonth = findCol(headers, ['MONTH', 'BOOKING MONTH', 'ISSUE MONTH']);

    const detected = {colNet, colType, colICat, colBA, colLOB, colProd, colPol, colInterm, colMonth};
    console.log('[SME] Detected columns:', detected);

    const missing = [];
    if (!colNet) missing.push('USGI NET PREMIUM (or NET PREMIUM)');
    if (!colType) missing.push('BUSINESS TYPE FRESH RENEWAL (or BUSINESS TYPE)');
    if (!colBA) missing.push('BA NAME (or ADVISOR NAME)');
    if (!colLOB) missing.push('LINE OF BUSINESS (or LOB)');
    if (!colICat) missing.push('INTERMEDIARY CATEGORY');
    if (missing.length){
      return { error: 'Missing required headers: ' + missing.join(', ') + '.', detected };
    }

    const active = [], all = [];
    for (const r of rows){
      const net = toNum(r[colNet]);
      const typ = String(r[colType] ?? '').toUpperCase();
      const isActive = net > 0 && (typ.includes('NEW') || typ.includes('RENEW'));
      const cleaned = {
        NET: net,
        TYPE: typ,
        ICAT: String(r[colICat] ?? '').trim() || 'UNKNOWN',
        BA: String(r[colBA] ?? '').trim() || 'UNKNOWN',
        LOB: String(r[colLOB] ?? '').trim() || 'UNKNOWN',
        PRODUCT: String((colProd ? r[colProd] : '') ?? '').trim() || 'UNKNOWN',
        POLICY: String((colPol ? r[colPol] : '') ?? '').trim(),
        INTERMEDIARY: String((colInterm ? r[colInterm] : '') ?? '').trim() || 'UNKNOWN',
        MONTH: String((colMonth ? r[colMonth] : '') ?? '').trim()
      };
      all.push(cleaned);
      if (isActive) active.push(cleaned);
    }

    let used = active;
    let warning = null;
    if (!active.length){
      used = all;
      warning = 'No active records (NET PREMIUM > 0 and BUSINESS TYPE contains NEW/RENEW) — using unfiltered dataset for debugging.';
      console.warn('[SME] No active rows; using all rows.');
    }

    const totalNet = used.reduce((s,r)=>s+r.NET,0);
    const policySet = new Set();
    for (const r of used){ if (r.POLICY) policySet.add(r.POLICY); }
    const policyCount = policySet.size ? policySet.size : used.length;
    const avgPrem = policyCount ? totalNet / policyCount : 0;

    const advisors = new Set(used.map(r=>r.BA).filter(Boolean));
    const advisorCount = advisors.size;
    const policiesPerAdvisorAvg = advisorCount ? (policyCount / advisorCount) : 0;

    const kpis = {
      total_usgi_net_premium: totalNet,
      policy_count: policyCount,
      avg_premium_per_policy: avgPrem,
      advisor_count: advisorCount,
      policies_per_advisor_avg: policiesPerAdvisorAvg,
      used_rows: used.length,
      active_rows: active.length,
      total_rows: all.length
    };

    const groupAgg = (keyFn) => {
      const m = new Map();
      for (const r of used){
        const key = keyFn(r);
        if (!m.has(key)) m.set(key, {key, sum:0, count:0});
        const o = m.get(key);
        o.sum += r.NET; o.count += 1;
      }
      const arr = Array.from(m.values()).map(o => ({ key:o.key, sum:o.sum, count:o.count, mean:o.count?o.sum/o.count:0 }));
      arr.sort((a,b)=>b.sum-a.sum);
      return arr;
    };

    const byICat = groupAgg(r=>r.ICAT);
    const byLOB = groupAgg(r=>r.LOB);

    const thr = quintileThresholds(used.map(r=>r.NET));
    let byBand = null;
    if (thr){
      const order = ['Very Low','Low','Mid','High','Very High'];
      const m = new Map(order.map(k => [k,{key:k,sum:0,count:0}]));
      for (const r of used){
        const b = bandLabel(r.NET, thr);
        const o = m.get(b); o.sum += r.NET; o.count += 1;
      }
      byBand = order.map(k => { const o = m.get(k); return {key:k,sum:o.sum,count:o.count,mean:o.count?o.sum/o.count:0};});
    }

    const byBA = groupAgg(r=>r.BA).map(o => ({ name:o.key, policies:o.count, net_premium:o.sum, avg_premium:o.mean }))
      .sort((a,b)=>b.net_premium-a.net_premium).slice(0,20);

    console.log('[SME] Active rows:', active.length, 'Used rows:', used.length);
    console.log('[SME] KPI snapshot:', JSON.parse(JSON.stringify(kpis)));

    return {
      kpis,
      intermediary_category: byICat,
      line_of_business: byLOB,
      premium_banding: byBand,
      top_ba: byBA,
      meta: { detected_columns: detected, warning, generated_at: new Date().toISOString() }
    };
  }

  function setupCanvas(canvas){
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    const w = Math.max(320, Math.floor(rect.width));
    canvas.width = Math.floor(w * dpr);
    canvas.height = Math.floor((canvas.getAttribute('height') ? parseInt(canvas.getAttribute('height'),10) : 260) * dpr);
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr,0,0,dpr,0,0);
    return {ctx, w, h: canvas.height/dpr};
  }
  function drawText(ctx, s, x, y, color='#374151', size=12, align='left'){
    ctx.fillStyle = color;
    ctx.font = `${size}px ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial`;
    ctx.textAlign = align;
    ctx.textBaseline = 'middle';
    ctx.fillText(s, x, y);
  }
  function MicroHBarChart(canvas, labels, values, color='#4f46e5'){
    const {ctx, w, h} = setupCanvas(canvas);
    ctx.clearRect(0,0,w,h);
    const padL=140,padR=16,padT=18,padB=24;
    const innerW=w-padL-padR, innerH=h-padT-padB;
    const n=Math.min(labels.length,12);
    const L=labels.slice(0,n), V=values.slice(0,n);
    const maxV=Math.max(1,...V);
    ctx.strokeStyle='#e5e7eb'; ctx.lineWidth=1;
    for (let t=0;t<=4;t++){
      const x=padL+(innerW*t/4);
      ctx.beginPath(); ctx.moveTo(x,padT); ctx.lineTo(x,padT+innerH); ctx.stroke();
      drawText(ctx, rupee(maxV*t/4), x, h-padB/2, '#6b7280', 10, 'center');
    }
    const rowH=innerH/Math.max(1,n);
    for (let i=0;i<n;i++){
      const y=padT+rowH*i+rowH/2;
      const bw=innerW*(V[i]/maxV);
      drawText(ctx, String(L[i]).slice(0,18), 8, y, '#374151', 12, 'left');
      ctx.fillStyle=color; ctx.fillRect(padL, y-rowH*0.28, Math.max(0,bw), rowH*0.56);
      drawText(ctx, rupee(V[i]), padL+Math.max(0,bw)+6, y, '#111827', 11, 'left');
    }
    ctx.strokeStyle='#9ca3af'; ctx.beginPath(); ctx.moveTo(padL,padT); ctx.lineTo(padL,padT+innerH); ctx.stroke();
  }
  function MicroLineChart(canvas, labels, values, targetValue){
    const {ctx, w, h} = setupCanvas(canvas);
    ctx.clearRect(0,0,w,h);
    const padL=52,padR=16,padT=18,padB=30;
    const innerW=w-padL-padR, innerH=h-padT-padB;
    const maxV=Math.max(1,...values,(targetValue||0));
    ctx.strokeStyle='#e5e7eb'; ctx.lineWidth=1;
    for (let t=0;t<=4;t++){
      const y=padT+innerH*(1-t/4);
      ctx.beginPath(); ctx.moveTo(padL,y); ctx.lineTo(padL+innerW,y); ctx.stroke();
      drawText(ctx, rupee(maxV*t/4), padL-6, y, '#6b7280', 10, 'right');
    }
    const n=values.length;
    for (let i=0;i<n;i++){
      const x=padL+innerW*(n===1?0:i/(n-1));
      drawText(ctx, labels[i], x, h-padB/2, '#6b7280', 10, 'center');
    }
    if (targetValue !== null && targetValue !== undefined){
      const ty=padT+innerH*(1-(targetValue/maxV));
      ctx.save(); ctx.setLineDash([6,4]); ctx.strokeStyle='#dc2626';
      ctx.beginPath(); ctx.moveTo(padL,ty); ctx.lineTo(padL+innerW,ty); ctx.stroke();
      ctx.restore(); drawText(ctx,'Target',padL+innerW-4,ty-10,'#dc2626',10,'right');
    }
    ctx.strokeStyle='#4f46e5'; ctx.lineWidth=2; ctx.beginPath();
    for (let i=0;i<n;i++){
      const x=padL+innerW*(n===1?0:i/(n-1));
      const y=padT+innerH*(1-(values[i]/maxV));
      if (i===0) ctx.moveTo(x,y); else ctx.lineTo(x,y);
    }
    ctx.stroke();
    ctx.fillStyle='#4f46e5';
    for (let i=0;i<n;i++){
      const x=padL+innerW*(n===1?0:i/(n-1));
      const y=padT+innerH*(1-(values[i]/maxV));
      ctx.beginPath(); ctx.arc(x,y,3,0,Math.PI*2); ctx.fill();
    }
    drawText(ctx,'Projected Net Premium',padL,10,'#111827',12,'left');
  }

  let SME_DATA = null;
  let BASE = null;

  function computeBaseline(){
    const k = SME_DATA?.kpis || { advisor_count: 10, policies_per_advisor_avg: 24, avg_premium_per_policy: 18000 };
    const advisors = Math.max(1, Math.round(k.advisor_count || 10));
    const ppaMonthly = Math.max(0, Math.round((k.policies_per_advisor_avg || 0)/12));
    const avgPrem = Math.max(100, Math.round(k.avg_premium_per_policy || 18000));
    return { years:3, advisors, ppa:ppaMonthly, avgPrem, employees:50, hirings:2, target:20 };
  }

  function setSliderDefaults(){
    if (!BASE) return;
    $('years').value = BASE.years;
    const maxAdv = Math.max(50, Math.round(BASE.advisors*5));
    $('advisors').min=1; $('advisors').max=String(maxAdv); $('advisors').value=BASE.advisors;
    $('ppa').min=0; $('ppa').max=String(Math.max(50, BASE.ppa*5 + 10)); $('ppa').value=BASE.ppa;
    $('avgPrem').min=100; $('avgPrem').max=String(Math.max(200000, BASE.avgPrem*3)); $('avgPrem').step=100; $('avgPrem').value=BASE.avgPrem;
    $('employees').min=1; $('employees').max=String(Math.max(1000, BASE.employees*5)); $('employees').value=BASE.employees;
    $('hirings').min=0; $('hirings').max=String(Math.max(50, BASE.hirings*10+10)); $('hirings').value=BASE.hirings;
    $('target').value=BASE.target;
    updateSliderLabels();
  }
  function readSliders(){
    return {
      years: parseInt($('years').value,10),
      advisors: parseInt($('advisors').value,10),
      ppa: parseInt($('ppa').value,10),
      avgPrem: parseInt($('avgPrem').value,10),
      employees: parseInt($('employees').value,10),
      hirings: parseInt($('hirings').value,10),
      target: parseInt($('target').value,10)
    };
  }
  function updateSliderLabels(){
    const s = readSliders();
    text($('yearsVal'), String(s.years));
    text($('advisorsVal'), String(s.advisors));
    text($('ppaVal'), String(s.ppa));
    text($('avgPremVal'), rupee(s.avgPrem));
    text($('employeesVal'), String(s.employees));
    text($('hiringsVal'), String(s.hirings));
    text($('targetVal'), String(s.target) + '%');
  }
  function baselineAnnualPremium(base){ return base.advisors*base.ppa*12*base.avgPrem; }

  function runSim(){
    if (!BASE) return;
    const s = readSliders(); updateSliderLabels();
    const capacityAdvisors = (s.employees + (s.hirings*12*s.years)/s.years) * 2.0;
    const effectiveMax = Math.min(s.advisors, Math.max(0, Math.round(capacityAdvisors)));
    let currentAdvisors = Math.max(0, Math.min(effectiveMax, BASE.advisors));
    const labels=[], values=[];
    for (let y=1;y<=s.years;y++){
      const uplift = 1 + Math.min(0.05*(y-1), 0.25);
      values.push(currentAdvisors*s.ppa*12*s.avgPrem*uplift);
      labels.push('Y'+y);
      currentAdvisors = Math.min(effectiveMax, currentAdvisors + Math.round(s.hirings*6));
    }
    const targetAnnual = baselineAnnualPremium(BASE) * (1 + s.target/100);
    MicroLineChart($('simChart'), labels, values, targetAnnual);
  }

  function setKPIs(){
    const k = SME_DATA?.kpis;
    if (!k){ text($('kpiNet'),'₹0'); text($('kpiPolicies'),'0'); text($('kpiAvg'),'₹0'); text($('kpiAdvisors'),'0'); return; }
    text($('kpiNet'), rupee(k.total_usgi_net_premium));
    text($('kpiPolicies'), fmt(k.policy_count));
    text($('kpiAvg'), rupee(k.avg_premium_per_policy));
    text($('kpiAdvisors'), fmt(k.advisor_count));
  }
  function renderPortfolio(){
    const icat = SME_DATA?.intermediary_category || [];
    const lob = SME_DATA?.line_of_business || [];
    const band = SME_DATA?.premium_banding;
    MicroHBarChart($('icatChart'), icat.map(d=>d.key), icat.map(d=>d.sum), '#4f46e5');
    MicroHBarChart($('lobChart'), lob.map(d=>d.key), lob.map(d=>d.sum), '#059669');
    if (band) MicroHBarChart($('bandChart'), band.map(d=>d.key), band.map(d=>d.sum), '#d97706');
    else {
      const {ctx,w,h} = setupCanvas($('bandChart')); ctx.clearRect(0,0,w,h);
      drawText(ctx, 'Premium banding skipped (need ≥10 rows).', 10, h/2, '#6b7280', 12, 'left');
    }
    const tbody = $('baTable'); tbody.innerHTML='';
    for (const r of (SME_DATA?.top_ba || [])){
      const tr = document.createElement('tr'); tr.className='border-b';
      const td1=document.createElement('td'); td1.className='py-1'; td1.textContent=r.name;
      const td2=document.createElement('td'); td2.className='py-1'; td2.style.textAlign='right'; td2.textContent=fmt(r.policies);
      const td3=document.createElement('td'); td3.className='py-1'; td3.style.textAlign='right'; td3.textContent=rupee(r.net_premium);
      const td4=document.createElement('td'); td4.className='py-1'; td4.style.textAlign='right'; td4.textContent=rupee(r.avg_premium);
      tr.append(td1,td2,td3,td4); tbody.appendChild(tr);
    }
  }

  function logStatus(msg){ text($('uploadStatus'), msg); console.log('[SME] Status:', msg); }
  function showWarning(msg){
    const b=$('warningBanner'), t=$('warningText');
    if (!msg){ b.classList.add('hidden'); text(t,''); return; }
    b.classList.remove('hidden'); text(t,msg);
  }

  function exportJSON(){
    if (!SME_DATA) return;
    const blob = new Blob([JSON.stringify(SME_DATA, null, 2)], {type:'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'SME_DATA_export.json';
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); },0);
  }

  async function handleFile(file){
    if (!file) return;
    showWarning(null);
    const name=file.name||'';
    const ext=name.split('.').pop().toLowerCase();
    logStatus('Reading: '+name+' ...');

    try {
      const isExcel = (ext==='xlsx' || ext==='xlsm');
      const hasXLSX = (window.XLSX && typeof window.XLSX.read === 'function');

      // IMPORTANT: don't try to parse Excel as CSV (it looks like "no change").
      if (isExcel && !hasXLSX){
        showWarning('Excel library missing — cannot read .xlsx/.xlsm in CSV-only mode. Replace assets/js/xlsx.full.min.js with the real SheetJS full build, or export your PR sheet to CSV.');
        logStatus('Excel upload blocked (CSV-only mode).');
        console.warn('[SME] XLSX missing. window.XLSX=', window.XLSX);
        return;
      }

      if (isExcel && hasXLSX){
        const ab = await file.arrayBuffer();
        const wb = window.XLSX.read(ab, {type:'array'});
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const aoa = window.XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
        const {rows, headerRowIndex} = aoaToObjects(aoa);
        console.log('[SME] Excel header row detected at (0-based):', headerRowIndex);
        logStatus('Loaded: '+sheetName+' · Records: '+fmt(rows.length)+' · Header row: '+(headerRowIndex+1));
        postParse(rows);
        return;
      }

      // CSV
      const txt = await file.text();
      const aoa = parseCSVtoAOA(txt);
      const {rows, headerRowIndex} = aoaToObjects(aoa);
      console.log('[SME] CSV header row detected at (0-based):', headerRowIndex);
      logStatus('Loaded: CSV · Records: '+fmt(rows.length)+' · Header row: '+(headerRowIndex+1));
      postParse(rows);

    } catch (e){
      console.error(e);
      logStatus('Error reading file: ' + (e?.message || e));
      showWarning('Upload failed. Check file format and headers. Details in console.');
    }
  }

  function postParse(rows){
    const built = buildDatasetFromRows(rows);
    if (built.error){ showWarning(built.error); logStatus('Header error — see banner.'); return; }
    SME_DATA = built;
    if (SME_DATA.meta?.warning) showWarning(SME_DATA.meta.warning); else showWarning(null);
    BASE = computeBaseline();
    setSliderDefaults();
    setKPIs();
    renderPortfolio();
    runSim();
  }

  function bootSample(){
    const sampleRows=[];
    const icats=['POSP','BROKER','CORPORATE AGENT','DIRECT'];
    const lobs=['MOTOR','HEALTH','FIRE','MARINE'];
    const types=['NEW BUSINESS','RENEWAL'];
    for (let i=0;i<60;i++){
      sampleRows.push({
        'USGI NET PREMIUM': 5000+(i%10)*2500+(lobs[i%4]==='FIRE'?7000:0),
        'BUSINESS TYPE FRESH RENEWAL': types[i%2],
        'INTERMEDIARY CATEGORY': icats[i%4],
        'BA NAME': 'Advisor '+(1+(i%12)),
        'LINE OF BUSINESS': lobs[i%4],
        'PRODUCT NAME': lobs[i%4]+' SME',
        'POLICY NO': 'P'+(10000+i),
        'MONTH': '2025-'+String(1+(i%12)).padStart(2,'0')
      });
    }
    postParse(sampleRows);
    logStatus('Loaded: Sample (offline) · Records: '+fmt(sampleRows.length));
  }

  function excelCapabilityFlag(){
    const enabled = (window.XLSX && typeof window.XLSX.read === 'function');
    text($('excelFlag'), enabled ? 'Enabled' : 'CSV only; add xlsx.full.min.js for Excel');
    if (!enabled) console.warn('[SME] Excel disabled — XLSX library not detected.');
  }

  function wire(){
    $('fileInput').addEventListener('change', (e)=>handleFile(e.target.files[0]));
    for (const id of ['years','advisors','ppa','avgPrem','employees','hirings','target']){
      $(id).addEventListener('input', ()=>runSim());
    }
    $('exportBtn').addEventListener('click', exportJSON);
    $('resetBtn').addEventListener('click', ()=>{ BASE = computeBaseline(); setSliderDefaults(); runSim(); });
    window.addEventListener('resize', ()=>{ setKPIs(); renderPortfolio(); runSim(); });
  }

  function init(){
    excelCapabilityFlag();
    wire();
    bootSample();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
