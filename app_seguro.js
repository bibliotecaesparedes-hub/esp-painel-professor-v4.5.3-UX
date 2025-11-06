/* ESP.EE v4.5.3-UX — app.js unificado com correções e afinações
   - Robustez (offline-first)
   - Gate Admin por config/claims (fallback legado)
   - Sugestão Nº da lição (aluno+disciplina+professor)
   - Atrasos ilimitados com badges
   - Export por disciplina (professor autenticado)
   - Admin: CRUD + validações + paginação + ordenação + pesquisa + export
   - Máscara HH:MM (Nova oficina + CRUD Oficinas)
*/
const SITE_ID       = 'esparedes-my.sharepoint.com,540a0485-2578-481e-b4d8-220b41fb5c43,7335dc42-69c8-42d6-8282-151e3783162d';
const CFG_PATH      = '/Documents/GestaoAlunos-OneDrive/config_especial.json';
const REG_PATH      = '/Documents/GestaoAlunos-OneDrive/2registos_alunos.json';
const BACKUP_FOLDER = '/Documents/GestaoAlunos-OneDrive/backup';
const LEGACY_ADMINS = ['biblioteca@esparedes.pt'];

const MSAL_CONFIG = {
  auth: {
    clientId: 'c5573063-8a04-40d3-92bf-eb229ad4701c',
    authority: 'https://login.microsoftonline.com/d650692c-6e73-48b3-af84-e3497ff3e1f1',
    redirectUri: 'https://bibliotecaesparedes-hub.github.io/esp-painel-professor-v4.5.3-UX/'
  },
  cache: { cacheLocation:'localStorage', storeAuthStateInCookie:false }
};
const MSAL_SCOPES = { scopes: ['Files.ReadWrite.All','User.Read','openid','profile','offline_access'] };

let msalApp, account, accessToken;
const state = { config: null, reg: { versao:'v2', registos:[] } };
const $ = s => document.querySelector(s);

function updateSync(t){ const el=$('#syncIndicator'); if(el) el.textContent=t; }
function toast(t){ try{ Swal.fire({toast:true,position:'top-end',timer:1500,showConfirmButton:false,title:t}); }catch{} }
function checkDeps(){
  const missing = [];
  if (typeof Swal === 'undefined') missing.push('SweetAlert2');
  if (typeof XLSX === 'undefined') missing.push('SheetJS (XLSX)');
  if (!(window.jspdf && window.jspdf.jsPDF)) missing.push('jsPDF');
  if (missing.length){
    const msg = 'Dependências em falta: ' + missing.join(', ') + '. Carregue as bibliotecas no index.html (CDN) antes de app.js.';
    console.warn(msg);
    try { Swal.fire('Erro', msg, 'error'); }catch{}
  }
}
function getAccountEmail(){ return (account?.username || '').trim().toLowerCase(); }
function isAdmin(){
  const email = getAccountEmail();
  const cfgAdmins = (state.config?.professores || [])
    .filter(p => (p.role||'').toLowerCase()==='admin')
    .map(p => (p.email||'').trim().toLowerCase());
  const acc    = msalApp?.getActiveAccount?.();
  const roles  = (acc?.idTokenClaims?.roles)  || [];
  const groups = (acc?.idTokenClaims?.groups) || [];
  const hasClaim = roles.includes('esp-admin') || groups.includes('esp-admin');
  return cfgAdmins.includes(email) || hasClaim || LEGACY_ADMINS.includes(email);
}
function setSessionName(){ const el=$('#sessNome'); if(!el) return; el.textContent = account ? `Sessão: ${account.name || account.username}` : 'Sessão: não iniciada'; }
function applyRoleVisibilityHard(){ const admin = isAdmin(); $('#admin')?.classList.toggle('hidden', !admin); $('#secHoje')?.classList.toggle('hidden', admin); $('#secRegistos')?.classList.toggle('hidden', admin); }
function updateAuthButtons(){ const logged = !!account; $('#btnMsLogin')?.classList.toggle('hidden', logged); $('#btnMsLogout')?.classList.toggle('hidden', !logged); setSessionName(); }

/* ====== MSAL ====== */
async function initMsal(){
  if (typeof msal === 'undefined'){ console.error('MSAL missing'); return; }
  msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  try{
    const resp = await msalApp.handleRedirectPromise();
    if (resp && resp.account){ account = resp.account; msalApp.setActiveAccount(account); await acquireToken(); }
    const accs = msalApp.getAllAccounts();
    if (accs.length && !account){ account = accs[0]; msalApp.setActiveAccount(account); await acquireToken(); }
  }catch(e){ console.warn('msal init', e); }
  setSessionName(); updateAuthButtons(); applyRoleVisibilityHard();
}
async function acquireToken(){ if(!msalApp) return; try{ const r = await msalApp.acquireTokenSilent(MSAL_SCOPES); accessToken=r.accessToken; return accessToken; }catch(e){ try{ await msalApp.acquireTokenRedirect(MSAL_SCOPES);}catch(err){ console.error(err);} } }
function ensureLogin(){ if(msalApp) msalApp.loginRedirect(MSAL_SCOPES); }
function ensureLogout(){ if(msalApp) msalApp.logoutRedirect(); }

/* ====== Graph ====== */
async function graphLoad(path){
  if(!accessToken) await acquireToken();
  try{
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`;
    const r = await fetch(url, { headers:{ Authorization:`Bearer ${accessToken}` } });
    if(r.ok){ const txt = await r.text(); return txt ? JSON.parse(txt) : null; }
    if(r.status===404) return null;
    throw new Error('Graph '+r.status);
  }catch(e){ console.warn('graphLoad', e); return null; }
}
async function graphSave(path,obj){
  if(!accessToken) await acquireToken();
  try{
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`;
    const r = await fetch(url, { method:'PUT', headers:{ Authorization:`Bearer ${accessToken}` }, body: JSON.stringify(obj, null, 2) });
    if(!r.ok) throw new Error('save '+r.status);
    return await r.json();
  }catch(e){ console.warn('graphSave', e); throw e; }
}
async function graphList(folderPath){
  if(!accessToken) await acquireToken();
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${folderPath}:/children`;
  try{
    const r = await fetch(url, { headers:{ Authorization:`Bearer ${accessToken}` } });
    if(!r.ok) throw new Error('list '+r.status);
    const data = await r.json();
    return Array.isArray(data.value) ? data.value : [];
  }catch(e){ console.warn('graphList', e); return []; }
}

/* ====== Onboarding/Migração mínima ====== */
function isRegData(o){ return o && typeof o==='object' && (o.versao||o.version) && Array.isArray(o.registos); }
function isCfg(o){ return o && typeof o==='object' && Array.isArray(o.professores); }
async function onboardingIfNeeded(){ return true; }

/* ====== Carregamento ====== */
async function loadConfigAndReg(){
  updateSync('🔁 sincronizando...');
  let cfg = await graphLoad(CFG_PATH);
  let reg = await graphLoad(REG_PATH);

  // Auto-migração caso a config antiga contenha registos
  if (isRegData(cfg) && (!reg || !Array.isArray(reg.registos) || reg.registos.length===0)){
    try{
      await graphSave(REG_PATH, cfg);
      reg = cfg;
      cfg = {professores:[], alunos:[], disciplinas:[], oficinas:[], calendario:{}};
      await graphSave(CFG_PATH, cfg);
      toast('Config/Registos migrados automaticamente');
    }catch(e){ console.warn('auto-migração', e); }
  }

  state.config = isCfg(cfg) ? cfg : (JSON.parse(localStorage.getItem('esp_config')||'{}')||{});
  if(!isCfg(state.config)) state.config = {professores:[], alunos:[], disciplinas:[], oficinas:[], calendario:{}};

  state.reg = isRegData(reg) ? reg : (JSON.parse(localStorage.getItem('esp_reg')||'{}')||{versao:'v2', registos:[]});
  if(!isRegData(state.reg)) state.reg = {versao:'v2', registos:[]};

  localStorage.setItem('esp_config', JSON.stringify(state.config));
  localStorage.setItem('esp_reg',    JSON.stringify(state.reg));

  await onboardingIfNeeded();
  updateSync('💾 guardado');
  applyRoleVisibilityHard(); updateAuthButtons();
  renderHoje(); renderRegList();
}

/* ====== Hoje ====== */
function diaSemana(dateStr){ const d=new Date(dateStr); const g=d.getDay(); return g===0?7:g; }
function getOficinasHoje(profId,dateStr){ const dw=diaSemana(dateStr); return (state.config.oficinas||[]).filter(s=> s.professorId===profId && Number(s.diaSemana)===Number(dw)); }
function nextNumeroLicao(alunoId,disciplinaId,professorId){
  const nums=(state.reg.registos||[]).filter(r=> r.alunoId===alunoId && r.disciplinaId===disciplinaId && r.professorId===professorId)
    .map(r=> parseInt(r.numeroLicao,10)).filter(n=>!isNaN(n));
  return (nums.length? Math.max(...nums)+1 : 1).toString();
}
function renderHoje(){
  const date=$('#dataHoje')?.value || new Date().toISOString().slice(0,10);
  if($('#dataHoje')) $('#dataHoje').value=date;
  const out=$('#sessoesHoje'); if(!out) return; out.innerHTML='';

  if(isAdmin()){ out.innerHTML='<div class="muted">Perfil admin — use Administração.</div>'; return; }

  const email=getAccountEmail();
  const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email);
  if(!prof){ out.innerHTML='<div class="muted">Professor não reconhecido.</div>'; return; }

  const oficinas=getOficinasHoje(prof.id,date);
  if(!oficinas.length){ out.innerHTML='<div class="muted">Sem oficinas para hoje.</div>'; return; }

  const alunosById=Object.fromEntries((state.config.alunos||[]).map(a=>[String(a.id),a]));
  const discById  =Object.fromEntries((state.config.disciplinas||[]).map(d=>[String(d.id),d]));

  oficinas.forEach(sess=>{
    const disc=discById[sess.disciplinaId] || {nome:sess.disciplinaId};
    const alunos=(sess.alunoIds||[]).map(id=>alunosById[id]).filter(Boolean);
    const card=document.createElement('div'); card.className='card';
    card.innerHTML=`
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap">
        <div><strong>${disc.nome}</strong> <span class="muted">• Sala ${sess.sala||'-'}</span></div>
        <div class="muted">${sess.horaInicio||''} – ${sess.horaFim||''}</div>
      </div>
      <div style="margin-top:10px">
        ${alunos.map(a=>`
          <div style="display:grid;grid-template-columns:120px 1fr 120px 200px;gap:6px;align-items:center;margin:6px 0">
            <div><strong>${a?.numero||''}</strong> ${a?.nome||a?.id||''}</div>
            <input class="input sumario" data-aluno="${a.id}" placeholder="Sumário (por aluno)">
            <input class="input nlec" data-aluno="${a.id}" placeholder="Nº lição">
            <select class="input status" data-aluno="${a.id}">
              <option value="P">Presente</option>
              <option value="A">Ausente (injust.)</option>
              <option value="J">J (just.)</option>
            </select>
          </div>`).join('')}
      </div>
      <div class="controls-row"><button class="btn" data-saveSess>Guardar registos desta oficina</button></div>
    `;
    out.appendChild(card);

    // Nº da lição sugerido
    card.querySelectorAll('.nlec').forEach(inp=>{
      const aid = inp.dataset.aluno;
      if(aid && !inp.value){ inp.value = nextNumeroLicao(aid, sess.disciplinaId, prof.id); }
    });

    card.querySelector('[data-saveSess]')?.addEventListener('click', async ()=>{
      const inputsSum=[...card.querySelectorAll('.sumario')];
      const inputsNum=[...card.querySelectorAll('.nlec')];
      const inputsSts=[...card.querySelectorAll('.status')];

      const mapSum=Object.fromEntries(inputsSum.map(i=>[i.dataset.aluno,i.value.trim()]));
      const mapNum=Object.fromEntries(inputsNum.map(i=>[i.dataset.aluno,i.value.trim()]));
      const mapSts=Object.fromEntries(inputsSts.map(i=>[i.dataset.aluno,i.value]));

      const batch=(sess.alunoIds||[]).map(aid=>({
        id:'R'+crypto.randomUUID()+aid, data:date, professorId:prof.id, disciplinaId:sess.disciplinaId,
        alunoId:aid, sessaoId:sess.id||'',
        numeroLicao:(mapNum[aid]&&mapNum[aid].trim())?mapNum[aid].trim():nextNumeroLicao(aid, sess.disciplinaId, prof.id),
        sumario:mapSum[aid]||'', status:mapSts[aid]||'P', justificacao:'', criadoEm:new Date().toISOString()
      }));
      state.reg.registos.push(...batch);
      await persistReg();
      toast(`Guardado: ${batch.length} registos`);
      renderRegList();
    });
  });
}

/* ====== Registos + Atrasos ====== */
function expectedSessDates(sess, startISO, endISO){
  const res=[]; const start=new Date(startISO); const end=new Date(endISO);
  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){
    const ds=d.toISOString().slice(0,10);
    const dw=diaSemana(ds);
    if(Number(dw)===Number(sess.diaSemana)) res.push(ds);
  }
  return res;
}
function inicioAnoLetivoISO(){
  const s = state.config?.calendario?.anoLetivoInicio;
  if(typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const now=new Date(); const sept1=new Date(now.getFullYear(),8,1);
  return sept1.toISOString().slice(0,10);
}
function getAtrasos(profId){
  const today=new Date().toISOString().slice(0,10);
  const iniUI=$('#fltIni')?.value; const fimUI=$('#fltFim')?.value;
  const startISO=iniUI || inicioAnoLetivoISO(); const endISO=fimUI || today;

  const regKey=new Map();
  (state.reg.registos||[]).forEach(r=>{
    const k=`${r.data}|${r.professorId}|${r.disciplinaId}|${r.alunoId}|${r.sessaoId||''}`;
    regKey.set(k,r);
  });

  const atrasos=[];
  (state.config.oficinas||[]).filter(s=> s.professorId===profId).forEach(sess=>{
    const days=expectedSessDates(sess,startISO,endISO);
    (sess.alunoIds||[]).forEach(aid=>{
      days.forEach(ds=>{
        const key=`${ds}|${sess.professorId}|${sess.disciplinaId}|${aid}|${sess.id||''}`;
        const r=regKey.get(key);
        const miss={inexistente:false,numero:false,sumario:false,status:false};
        if(!r){ miss.inexistente=true; }
        else{ if(!r.numeroLicao) miss.numero=true; if(!r.sumario) miss.sumario=true; if(!r.status) miss.status=true; }
        if(miss.inexistente||miss.numero||miss.sumario||miss.status){
          atrasos.push({data:ds,sessaoId:sess.id,alunoId:aid,disciplinaId:sess.disciplinaId,missing:miss});
        }
      });
    });
  });
  return atrasos.sort((a,b)=> a.data<b.data? -1:1);
}
function renderRegList(){
  const el=$('#regList'); if(!el) return; el.innerHTML='';

  // Em atraso (professor)
  if(!isAdmin()){
    const email=getAccountEmail();
    const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email);
    if(prof){
      const atrasos=getAtrasos(prof.id);
      if(atrasos.length){
        const wrap=document.createElement('div'); wrap.className='card';
        wrap.innerHTML = `<h4>Registos em atraso (${atrasos.length})</h4>` + atrasos.map(a=>{
          const labels=[];
          if(a.missing?.inexistente) labels.push('<span class="badge badge-miss">Sem registo</span>');
          if(a.missing?.numero)      labels.push('<span class="badge badge-miss">Falta Nº</span>');
          if(a.missing?.sumario)     labels.push('<span class="badge badge-miss">Falta sumário</span>');
          if(a.missing?.status)      labels.push('<span class="badge badge-miss">Falta presença</span>');
          const badges=labels.join(' ');
          return `<div style="padding:6px;border-bottom:1px solid var(--primary-border)">
            ${a.data} · ${a.disciplinaId} · aluno ${a.alunoId}
            ${badges}
            <button class="btn" data-completar="${a.data}|${a.disciplinaId}|${a.alunoId}|${a.sessaoId||''}">Completar</button>
          </div>`;
        }).join('');
        el.appendChild(wrap);
        wrap.querySelectorAll('[data-completar]').forEach(b =>
          b.addEventListener('click', ()=> openCompletarModal(b.dataset.completar)));
      }
    }
  }

  // Lista de registos (com filtro por datas)
  const ini=$('#fltIni')?.value, fim=$('#fltFim')?.value;
  (state.reg.registos||[]).filter(r=>{
    if(!ini && !fim) return true;
    const d=r.data;
    if(ini && d<ini) return false;
    if(fim && d>fim) return false;
    return true;
  }).slice().reverse().forEach(r=>{
    const div=document.createElement('div'); div.className='card';
    const status = r.status==='P'?'Presente' : (r.status==='A'?'Ausente (injust.)' : (r.status==='J'?'J (just.)' : (r.presenca===true?'Presente':r.presenca===false?'Ausente':'-')));
    div.textContent = `${r.data} • ${r.disciplinaId} • aluno ${r.alunoId||'-'} • Nº ${r.numeroLicao||'-'} • ${r.sumario||'-'} • ${status}`;
    el.appendChild(div);
  });
}
async function openCompletarModal(key){
  const [data,disc,alunoId,sessId]=key.split('|');
  const { value: form } = await Swal.fire({
    title:`Completar registo ${data}`,
    html:`<input id="nlec" class="swal2-input" placeholder="Nº lição">
          <input id="sum" class="swal2-input" placeholder="Sumário">
          <select id="sts" class="swal2-input">
            <option value="P">Presente</option>
            <option value="A">Ausente (injust.)</option>
            <option value="J">J (just.)</option>
          </select>
          <input id="just" class="swal2-input" placeholder="Justificação (se J)">`,
    confirmButtonText:'Guardar', showCancelButton:true,
    preConfirm:()=>({ n:$('#nlec').value.trim(), s:$('#sum').value.trim(), st:$('#sts').value, j:$('#just').value.trim() })
  });
  if(!form) return;
  const email=getAccountEmail();
  const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email);
  state.reg.registos.push({
    id:'R'+crypto.randomUUID()+alunoId, data, professorId:prof?.id, disciplinaId:disc,
    alunoId, sessaoId:sessId||'', numeroLicao:form.n, sumario:form.s,
    status:form.st, justificacao:form.j, criadoEm:new Date().toISOString()
  });
  await persistReg();
  renderRegList();
}

/* ====== Persistência ====== */
async function persistReg(){
  try{
    updateSync('🔁 sincronizando...');
    await graphSave(REG_PATH,state.reg);
    localStorage.setItem('esp_reg',JSON.stringify(state.reg));
    updateSync('💾 guardado');
  }catch(e){
    console.warn('save failed',e);
    localStorage.setItem('esp_reg',JSON.stringify(state.reg));
    updateSync('⚠ offline');
    Swal.fire('Aviso','Guardado localmente. Será sincronizado quando online.','warning');
  }
}

/* ====== Exportações (PDF/XLSX) ====== */
function semanaRange(){ const hoje=new Date(); const ini=new Date(hoje); ini.setDate(hoje.getDate()-hoje.getDay()+1); const fim=new Date(ini); fim.setDate(ini.getDate()+6); return [ini.toISOString().slice(0,10), fim.toISOString().slice(0,10)]; }
async function exportSemanalPDF(){ if(!(window.jspdf&&window.jspdf.jsPDF)){ Swal.fire('Erro','jsPDF não disponível','error'); return; }
  const email=getAccountEmail(); const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email); if(!prof) return;
  const [sISO,eISO]=semanaRange(); const rows=(state.reg.registos||[])
    .filter(r=> r.professorId===prof.id && r.data>=sISO && r.data<=eISO)
    .map(r=> [r.data, r.alunoId, r.disciplinaId, r.numeroLicao||'', r.sumario||'', (r.status==='P'?'Presente':(r.status==='A'?'Ausente (injust.)':(r.status==='J'?'J (just.)':'')))]);
  const doc=new window.jspdf.jsPDF({unit:'pt',format:'a4'}); doc.text(`Registos semanais • ${sISO} a ${eISO}`,40,40);
  doc.autoTable({startY:60, head:[['Data','Aluno','Oficina','Nº','Sumário','Presença']], body:rows, styles:{fontSize:9}});
  doc.save(`registos_${sISO}_${eISO}.pdf`);
}
async function exportSemanalXLSX(){ const email=getAccountEmail(); const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email); if(!prof){ Swal.fire('Erro','Sem professor','error'); return; }
  const [sISO,eISO]=semanaRange(); const rows=(state.reg.registos||[])
    .filter(r=> r.professorId===prof.id && r.data>=sISO && r.data<=eISO)
    .map(r=>({ Data:r.data, Aluno:r.alunoId, Oficina:r.disciplinaId, Numero:r.numeroLicao||'', Sumario:r.sumario||'', Presenca:(r.status==='P'?'Presente':(r.status==='A'?'Ausente (injust.)':(r.status==='J'?'J (just.)':''))) }));
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows),'Semana');
  const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  a.download=`registos_${sISO}_${eISO}.xlsx`; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1200);
}
async function exportAlunoPDF(){ if(!(window.jspdf&&window.jspdf.jsPDF)){ Swal.fire('Erro','jsPDF não disponível','error'); return; }
  const { value: form } = await Swal.fire({ title:'Exportar por aluno (PDF)', html:`<input id="al" class="swal2-input" placeholder="ID do aluno"><input id="di" class="swal2-input" type="date"><input id="df" class="swal2-input" type="date">`, confirmButtonText:'Exportar', showCancelButton:true, preConfirm:()=>({ a:$('#al').value.trim(), i:$('#di').value, f:$('#df').value }) });
  if(!form || !form.a) return;
  const rows=(state.reg.registos||[]).filter(r=> r.alunoId===form.a && (!form.i || r.data>=form.i) && (!form.f || r.data<=form.f))
    .map(r=> [r.data, r.disciplinaId, r.numeroLicao||'', r.sumario||'', (r.status==='P'?'Presente':(r.status==='A'?'Ausente (injust.)':(r.status==='J'?'J (just.)':'')))]);
  const doc=new window.jspdf.jsPDF({unit:'pt',format:'a4'}); doc.text(`Aluno ${form.a} • ${form.i||'…'} a ${form.f||'…'}`,40,40);
  doc.autoTable({startY:60, head:[['Data','Oficina','Nº','Sumário','Presença']], body:rows, styles:{fontSize:9}});
  doc.save(`aluno_${form.a}_${form.i||'ini'}_${form.f||'fim'}.pdf`);
}
async function exportAlunoXLSX(){ const { value: form } = await Swal.fire({ title:'Exportar por aluno (XLSX)', html:`<input id="alx" class="swal2-input" placeholder="ID do aluno"><input id="dix" class="swal2-input" type="date"><input id="dfx" class="swal2-input" type="date">`, confirmButtonText:'Exportar', showCancelButton:true, preConfirm:()=>({ a:$('#alx').value.trim(), i:$('#dix').value, f:$('#dfx').value }) });
  if(!form || !form.a) return;
  const rows=(state.reg.registos||[]).filter(r=> r.alunoId===form.a && (!form.i || r.data>=form.i) && (!form.f || r.data<=form.f))
    .map(r=> ({ Data:r.data, Oficina:r.disciplinaId, Numero:r.numeroLicao||'', Sumario:r.sumario||'', Presenca:(r.status==='P'?'Presente':(r.status==='A'?'Ausente (injust.)':(r.status==='J'?'J (just.)':''))) }));
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows),'Aluno');
  const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  a.download=`aluno_${form.a}_${form.i||'ini'}_${form.f||'fim'}.xlsx`;
  a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1200);
}
function humanStatus(r){ return r.status==='P'?'Presente':(r.status==='A'?'Ausente (injust.)':(r.status==='J'?'J (just.)':'')); }
async function pickDisciplinaSwal(){ const all=(state.config.disciplinas||[]).map(d=>[String(d.id), d.nome||d.id]); const opts=Object.fromEntries(all);
  const {value:did}=await Swal.fire({title:'Escolhe a disciplina', input:'select', inputOptions:opts, inputPlaceholder:'Disciplina/Oficina', showCancelButton:true}); return did||null;
}
async function exportDisciplinaPDF(){ if(!(window.jspdf&&window.jspdf.jsPDF)){ Swal.fire('Erro','jsPDF não disponível','error'); return; }
  const email=getAccountEmail(); const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email); if(!prof){ Swal.fire('Erro','Professor não reconhecido','error'); return; }
  const did=await pickDisciplinaSwal(); if(!did) return;
  const { value: form } = await Swal.fire({ title:'Intervalo de datas (opcional)', html:`<input id="di" class="swal2-input" type="date"><input id="df" class="swal2-input" type="date">`, confirmButtonText:'Exportar', showCancelButton:true, preConfirm:()=>({ i:$('#di').value, f:$('#df').value }) }); if(form===undefined) return;
  const rows=(state.reg.registos||[]).filter(r=> r.professorId===prof.id && r.disciplinaId===did && (!form.i || r.data>=form.i) && (!form.f || r.data<=form.f))
    .map(r=> [r.data, r.alunoId, r.numeroLicao||'', r.sumario||'', humanStatus(r)]);
  const doc=new window.jspdf.jsPDF({unit:'pt',format:'a4'}); const discNome=(state.config.disciplinas||[]).find(d=>String(d.id)===String(did))?.nome||did;
  const hdr=`${discNome} • ${prof.nome||prof.id} • ${form?.i||'…'} a ${form?.f||'…'}`; doc.text(hdr,40,40);
  doc.autoTable({startY:60, head:[['Data','Aluno','Nº','Sumário','Presença']], body:rows, styles:{fontSize:9}});
  doc.save(`disciplina_${did}_${prof.id}_${form?.i||'ini'}_${form?.f||'fim'}.pdf`);
}
async function exportDisciplinaXLSX(){ const email=getAccountEmail(); const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email); if(!prof){ Swal.fire('Erro','Professor não reconhecido','error'); return; }
  const did=await pickDisciplinaSwal(); if(!did) return;
  const { value: form } = await Swal.fire({ title:'Intervalo de datas (opcional)', html:`<input id="dix" class="swal2-input" type="date"><input id="dfx" class="swal2-input" type="date">`, confirmButtonText:'Exportar', showCancelButton:true, preConfirm:()=>({ i:$('#dix').value, f:$('#dfx').value }) }); if(form===undefined) return;
  const rows=(state.reg.registos||[]).filter(r=> r.professorId===prof.id && r.disciplinaId===did && (!form.i || r.data>=form.i) && (!form.f || r.data<=form.f))
    .map(r=> ({ Data:r.data, Aluno:r.alunoId, Numero:r.numeroLicao||'', Sumario:r.sumario||'', Presenca:humanStatus(r) }));
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Disciplina');
  const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  a.download=`disciplina_${did}_${prof.id}_${form?.i||'ini'}_${form?.f||'fim'}.xlsx`; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1200);
}

/* ====== Import simples (XLSX/JSON) ====== */
document.addEventListener('change', async ev=>{
  if(ev.target && ev.target.id==='fileImport'){
    const files=ev.target.files; if(!files || !files.length) return;
    for(const f of files){
      const name=f.name.toLowerCase();
      if(name.endsWith('.xlsx')){
        const data=await f.arrayBuffer();
        const wb=XLSX.read(data);
        const rows = wb.SheetNames.includes('Oficinas') ? XLSX.utils.sheet_to_json(wb.Sheets['Oficinas']) : [];
        if(rows.length){
          const novas = rows.map(r=>({
            id:String(r.id||'').trim(),
            professorId:String(r.professorId||'').trim(),
            disciplinaId:String(r.disciplinaId||'').trim(),
            alunoIds:String(r.alunoIds||'').split(',').map(s=>s.trim()).filter(Boolean),
            diaSemana:Number(r.diaSemana||0),
            horaInicio:String(r.horaInicio||'').trim(),
            horaFim:String(r.horaFim||'').trim(),
            sala:String(r.sala||'').trim()
          })).filter(x=> x.id && x.professorId && x.disciplinaId && x.alunoIds.length && x.diaSemana>=1 && x.diaSemana<=7 && x.horaInicio && x.horaFim);
          state.config.oficinas = []; // será afinado (merge/substituir) na fase de Importações
          const byId=new Map(state.config.oficinas.map(o=>[String(o.id),o]));
          novas.forEach(n=> byId.set(String(n.id), n));
          state.config.oficinas = Array.from(byId.values());
        }
        const first = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]] || {});
        if(first.length && !(state.config.professores||[]).length){
          const map=first.map(r=>({ id:r.id||r.ID||r.Codigo||r.codigo, nome:r.nome||r.Nome||r.NOME, email:r.email||r.Email||r.EMAIL })).filter(x=>x.id && x.nome);
          if(map.length) state.config.professores = map;
        }
        await graphSave(CFG_PATH,state.config);
        localStorage.setItem('esp_config',JSON.stringify(state.config));
        Swal.fire('Importado','XLSX importado.','success');
      } else if(name.endsWith('.json')){
        const txt=await f.text();
        try{
          state.config=JSON.parse(txt);
          await graphSave(CFG_PATH, state.config);
          localStorage.setItem('esp_config',JSON.stringify(state.config));
          Swal.fire('Importado','JSON importado e guardado.','success');
        }catch{ Swal.fire('Erro','JSON inválido','error'); }
      }
    }
  }
});

/* ====== Auto-save & Backup ====== */
let autosaveTimer=null;
function autoSaveConfig(){
  if(autosaveTimer) clearTimeout(autosaveTimer);
  autosaveTimer=setTimeout(async()=>{
    try{
      await graphSave(CFG_PATH, state.config);
      localStorage.setItem('esp_config', JSON.stringify(state.config));
      updateSync('💾 guardado');
    }catch(e){
      console.warn('auto-save', e);
      updateSync('⚠ offline');
    }
  }, 800);
}
async function createBackupIfExists(){
  try{
    const current = state.config || JSON.parse(localStorage.getItem('esp_config')||'{}');
    if(!current) return null;
    const now=new Date();
    const ts = now.getFullYear().toString().padStart(4,'0')
              +(now.getMonth()+1).toString().padStart(2,'0')
              +now.getDate().toString().padStart(2,'0')+'_'
              +now.getHours().toString().padStart(2,'0')
              +now.getMinutes().toString().padStart(2,'0');
    const backupPath = BACKUP_FOLDER+`/config_especial_${ts}.json`;
    await graphSave(backupPath, current);
    toast('Backup criado'); return backupPath;
  }catch(e){ console.warn(e); return null; }
}
async function restoreBackup(){
  try{
    updateSync('🔁 a ler backups...');
    const items=await graphList(BACKUP_FOLDER);
    const onlyCfg=items.filter(it=> it?.name?.startsWith('config_especial_') && it?.name?.endsWith('.json'))
                      .sort((a,b)=> a.name<b.name?1:-1);
    if(!onlyCfg.length){ Swal.fire('Restauração','Sem backups.','info'); updateSync('—'); return; }
    const options={}; onlyCfg.forEach(f=> options[f.name]=f.name);
    const { value: pick }=await Swal.fire({title:'Restaurar backup',input:'select',inputOptions:options,inputPlaceholder:'Escolhe o ficheiro',showCancelButton:true});
    if(!pick){ updateSync('—'); return; }
    updateSync('🔁 a restaurar...');
    const content=await graphLoad(`${BACKUP_FOLDER}/${pick}`);
    if(!content){ Swal.fire('Erro','Falha a ler o backup.','error'); updateSync('⚠ offline'); return; }
    await graphSave(CFG_PATH,content);
    state.config=content;
    localStorage.setItem('esp_config',JSON.stringify(state.config));
    toast('Configuração restaurada');
    renderHoje(); renderRegList();
    updateSync('💾 guardado');
  }catch(e){
    console.warn(e);
    Swal.fire('Aviso','Não foi possível restaurar.','warning');
    updateSync('⚠ offline');
  }
}

/* ====== Export conf/reg ====== */
function downloadBlob(filename, blob){ const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1200); }
function download(filename, data){ const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'}); downloadBlob(filename, blob); }
function exportConfigXlsx(){
  if(typeof XLSX==='undefined'){ alert('XLSX não carregou'); return; }
  const cfg = state.config || {professores:[],alunos:[],disciplinas:[],oficinas:[],calendario:{}};
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cfg.professores||[]),'Professores');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cfg.alunos||[]),'Alunos');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cfg.disciplinas||[]),'Disciplinas');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cfg.oficinas||[]),'Oficinas');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([cfg.calendario||{}]),'Calendario');
  const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  downloadBlob(`config_${new Date().toISOString().slice(0,10)}.xlsx`, new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
}
function exportRegXlsx(){
  if(typeof XLSX==='undefined'){ alert('XLSX não carregou'); return; }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet((state.reg?.registos)||[]), 'Registos');
  const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  downloadBlob(`registos_${new Date().toISOString().slice(0,10)}.xlsx`, new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
}

/* ====== Máscara HH:MM ====== */
function attachHHMMMask(input){
  if(!input) return;
  input.addEventListener('input', ()=>{
    let v=input.value.replace(/[^0-9]/g,'');
    if(v.length>4) v=v.slice(0,4);
    if(v.length>=3) v=v.slice(0,2)+':'+v.slice(2);
    input.value=v;
  });
  input.addEventListener('blur', ()=>{
    const m=/(\d{1,2}):?(\d{0,2})/.exec(input.value.replace(/[^0-9]/g,''));
    if(!m){ input.value=''; return; }
    let h=m[1], mm=m[2]||'';
    if(h.length===1) h='0'+h;
    if(mm.length===1) mm='0'+mm;
    h = Math.min(23, Math.max(0, parseInt(h||'0',10))).toString().padStart(2,'0');
    if(mm===''){ input.value=h+':00'; return; }
    mm= Math.min(59, Math.max(0, parseInt(mm||'0',10))).toString().padStart(2,'0');
    input.value = `${h}:${mm}`;
  });
}

/* ====== Admin (CRUD + validações + paginação + ordenação + export) ====== */
const ADMIN_SCHEMAS = {
  professores: [
    { key:'id',    label:'ID',    type:'text',   required:true },
    { key:'nome',  label:'Nome',  type:'text',   required:true },
    { key:'email', label:'Email', type:'email',  required:false },
    { key:'role',  label:'Role',  type:'select', required:false, options:['','admin'] }
  ],
  alunos: [
    { key:'id',     label:'ID',     type:'text',  required:true },
    { key:'nome',   label:'Nome',   type:'text',  required:true },
    { key:'numero', label:'Número', type:'text',  required:false },
    { key:'turma',  label:'Turma',  type:'text',  required:false }
  ],
  disciplinas: [
    { key:'id',   label:'ID',   type:'text', required:true },
    { key:'nome', label:'Nome', type:'text', required:true }
  ],
  oficinas: [
    { key:'id',          label:'ID',                               type:'text',     required:true },
    { key:'professorId', label:'Professor',                        type:'select',   required:true, optionsFrom:'professores' },
    { key:'disciplinaId',label:'Disciplina',                       type:'select',   required:true, optionsFrom:'disciplinas' },
    { key:'alunoIds',    label:'Alunos (IDs separados por vírg.)', type:'textarea', required:true },
    { key:'diaSemana',   label:'Dia da semana (1=Seg..7=Dom)',     type:'number',   required:true, min:1, max:7 },
    { key:'horaInicio',  label:'Hora início (HH:MM)',              type:'text',     required:true },
    { key:'horaFim',     label:'Hora fim (HH:MM)',                 type:'text',     required:true },
    { key:'sala',        label:'Sala',                             type:'text',     required:false }
  ]
};
function adminGetData(entity){
  const cfg = state.config || {};
  if(entity==='professores') return cfg.professores || [];
  if(entity==='alunos')      return cfg.alunos      || [];
  if(entity==='disciplinas') return cfg.disciplinas || [];
  if(entity==='oficinas')    return cfg.oficinas    || [];
  return [];
}
function adminSetData(entity, rows){
  if(!Array.isArray(rows)) rows = [];
  if(entity==='professores') state.config.professores = rows;
  if(entity==='alunos')      state.config.alunos      = rows;
  if(entity==='disciplinas') state.config.disciplinas = rows;
  if(entity==='oficinas')    state.config.oficinas    = rows;
}

/* Estado de grelha */
let ADMIN_SELECTED_ID=null, ADMIN_LAST_FILTER='', ADMIN_SORT_KEY=null, ADMIN_SORT_DIR='asc', ADMIN_PAGE=1, ADMIN_PAGE_SIZE=25;

function adminOptionsFrom(entity){
  const list=adminGetData(entity);
  if(entity==='professores') return list.map(p=>[String(p.id), `${p.nome||p.id} <${p.email||''}>`]);
  if(entity==='disciplinas') return list.map(d=>[String(d.id), d.nome||d.id]);
  return list.map(x=>[String(x.id), String(x.id)]);
}
function adminBuildFormHtml(entity,item){
  const schema=ADMIN_SCHEMAS[entity]||[];
  const asSelect=(s,value)=>{
    let opts='';
    if(s.optionsFrom){
      const pairs=adminOptionsFrom(s.optionsFrom);
      if(value && !pairs.some(([v])=>String(v)===String(value))) pairs.unshift([String(value),String(value)]);
      opts=pairs.map(([v,l])=>`<option value="${String(v)}"${String(v)===String(value)?' selected':''}>${l}</option>`).join('');
    }else{
      const pairs=(s.options||[]).map(v=>[v,v]);
      if(value && !pairs.some(([v])=>String(v)===String(value))) pairs.unshift([String(value),String(value)]);
      opts=pairs.map(([v,l])=>`<option value="${String(v)}"${String(v)===String(value)?' selected':''}>${l}</option>`).join('');
    }
    return `<select id="f_${s.key}" class="swal2-input">${opts}</select>`;
  };
  return `<div style="text-align:left">
    ${schema.map(s=>{
      const v=(item?.[s.key] ?? '');
      const label=`<label style="display:block;margin-top:6px">${s.label}${s.required?' *':''}</label>`;
      if(s.type==='select')  return `${label}${asSelect(s,v)}`;
      if(s.type==='textarea')return `${label}<textarea id="f_${s.key}" class="swal2-textarea" rows="3" placeholder="${s.label}">${(Array.isArray(v)?v.join(','):v)||''}</textarea>`;
      if(s.type==='number')  return `${label}<input id="f_${s.key}" class="swal2-input" type="number" ${s.min?'min="'+s.min+'"':''} ${s.max?'max="'+s.max+'"':''} value="${v}">`;
      const type=s.type||'text';
      return `${label}<input id="f_${s.key}" class="swal2-input" type="${type}" value="${v}" placeholder="${s.label}">`;
    }).join('')}
  </div>`;
}
function adminReadForm(entity){
  const schema=ADMIN_SCHEMAS[entity]||[];
  const obj={};
  schema.forEach(s=>{
    const el=document.getElementById(`f_${s.key}`);
    let val=(el?.value??'').trim();
    if(s.type==='number' && val!=='') val=Number(val);
    if(entity==='oficinas' && s.key==='alunoIds'){
      val = val ? val.split(',').map(x=>x.trim()).filter(Boolean) : [];
    }
    obj[s.key]=val;
  });
  return obj;
}
function adminValidate(entity,obj,isEdit=false,oldId=null){
  if(!obj || typeof obj!=='object') return 'Objeto inválido.';
  const schema=ADMIN_SCHEMAS[entity]||[];
  for(const s of schema){
    if(s.required){
      if(s.key==='alunoIds' && Array.isArray(obj.alunoIds) && obj.alunoIds.length===0) return 'A lista de alunos não pode estar vazia.';
      if(s.key!=='alunoIds' && (obj[s.key]===undefined || obj[s.key]==='')) return `Campo obrigatório: ${s.label}`;
    }
    if(s.type==='email' && obj[s.key]){
      const ok=/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(obj[s.key]);
      if(!ok) return 'Email inválido.';
    }
  }
  if(!isEdit || (isEdit && String(obj.id)!==String(oldId))){
    const exists=adminGetData(entity).some(x=> String(x.id)===String(obj.id));
    if(exists) return `ID "${obj.id}" já existe em ${entity}.`;
  }
  if(entity==='oficinas'){
    if(!(Number(obj.diaSemana)>=1 && Number(obj.diaSemana)<=7)) return 'Dia da semana inválido (1..7).';
    const hhmm=/^\d{2}:\d{2}$/;
    if(!hhmm.test(obj.horaInicio)||!hhmm.test(obj.horaFim)) return 'Hora início/fim inválida (use HH:MM).';
    const profOk=adminGetData('professores').some(p=> String(p.id)===String(obj.professorId));
    if(!profOk) return `ProfessorId "${obj.professorId}" não existe.`;
    const discOk=adminGetData('disciplinas').some(d=> String(d.id)===String(obj.disciplinaId));
    if(!discOk) return `DisciplinaId "${obj.disciplinaId}" não existe.`;
    const alunos=new Set(adminGetData('alunos').map(a=>String(a.id)));
    const desconhecidos=(obj.alunoIds||[]).filter(a=> !alunos.has(String(a)));
    if(desconhecidos.length) return `Aluno(s) desconhecido(s): ${desconhecidos.join(', ')}`;
  }
  return null;
}
function adminRenderGrid(){
  const entity=$('#adminEntity')?.value || 'professores';
  const target=$('#adminGrid'); if(!target) return;

  const all=adminGetData(entity).slice();
  const filter=($('#adminSearch')?.value||'').trim().toLowerCase();
  ADMIN_LAST_FILTER=filter;

  let data = filter ? all.filter(row => JSON.stringify(row).toLowerCase().includes(filter)) : all;

  const schema=ADMIN_SCHEMAS[entity]||[];
  const cols=schema.map(s=>s.key);
  const headers=schema.map(s=>s.label);

  if(ADMIN_SORT_KEY && cols.includes(ADMIN_SORT_KEY)){
    const key=ADMIN_SORT_KEY, dir=(ADMIN_SORT_DIR==='desc'?-1:1);
    data.sort((a,b)=>{
      const va=(a?.[key]??'').toString().toLowerCase();
      const vb=(b?.[key]??'').toString().toLowerCase();
      if(va<vb) return -1*dir; if(va>vb) return 1*dir; return 0;
    });
  }

  const total=data.length;
  const pages=Math.max(1,Math.ceil(total/ADMIN_PAGE_SIZE));
  if(ADMIN_PAGE>pages) ADMIN_PAGE=pages; if(ADMIN_PAGE<1) ADMIN_PAGE=1;
  const ini=(ADMIN_PAGE-1)*ADMIN_PAGE_SIZE; const fim=ini+ADMIN_PAGE_SIZE;
  const pageRows=data.slice(ini,fim);

  const ths=headers.map((h,i)=>{
    const k=cols[i]; const active=(k===ADMIN_SORT_KEY); const arrow=active?(ADMIN_SORT_DIR==='asc'?' ▲':' ▼'):'';
    return `<th data-sort="${k}" style="cursor:pointer">${h}${arrow}</th>`;
  }).join('');

  const trs=pageRows.map(row=>{
    const id=String(row.id??'');
    const tds=cols.map(k=>`<td>${Array.isArray(row[k])?row[k].join(', '):(row[k]??'')}</td>`).join('');
    const selected=(ADMIN_SELECTED_ID && ADMIN_SELECTED_ID===id)?' style="background:var(--primary-hover)"':'';
    return `<tr data-id="${id}"${selected}>${tds}</tr>`;
  }).join('') || `<tr><td colspan="${cols.length}" class="muted">Sem registos</td></tr>`;

  target.innerHTML=`
    <div style="overflow:auto">
      <table class="table" style="width:100%">
        <thead><tr>${ths}</tr></thead>
        <tbody>${trs}</tbody>
      </table>
    </div>
    <div class="sec-head" style="margin-top:8px">
      <div class="muted">Total ${total} • Página ${ADMIN_PAGE}/${pages}</div>
      <div class="sec-actions">
        <label class="muted">Linhas:</label>
        <select id="adminPageSize" class="input" style="width:auto">
          ${[10,25,50,100].map(n=>`<option value="${n}" ${n===ADMIN_PAGE_SIZE?'selected':''}>${n}</option>`).join('')}
        </select>
        <button id="adminPrevPage" class="btn" ${ADMIN_PAGE<=1?'disabled':''}>◀</button>
        <button id="adminNextPage" class="btn" ${ADMIN_PAGE>=pages?'disabled':''}>▶</button>
      </div>
    </div>
  `;

  // Seleção de linha
  target.querySelectorAll('tbody tr[data-id]').forEach(tr=>{
    tr.addEventListener('click', ()=>{
      const id=tr.getAttribute('data-id');
      ADMIN_SELECTED_ID=id;
      target.querySelectorAll('tbody tr').forEach(x=> x.style.background='');
      tr.style.background='var(--primary-hover)';
      $('#btnAdminEditar')?.removeAttribute('disabled');
      $('#btnAdminApagar')?.removeAttribute('disabled');
    });
  });

  // Ordenação
  target.querySelectorAll('th[data-sort]').forEach(th=>{
    th.addEventListener('click', ()=>{
      const key=th.getAttribute('data-sort');
      if(ADMIN_SORT_KEY===key){ ADMIN_SORT_DIR=(ADMIN_SORT_DIR==='asc'?'desc':'asc'); }
      else { ADMIN_SORT_KEY=key; ADMIN_SORT_DIR='asc'; }
      adminRenderGrid();
    });
  });

  // Paginação
  $('#adminPrevPage')?.addEventListener('click', ()=>{ ADMIN_PAGE=Math.max(1,ADMIN_PAGE-1); adminRenderGrid(); });
  $('#adminNextPage')?.addEventListener('click', ()=>{ ADMIN_PAGE=ADMIN_PAGE+1; adminRenderGrid(); });
  $('#adminPageSize')?.addEventListener('change', (e)=>{ ADMIN_PAGE_SIZE=parseInt(e.target.value,10)||25; ADMIN_PAGE=1; adminRenderGrid(); });

  if(!ADMIN_SELECTED_ID){
    $('#btnAdminEditar')?.setAttribute('disabled','true');
    $('#btnAdminApagar')?.setAttribute('disabled','true');
  }
}

/* CRUD: Novo/Editar/Apagar */
async function adminNovo(){
  const entity=$('#adminEntity')?.value || 'professores';
  const html=adminBuildFormHtml(entity,null);
  const { value: ok } = await Swal.fire({
    title:`Novo ${entity}`, html, showCancelButton:true, confirmButtonText:'Guardar',
    didOpen:(el)=>{ if(entity==='oficinas'){ attachHHMMMask(el.querySelector('#f_horaInicio')); attachHHMMMask(el.querySelector('#f_horaFim')); } },
    preConfirm:()=> true
  });
  if(!ok) return;
  const obj=adminReadForm(entity);
  const err=adminValidate(entity,obj,false,null);
  if(err){ Swal.fire('Erro', err, 'error'); return; }

  const arr=adminGetData(entity).slice();
  arr.push(obj);
  adminSetData(entity,arr);
  await graphSave(CFG_PATH,state.config);
  localStorage.setItem('esp_config',JSON.stringify(state.config));
  autoSaveConfig();
  toast('Registo criado');
  ADMIN_SELECTED_ID=String(obj.id);
  adminRenderGrid();
  renderHoje();
}
async function adminEditar(){
  const entity=$('#adminEntity')?.value || 'professores';
  if(!ADMIN_SELECTED_ID){ Swal.fire('Info','Seleciona uma linha para editar.','info'); return; }
  const arr=adminGetData(entity).slice();
  const idx=arr.findIndex(x=> String(x.id)===String(ADMIN_SELECTED_ID));
  if(idx<0){ Swal.fire('Erro','Registo não encontrado.','error'); return; }

  const original=arr[idx];
  const html=adminBuildFormHtml(entity,original);
  const { value: ok } = await Swal.fire({
    title:`Editar ${entity}`, html, showCancelButton:true, confirmButtonText:'Guardar',
    didOpen:(el)=>{ if(entity==='oficinas'){ attachHHMMMask(el.querySelector('#f_horaInicio')); attachHHMMMask(el.querySelector('#f_horaFim')); } },
    preConfirm:()=> true
  });
  if(!ok) return;

  const updated=adminReadForm(entity);
  const err=adminValidate(entity,updated,true,original.id);
  if(err){ Swal.fire('Erro', err, 'error'); return; }

  arr[idx]=updated;
  adminSetData(entity,arr);
  await graphSave(CFG_PATH,state.config);
  localStorage.setItem('esp_config',JSON.stringify(state.config));
  autoSaveConfig();
  toast('Registo atualizado');
  ADMIN_SELECTED_ID=String(updated.id);
  adminRenderGrid();
  renderHoje();
}
async function adminApagar(){
  const entity=$('#adminEntity')?.value || 'professores';
  if(!ADMIN_SELECTED_ID){ Swal.fire('Info','Seleciona uma linha para apagar.','info'); return; }

  const { isConfirmed } = await Swal.fire({
    title:'Confirmar remoção',
    text:`Quer mesmo apagar o ID "${ADMIN_SELECTED_ID}" de ${entity}?`,
    icon:'warning', showCancelButton:true, confirmButtonText:'Sim, apagar'
  });
  if(!isConfirmed) return;

  let arr=adminGetData(entity).slice();
  arr=arr.filter(x=> String(x.id)!==String(ADMIN_SELECTED_ID));
  adminSetData(entity,arr);
  await graphSave(CFG_PATH,state.config);
  localStorage.setItem('esp_config',JSON.stringify(state.config));
  autoSaveConfig();
  toast('Removido');
  ADMIN_SELECTED_ID=null;
  adminRenderGrid();
  renderHoje();
}

/* Export lista & Export registos (Admin) */
function adminExportListaXLSX(){
  const entity=$('#adminEntity')?.value || 'professores';
  const rows=adminGetData(entity);
  if(typeof XLSX==='undefined'){ alert('XLSX não carregou'); return; }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), entity);
  const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  a.download=`${entity}_${new Date().toISOString().slice(0,10)}.xlsx`;
  a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1200);
}
async function adminExportRegistos(){
  const { value: step1 } = await Swal.fire({
    title:'Exportar registos',
    input:'select',
    inputOptions:{ aluno:'Por Aluno', professor:'Por Professor', disciplina:'Por Disciplina' },
    inputPlaceholder:'Escolha o tipo', showCancelButton:true
  });
  if(!step1) return;

  let pairs=[];
  if(step1==='aluno'){
    pairs=(state.config.alunos||[]).map(a=>[String(a.id), `${a.nome||a.id} (${a.id})`]);
  }else if(step1==='professor'){
    pairs=(state.config.professores||[]).map(p=>[String(p.id), `${p.nome||p.id} <${p.email||''}>`]);
  }else{
    pairs=(state.config.disciplinas||[]).map(d=>[String(d.id), d.nome||d.id]);
  }
  const options=Object.fromEntries(pairs);

  const { value: pick } = await Swal.fire({
    title: step1==='aluno'?'Aluno':(step1==='professor'?'Professor':'Disciplina'),
    input:'select', inputOptions:options, inputPlaceholder:'Escolha', showCancelButton:true
  });
  if(!pick) return;

  const { value: form } = await Swal.fire({
    title:'Intervalo e formato',
    html:`<input id="ex_di" class="swal2-input" type="date" placeholder="De (opcional)">
          <input id="ex_df" class="swal2-input" type="date" placeholder="Até (opcional)">
          <select id="ex_fmt" class="swal2-input"><option value="pdf">PDF</option><option value="xlsx">XLSX</option></select>`,
    confirmButtonText:'Exportar', showCancelButton:true,
    preConfirm:()=>({ i:$('#ex_di').value, f:$('#ex_df').value, fmt:$('#ex_fmt').value })
  });
  if(!form) return;

  const reg=(state.reg.registos||[]);
  let rows=[];
  if(step1==='aluno')      rows=reg.filter(r => r.alunoId===pick);
  else if(step1==='professor') rows=reg.filter(r => String(r.professorId)===String(pick));
  else                      rows=reg.filter(r => String(r.disciplinaId)===String(pick));
  if(form.i) rows=rows.filter(r=> r.data>=form.i);
  if(form.f) rows=rows.filter(r=> r.data<=form.f);
  rows.sort((a,b)=> a.data<b.data ? -1 : 1);

  if(form.fmt==='xlsx'){
    if(typeof XLSX==='undefined'){ alert('XLSX não carregou'); return; }
    const out=rows.map(r=>({ Data:r.data, Aluno:r.alunoId, Professor:r.professorId, Disciplina:r.disciplinaId, Numero:r.numeroLicao||'', Sumario:r.sumario||'', Presenca:humanStatus(r) }));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(out),'Registos');
    const bin=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(new Blob([bin],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    a.download=`registos_${step1}_${pick}_${form.i||'ini'}_${form.f||'fim'}.xlsx`;
    a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1200);
  }else{
    if(!(window.jspdf && window.jspdf.jsPDF)){ Swal.fire('Erro','jsPDF não disponível','error'); return; }
    const hdr = step1==='aluno'?`Aluno ${pick}`:(step1==='professor'?`Professor ${pick}`:`Disciplina ${pick}`);
    const body = rows.map(r => [r.data, r.alunoId, r.professorId, r.disciplinaId, r.numeroLicao||'', r.sumario||'', humanStatus(r)]);
    const doc=new window.jspdf.jsPDF({unit:'pt',format:'a4'});
    doc.text(`${hdr} • ${form.i||'…'} a ${form.f||'…'}`,40,40);
    doc.autoTable({ startY:60, head:[['Data','Aluno','Professor','Disciplina','Nº','Sumário','Presença']], body, styles:{fontSize:9} });
    doc.save(`registos_${step1}_${pick}_${form.i||'ini'}_${form.f||'fim'}.pdf`);
  }
}

/* ====== Nova oficina rápida ====== */
async function novaOficinaRapida(){
  if(isAdmin()){ Swal.fire('Nota','Crie oficinas em massa via XLSX na Administração.','info'); return; }
  const email=getAccountEmail();
  const prof=(state.config.professores||[]).find(p=> (p.email||'').toLowerCase()===email);
  if(!prof){ Swal.fire('Aviso','Professor não reconhecido na configuração.','warning'); return; }

  const { value: form } = await Swal.fire({
    title:'Nova oficina',
    html:`<div style="text-align:left">
      <label>ID</label><input id="o_id" class="swal2-input" value="sess${crypto.randomUUID().toString().slice(-4)}">
      <label>Disciplina/Oficina (id)</label><input id="o_disc" class="swal2-input" value="${(state.config.disciplinas?.[0]?.id||'of_port')}">
      <label>Alunos (IDs separados por vírgulas)</label><input id="o_al" class="swal2-input" placeholder="a001,a002">
      <label>Dia da semana (1=Seg,..7=Dom)</label><input id="o_dw" class="swal2-input" value="${diaSemana(new Date().toISOString().slice(0,10))}">
      <label>Hora início</label><input id="o_ini" class="swal2-input" value="10:00">
      <label>Hora fim</label><input id="o_fim" class="swal2-input" value="10:50">
      <label>Sala</label><input id="o_sala" class="swal2-input" value="CAA">
    </div>`,
    confirmButtonText:'Guardar', showCancelButton:true,
    didOpen:(el)=>{ attachHHMMMask(el.querySelector('#o_ini')); attachHHMMMask(el.querySelector('#o_fim')); },
    preConfirm:()=>({
      id:document.getElementById('o_id').value.trim(),
      disciplinaId:document.getElementById('o_disc').value.trim(),
      alunoIds:(document.getElementById('o_al').value||'').split(',').map(s=>s.trim()).filter(Boolean),
      diaSemana:Number(document.getElementById('o_dw').value||1),
      horaInicio:document.getElementById('o_ini').value.trim(),
      horaFim:document.getElementById('o_fim').value.trim(),
      sala:document.getElementById('o_sala').value.trim()
    })
  });
  if(!form || !form.id) return;
  form.professorId = prof.id;

  if(!Array.isArray(state.config.oficinas)) state.config.oficinas = [];
  state.config.oficinas.push(form);

  await graphSave(CFG_PATH,state.config);
  localStorage.setItem('esp_config',JSON.stringify(state.config));
  toast('Oficina criada');
  renderHoje();
}

/* ====== Bindings ====== */
document.addEventListener('DOMContentLoaded', async ()=>{
  checkDeps();

  $('#btnMsLogin')?.addEventListener('click', ()=>ensureLogin());
  $('#btnMsLogout')?.addEventListener('click', ()=>ensureLogout());
  $('#btnRefreshDay')?.addEventListener('click', ()=>renderHoje());
  $('#btnCriarOficina')?.addEventListener('click', ()=>novaOficinaRapida());

  $('#btnBackupNow')?.addEventListener('click', async ()=>{ const b=await createBackupIfExists(); if(b) Swal.fire('Backup criado', b, 'success'); });
  $('#btnExportCfgJson')?.addEventListener('click', ()=>download('config_especial.json', state.config||{}));
  $('#btnExportRegJson')?.addEventListener('click', ()=>download('2registos_alunos.json', state.reg||{versao:'v2', registos:[]}));
  $('#btnExportCfgXlsx')?.addEventListener('click', ()=>exportConfigXlsx());
  $('#btnExportRegXlsx')?.addEventListener('click', ()=>exportRegXlsx());
  $('#btnRestoreBackup')?.addEventListener('click', ()=>restoreBackup());

  $('#btnFiltrar')?.addEventListener('click', ()=>renderRegList());
  $('#btnPdfSemana')?.addEventListener('click', ()=>exportSemanalPDF());
  $('#btnXlsxSemana')?.addEventListener('click', ()=>exportSemanalXLSX());
  $('#btnPdfAluno')?.addEventListener('click', ()=>exportAlunoPDF());
  $('#btnXlsxAluno')?.addEventListener('click', ()=>exportAlunoXLSX());
  $('#btnPdfDisciplina')?.addEventListener('click', ()=>exportDisciplinaPDF());
  $('#btnXlsxDisciplina')?.addEventListener('click', ()=>exportDisciplinaXLSX());

  const theme = localStorage.getItem('esp_theme')
    || (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  if(theme==='dark') document.documentElement.setAttribute('data-theme','dark');

  await initMsal();

  const c=localStorage.getItem('esp_config'); if(c) state.config=JSON.parse(c);
  const r=localStorage.getItem('esp_reg');    if(r) state.reg=JSON.parse(r);
  if(!state.config) state.config={professores:[],alunos:[],disciplinas:[],oficinas:[],calendario:{}};
  if(!state.reg)    state.reg={versao:'v2',registos:[]};

  await loadConfigAndReg();
});