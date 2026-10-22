// ESJA Legal — símaforrit. Les stöðu tölvanna og skrifar beiðnir í OneDrive notandans
// (Microsoft Graph, innskráning notandans). Ekkert efni er í kóðanum og ekkert er geymt
// í símanum nema innskráningin og listi yfir eigin beiðnir (auðkenni).
'use strict';
const ST = window.SIMI_STILLINGAR || {};
let TOLVUR = [];   // fundnar úr stada-<TÖLVA>.json í möppunni
const SCOPES = ['Files.ReadWrite'];
const $ = s => document.querySelector(s);
const esc = t => String(t ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const bida = ms => new Promise(r => setTimeout(r, ms));

// ---------------------------------------------------------------- innskráning og Graph
let MSAL = null, REIKN = null;
// Greining innskráningar: hvort svar Microsoft barst, og síðustu skilaboð MSAL (án persónugagna).
const KOMA = { slod: location.href.replace(/(code|id_token|access_token|client_info|session_state)=[^&#]*/g, '$1=…'),
               svar: /[#&?](code|error|state)=/.test(location.href) };
const MSAL_LOG = [];
function greining() {
  let ss = []; try { ss = Object.keys(sessionStorage).filter(k => k.startsWith('msal')).map(k => k.split('.').slice(-2).join('.')); } catch (e) {}
  return 'Svar frá Microsoft: ' + (KOMA.svar ? 'já' : 'nei') + '\nSlóð við komu: ' + KOMA.slod +
    '\nReikningar í minni: ' + (MSAL ? MSAL.getAllAccounts().length : '-') + '\nsessionStorage: ' + (ss.join(', ') || 'tómt') +
    '\nVafri: ' + (navigator.standalone ? 'heimaskjár' : 'Safari/vafri') + '\n' + MSAL_LOG.slice(-10).join('\n');
}
async function raesaMsal() {
  if (ST.profun) return true;
  if (!ST.clientId || !ST.tenantId) { $('#inn-villa').textContent = 'Uppsetningu vantar (app-skráning í Entra).'; return false; }
  MSAL = new msal.PublicClientApplication({
    auth: { clientId: ST.clientId, authority: 'https://login.microsoftonline.com/' + ST.tenantId,
            redirectUri: location.origin + location.pathname.replace(/index\.html$/, ''),
            // Síðan notar #-leiðsögn; MSAL má ekki fletta aftur á upphafsslóðina eftir endurkomu.
            navigateToLoginRequestUrl: false },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true },
    system: { loggerOptions: { piiLoggingEnabled: false, logLevel: msal.LogLevel.Info,
      loggerCallback: (lvl, m, pii) => { if (!pii) MSAL_LOG.push(m.replace(/^\[[^\]]*\]\s*:\s*/, '').slice(0, 220)); } } } });
  await MSAL.initialize();
  try {
    const r = await MSAL.handleRedirectPromise();
    if (r && r.account) { REIKN = r.account; MSAL.setActiveAccount(r.account); history.replaceState(null, '', location.pathname + '#idag'); }
  } catch (e) { $('#inn-villa').textContent = 'Innskráning mistókst: ' + (e.errorCode || '') + ' ' + e.message; }
  REIKN = REIKN || MSAL.getActiveAccount() || MSAL.getAllAccounts()[0] || null;
  if (!REIKN && KOMA.svar && !$('#inn-villa').textContent)
    $('#inn-villa').textContent = 'Svar barst frá Microsoft en innskráningin vistaðist ekki.';
  let reynt = false; try { reynt = !!localStorage.getItem('esja_innskra_reynt'); } catch (e) {}
  if (REIKN) { try { localStorage.removeItem('esja_innskra_reynt'); } catch (e) {} }
  else if (reynt || KOMA.svar || $('#inn-villa').textContent) { const g = $('#greining'); g.textContent = 'Greining (sendu Claude):\n' + greining(); g.hidden = false; }
  return !!REIKN;
}
async function innskra() {
  if (!MSAL) return;
  try { localStorage.setItem('esja_innskra_reynt', new Date().toISOString()); } catch (e) {}
  try { await MSAL.loginRedirect({ scopes: SCOPES, prompt: 'select_account', redirectStartPage: location.pathname + '#idag' }); }
  catch (e) {
    if (e.errorCode === 'interaction_in_progress') {   // leif eftir fyrri tilraun
      try { Object.keys(sessionStorage).filter(k => k.includes('interaction.status')).forEach(k => sessionStorage.removeItem(k)); } catch (x) {}
      $('#inn-villa').textContent = 'Fyrri innskráning hékk. Ýttu aftur á Skrá inn.';
    } else $('#inn-villa').textContent = 'Innskráning hófst ekki: ' + (e.errorCode || '') + ' ' + e.message;
  }
}
async function tokn() {
  if (ST.profun) return 'profun';
  try { return (await MSAL.acquireTokenSilent({ scopes: SCOPES, account: REIKN })).accessToken; }
  catch (e) { await MSAL.acquireTokenRedirect({ scopes: SCOPES, account: REIKN }); throw e; }
}
const skraSlod = p => ST.graph + '/me/drive/root:/' + [ST.mappa, ...p.split('/').filter(Boolean)].map(encodeURIComponent).join('/');
async function graph(url, opt = {}) {
  const r = await fetch(url, { ...opt, cache: 'no-store',
    headers: { Authorization: 'Bearer ' + await tokn(), ...(opt.headers || {}) } });
  return r;
}
async function lesa(p) {           // JSON-skrá eða null ef hún er ekki til
  const r = await graph(skraSlod(p) + '?select=id,%40microsoft.graph.downloadUrl');
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('OneDrive svaraði ' + r.status);
  const j = await r.json(); const d = j['@microsoft.graph.downloadUrl'];
  if (!d) return null;
  const s = await fetch(d, { cache: 'no-store' });
  if (!s.ok) return null;
  return s.json();
}
async function lesaTexta(p) {
  const r = await graph(skraSlod(p) + '?select=id,%40microsoft.graph.downloadUrl');
  if (!r.ok) return null;
  const d = (await r.json())['@microsoft.graph.downloadUrl'];
  return d ? (await fetch(d, { cache: 'no-store' })).text() : null;
}
async function skrifa(p, efni, teg = 'application/json') {
  if (efni.size > 4 * 1024 * 1024) throw new Error('Skráin er stærri en 4 MB.');
  const r = await graph(skraSlod(p) + ':/content', { method: 'PUT', headers: { 'Content-Type': teg },
    body: typeof efni === 'string' || efni instanceof Blob ? efni : JSON.stringify(efni) });
  if (!r.ok) throw new Error('Vistun í OneDrive mistókst (' + r.status + ').');
  return r.json();
}
async function listi(p) {
  const r = await graph(skraSlod(p) + ':/children?$select=name,lastModifiedDateTime&$top=500');
  if (r.status === 404) return [];
  if (!r.ok) throw new Error('OneDrive svaraði ' + r.status);
  return (await r.json()).value || [];
}

// ---------------------------------------------------------------- staðan
let STADA = null, ALLAR = {}, BEIDNIR = {};
const iDag = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const aMorgun = () => { const d = new Date(Date.now() + 864e5); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const klukka = iso => iso ? iso.slice(11, 16) : '';
const VIKUDAGAR = ['sun.', 'mán.', 'þri.', 'mið.', 'fim.', 'fös.', 'lau.'];
const dags = iso => { if (!iso) return ''; const d = new Date(iso.slice(0, 10) + 'T12:00'); return VIKUDAGAR[d.getDay()] + ' ' + d.getDate() + '.' + (d.getMonth() + 1) + '.'; };
async function saekjaStodu() {
  TOLVUR = (await listi('')).map(f => (f.name.match(/^stada-(.+)\.json$/) || [])[1]).filter(Boolean);
  const s = await Promise.all(TOLVUR.map(t => lesa('stada-' + t + '.json').catch(() => null)));
  ALLAR = {}; TOLVUR.forEach((t, i) => { if (s[i]) ALLAR[t] = s[i]; });
  const radad = Object.values(ALLAR).sort((a, b) => (b.kl || '').localeCompare(a.kl || ''));
  STADA = radad[0] || null;
  BEIDNIR = {};
  radad.slice().reverse().forEach(a => Object.assign(BEIDNIR, a.beidnir || {}));
  return STADA;
}
function ferskleiki() {
  if (!STADA) return '<p class="s-villa">Engin staða fannst enn frá tölvunum. Kveiktu á annarri hvorri tölvunni.</p>';
  const aldur = (Date.now() - new Date(STADA.kl).getTime()) / 6e4;
  const gamalt = aldur > 30;
  const hin = TOLVUR.filter(t => t !== STADA.tolva).map(t => ALLAR[t] ? t + ' ' + dags(ALLAR[t].kl) + ' kl. ' + klukka(ALLAR[t].kl) : t + ': engin staða').join(' · ');
  return `<p class="s-ferskt ${gamalt ? 'gamalt' : ''}">Staða frá ${esc(STADA.tolva)} ${esc(dags(STADA.kl))} kl. ${esc(klukka(STADA.kl))}${gamalt ? ' — báðar tölvur gætu verið slökktar; aðgerðir bíða' : ''}. <span>${esc(hin)}</span></p>`;
}

// ---------------------------------------------------------------- sýnir
const SYNIR = ['idag', 'postur', 'athuga', 'minni'];
function syna() {
  const h = (location.hash || '#idag').slice(1).split('/')[0];
  const s = SYNIR.includes(h) ? h : 'idag';
  document.querySelectorAll('.s-syn').forEach(x => x.classList.toggle('virk', x.id === 'syn-' + s));
  document.querySelectorAll('.s-stika a').forEach(a => a.classList.toggle('virkt', a.dataset.s === s));
  $('#fab').hidden = s === 'minni';
  ({ idag: synaDag, postur: synaPost, athuga: synaAthuga, minni: synaMinni })[s]();
}
window.addEventListener('hashchange', syna);

function lina(x) {
  if (x.teg === 'thinghald') return `<div class="s-lina hreimur"><div class="s-titill">${esc(x.timi || '')} ${esc(x.tegund)}</div>
    <div class="s-undir">${esc(x.domstoll)}${x.salur ? ' · salur ' + esc(x.salur) : ''}${x.lengd ? ' · ' + esc(x.lengd) : ''}</div><div class="s-meta">${esc(x.mal)}</div></div>`;
  const brynt = x.logbundinn && x.dags <= aMorgun();
  return `<div class="s-lina ${brynt ? 'brynt' : 'vidv'}"><div class="s-titill">${esc(x.heiti)}</div>
    <div class="s-undir">${esc(x.flokkur)}${x.dags < iDag() ? ' · ' + esc(dags(x.dags)) : ''}${x.i_dagatali ? '' : ' · <span class="s-merki brynt">ekki í dagatali</span>'}</div>
    <div class="s-meta">${esc(x.mal)}</div></div>`;
}
async function synaDag(nytt) {
  if (nytt || !STADA) { $('#dagur').innerHTML = '<p class="s-tomt">Les …</p>'; try { await saekjaStodu(); } catch (e) { $('#dagur').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; return; } }
  const g = STADA && STADA.dagur;
  if (!g) { $('#dagur').innerHTML = ferskleiki(); return; }
  const a = iDag(), b = aMorgun();
  const dagsins = d => [...g.thinghold.filter(x => x.dags === d).sort((x, y) => (x.timi || '').localeCompare(y.timi || '')), ...g.frestir.filter(x => x.dags === d)];
  const idag = dagsins(a), amorgun = dagsins(b), lidid = g.frestir.filter(x => x.dags < a);
  const fram = d => { const t = new Date(Date.now() + d * 864e5).toISOString().slice(0, 10); return t; };
  const n14f = g.frestir.filter(x => x.dags > b && x.dags <= fram(14)).length, n14t = g.thinghold.filter(x => x.dags > b && x.dags <= fram(14)).length;
  $('#undirhaus').textContent = (idag.length + amorgun.length ? (idag.length + amorgun.length) + ' atriði í dag og á morgun' : 'Ekkert í dag og á morgun');
  $('#dagur').innerHTML = ferskleiki() + (g.villur.length ? `<p class="s-villa">${g.villur.map(esc).join('<br>')}</p>` : '') +
    `<section class="s-kafli"><h2>Í dag</h2>${idag.length ? idag.map(lina).join('') : '<p class="s-tomt">Ekkert í dag.</p>'}</section>
     <section class="s-kafli"><h2>Á morgun</h2>${amorgun.length ? amorgun.map(lina).join('') : '<p class="s-tomt">Ekkert á morgun.</p>'}</section>
     ${lidid.length ? `<section class="s-kafli"><details><summary>Liðið og óuppfyllt · ${lidid.length}</summary>${lidid.map(lina).join('')}</details></section>` : ''}
     <p class="s-tomt">Næstu 14 dagar: ${n14f} frestir · ${n14t} þinghöld.</p>
     <div><button class="s-takki" onclick="synaDag(true)">Uppfæra</button></div>`;
}

// ---------------------------------------------------------------- beiðnir
function nyttId() { const d = new Date(); return d.toISOString().slice(0, 19).replace(/[-:]/g, '') + '-' + Math.random().toString(36).slice(2, 8); }
function minarBeidnir() { try { return JSON.parse(localStorage.getItem('esja_beidnir') || '[]'); } catch (e) { return []; } }
function geymaBeidni(b) { try { const l = minarBeidnir().filter(x => x.id !== b.id); l.unshift(b); localStorage.setItem('esja_beidnir', JSON.stringify(l.slice(0, 40))); } catch (e) {} }
async function setjaIBidrod(b) {
  b.id = nyttId(); b.kl = new Date().toISOString().slice(0, 19); b.fra = 'simi';
  await skrifa('bidrod/' + b.id + '.json', b);
  geymaBeidni({ id: b.id, tegund: b.tegund, kl: b.kl, lysing: b.lysing || '' });
  return b;
}
const STADA_ORD = { sent: 'Sent', ostadfest: 'Sent — óstaðfest', villa: 'Ekki framkvæmt', lokid: 'Lokið', utrunnin: 'Útrunnin', thurr: 'Þurrkeyrsla' };
function beidniMerki(id) {
  const n = BEIDNIR[id];
  if (!n) return '<span class="s-merki bidur">Bíður tölvu</span>';
  return `<span class="s-merki ${n.stada === 'villa' ? 'brynt' : ''}">${esc(STADA_ORD[n.stada] || n.stada)} · ${esc(n.tolva)}</span>`;
}

// ---------------------------------------------------------------- póstur
let DRAG = null;
async function synaPost(nytt) {
  const h = location.hash.split('/');
  if (nytt || !STADA) { $('#postur').innerHTML = '<p class="s-tomt">Les …</p>'; try { await saekjaStodu(); } catch (e) { $('#postur').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; return; } }
  const drog = (STADA && STADA.drog) || [];
  $('#t-postur').textContent = drog.length || '';
  const sendar = minarBeidnir().filter(b => b.tegund === 'senda');
  if (h[1]) { DRAG = drog.find(d => d.id === decodeURIComponent(h[1])); return teiknaDrag(); }
  $('#postur').innerHTML = ferskleiki() + `<section class="s-kafli"><h2>Tillögur Claude · ${drog.length}</h2>
    ${drog.length ? drog.map((x, i) => `<button class="s-spjald" onclick="location.hash='#postur/'+encodeURIComponent(STADA.drog[${i}].id)">
      <div class="s-titill">${esc(x.sendandi || x.efni)}</div><div class="s-undir">${esc(x.efni)}</div>
      <div class="s-meta">${esc(x.tegund)}${x.aldur != null ? ' · ' + x.aldur + ' d.' : ''}${x.mal ? ' · ' + esc(x.mal) : ''}</div></button>`).join('')
      : '<p class="s-tomt">Engar tillögur bíða. Claude semur tillögur á tölvunni; þær birtast hér.</p>'}</section>
    ${sendar.length ? `<section class="s-kafli"><h2>Sendingar úr símanum</h2>${sendar.map(b => `<div class="s-lina"><div class="s-undir">${esc(b.lysing)}</div>
      <div class="s-meta">${esc(dags(b.kl))} kl. ${esc(klukka(b.kl))} · ${beidniMerki(b.id)}${BEIDNIR[b.id] ? '<br>' + esc(BEIDNIR[b.id].skilabod || '') : ''}</div></div>`).join('')}</section>` : ''}
    <div><button class="s-takki" onclick="synaPost(true)">Uppfæra</button></div>`;
}
function teiknaDrag(stadfesta) {
  const d = DRAG;
  if (!d) { $('#postur').innerHTML = '<p class="s-villa">Dragið er ekki lengur í stöðunni — það gæti hafa verið sent.</p><div><button class="s-takki" onclick="location.hash=\'#postur\'">‹ Póstur</button></div>'; return; }
  const send = minarBeidnir().find(b => b.tegund === 'senda' && b.drag === d.id);
  $('#postur').innerHTML = `<div><button class="s-takki" onclick="location.hash='#postur'">‹ Póstur</button></div>
  <section class="s-kafli"><div class="s-meta">Til: ${esc(d.til)}${d.afrit ? '<br>Afrit: ' + esc(d.afrit) : ''}</div>
   <div class="s-titill" style="margin-top:6px">${esc(d.efni)}</div>
   <div class="s-texti">${esc(d.texti)}</div>
   ${d.vidhengi.length ? `<div class="s-meta">Viðhengi: ${d.vidhengi.map(esc).join(', ')}</div>` : ''}
   <details><summary>Allur pósturinn með þræðinum</summary><div class="s-texti" style="font-size:14px">${esc(d.allt)}</div></details>
   <div id="sending">${send ? `<p>${beidniMerki(send.id)} ${esc((BEIDNIR[send.id] || {}).skilabod || 'Tölvan sendir þegar hún kemur upp.')}</p>` : ''}</div>
   <div class="s-takkar">${send && (!BEIDNIR[send.id] || BEIDNIR[send.id].stada !== 'villa') ? '' :
     (stadfesta ? `<button class="s-takki haetta" onclick="senda()">Já, senda til ${esc(d.til.split(';')[0])}${d.til.includes(';') ? ' o.fl.' : ''}</button>
       <button class="s-takki" onclick="teiknaDrag(false)">Hætta við</button>`
     : `<button class="s-takki adal" onclick="teiknaDrag(true)">Senda</button>`)}</div>
   <p class="s-meta">Tölvan sendir póstinn úr Outlook með undirskriftinni þinni og staðfestir hann í Sent Items. Sé slökkt á báðum tölvum bíður hann. Viltu breyta texta? Gerðu það í Outlook á tölvunni.</p>
  </section>`;
}
async function senda() {
  const d = DRAG;
  $('#sending').innerHTML = '<p>Setur í biðröð …</p>';
  try {
    const b = await setjaIBidrod({ tegund: 'senda', drag: d.id, breytt: d.breytt, lysing: d.efni + ' → ' + d.til });
    geymaBeidni({ id: b.id, tegund: 'senda', kl: b.kl, lysing: b.lysing, drag: d.id });
  } catch (e) { $('#sending').innerHTML = `<p class="s-villa">${esc(e.message)} Ekkert fór í biðröð.</p>`; return; }
  teiknaDrag(); fylgja();
}
async function fylgja() {
  for (let i = 0; i < 30; i++) {
    await bida(10000);
    if (!location.hash.startsWith('#postur/')) return;
    try { await saekjaStodu(); } catch (e) { continue; }
    const send = minarBeidnir().find(b => b.tegund === 'senda' && DRAG && b.drag === DRAG.id);
    teiknaDrag();
    if (send && BEIDNIR[send.id]) return;
  }
}

// ---------------------------------------------------------------- árekstrar
async function athuga() {
  const nafn = $('#nafn').value.trim(); if (!nafn) return;
  try { await setjaIBidrod({ tegund: 'arekstur', nafn, lysing: nafn }); $('#nafn').value = ''; }
  catch (e) { $('#ar-listi').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; return; }
  synaAthuga(); fylgjaArekstri();
}
async function fylgjaArekstri() {
  for (let i = 0; i < 60; i++) {
    await bida(10000);
    if (location.hash !== '#athuga') return;
    const opnar = minarBeidnir().filter(b => b.tegund === 'arekstur' && !NID[b.id]);
    if (!opnar.length) return;
    await synaAthuga();
  }
}
const NID = {};
async function synaAthuga() {
  const l = minarBeidnir().filter(b => b.tegund === 'arekstur').slice(0, 8);
  await Promise.all(l.filter(b => !NID[b.id]).map(async b => { try { const n = await lesa('nidurstodur/' + b.id + '.json'); if (n) NID[b.id] = n; } catch (e) {} }));
  $('#ar-listi').innerHTML = l.map(b => {
    const n = NID[b.id];
    if (!n) return `<section class="s-kafli"><div class="s-titill">${esc(b.lysing)}</div><p class="s-meta"><span class="s-merki bidur">Bíður tölvu</span> Sett í biðröð ${esc(dags(b.kl))} kl. ${esc(klukka(b.kl))}.</p></section>`;
    const st = n.stadbundid;
    return `<section class="s-kafli"><div class="s-titill">${esc(b.lysing)}</div>
      ${n.svar ? `<div class="s-texti">${esc(n.svar)}</div>` : ''}${n.skilabod ? `<p class="s-villa">${esc(n.skilabod)}</p>` : ''}
      ${st ? `<details><summary>Tölvan: póstur og málsmöppur</summary>
        ${st.postur.map(p => `<div class="s-lina"><div class="s-undir">${esc(p.postholf)} · ${esc(p.mappa)} · ${p.fjoldi}${p.takmarkad ? '+' : ''} póstar</div><div class="s-meta">${esc(p.fyrst)} – ${esc(p.sidast)}</div></div>`).join('') || '<p class="s-tomt">Enginn póstur.</p>'}
        ${st.moppur.map(m => `<div class="s-lina"><div class="s-undir">${esc(m.mappa)}</div><div class="s-meta">${m.i_heiti ? 'nafnið í heiti möppu' : ''}${m.skrar ? ' · ' + m.skrar + ' skjöl með nafninu' : ''}</div></div>`).join('') || '<p class="s-tomt">Engin málsmappa.</p>'}
      </details>` : ''}<p class="s-meta">Kannað á ${esc(n.tolva)} ${esc(dags(n.kl))} kl. ${esc(klukka(n.kl))}</p></section>`;
  }).join('');
}

// ---------------------------------------------------------------- minnisatriði
async function synaMinni() {
  if (!STADA) { try { await saekjaStodu(); } catch (e) {} }
  malSia();
  $('#upptaka').hidden = !(window.isSecureContext && navigator.mediaDevices && window.MediaRecorder);
  try {
    const l = (await listi('minni')).filter(f => !f.name.endsWith('.json')).sort((a, b) => b.name.localeCompare(a.name)).slice(0, 12);
    $('#minni-listi').innerHTML = l.map(f => `<button class="s-spjald" style="margin-bottom:8px" onclick="lesaMinni(this, ${JSON.stringify(f.name).replace(/"/g, '&quot;')})"><div class="s-undir">${esc(f.name)}</div></button>`).join('') || '<p class="s-tomt">Engin minnisatriði enn.</p>';
  } catch (e) { $('#minni-listi').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; }
}
async function lesaMinni(el, nafn) {
  if (!/\.md$/.test(nafn)) { el.insertAdjacentHTML('beforeend', '<div class="s-meta">Upptaka — umritunin birtist sem „… - UMRITUN.md“ þegar tölva hefur unnið hana.</div>'); return; }
  const t = await lesaTexta('minni/' + nafn);
  el.insertAdjacentHTML('beforeend', `<div class="s-texti" style="font-size:14px">${esc(t || '(tómt)')}</div>`);
}
function malSia() {
  const q = ($('#mal-leit').value || '').toLowerCase(); const s = $('#mal'); const v = s.value;
  const mal = (STADA && STADA.mal) || [];
  s.innerHTML = '<option value="">Ekkert mál valið</option>' + mal.filter(m => !q || m.birting.toLowerCase().includes(q)).slice(0, 60)
    .map(m => `<option value="${esc(m.mappa)}">${esc(m.birting)}</option>`).join('');
  if (q && s.options.length > 1) s.selectedIndex = 1; else s.value = v;
}
const stimpill = () => new Date().toISOString().slice(0, 19).replace('T', ' ').replace(/:/g, '');
const malHeiti = () => { const o = $('#mal').selectedOptions[0]; return $('#mal').value ? o.textContent : 'án máls'; };
const hreint = t => t.replace(/[\\/:*?"<>|#%]+/g, '_').slice(0, 60);
async function vistaMinni() {
  const t = $('#minni-texti').value.trim(); if (!t) return;
  const nafn = stimpill() + ' - ' + hreint(malHeiti()) + ' - MINNI.md';
  const n = new Date();
  const efni = '# Minnisatriði ' + n.getDate() + '.' + (n.getMonth() + 1) + '.' + n.getFullYear() + ' kl. ' +
    String(n.getHours()).padStart(2, '0') + ':' + String(n.getMinutes()).padStart(2, '0') + '\n\nMál: ' + malHeiti() + '\nSkráð í síma.\n\n' + t + '\n';
  try { await skrifa('minni/' + nafn, efni, 'text/markdown; charset=utf-8'); $('#minni-texti').value = ''; $('#minni-stada').textContent = 'Vistað í OneDrive.'; synaMinni(); }
  catch (e) { $('#minni-stada').textContent = e.message + ' Textinn er enn í reitnum.'; }
}
let REC = null, BUTAR = [];
async function upptaka() {
  if (REC) { REC.stop(); return; }
  try {
    const s = await navigator.mediaDevices.getUserMedia({ audio: true });
    const teg = MediaRecorder.isTypeSupported('audio/mp4') ? 'audio/mp4' : 'audio/webm';
    REC = new MediaRecorder(s, { mimeType: teg, audioBitsPerSecond: 64000 }); BUTAR = [];
    REC.ondataavailable = e => BUTAR.push(e.data);
    REC.onstop = async () => {
      s.getTracks().forEach(t => t.stop()); REC = null; $('#upptaka').textContent = '● Taka upp';
      const b = new Blob(BUTAR, { type: teg });
      const nafn = stimpill() + ' - ' + hreint(malHeiti()) + ' - upptaka' + (teg === 'audio/mp4' ? '.m4a' : '.webm');
      $('#minni-stada').textContent = 'Vistar upptöku …';
      try { await skrifa('minni/' + nafn, b, teg); await setjaIBidrod({ tegund: 'umrita', skra: nafn, lysing: nafn });
        $('#minni-stada').textContent = 'Upptakan er í OneDrive. Tölva umritar hana þegar hún kemur upp.'; synaMinni(); }
      catch (e) { $('#minni-stada').textContent = e.message; }
    };
    REC.start(); $('#upptaka').textContent = '■ Ljúka upptöku'; $('#minni-stada').textContent = 'Tekur upp … (hámark um 8 mín.)';
  } catch (e) { $('#minni-stada').textContent = 'Hljóðneminn fékkst ekki: ' + e.message; }
}

// ---------------------------------------------------------------- ræsing
(async () => {
  if ('serviceWorker' in navigator && window.isSecureContext) navigator.serviceWorker.register('sw.js').catch(() => {});
  const inni = await raesaMsal();
  if (!inni) { document.querySelectorAll('.s-syn').forEach(x => x.classList.toggle('virk', x.id === 'syn-inn')); return; }
  $('#stika').hidden = false;
  try { await saekjaStodu(); } catch (e) {}
  syna();
})();
