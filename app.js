// ESJA Legal — símaforrit. Les stöðu tölvanna og skrifar beiðnir í OneDrive notandans
// (Microsoft Graph, innskráning notandans). Ekkert efni er í kóðanum og ekkert er geymt
// í símanum nema innskráningin og listi yfir eigin beiðnir (auðkenni).
'use strict';
const UTGAFA = 8;
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
  let ss = [], ls = [];
  try { ss = Object.keys(sessionStorage).filter(k => k.startsWith('msal')).map(k => k.split('.').slice(-2).join('.')); } catch (e) {}
  try { ls = Object.keys(localStorage).filter(k => k.startsWith('msal')).map(k => k.split('.').slice(-2).join('.')); } catch (e) {}
  return 'Útgáfa: ' + UTGAFA + '\nSvar frá Microsoft: ' + (KOMA.svar ? 'já' : 'nei') + '\nSlóð við komu: ' + KOMA.slod +
    '\nReikningar í minni: ' + (MSAL ? MSAL.getAllAccounts().length : '-') + '\nsessionStorage: ' + (ss.join(', ') || 'tómt') +
    '\nlocalStorage: ' + (ls.slice(0, 8).join(', ') || 'tómt') +
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
    // temporaryCacheLocation: iPhone opnar endurkomuna frá Microsoft stundum í nýjum flipa (t.d. um
    // Authenticator), og þá er sessionStorage tómt; localStorage lifir það af (greining 22.10.2026).
    cache: { cacheLocation: 'localStorage', temporaryCacheLocation: 'localStorage', storeAuthStateInCookie: true },
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
async function innskraGluggi() {          // varaleið: Microsoft í nýjum glugga, síðan les svarið þaðan
  if (!MSAL) return;
  try {
    const r = await MSAL.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
    if (r && r.account) { REIKN = r.account; MSAL.setActiveAccount(r.account); location.replace(location.pathname + '#idag'); location.reload(); }
  } catch (e) {
    $('#inn-villa').textContent = 'Innskráning í glugga mistókst: ' + (e.errorCode || '') + ' ' + e.message;
    const g = $('#greining'); g.textContent = 'Greining (sendu Claude):\n' + greining(); g.hidden = false;
  }
}
async function innskra() {
  if (!MSAL) return;
  try { localStorage.setItem('esja_innskra_reynt', new Date().toISOString()); } catch (e) {}
  try { await MSAL.loginRedirect({ scopes: SCOPES, prompt: 'select_account', redirectStartPage: location.pathname + '#idag' }); }
  catch (e) {
    if (e.errorCode === 'interaction_in_progress') {   // leif eftir fyrri tilraun
      try { [sessionStorage, localStorage].forEach(s => Object.keys(s).filter(k => k.includes('interaction.status')).forEach(k => s.removeItem(k))); } catch (x) {}
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
const BUTUR = 16 * 327680;   // 5 MiB, margfeldi af 320 KiB eins og Graph krefst
async function skrifaStort(p, blob, framvinda) {    // upphleðslulota (createUploadSession) fyrir > 4 MB
  const r = await graph(skraSlod(p) + ':/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename' } }) });
  if (!r.ok) throw new Error('Upphleðsla hófst ekki (' + r.status + ').');
  const url = (await r.json()).uploadUrl;
  for (let a = 0; a < blob.size; a += BUTUR) {
    const b = Math.min(a + BUTUR, blob.size);
    let s = null;
    for (let tilraun = 0; tilraun < 3; tilraun++) {
      try { s = await fetch(url, { method: 'PUT', headers: { 'Content-Range': 'bytes ' + a + '-' + (b - 1) + '/' + blob.size }, body: blob.slice(a, b) }); }
      catch (e) { s = null; }
      if (s && (s.ok || s.status === 202)) break;
      await bida(2000 * (tilraun + 1));
    }
    if (!s || !(s.ok || s.status === 202)) throw new Error('Upphleðsla rofnaði við ' + Math.round(a / 1048576) + ' MB.');
    if (framvinda) framvinda(b / blob.size);
  }
  return true;
}
async function skrifa(p, efni, teg = 'application/json', framvinda) {
  if (efni instanceof Blob && efni.size > 4 * 1024 * 1024) return skrifaStort(p, efni, framvinda);
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
let STADA = null, ALLAR = {}, BEIDNIR = {}, DROG = [], VERK = null, VERK_TOLVA = '';
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
  // Drög og verkefnalisti koma úr þeirri stöðu sem hefur nýjustu útgáfu þeirra (G252):
  // tölva sem snerti ekki Outlook í umferðinni ber eldri drög en nýjasta stöðu.
  const nyjast = (k, kl) => radad.filter(a => a[k]).sort((a, b) => (kl(b) || '').localeCompare(kl(a) || ''))[0];
  const d = nyjast('drog', a => a.drog_kl || a.kl);
  DROG = d ? d.drog : [];
  const v = nyjast('verkefni', a => a.verkefni.kl);
  VERK = v ? v.verkefni : null; VERK_TOLVA = v ? v.tolva : '';
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
const SYNIR = ['idag', 'postur', 'frestir', 'verkefni', 'athuga', 'minni'];
function syna() {
  const h = (location.hash || '#idag').slice(1).split('/')[0];
  const s = SYNIR.includes(h) ? h : 'idag';
  document.querySelectorAll('.s-syn').forEach(x => x.classList.toggle('virk', x.id === 'syn-' + s));
  document.querySelectorAll('.s-stika a').forEach(a => a.classList.toggle('virkt', a.dataset.s === s));
  fabStada(s);
  if (s === 'athuga') document.querySelector('.s-stika a[data-s=idag]').classList.add('virkt');
  ({ idag: synaDag, verkefni: synaVerk, postur: synaPost, frestir: synaFresti, athuga: synaAthuga, minni: synaMinni })[s]();
}
window.addEventListener('hashchange', syna);

function lina(x) {
  if (x.teg === 'thinghald') return `<div class="s-lina hreimur"><div class="s-tg th">${esc(x.timi || '')} ${esc(x.tegund)}</div>
    <div class="s-titill">${esc(x.mal)}</div><div class="s-meta">${esc(x.domstoll)}${x.salur ? ' · salur ' + esc(x.salur) : ''}${x.lengd ? ' · ' + esc(x.lengd) : ''}</div></div>`;
  return `<div class="s-lina ${x.logbundinn ? 'brynt' : 'vidv'}"><div class="s-tg ${x.logbundinn ? 'lb' : ''}">${esc(x.flokkur)}${x.dags < iDag() ? ' · ' + esc(dags(x.dags)) : ''}</div>
    <div class="s-titill">${esc(x.heiti)}</div>
    <div class="s-meta">${esc(x.mal)}${x.i_dagatali ? '' : ' <span class="s-merki brynt">ekki í dagatali</span>'}</div></div>`;
}
const fram = n => { const d = new Date(Date.now() + n * 864e5); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const tilSunnudags = () => (7 - new Date().getDay()) % 7;
function dagaRuna(atr, hopar) {        // atriði flokkuð eftir dögum innan hópa
  const radad = atr.slice().sort((a, b) => (a.dags + (a.timi || '~')).localeCompare(b.dags + (b.timi || '~')));
  return hopar.map(h => {
    if (h.fra > h.til) return '';
    const x = radad.filter(a => a.dags >= h.fra && a.dags <= h.til);
    if (!x.length && !h.alltaf) return '';
    const dagar = [...new Set(x.map(a => a.dags))];
    return `<h3 class="s-hopur">${esc(h.h)}</h3>` + (dagar.length ? dagar.map(d => `<div class="s-dagur"><div class="s-dh">${esc(dags(d))}${d === iDag() ? ' · í dag' : d === aMorgun() ? ' · á morgun' : ''}</div>${x.filter(a => a.dags === d).map(lina).join('')}</div>`).join('') : '<p class="s-tomt">Ekkert.</p>');
  }).join('');
}
let SIA = '';
function siaSmellt(s) {
  if (s === 'postur') { location.hash = '#postur'; return; }
  if (s === 'vl') { location.hash = '#verkefni'; return; }
  SIA = SIA === s ? '' : s; synaDag();
}
async function synaDag(nytt) {
  if (nytt || !STADA) { $('#dagur').innerHTML = '<p class="s-tomt">Les …</p>'; try { await saekjaStodu(); } catch (e) { $('#dagur').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; return; } }
  const g = STADA && STADA.dagur;
  if (!g) { $('#dagur').innerHTML = ferskleiki(); return; }
  const a = iDag(), b = aMorgun(), f14 = fram(14), sun = tilSunnudags();
  const fr = g.frestir.filter(x => x.dags >= a && x.dags <= f14), th = g.thinghold.filter(x => x.dags >= a && x.dags <= f14);
  const lidid = g.frestir.filter(x => x.dags < a);
  const nTh = th.slice().sort((x, y) => (x.dags + (x.timi || '')).localeCompare(y.dags + (y.timi || '')))[0];
  const bid = hreinsaVerkBid(), opin = VERK && VERK.lidir ? VERK.lidir.filter(l => !verkLokid(l, bid)).length : null;
  const tv = (STADA.tilvonandi && STADA.tilvonandi.bidur) || [];
  $('#t-postur').textContent = (DROG || []).length || '';
  if (opin != null) $('#t-verkefni').textContent = opin || '';
  const reitur = (s, n, t, u, r) => `<button class="s-reitur${SIA === s ? ' virk' : ''}${r ? ' r' : ''}" onclick="siaSmellt('${s}')"><span class="n">${n}</span><span class="t">${t}</span><span class="u">${esc(u)}</span></button>`;
  $('#undirhaus').textContent = 'Næstu 14 dagar';
  let runa;
  if (SIA === 'tv') runa = tilvonandiHTML();
  else {
    const atr = SIA === 'fr' ? fr : SIA === 'th' ? th : [...th, ...fr];
    runa = dagaRuna(atr, [{ h: 'Í dag og á morgun', fra: a, til: b, alltaf: true },
      { h: 'Það sem eftir er vikunnar', fra: fram(2), til: fram(sun) },
      { h: 'Næsta vika og fram að 14 dögum', fra: fram(Math.max(sun + 1, 2)), til: f14 }]);
  }
  $('#dagur').innerHTML = ferskleiki() + (g.villur.length ? `<p class="s-villa">${g.villur.map(esc).join('<br>')}</p>` : '') +
    `<div class="s-reitir">${reitur('fr', fr.length, 'Frestir · 14 d', fr.filter(x => !x.i_dagatali).length + ' ekki í dagatali', fr.some(x => x.logbundinn && x.dags <= b))}
     ${reitur('th', th.length, 'Þinghöld', nTh ? 'næst ' + dags(nTh.dags) + (nTh.timi ? ' kl. ' + nTh.timi : '') : 'ekkert')}
     ${reitur('postur', (DROG || []).length, 'Póstur', 'bíður svars')}
     ${reitur('vl', opin == null ? '…' : opin, 'Verkefni', 'opin á lista')}
     ${reitur('tv', tv.length, 'Tilvonandi', 'upptaka bíður máls')}</div>
     ${SIA ? `<p class="s-meta">Sía: ${({ fr: 'aðeins frestir', th: 'aðeins þinghöld', tv: 'tilvonandi viðskiptavinir' })[SIA]} · <a href="#idag" onclick="SIA='';synaDag();return false">sýna allt</a></p>` : ''}
     <section class="s-kafli">${runa}</section>
     ${lidid.length && !SIA ? `<section class="s-kafli"><details><summary>Liðið og óuppfyllt · ${lidid.length}</summary>${lidid.map(lina).join('')}</details></section>` : ''}
     <div class="s-takkar"><button class="s-takki" onclick="synaDag(true)">Uppfæra</button><button class="s-takki" onclick="location.hash='#athuga'">Athuga hagsmunaárekstur</button></div>`;
}
async function synaFresti(nytt) {
  if (nytt || !STADA) { $('#frestir').innerHTML = '<p class="s-tomt">Les …</p>'; try { await saekjaStodu(); } catch (e) { $('#frestir').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; return; } }
  const g = STADA && STADA.dagur;
  if (!g) { $('#frestir').innerHTML = ferskleiki(); return; }
  const a = iDag(), allt = [...g.thinghold, ...g.frestir];
  $('#undirhaus').textContent = g.frestir.length + ' frestir · ' + g.thinghold.length + ' þinghöld';
  $('#frestir').innerHTML = ferskleiki() + `<section class="s-kafli">${dagaRuna(allt, [{ h: 'Liðið og óuppfyllt', fra: '0000', til: fram(-1) },
    { h: 'Í dag og á morgun', fra: a, til: aMorgun(), alltaf: true }, { h: 'Næstu 14 dagar', fra: fram(2), til: fram(14) }, { h: 'Síðar', fra: fram(15), til: '9999' }])}</section>
    <div class="s-takkar"><button class="s-takki" onclick="synaFresti(true)">Uppfæra</button></div>`;
}

// ---------------------------------------------------------------- tilvonandi viðskiptavinir (Ómar 7.11.2026)
// Upptaka sem bíður máls fer í málsmöppu með smelli: tölvan afritar hana í rót málsins og
// umritunina í Claude vinnuskjöl\Lestur. Tveir smellir (staðfesting í hnappnum) áður en nokkuð fer.
let TV_STADFESTA = '';
function tilvonandiHTML() {
  const t = (STADA && STADA.tilvonandi) || { bidur: [], komid: [] };
  const min = minarBeidnir().filter(b => b.tegund === 'tilvonandi');
  const kort = x => {
    const send = min.find(b => b.skra === x.skra && (!BEIDNIR[b.id] || BEIDNIR[b.id].stada !== 'villa'));
    const villa = min.filter(b => b.skra === x.skra && BEIDNIR[b.id] && BEIDNIR[b.id].stada === 'villa').pop();
    const takki = m => { const k = x.skra + '|' + m, st = TV_STADFESTA === k;
      return `<button class="s-takki ${st ? 'haetta' : 'adal'}" data-k="${esc(k)}" onclick="tvSmellt(this.dataset.k)">${st ? 'Staðfesta: setja í ' : 'Setja í málið: '}${esc(m)}</button>`; };
    const valid = TV_STADFESTA.startsWith(x.skra + '|') && !(x.uppastungur || []).includes(TV_STADFESTA.slice(x.skra.length + 1)) ? TV_STADFESTA.slice(x.skra.length + 1) : '';
    return `<div class="s-lina hreimur"><div class="s-titill">${esc(x.nafn)}</div>
      <div class="s-meta">${esc(dags(x.dags))} kl. ${esc(x.kl)} · ${x.staerd_mb} MB · ${x.umritun ? 'umritað' : 'umritun bíður'}</div>
      ${send ? `<p class="s-meta">${beidniMerki(send.id)} ${esc((BEIDNIR[send.id] || {}).skilabod || 'Tölvan setur upptökuna í málið þegar hún kemur upp.')}</p>` : `
      ${villa ? `<p class="s-villa">${esc(BEIDNIR[villa.id].skilabod || 'Ekki framkvæmt.')}</p>` : ''}
      <div class="s-takkar">${(x.uppastungur || []).map(takki).join('') || '<span class="s-meta">Engin málsmappa ber nafnið enn. Stofnaðu málið í Vergo, eða veldu mál hér:</span>'}${valid ? takki(valid) : ''}</div>
      <select class="s-inn" style="margin-top:8px" data-skra="${esc(x.skra)}" onchange="if(this.value)tvSmellt(this.dataset.skra+'|'+this.value)"><option value="">Annað mál …</option>
        ${((STADA && STADA.mal) || []).map(m => `<option value="${esc(m.mappa)}">${esc(m.birting)}</option>`).join('')}</select>`}</div>`;
  };
  return `<h3 class="s-hopur">Upptökur sem bíða máls · ${t.bidur.length}</h3>` + (t.bidur.length ? t.bidur.map(kort).join('') : '<p class="s-tomt">Engin upptaka bíður máls. Tilvonandi viðskiptavinur er tekinn upp undir Minni.</p>') +
    (t.komid.length ? `<details><summary>Komið í mál síðustu 14 daga · ${t.komid.length}</summary>${t.komid.map(x => `<div class="s-lina"><div class="s-undir">${esc(x.nafn)} → ${esc(x.mappa)}</div><div class="s-meta">${esc(dags(x.dags))}</div></div>`).join('')}</details>` : '');
}
async function tvSmellt(k) {
  if (TV_STADFESTA !== k) { TV_STADFESTA = k; synaDag(); return; }
  TV_STADFESTA = '';
  const skra = k.slice(0, k.indexOf('|')), mappa = k.slice(k.indexOf('|') + 1);
  try {
    const b = await setjaIBidrod({ tegund: 'tilvonandi', skra, mappa, lysing: skra + ' → ' + mappa });
    geymaBeidni({ id: b.id, tegund: 'tilvonandi', kl: b.kl, lysing: b.lysing, skra });
  } catch (e) { tilkynna(e.message + ' Ekkert fór í biðröð.'); }
  synaDag();
}
function tilkynna(t) { const d = $('#tilk'); d.textContent = t; d.hidden = false; setTimeout(() => { d.hidden = true; }, 6000); }

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
  const drog = DROG || [];
  $('#t-postur').textContent = drog.length || '';
  const sendar = minarBeidnir().filter(b => b.tegund === 'senda');
  if (h[1]) { DRAG = drog.find(d => d.id === decodeURIComponent(h[1])); return teiknaDrag(); }
  $('#postur').innerHTML = ferskleiki() + `<section class="s-kafli"><h2>Tillögur Claude · ${drog.length}</h2>
    ${drog.length ? drog.map((x, i) => `<button class="s-spjald" onclick="location.hash='#postur/'+encodeURIComponent(DROG[${i}].id)">
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

// ---------------------------------------------------------------- verkefnalistinn (G252)
// Listinn kemur úr stöðu tölvunnar (verkefni). Breyting fer í biðröðina og tölvan setur hana
// á verkefnalistann; þangað til man síminn hana sjálfur (esja_verk) svo verkið hverfi strax.
let SYNA_LOKID = false;
function verkBid() { try { return JSON.parse(localStorage.getItem('esja_verk') || '{}'); } catch (e) { return {}; } }
function geymaVerkBid(o) { try { localStorage.setItem('esja_verk', JSON.stringify(o)); } catch (e) {} }
const aldurMs = b => Date.now() - new Date(b.kl).getTime();
function hreinsaVerkBid() {     // það sem tölvan hefur sett á listann gleymist; villur sjást í sólarhring
  const o = verkBid(), ny = {};
  Object.entries(o).forEach(([lid, l]) => {
    const halda = l.filter(b => BEIDNIR[b.id] ? (BEIDNIR[b.id].stada === 'villa' && aldurMs(b) < 864e5) : aldurMs(b) < 2 * 864e5);
    if (halda.length) ny[lid] = halda;
  });
  geymaVerkBid(ny); return ny;
}
const bidurTolvu = b => !BEIDNIR[b.id];
function verkLokid(l, bid) {
  const b = (bid[l.id] || []).filter(x => bidurTolvu(x) && ['afgreitt', 'opna', 'framkvaemdu'].includes(x.adgerd)).pop();
  return b ? b.adgerd === 'afgreitt' : l.lokid;
}
function svarHTML(v) {
  const k = { gert: 'lokid', ekki: 'brynt', seint: 'brynt', bidur: 'bidur' }[v.k] || '';
  return `<div class="s-svar"><div class="s-meta">${esc(v.dags)}${v.framkvaema ? ' · Framkvæmdu' : ''}</div>
    <div class="s-undir">${esc(v.texti)}</div>
    ${v.stada ? `<div><span class="s-merki ${k}">${esc(v.stada)}</span></div>` : ''}
    ${v.skyr ? `<div class="s-meta">${v.k === 'ekki' ? 'Ástæða' : 'Gert'}: ${esc(v.skyr)}</div>` : ''}</div>`;
}
function verkHTML(l, bid) {
  const min = bid[l.id] || [];
  const lokid = verkLokid(l, bid);
  const bidur = min.filter(bidurTolvu);
  const villa = min.filter(b => !bidurTolvu(b) && BEIDNIR[b.id].stada === 'villa');
  const nySvor = bidur.filter(b => b.adgerd === 'svar' || b.adgerd === 'framkvaemdu')
    .map(b => svarHTML({ dags: dags(b.kl) + ' kl. ' + klukka(b.kl), texti: b.texti, framkvaema: b.adgerd === 'framkvaemdu', k: 'bidur', stada: 'Í símanum – bíður tölvu' }));
  const seint = l.svor.some(v => v.k === 'seint');
  return `<section class="s-kafli s-verk${seint ? ' seint' : ''}" data-lid="${esc(l.id)}">
    <div class="s-titill">${esc(l.texti)}${l.mal ? ` <span class="s-merki">${esc(l.mal)}</span>` : ''}</div>
    ${l.svor.map(svarHTML).join('')}${nySvor.join('')}
    ${villa.map(b => `<p class="s-villa">${esc(BEIDNIR[b.id].skilabod || 'Ekki framkvæmt.')}</p>`).join('')}
    ${bidur.some(b => b.adgerd === 'afgreitt' || b.adgerd === 'opna') ? '<p class="s-meta"><span class="s-merki bidur">Bíður tölvu</span></p>' : ''}
    <details class="s-verk-svar"><summary>Svara</summary>
      <textarea class="s-inn" placeholder="Svar eða fyrirmæli til Claude"></textarea>
      <div class="s-takkar"><button class="s-takki" onclick="verkSvar(this,'svar')">Senda svar</button>
        <button class="s-takki adal" onclick="verkSvar(this,'framkvaemdu')">Framkvæmdu</button></div></details>
    <div class="s-takkar">${lokid
      ? `<button class="s-takki" onclick="verkMerkja(this,'opna')">Aftur á listann</button>`
      : `<button class="s-takki adal" onclick="verkMerkja(this,'afgreitt')">✓ Afgreitt</button>`}</div>
  </section>`;
}
async function synaVerk(nytt) {
  if (nytt || !STADA) { $('#verkefni').innerHTML = '<p class="s-tomt">Les …</p>'; try { await saekjaStodu(); } catch (e) { $('#verkefni').innerHTML = `<p class="s-villa">${esc(e.message)}</p>`; return; } }
  const bid = hreinsaVerkBid();
  if (!VERK || !VERK.lidir) { $('#verkefni').innerHTML = ferskleiki() + '<p class="s-tomt">Verkefnalistinn hefur ekki borist frá tölvunum enn.</p>'; return; }
  const opin = VERK.lidir.filter(l => !verkLokid(l, bid)), lokin = VERK.lidir.filter(l => verkLokid(l, bid));
  const seintN = opin.reduce((n, l) => n + l.svor.filter(v => v.k === 'seint').length, 0);
  $('#t-verkefni').textContent = opin.length || '';
  $('#undirhaus').textContent = opin.length + ' opin verk';
  const kaflar = [];
  opin.forEach(l => { let k = kaflar.find(x => x.h === l.kafli); if (!k) kaflar.push(k = { h: l.kafli, l: [] }); k.l.push(l); });
  const aldur = (Date.now() - new Date(VERK.kl).getTime()) / 6e4;
  $('#verkefni').innerHTML = `<p class="s-ferskt ${aldur > 30 ? 'gamalt' : ''}">Listi ${esc(dags(VERK.listi))} · frá ${esc(VERK_TOLVA)} ${esc(dags(VERK.kl))} kl. ${esc(klukka(VERK.kl))}${aldur > 30 ? ' — tölvurnar gætu verið slökktar; breytingar bíða' : ''}</p>` +
    (seintN ? `<p class="s-villa">⚠ ${seintN} ${seintN === 1 ? 'svar bíður' : 'svör bíða'} Claude lengur en 2 klst.</p>` : '') +
    (kaflar.length ? kaflar.map(k => `<h2 class="s-verk-kafli">${esc(k.h)}</h2>` + k.l.map(l => verkHTML(l, bid)).join('')).join('')
      : '<p class="s-tomt">Ekkert opið á listanum.</p>') +
    (lokin.length ? `<div class="s-takkar"><button class="s-takki" onclick="SYNA_LOKID=!SYNA_LOKID;synaVerk()">${SYNA_LOKID ? 'Fela afgreidd' : 'Afgreitt'} (${lokin.length})</button></div>` +
      (SYNA_LOKID ? '<p class="s-meta">Afgreidd verk. „Aftur á listann“ afturkallar afgreiðslu.</p>' + lokin.map(l => verkHTML(l, bid)).join('') : '') : '') +
    `<div class="s-takkar"><button class="s-takki" onclick="synaVerk(true)">Uppfæra</button></div>`;
}
async function verkBeidni(lid, adgerd, texti) {
  const b = { tegund: 'verkefni', lid, adgerd, lysing: adgerd + ' ' + lid };
  if (texti) { b.texti = texti; b.svar_ms = Date.now(); }
  await setjaIBidrod(b);
  const o = verkBid(); (o[lid] = o[lid] || []).push({ id: b.id, kl: b.kl, adgerd, texti: texti || '' }); geymaVerkBid(o);
  synaVerk(); fylgjaVerki();
}
async function verkMerkja(el, adgerd) {
  const lid = el.closest('.s-verk').dataset.lid; el.disabled = true;
  try { await verkBeidni(lid, adgerd); }
  catch (e) { el.disabled = false; el.insertAdjacentHTML('afterend', `<p class="s-villa">${esc(e.message)} Ekkert fór í biðröð.</p>`); }
}
async function verkSvar(el, adgerd) {
  const s = el.closest('.s-verk'), ta = s.querySelector('textarea'), t = ta.value.trim();
  if (!t) { ta.focus(); return; }
  el.disabled = true;
  try { await verkBeidni(s.dataset.lid, adgerd, t); }
  catch (e) { el.disabled = false; el.insertAdjacentHTML('afterend', `<p class="s-villa">${esc(e.message)} Textinn er enn í reitnum.</p>`); }
}
let FYLGIR_VERKI = false;
async function fylgjaVerki() {          // fylgist með þar til tölvan hefur sett breytingarnar á listann
  if (FYLGIR_VERKI) return; FYLGIR_VERKI = true;
  try {
    for (let i = 0; i < 40; i++) {
      await bida(15000);
      if (location.hash !== '#verkefni') return;
      if (document.activeElement && document.activeElement.tagName === 'TEXTAREA') continue;
      try { await saekjaStodu(); } catch (e) { continue; }
      synaVerk();
      if (!Object.values(hreinsaVerkBid()).some(l => l.some(bidurTolvu))) return;
    }
  } finally { FYLGIR_VERKI = false; }
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
const tvNafn = () => ($('#tv-nafn').value || '').trim();
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
let REC = null, BUTAR = [], VAKA = null, BIDUR_UPPHL = null, KLUKKA = null;
// Ekkert hámark á lengd (Ómar 7.11.2026). Bútar á sekúndu svo ekkert tapist; skjárinn haldinn vakandi.
// Ómar tilkynnir viðmælanda sjálfur að upptaka sé hafin.
async function upptaka() {
  if (REC) { REC.stop(); return; }
  if (BIDUR_UPPHL) { $('#minni-stada').textContent = 'Fyrri upptaka er enn óvistuð. Ýttu á „Reyna aftur“ fyrst.'; return; }
  try {
    const s = await navigator.mediaDevices.getUserMedia({ audio: true });
    const teg = MediaRecorder.isTypeSupported('audio/mp4') ? 'audio/mp4' : 'audio/webm';
    const tv = tvNafn();
    const nafn = stimpill() + (tv ? ' - BÍÐUR MÁLS - ' + hreint(tv) : ' - ' + hreint(malHeiti())) + ' - upptaka' + (teg === 'audio/mp4' ? '.m4a' : '.webm');
    REC = new MediaRecorder(s, { mimeType: teg, audioBitsPerSecond: 64000 }); BUTAR = [];
    const byrjun = Date.now();
    REC.ondataavailable = e => { if (e.data && e.data.size) BUTAR.push(e.data); };
    REC.onstop = async () => {
      s.getTracks().forEach(t => t.stop()); REC = null; clearInterval(KLUKKA); $('#upptaka').textContent = '● Taka upp'; fabStada();
      try { if (VAKA) await VAKA.release(); } catch (e) {} VAKA = null;
      BIDUR_UPPHL = { blob: new Blob(BUTAR, { type: teg }), nafn, teg, tv }; BUTAR = [];
      if (tv) $('#tv-nafn').value = '';
      hladaUpp();
    };
    REC.start(1000);
    try { VAKA = await navigator.wakeLock.request('screen'); } catch (e) {}
    $('#upptaka').textContent = '■ Ljúka upptöku'; fabStada();
    $('#minni-stada').textContent = 'Tekur upp …';
    KLUKKA = setInterval(() => { const m = Math.floor((Date.now() - byrjun) / 1000);
      $('#minni-stada').textContent = 'Tekur upp' + (tv ? ' — ' + tv + ' (bíður máls)' : '') + ' · ' + Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0') + ' · haltu síðunni opinni'; }, 1000);
  } catch (e) { $('#minni-stada').textContent = 'Hljóðneminn fékkst ekki: ' + e.message; }
}
async function hladaUpp() {        // upptakan geymist í minni símans þar til hún er komin í OneDrive
  const u = BIDUR_UPPHL; if (!u) return;
  $('#upphl-aftur').hidden = true;
  const mb = (u.blob.size / 1048576).toFixed(1);
  $('#minni-stada').textContent = 'Vistar upptöku (' + mb + ' MB) …';
  try {
    await skrifa('minni/' + u.nafn, u.blob, u.teg, h => { $('#minni-stada').textContent = 'Vistar upptöku … ' + Math.round(h * 100) + '% af ' + mb + ' MB'; });
    BIDUR_UPPHL = null;
    await setjaIBidrod({ tegund: 'umrita', skra: u.nafn, lysing: u.nafn });
    $('#minni-stada').textContent = 'Upptakan er í OneDrive' + (u.tv ? ', merkt „Bíður máls“ (' + u.tv + ')' : '') + '. Tölva umritar hana þegar hún kemur upp.';
    synaMinni();
  } catch (e) {
    $('#minni-stada').textContent = e.message + ' Upptakan er enn í símanum — ekki loka síðunni.';
    $('#upphl-aftur').hidden = false;
  }
}
// Record-hnappurinn (Ómar 7.11.2026): einn smellur byrjar upptöku strax og opnar Minni, annar stöðvar.
// Meðan tekið er upp sést hann á öllum flipum, líka Minni, og púlsar með ferningi.
const upptakaHaegt = () => !!(window.isSecureContext && navigator.mediaDevices && window.MediaRecorder);
function fabStada(s) {
  s = s || (location.hash || '#idag').slice(1).split('/')[0];
  const f = $('#fab'), rec = !!REC;
  f.classList.toggle('rec', rec);
  f.hidden = !rec && s === 'minni';
  f.setAttribute('aria-label', rec ? 'Stöðva upptöku' : 'Byrja upptöku');
  f.querySelector('span').textContent = rec ? 'Stop' : 'Record';
}
function fabSmellt() {
  if (REC) { upptaka(); return; }
  if (location.hash !== '#minni') location.hash = '#minni';
  if (upptakaHaegt()) upptaka();
}
window.addEventListener('beforeunload', e => { if (REC || BIDUR_UPPHL) { e.preventDefault(); e.returnValue = ''; } });

// ---------------------------------------------------------------- ræsing
(async () => {
  if ('serviceWorker' in navigator && window.isSecureContext) navigator.serviceWorker.register('sw.js').catch(() => {});
  const inni = await raesaMsal();
  if (!inni) { document.querySelectorAll('.s-syn').forEach(x => x.classList.toggle('virk', x.id === 'syn-inn')); return; }
  $('#stika').hidden = false;
  try { await saekjaStodu(); } catch (e) {}
  syna();
})();
