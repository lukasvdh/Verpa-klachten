/* ═══════════════════════════════════════════════════
   Verpa Klachtenmeldingen – app.js
   Stack: MSAL.js + Microsoft Graph API + SharePoint Lists + SheetJS
   ═══════════════════════════════════════════════════ */

/* ── CONFIG ─────────────────────────────────────────
   Pas onderstaande waarden aan na aanmaken App Registration in Azure AD.
   Vereiste Graph permissies (Delegated):
     - Sites.ReadWrite.All  (of Sites.Manage.All)
     - User.Read
   SharePoint List: "KlachtenMeldingen" met onderstaande kolommen.
   ─────────────────────────────────────────────────── */
const CONFIG = {
  clientId:    'e1f6ac61-a64c-4a4b-a2f5-f061c989983f', // Verpa Klachten App
  tenantId:    'e65dbe4b-d1e2-4283-b0f5-aa7717e81077', // Verpa tenant
  redirectUri: window.location.origin,

  // SharePoint site – zet de correcte site-naam/subsite
  spSiteUrl:   'https://verpabenelux.sharepoint.com/sites/OfficeData',
  spListName:  'KlachtenMeldingen',

  // UPN's van beheerder(s) – kleine letters
  // Object-ID van de Entra-groep waarvan leden als beheerder gelden
  adminGroupId: 'JOUW_ENTRA_GROEP_OBJECT_ID',
};

/* ─────────────────────────────────────────────────── */

const msalConfig = {
  auth: {
    clientId:    CONFIG.clientId,
    authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
};

const GRAPH_SCOPES = ['User.Read', 'Sites.ReadWrite.All', 'GroupMember.Read.All'];

// MSAL v3 exporteert via window.msalBrowser, v2 via window.msal
const msalLib = window.msalBrowser || window.msal;

let msalInstance;
let currentUser  = null;
let allKlachten  = [];
let currentRejectId = null;

/* ══════════════════ INIT ══════════════════ */
async function init() {
  showLogin();

  // Koppel loginknop direct
  document.getElementById('btnLogin').addEventListener('click', () => {
    if (msalInstance) {
      msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
    } else {
      window.location.reload();
    }
  });

  if (!msalLib) {
    console.error('MSAL bibliotheek niet geladen.');
    showLogin();
    return;
  }

  try {
    msalInstance = new msalLib.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    const resp = await msalInstance.handleRedirectPromise().catch(err => {
      console.warn('handleRedirectPromise fout:', err);
      return null;
    });
    if (resp && resp.account) { await onSignedIn(resp.account); return; }

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      await onSignedIn(accounts[0]);
    }
  } catch (err) {
    console.error('MSAL init fout:', err);
    showLogin();
  }
}

function showLogin() {
  document.getElementById('loginScreen').classList.remove('hidden');
  document.getElementById('appShell').classList.add('hidden');
}

async function onSignedIn(account) {
  // Verberg login scherm zodra we een account hebben
  document.getElementById('loginScreen').classList.add('hidden');
  const token = await getToken(account);
  if (!token) { showLogin(); return; }

  // Haal gebruikersprofiel op
  const profile = await graphGet('/me', token);
  const email = (profile.mail || profile.userPrincipalName || '').toLowerCase();

  // Check groepslidmaatschap via Entra
  let isAdmin = false;
  try {
    const groupCheck = await fetch('https://graph.microsoft.com/v1.0/me/checkMemberOf', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ groupIds: [CONFIG.adminGroupId] }),
    });
    if (groupCheck.ok) {
      const groupData = await groupCheck.json();
      isAdmin = Array.isArray(groupData.value) && groupData.value.includes(CONFIG.adminGroupId);
    }
  } catch (e) {
    console.warn('Admin groepscheck mislukt:', e);
  }

  currentUser = {
    name:    profile.displayName || account.name || 'Gebruiker',
    email,
    isAdmin,
    token,
  };

  // UI bijwerken
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('appShell').classList.remove('hidden');
  document.getElementById('userName').textContent   = currentUser.name.split(' ')[0];
  document.getElementById('userRole').textContent   = currentUser.isAdmin ? 'Beheerder' : 'Melder';
  document.getElementById('dashSubtitle').textContent =
    currentUser.isAdmin ? 'Alle ingediende klachten' : 'Uw ingediende klachten';

  // Admin-elementen tonen
  if (currentUser.isAdmin) {
    document.querySelectorAll('.admin-only').forEach(el => el.classList.remove('hidden'));
  }

  setupNavigation();
  setupForm();
  setupImport();
  setupModals();

  await loadDashboard();
}

/* ══════════════════ MSAL HELPERS ══════════════════ */
async function getToken(account) {
  try {
    const resp = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return resp.accessToken;
  } catch {
    try {
      const resp = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      return resp.accessToken;
    } catch (e) {
      console.error('Token fout:', e);
      return null;
    }
  }
}

async function refreshToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  const tok = await getToken(accounts[0]);
  currentUser.token = tok;
  return tok;
}

/* ══════════════════ GRAPH / SHAREPOINT HELPERS ══════════════════ */
async function graphGet(path, token) {
  const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!r.ok) throw new Error(`Graph GET ${path}: ${r.status}`);
  return r.json();
}

let _siteId = null;
async function getSiteId() {
  if (_siteId) return _siteId;
  const tok = await refreshToken();
  const url  = new URL(CONFIG.spSiteUrl);
  const host = url.hostname;
  const path = url.pathname.replace(/^\//, '');
  const data = await graphGet(`/sites/${host}:/${path}`, tok);
  _siteId = data.id;
  return _siteId;
}

let _listId = null;
async function getListId() {
  if (_listId) return _listId;
  const siteId = await getSiteId();
  const tok    = await refreshToken();
  const data   = await graphGet(`/sites/${siteId}/lists?$filter=displayName eq '${CONFIG.spListName}'`, tok);
  if (!data.value.length) throw new Error(`Lijst "${CONFIG.spListName}" niet gevonden.`);
  _listId = data.value[0].id;
  return _listId;
}

async function spGetItems(filter = '', orderby = 'fields/DatumMelding desc') {
  const siteId = await getSiteId();
  const listId = await getListId();
  const tok    = await refreshToken();

  const params = new URLSearchParams({ $top: '5000', $orderby: orderby });
  if (filter) params.append('$filter', filter);

  const data = await graphGet(`/sites/${siteId}/lists/${listId}/items?${params}&$expand=fields`, tok);
  return data.value.map(i => ({ id: i.id, ...i.fields }));
}

async function spCreateItem(fields) {
  const siteId = await getSiteId();
  const listId = await getListId();
  const tok    = await refreshToken();

  const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${tok}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields }),
  });
  if (!r.ok) { const e = await r.text(); throw new Error(e); }
  return r.json();
}

async function spUpdateItem(itemId, fields) {
  const siteId = await getSiteId();
  const listId = await getListId();
  const tok    = await refreshToken();

  const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${tok}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(fields),
  });
  if (!r.ok) { const e = await r.text(); throw new Error(e); }
  return r.json();
}

/* ══════════════════ DOSSIERNUMMER LOGICA ══════════════════ */
async function generateDossierNumber() {
  const year = new Date().getFullYear();
  // Haal hoogste volgnummer op van dit jaar
  const filter = `fields/Dossiernummer ge '${year}-0000' and fields/Dossiernummer le '${year}-9999'`;
  let items = [];
  try { items = await spGetItems(filter, 'fields/Dossiernummer desc'); } catch {}

  let seq = 1;
  if (items.length > 0) {
    const last = items[0].Dossiernummer || '';
    const parts = last.split('-');
    const lastSeq = parseInt(parts[parts.length - 1], 10);
    if (!isNaN(lastSeq)) seq = lastSeq + 1;
  }
  return `${year}-${String(seq).padStart(4, '0')}`;
}

/* ══════════════════ NAVIGATION ══════════════════ */
function setupNavigation() {
  document.querySelectorAll('[data-view]').forEach(el => {
    el.addEventListener('click', () => navigateTo(el.dataset.view));
  });

  document.getElementById('btnLogout').addEventListener('click', async () => {
    await msalInstance.logoutPopup();
    window.location.reload();
  });
}

function navigateTo(view) {
  // Views
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  const target = document.getElementById(`view${capitalize(view)}`);
  if (target) target.classList.add('active');

  // Nav items
  document.querySelectorAll('.tn').forEach(n => {
    n.classList.toggle('active', n.dataset.view === view);
  });

  // Lazy loads
  if (view === 'dashboard') loadDashboard();
  if (view === 'review' && currentUser?.isAdmin) loadReview();
  if (view === 'new') resetForm();
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/* ══════════════════ DASHBOARD ══════════════════ */
async function loadDashboard() {
  show('loadingDash');
  hide('emptyState');
  document.getElementById('melding-list').innerHTML = '';

  try {
    let filter = '';
    if (!currentUser.isAdmin) {
      filter = `fields/Melder eq '${currentUser.email}'`;
    }
    allKlachten = await spGetItems(filter);
    hide('loadingDash');
    renderList();
  } catch (e) {
    hide('loadingDash');
    showToast('Fout bij laden: ' + e.message, 'error');
  }
}

function filterFase(type, el) {
  document.querySelectorAll('.fase-card').forEach(function(c){c.classList.remove('active-fase');});
  el.classList.add('active-fase');
  document.getElementById('fType').value = type;
  renderList();
}

function renderList(){
  if (!allKlachten) return;
  updateFaseCounts();
  var search=(document.getElementById('searchInput')?.value||'').toLowerCase();
  var ftype=(document.getElementById('fType')?.value)||'';
  var fstat=(document.getElementById('fStatus')?.value)||'';
  var items=allKlachten.filter(function(k){return(!fstat||k.Status===fstat)&&(!ftype||k.TypeKlacht===ftype)&&(!search||k.Klantnaam.toLowerCase().includes(search)||k.Dossiernummer.toLowerCase().includes(search)||(k.TypeKlacht||'').toLowerCase().includes(search));});
  var cntOpen=allKlachten.filter(function(k){return k.Status==='Wachtend op goedkeuring';}).length;
  var cntDone=allKlachten.filter(function(k){return k.Status==='Goedgekeurd'||k.Status==='Geklasseerd';}).length;
  if(document.getElementById('cnt-open'))document.getElementById('cnt-open').textContent=cntOpen+' open';
  if(document.getElementById('cnt-hoog'))document.getElementById('cnt-hoog').textContent='0 hoog';
  if(document.getElementById('cnt-done'))document.getElementById('cnt-done').textContent=cntDone+' afgehandeld';
  const badge=document.getElementById('badgePending');
  if(badge){badge.textContent=cntOpen;badge.classList.toggle('hidden',cntOpen===0);}
  var el=document.getElementById('melding-list');
  var empty=document.getElementById('emptyState');
  if(!items.length){if(el)el.innerHTML='';if(empty)empty.classList.remove('hidden');return;}
  if(empty)empty.classList.add('hidden');
  el.innerHTML=items.map(function(k){
    var sc=typeClass[k.TypeKlacht]||'';var sp=typePill[k.TypeKlacht]||'';var sb=statusBadge(k.Status);
    var artCount=k.Artikelregels?k.Artikelregels.length:0;
    var cnBadge=k.CreditnotaNr?'<div class="meta-item cn"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>'+k.CreditnotaNr+'</div>':'';
    var artBadge=artCount?'<div class="meta-item"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>'+artCount+' artikel'+(artCount>1?'en':'')+'</div>':'';
    return '<div class="melding-row '+sc+'" onclick="openDetail(\''+k.id+'\')">'
      +'<div class="melding-body">'
        +'<div class="melding-top"><span class="melding-nr">'+esc(k.Dossiernummer)+'</span>'+sp+'</div>'
        +'<div class="melding-main"><span class="knr">'+esc(k.Klantnummer)+' &middot;</span>'+esc(k.Klantnaam)+'</div>'
        +'<div class="melding-meta">'
          +'<div class="meta-item"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>'+formatDate(k.DatumMelding)+'</div>'
          +'<div class="meta-item"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>'+esc(k.MelderNaam||'\u2013')+'</div>'
          +artBadge+cnBadge
        +'</div>'
      +'</div>'
      +'<div class="melding-right"><span class="melding-bedrag">\u20ac '+formatBedrag(k.Bedrag)+'</span>'+sb+'</div>'
      +'</div>';
  }).join('');
}

function updateFaseCounts(){
  if(!allKlachten) return;
  ['Foute bestelling','Beschadiging','Prijsverschil','Kwaliteit'].forEach(function(t,i){
    var n=allKlachten.filter(function(k){return k.TypeKlacht===t&&k.Status==='Wachtend op goedkeuring';}).length;
    var el=document.getElementById('fc'+(i+1));
    if(el)el.textContent=n+' open';
  });
}

// Status filter change
document.getElementById('filterStatus')?.addEventListener('change', loadDashboard);

/* ══════════════════ REVIEW VIEW ══════════════════ */
async function loadReview() {
  const container = document.getElementById('reviewCards');
  show('loadingReview');
  hide('emptyReview');
  container.innerHTML = '';

  try {
    const items = await spGetItems(`fields/Status eq 'Wachtend op goedkeuring'`);
    hide('loadingReview');

    if (!items.length) {
      show('emptyReview');
      return;
    }

    items.forEach(item => {
      const card = document.createElement('div');
      card.className = 'review-card';
      card.innerHTML = `
        <div class="review-card-header">
          <div>
            <span class="dossier-nr" style="font-size:1rem">${esc(item.Dossiernummer)}</span>
            <span style="margin-left:0.75rem;color:var(--text-muted);font-size:0.82rem">${formatDate(item.DatumMelding)}</span>
          </div>
          ${statusPill(item.Status)}
        </div>
        <div class="review-card-meta">
          <div class="meta-item"><span class="meta-label">Klant</span><span class="meta-value">${esc(item.Klantnaam)}</span></div>
          <div class="meta-item"><span class="meta-label">Klantnr</span><span class="meta-value">${esc(item.Klantnummer)}</span></div>
          <div class="meta-item"><span class="meta-label">Factuurnr</span><span class="meta-value">${esc(item.Factuurnummer)}</span></div>
          <div class="meta-item"><span class="meta-label">Type</span><span class="meta-value">${esc(item.TypeKlacht)}</span></div>
          <div class="meta-item"><span class="meta-label">Bedrag</span><span class="meta-value" style="color:var(--brand);font-weight:700">€ ${formatBedrag(item.Bedrag)}</span></div>
          <div class="meta-item"><span class="meta-label">Ingediend door</span><span class="meta-value">${esc(item.MelderNaam || item.Melder || '–')}</span></div>
        </div>
        <div class="review-desc">${esc(item.Omschrijving)}</div>
        <div class="review-actions">
          <button class="btn btn-success" onclick="approveKlacht('${item.id}', true)">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 6L9 17l-5-5"/></svg>
            Goedkeuren
          </button>
          <button class="btn btn-danger" onclick="openReject('${item.id}')">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            Weigeren
          </button>
          <button class="btn btn-ghost btn-sm" onclick="openDetail('${item.id}')">Detail bekijken</button>
        </div>
      `;
      container.appendChild(card);
    });
  } catch (e) {
    hide('loadingReview');
    showToast('Fout: ' + e.message, 'error');
  }
}

/* ══════════════════ APPROVE / REJECT ══════════════════ */
async function approveKlacht(itemId, fromReview = false) {
  try {
    await spUpdateItem(itemId, {
      Status: 'Goedgekeurd',
      DatumGoedkeuring: new Date().toISOString(),
      BeoordeeldDoor: currentUser.email,
    });
    showToast('Klacht goedgekeurd. Creditnota staat klaar.', 'success');
    closeModal();
    await loadDashboard();
    if (fromReview) await loadReview();
  } catch (e) {
    showToast('Fout bij goedkeuren: ' + e.message, 'error');
  }
}

function openReject(itemId) {
  currentRejectId = itemId;
  document.getElementById('rejectReason').value = '';
  document.getElementById('rejectError').classList.add('hidden');
  show('rejectOverlay');
}

async function confirmReject() {
  const reason = document.getElementById('rejectReason').value.trim();
  if (!reason) {
    document.getElementById('rejectError').classList.remove('hidden');
    return;
  }

  try {
    await spUpdateItem(currentRejectId, {
      Status: 'Geweigerd',
      WeigeringReden: reason,
      DatumGoedkeuring: new Date().toISOString(),
      BeoordeeldDoor: currentUser.email,
    });
    showToast('Klacht geweigerd en gearchiveerd.', 'success');
    hide('rejectOverlay');
    closeModal();
    currentRejectId = null;
    await loadDashboard();
    if (document.getElementById('viewReview').classList.contains('active')) await loadReview();
  } catch (e) {
    showToast('Fout bij weigeren: ' + e.message, 'error');
  }
}

/* ══════════════════ DETAIL MODAL ══════════════════ */
function openDetail(id){
  var k=klachten.find(function(x){return x.id===id;});if(!k)return;
  document.getElementById('modal-title').textContent='Dossier '+k.Dossiernummer;
  var artHtml='';
  if(k.Artikelregels&&k.Artikelregels.length){
    var rows=k.Artikelregels.map(function(r){var a=parseFloat(r.aantal)||0;var p=parseFloat(String(r.prijs||0).replace(',','.'))||0;return'<tr><td style="font-family:monospace;font-weight:700;font-size:12px;color:var(--navy)">'+(r.artnr||'\u2013')+'</td><td>'+(r.naam||'\u2013')+'</td><td><span style="background:var(--gray-bg);padding:1px 7px;border-radius:4px;font-size:11px;font-weight:600;color:var(--muted)">'+(r.uom||'ST')+'</span></td><td style="text-align:right;font-weight:600">'+a+'</td><td style="text-align:right">'+(p?'\u20ac '+fmtB(p):'\u2013')+'</td><td style="text-align:right;font-weight:600;color:var(--navy)">'+(p?'\u20ac '+fmtB(a*p):'\u2013')+'</td></tr>';}).join('');
    var tot=k.Artikelregels.reduce(function(s,r){var a=parseFloat(r.aantal)||0;var p=parseFloat(String(r.prijs||0).replace(',','.'))||0;return s+a*p;},0);
    artHtml='<div style="margin-bottom:16px"><div class="d-lbl" style="margin-bottom:6px">Artikelregels</div><div style="border:1px solid var(--border);border-radius:8px;overflow:hidden"><table class="art-detail-table"><thead><tr><th>Artikelnummer</th><th>Artikelnaam</th><th>UOM</th><th style="text-align:right">Aantal</th><th style="text-align:right">Eenheidsprijs</th><th style="text-align:right">Totaal</th></tr></thead><tbody>'+rows+'</tbody><tfoot><tr><td colspan="5" style="text-align:right;font-size:12px">Totaal (excl. BTW)</td><td style="text-align:right">\u20ac '+fmtB(tot)+'</td></tr></tfoot></table></div></div>';
  } else if(k.Bedrag){artHtml='<div style="margin-bottom:16px"><div class="d-lbl" style="margin-bottom:4px">Bedrag (excl. BTW)</div><div class="d-val big">\u20ac '+fmtB(k.Bedrag)+'</div></div>';}
  var rejectHtml=k.WeigeringReden?'<div class="reject-box"><div class="d-lbl">Reden weigering</div><div class="d-val" style="font-weight:400;margin-top:4px">'+k.WeigeringReden+'</div></div>':'';
  var cnHtml=k.CreditnotaNr?'<div><div class="d-lbl">Creditnota</div><div class="d-val" style="font-family:monospace;font-weight:700;color:var(--green)">'+k.CreditnotaNr+'</div></div>':'';
  document.getElementById('modal-body').innerHTML='<div class="detail-grid"><div><div class="d-lbl">Dossiernummer</div><div class="d-val" style="font-family:monospace;font-size:15px;font-weight:700;color:var(--navy)">'+k.Dossiernummer+'</div></div><div><div class="d-lbl">Status</div><div class="d-val">'+statusBadge(k.Status)+'</div></div><div><div class="d-lbl">Datum melding</div><div class="d-val">'+fmtDate(k.DatumMelding)+'</div></div><div><div class="d-lbl">Type klacht</div><div class="d-val">'+(typePill[k.TypeKlacht]||k.TypeKlacht)+'</div></div><div><div class="d-lbl">Klantnaam</div><div class="d-val">'+k.Klantnaam+'</div></div><div><div class="d-lbl">Klantnummer</div><div class="d-val">'+k.Klantnummer+'</div></div><div><div class="d-lbl">Factuurnummer</div><div class="d-val">'+k.Factuurnummer+'</div></div>'+(k.BeoordeeldDoor?'<div><div class="d-lbl">Beoordeeld door</div><div class="d-val">'+k.BeoordeeldDoor+'</div></div>':'')+cnHtml+'<div class="d-full"><div class="d-lbl">Omschrijving</div><div class="d-val desc">'+k.Omschrijving+'</div></div><div><div class="d-lbl">Ingediend door</div><div class="d-val">'+(k.MelderNaam||'\u2013')+'</div></div></div>'+artHtml+rejectHtml;
  var foot=document.getElementById('modal-foot');
  if(k.Status==='Wachtend op goedkeuring'){foot.innerHTML='<button class="btn btn-success" onclick="approve(\''+k.id+'\')">&#10003; Goedkeuren</button><button class="btn btn-danger" onclick="openRejectModal(\''+k.id+'\')">&#10007; Weigeren</button><button class="btn btn-ghost" onclick="closeModal()">Sluiten</button>';}
  else if(k.Status==='Goedgekeurd'){foot.innerHTML='<div style="display:flex;align-items:center;gap:8px;flex:1;flex-wrap:wrap"><div style="display:flex;align-items:center;border:1.5px solid var(--border);border-radius:8px;overflow:hidden;background:var(--surface)"><span style="padding:6px 10px;background:var(--gray-bg);color:var(--muted);font-size:12px;font-weight:600;border-right:1px solid var(--border);white-space:nowrap">Creditnota</span><input id="creditnota-input" type="text" placeholder="bijv. CN2026-00123 (optioneel)" value="'+(k.CreditnotaNr||'')+'" style="border:none;padding:6px 10px;font-size:13px;color:var(--text);outline:none;width:220px;font-family:monospace;font-weight:600"/></div><button class="btn btn-success btn-sm" onclick="saveCreditnota(\''+k.id+'\')">Opslaan</button></div><button class="btn btn-ghost" onclick="closeModal()">Sluiten</button>';}
  else{foot.innerHTML='<button class="btn" onclick="closeModal()">Sluiten</button>';}
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function setupModals() {
  document.getElementById('modalClose').addEventListener('click', closeModal);
  document.getElementById('rejectClose').addEventListener('click', () => hide('rejectOverlay'));
  document.getElementById('btnRejectConfirm').addEventListener('click', confirmReject);
  document.getElementById('btnRejectCancel').addEventListener('click', () => hide('rejectOverlay'));

  // Klik buiten modal sluit
  document.getElementById('modalOverlay').addEventListener('click', e => {
    if (e.target === document.getElementById('modalOverlay')) closeModal();
  });
  document.getElementById('rejectOverlay').addEventListener('click', e => {
    if (e.target === document.getElementById('rejectOverlay')) hide('rejectOverlay');
  });
}

function closeModal() { hide('modalOverlay'); }

/* ══════════════════ FORM ══════════════════ */
let selectedFiles = [];

function setupForm() {
  document.getElementById('btnSubmit').addEventListener('click', submitKlacht);

  // File drag & drop
  const drop = document.getElementById('fileDrop');
  const input = document.getElementById('fBijlage');

  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('drag-over');
    addFiles(e.dataTransfer.files);
  });

  input.addEventListener('change', () => addFiles(input.files));
}

function addFiles(fileList) {
  Array.from(fileList).forEach(f => {
    if (f.size > 10 * 1024 * 1024) {
      showToast(`${f.name} is te groot (max 10 MB).`, 'error');
      return;
    }
    if (!selectedFiles.find(x => x.name === f.name && x.size === f.size)) {
      selectedFiles.push(f);
    }
  });
  renderFileList();
}

function renderFileList() {
  const list = document.getElementById('fileList');
  list.innerHTML = '';
  selectedFiles.forEach((f, i) => {
    const item = document.createElement('div');
    item.className = 'file-item';
    item.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/></svg>
      <span>${esc(f.name)}</span>
      <small style="color:var(--text-light);white-space:nowrap">${(f.size/1024).toFixed(0)} KB</small>
      <button onclick="removeFile(${i})" title="Verwijderen">✕</button>
    `;
    list.appendChild(item);
  });
}

function removeFile(i) {
  selectedFiles.splice(i, 1);
  renderFileList();
}

function resetForm() {
  ['fKlant','fKlantnr','fFactuurnr','fOmschrijving','fBedrag'].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.value = ''; el.classList.remove('invalid'); }
  });
  document.getElementById('fType').value = '';
  selectedFiles = [];
  renderFileList();
  hide('formSuccess');
  hide('formError');
}

async function submitKlacht() {
  hide('formSuccess');
  hide('formError');

  // Validatie
  const fields = {
    Klantnaam:    document.getElementById('fKlant').value.trim(),
    Klantnummer:  document.getElementById('fKlantnr').value.trim(),
    Factuurnummer: document.getElementById('fFactuurnr').value.trim(),
    TypeKlacht:   document.getElementById('fType').value,
    Omschrijving: document.getElementById('fOmschrijving').value.trim(),
    Bedrag:       parseFloat(document.getElementById('fBedrag').value),
  };

  let valid = true;
  Object.entries(fields).forEach(([key, val]) => {
    let elId;
    if (key === 'Klantnaam')    elId = 'fKlant';
    if (key === 'Klantnummer')  elId = 'fKlantnr';
    if (key === 'Factuurnummer') elId = 'fFactuurnr';
    if (key === 'TypeKlacht')   elId = 'fType';
    if (key === 'Omschrijving') elId = 'fOmschrijving';
    if (key === 'Bedrag')       elId = 'fBedrag';
    const el = document.getElementById(elId);
    if (!val && val !== 0) { el.classList.add('invalid'); valid = false; }
    else if (key === 'Bedrag' && (isNaN(val) || val < 0)) { el.classList.add('invalid'); valid = false; }
    else el.classList.remove('invalid');
  });

  if (!valid) {
    showError('formErrorMsg', 'Vul alle verplichte velden correct in.');
    show('formError');
    return;
  }

  const btn = document.getElementById('btnSubmit');
  btn.disabled = true;
  btn.textContent = 'Bezig met indienen…';

  try {
    const dossiernummer = await generateDossierNumber();

    await spCreateItem({
      ...fields,
      Dossiernummer: dossiernummer,
      Status:        'Wachtend op goedkeuring',
      DatumMelding:  new Date().toISOString(),
      Melder:        currentUser.email,
      MelderNaam:    currentUser.name,
    });

    document.getElementById('successDossier').textContent = ` Dossiernummer: ${dossiernummer}`;
    show('formSuccess');
    resetForm();
    await loadDashboard();
  } catch (e) {
    showError('formErrorMsg', 'Fout bij indienen: ' + e.message);
    show('formError');
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 2L11 13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg> Klacht indienen`;
  }
}

/* ══════════════════ EXCEL EXPORT ══════════════════ */
document.getElementById('btnExport').addEventListener('click', exportToExcel);

function exportToExcel() {
  if (!allKlachten.length) { showToast('Geen data om te exporteren.', 'error'); return; }

  const rows = allKlachten.map(k => ({
    'Dossiernummer':    k.Dossiernummer || '',
    'Datum Melding':    k.DatumMelding ? formatDateExcel(k.DatumMelding) : '',
    'Klantnaam':        k.Klantnaam || '',
    'Klantnummer':      k.Klantnummer || '',
    'Factuurnummer':    k.Factuurnummer || '',
    'Type Klacht':      k.TypeKlacht || '',
    'Omschrijving':     k.Omschrijving || '',
    'Bedrag excl. BTW': typeof k.Bedrag === 'number' ? k.Bedrag : parseFloat(k.Bedrag) || 0,
    'Status':           k.Status || '',
    'Melder':           k.MelderNaam || k.Melder || '',
    'Beoordeeld door':  k.BeoordeeldDoor || '',
    'Datum beoordeling': k.DatumGoedkeuring ? formatDateExcel(k.DatumGoedkeuring) : '',
    'Reden weigering':  k.WeigeringReden || '',
  }));

  const ws = XLSX.utils.json_to_sheet(rows);

  // Opmaak: brede kolommen
  ws['!cols'] = [
    {wch:14}, {wch:14}, {wch:28}, {wch:14}, {wch:16},
    {wch:18}, {wch:50}, {wch:18}, {wch:24}, {wch:24},
    {wch:22}, {wch:18}, {wch:40},
  ];

  // Valuta opmaak kolom H (index 7 = Bedrag)
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 7 })];
    if (cell && cell.t === 'n') cell.z = '€ #,##0.00';
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Klachten');

  const filename = `Verpa_Klachten_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast(`Export klaar: ${filename}`, 'success');
}

/* ══════════════════ IMPORT ══════════════════ */
let importData = [];

function setupImport() {
  const drop    = document.getElementById('importDrop');
  const input   = document.getElementById('importFile');
  const btnTmpl = document.getElementById('btnDownloadTemplate');

  // Template download
  btnTmpl.addEventListener('click', e => {
    e.preventDefault();
    downloadImportTemplate();
  });

  // Drag & drop
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', e => {
    e.preventDefault(); drop.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) processImportFile(e.dataTransfer.files[0]);
  });

  input.addEventListener('change', () => { if (input.files[0]) processImportFile(input.files[0]); });

  document.getElementById('btnImportConfirm').addEventListener('click', runImport);
  document.getElementById('btnImportCancel').addEventListener('click', () => {
    hide('importPreview');
    hide('importResult');
    importData = [];
  });
}

function downloadImportTemplate() {
  const headers = ['Dossiernummer','DatumMelding','Klantnaam','Klantnummer','Factuurnummer',
                   'TypeKlacht','Omschrijving','Bedrag','Status','Melder','MelderNaam'];
  const ws = XLSX.utils.aoa_to_sheet([headers, ['2025-0001','15-06-2025','Voorbeeldbedrijf BV','K00001','F2025-001','Kwaliteit','Voorbeeld omschrijving',125.50,'Goedgekeurd','gebruiker@verpa.be','Jan Janssen']]);
  ws['!cols'] = headers.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Import');
  XLSX.writeFile(wb, 'Verpa_Klachten_Import_Template.xlsx');
}

function processImportFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { raw: false, dateNF: 'DD-MM-YYYY' });

      if (!raw.length) { showToast('Bestand is leeg.', 'error'); return; }

      importData = raw;
      renderImportPreview(raw);
    } catch (err) {
      showToast('Fout bij lezen bestand: ' + err.message, 'error');
    }
  };
  reader.readAsBinaryString(file);
}

function renderImportPreview(rows) {
  const tbl = document.getElementById('tblImportPreview');
  const preview = rows.slice(0, 5);
  const headers = Object.keys(rows[0]);

  let html = '<thead><tr>' + headers.map(h => `<th>${esc(h)}</th>`).join('') + '</tr></thead><tbody>';
  preview.forEach(row => {
    html += '<tr>' + headers.map(h => `<td>${esc(String(row[h] ?? ''))}</td>`).join('') + '</tr>';
  });
  html += '</tbody>';

  tbl.innerHTML = html;
  document.getElementById('importInfo').textContent = `${rows.length} records gevonden – originele datums en dossiernummers worden behouden.`;
  show('importPreview');
  hide('importResult');
}

async function runImport() {
  if (!importData.length) return;

  const btnConfirm  = document.getElementById('btnImportConfirm');
  const progressDiv = document.getElementById('importProgress');
  const fill        = document.getElementById('progressFill');
  const label       = document.getElementById('progressLabel');
  const resultDiv   = document.getElementById('importResult');

  btnConfirm.disabled = true;
  hide('importPreview');
  show('importProgress');

  let ok = 0, fail = 0;

  for (let i = 0; i < importData.length; i++) {
    const row = importData[i];
    try {
      // Parseer datum naar ISO
      let datumISO = new Date().toISOString();
      if (row.DatumMelding) {
        const parts = row.DatumMelding.split(/[-/]/);
        if (parts.length === 3) {
          // DD-MM-YYYY
          if (parts[0].length === 2) {
            datumISO = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`).toISOString();
          } else {
            // YYYY-MM-DD
            datumISO = new Date(row.DatumMelding).toISOString();
          }
        }
      }

      await spCreateItem({
        Dossiernummer: row.Dossiernummer || '',
        DatumMelding:  datumISO,
        Klantnaam:     row.Klantnaam || '',
        Klantnummer:   row.Klantnummer || '',
        Factuurnummer: row.Factuurnummer || '',
        TypeKlacht:    row.TypeKlacht || '',
        Omschrijving:  row.Omschrijving || '',
        Bedrag:        parseFloat(String(row.Bedrag).replace(',', '.')) || 0,
        Status:        row.Status || 'Geklasseerd',
        Melder:        row.Melder || 'import@verpa',
        MelderNaam:    row.MelderNaam || 'Import',
        IsHistorisch:  true,
      });
      ok++;
    } catch {
      fail++;
    }

    const pct = Math.round(((i + 1) / importData.length) * 100);
    fill.style.width = pct + '%';
    label.textContent = `${i + 1} / ${importData.length} records verwerkt`;
  }

  hide('importProgress');

  resultDiv.className = ok > 0 ? 'alert alert-success' : 'alert alert-error';
  resultDiv.innerHTML = `
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
    <div><strong>Import voltooid:</strong> ${ok} records succesvol geïmporteerd${fail > 0 ? `, ${fail} mislukt` : ''}.</div>
  `;
  show('importResult');

  btnConfirm.disabled = false;
  importData = [];
  await loadDashboard();
}

/* ══════════════════ UTILS ══════════════════ */
function show(id) { document.getElementById(id)?.classList.remove('hidden'); }
function hide(id) { document.getElementById(id)?.classList.add('hidden'); }
function showError(id, msg) { const el = document.getElementById(id); if (el) el.textContent = msg; }
function esc(s) { if (!s) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function formatDate(iso) {
  if (!iso) return '–';
  try {
    const d = new Date(iso);
    return d.toLocaleDateString('nl-BE', { day:'2-digit', month:'2-digit', year:'numeric' });
  } catch { return iso; }
}

function formatDateExcel(iso) {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    return `${dd}-${mm}-${d.getFullYear()}`;
  } catch { return iso; }
}

function formatBedrag(val) {
  const n = parseFloat(val);
  if (isNaN(n)) return '–';
  return n.toLocaleString('nl-BE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function statusPill(status) {
  const map = {
    'Wachtend op goedkeuring': ['pill-pending',  '⏳'],
    'Goedgekeurd':             ['pill-approved', '✓'],
    'Geweigerd':               ['pill-rejected', '✗'],
    'Geklasseerd':             ['pill-archived', '📁'],
  };
  const [cls, icon] = map[status] || ['pill-archived', '–'];
  return `<span class="status-pill ${cls}"><span class="pill-dot"></span>${esc(status)}</span>`;
}

function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = `toast${type ? ' toast-'+type : ''}`;
  t.classList.remove('hidden');
  setTimeout(() => t.classList.add('hidden'), 3500);
}

/* ══════════════════ START ══════════════════ */
init();
