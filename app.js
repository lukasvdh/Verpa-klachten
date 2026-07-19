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

const GRAPH_SCOPES = ['User.Read', 'Sites.ReadWrite.All'];

/* ── BUSINESS CENTRAL CONFIG ─────────────────────────────────────────
   Vereiste App Registration permissie (Delegated):
     - Dynamics 365 Business Central → Financials.ReadWrite.All
   ─────────────────────────────────────────────────────────────────── */
const BC_SCOPE      = 'https://api.businesscentral.dynamics.com/user_impersonation';
const BC_BASE       = 'https://api.businesscentral.dynamics.com/v2.0';
const BC_TENANT     = CONFIG.tenantId;
const BC_ENV        = 'Verpa_Accept_09022026';
let   BC_COMPANY_ID = null;         // wordt opgehaald bij eerste gebruik

async function getBCToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  try {
    const resp = await msalInstance.acquireTokenSilent({ scopes: [BC_SCOPE], account: accounts[0] });
    return resp.accessToken;
  } catch {
    try {
      const resp = await msalInstance.acquireTokenPopup({ scopes: [BC_SCOPE] });
      return resp.accessToken;
    } catch (e) {
      console.error('BC token fout:', e);
      return null;
    }
  }
}

async function getBCCompanyId() {
  if (BC_COMPANY_ID) return BC_COMPANY_ID;
  const tok = await getBCToken();
  if (!tok) throw new Error('Geen BC-token beschikbaar');
  const r = await fetch(`${BC_BASE}/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`, {
    headers: { Authorization: `Bearer ${tok}` },
  });
  if (!r.ok) throw new Error(`BC companies: ${r.status}`);
  const data = await r.json();
  if (!data.value.length) throw new Error('Geen BC-company gevonden');
  BC_COMPANY_ID = data.value[0].id;
  return BC_COMPANY_ID;
}

async function bcZoekKlanten(zoekterm) {
  const tok       = await getBCToken();
  const companyId = await getBCCompanyId();
  const term = zoekterm.replace(/'/g, "''"); // UI toont hoofdletters via CSS, BC zoekt op originele input
  const select    = 'id,number,displayName,email,phoneNumber,addressLine1,city,postalCode';
  const base      = `${BC_BASE}/${BC_TENANT}/${BC_ENV}/api/v2.0/companies(${companyId})/customers`;

  const headers = { Authorization: `Bearer ${tok}` };

  // BC ondersteunt geen 'or' over verschillende velden → twee aparte calls
  const [r1, r2] = await Promise.all([
    fetch(`${base}?$filter=${encodeURIComponent("startswith(number,'" + term + "')")}&$top=8&$select=${select}`, { headers }),
    fetch(`${base}?$filter=${encodeURIComponent("contains(displayName,'" + term + "')")}&$top=8&$select=${select}`, { headers }),
  ]);

  if (!r1.ok && !r2.ok) throw new Error(`BC klanten: ${r1.status}`);

  const [d1, d2] = await Promise.all([
    r1.ok ? r1.json() : { value: [] },
    r2.ok ? r2.json() : { value: [] },
  ]);

  // Samenvoegen zonder duplicaten op klantnummer
  const seen = new Set();
  return [...d1.value, ...d2.value].filter(k => {
    if (seen.has(k.number)) return false;
    seen.add(k.number);
    return true;
  }).slice(0, 12);
}

const BC_DEMO_KLANTEN = [
  { number: 'K00001', displayName: 'Carrefour Belgium NV',       addressLine1: 'Olympiadenlaan 20',  postalCode: '1140', city: 'Evere'       },
  { number: 'K00002', displayName: 'Colruyt Group NV',           addressLine1: 'Edingensesteenweg 196', postalCode: '1500', city: 'Halle'    },
  { number: 'K00003', displayName: 'Delhaize Belgium',           addressLine1: 'Square Marie Curie 40', postalCode: '1070', city: 'Anderlecht' },
  { number: 'K00004', displayName: 'AZ Turnhout',                addressLine1: 'Rubensstraat 166',   postalCode: '2300', city: 'Turnhout'    },
  { number: 'K00005', displayName: 'Gemeente Mol',               addressLine1: 'Molenhoek 2',        postalCode: '2400', city: 'Mol'         },
  { number: 'K00006', displayName: 'Hotel Industria',            addressLine1: 'Kleinhoefstraat 9',  postalCode: '2440', city: 'Geel'        },
  { number: 'K00007', displayName: 'Lunch Garden Hasselt',       addressLine1: 'Genkersteenweg 14',  postalCode: '3500', city: 'Hasselt'     },
  { number: 'K00008', displayName: 'Verpa Benelux NV (intern)',  addressLine1: 'Industrielaan 10',   postalCode: '2430', city: 'Laakdal'     },
];

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

  // Check app-rol via ID token claims (zelfde aanpak als andere Verpa apps)
  const roles = (account && account.idTokenClaims && account.idTokenClaims.roles) || [];
  const isAdmin = roles.includes('admin');

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

  // Deep-link: open dossier direct via ?dossier=2026-XXXX
  const urlParams = new URLSearchParams(window.location.search);
  const dossierParam = urlParams.get('dossier');
  if (dossierParam) {
    const match = allKlachten.find(k => k.Dossiernummer === dossierParam);
    if (match) openDetail(match.id);
  }
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
    headers: {
      Authorization: `Bearer ${token}`,
      'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly',
    },
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

async function spDeleteItem(itemId) {
  const listId = await getListId();
  const siteId = await getSiteId();
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}`;
  const r = await fetch(url, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${currentUser.token}` },
  });
  if (!r.ok && r.status !== 204) throw new Error(`DELETE ${r.status}`);
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
  const prefix = `${year}-`;

  // Haal alle items op met startswith filter (Graph ondersteunt geen ge/le op tekst)
  let items = [];
  try {
    items = await spGetItems(
      `startswith(fields/Dossiernummer,'${prefix}')`,
      'fields/Dossiernummer desc'
    );
  } catch {}

  let seq = 1;
  if (items.length > 0) {
    // Zoek het hoogste volgnummer manueel (sortering op tekst kan fout gaan)
    let max = 0;
    items.forEach(item => {
      const nr = item.Dossiernummer || '';
      if (nr.startsWith(prefix)) {
        const n = parseInt(nr.slice(prefix.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
    seq = max + 1;
  }
  return `${prefix}${String(seq).padStart(4, '0')}`;
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
  document.querySelectorAll('.tn, .mobile-nav-item').forEach(n => {
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
  document.getElementById('filterType').value = type;
  renderList();
}

/* ══════════════════ RENDER HELPERS ══════════════════ */
const typeClass = {
  'Foute bestelling': 'fase-ontvangst',
  'Beschadiging':     'fase-voorraad',
  'Prijsverschil':    'fase-verzending',
  'Administratief':   'fase-admin',
  'Kwaliteit':        'fase-levering',
};

const typePill = {
  'Foute bestelling': '<span class="fase-pill fp-ontvangst">Foute bestelling</span>',
  'Beschadiging':     '<span class="fase-pill fp-voorraad">Beschadiging</span>',
  'Prijsverschil':    '<span class="fase-pill fp-verzending">Prijsverschil</span>',
  'Administratief':   '<span class="fase-pill fp-admin">Administratief</span>',
  'Kwaliteit':        '<span class="fase-pill fp-levering">Kwaliteit</span>',
};

function statusBadge(status) {
  const map = {
    'Wachtend op goedkeuring': 's-wachtend',
    'Goedgekeurd':             's-goedgekeurd',
    'Geweigerd':               's-geweigerd',
    'Geklasseerd':             's-geklasseerd',
  };
  const cls = map[status] || 's-geklasseerd';
  return `<span class="status-badge ${cls}"><span class="sb-dot"></span>${esc(status)}</span>`;
}

function renderList(){
  if (!allKlachten) return;
  updateFaseCounts();
  var search=(document.getElementById('searchInput')?.value||'').toLowerCase();
  var ftype=(document.getElementById('filterType')?.value)||'';
  var fstat=(document.getElementById('filterStatus')?.value)||'';
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
    var artCount=0;try{if(k.Artikelregels){var _ar=JSON.parse(k.Artikelregels);artCount=_ar.filter(function(r){return r.artnr||r.naam;}).length;}}catch(e){}
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
function fmtB(n){return parseFloat(n).toLocaleString('nl-BE',{minimumFractionDigits:2,maximumFractionDigits:2});}
function fmtDate(iso){return formatDate(iso);}
function openDetail(id){
  var k=allKlachten.find(function(x){return x.id===id;});if(!k)return;
  document.getElementById('modalTitle').textContent='Dossier '+k.Dossiernummer;
  var artHtml='';
  var _artParsed=[];try{if(k.Artikelregels)_artParsed=JSON.parse(k.Artikelregels).filter(function(r){return r.artnr||r.naam;});}catch(e){}
  if(_artParsed.length){
    var rows=_artParsed.map(function(r){var a=parseFloat(r.aantal)||0;var p=parseFloat(String(r.prijs||0).replace(',','.'))||0;return'<tr><td style="font-family:monospace;font-weight:700;font-size:12px;color:var(--navy)">'+(r.artnr||'\u2013')+'</td><td>'+(r.naam||'\u2013')+'</td><td><span style="background:var(--gray-bg);padding:1px 7px;border-radius:4px;font-size:11px;font-weight:600;color:var(--muted)">'+(r.uom||'ST')+'</span></td><td style="text-align:right;font-weight:600">'+a+'</td><td style="text-align:right">'+(p?'\u20ac '+fmtB(p):'\u2013')+'</td><td style="text-align:right;font-weight:600;color:var(--navy)">'+(p?'\u20ac '+fmtB(a*p):'\u2013')+'</td></tr>';}).join('');
    var tot=_artParsed.reduce(function(s,r){var a=parseFloat(r.aantal)||0;var p=parseFloat(String(r.prijs||0).replace(',','.'))||0;return s+a*p;},0);
    artHtml='<div style="margin-bottom:16px"><div class="d-lbl" style="margin-bottom:6px">Artikelregels</div><div style="border:1px solid var(--border);border-radius:8px;overflow:hidden"><table class="art-detail-table"><thead><tr><th>Artikelnummer</th><th>Artikelnaam</th><th>UOM</th><th style="text-align:right">Aantal</th><th style="text-align:right">Eenheidsprijs</th><th style="text-align:right">Totaal</th></tr></thead><tbody>'+rows+'</tbody><tfoot><tr><td colspan="5" style="text-align:right;font-size:12px">Totaal (excl. BTW)</td><td style="text-align:right">\u20ac '+fmtB(tot)+'</td></tr></tfoot></table></div></div>';
  } else if(k.Bedrag){artHtml='<div style="margin-bottom:16px"><div class="d-lbl" style="margin-bottom:4px">Bedrag (excl. BTW)</div><div class="d-val big">\u20ac '+fmtB(k.Bedrag)+'</div></div>';}
  var rejectHtml=k.WeigeringReden?'<div class="reject-box"><div class="d-lbl">Reden weigering</div><div class="d-val" style="font-weight:400;margin-top:4px">'+k.WeigeringReden+'</div></div>':'';
  var cnHtml=k.CreditnotaNr?'<div><div class="d-lbl">Creditnota</div><div class="d-val" style="font-family:monospace;font-weight:700;color:var(--green)">'+k.CreditnotaNr+'</div></div>':'';
  document.getElementById('modalBody').innerHTML='<div class="detail-grid"><div><div class="d-lbl">Dossiernummer</div><div class="d-val" style="font-family:monospace;font-size:15px;font-weight:700;color:var(--navy)">'+k.Dossiernummer+'</div></div><div><div class="d-lbl">Status</div><div class="d-val">'+statusBadge(k.Status)+'</div></div><div><div class="d-lbl">Datum melding</div><div class="d-val">'+fmtDate(k.DatumMelding)+'</div></div><div><div class="d-lbl">Type klacht</div><div class="d-val">'+(typePill[k.TypeKlacht]||k.TypeKlacht)+'</div></div><div><div class="d-lbl">Klantnaam</div><div class="d-val">'+k.Klantnaam+'</div></div><div><div class="d-lbl">Klantnummer</div><div class="d-val">'+k.Klantnummer+'</div></div><div><div class="d-lbl">Factuurnummer</div><div class="d-val">'+k.Factuurnummer+'</div></div>'+(k.BeoordeeldDoor?'<div><div class="d-lbl">Beoordeeld door</div><div class="d-val">'+k.BeoordeeldDoor+'</div></div>':'')+cnHtml+'<div class="d-full"><div class="d-lbl">Omschrijving</div><div class="d-val desc">'+k.Omschrijving+'</div></div><div><div class="d-lbl">Ingediend door</div><div class="d-val">'+(k.MelderNaam||'\u2013')+'</div></div></div>'+artHtml+rejectHtml;
  var foot=document.getElementById('modalFooter');
  var retourBtn='<button class="btn btn-secondary" onclick="printRetour(\''+k.id+'\')">&#128196; Retourkaart</button>';
  var delBtn=currentUser.isAdmin?'<button class="btn btn-danger" style="margin-left:auto" onclick="deleteKlacht(\''+k.id+'\',\''+k.Dossiernummer+'\')" title="Verwijderen">&#128465; Verwijderen</button>':'';
  if(k.Status==='Wachtend op goedkeuring'){foot.innerHTML='<button class="btn btn-success" onclick="approveKlacht(\''+k.id+'\')">&#10003; Goedkeuren</button><button class="btn btn-danger" onclick="openReject(\''+k.id+'\')">&#10007; Weigeren</button><button class="btn btn-ghost" onclick="closeModal()">Sluiten</button>'+retourBtn+delBtn;}
  else if(k.Status==='Goedgekeurd'){foot.innerHTML='<div style="display:flex;align-items:center;gap:8px;flex:1;flex-wrap:wrap"><div style="display:flex;align-items:center;border:1.5px solid var(--border);border-radius:8px;overflow:hidden;background:var(--surface)"><span style="padding:6px 10px;background:var(--gray-bg);color:var(--muted);font-size:12px;font-weight:600;border-right:1px solid var(--border);white-space:nowrap">Creditnota</span><input id="creditnota-input" type="text" placeholder="bijv. CN2026-00123 (optioneel)" value="'+(k.CreditnotaNr||'')+'" style="border:none;padding:6px 10px;font-size:13px;color:var(--text);outline:none;width:220px;font-family:monospace;font-weight:600"/></div><button class="btn btn-success btn-sm" onclick="saveCreditnota(\''+k.id+'\')">Opslaan</button></div><button class="btn btn-ghost" onclick="closeModal()">Sluiten</button>'+retourBtn+delBtn;}
  else{foot.innerHTML='<button class="btn" onclick="closeModal()">Sluiten</button>'+retourBtn+delBtn;}
  document.getElementById('modalOverlay').classList.remove('hidden');
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

async function deleteKlacht(itemId, dossiernummer) {
  if (!confirm('Dossier ' + dossiernummer + ' definitief verwijderen? Dit kan niet ongedaan worden gemaakt.')) return;
  try {
    await spDeleteItem(itemId);
    closeModal();
    showToast('Dossier ' + dossiernummer + ' verwijderd.', 'success');
    await loadDashboard();
  } catch(e) {
    showToast('Fout bij verwijderen: ' + e.message, 'error');
  }
}

async function saveCreditnota(itemId) {
  const val = document.getElementById('creditnota-input')?.value.trim();
  try {
    await spUpdateItem(itemId, { CreditnotaNr: val });
    showToast('Creditnota opgeslagen.', 'success');
    closeModal();
    await loadDashboard();
  } catch(e) {
    showToast('Fout bij opslaan: ' + e.message, 'error');
  }
}

/* ══════════════════ BC KLANTZOEKER ══════════════════ */
let bcZoekTimeout = null;

function bcZoekKlant(waarde) {
  clearTimeout(bcZoekTimeout);
  const status     = document.getElementById('bcZoekStatus');
  const suggesties = document.getElementById('bcSuggesties');

  if (waarde.trim().length < 2) {
    status.classList.add('hidden');
    suggesties.classList.add('hidden');
    return;
  }

  status.textContent = 'Zoeken…';
  status.classList.remove('hidden');
  suggesties.classList.add('hidden');

  bcZoekTimeout = setTimeout(async () => {
    try {
      const klanten = await bcZoekKlanten(waarde.trim());
      status.classList.add('hidden');

      if (!klanten.length) {
        suggesties.innerHTML = '<div class="bc-sug-leeg">Geen klanten gevonden in BC</div>';
        suggesties.classList.remove('hidden');
        return;
      }

      toonSuggesties(klanten, false, suggesties);
    } catch (e) {
      // BC niet bereikbaar → filter op demodata
      const term    = waarde.trim().toLowerCase();
      const matches = BC_DEMO_KLANTEN.filter(k =>
        k.number.toLowerCase().includes(term) ||
        k.displayName.toLowerCase().includes(term) ||
        k.city.toLowerCase().includes(term)
      );

      status.textContent = '⚠ BC niet beschikbaar – demodata';
      status.classList.remove('hidden');

      if (!matches.length) {
        suggesties.innerHTML = '<div class="bc-sug-leeg">Geen demo-klanten gevonden</div>';
      } else {
        toonSuggesties(matches, true, suggesties);
      }
      suggesties.classList.remove('hidden');
    }
  }, 320);
}

function toonSuggesties(klanten, isDemo, container) {
  container.innerHTML = klanten.map(k => `
    <div class="bc-sug-item${isDemo ? ' bc-sug-demo' : ''}" onclick="bcSelecteerKlant(${JSON.stringify(k).replace(/"/g, '&quot;')})">
      <div class="bc-sug-nr">${esc(k.number)}${isDemo ? ' <span class="bc-demo-tag">DEMO</span>' : ''}</div>
      <div class="bc-sug-naam">${esc(k.displayName)}</div>
      <div class="bc-sug-adres">${esc([k.addressLine1, k.postalCode, k.city].filter(Boolean).join(' · '))}</div>
    </div>
  `).join('');
  container.classList.remove('hidden');
}

let bcHuidigeLeveradressen = [];

async function bcHaalLeveradressen(klantNr) {
  const tok       = await getBCToken();
  const companyId = await getBCCompanyId();
  // Custom Verpa API page (verpa-bc-extension)
  const url = `${BC_BASE}/${BC_TENANT}/${BC_ENV}/api/verpa/klachten/v1.0/companies(${companyId})/shipToAddresses?$filter=customerNo eq '${klantNr.replace(/'/g, "''")}'&$select=code,name,address,city,postCode`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${tok}` } });
  if (!r.ok) return [];
  const data = await r.json();
  return (data.value || []).map(a => ({
    code:         a.code,
    displayName:  a.name,
    addressLine1: a.address,
    city:         a.city,
    postalCode:   a.postCode,
  }));
}

function bcVulAdres(straat, postcode, gemeente) {
  const set = (id, val) => { const el = document.getElementById(id); if (el) { el.value = val || ''; el.classList.remove('invalid'); } };
  set('fStraat',   straat);
  set('fPostcode', postcode);
  set('fGemeente', gemeente);
}

function bcSelecteerLeveradres(code) {
  if (!code) {
    // Terug naar facturatieadres
    const klant = bcHuidigeLeveradressen._klant;
    if (klant) bcVulAdres(klant.addressLine1, klant.postalCode, klant.city);
    return;
  }
  const adres = bcHuidigeLeveradressen.find(a => a.code === code);
  if (adres) bcVulAdres(adres.addressLine1, adres.postalCode, adres.city);
}

async function bcSelecteerKlant(klant) {
  // Vul basisvelden in
  const set = (id, val) => { const el = document.getElementById(id); if (el) { el.value = val || ''; el.classList.remove('invalid'); } };
  set('fKlant',    klant.displayName);
  set('fKlantnr',  klant.number);
  bcVulAdres(klant.addressLine1, klant.postalCode, klant.city);

  // Toon bevestiging
  document.getElementById('bcGeselecteerdNaam').textContent = `${klant.number} – ${klant.displayName}`;
  document.getElementById('bcGeselecteerd').classList.remove('hidden');

  // Sluit dropdown en leeg zoekveld
  document.getElementById('bcKlantZoek').value = '';
  document.getElementById('bcSuggesties').classList.add('hidden');
  document.getElementById('bcZoekStatus').classList.add('hidden');

  // Haal leveradressen op (enkel als klant een BC id heeft)
  const leveradresBlock  = document.getElementById('bcLeveradresBlock');
  const leveradresSelect = document.getElementById('bcLeveradresSelect');
  leveradresBlock.classList.add('hidden');
  leveradresSelect.innerHTML = '<option value="">— Facturatieadres gebruiken —</option>';
  bcHuidigeLeveradressen = [];

  if (klant.number) {
    try {
      const adressen = await bcHaalLeveradressen(klant.number);
      if (adressen.length) {
        bcHuidigeLeveradressen = adressen;
        bcHuidigeLeveradressen._klant = klant; // bewaar voor reset
        adressen.forEach(a => {
          const opt = document.createElement('option');
          opt.value = a.code;
          opt.textContent = `${a.code}${a.displayName ? ' – ' + a.displayName : ''} · ${[a.addressLine1, a.postalCode, a.city].filter(Boolean).join(', ')}`;
          leveradresSelect.appendChild(opt);
        });
        leveradresBlock.classList.remove('hidden');
      }
    } catch (e) {
      // Geen leveradressen beschikbaar – stilletjes negeren
    }
  }
}

function bcResetKlant() {
  ['fKlant','fKlantnr','fStraat','fPostcode','fGemeente'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.getElementById('bcGeselecteerd').classList.add('hidden');
  document.getElementById('bcLeveradresBlock').classList.add('hidden');
  document.getElementById('bcLeveradresSelect').innerHTML = '<option value="">— Facturatieadres gebruiken —</option>';
  bcHuidigeLeveradressen = [];
  document.getElementById('bcKlantZoek').value = '';
  document.getElementById('bcKlantZoek').focus();
}

// Sluit dropdown bij klik buiten de zoeker
document.addEventListener('click', e => {
  if (!e.target.closest('.bc-search-wrap') && !e.target.closest('.bc-sug-item')) {
    const sug = document.getElementById('bcSuggesties');
    if (sug) sug.classList.add('hidden');
  }
});

/* ══════════════════ BC SYNC ══════════════════ */
async function bcSyncKlacht(klacht) {
  const tok       = await getBCToken();
  const companyId = await getBCCompanyId();
  const url       = `${BC_BASE}/${BC_TENANT}/${BC_ENV}/api/verpa/klachten/v1.0/companies(${companyId})/klachten`;

  // Map app-waarden naar BC enum-namen
  const typeMap = {
    'Foute bestelling': 'FouteBestellling',
    'Beschadiging':     'Beschadiging',
    'Prijsverschil':    'Prijsverschil',
    'Administratief':   'Administratief',
    'Kwaliteit':        'Kwaliteit',
  };

  const statusMap = {
    'Wachtend op goedkeuring': 'Open',
    'Goedgekeurd':             'Goedgekeurd',
    'Geweigerd':               'Geweigerd',
    'Geklasseerd':             'Geklasseerd',
  };

  const body = {
    dossiernummer:  klacht.dossiernummer,
    klantnummer:    klacht.klantnummer,
    datumMelding:   klacht.datumMelding,
    typeKlacht:     typeMap[klacht.typeKlacht] || klacht.typeKlacht,
    omschrijving:   klacht.omschrijving,
    factuurnummer:  klacht.factuurnummer,
    bedrag:         klacht.bedrag,
    status:         statusMap[klacht.status] || 'Open',
    melder:         klacht.melder,
    melderNaam:     klacht.melderNaam,
    straat:         klacht.straat,
    postcode:       klacht.postcode,
    gemeente:       klacht.gemeente,
    leveradresCode: klacht.leveradresCode,
    sharePointId:   klacht.sharePointId,
  };

  const r = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization:  `Bearer ${tok}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  if (!r.ok) {
    const err = await r.json().catch(() => ({}));
    const msg = err?.error?.message || `BC POST mislukt: ${r.status}`;
    // 409 = record bestaat al → als "already exists" behandelen
    if (r.status === 409 || msg.includes('already exists') || msg.includes('al bestaat')) {
      throw new Error('already exists');
    }
    throw new Error(msg);
  }

  return await r.json();
}

/* ══════════════════ BC SYNC ALLE KLACHTEN ══════════════════ */
async function bcSyncAlleKlachten() {
  const btn      = document.getElementById('btnBCSyncAll');
  const progress = document.getElementById('bcSyncProgress');
  const fill     = document.getElementById('bcSyncFill');
  const label    = document.getElementById('bcSyncLabel');
  const result   = document.getElementById('bcSyncResult');

  btn.disabled = true;
  btn.textContent = 'Bezig…';
  progress.classList.remove('hidden');
  result.classList.add('hidden');

  try {
    // 1. Haal alle klachten op uit SharePoint
    const siteId = await getSiteId();
    const listId = await getListId();
    const tok    = await refreshToken();

    let all = [], url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;
    while (url) {
      const r = await fetch(url, { headers: { Authorization: `Bearer ${tok}` } });
      const d = await r.json();
      all = all.concat(d.value || []);
      url = d['@odata.nextLink'] || null;
    }

    const totaal = all.length;
    let geslaagd = 0, overgeslagen = 0, mislukt = 0;

    for (let i = 0; i < all.length; i++) {
      const f = all[i].fields;
      const pct = Math.round(((i + 1) / totaal) * 100);
      fill.style.width = pct + '%';
      label.textContent = `${i + 1} / ${totaal} records verwerkt`;

      try {
        await bcSyncKlacht({
          dossiernummer:  f.Dossiernummer || '',
          klantnummer:    f.Klantnummer   || '',
          datumMelding:   f.DatumMelding  || new Date().toISOString(),
          typeKlacht:     f.TypeKlacht    || '',
          omschrijving:   f.Omschrijving  || '',
          factuurnummer:  f.Factuurnummer || '',
          bedrag:         parseFloat(f.Bedrag) || 0,
          status:         f.Status        || 'Wachtend op goedkeuring',
          melder:         f.Melder        || '',
          melderNaam:     f.MelderNaam    || '',
          straat:         f.Straat        || '',
          postcode:       f.Postcode      || '',
          gemeente:       f.Gemeente      || '',
          leveradresCode: '',
          sharePointId:   String(all[i].id || ''),
        });
        geslaagd++;
      } catch (e) {
        if (e.message && e.message.includes('already exists')) {
          overgeslagen++;
        } else {
          mislukt++;
          console.warn(`BC sync mislukt voor ${f.Dossiernummer}:`, e.message);
        }
      }

      // Kleine pauze om BC niet te overbelasten
      if (i % 10 === 9) await new Promise(r => setTimeout(r, 500));
    }

    result.innerHTML = `
      <div class="alert alert-success">
        <strong>Synchronisatie voltooid</strong><br>
        ✅ ${geslaagd} gesynchroniseerd &nbsp;·&nbsp;
        ⏭ ${overgeslagen} reeds aanwezig &nbsp;·&nbsp;
        ❌ ${mislukt} mislukt
      </div>`;
    result.classList.remove('hidden');

  } catch (e) {
    result.innerHTML = `<div class="alert alert-error">Fout: ${esc(e.message)}</div>`;
    result.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg> Alle klachten naar BC synchroniseren`;
  }
}

/* ══════════════════ FORM ══════════════════ */
let selectedFiles = [];

function setupForm() {
  document.getElementById('btnSubmit').addEventListener('click', submitKlacht);
  addArtRow(); // start met 1 lege rij

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
  // Reset BC klantzoeker
  const bcZoek = document.getElementById('bcKlantZoek');
  if (bcZoek) bcZoek.value = '';
  const bcGes = document.getElementById('bcGeselecteerd');
  if (bcGes) bcGes.classList.add('hidden');
  const bcSug = document.getElementById('bcSuggesties');
  if (bcSug) bcSug.classList.add('hidden');

  ['fKlant','fKlantnr','fFactuurnr','fOmschrijving','fStraat','fPostcode','fGemeente'].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.value = ''; el.classList.remove('invalid'); }
  });
  document.getElementById('fType').value = '';
  selectedFiles = [];
  renderFileList();
  // Reset artikelregels
  document.getElementById('artBody').innerHTML = '';
  addArtRow();
  updateArtTotaal();
  hide('formSuccess');
  hide('formError');
}

/* ══════════════════ ARTIKELREGELS ══════════════════ */
/* ══════════════════ BC ARTIKELZOEKER ══════════════════ */
let artZoekTimers = {};

async function bcZoekArtikelen(zoekterm) {
  const tok       = await getBCToken();
  const companyId = await getBCCompanyId();
  const term      = zoekterm.replace(/'/g, "''");
  const select    = 'id,number,displayName,baseUnitOfMeasureCode,unitPrice,description';
  const base      = `${BC_BASE}/${BC_TENANT}/${BC_ENV}/api/v2.0/companies(${companyId})/items`;
  const headers   = { Authorization: `Bearer ${tok}` };

  const [r1, r2] = await Promise.all([
    fetch(`${base}?$filter=startswith(number,'${term}')&$top=8&$select=${select}`, { headers }),
    fetch(`${base}?$filter=contains(displayName,'${term}')&$top=8&$select=${select}`, { headers }),
  ]);

  const [d1, d2] = await Promise.all([
    r1.ok ? r1.json() : { value: [] },
    r2.ok ? r2.json() : { value: [] },
  ]);

  const seen = new Set();
  return [...d1.value, ...d2.value].filter(a => {
    if (seen.has(a.number)) return false;
    seen.add(a.number);
    return true;
  }).slice(0, 10);
}

function artZoekDebounce(rowId, waarde) {
  updateArtTotaal();
  clearTimeout(artZoekTimers[rowId]);
  const sug = document.getElementById(`art-sug-${rowId}`);
  if (waarde.trim().length < 2) { sug.classList.add('hidden'); return; }
  artZoekTimers[rowId] = setTimeout(() => artZoekVoerUit(rowId, waarde), 320);
}

async function artZoekVoerUit(rowId, waarde) {
  const sug = document.getElementById(`art-sug-${rowId}`);
  if (!sug) return;
  sug.innerHTML = '<div class="art-sug-item art-sug-leeg">Zoeken…</div>';
  sug.classList.remove('hidden');
  try {
    const artikelen = await bcZoekArtikelen(waarde.trim());
    if (!artikelen.length) {
      sug.innerHTML = '<div class="art-sug-item art-sug-leeg">Geen artikelen gevonden</div>';
      return;
    }
    sug.innerHTML = artikelen.map(a => `
      <div class="art-sug-item" onmousedown="artSelecteer(${rowId},${JSON.stringify(a).replace(/"/g,'&quot;')})">
        <div class="art-sug-nr">${esc(a.number)}</div>
        <div class="art-sug-naam">${esc(a.displayName)}</div>
        <div class="art-sug-adres">${esc(a.baseUnitOfMeasureCode||'')} · € ${(a.unitPrice||0).toFixed(2)}</div>
      </div>
    `).join('');
  } catch (e) {
    sug.innerHTML = '<div class="art-sug-item art-sug-leeg">BC niet beschikbaar</div>';
  }
}

function artSelecteer(rowId, artikel) {
  const tr   = document.getElementById(`art-row-${rowId}`);
  if (!tr) return;
  const set  = (field, val) => { const el = tr.querySelector(`[data-field="${field}"]`); if (el) { if (el.tagName === 'SELECT') { [...el.options].forEach(o => { if (o.value === val) o.selected = true; }); } else el.value = val; } };
  set('artnr', artikel.number);
  set('naam',  artikel.displayName);
  // prijs NIET automatisch invullen — melder vult zelf in
  // UOM matchen
  const uomMap = { 'ST':'ST','EA':'ST','DS':'DS','KG':'KG','LT':'LT','M2':'M2','PAL':'PAL' };
  const uom = uomMap[artikel.baseUnitOfMeasureCode] || 'ST';
  set('uom', uom);
  artSluitDropdown(rowId);
  updateArtTotaal();
}

function artSluitDropdown(rowId) {
  const sug = document.getElementById(`art-sug-${rowId}`);
  if (sug) sug.classList.add('hidden');
}

let artRowId = 0;

function addArtRow() {
  const id = ++artRowId;
  const tbody = document.getElementById('artBody');
  const tr = document.createElement('tr');
  tr.id = 'art-row-' + id;
  tr.innerHTML = `
    <td style="position:relative">
      <div class="art-zoek-wrap">
        <input class="art-input" placeholder="bijv. VB-4421"
          oninput="this.value=this.value.toUpperCase();artZoekDebounce(${id},this.value)"
          onblur="setTimeout(()=>artSluitDropdown(${id}),200)"
          data-field="artnr" autocomplete="off" />
        <div id="art-sug-${id}" class="art-suggesties hidden"></div>
      </div>
    </td>
    <td><input class="art-input" placeholder="Artikelnaam" oninput="updateArtTotaal()" data-field="naam" /></td>
    <td>
      <select class="art-select" onchange="updateArtTotaal()" data-field="uom">
        <option>ST</option><option>DS</option><option>KG</option><option>LT</option><option>M2</option><option>PAL</option>
      </select>
    </td>
    <td><input class="art-input" type="number" min="0" step="1" value="1" oninput="updateArtTotaal()" data-field="aantal" style="text-align:right" /></td>
    <td>
      <div style="display:flex;align-items:center;gap:3px">
        <span style="color:var(--muted);font-size:12px">€</span>
        <input class="art-input" type="number" min="0" step="0.01" value="0" oninput="updateArtTotaal()" data-field="prijs" style="text-align:right" />
      </div>
    </td>
    <td>
      <button type="button" class="art-del" onclick="delArtRow(${id})" title="Verwijderen">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>
      </button>
    </td>
  `;
  tbody.appendChild(tr);
  updateArtTotaal();
}

function delArtRow(id) {
  const row = document.getElementById('art-row-' + id);
  if (row) row.remove();
  updateArtTotaal();
}

function updateArtTotaal() {
  const rows = document.querySelectorAll('#artBody tr');
  let total = 0;
  rows.forEach(tr => {
    const aantal = parseFloat(tr.querySelector('[data-field="aantal"]')?.value) || 0;
    const prijs  = parseFloat(tr.querySelector('[data-field="prijs"]')?.value)  || 0;
    total += aantal * prijs;
  });
  const fmt = total.toLocaleString('nl-BE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const el1 = document.getElementById('artTotaal');
  const el2 = document.getElementById('artTotaalBadge');
  if (el1) el1.textContent = '€ ' + fmt;
  if (el2) el2.textContent = '€ ' + fmt;
}

function getArtRegels() {
  const rows = document.querySelectorAll('#artBody tr');
  return Array.from(rows).map(tr => ({
    artnr: tr.querySelector('[data-field="artnr"]')?.value.trim() || '',
    naam:  tr.querySelector('[data-field="naam"]')?.value.trim()  || '',
    uom:   tr.querySelector('[data-field="uom"]')?.value          || 'ST',
    aantal: parseFloat(tr.querySelector('[data-field="aantal"]')?.value) || 0,
    prijs:  parseFloat(tr.querySelector('[data-field="prijs"]')?.value)  || 0,
  }));
}

async function submitKlacht() {
  hide('formSuccess');
  hide('formError');

  // Validatie basisvelden
  const basicFields = {
    Klantnaam:     { id: 'fKlant',      val: document.getElementById('fKlant').value.trim() },
    Klantnummer:   { id: 'fKlantnr',    val: document.getElementById('fKlantnr').value.trim() },
    Factuurnummer: { id: 'fFactuurnr',  val: document.getElementById('fFactuurnr').value.trim() },
    TypeKlacht:    { id: 'fType',       val: document.getElementById('fType').value },
    Omschrijving:  { id: 'fOmschrijving', val: document.getElementById('fOmschrijving').value.trim() },
  };

  let valid = true;
  Object.values(basicFields).forEach(({ id, val }) => {
    const el = document.getElementById(id);
    if (!val) { el.classList.add('invalid'); valid = false; }
    else el.classList.remove('invalid');
  });

  // Validatie artikelregels
  const artikelregels = getArtRegels();
  const hasValidRow = artikelregels.some(r => r.artnr || r.naam);
  if (!hasValidRow) {
    valid = false;
    showError('formErrorMsg', 'Voeg minimaal 1 artikelregel toe.');
    show('formError');
  }

  if (!valid) {
    if (hasValidRow) {
      showError('formErrorMsg', 'Vul alle verplichte velden correct in.');
      show('formError');
    }
    return;
  }

  // Bereken totaalbedrag uit artikelregels
  const bedrag = artikelregels.reduce((sum, r) => sum + (r.aantal * r.prijs), 0);

  const fields = {
    Klantnaam:    basicFields.Klantnaam.val,
    Klantnummer:  basicFields.Klantnummer.val,
    Factuurnummer: basicFields.Factuurnummer.val,
    TypeKlacht:   basicFields.TypeKlacht.val,
    Omschrijving: basicFields.Omschrijving.val,
    Bedrag:       bedrag,
    Straat:       document.getElementById('fStraat')?.value.trim() || '',
    Postcode:     document.getElementById('fPostcode')?.value.trim() || '',
    Gemeente:     document.getElementById('fGemeente')?.value.trim() || '',
  };

  const btn = document.getElementById('btnSubmit');
  btn.disabled = true;
  btn.textContent = 'Bezig met indienen…';

  try {
    const dossiernummer = await generateDossierNumber();

    const spItem = await spCreateItem({
      ...fields,
      Dossiernummer: dossiernummer,
      Status:        'Wachtend op goedkeuring',
      DatumMelding:  new Date().toISOString(),
      Melder:        currentUser.email,
      MelderNaam:    currentUser.name,
      Artikelregels: JSON.stringify(artikelregels),
    });

    // BC sync – na SharePoint, zodat dataverlies onmogelijk is
    bcSyncKlacht({
      dossiernummer,
      klantnummer:    fields.Klantnummer,
      datumMelding:   new Date().toISOString(),
      typeKlacht:     fields.TypeKlacht,
      omschrijving:   fields.Omschrijving,
      factuurnummer:  fields.Factuurnummer,
      bedrag:         fields.Bedrag,
      straat:         fields.Straat,
      postcode:       fields.Postcode,
      gemeente:       fields.Gemeente,
      melder:         currentUser.email,
      melderNaam:     currentUser.name,
      leveradresCode: document.getElementById('bcLeveradresSelect')?.value || '',
      sharePointId:   String(spItem?.id || ''),
      status:         'Wachtend op goedkeuring',
    }).catch(e => console.warn('BC sync mislukt (niet kritiek):', e.message));

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

  // Deep-link: open dossier direct via ?dossier=2026-XXXX
  const urlParams = new URLSearchParams(window.location.search);
  const dossierParam = urlParams.get('dossier');
  if (dossierParam) {
    const match = allKlachten.find(k => k.Dossiernummer === dossierParam);
    if (match) openDetail(match.id);
  }
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

/* ══════════════════ RETOUR DOCUMENT ══════════════════ */
function printRetour(itemId) {
  var k = allKlachten.find(function(x){ return x.id === itemId; });
  if (!k) return;

  var artikelregels = [];
  try { artikelregels = JSON.parse(k.Artikelregels || '[]').filter(function(r){ return r.artnr || r.naam; }); } catch(e){}

  var totaal = artikelregels.reduce(function(s,r){ return s + (r.aantal * r.prijs); }, 0);
  var fmtTot = totaal.toLocaleString('nl-BE', {minimumFractionDigits:2, maximumFractionDigits:2});

  var qrUrl = 'https://verpa-klachten.pages.dev/?dossier=' + encodeURIComponent(k.Dossiernummer);
  var qrSrc = 'https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=' + encodeURIComponent(qrUrl);

  var artRows = artikelregels.map(function(r){
    var lijn = (r.aantal * r.prijs).toLocaleString('nl-BE', {minimumFractionDigits:2, maximumFractionDigits:2});
    return '<tr><td>' + esc(r.artnr) + '</td><td>' + esc(r.naam) + '</td><td style="text-align:center">' + esc(r.uom) + '</td><td style="text-align:right">' + r.aantal + '</td><td style="text-align:right">€ ' + r.prijs.toLocaleString('nl-BE',{minimumFractionDigits:2,maximumFractionDigits:2}) + '</td><td style="text-align:right">€ ' + lijn + '</td></tr>';
  }).join('');

  var datumFormatted = k.DatumMelding ? new Date(k.DatumMelding).toLocaleDateString('nl-BE') : '–';

  var html = `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8"/>
<title>Retour ${esc(k.Dossiernummer)}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Helvetica Neue', Arial, sans-serif; font-size: 12px; color: #111; background: #fff; padding: 28px 32px; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 24px; padding-bottom: 16px; border-bottom: 2px solid #1B3F6A; }
  .header-left h1 { font-size: 22px; font-weight: 800; color: #1B3F6A; letter-spacing: -.5px; }
  .header-left .sub { font-size: 11px; color: #64748B; margin-top: 2px; }
  .dossier-badge { background: #1B3F6A; color: #fff; font-size: 15px; font-weight: 700; padding: 6px 14px; border-radius: 6px; letter-spacing: .5px; }
  .section { margin-bottom: 20px; }
  .section-title { font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: .1em; color: #94A3B8; margin-bottom: 8px; }
  .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 24px; }
  .info-item label { font-size: 9px; text-transform: uppercase; letter-spacing: .07em; color: #94A3B8; display: block; margin-bottom: 2px; }
  .info-item span { font-size: 13px; font-weight: 600; color: #0F172A; }
  .omschrijving-box { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 10px 12px; font-size: 12px; color: #334155; line-height: 1.5; }
  table { width: 100%; border-collapse: collapse; font-size: 11.5px; }
  thead tr { background: #1B3F6A; color: #fff; }
  thead th { padding: 7px 10px; text-align: left; font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: .05em; }
  tbody tr:nth-child(even) { background: #F8FAFC; }
  tbody td { padding: 6px 10px; border-bottom: 1px solid #E2E8F0; }
  .totaal-row td { font-weight: 700; font-size: 13px; border-top: 2px solid #1B3F6A; border-bottom: none; padding-top: 8px; }
  .bottom { display: flex; gap: 24px; margin-top: 24px; padding-top: 16px; border-top: 1px solid #E2E8F0; }
  .sign-box { flex: 1; border: 1.5px dashed #CBD5E1; border-radius: 8px; padding: 12px 16px; min-height: 100px; }
  .sign-box .sign-label { font-size: 9px; text-transform: uppercase; letter-spacing: .08em; color: #94A3B8; font-weight: 700; margin-bottom: 4px; }
  .sign-box .sign-name { font-size: 11px; color: #64748B; margin-top: 6px; }
  .qr-box { display: flex; flex-direction: column; align-items: center; gap: 6px; }
  .qr-box img { width: 110px; height: 110px; }
  .qr-box .qr-label { font-size: 9px; color: #94A3B8; text-align: center; max-width: 110px; line-height: 1.4; }
  .type-pill { display: inline-block; background: #EBF3FA; color: #1B3F6A; font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 20px; }
  .footer { margin-top: 20px; font-size: 9px; color: #94A3B8; text-align: center; border-top: 1px solid #E2E8F0; padding-top: 10px; }
  @media print {
    body { padding: 16px 20px; }
    @page { margin: 12mm; }
  }
</style>
</head>
<body>
  <div class="header">
    <div class="header-left">
      <div style="background:#1B3F6A;border-radius:8px;padding:8px 16px;display:inline-block;margin-bottom:6px">
        <img src="https://verpa.be/wp-content/uploads/2023/03/cropped-Transparant-logo-Verpa_Lukas-1-2048x594.png" alt="Verpa" style="height:36px;display:block" />
      </div>
      <div class="sub">Verkoop Retour Verzending</div>
    </div>
    <div style="text-align:right">
      <div class="dossier-badge">${esc(k.Dossiernummer)}</div>
      <div style="font-size:10px;color:#64748B;margin-top:6px">Opgemaakt op ${new Date().toLocaleDateString('nl-BE')}</div>
    </div>
  </div>

  <div class="section" style="display:flex;gap:24px">
    <div style="flex:1">
      <div class="section-title">Klantgegevens</div>
      <div class="info-grid">
        <div class="info-item"><label>Klantnaam</label><span>${esc(k.Klantnaam)}</span></div>
        <div class="info-item"><label>Klantnummer</label><span>${esc(k.Klantnummer)}</span></div>
        ${k.Straat ? '<div class="info-item" style="grid-column:1/-1"><label>Adres</label><span>' + esc(k.Straat) + '<br>' + esc((k.Postcode||'') + ' ' + (k.Gemeente||'')).trim() + '<br>België</span></div>' : ''}
        <div class="info-item"><label>Factuurnummer</label><span>${esc(k.Factuurnummer)}</span></div>
        <div class="info-item"><label>Datum melding</label><span>${datumFormatted}</span></div>
        <div class="info-item"><label>Type klacht</label><span class="type-pill">${esc(k.TypeKlacht)}</span></div>
        <div class="info-item"><label>Ingediend door</label><span>${esc(k.MelderNaam||k.Melder)}</span></div>
      </div>
    </div>
    ${k.Straat ? `<div style="min-width:160px;background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;padding:14px 16px">
      <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#94A3B8;margin-bottom:8px">Retouradres klant</div>
      <div style="font-size:13px;font-weight:600;line-height:1.7;color:#0F172A">
        ${esc(k.Klantnaam)}<br>
        ${esc(k.Straat)}<br>
        ${esc((k.Postcode||'') + ' ' + (k.Gemeente||'')).trim()}<br>
        België
      </div>
    </div>` : ''}
  </div>

  <div class="section">
    <div class="section-title">Omschrijving</div>
    <div class="omschrijving-box">${esc(k.Omschrijving)}</div>
  </div>

  <div class="section">
    <div class="section-title">Te retourneren artikelen</div>
    <table>
      <thead><tr>
        <th>Artikelnr.</th><th>Artikelnaam</th><th style="text-align:center">UOM</th>
        <th style="text-align:right">Aantal</th><th style="text-align:right">Prijs/st.</th><th style="text-align:right">Totaal</th>
      </tr></thead>
      <tbody>
        ${artRows}
        <tr class="totaal-row">
          <td colspan="5" style="text-align:right">Totaal (excl. BTW)</td>
          <td style="text-align:right">€ ${fmtTot}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <div class="bottom">
    <div class="sign-box" style="flex:2">
      <div class="sign-label">Handtekening klant voor ontvangst retour</div>
      <div style="height:60px"></div>
      <div class="sign-name">Naam: _____________________________ &nbsp;&nbsp; Datum: _______________</div>
    </div>
    <div class="sign-box" style="flex:1.2">
      <div class="sign-label">Handtekening chauffeur</div>
      <div style="height:60px"></div>
      <div class="sign-name">Naam: _____________________________</div>
    </div>
    <div class="qr-box">
      <img src="${qrSrc}" alt="QR code dossier" />
      <div class="qr-label">Scan voor digitaal dossier ${esc(k.Dossiernummer)}</div>
    </div>
  </div>

  <div class="footer">
    Verpa Benelux NV &nbsp;·&nbsp; Laakdal &amp; Stora &nbsp;·&nbsp; verpabenelux.be &nbsp;·&nbsp; Dossier ${esc(k.Dossiernummer)}
  </div>

  <script>window.onload = function(){ window.print(); }<\/script>
</body>
</html>`;

  var win = window.open('', '_blank');
  win.document.write(html);
  win.document.close();
}
