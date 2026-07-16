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
  adminUsers: ['ils@verpa.be'],
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

let msalInstance;
let currentUser  = null;  // { name, email, isAdmin }
let allKlachten  = [];    // lokale cache
let currentRejectId = null;

/* ══════════════════ INIT ══════════════════ */
async function init() {
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Verwerk redirect na login
  const resp = await msalInstance.handleRedirectPromise().catch(console.error);
  if (resp) { await onSignedIn(resp.account); return; }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await onSignedIn(accounts[0]);
  } else {
    showLogin();
  }
}

function showLogin() {
  document.getElementById('loginScreen').classList.remove('hidden');
  document.getElementById('appShell').classList.add('hidden');
}

async function onSignedIn(account) {
  const token = await getToken(account);
  if (!token) { showLogin(); return; }

  // Haal gebruikersprofiel op
  const profile = await graphGet('/me', token);
  const email = (profile.mail || profile.userPrincipalName || '').toLowerCase();

  currentUser = {
    name:    profile.displayName || account.name || 'Gebruiker',
    email,
    isAdmin: CONFIG.adminUsers.includes(email),
    token,
  };

  // UI bijwerken
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('appShell').classList.remove('hidden');
  document.getElementById('userName').textContent   = currentUser.name.split(' ')[0];
  document.getElementById('userRole').textContent   = currentUser.isAdmin ? 'Beheerder' : 'Melder';
  document.getElementById('userAvatar').textContent = currentUser.name.charAt(0).toUpperCase();
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

  document.getElementById('btnLogin').addEventListener('click', () => {
    msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
  });
}

function navigateTo(view) {
  // Views
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  const target = document.getElementById(`view${capitalize(view)}`);
  if (target) target.classList.add('active');

  // Nav items
  document.querySelectorAll('.nav-item').forEach(n => {
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
  hide('tblKlachten');
  hide('emptyState');

  try {
    let filter = '';
    if (!currentUser.isAdmin) {
      filter = `fields/Melder eq '${currentUser.email}'`;
    }

    // Status filter (admin)
    const statusFilter = document.getElementById('filterStatus')?.value;
    if (statusFilter && currentUser.isAdmin) {
      const extra = `fields/Status eq '${statusFilter}'`;
      filter = filter ? `${filter} and ${extra}` : extra;
    }

    allKlachten = await spGetItems(filter);
    renderDashboard(allKlachten);
  } catch (e) {
    hide('loadingDash');
    showToast('Fout bij laden: ' + e.message, 'error');
  }
}

function renderDashboard(items) {
  hide('loadingDash');

  // Stats
  document.getElementById('statTotal').textContent    = items.length;
  document.getElementById('statPending').textContent  = items.filter(i => i.Status === 'Wachtend op goedkeuring').length;
  document.getElementById('statApproved').textContent = items.filter(i => i.Status === 'Goedgekeurd').length;
  document.getElementById('statRejected').textContent = items.filter(i => i.Status === 'Geweigerd').length;

  // Badge
  const pendingCount = items.filter(i => i.Status === 'Wachtend op goedkeuring').length;
  const badge = document.getElementById('badgePending');
  if (pendingCount > 0 && currentUser.isAdmin) {
    badge.textContent = pendingCount;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }

  if (!items.length) {
    show('emptyState');
    return;
  }

  const tbody = document.getElementById('tblBody');
  tbody.innerHTML = '';

  // Admin columns
  document.querySelectorAll('.data-table .admin-only').forEach(th => {
    th.classList.toggle('hidden', !currentUser.isAdmin);
  });

  items.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="dossier-nr">${esc(item.Dossiernummer)}</span></td>
      <td>${formatDate(item.DatumMelding)}</td>
      <td>${esc(item.Klantnaam)}</td>
      <td>${esc(item.TypeKlacht)}</td>
      <td>€ ${formatBedrag(item.Bedrag)}</td>
      <td>${statusPill(item.Status)}</td>
      <td>${esc(item.MelderNaam || item.Melder || '–')}</td>
      ${currentUser.isAdmin ? `
      <td>
        <div style="display:flex;gap:0.4rem;align-items:center">
          <button class="btn btn-sm btn-secondary" onclick="openDetail('${item.id}')">Detail</button>
          ${item.Status === 'Wachtend op goedkeuring' ? `
            <button class="btn btn-sm btn-success" onclick="approveKlacht('${item.id}')">✓ Goed</button>
            <button class="btn btn-sm btn-danger"  onclick="openReject('${item.id}')">✗ Weiger</button>
          ` : ''}
        </div>
      </td>` : '<td></td>'}
    `;
    tbody.appendChild(tr);
  });

  show('tblKlachten');
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
function openDetail(itemId) {
  const item = allKlachten.find(k => k.id === itemId);
  if (!item) return;

  document.getElementById('modalTitle').textContent = `Dossier ${item.Dossiernummer}`;

  document.getElementById('modalBody').innerHTML = `
    <div class="detail-grid">
      <div><div class="detail-label">Dossiernummer</div><div class="detail-value" style="font-family:monospace;font-weight:700">${esc(item.Dossiernummer)}</div></div>
      <div><div class="detail-label">Status</div><div class="detail-value">${statusPill(item.Status)}</div></div>
      <div><div class="detail-label">Datum melding</div><div class="detail-value">${formatDate(item.DatumMelding)}</div></div>
      <div><div class="detail-label">Type klacht</div><div class="detail-value">${esc(item.TypeKlacht)}</div></div>
      <div><div class="detail-label">Klantnaam</div><div class="detail-value">${esc(item.Klantnaam)}</div></div>
      <div><div class="detail-label">Klantnummer</div><div class="detail-value">${esc(item.Klantnummer)}</div></div>
      <div><div class="detail-label">Factuurnummer</div><div class="detail-value">${esc(item.Factuurnummer)}</div></div>
      <div><div class="detail-label">Bedrag (excl. BTW)</div><div class="detail-value big">€ ${formatBedrag(item.Bedrag)}</div></div>
      <div class="detail-full"><div class="detail-label">Omschrijving</div><div class="detail-value desc">${esc(item.Omschrijving)}</div></div>
      <div><div class="detail-label">Ingediend door</div><div class="detail-value">${esc(item.MelderNaam || item.Melder || '–')}</div></div>
      ${item.BeoordeeldDoor ? `<div><div class="detail-label">Beoordeeld door</div><div class="detail-value">${esc(item.BeoordeeldDoor)}</div></div>` : ''}
    </div>
    ${item.WeigeringReden ? `
      <div class="rejection-reason">
        <div class="detail-label">Reden weigering</div>
        <div class="detail-value">${esc(item.WeigeringReden)}</div>
      </div>` : ''}
  `;

  const footer = document.getElementById('modalFooter');
  footer.innerHTML = '';

  if (currentUser.isAdmin && item.Status === 'Wachtend op goedkeuring') {
    footer.innerHTML = `
      <button class="btn btn-success" onclick="approveKlacht('${item.id}')">✓ Goedkeuren</button>
      <button class="btn btn-danger"  onclick="openReject('${item.id}')">✗ Weigeren</button>
      <button class="btn btn-ghost modal-close">Sluiten</button>
    `;
    footer.querySelectorAll('.modal-close').forEach(b => b.addEventListener('click', closeModal));
  } else {
    footer.innerHTML = `<button class="btn btn-secondary modal-close">Sluiten</button>`;
    footer.querySelector('.modal-close').addEventListener('click', closeModal);
  }

  show('modalOverlay');
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
