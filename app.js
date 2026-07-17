/* ============================================================================
 *  SCHADEMELDINGEN — Verpa Benelux
 *  MSAL (delegated) + Microsoft Graph -> SharePoint-lijsten + fotobibliotheek
 * ============================================================================ */

const CONFIG = {
  clientId:    "b2bbb9ba-b6f1-4046-91fe-cc8fa5df7a7b",
  tenantId:    "e65dbe4b-d1e2-4283-b0f5-aa7717e81077",
  redirectUri: "https://verpa-schade.pages.dev",
  siteHostname:"verpabenelux.sharepoint.com",
  sitePath:    "/sites/OfficeData",
  listName:    "Schademeldingen",
  attachFolder:"Schadefotos",
  settingsList:"SchadeInstellingen",
  commentsList:"SchadeReacties",
  adminRole:   "Admin",
  scopes:      ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"],
  mailEnabled: false,
  demoMode:    !window.location.hostname.includes("pages.dev"),
};

const GRAPH = "https://graph.microsoft.com/v1.0";

/* ── Fasen + velden ─────────────────────────────────────── */
const COMMON_HEAD = [
  { name: "warehouse",     label: "Magazijn",          type: "select",   options: ["LAAKDAL", "STORA"], required: true },
  { name: "artikelnummer", label: "Artikelnummer",     type: "text",     placeholder: "bv. 100234", required: true, mono: true },
  { name: "omschrijving",  label: "Omschrijving",      type: "text",     placeholder: "Artikelomschrijving", full: true },
  { name: "aantal",        label: "Aantal beschadigd", type: "number",   placeholder: "0", required: true },
];
const COMMON_TAIL = [
  { name: "prioriteit",   label: "Prioriteit",   type: "select",   options: ["Laag", "Normaal", "Hoog"], required: true },
  { name: "aangemaakt",   label: "Datum melding", type: "date",    required: true },
  { name: "melder",       label: "Gemeld door",  type: "text",    placeholder: "Naam" },
  { name: "opmerkingen",  label: "Opmerkingen",  type: "textarea", placeholder: "Extra details…", full: true },
];
const STAGES = [
  { key: "inkomend", stap: "Ontvangst", label: "Inkomende goederen", sub: "Schade bij ontvangst", icon: "arrow-down-to-line",
    fields: [
      { name: "leverancier",    label: "Leverancier",          type: "text",   placeholder: "Naam leverancier" },
      { name: "inkooporder",    label: "Inkooporder-nr.",      type: "text",   placeholder: "BC inkooporder", mono: true },
      { name: "ontvangstdatum", label: "Ontvangstdatum",       type: "date" },
      { name: "vervoerder",     label: "Vervoerder",           type: "text",   placeholder: "Transporteur" },
      { name: "typeSchade",     label: "Soort schade",         type: "select", options: ["Transportschade","Verpakkingsschade","Productdefect","Nat / vochtig","Ontbrekend"] },
      { name: "actie",          label: "Actie bij ontvangst",  type: "select", options: ["Geweigerd","Onder voorbehoud aangenomen","Volledig aangenomen","Retour naar leverancier"] },
    ] },
  { key: "voorraad", stap: "Voorraad", label: "Schade aan voorraad", sub: "Tijdens opslag of orderverwerking", icon: "warehouse",
    fields: [
      { name: "locatie",        label: "Locatie / bin",        type: "text",   placeholder: "bv. A-12-03", mono: true },
      { name: "oorzaak",        label: "Oorzaak",              type: "select", options: ["Handling (val)","Heftruck / transpallet","THT verlopen","Waterschade","Mispick-schade","Onbekend"] },
      { name: "ontdektBij",     label: "Ontdekt bij",          type: "select", options: ["Orderpicking","Cyclustelling","Routine-inspectie","Bijvullen"] },
      { name: "gekoppeldeOrder",label: "Gekoppelde order",     type: "text",   placeholder: "Verkooporder tijdens picking", mono: true },
      { name: "afhandeling",    label: "Afhandeling",          type: "select", options: ["Afgeschreven","Afgeprijsd","Hersteld","In quarantaine"] },
    ] },
  { key: "uitgaand", stap: "Verzending", label: "Uitgaande goederen", sub: "Schade vóór of bij verzending", icon: "package-check",
    fields: [
      { name: "klant",          label: "Klant",                type: "text",   placeholder: "Naam klant" },
      { name: "verkooporder",   label: "Verkooporder-nr.",     type: "text",   placeholder: "BC verkooporder", mono: true },
      { name: "ontdektTijdens", label: "Ontdekt tijdens",      type: "select", options: ["Verpakken","Laden","Eindcontrole"] },
      { name: "vervoerder",     label: "Vervoerder",           type: "text",   placeholder: "Transporteur" },
      { name: "actie",          label: "Actie",                type: "select", options: ["Opnieuw picken","Vervangen","Order gesplitst","Verzending uitgesteld"] },
    ] },
  { key: "levering", stap: "Levering", label: "Schade bij levering", sub: "Na aflevering gemeld", icon: "truck",
    fields: [
      { name: "klant",          label: "Klant",                type: "text",   placeholder: "Naam klant" },
      { name: "verkooporder",   label: "Verkooporder-nr.",     type: "text",   placeholder: "BC verkooporder", mono: true },
      { name: "leveringsdatum", label: "Leveringsdatum",       type: "date" },
      { name: "vervoerder",     label: "Vervoerder",           type: "text",   placeholder: "Transporteur" },
      { name: "gemeldDoor",     label: "Gemeld door",          type: "select", options: ["Klant","Chauffeur","Vertegenwoordiger"] },
      { name: "claimVervoerder",label: "Claim bij vervoerder", type: "select", options: ["Ja","Nee","In behandeling"] },
      { name: "oplossing",      label: "Oplossing",            type: "select", options: ["Creditnota","Herlevering","Retour","Geen actie"] },
    ] },
];

const stageByKey = (k) => STAGES.find((s) => s.key === k);
const KEY_TO_FASE = { inkomend:"Ontvangst", voorraad:"Voorraad", uitgaand:"Verzending", levering:"Levering" };
const FASE_TO_KEY = Object.fromEntries(Object.entries(KEY_TO_FASE).map(([k,v])=>[v,k]));
const STATUS_FLOW = ["Nieuw","In behandeling","Afgehandeld"];

/* status chip CSS class mapping */
function statusChipClass(status) {
  if (status === "Nieuw")          return "status-chip status-chip--new";
  if (status === "In behandeling") return "status-chip status-chip--processing";
  if (status === "Afgehandeld")    return "status-chip status-chip--done";
  return "status-chip status-chip--new";
}
function statusIconName(status) {
  if (status === "Nieuw")          return "circle";
  if (status === "In behandeling") return "clock";
  if (status === "Afgehandeld")    return "check-circle-2";
  return "circle";
}

/* fase chip */
function faseChipClass(key) {
  return `fase-chip fase-chip--${key}`;
}

/* ── SharePoint kolomdefinities ─────────────────────────── */
const COLS = {
  prioriteit:{col:"Prioriteit"}, warehouse:{col:"Magazijn"}, artikelnummer:{col:"Artikelnummer"},
  omschrijving:{col:"Omschrijving"}, aantal:{col:"Aantal",kind:"number"}, melder:{col:"Melder"},
  opmerkingen:{col:"Opmerkingen"}, aangemaakt:{col:"Melddatum",kind:"date"}, fotos:{col:"Fotos",kind:"json"},
  leverancier:{col:"Leverancier"}, inkooporder:{col:"Inkooporder"},
  ontvangstdatum:{col:"Ontvangstdatum",kind:"date"}, vervoerder:{col:"Vervoerder"},
  typeSchade:{col:"SoortSchade"}, actie:{col:"Actie"}, locatie:{col:"Locatie"},
  oorzaak:{col:"Oorzaak"}, ontdektBij:{col:"OntdektBij"}, gekoppeldeOrder:{col:"GekoppeldeOrder"},
  afhandeling:{col:"Afhandeling"}, klant:{col:"Klant"}, verkooporder:{col:"Verkooporder"},
  ontdektTijdens:{col:"OntdektTijdens"}, leveringsdatum:{col:"Leveringsdatum",kind:"date"},
  gemeldDoor:{col:"GemeldDoor"}, claimVervoerder:{col:"ClaimVervoerder"}, oplossing:{col:"Oplossing"},
};

/* ── Helpers ─────────────────────────────────────────────── */
const $ = (sel, root=document) => root.querySelector(sel);
const todayISO = () => new Date().toISOString().slice(0,10);
const esc = (v) => String(v==null?"":v).replace(/[&<>"']/g,(c)=>({'&':"&amp;",'<':"&lt;",'>':'&gt;','"':"&quot;","'":'&#39;'}[c]));
const fmtDate = (iso) => { if(!iso) return "—"; try { return new Date(iso).toLocaleDateString("nl-BE",{day:"2-digit",month:"short",year:"numeric"}); } catch{return iso;} };
const fmtDateTime = (iso) => { if(!iso) return ""; try { return new Date(iso).toLocaleString("nl-BE",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"}); } catch{return iso;} };
const icons = () => window.lucide && window.lucide.createIcons();

/* ── MSAL ────────────────────────────────────────────────── */
function loadScript(src) {
  return new Promise((resolve,reject) => {
    const s = document.createElement("script"); s.src=src; s.async=true;
    s.onload=()=>resolve(src); s.onerror=()=>reject(new Error("kon niet laden: "+src));
    document.head.appendChild(s);
  });
}
async function loadMsal() {
  if (window.msal) return;
  const sources = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
    "https://alcdn.msauth.net/browser/2.38.4/js/msal-browser.min.js",
  ];
  for (const src of sources) {
    try { await loadScript(src); if (window.msal) return; } catch(e) { console.warn("MSAL-bron faalde:", src); }
  }
  throw new Error("MSAL kon niet geladen worden.");
}

let pca = null;
async function initAuth() {
  await loadMsal();
  pca = new msal.PublicClientApplication({
    auth: { clientId:CONFIG.clientId, authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri:CONFIG.redirectUri },
    cache: { cacheLocation:"localStorage", storeAuthStateInCookie:false },
  });
  const resp = await pca.handleRedirectPromise();
  if (resp && resp.account) pca.setActiveAccount(resp.account);
  const acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (acc) pca.setActiveAccount(acc);
  return acc || null;
}
function login() { pca.loginRedirect({ scopes:CONFIG.scopes }); }
function logout() { pca.logoutRedirect(); }
async function getToken() {
  const account = pca.getActiveAccount();
  const r = await pca.acquireTokenSilent({ scopes:CONFIG.scopes, account }).catch(async () => {
    await pca.acquireTokenRedirect({ scopes:CONFIG.scopes, account }); return null;
  });
  return r && r.accessToken;
}
function isAdmin() {
  const a = pca && pca.getActiveAccount();
  const roles = (a && a.idTokenClaims && a.idTokenClaims.roles) || [];
  return roles.includes(CONFIG.adminRole);
}

/* ── Graph datalaag ──────────────────────────────────────── */
async function graph(pathOrUrl, { method="GET", body, headers }={}) {
  const token = await getToken();
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH+pathOrUrl;
  const res = await fetch(url, {
    method,
    headers: { Authorization:`Bearer ${token}`, ...(body?{"Content-Type":"application/json"}:{}), ...(headers||{}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) throw new Error(`Graph ${method} ${res.status}: ${await res.text().catch(()=>"")}`);
  return res.status===204 ? null : res.json();
}

let _ctx=null, INTERNAL=null;
function internal(name) { return (INTERNAL && INTERNAL[name]) || name; }
async function getSiteId() {
  const site = await graph(`/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`);
  return site.id;
}
async function ensureList() {
  if (_ctx) return _ctx;
  const siteId = await getSiteId();
  const found = await graph(`/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(CONFIG.listName)}'`);
  if (!found.value||!found.value.length) throw new Error(`SharePoint-lijst "${CONFIG.listName}" niet gevonden.`);
  const listId = found.value[0].id;
  _ctx = { siteId, listId };
  const cols = await graph(`/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
  INTERNAL = {};
  (cols.value||[]).forEach((c)=>{ if(c.displayName) INTERNAL[c.displayName]=c.name; INTERNAL[c.name]=c.name; });
  return _ctx;
}
function encode(def,value) {
  if(def.kind==="number") return Number(value)||0;
  if(def.kind==="date")   return `${String(value).slice(0,10)}T00:00:00Z`;
  if(def.kind==="json")   return JSON.stringify(value||[]);
  return value;
}
function toFields(rec) {
  const f={};
  if(rec.id!=null) f[internal("Title")]=rec.id;
  if(rec.type) f[internal("Fase")]=KEY_TO_FASE[rec.type]||rec.type;
  if(rec.status) f[internal("Status")]=rec.status;
  for(const [k,def] of Object.entries(COLS)) {
    const v=rec[k]; if(v===undefined||v===null||v==="") continue;
    f[internal(def.col)]=encode(def,v);
  }
  return f;
}
function fromItem(item) {
  const f=item.fields||{};
  const gv=(n)=>f[internal(n)];
  const rec={spId:item.id, id:gv("Title"), status:gv("Status")||"Nieuw", type:FASE_TO_KEY[gv("Fase")]||gv("Fase")};
  for(const [k,def] of Object.entries(COLS)) {
    const raw=gv(def.col); if(raw===undefined||raw===null||raw==="") continue;
    if(def.kind==="date") rec[k]=String(raw).slice(0,10);
    else if(def.kind==="number") rec[k]=Number(raw);
    else if(def.kind==="json") { try{rec[k]=JSON.parse(raw);}catch{rec[k]=[];} }
    else rec[k]=raw;
  }
  return rec;
}
async function fetchReports() {
  const { siteId, listId } = await ensureList();
  const out=[]; let url=`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`;
  while(url) { const d=await graph(url); out.push(...(d.value||[]).map(fromItem)); url=d["@odata.nextLink"]||null; }
  return out.sort((a,b)=>(String(a.aangemaakt)<String(b.aangemaakt)?1:-1));
}
async function createReport(rec) {
  const { siteId, listId } = await ensureList();
  const created = await graph(`/sites/${siteId}/lists/${listId}/items`,{method:"POST",body:{fields:toFields(rec)}});
  return fromItem(created);
}
async function updateReport(spId, patch) {
  const { siteId, listId } = await ensureList();
  await graph(`/sites/${siteId}/lists/${listId}/items/${spId}/fields`,{method:"PATCH",body:toFields(patch)});
}
async function deleteReport(spId) {
  const { siteId, listId } = await ensureList();
  await graph(`/sites/${siteId}/lists/${listId}/items/${spId}`,{method:"DELETE"});
}

/* ── Instellingen ────────────────────────────────────────── */
let _settingsCtx=null;
async function ensureSettingsList() {
  if(_settingsCtx) return _settingsCtx;
  const siteId = await getSiteId();
  const found = await graph(`/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(CONFIG.settingsList)}'`);
  if(!found.value||!found.value.length) throw new Error(`Lijst "${CONFIG.settingsList}" niet gevonden.`);
  _settingsCtx = { siteId, listId:found.value[0].id };
  return _settingsCtx;
}
async function loadSettings() {
  try {
    const { siteId, listId } = await ensureSettingsList();
    const d = await graph(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=10`);
    const items=d.value||[];
    const item=items.find((it)=>(it.fields||{}).Title==="config")||items[0];
    if(!item) return { spId:null, enabled:true, recipients:[] };
    const f=item.fields||{}; let recipients=[];
    try{recipients=JSON.parse(f.Ontvangers||"[]");}catch{recipients=[];}
    return { spId:item.id, enabled:f.MeldingenAan!==false, recipients };
  } catch(e) { console.warn("Instellingen laden mislukt:",e); return { spId:null, enabled:true, recipients:[] }; }
}
async function saveSettings(settings) {
  const { siteId, listId } = await ensureSettingsList();
  const fields={ Title:"config", Ontvangers:JSON.stringify(settings.recipients||[]), MeldingenAan:!!settings.enabled };
  if(settings.spId) await graph(`/sites/${siteId}/lists/${listId}/items/${settings.spId}/fields`,{method:"PATCH",body:fields});
  else { const created=await graph(`/sites/${siteId}/lists/${listId}/items`,{method:"POST",body:{fields}}); settings.spId=created.id; }
  return settings;
}

/* ── E-mail ──────────────────────────────────────────────── */
function appCaseUrl(id) { return `${location.origin}${location.pathname}?case=${encodeURIComponent(id)}`; }
async function sendMail(recipients,subject,html) {
  const to=recipients.map((r)=>({emailAddress:{address:typeof r==="string"?r:r.email}}));
  await graph(`/me/sendMail`,{method:"POST",body:{message:{subject,body:{contentType:"HTML",content:html},toRecipients:to},saveToSentItems:true}});
}
async function notifyNewCase(rec) {
  if(!CONFIG.mailEnabled) return;
  const s=S.settings||{}; if(!s.enabled) return;
  const actief=(s.recipients||[]).filter((r)=>{
    if(!r||!r.email) return false;
    const fasen=Array.isArray(r.fasen)&&r.fasen.length?r.fasen:["inkomend","voorraad","uitgaand","levering"];
    return fasen.includes(rec.type);
  });
  if(!actief.length) return;
  await sendMail(actief,`Nieuwe schademelding ${rec.id} — ${stageByKey(rec.type).stap}`,buildEmailHtml(rec));
}
function buildEmailHtml(rec) {
  const stage=stageByKey(rec.type);
  const HEX={inkomend:["#0284c7","#e0f2fe","#0369a1"],voorraad:["#f59e0b","#fef3c7","#92400e"],uitgaand:["#7c3aed","#ede9fe","#6d28d9"],levering:["#e11d48","#ffe4e6","#be123c"]};
  const [main,soft,dark]=HEX[rec.type]||HEX.inkomend; const url=appCaseUrl(rec.id);
  const row=(label,val)=>val?`<tr><td style="padding:6px 0;color:#64748b;font-size:13px;width:150px;vertical-align:top;">${label}</td><td style="padding:6px 0;color:#0f172a;font-size:13px;font-weight:600;">${esc(val)}</td></tr>`:"";
  const extra=rec.klant?row("Klant",rec.klant):(rec.leverancier?row("Leverancier",rec.leverancier):"");
  const prio=rec.prioriteit==="Hoog"
    ?`<span style="display:inline-block;background:#ffe4e6;color:#be123c;font-size:12px;font-weight:700;padding:2px 8px;border-radius:9999px;">Hoog</span>`
    :`<span style="color:#0f172a;font-size:13px;font-weight:600;">${esc(rec.prioriteit||"Normaal")}</span>`;
  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f1f5f9;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;padding:24px 12px;"><tr><td align="center"><table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid #e2e8f0;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;"><tr><td style="background:${main};padding:20px 24px;"><div style="color:#fff;font-size:12px;font-weight:600;letter-spacing:.4px;text-transform:uppercase;opacity:.9;">Nieuwe schademelding</div><div style="color:#fff;font-size:20px;font-weight:700;margin-top:2px;">${esc(stage.stap)} · ${esc(rec.id)}</div></td></tr><tr><td style="padding:22px 24px 6px;"><span style="display:inline-block;background:${soft};color:${dark};font-size:12px;font-weight:700;padding:4px 10px;border-radius:9999px;">${esc(stage.label)}</span><p style="color:#334155;font-size:14px;line-height:1.5;margin:14px 0 8px;">Er is een nieuw schadegeval geregistreerd${rec.melder?` door ${esc(rec.melder)}`:""}.</p><table role="presentation" width="100%" cellpadding="0" cellspacing="0">${row("Magazijn",rec.warehouse)}${row("Artikel",[rec.artikelnummer,rec.omschrijving].filter(Boolean).join(" — "))}${row("Aantal",rec.aantal!=null?rec.aantal+" stuks":"")}${extra}<tr><td style="padding:6px 0;color:#64748b;font-size:13px;">Prioriteit</td><td style="padding:6px 0;">${prio}</td></tr>${row("Datum",fmtDate(rec.aangemaakt))}</table></td></tr><tr><td style="padding:10px 24px 26px;"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td style="border-radius:10px;background:${main};"><a href="${url}" style="display:inline-block;padding:12px 22px;color:#fff;font-size:14px;font-weight:600;text-decoration:none;border-radius:10px;">Bekijk in de app →</a></td></tr></table></td></tr><tr><td style="padding:14px 24px;background:#f8fafc;border-top:1px solid #e2e8f0;color:#94a3b8;font-size:12px;">Verpa Benelux · Schademeldingen</td></tr></table></td></tr></table></body></html>`;
}

/* ── Gesprek ─────────────────────────────────────────────── */
let _commentsCtx=null, INTERNAL_COMMENTS=null;
async function ensureCommentsList() {
  if(_commentsCtx) return _commentsCtx;
  const siteId=await getSiteId();
  const found=await graph(`/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(CONFIG.commentsList)}'`);
  if(!found.value||!found.value.length) throw new Error(`LIJST_ONTBREEKT:${CONFIG.commentsList}`);
  const listId=found.value[0].id;
  _commentsCtx={siteId,listId};
  const cols=await graph(`/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
  INTERNAL_COMMENTS={};
  (cols.value||[]).forEach((c)=>{ if(c.displayName) INTERNAL_COMMENTS[c.displayName]=c.name; INTERNAL_COMMENTS[c.name]=c.name; });
  return _commentsCtx;
}
function ic(name) { return (INTERNAL_COMMENTS&&INTERNAL_COMMENTS[name])||name; }
async function fetchComments(meldingId) {
  const { siteId, listId }=await ensureCommentsList();
  const d=await graph(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=999&$select=id,createdDateTime,fields`);
  return (d.value||[]).filter((it)=>String((it.fields||{}).Title||"")===String(meldingId))
    .map((it)=>{ const f=it.fields||{}; return { id:it.id, auteur:f[ic("Auteur")]||f.Auteur, bericht:f[ic("Bericht")]||f.Bericht, datum:it.createdDateTime }; })
    .sort((a,b)=>(String(a.datum)<String(b.datum)?-1:1));
}
async function addComment(meldingId,auteur,bericht) {
  const { siteId, listId }=await ensureCommentsList();
  const fields={ Title:String(meldingId) };
  fields[ic("Auteur")]=auteur; fields[ic("Bericht")]=bericht;
  const created=await graph(`/sites/${siteId}/lists/${listId}/items`,{method:"POST",body:{fields}});
  return { id:created.id, auteur, bericht, datum:created.createdDateTime||new Date().toISOString() };
}

/* ── Foto's ──────────────────────────────────────────────── */
const CHUNK=5*1024*1024;
const encPath=(p)=>p.split("/").map(encodeURIComponent).join("/");
async function uploadFotos(meldingId,files) {
  const { siteId }=await ensureList(); const refs=[];
  for(const file of files) {
    const path=`${CONFIG.attachFolder}/${meldingId}/${file.name}`;
    const session=await graph(`/sites/${siteId}/drive/root:/${encPath(path)}:/createUploadSession`,{method:"POST",body:{item:{"@microsoft.graph.conflictBehavior":"rename"}}});
    const uploadUrl=session.uploadUrl; const total=file.size; let start=0, item=null;
    do {
      const end=Math.min(start+CHUNK,total);
      const res=await fetch(uploadUrl,{method:"PUT",headers:{"Content-Range":`bytes ${start}-${end-1}/${total}`},body:file.slice(start,end)});
      if(res.status===200||res.status===201) item=await res.json(); else if(res.status!==202) throw new Error(`Foto-upload ${res.status}`);
      start=end;
    } while(start<total);
    refs.push({ name:file.name, driveId:item.parentReference.driveId, itemId:item.id });
  }
  return refs;
}
const fotoCache=new Map();
async function getFotoBlobUrl(ref) {
  if(ref.url) return ref.url;
  if(ref.itemId&&fotoCache.has(ref.itemId)) return fotoCache.get(ref.itemId);
  const token=await getToken();
  const res=await fetch(`${GRAPH}/drives/${ref.driveId}/items/${ref.itemId}/content`,{headers:{Authorization:`Bearer ${token}`}});
  if(!res.ok) throw new Error("Foto laden mislukt");
  const url=URL.createObjectURL(await res.blob());
  if(ref.itemId) fotoCache.set(ref.itemId,url);
  return url;
}

/* ── DEMO data ───────────────────────────────────────────── */
function demoReports() {
  return [
    { spId:"1", id:"SCH-2026-0007", type:"inkomend", status:"Nieuw", prioriteit:"Hoog", aangemaakt:"2026-07-16", warehouse:"LAAKDAL", artikelnummer:"VB-7012", omschrijving:"Reinigingsmiddel concentraat", aantal:3, leverancier:"Ecolab NV", vervoerder:"DHL", typeSchade:"Transportschade", melder:"Jan Pieters", fotos:[] },
    { spId:"2", id:"SCH-2026-0006", type:"voorraad", status:"In behandeling", prioriteit:"Normaal", aangemaakt:"2026-07-14", warehouse:"STORA", artikelnummer:"VB-4421", omschrijving:"Handzeep navulling 5L", aantal:8, locatie:"B-05-12", oorzaak:"Handling (val)", melder:"Sophie Maes", fotos:[] },
    { spId:"3", id:"SCH-2026-0005", type:"levering",  status:"Afgehandeld",    prioriteit:"Laag",   aangemaakt:"2026-07-11", warehouse:"LAAKDAL", artikelnummer:"VB-3301", omschrijving:"Vaatwasmiddel 10L", aantal:2, klant:"Colruyt NV", oplossing:"Creditnota", melder:"Tom Declercq", fotos:[] },
    { spId:"4", id:"SCH-2026-0004", type:"uitgaand",  status:"Afgehandeld",    prioriteit:"Normaal",aangemaakt:"2026-07-10", warehouse:"LAAKDAL", artikelnummer:"VB-9910", omschrijving:"Ontvetter spray 750ml", aantal:12, klant:"Makro Belgium", actie:"Vervangen", melder:"An Luyckx", fotos:[] },
  ];
}

/* ── STATE ───────────────────────────────────────────────── */
const S = {
  account:null, isAdmin:false, settings:null, reports:[], loading:true, error:null,
  view:{ name:"dashboard" },
  filters:{ q:"", status:"Alle", stage:"Alle", warehouse:"Alle", page:0 },
};
let formPhotos=[];

function go(view) { S.view=view; render(); window.scrollTo(0,0); }

function nextId() {
  const year=new Date().getFullYear();
  const nums=S.reports.map((r)=>r.id).filter((id)=>id&&id.startsWith(`SCH-${year}-`)).map((id)=>parseInt(id.split("-")[2],10));
  return `SCH-${year}-${String((nums.length?Math.max(...nums):0)+1).padStart(4,"0")}`;
}

/* ── LOGO base64 (Verpa) ─────────────────────────────────── */
const LOGO_B64 = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCABQAFADASIAAhEBAxEB/8QAHQAAAQQDAQEAAAAAAAAAAAAAAAMGBwgBAgUECf/EAD4QAAECBAMEBwQHCAMAAAAAAAECAwAEBREGEiEHMVFxCBMyQWGBsRUiNHMUQlKCkZKhFhcjM2JyosGz0dL/xAAbAQACAgMBAAAAAAAAAAAAAAAFBwQGAAIIAf/EADARAAEDAwEFCAIBBQAAAAAAAAECAwQAESEFMUFRcaEGEhMUMoGRsWHB0iJCQ+Hw/9oADAMBAAIRAxEAPwBxEDMdO+MWHCMntHnGjqsjS12vlSVW42F4UABJsKdVwBc1tYcILDhEffvIXlv7HRuv8Qf/ADDoxLXkUajNT5ZDrjqkpQ0VWvcXOvgIPSezGqRXW2XWrKcNki6TcjbsONu+g0ftDp0htx1ty6UC5wRb5GfauzYcILDhDQwzjGYrdXbkEUtDYUCpaw8TlSBvtbkPOHjlV9lX4RA1LTJOmOhmUnuqIva4OPYmpkDUI+oNl2ObpvbYRn3ArWw4QWHCMwQPqdWLDhBYcIzBGVlB7R5wlNfCvfLV6GFT2jzhKa+Fe+Wr0Mbt+sc60X6DUCHsfd/1Dr2j1D6TPykig+5KS6Mw/rUkE/paGqOyOUe+nSk9iCvy0hLAuz1QmUMt+K1kJHkPQR0vJhIdltSV/wCMKtzVYX+Afmuf2JamozjCf7ym/IXP3b4qwPRf2X0qr4XfxRiBmYdE28pqVaS+tpCmkaFRykE3XmG+3uxNJ2a4GLeT9mpH+6ys35r3/WPXKrw5s/wdISM7UZOm06Ql0S6HZh1LYVlHjvUdTprrHDY2zbMHnwwjGNPCibXWFpT+YpA/WKnI8WU8p5CSb8Bu3dKItr8JAQVbPzXhxBsiprjancPT8zTngLpZfWqYYV4e8c6eYVpwMRZVqfUKPU10yryipScQnME5syHEbs7avrJ/AjcQDFmJCclKhKNzcjNMTUu4LodZcC0KHEEaGOPjzC8niqiLknyGplu7kpMgXUw7bRXiDuUO8EiK1qOix5aTZISviMfPHnt+jYtL1+RDWAtRUjeDn4P62fdV3ggW2+y87LTbXUzLDimX273yOJNlC/eL7j3gg98ELlxtTSyhYsRg0zW3EuoC0G4ORQe0ecJTXwr3y1ehhU7zzhKa+Fe+Wr0MY36xWL9BqBRuHIRMHRipUmxXavjyrjLTMMyS3yq295STa3iEBXmpMQ8CAkE7gNYs3VcNOYO6IM7LuI6qeqKWJqc45nXm7JPJGVPkY6Y1FyyA1vWQPY7elc8xk3JVwF6gnaLjOsY5xI9Wau6rUkS0uFe5LN9yEj1O8nWG3m17WvOML7KiDrYxZbpD1fAVN2ZSuCJRiW9tyyJVcswyxZUqkhKisqtb3k3G8klWsbuOiOptpCLg4xuHHrXiUFwKWTsqMdgWP57BONJNhcyv2JPvpZnZcq9xOY5Q6B3KSSCSN4uD3Wu/HzYUSEqKd4BI5x9HaMt1ykya3r9aphsrvxyi8AtfZSlaHBtN7+1TYCiQU8KhTbTT0SOPPpTacqKlKJeV4uNnIo+aS3+EM2JJ6QKkGr0BItnDE0TyzM/7iNoTfaNATOURvAPS36pvdmVqVp6AdxI60HeecJTfwr3y1ehhU7zzhKb+Fe+Wr0MBW/WKOL9BqMdjtAl8TbSqDRpxaEyr0ylb2c2C0IBWUc1ZcvnFq+lHpsSrIG7rJb/nRFbtlNMU1TnatmWy+4tIl3UaKbyEELT45wD92Ji2sYzYxTsCrcvMlDFZlDLGalxoFjr0DrW+KD/idD3EvJ7V2ZmseWQctqA57Lkcjg8qSx0l6LpyZKhhYJ5cL8xkVVZfYVyMSR0kFA7V5lOYZvoElpfX+QmI3X2FcjH0VpUjJKk5SaVJy5fLDd3S0kr0SLa2vFg1KZ5RaHLXwR9UGjM+MlSb22fuqg7EdkVdxdXpOo1WnvyOHmHUuvPTCCgzIBv1bYOpvaxVuAvvMXQGgghkbUsbN4dkTTqc4hytzKP4KO0JdJ065Y4DuH1jpuuRVdQ1EyD4jv8ASlPQUXiRCCG2xdR61G+1iqoq2PZvqVBTFOaTJJIOhWCVunyJSnmgw140aQG2wgKUq2pUo3Uok3JJ7ySSSeJjeFFqMvzklb247OQwOlOHTYfk4qGeAzzOT1oO8840eSVsuIBsVIKQeYtG53nnBEIGxuKmEXFq81MlGpCnS8kyPcYbSgeNhv8AM6xzcb096pYefl5VgPTOZJbFwD2hfU+EduCJcac7Hlplg3WFd7O83vnnvqM/DaejKjEWSR3cbha2OW6ofODMSKSR7OtcW1dR/wBxbaS2uUmXprDJoNbU820lJADFrgAb+siKYItEztzPl28RCMfhX8qrTPYqC1furXn8j+NPqv7VcQVBpTFJkmKM2oWLy1h9/wC6LBCT4nNDEsouuPOOOPPPKzuuurK1uK+0pR1JjMEV2dqsmbh1WOAwP9+96OwdJiwctJzxOT/3KiCCCBtEq//Z";

/* ── SHELL ───────────────────────────────────────────────── */
function shell(content, activeTab) {
  const name = S.account ? esc(S.account.name || S.account.username || "") : "Demo";
  const pendingCount = S.reports.filter((r) => r.status === "Nieuw" || r.status === "In behandeling").length;
  const demoBar = CONFIG.demoMode
    ? `<div class="demo-bar">DEMO MODUS <span>— geen echte data, geen Microsoft-aanmelding vereist</span></div>` : "";

  const tabs = [
    { id:"dashboard",  label:"Dashboard" },
    { id:"form",       label:"Melding indienen" },
    { id:"pending",    label:"Te beoordelen", badge: pendingCount },
  ];

  const tabsHtml = tabs.map((t) => {
    const active = (activeTab === t.id || (t.id === "form" && activeTab === "picker")) ? "active" : "";
    const badge  = t.badge ? `<span class="nav-tab__badge">${t.badge}</span>` : "";
    return `<button class="nav-tab ${active}" data-tab="${t.id}">${esc(t.label)}${badge}</button>`;
  }).join("");

  return `${demoBar}
  <nav class="app-nav">
    <div class="app-nav__inner">
      <button class="app-nav__brand" data-nav="dashboard">
        <img src="${LOGO_B64}" alt="Verpa" class="app-nav__logo" />
        <span>
          <span class="app-nav__title">Schademeldingen</span>
          <span class="app-nav__sub">Verpa Benelux</span>
        </span>
      </button>
      <div class="app-nav__tabs--desktop">${tabsHtml}</div>
      <div class="app-nav__right">
        ${S.isAdmin ? `<span class="admin-badge">Beheerder</span>` : ""}
        <span class="nav-user">${name}</span>
        ${S.isAdmin && CONFIG.mailEnabled ? `<button class="nav-icon-btn" data-nav="settings" title="Instellingen"><i data-lucide="settings" class="w-4 h-4"></i></button>` : ""}
        <button class="nav-icon-btn" data-action="logout" title="Afmelden"><i data-lucide="log-out" class="w-4 h-4"></i></button>
      </div>
    </div>
    <div class="app-nav__tab-row">${tabsHtml}</div>
  </nav>
  <main class="app-main">${content}</main>`;
}

/* ── DASHBOARD ───────────────────────────────────────────── */
function viewDashboard() {
  const openCounts = {}; STAGES.forEach((s)=>(openCounts[s.key]=0));
  S.reports.forEach((r)=>{ if(r.status!=="Afgehandeld") openCounts[r.type]=(openCounts[r.type]||0)+1; });
  const open = S.reports.filter((r)=>r.status!=="Afgehandeld").length;
  const hoog = S.reports.filter((r)=>r.prioriteit==="Hoog"&&r.status!=="Afgehandeld").length;
  const done = S.reports.filter((r)=>r.status==="Afgehandeld").length;

  const cards = STAGES.map((s,i)=>{
    const count = openCounts[s.key];
    const badge = count > 0
      ? `<span class="stage-card__badge stage-card__badge--${s.key}">${count} open</span>`
      : `<span class="stage-card__badge" style="background:#f1f5f9;color:#94a3b8;">0 open</span>`;
    return `
      <button class="stage-card stage-card--${s.key}" data-newstage="${s.key}">
        <div class="stage-card__top">
          <span class="stage-card__icon"><i data-lucide="${s.icon}" class="w-5 h-5"></i></span>
          ${badge}
        </div>
        <div class="stage-card__label">Type ${i+1}</div>
        <div class="stage-card__title">${s.stap}</div>
        <div class="stage-card__sub">${s.sub}</div>
        <div class="stage-card__action"><i data-lucide="plus" class="w-3.5 h-3.5"></i> Melding indienen</div>
      </button>`;
  }).join("");

  const PAGE_SIZE=6;
  const all=filteredReports();
  const totalPages=Math.max(1,Math.ceil(all.length/PAGE_SIZE));
  const page=Math.min(S.filters.page||0,totalPages-1);
  const paged=all.slice(page*PAGE_SIZE,page*PAGE_SIZE+PAGE_SIZE);

  const rows = paged.map((r)=>{
    const stage=stageByKey(r.type);
    const beschrijving = [r.klant,r.leverancier,r.omschrijving].filter(Boolean).join(" · ");
    return `
      <button class="report-item-card report-item-card--${r.type}" data-open="${esc(r.id)}">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.625rem;">
          <div style="display:flex;align-items:center;gap:0.5rem;">
            <span class="report-card__dossier">${esc(r.id)}</span>
            <span class="${faseChipClass(r.type)}">${stage.stap}</span>
          </div>
          <span class="${statusChipClass(r.status)}">
            <i data-lucide="${statusIconName(r.status)}" class="w-3 h-3"></i>
            ${esc(r.status)}
          </span>
        </div>
        <div class="report-card__meta">
          <div class="report-card__field"><span class="report-card__field-label">Klant / Leverancier</span><span class="report-card__field-value">${esc(r.klant||r.leverancier||"—")}</span></div>
          <div class="report-card__field"><span class="report-card__field-label">Artikelnr.</span><span class="report-card__field-value mono">${esc(r.artikelnummer||"—")}</span></div>
          <div class="report-card__field"><span class="report-card__field-label">Aantal</span><span class="report-card__field-value">${esc(r.aantal||0)} st.</span></div>
          <div class="report-card__field"><span class="report-card__field-label">Ingediend door</span><span class="report-card__field-value">${esc(r.melder||"—")}</span></div>
          <div class="report-card__field"><span class="report-card__field-label">Datum</span><span class="report-card__field-value">${fmtDate(r.aangemaakt)}</span></div>
        </div>
        ${beschrijving?`<div class="report-card__desc">${esc(beschrijving)}</div>`:""}
      </button>`;
  }).join("");

  const pagination = totalPages>1 ? `
    <div class="pagination">
      <button class="btn-page" data-page="${page-1}" ${page===0?"disabled":""}>
        <i data-lucide="chevron-left" class="w-4 h-4"></i> Vorige
      </button>
      <span>Pagina ${page+1} van ${totalPages} <span style="color:#94a3b8;">(${all.length} meldingen)</span></span>
      <button class="btn-page" data-page="${page+1}" ${page>=totalPages-1?"disabled":""}>
        Volgende <i data-lucide="chevron-right" class="w-4 h-4"></i>
      </button>
    </div>` : "";

  const listBlock = all.length
    ? `<div class="report-card-list">${rows}${pagination}</div>`
    : `<div class="empty-state">
         <i data-lucide="clipboard-list" class="empty-state__icon"></i>
         <p class="empty-state__title">Geen meldingen gevonden</p>
         <p class="empty-state__sub">Pas de filters aan of maak een nieuwe melding via de fasen hierboven.</p>
       </div>`;

  const filterSel = (id,value,options,labelFn) =>
    `<select data-filter="${id}">${options.map((o)=>`<option value="${esc(o)}" ${o===value?"selected":""}>${esc(labelFn(o))}</option>`).join("")}</select>`;

  return shell(`
    <div class="page-header">
      <h1>Waar is de schade ontstaan?</h1>
      <p>Bekijk alle ingediende schademeldingen of <a href="#" data-tab="form" style="color:#3b82f6;font-weight:600;">maak een nieuwe melding</a>.</p>
    </div>
    <div class="stage-cards">${cards}</div>

    <div class="section-mt">
      <div class="section-header">
        <div class="section-title">
          <h2>Meldingen</h2>
          <div class="stat-dots">
            <span class="stat-dot stat-dot--open">${open} open</span>
            <span class="stat-dot stat-dot--hoog">${hoog} hoog</span>
            <span class="stat-dot stat-dot--done">${done} afgehandeld</span>
          </div>
        </div>
        <button class="btn-export" data-action="export">
          <i data-lucide="download" class="w-3.5 h-3.5"></i> Export CSV
        </button>
      </div>
      <div class="filters">
        <div class="filter-search">
          <i data-lucide="search" class="filter-search__icon"></i>
          <input data-filter="q" value="${esc(S.filters.q)}" placeholder="Zoek op nr., artikel, klant of leverancier…" />
        </div>
        <div class="filter-selects">
          ${filterSel("stage",S.filters.stage,["Alle",...STAGES.map((s)=>s.key)],(v)=>v==="Alle"?"Alle types":stageByKey(v).stap)}
          ${filterSel("status",S.filters.status,["Alle",...STATUS_FLOW],(v)=>v==="Alle"?"Alle statussen":v)}
          ${filterSel("warehouse",S.filters.warehouse,["Alle","LAAKDAL","STORA"],(v)=>v==="Alle"?"Alle magazijnen":v)}
        </div>
      </div>
      ${listBlock}
    </div>`, "dashboard");
}

/* ── TE BEOORDELEN (pending tab) ─────────────────────────── */
function viewPending() {
  const pending = S.reports.filter((r)=>r.status==="Nieuw"||r.status==="In behandeling");
  const cards = pending.map((r)=>{
    const stage = stageByKey(r.type);
    return `
      <div class="review-card">
        <div class="review-card__header">
          <span class="report-card__dossier">${esc(r.id)}</span>
          <span class="${statusChipClass(r.status)}">
            <i data-lucide="${statusIconName(r.status)}" class="w-3 h-3"></i>
            ${esc(r.status)}
          </span>
        </div>
        <div class="review-card__grid">
          <div class="review-card__field"><label>Klant / Leverancier</label><p>${esc(r.klant||r.leverancier||"—")}</p></div>
          <div class="review-card__field"><label>Klantnr.</label><p class="mono">${esc(r.artikelnummer||"—")}</p></div>
          <div class="review-card__field"><label>Fase</label><p><span class="${faseChipClass(r.type)}">${stage.stap}</span></p></div>
          <div class="review-card__field"><label>Aantal</label><p>${esc(r.aantal||0)} st.</p></div>
          <div class="review-card__field"><label>Ingediend door</label><p>${esc(r.melder||"—")}</p></div>
          <div class="review-card__field"><label>Datum</label><p>${fmtDate(r.aangemaakt)}</p></div>
        </div>
        ${r.opmerkingen ? `<div class="review-card__desc">${esc(r.opmerkingen)}</div>` : ""}
        <div class="review-card__footer">
          <button class="btn btn--approve" data-approve="${esc(r.id)}"><i data-lucide="check" class="w-3.5 h-3.5"></i> Afhandelen</button>
          <button class="btn btn--ghost" data-open="${esc(r.id)}">Detail</button>
        </div>
      </div>`;
  }).join("");

  const content = pending.length ? cards
    : `<div class="empty-state">
         <i data-lucide="check-circle-2" class="empty-state__icon" style="color:#10b981;"></i>
         <p class="empty-state__title">Alles beoordeeld</p>
         <p class="empty-state__sub">Er zijn geen meldingen meer die wachten op goedkeuring.</p>
       </div>`;

  return shell(`
    <div class="page-header">
      <h1>Te beoordelen</h1>
      <p>Meldingen met status "Nieuw" of "In behandeling".</p>
    </div>
    ${content}`, "pending");
}

/* ── FASE KIEZER ─────────────────────────────────────────── */
function viewStagePicker() {
  const iconStyle = {
    inkomend: "background:#f0f9ff;color:#0284c7;box-shadow:0 0 0 1px #bae6fd;",
    voorraad:  "background:#fffbeb;color:#d97706;box-shadow:0 0 0 1px #fde68a;",
    uitgaand:  "background:#f5f3ff;color:#7c3aed;box-shadow:0 0 0 1px #ddd6fe;",
    levering:  "background:#fff1f2;color:#e11d48;box-shadow:0 0 0 1px #fecdd3;",
  };
  const ctaColor = {
    inkomend: "#0284c7", voorraad: "#d97706", uitgaand: "#7c3aed", levering: "#e11d48",
  };

  const tiles = STAGES.map((s) => `
    <button class="fase-picker-card fase-picker-card--${s.key}" data-newstage="${s.key}">
      <span class="fase-picker-card__icon" style="${iconStyle[s.key]}">
        <i data-lucide="${s.icon}" class="w-6 h-6"></i>
      </span>
      <div>
        <div class="fase-picker-card__title">${s.stap}</div>
        <div class="fase-picker-card__sub">${s.label} — ${s.sub}</div>
      </div>
      <span class="fase-picker-card__cta" style="color:${ctaColor[s.key]};">
        <i data-lucide="plus" class="w-3.5 h-3.5"></i> Melding indienen
      </span>
    </button>`).join("");

  return shell(`
    <div class="page-header">
      <h1>Melding indienen</h1>
      <p>Kies de fase waarin de schade is ontstaan.</p>
    </div>
    <div class="fase-picker">${tiles}</div>`, "picker");
}

/* ── FORMULIER ───────────────────────────────────────────── */
function fieldHtml(f, value, stageKey) {
  const req = f.required ? `<span> *</span>` : "";
  let control;
  const cls = "field-control" + (f.mono ? " mono" : "");
  if(f.type==="select") {
    control=`<select data-field="${f.name}" class="${cls}"><option value="">— Kies —</option>${f.options.map((o)=>`<option ${o===value?"selected":""}>${esc(o)}</option>`).join("")}</select>`;
  } else if(f.type==="textarea") {
    control=`<textarea data-field="${f.name}" rows="3" placeholder="${esc(f.placeholder||"")}" class="${cls}">${esc(value||"")}</textarea>`;
  } else {
    const t=f.type==="number"?"number":f.type==="date"?"date":"text";
    control=`<input data-field="${f.name}" type="${t}" value="${esc(value||"")}" placeholder="${esc(f.placeholder||"")}" class="${cls}" />`;
  }
  return `<div class="field-group ${f.full?"field-group--full":""}"><label class="field-label">${esc(f.label)}${req}</label>${control}</div>`;
}

function viewForm(stageKey, editRecord) {
  const stage=stageByKey(stageKey);
  const fields=[...COMMON_HEAD,...stage.fields,...COMMON_TAIL];
  const isEdit=!!editRecord;
  const id=isEdit?editRecord.id:nextId();
  const defaults={};
  if(isEdit) { fields.forEach((f)=>{ defaults[f.name]=editRecord[f.name]??""; }); }
  else { fields.forEach((f)=>{ if(f.type==="date") defaults[f.name]=todayISO(); if(f.name==="prioriteit") defaults[f.name]="Normaal"; }); }

  // icon color per stage
  const iconStyle = {
    inkomend: "background:#f0f9ff;color:#0284c7;box-shadow:0 0 0 1px #bae6fd;",
    voorraad:  "background:#fffbeb;color:#d97706;box-shadow:0 0 0 1px #fde68a;",
    uitgaand:  "background:#f5f3ff;color:#7c3aed;box-shadow:0 0 0 1px #ddd6fe;",
    levering:  "background:#fff1f2;color:#e11d48;box-shadow:0 0 0 1px #fecdd3;",
  }[stageKey]||"";

  return shell(`
    <button class="back-link" data-nav="${isEdit?"detail:"+id:"picker"}">
      <i data-lucide="chevron-left" class="w-4 h-4"></i> ${isEdit?"Terug naar melding":"Terug"}
    </button>
    <div class="form-card">
      <div class="form-card__head">
        <span class="form-card__icon" style="${iconStyle}"><i data-lucide="${stage.icon}" class="w-6 h-6"></i></span>
        <div>
          <div class="form-card__eyebrow">${isEdit?"Melding bewerken":"Nieuwe schademelding"}</div>
          <div class="form-card__title">${stage.label}</div>
        </div>
        <span class="form-card__id">${id}</span>
      </div>
      <div class="form-card__body" id="formfields" data-stage="${stageKey}" data-id="${id}" data-edit="${isEdit?"1":"0"}" data-spid="${isEdit?esc(editRecord.spId||""):""}">
        ${fields.map((f)=>fieldHtml(f,defaults[f.name],stageKey)).join("")}
        <div class="field-group field-group--full">
          <label class="field-label">Foto's</label>
          <div style="display:flex;flex-wrap:wrap;gap:0.5rem;">
            <label class="btn btn--ghost" style="cursor:pointer;">
              <i data-lucide="camera" class="w-4 h-4"></i> Foto maken
              <input type="file" accept="image/*" capture="environment" class="hidden" data-addfiles style="display:none;" />
            </label>
            <label class="btn btn--ghost" style="cursor:pointer;">
              <i data-lucide="image-plus" class="w-4 h-4"></i> Uit galerij
              <input type="file" accept="image/*" multiple class="hidden" data-addfiles style="display:none;" />
            </label>
          </div>
          <div id="formthumbs" class="thumbs" style="margin-top:0.75rem;"></div>
        </div>
      </div>
      <div class="form-card__footer">
        <button class="btn btn--ghost" data-nav="${isEdit?"detail:"+id:"picker"}">Annuleren</button>
        <button class="btn btn--primary" data-action="${isEdit?"update":"save"}">${isEdit?"Wijzigingen opslaan":"Melding opslaan"}</button>
      </div>
    </div>`, "form");
}

function renderFormThumbs() {
  const box=$("#formthumbs"); if(!box) return;
  box.innerHTML=thumbsHtml(formPhotos,true); icons(); hydrateThumbs(box,formPhotos);
}

/* ── DETAIL ──────────────────────────────────────────────── */
function viewDetail(id) {
  const r=S.reports.find((x)=>x.id===id);
  if(!r) return shell(`<p style="color:#64748b;">Melding niet gevonden.</p>`, "dashboard");
  const stage=stageByKey(r.type);
  const detailFields=[...COMMON_HEAD,...stage.fields,...COMMON_TAIL];
  const defs = detailFields.map((f)=>{
    const v=r[f.name]; if(v===undefined||v===""||v===null) return "";
    const val=f.name==="aantal"?`${v} stuks`:esc(v);
    return `<dl class="detail-field ${f.full?"detail-field--full":""}"><dt>${esc(f.label)}</dt><dd class="${f.mono?"mono":""}">${val}</dd></dl>`;
  }).join("");
  const statusBtns = STATUS_FLOW.map((s)=>{
    const active=r.status===s;
    return `<button data-status="${s}" class="status-btn ${active?"active":""}">${s}</button>`;
  }).join("");
  const fotoBlock = (r.fotos&&r.fotos.length)
    ? `<dl class="detail-field detail-field--full"><dt>Foto's</dt><dd id="detailthumbs" class="thumbs" style="margin-top:0.5rem;"></dd></dl>` : "";

  const iconStyle = {
    inkomend:"background:#f0f9ff;color:#0284c7;box-shadow:0 0 0 1px #bae6fd;",
    voorraad: "background:#fffbeb;color:#d97706;box-shadow:0 0 0 1px #fde68a;",
    uitgaand: "background:#f5f3ff;color:#7c3aed;box-shadow:0 0 0 1px #ddd6fe;",
    levering: "background:#fff1f2;color:#e11d48;box-shadow:0 0 0 1px #fecdd3;",
  }[r.type]||"";

  return shell(`
    <button class="back-link" data-nav="dashboard"><i data-lucide="chevron-left" class="w-4 h-4"></i> Terug naar overzicht</button>
    <div class="detail-card">
      <div class="detail-card__head">
        <span class="form-card__icon" style="${iconStyle}"><i data-lucide="${stage.icon}" class="w-6 h-6"></i></span>
        <div style="min-width:0;">
          <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.2rem;">
            <span style="font-family:ui-monospace,monospace;font-size:0.75rem;color:#94a3b8;">${esc(r.id)}</span>
            <span class="${faseChipClass(r.type)}">${stage.stap}</span>
          </div>
          <div style="font-size:1rem;font-weight:700;color:#0f172a;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(r.omschrijving||r.artikelnummer||"")}</div>
        </div>
        <span class="${statusChipClass(r.status)}" style="margin-left:auto;">
          <i data-lucide="${statusIconName(r.status)}" class="w-3.5 h-3.5"></i> ${r.status}
        </span>
      </div>
      <div class="detail-card__status-bar">
        <span style="font-size:0.75rem;font-weight:600;color:#64748b;margin-right:0.25rem;">Status:</span>
        ${statusBtns}
      </div>
      <div class="detail-card__body">
        ${defs}
        ${fotoBlock}
      </div>
      <div class="detail-card__footer">
        ${S.isAdmin
          ? `<button class="btn btn--danger-ghost" data-action="delete" data-id="${esc(r.id)}"><i data-lucide="trash-2" class="w-4 h-4"></i> Verwijderen</button>
             <button class="btn btn--ghost" data-action="edit" data-id="${esc(r.id)}"><i data-lucide="pencil" class="w-4 h-4"></i> Bewerken</button>`
          : `<span style="display:inline-flex;align-items:center;gap:0.375rem;font-size:0.75rem;color:#94a3b8;"><i data-lucide="lock" class="w-3.5 h-3.5"></i> Alleen beheerders kunnen verwijderen</span>`}
      </div>
    </div>

    <div class="chat-card">
      <div class="chat-card__head">
        <i data-lucide="messages-square" class="w-4 h-4" style="color:#64748b;"></i> Gesprek
      </div>
      <div id="gesprek" class="chat-card__messages"></div>
      <div class="chat-card__compose">
        <textarea id="comment-input" rows="2" placeholder="Schrijf een bericht… (Ctrl/⌘ + Enter om te versturen)"></textarea>
        <div class="chat-card__compose-footer">
          <button class="btn btn--primary" data-comment-send>
            <i data-lucide="send" class="w-4 h-4"></i> Verstuur
          </button>
        </div>
      </div>
    </div>`, "dashboard");
}

/* ── GESPREK ─────────────────────────────────────────────── */
function commentListHtml(comments) {
  if(!comments.length) return `<p style="font-size:0.8125rem;color:#94a3b8;padding:0.5rem 0;">Nog geen berichten. Start het gesprek hieronder.</p>`;
  const me=(S.account&&(S.account.name||S.account.username))||"";
  return `<div style="display:flex;flex-direction:column;gap:0.75rem;">${comments.map((c)=>{
    const mine=c.auteur===me;
    const initials=(c.auteur||"?").trim().split(/\s+/).map((w)=>w[0]).slice(0,2).join("").toUpperCase();
    return `<div class="comment">
      <span class="comment__avatar ${mine?"comment__avatar--me":"comment__avatar--other"}">${esc(initials||"?")}</span>
      <div class="comment__body">
        <div class="comment__meta">
          <span class="comment__author">${esc(c.auteur||"Onbekend")}</span>
          <span class="comment__time">${fmtDateTime(c.datum)}</span>
        </div>
        <div class="comment__text">${esc(c.bericht||"")}</div>
      </div>
    </div>`;
  }).join("")}</div>`;
}

async function loadAndRenderComments(meldingId) {
  const box=$("#gesprek"); if(!box) return;
  box.innerHTML=`<div style="padding:1rem;text-align:center;color:#cbd5e1;"><i data-lucide="loader-2" class="w-5 h-5 spin" style="display:inline-block;"></i></div>`; icons();
  try {
    const comments=await fetchComments(meldingId);
    box.innerHTML=commentListHtml(comments); icons(); box.scrollTop=box.scrollHeight;
  } catch(e) {
    if(e.message&&e.message.startsWith("LIJST_ONTBREEKT")) {
      box.innerHTML=`<p style="font-size:0.8125rem;color:#94a3b8;">De lijst <strong>SchadeReacties</strong> bestaat nog niet in SharePoint.</p>`;
    } else {
      box.innerHTML=`<p style="font-size:0.8125rem;color:#ef4444;">Fout: ${esc(e.message||"onbekend")}</p>`;
    }
  }
}

async function postComment(meldingId) {
  const ta=$("#comment-input"); if(!ta) return;
  const text=ta.value.trim(); if(!text) { ta.focus(); return; }
  const btn=$("[data-comment-send]");
  const auteur=(S.account&&(S.account.name||S.account.username))||"Onbekend";
  if(btn) btn.disabled=true; ta.disabled=true;
  try {
    await addComment(meldingId,auteur,text); ta.value="";
    await loadAndRenderComments(meldingId);
  } catch(e) {
    if(e.message&&e.message.startsWith("LIJST_ONTBREEKT")) toast("Maak eerst de SharePoint-lijst 'SchadeReacties' aan");
    else toast("Bericht versturen mislukt");
  } finally { if(btn) btn.disabled=false; ta.disabled=false; ta.focus(); }
}

/* ── INSTELLINGEN ────────────────────────────────────────── */
function viewSettings() {
  if(!S.isAdmin) return shell(`<p style="color:#64748b;">Geen toegang. Deze pagina is enkel voor beheerders.</p>`, "dashboard");
  const s=S.settings||{enabled:true,recipients:[]};
  const FASEN=[{key:"inkomend",label:"Ontvangst",color:"#0284c7"},{key:"voorraad",label:"Voorraad",color:"#f59e0b"},{key:"uitgaand",label:"Verzending",color:"#7c3aed"},{key:"levering",label:"Levering",color:"#e11d48"}];
  const ALL_KEYS=FASEN.map((f)=>f.key);
  const rows=(s.recipients||[]).map((r,i)=>{
    const fasen=Array.isArray(r.fasen)&&r.fasen.length?r.fasen:ALL_KEYS;
    const faseBoxes=FASEN.map((f)=>`<label style="display:inline-flex;align-items:center;gap:0.25rem;font-size:0.75rem;cursor:pointer;">
      <input type="checkbox" data-recip-fase="${i}" data-fase="${f.key}" ${fasen.includes(f.key)?"checked":""} style="width:0.875rem;height:0.875rem;" />
      <span style="display:inline-flex;align-items:center;gap:0.25rem;color:#475569;">
        <span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${f.color};"></span>${f.label}
      </span></label>`).join("");
    return `<div class="recip-row">
      <div class="recip-row__inputs">
        <input data-recip-email="${i}" value="${esc(r.email||"")}" placeholder="naam@verpa.be" type="email" class="field-control" style="flex:1;" />
        <input data-recip-naam="${i}" value="${esc(r.naam||"")}" placeholder="Naam" class="field-control" style="width:8rem;" />
        <button data-recip-remove="${i}" class="nav-icon-btn" title="Verwijderen" style="color:#94a3b8;"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
      </div>
      <div class="recip-row__fasen">
        <span style="font-size:0.75rem;font-weight:600;color:#64748b;">Ontvangt:</span>${faseBoxes}
      </div>
    </div>`;
  }).join("");

  return shell(`
    <button class="back-link" data-nav="dashboard"><i data-lucide="chevron-left" class="w-4 h-4"></i> Terug</button>
    <div class="settings-card">
      <div class="settings-card__head">
        <div class="form-card__eyebrow">Beheer</div>
        <div class="form-card__title">E-mailinstellingen</div>
        <p style="font-size:0.8125rem;color:#64748b;margin-top:0.25rem;">Stel per ontvanger in voor welke fasen ze een e-mail ontvangen bij een nieuwe melding.</p>
      </div>
      <div class="settings-card__body">
        <label style="display:flex;align-items:center;gap:0.75rem;cursor:pointer;font-size:0.875rem;color:#374151;">
          <input type="checkbox" data-notify-toggle ${s.enabled?"checked":""} style="width:1rem;height:1rem;" />
          E-mailmeldingen inschakelen
        </label>
        <div>
          <div style="font-size:0.875rem;font-weight:600;color:#374151;margin-bottom:0.625rem;">Ontvangers</div>
          <div id="recip-list" style="display:flex;flex-direction:column;gap:0.5rem;">${rows||`<p style="font-size:0.8125rem;color:#94a3b8;">Nog geen ontvangers toegevoegd.</p>`}</div>
          <button data-recip-add class="btn btn--ghost" style="margin-top:0.75rem;">
            <i data-lucide="plus" class="w-4 h-4"></i> Ontvanger toevoegen
          </button>
        </div>
      </div>
      <div class="settings-card__footer">
        <button class="btn btn--primary" data-action="settings-save">Opslaan</button>
      </div>
    </div>`, "dashboard");
}

function syncRecipientsFromDom() {
  if(!S.settings) S.settings={enabled:true,recipients:[]};
  const list=[];
  document.querySelectorAll("[data-recip-email]").forEach((el)=>{
    const i=+el.getAttribute("data-recip-email");
    const naam=document.querySelector(`[data-recip-naam='${i}']`);
    const fasen=Array.from(document.querySelectorAll(`[data-recip-fase='${i}']:checked`)).map((cb)=>cb.getAttribute("data-fase"));
    list.push({email:el.value.trim(),naam:naam?naam.value.trim():"",fasen,actief:true});
  });
  S.settings.recipients=list;
  const t=document.querySelector("[data-notify-toggle]");
  if(t) S.settings.enabled=t.checked;
}

function wireSettings() {
  const ALL_KEYS=["inkomend","voorraad","uitgaand","levering"];
  const add=document.querySelector("[data-recip-add]");
  if(add) add.onclick=()=>{ syncRecipientsFromDom(); S.settings.recipients.push({email:"",naam:"",fasen:ALL_KEYS,actief:true}); render(); };
  document.querySelectorAll("[data-recip-remove]").forEach((b)=>b.onclick=()=>{
    syncRecipientsFromDom(); S.settings.recipients.splice(+b.getAttribute("data-recip-remove"),1); render();
  });
}

/* ── THUMBNAILS + VIEWER ─────────────────────────────────── */
function thumbsHtml(photos,removable) {
  if(!photos||!photos.length) return "";
  return photos.map((p,i)=>`
    <div class="thumb">
      <button data-thumb="${i}" style="display:grid;place-items:center;width:100%;height:100%;border:none;background:none;cursor:pointer;color:#cbd5e1;">
        <i data-lucide="loader-2" class="w-5 h-5 spin"></i>
      </button>
      ${removable?`<button data-removefoto="${i}" class="thumb__remove"><i data-lucide="x" class="w-3 h-3"></i></button>`:""}
    </div>`).join("");
}

async function hydrateThumbs(container,photos) {
  const btns=container.querySelectorAll("[data-thumb]");
  btns.forEach((b)=>b.onclick=()=>openThumb(b));
  container.querySelectorAll("[data-removefoto]").forEach((b)=>b.onclick=()=>{ formPhotos.splice(+b.getAttribute("data-removefoto"),1); renderFormThumbs(); });
  photos.forEach(async(p,i)=>{
    try {
      const url=await getFotoBlobUrl(p);
      const btn=btns[i]; if(!btn) return;
      btn.innerHTML=`<img src="${url}" alt="${esc(p.name)}" style="width:100%;height:100%;object-fit:cover;" />`;
    } catch(e) { const btn=btns[i]; if(btn) { btn.innerHTML=`<i data-lucide="image" class="w-5 h-5"></i>`; icons(); } }
  });
}

function openViewer(photos,startIndex) {
  let index=startIndex; const root=$("#overlay-root");
  root.innerHTML=`<div id="viewer" style="position:fixed;inset:0;z-index:40;display:flex;flex-direction:column;background:rgba(0,0,0,.92);">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:0.75rem 1rem;color:#fff;">
      <div style="min-width:0;"><div id="v-name" style="font-size:0.875rem;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"></div>${photos.length>1?`<div id="v-count" style="font-size:0.75rem;color:rgba(255,255,255,.6);"></div>`:""}</div>
      <div style="display:flex;gap:0.25rem;">
        <a id="v-download" href="#" download style="display:grid;place-items:center;width:2.25rem;height:2.25rem;border-radius:9999px;color:rgba(255,255,255,.8);"><i data-lucide="download" class="w-5 h-5"></i></a>
        <button data-vclose style="display:grid;place-items:center;width:2.25rem;height:2.25rem;border-radius:9999px;background:none;border:none;cursor:pointer;color:rgba(255,255,255,.8);"><i data-lucide="x" class="w-5 h-5"></i></button>
      </div>
    </div>
    <div id="v-stage" style="position:relative;flex:1;display:flex;align-items:center;justify-content:center;overflow:hidden;padding:0 0.5rem 1rem;">
      <i data-lucide="loader-2" class="w-8 h-8 spin" style="color:rgba(255,255,255,.6);"></i>
      ${photos.length>1?`
        <button data-vprev style="position:absolute;left:0.5rem;top:50%;transform:translateY(-50%);display:grid;place-items:center;width:2.75rem;height:2.75rem;border-radius:9999px;background:rgba(255,255,255,.12);border:none;cursor:pointer;color:#fff;"><i data-lucide="chevron-left" class="w-6 h-6"></i></button>
        <button data-vnext style="position:absolute;right:0.5rem;top:50%;transform:translateY(-50%);display:grid;place-items:center;width:2.75rem;height:2.75rem;border-radius:9999px;background:rgba(255,255,255,.12);border:none;cursor:pointer;color:#fff;"><i data-lucide="chevron-right" class="w-6 h-6"></i></button>`:""}
    </div>
  </div>`;
  icons();
  const stage=$("#v-stage"), nameEl=$("#v-name"), countEl=$("#v-count"), dl=$("#v-download");
  async function show() {
    nameEl.textContent=photos[index].name||"";
    if(countEl) countEl.textContent=`${index+1} / ${photos.length}`;
    [...stage.querySelectorAll("img")].forEach((n)=>n.remove());
    try {
      const url=await getFotoBlobUrl(photos[index]);
      const img=document.createElement("img");
      img.src=url; img.alt=photos[index].name||"";
      img.style.cssText="max-height:100%;max-width:100%;border-radius:0.5rem;object-fit:contain;user-select:none;";
      img.onclick=(e)=>e.stopPropagation();
      stage.insertBefore(img,stage.firstChild);
      dl.href=url; dl.setAttribute("download",photos[index].name||"foto");
    } catch(e){}
  }
  const fwd=(d)=>{ index=(index+d+photos.length)%photos.length; show(); };
  function close() { root.innerHTML=""; document.removeEventListener("keydown",onKey); }
  function onKey(e) { if(e.key==="Escape") close(); else if(e.key==="ArrowRight"&&photos.length>1) fwd(1); else if(e.key==="ArrowLeft"&&photos.length>1) fwd(-1); }
  document.addEventListener("keydown",onKey);
  const viewer=$("#viewer");
  viewer.addEventListener("click",(e)=>{ if(e.target.closest("#v-stage")&&!e.target.closest("button")&&!e.target.closest("a")&&e.target.tagName!=="IMG") close(); });
  $("[data-vclose]").onclick=close;
  const prev=$("[data-vprev]"), next=$("[data-vnext]");
  if(prev) prev.onclick=(e)=>{ e.stopPropagation(); fwd(-1); };
  if(next) next.onclick=(e)=>{ e.stopPropagation(); fwd(1); };
  let startX=null;
  viewer.addEventListener("touchstart",(e)=>(startX=e.touches[0].clientX),{passive:true});
  viewer.addEventListener("touchend",(e)=>{ if(startX==null) return; const dx=e.changedTouches[0].clientX-startX; if(Math.abs(dx)>50&&photos.length>1) fwd(dx<0?1:-1); startX=null; });
  show();
}

/* ── CSV EXPORT ──────────────────────────────────────────── */
function exportCSV() {
  const cols=["id","type","status","prioriteit","aangemaakt","warehouse","artikelnummer","omschrijving","aantal","melder"];
  const q=(v)=>`"${String(v==null?"":v).replace(/"/g,'""')}"`;
  const rows=filteredReports().map((r)=>{
    const stage=stageByKey(r.type);
    const details=stage.fields.map((f)=>(r[f.name]?`${f.label}: ${r[f.name]}`:"")).filter(Boolean).join(" | ");
    return [...cols.map((c)=>q(c==="type"?stage.label:r[c])),q(details)].join(",");
  });
  const csv="\uFEFF"+[...cols,"details"].join(",")+"\n"+rows.join("\n");
  const url=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"}));
  const a=document.createElement("a"); a.href=url; a.download=`schademeldingen_${todayISO()}.csv`; a.click();
  URL.revokeObjectURL(url);
}

/* ── TOAST ───────────────────────────────────────────────── */
function toast(msg) {
  const root=$("#overlay-root");
  const t=document.createElement("div"); t.className="toast";
  t.innerHTML=`<i data-lucide="check-circle-2" class="w-4 h-4" style="color:#34d399;"></i> ${esc(msg)}`;
  root.appendChild(t); icons();
  setTimeout(()=>t.remove(),2600);
}

/* ── RENDER + EVENTS ─────────────────────────────────────── */
function filteredReports() {
  const f=S.filters; const needle=f.q.trim().toLowerCase();
  return S.reports
    .filter((r)=>f.status==="Alle"||r.status===f.status)
    .filter((r)=>f.stage==="Alle"||r.type===f.stage)
    .filter((r)=>f.warehouse==="Alle"||r.warehouse===f.warehouse)
    .filter((r)=>!needle||[r.id,r.artikelnummer,r.omschrijving,r.klant,r.leverancier,r.verkooporder,r.inkooporder].filter(Boolean).join(" ").toLowerCase().includes(needle))
    .sort((a,b)=>(String(a.aangemaakt)<String(b.aangemaakt)?1:-1));
}

function render() {
  const app=$("#app");
  if(S.loading) { app.innerHTML=loadingScreen(); icons(); return; }
  if(S.error)   { app.innerHTML=errorScreen(S.error); icons(); wireEvents(); return; }
  if(S.view.name==="picker")    app.innerHTML=viewStagePicker();
  else if(S.view.name==="form")      app.innerHTML=viewForm(S.view.stageKey);
  else if(S.view.name==="edit") { const r=S.reports.find((x)=>x.id===S.view.id); app.innerHTML=r?viewForm(r.type,r):viewDashboard(); }
  else if(S.view.name==="detail")   app.innerHTML=viewDetail(S.view.id);
  else if(S.view.name==="pending")  app.innerHTML=viewPending();
  else if(S.view.name==="settings") app.innerHTML=viewSettings();
  else app.innerHTML=viewDashboard();
  icons();
  wireEvents();
  if(S.view.name==="settings") wireSettings();
  if(S.view.name==="form"||S.view.name==="edit") { formPhotos=formPhotos||[]; renderFormThumbs(); }
  if(S.view.name==="detail") {
    const r=S.reports.find((x)=>x.id===S.view.id);
    const box=$("#detailthumbs");
    if(r&&r.fotos&&r.fotos.length&&box) { box.innerHTML=thumbsHtml(r.fotos,false); icons(); hydrateThumbs(box,r.fotos); }
    if(r) {
      loadAndRenderComments(r.id);
      const sendBtn=$("[data-comment-send]"); const ta=$("#comment-input");
      if(sendBtn) sendBtn.onclick=(e)=>{ e.stopPropagation(); postComment(r.id); };
      if(ta) {
        ta.onkeydown=(e)=>{ if((e.ctrlKey||e.metaKey)&&e.key==="Enter"){e.preventDefault();postComment(r.id);} };
        ta.onclick=(e)=>e.stopPropagation();
      }
    }
  }
}

function loadingScreen() {
  return `<div class="fullpage"><div style="text-align:center;color:#94a3b8;"><i data-lucide="loader-2" class="w-8 h-8 spin" style="display:block;margin:0 auto;"></i><p style="margin-top:0.75rem;font-size:0.875rem;">Laden…</p></div></div>`;
}
function loginScreen() {
  return `<div class="fullpage"><div class="card-center">
    <img src="${LOGO_B64}" alt="Verpa" style="width:4rem;height:4rem;border-radius:0.75rem;object-fit:cover;margin:0 auto;display:block;" />
    <h1>Schademeldingen</h1>
    <p>Magazijn · Verpa Benelux</p>
    <button class="btn btn--primary" data-action="login" style="width:100%;margin-top:1.5rem;justify-content:center;">Aanmelden met Microsoft</button>
  </div></div>`;
}
function errorScreen(msg) {
  return `<div class="fullpage"><div class="card-center" style="border-color:#fecaca;">
    <span style="display:grid;place-items:center;width:3rem;height:3rem;border-radius:0.75rem;background:#fee2e2;color:#dc2626;margin:0 auto;"><i data-lucide="alert-triangle" class="w-6 h-6"></i></span>
    <h1>Er ging iets mis</h1>
    <p style="word-break:break-words;">${esc(msg)}</p>
    <button class="btn btn--primary" data-action="reload" style="margin-top:1.5rem;width:100%;justify-content:center;">Opnieuw proberen</button>
  </div></div>`;
}
function setupScreen() {
  return `<div class="fullpage"><div class="card-center" style="border-color:#fde68a;max-width:26rem;text-align:left;">
    <h1>Nog instellen</h1>
    <p style="margin-top:0.5rem;">Vul bovenaan in <code style="background:#f1f5f9;padding:0.1rem 0.3rem;border-radius:0.25rem;">app.js</code> de <code style="background:#f1f5f9;padding:0.1rem 0.3rem;border-radius:0.25rem;">CONFIG</code> in: clientId, tenantId, siteHostname en sitePath.</p>
  </div></div>`;
}

function wireEvents() {
  /* nav tabs */
  document.querySelectorAll("[data-tab]").forEach((b)=>b.onclick=()=>{
    const tab=b.getAttribute("data-tab");
    if(tab==="form")    { go({ name:"picker" }); }
    else if(tab==="pending") go({ name:"pending" });
    else go({ name:"dashboard" });
  });

  document.querySelectorAll("[data-nav]").forEach((b)=>b.onclick=()=>{
    const nav=b.getAttribute("data-nav");
    if(nav.startsWith("detail:")) go({ name:"detail", id:nav.split(":")[1] });
    else if(nav==="picker") go({ name:"picker" });
    else go({ name:nav });
  });
  document.querySelectorAll("[data-newstage]").forEach((b)=>b.onclick=()=>{ formPhotos=[]; go({ name:"form", stageKey:b.getAttribute("data-newstage") }); });
  document.querySelectorAll("[data-open]").forEach((b)=>b.onclick=()=>go({ name:"detail", id:b.getAttribute("data-open") }));
  document.querySelectorAll("[data-status]").forEach((b)=>b.onclick=()=>changeStatus(S.view.id,b.getAttribute("data-status")));
  document.querySelectorAll("[data-thumb]").forEach((b)=>b.onclick=()=>openThumb(b));
  document.querySelectorAll("[data-removefoto]").forEach((b)=>b.onclick=()=>{ formPhotos.splice(+b.getAttribute("data-removefoto"),1); renderFormThumbs(); });
  document.querySelectorAll("[data-addfiles]").forEach((inp)=>inp.onchange=(e)=>{ addFiles(e.target.files); e.target.value=""; });

  /* afhandelen in pending tab */
  document.querySelectorAll("[data-approve]").forEach((b)=>b.onclick=()=>changeStatus(b.getAttribute("data-approve"),"Afgehandeld"));

  document.querySelectorAll("[data-filter]").forEach((el)=>{
    const key=el.getAttribute("data-filter");
    if(key==="q") el.oninput=()=>{ S.filters.q=el.value; S.filters.page=0; refreshList(); };
    else el.onchange=()=>{ S.filters[key]=el.value; S.filters.page=0; refreshList(); };
  });
  document.querySelectorAll("[data-page]").forEach((b)=>{
    if(!b.disabled) b.onclick=()=>{ S.filters.page=+b.getAttribute("data-page"); render(); window.scrollTo(0,0); };
  });

  /* inline link in page-header */
  document.querySelectorAll("a[data-tab]").forEach((a)=>{
    a.onclick=(e)=>{ e.preventDefault(); const tab=a.getAttribute("data-tab"); if(tab==="form") { formPhotos=[]; go({ name:"form", stageKey:"inkomend" }); } };
  });

  const actions = {
    login, logout, reload:()=>location.reload(),
    export: exportCSV,
    save:   saveReport,
    update: updateReportFromForm,
    delete: (b)=>confirmDelete(b.getAttribute("data-id")),
    edit:   (b)=>{ formPhotos=[]; go({ name:"edit", id:b.getAttribute("data-id") }); },
    "settings-save": saveSettingsFromUI,
  };
  document.querySelectorAll("[data-action]").forEach((b)=>b.onclick=()=>actions[b.getAttribute("data-action")]&&actions[b.getAttribute("data-action")](b));
}

async function saveSettingsFromUI() {
  syncRecipientsFromDom();
  S.settings.recipients=(S.settings.recipients||[]).filter((r)=>r.email);
  const btn=document.querySelector("[data-action='settings-save']");
  if(btn){btn.disabled=true;btn.textContent="Opslaan…";}
  try { await saveSettings(S.settings); toast("Instellingen opgeslagen"); go({ name:"dashboard" }); }
  catch(e) { console.error(e); if(btn){btn.disabled=false;btn.textContent="Opslaan";} toast("Opslaan mislukt"); }
}

function refreshList() {
  if(S.view.name!=="dashboard") return;
  const hadFocus=document.activeElement&&document.activeElement.getAttribute&&document.activeElement.getAttribute("data-filter")==="q";
  const caret=hadFocus?document.activeElement.selectionStart:null;
  render();
  if(hadFocus) { const q=$("[data-filter='q']"); if(q){ q.focus(); if(caret!=null) q.setSelectionRange(caret,caret); } }
}

function openThumb(btn) {
  const idx=+btn.getAttribute("data-thumb");
  const photos=S.view.name==="form"?formPhotos:(S.reports.find((r)=>r.id===S.view.id)?.fotos||[]);
  if(photos.length) openViewer(photos,idx);
}

function addFiles(fileList) {
  const items=Array.from(fileList||[]).map((file)=>({name:file.name,url:URL.createObjectURL(file),file}));
  formPhotos.push(...items); renderFormThumbs();
}

async function saveReport() {
  const box=$("#formfields"); if(!box) return;
  const stageKey=box.getAttribute("data-stage"); const id=box.getAttribute("data-id");
  const stage=stageByKey(stageKey);
  const fields=[...COMMON_HEAD,...stage.fields,...COMMON_TAIL];
  const data={}; box.querySelectorAll("[data-field]").forEach((el)=>{ data[el.getAttribute("data-field")]=el.value; });
  let bad=null;
  box.querySelectorAll("[data-field]").forEach((el)=>el.classList.remove("error"));
  fields.forEach((f)=>{ if(f.required&&!String(data[f.name]||"").trim()){ const el=box.querySelector(`[data-field='${f.name}']`); if(el) el.classList.add("error"); if(!bad) bad=el; } });
  if(bad) { bad.scrollIntoView({behavior:"smooth",block:"center"}); return; }
  const btn=$("[data-action='save']"); btn.disabled=true; btn.textContent=formPhotos.length?"Foto's uploaden…":"Opslaan…";
  try {
    let fotoRefs=[];
    if(formPhotos.length) fotoRefs=await uploadFotos(id,formPhotos.map((p)=>p.file));
    const rec=await createReport({...data,id,type:stageKey,status:"Nieuw",aangemaakt:data.aangemaakt||todayISO(),aantal:Number(data.aantal)||0,fotos:fotoRefs});
    S.reports.unshift(rec); formPhotos=[];
    notifyNewCase(rec).catch((e)=>console.warn("E-mailmelding niet verstuurd:",e));
    toast(`Melding ${rec.id} opgeslagen`); go({ name:"detail", id:rec.id });
  } catch(e) { btn.disabled=false; btn.textContent="Melding opslaan"; toast("Opslaan mislukt — probeer opnieuw"); console.error(e); }
}

async function updateReportFromForm() {
  const box=$("#formfields"); if(!box) return;
  const stageKey=box.getAttribute("data-stage"); const id=box.getAttribute("data-id"); const spId=box.getAttribute("data-spid");
  const stage=stageByKey(stageKey);
  const fields=[...COMMON_HEAD,...stage.fields,...COMMON_TAIL];
  const data={}; box.querySelectorAll("[data-field]").forEach((el)=>{ data[el.getAttribute("data-field")]=el.value; });
  let bad=null;
  box.querySelectorAll("[data-field]").forEach((el)=>el.classList.remove("error"));
  fields.forEach((f)=>{ if(f.required&&!String(data[f.name]||"").trim()){ const el=box.querySelector(`[data-field='${f.name}']`); if(el) el.classList.add("error"); if(!bad) bad=el; } });
  if(bad) { bad.scrollIntoView({behavior:"smooth",block:"center"}); return; }
  const btn=$("[data-action='update']"); if(btn){btn.disabled=true;btn.textContent="Opslaan…";}
  try {
    const patch={...data,aantal:Number(data.aantal)||0,aangemaakt:data.aangemaakt||todayISO()};
    await updateReport(spId,patch);
    const idx=S.reports.findIndex((r)=>r.id===id); if(idx!==-1) S.reports[idx]={...S.reports[idx],...patch};
    toast("Wijzigingen opgeslagen"); go({ name:"detail", id });
  } catch(e) { if(btn){btn.disabled=false;btn.textContent="Wijzigingen opslaan";} toast("Opslaan mislukt"); console.error(e); }
}

async function changeStatus(id,status) {
  const r=S.reports.find((x)=>x.id===id); if(!r) return;
  const prev=r.status; r.status=status; render();
  try { await updateReport(r.spId,{status}); } catch(e) { r.status=prev; render(); toast("Status bijwerken mislukt"); }
}

function confirmDelete(id) {
  if(!S.isAdmin) { toast("Alleen beheerders kunnen verwijderen"); return; }
  const root=$("#overlay-root");
  root.innerHTML=`<div style="position:fixed;inset:0;z-index:50;display:grid;place-items:center;background:rgba(0,0,0,.4);padding:1.5rem;">
    <div style="width:100%;max-width:22rem;background:#fff;border-radius:1rem;padding:1.5rem;box-shadow:0 20px 60px rgba(0,0,0,.2);">
      <h3 style="font-size:0.9375rem;font-weight:700;color:#0f172a;">Melding verwijderen?</h3>
      <p style="margin-top:0.375rem;font-size:0.8125rem;color:#64748b;">${esc(id)} wordt definitief uit SharePoint verwijderd.</p>
      <div style="margin-top:1.25rem;display:flex;justify-content:flex-end;gap:0.5rem;">
        <button data-cancel class="btn btn--ghost">Annuleren</button>
        <button data-confirm class="btn btn--reject">Verwijderen</button>
      </div>
    </div></div>`;
  $("[data-cancel]").onclick=()=>(root.innerHTML="");
  $("[data-confirm]").onclick=async()=>{
    root.innerHTML="";
    const r=S.reports.find((x)=>x.id===id); if(!r) return;
    try { await deleteReport(r.spId); S.reports=S.reports.filter((x)=>x.id!==id); toast("Melding verwijderd"); go({ name:"dashboard" }); }
    catch(e) { toast("Verwijderen mislukt"); }
  };
}

/* ── INIT ────────────────────────────────────────────────── */
(async function init() {
  render(); // loading screen
  if(!CONFIG.clientId||!CONFIG.tenantId) { S.loading=false; $("#app").innerHTML=setupScreen(); icons(); return; }

  /* Demo mode: skip MSAL, use dummy data */
  if(CONFIG.demoMode) {
    S.account={ name:"IIs", username:"ils@verpa.be" };
    S.isAdmin=true;
    S.reports=demoReports();
    S.settings={ spId:null, enabled:false, recipients:[] };
    const caseId=new URLSearchParams(location.search).get("case");
    if(caseId&&S.reports.find((r)=>r.id===caseId)) S.view={ name:"detail", id:caseId };
    S.loading=false; render(); return;
  }

  try {
    S.account=await initAuth();
    if(!S.account) { S.loading=false; $("#app").innerHTML=loginScreen(); icons(); wireEvents(); return; }
    S.isAdmin=isAdmin();
    S.reports=await fetchReports();
    S.settings=await loadSettings().catch(()=>({ spId:null, enabled:false, recipients:[] }));
    const caseId=new URLSearchParams(location.search).get("case");
    if(caseId&&S.reports.find((r)=>r.id===caseId)) S.view={ name:"detail", id:caseId };
    S.loading=false; render();
  } catch(e) {
    console.error(e); S.loading=false; S.error=e.message||String(e); render();
  }
})();
