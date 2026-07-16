# Verpa Klachtenmeldingen – Setup Instructies

## 1. Azure App Registration

1. Ga naar **portal.azure.com** → Azure Active Directory → App Registrations → New Registration
2. Naam: `Verpa Klachten App`
3. Redirect URI: `https://verpa-klachten.pages.dev` (Single-page application)
4. Kopieer de **Client ID** en zet in `app.js` bij `CONFIG.clientId`
5. **API Permissions** (Delegated):
   - `User.Read`
   - `Sites.ReadWrite.All`
6. Grant admin consent

---

## 2. SharePoint Lijst aanmaken

Maak op de Verpa SharePoint site een lijst genaamd **`KlachtenMeldingen`** aan met de volgende kolommen:

| Kolomnaam           | Type                | Verplicht |
|---------------------|---------------------|-----------|
| `Dossiernummer`     | Tekst (één regel)   | Ja        |
| `DatumMelding`      | Datum en tijd       | Ja        |
| `Klantnaam`         | Tekst (één regel)   | Ja        |
| `Klantnummer`       | Tekst (één regel)   | Ja        |
| `Factuurnummer`     | Tekst (één regel)   | Ja        |
| `TypeKlacht`        | Keuze               | Ja        |
| `Omschrijving`      | Meerdere regels     | Ja        |
| `Bedrag`            | Getal               | Ja        |
| `Status`            | Keuze               | Ja        |
| `Melder`            | Tekst (één regel)   | Ja        |
| `MelderNaam`        | Tekst (één regel)   | Nee       |
| `BeoordeeldDoor`    | Tekst (één regel)   | Nee       |
| `DatumGoedkeuring`  | Datum en tijd       | Nee       |
| `WeigeringReden`    | Meerdere regels     | Nee       |
| `IsHistorisch`      | Ja/Nee              | Nee       |

**Keuze-opties voor `Status`:**
- Wachtend op goedkeuring
- Goedgekeurd
- Geweigerd
- Geklasseerd

**Keuze-opties voor `TypeKlacht`:**
- Logistieke fout
- Beschadiging
- Prijsverschil
- Administratief
- Kwaliteit

---

## 3. Configuratie in app.js aanpassen

```javascript
const CONFIG = {
  clientId:    'JOUW_CLIENT_ID',           // stap 1
  tenantId:    'e65dbe4b-d1e2-4283-b0f5-aa7717e81077',
  spSiteUrl:   'https://verpabenelux.sharepoint.com/sites/Intranet', // aanpassen
  spListName:  'KlachtenMeldingen',
  adminUsers:  ['ils@verpabenelux.be'],    // UPN(s) van beheerder(s)
};
```

---

## 4. Cloudflare Pages deployment

```bash
# Installeer Wrangler
npm install -g wrangler

# Publiceer alle bestanden (index.html, app.js, styles.css)
npx wrangler pages deploy . --project-name verpa-klachten
```

Of via Cloudflare Dashboard → Pages → Direct Upload → 3 bestanden uploaden.

**Custom domain (optioneel):** verpa-klachten.pages.dev

---

## 5. Rollen

| Rol        | Wie                        | Rechten                                               |
|------------|----------------------------|-------------------------------------------------------|
| Melder     | Iedereen met M365 account  | Eigen klachten indienen en bekijken                   |
| Beheerder  | UPN in `adminUsers`        | Alle klachten zien, goedkeuren/weigeren, importeren   |

---

## 6. Bestandsstructuur

```
verpa-klachten/
├── index.html   – HTML structuur + modals
├── styles.css   – Stijlen (Verpa huisstijl)
├── app.js       – Alle logica (MSAL, SharePoint, export, import)
└── README.md    – Dit bestand
```
