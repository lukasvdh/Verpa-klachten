# Verpa Klachtenmeldingen – Setup Instructies

## 1. Azure App Registration

1. Ga naar **portal.azure.com** → Azure Active Directory → App Registrations → New Registration
2. Naam: `Verpa Klachten App`
3. Redirect URI: `https://verpa-klachten.pages.dev` (Single-page application)
4. Kopieer de **Client ID** en zet in `app.js` bij `CONFIG.clientId`
5. **API Permissions** (Delegated):
   - `User.Read`
   - `Sites.ReadWrite.All`
   - `Dynamics 365 Business Central` → `Financials.ReadWrite.All`
   - `Dynamics 365 Business Central` → `API.ReadWrite.All` *(vereist voor schrijven naar custom BC API)*
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
| `BehandelStatus`    | Keuze               | Nee       |
| `DatumAfhandeling`  | Datum en tijd       | Nee       |

**Keuze-opties voor `Status`:**
- Wachtend op goedkeuring
- Goedgekeurd
- Geweigerd
- Geklasseerd

**Keuze-opties voor `BehandelStatus`:**
- Nieuw
- In behandeling
- Afgehandeld

**Keuze-opties voor `TypeKlacht`:**
- Foute bestelling
- Beschadiging
- Prijsverschil
- Administratief
- Kwaliteit

---

## 3. Configuratie in app.js aanpassen

```javascript
const CONFIG = {
  clientId:     'e1f6ac61-a64c-4a4b-a2f5-f061c989983f', // Verpa Klachten App
  tenantId:     'e65dbe4b-d1e2-4283-b0f5-aa7717e81077',
  spSiteUrl:    'https://verpabenelux.sharepoint.com/sites/OfficeData',
  spListName:   'KlachtenMeldingen',
  adminGroupId: 'JOUW_ENTRA_GROEP_OBJECT_ID', // Object-ID van de Entra-groep
};

// Business Central
const BC_ENV = 'Verpa_Accept_09022026'; // naam van de BC-omgeving
```

---

## 4. Business Central Extensie

De app maakt gebruik van een custom AL-extensie (**Verpa Klachten API v1.2.0.0**) die apart gepubliceerd moet worden in BC.

### Wat de extensie toevoegt
| Object | ID | Omschrijving |
|---|---|---|
| Table | 51000 | Verpa Klacht – opslag van alle klachtgegevens |
| Page (List) | 50501 | Verpa Klachten – overzicht in BC UI |
| Page (Card) | 50502 | Verpa Klacht Card – detailpagina in BC UI |
| Page (API) | 50503 | Verpa Klacht API – REST endpoint voor de app |
| Page (API) | 50500 | Verpa ShipTo API – leveradressen per klant |
| Enum | 50500 | Verpa Klacht Status |
| Enum | 50501 | Verpa Klacht Type |

### API Endpoints na publicatie
```
# Klanten zoeken
GET /api/v2.0/companies({id})/customers?$filter=startswith(number,'K001')

# Artikelen zoeken
GET /api/v2.0/companies({id})/items?$filter=contains(displayName,'zeep')

# Leveradressen per klant
GET /api/verpa/klachten/v1.0/companies({id})/shipToAddresses?$filter=customerNo eq 'K16542'

# Klachten opvragen
GET /api/verpa/klachten/v1.0/companies({id})/klachten

# Klacht aanmaken (automatisch bij indienen in app)
POST /api/verpa/klachten/v1.0/companies({id})/klachten
```

### BC omgeving URL
```
https://businesscentral.dynamics.com/Verpa_Accept_09022026?company=Verpa&page=50501
```

---

## 5. Cloudflare Pages deployment

```bash
# Installeer Wrangler
npm install -g wrangler

# Publiceer alle bestanden (index.html, app.js, styles.css)
npx wrangler pages deploy . --project-name verpa-klachten
```

Of via Cloudflare Dashboard → Pages → Direct Upload → bestanden uploaden.

**URL:** https://verpa-klachten.pages.dev

---

## 6. Rollen

| Rol        | Wie                             | Rechten                                                        |
|------------|---------------------------------|----------------------------------------------------------------|
| Melder     | Iedereen met M365 account       | Eigen klachten indienen en bekijken                            |
| Beheerder  | Leden van de Entra-groep        | Alle klachten zien, goedkeuren/weigeren, importeren, BC-sync  |

---

## 7. Bestandsstructuur

```
verpa-klachten/
├── index.html   – HTML structuur + modals
├── styles.css   – Stijlen (Verpa huisstijl)
├── app.js       – Alle logica (MSAL, SharePoint, BC API, export, import)
├── .gitignore   – Git uitsluitingen
└── README.md    – Dit bestand
```

---

## 8. BC Sync – bestaande klachten

Bij de eerste ingebruikname kunnen bestaande SharePoint-klachten eenmalig gesynchroniseerd worden naar BC via de knop **"Alle klachten naar BC synchroniseren"** in de Import-tab (enkel zichtbaar voor beheerders). Vereist `API.ReadWrite.All` machtiging.

Nieuwe klachten worden automatisch gesynchroniseerd naar BC bij het indienen. Als de BC-sync mislukt (bv. geen verbinding), blijft de klacht altijd bewaard in SharePoint.
