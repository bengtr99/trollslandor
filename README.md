# Trollsländeapp – webb (online röststyrning + lokal databas offline)

Webbläsarversion av Trollsländeappen. Fungerar som originalet – inklusive den
**online-röststyrning** som fungerade bra – men har nu också en **lokal databas**
som gör att sökningar fungerar offline.

## Starta

1. Öppna en kommandotolk i projektmappen.
2. `npm install`
3. `npm run dev`
4. Öppna **http://localhost:1420/** i **Chrome eller Edge** (inte Firefox – röst-API:t
   saknas där).

## Röststyrning (online)
- Använder webbläsarens inbyggda taligenkänning (`webkitSpeechRecognition`), samma
  motor som i den ursprungliga appen. Den ger bra svensk igenkänning **men kräver
  internet** (tolkningen sker hos Google) och fungerar bara i Chrome/Edge.
- Detta är anledningen till att appen är en webbapp och inte Electron: i Electron
  saknas Googles API-nyckel och röststyrningen får felet `network`.

## Lokal databas (offline)
- Backas av **SQLite via `sql.js` (WASM)**, lagrad i webbläsarens **IndexedDB**.
- Kryssa i **Lokal databas** så körs alla sökningar offline mot den.
- All databaslogik ligger i `src/localDb.js` och exponerar samma kommandon som
  den tidigare Electron-versionen, så `App.jsx` är i stort sett oförändrad.

### Två knappar för att fylla databasen
- **Skapa lokal databas** – full ombyggnad: rensar den befintliga databasen och
  laddar om ALLA observationer från SLU (alla arter, alla år). Visar en
  bekräftelseruta först eftersom det kan ta lång tid. Kör en gång för att lägga
  grunden (och då och då om du vill fånga raderingar).
- **Uppdatera databas** – inkrementell synk: hämtar bara observationer som är
  **nya eller ändrade** sedan förra synken. Detta bygger på API-fältet
  `modified` (senast ändrad), så även observationer som rättats i efterhand
  fångas. Snabbt – en fråga per art. Kräver att en full databas byggts en gång.

## Vad kräver internet?
- **Röststyrningen** (Google-molnet) och att **bygga/uppdatera** den lokala databasen.
- **Kartrutorna** (OpenStreetMap). Själva datasökningarna fungerar offline.

## Obs om data från tidigare Electron-version
Den lokala databasen lagras per webbläsare (IndexedDB) och delar inte data med
Electron-versionens fil. Första gången bygger du därför databasen på nytt. (Det
går att lägga till en importfunktion för att återanvända en befintlig
`.sqlite`-fil – säg till om det önskas.)
