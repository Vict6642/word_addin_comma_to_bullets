# Word Add-in: Komma → Punktform (maks 3 kolonner)

## Sådan bruges

1. Kør en lokal webserver og host filerne (`taskpane.html`, `taskpane.js`, `taskpane.css`).
   - Eksempel: `npx http-server` eller `live-server` i mappen.
2. Rediger `manifest.xml` så `SourceLocation` peger på den lokale server.
3. Følg Microsofts vejledning for sideloading af Word add-ins:
   https://learn.microsoft.com/da-dk/office/dev/add-ins/testing/sideload-office-add-ins-for-testing
4. I Word:
   - Åbn add-in'et fra "Indsæt → Mine tilføjelsesprogrammer".
   - Marker tekst **inde i et tekstfelt**.
   - Klik **Konverter**.

Så opdeles teksten i punktopstilling og fordeles automatisk over op til **3 kolonner**.
