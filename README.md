# Check-in Art — presă și invitați (174 de persoane)

Șablon pentru aplicațiile de check-in la evenimente (Meno, Metro, Single Malt…). Scrii 3 litere din nume și bifezi **A venit?**, iar bifarea apare pe toate tabletele în ~2 secunde. Bifările stau în **Render Postgres**, deci un deploy în timpul evenimentului **nu mai produce întrerupere**.

## Ce face

- **Căutare:** după nume și detaliu, fără diacritice. Rezultatele au bifă, ora sosirii și, pentru invitațiile de mai multe persoane, corecția − / + a celor care au intrat efectiv.
- **Lista completă** (buton jos): tabel compact cu numele și numărul de persoane, filtrele Toți / Au venit / N-au venit și totaluri.
- **Statistici:** persoane și invitații sosite, **Descarcă situația** (CSV care se deschide direct în Excel, cu diacritice) și ștergerea tuturor bifărilor (confirmare cu `RESET`).
- **Fără conexiune:** dacă serverul nu răspunde (rețea, restart), bifarea rămâne pe tabletă, chiar și după reîncărcarea paginii, și se trimite automat când revine conexiunea.
- **Bifări simultane:** dacă două tablete bifează aceeași invitație, doar una reușește, iar cealaltă vede „era deja bifat(ă) de pe altă tabletă”.
- **Zoom-ul e blocat** pe tablete. Fără PIN.

## Pentru un eveniment nou

1. **Repo nou din șablon:** pe GitHub, *Use this template → Create a new repository* (ex. `webui-clientx-oras`).
2. **Lista de invitați** (rămâne fixă pe durata evenimentului):
   ```bash
   python3 import_excel.py "Lista invitati.xlsx"                  # A = nume, B = nr. persoane
   python3 import_excel.py lista.xlsx --name C --count E --detail D --sheet "Invitati"
   python3 import_excel.py lista.xlsx --count -                   # fără coloană de persoane (1 fiecare)
   ```
   Scriptul scrie `guests.json`, scoate telefoanele din nume, mută textul din paranteze în detaliu și listează numele care apar de mai multe ori, ca să le verifici cu clientul. Excel-urile nu se urcă pe GitHub (`.gitignore`).
3. **Textele:** în `event.json` completezi `brand`, `brandSub`, `title` (locația), `subtitle` (data), `id` și, opțional, `googleSheet` (vezi mai jos). `id` trebuie să fie unic pe eveniment, de exemplu `clientx-2026-12-01`, ca bifările să nu se amestece dacă refolosești baza de date.
4. **Culorile:** în `static/theme.css` schimbi variabilele din `:root` după site-ul clientului. Pentru o temă deschisă există un exemplu la finalul fișierului. Logo-ul (opțional) se pune în `static/logo.png`.
5. **`render.yaml`:** schimbi numele serviciului (devine adresa `https://<nume>.onrender.com`) și numele bazei de date, în ambele locuri.
6. **Render:** *New → Blueprint* și alegi repo-ul. Se creează baza de date Postgres și serviciul web. Orice push pe `main` face deploy automat, fără întrerupere.
7. **Înainte de eveniment:** deschizi adresa pe toate tabletele, faci o bifare de test, apoi *Statistici → Șterge toate bifările*.

## Copiere automată în Google Sheet (opțional)

La fiecare 2 minute (și imediat după pornire), aplicația trimite situația în Google Sheet-ul din `event.json` → `googleSheet`. Datele ajung într-o foaie (tab) cu numele evenimentului, de forma `brand · locație · dată`, sau `sheetName` dacă îl completezi. Dacă foaia nu există, e creată automat. La fiecare sincronizare foaia e rescrisă complet: titlu, ora ultimei sincronizări cu totalurile, apoi tabelul cu cine a venit (după oră) și cine nu (alfabetic). Nu edita manual în ea.

**Configurare, o singură dată pe sheet:**
1. În Google Sheet: *Extensions → Apps Script*. Lipești conținutul din `google-apps-script/sync.gs` și dai *Save*.
2. *Project Settings → Script properties → Add script property*: `TOKEN` = un șir lung, aleator.
3. *Deploy → New deployment → Web app*, cu *Execute as: Me* și *Who has access: Anyone*. Apoi *Deploy*, autorizezi cu contul tău Google și copiezi **Web app URL**. Deschis în browser, URL-ul trebuie să răspundă `{"ok":true,...}`.
4. `GSHEET_WEBAPP_URL` și `GSHEET_TOKEN` sunt deja completate în `render.yaml` pentru scriptul sheet-ului de check-in. Le schimbi doar dacă folosești alt sheet. Token-ul stă în cod, deci **repo-ul trebuie să rămână privat**.

Același script și același URL merg pentru toate evenimentele care scriu în acest sheet, fiecare în foaia lui. Starea sincronizării apare în **Statistici**, de exemplu „sincronizat la 19:42”. Dacă Google nu răspunde, check-in-ul merge normal și sincronizarea se reîncearcă la următorul pas.

## După eveniment

- **Exportul:** *Statistici → Descarcă situația*, sau direct `https://<serviciu>.onrender.com/api/export.csv`. Conține numerotarea, statusul, ora sosirii, numele, detaliul, persoanele invitate și cele sosite, cu totaluri.
- **Ștergerea:** după export, ștergi serviciul web și baza de date din Render ca să nu mai fie facturate.
- **Atenție la planul `free`:** baza de date **expiră după 30 de zile** și nu are backup. Fă exportul înainte de expirare. Dacă vrei datele păstrate, alege planul `0.1c-256mb` ($6/lună).

## Cum funcționează

| Fișier | Rol |
|---|---|
| `server.py` | Server Python (biblioteca standard + `psycopg`). Cu `DATABASE_URL` setat scrie în Postgres; fără el, în `checkins.json` local. |
| `index.html` | Interfața tabletelor. Nu trebuie modificată pentru un eveniment nou. |
| `static/theme.css` | Culorile și fontul evenimentului. |
| `event.json` | Brand, locație, dată, `id` eveniment, fus orar, Google Sheet. |
| `google-apps-script/sync.gs` | Scriptul care scrie situația în Google Sheet. |
| `guests.json` | Lista de invitați, generată de `import_excel.py`. |
| `render.yaml` | Postgres + serviciul web, fără disc. |

**De ce nu mai există downtime:** fără disc atașat, Render pornește versiunea nouă, așteaptă ca `/healthz` să răspundă (acesta verifică și baza de date) și abia apoi oprește versiunea veche. Ambele scriu în aceeași bază, iar fiecare modificare crește un contor de versiune în aceeași tranzacție, așa că tabletele văd mereu aceeași stare.

**Limita care rămâne:** dacă baza de date sau rețeaua cad, `/healthz` răspunde 503 și tabletele țin bifările în coadă până la revenire.

## Rulare locală (rezervă, fără internet)

```bash
pip install -r requirements.txt   # doar pentru Postgres; local nu e obligatoriu
python3 server.py                 # fără DATABASE_URL -> bifări în checkins.json
```

Tabletele din aceeași rețea intră pe `http://<IP-laptop>:8080`. Fontul se încarcă de la Google Fonts. Fără internet, pagina folosește un font de sistem.
