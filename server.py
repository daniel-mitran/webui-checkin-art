#!/usr/bin/env python3
"""
Check-in invitati pe tablete, sincronizat intre tablete.

- Lista de invitati e STATICA: guests.json (generat din Excel cu import_excel.py).
- Textele evenimentului (brand, titlu, data) vin din event.json; culorile din static/theme.css.
- Bifarile:
    * pe Render: in Postgres (variabila DATABASE_URL) -> deploy fara downtime,
      mai multe instante pot rula simultan pe aceeasi baza de date;
    * local, fara DATABASE_URL: in checkins.json din $DATA_DIR (rezerva pe laptop).
- /api/export.csv descarca situatia (cine a venit, la ce ora, cate persoane).
- Optional: la fiecare 2 minute situatia e copiata intr-un Google Sheet, printr-un
  script Apps Script (google-apps-script/sync.gs). Variabile: GSHEET_WEBAPP_URL, GSHEET_TOKEN.

Productie: Render.com (vezi render.yaml).   Local: python3 server.py
"""

import argparse
import csv
import io
import json
import os
import re
import signal
import socket
import threading
import time
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
GUESTS_FILE = os.path.join(HERE, "guests.json")
EVENT_FILE = os.path.join(HERE, "event.json")
INDEX_FILE = os.path.join(HERE, "index.html")
STATIC_DIR = os.path.join(HERE, "static")
STATIC_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".svg": "image/svg+xml",
                ".ico": "image/x-icon", ".webp": "image/webp", ".css": "text/css; charset=utf-8"}


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def now_ms():
    return int(time.time() * 1000)


def clamp_n(n):
    return max(1, min(int(n), 99))


# --------------------------------------------------------------------------- stocare bifari

class JsonStore:
    """Rezerva locala: un singur proces, bifarile in checkins.json."""
    kind = "fisier local"

    def __init__(self, data_dir):
        os.makedirs(data_dir, exist_ok=True)
        self.path = os.path.join(data_dir, "checkins.json")
        self.lock = threading.Lock()
        self.checkins = load_json(self.path, {})
        self.version = 1

    def _save(self):
        self.version += 1
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.checkins, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.path)

    def healthy(self):
        return True

    def snapshot(self):
        with self.lock:
            return self.version, dict(self.checkins)

    def checkin(self, gid, n):
        with self.lock:
            if gid in self.checkins:
                return False
            self.checkins[gid] = {"ts": now_ms(), "n": clamp_n(n)}
            self._save()
            return True

    def set_count(self, gid, n):
        with self.lock:
            if gid not in self.checkins:
                return False
            self.checkins[gid]["n"] = clamp_n(n)
            self._save()
            return True

    def undo(self, gid):
        with self.lock:
            if gid not in self.checkins:
                return False
            del self.checkins[gid]
            self._save()
            return True

    def reset(self):
        with self.lock:
            self.checkins = {}
            self._save()


class PgStore:
    """Postgres: sigur cu mai multe instante simultan (deploy fara downtime).
    Fiecare modificare creste un contor de versiune in aceeasi tranzactie, iar
    tabletele recitesc bifarile doar cand versiunea s-a schimbat."""
    kind = "Postgres"

    SCHEMA = """
    CREATE TABLE IF NOT EXISTS checkins (
        event    text   NOT NULL,
        guest_id text   NOT NULL,
        ts       bigint NOT NULL,
        n        int    NOT NULL,
        PRIMARY KEY (event, guest_id)
    );
    CREATE TABLE IF NOT EXISTS checkin_version (
        event text PRIMARY KEY,
        v     bigint NOT NULL
    );
    """

    def __init__(self, url, event):
        from psycopg_pool import ConnectionPool
        self.event = event
        self.pool = ConnectionPool(url, min_size=1, max_size=8, open=True, timeout=4,
                                   check=ConnectionPool.check_connection)
        with self.pool.connection() as c:
            c.execute(self.SCHEMA)
            c.execute("INSERT INTO checkin_version(event, v) VALUES (%s, 1) ON CONFLICT DO NOTHING", (event,))
        self._cache = (None, {})
        self._cache_lock = threading.Lock()

    def healthy(self):
        try:
            with self.pool.connection() as c:
                c.execute("SELECT 1")
            return True
        except Exception:
            return False

    def snapshot(self):
        with self.pool.connection() as c:
            v = c.execute("SELECT v FROM checkin_version WHERE event=%s", (self.event,)).fetchone()[0]
            with self._cache_lock:
                if self._cache[0] == v:
                    return v, self._cache[1]
            rows = c.execute("SELECT guest_id, ts, n FROM checkins WHERE event=%s", (self.event,)).fetchall()
        data = {g: {"ts": ts, "n": n} for g, ts, n in rows}
        with self._cache_lock:
            self._cache = (v, data)
        return v, data

    def _bump(self, c):
        c.execute("UPDATE checkin_version SET v = v + 1 WHERE event=%s", (self.event,))

    def checkin(self, gid, n):
        with self.pool.connection() as c:
            cur = c.execute("INSERT INTO checkins(event, guest_id, ts, n) VALUES (%s,%s,%s,%s) "
                            "ON CONFLICT DO NOTHING", (self.event, gid, now_ms(), clamp_n(n)))
            if cur.rowcount != 1:
                return False
            self._bump(c)
            return True

    def set_count(self, gid, n):
        with self.pool.connection() as c:
            cur = c.execute("UPDATE checkins SET n=%s WHERE event=%s AND guest_id=%s",
                            (clamp_n(n), self.event, gid))
            if cur.rowcount != 1:
                return False
            self._bump(c)
            return True

    def undo(self, gid):
        with self.pool.connection() as c:
            cur = c.execute("DELETE FROM checkins WHERE event=%s AND guest_id=%s", (self.event, gid))
            if cur.rowcount != 1:
                return False
            self._bump(c)
            return True

    def reset(self):
        with self.pool.connection() as c:
            c.execute("DELETE FROM checkins WHERE event=%s", (self.event,))
            self._bump(c)


# --------------------------------------------------------------------------- Google Sheet

class SheetSync(threading.Thread):
    """La fiecare GSHEET_INTERVAL secunde (implicit 120) trimite situatia catre scriptul
    Apps Script din google-apps-script/sync.gs, care o scrie intr-o foaie cu numele
    evenimentului (o creeaza daca lipseste). Ruleaza si imediat dupa pornire. O eroare
    (Google indisponibil etc.) nu afecteaza check-in-ul: se reincearca la pasul urmator."""

    def __init__(self, app, url, token, sheet_id, sheet_name, interval):
        super().__init__(daemon=True, name="sheet-sync")
        self.app, self.url, self.token = app, url, token
        self.sheet_id, self.sheet_name, self.interval = sheet_id, sheet_name, interval
        self.stop_event = threading.Event()
        self.status = {"enabled": True, "sheet": sheet_name, "ok": None, "at": None, "error": None}

    @staticmethod
    def from_env(app):
        url = (os.environ.get("GSHEET_WEBAPP_URL") or "").strip()
        if not url:
            return None
        ev = app.event
        raw_id = (os.environ.get("GSHEET_ID") or ev.get("googleSheet") or "").strip()
        m = re.search(r"/d/([A-Za-z0-9_-]{20,})", raw_id)
        sheet_id = m.group(1) if m else raw_id
        name = ev.get("sheetName") or " · ".join(x for x in (ev.get("brand"), ev.get("title"), ev.get("subtitle")) if x)
        name = re.sub(r"[\[\]*?/\\:]", "-", name or "Check-in").strip()[:99]   # caractere interzise in Google Sheets
        try:
            interval = max(30, int(os.environ.get("GSHEET_INTERVAL") or 120))
        except ValueError:
            interval = 120
        return SheetSync(app, url, os.environ.get("GSHEET_TOKEN", ""), sheet_id, name, interval)

    def payload(self):
        rows, totals, tot = self.app.situation()
        now = datetime.now(self.app.tz).strftime("%d.%m.%Y %H:%M:%S")
        ev = self.app.event
        return {
            "token": self.token, "spreadsheetId": self.sheet_id, "sheetName": self.sheet_name,
            "title": " · ".join(x for x in (ev.get("brand"), ev.get("title"), ev.get("subtitle")) if x) or "Check-in",
            "summary": "Ultima sincronizare: %s · Au venit %d din %d persoane (%d din %d invitații)" % (
                now, tot["persoane_sosite"], tot["persoane"], tot["invitatii_sosite"], tot["invitatii"]),
            "header": App.HEADER, "rows": rows, "totals": totals,
        }

    def sync_once(self):
        try:
            body = json.dumps(self.payload(), ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(self.url, data=body, method="POST",
                                         headers={"Content-Type": "application/json"})
            # Apps Script raspunde cu 302 catre rezultat; urllib urmeaza redirectarea cu GET
            with urllib.request.urlopen(req, timeout=45) as r:
                text = r.read().decode("utf-8", "replace")
            try:
                res = json.loads(text)
            except ValueError:
                raise RuntimeError("raspuns neasteptat de la Google (verifica in Apps Script: "
                                   "Deploy -> Web app, Who has access: Anyone)")
            if not res.get("ok"):
                raise RuntimeError(res.get("error") or "eroare necunoscuta in Apps Script")
            self.status.update(ok=True, at=now_ms(), error=None)
        except Exception as e:
            msg = "HTTP %d" % e.code if isinstance(e, urllib.error.HTTPError) else str(e)
            self.status.update(ok=False, at=now_ms(), error=msg[:200])
            print("Google Sheet: sincronizare esuata: %s" % msg, flush=True)
        return self.status["ok"]

    def run(self):
        if self.stop_event.wait(5):          # prima sincronizare la ~5 s dupa pornire
            return
        while True:
            self.sync_once()
            if self.stop_event.wait(self.interval):
                return


# --------------------------------------------------------------------------- aplicatia

class App:
    def __init__(self):
        db = load_json(GUESTS_FILE, None)
        if not db or not db.get("guests"):
            raise SystemExit("guests.json lipseste sau e gol - ruleaza import_excel.py")
        self.guests = db["guests"]
        self.by_id = {g["id"]: g for g in self.guests}
        self.event = load_json(EVENT_FILE, {})
        event_id = os.environ.get("EVENT_ID") or self.event.get("id") or "eveniment"
        url = os.environ.get("DATABASE_URL")
        self.store = PgStore(url, event_id) if url else JsonStore(os.environ.get("DATA_DIR") or HERE)
        try:
            from zoneinfo import ZoneInfo
            self.tz = ZoneInfo(os.environ.get("EVENT_TZ") or self.event.get("timezone") or "Europe/Bucharest")
        except Exception:
            self.tz = timezone.utc
        self.sheet_sync = SheetSync.from_env(self)

    def state(self, full=False):
        v, checkins = self.store.snapshot()
        # bifarile pentru id-uri care nu exista in lista curenta sunt ignorate
        out = {"version": v, "checkins": {k: c for k, c in checkins.items() if k in self.by_id},
               "serverTime": now_ms()}
        if self.sheet_sync:
            out["sheetSync"] = self.sheet_sync.status
        if full:
            out["event"] = self.event
            out["guests"] = self.guests
        return out

    HEADER = ["Nr.", "Status", "Ora sosirii", "Nume", "Detaliu", "Persoane invitate", "Persoane sosite"]

    def situation(self):
        """Tabelul situatiei: intai cei sositi (dupa ora), apoi cei care n-au venit (alfabetic)."""
        _, checkins = self.store.snapshot()
        checkins = {k: c for k, c in checkins.items() if k in self.by_id}
        came = sorted((g for g in self.guests if g["id"] in checkins), key=lambda g: checkins[g["id"]]["ts"])
        missing = sorted((g for g in self.guests if g["id"] not in checkins), key=lambda g: g["name"].lower())
        rows = []
        for i, g in enumerate(came, 1):
            c = checkins[g["id"]]
            t = datetime.fromtimestamp(c["ts"] / 1000, self.tz).strftime("%H:%M")
            rows.append([i, "A venit", t, g["name"], g.get("detail", ""), g["count"], c["n"]])
        for i, g in enumerate(missing, len(came) + 1):
            rows.append([i, "N-a venit", "", g["name"], g.get("detail", ""), g["count"], 0])
        tot = {"invitatii": len(self.guests), "invitatii_sosite": len(came),
               "persoane": sum(g["count"] for g in self.guests),
               "persoane_sosite": sum(checkins[g["id"]]["n"] for g in came)}
        totals = ["", "Total", "", "%d din %d invitații" % (tot["invitatii_sosite"], tot["invitatii"]), "",
                  tot["persoane"], tot["persoane_sosite"]]
        return rows, totals, tot

    def export_csv(self):
        rows, totals, _ = self.situation()
        buf = io.StringIO()
        w = csv.writer(buf, delimiter=";")          # ; -> Excel in romana deschide direct pe coloane
        w.writerow(self.HEADER)
        w.writerows(rows)
        w.writerow([])
        w.writerow(totals)
        return ("\ufeff" + buf.getvalue()).encode("utf-8")   # BOM -> diacritice corecte in Excel


APP = None


class Handler(BaseHTTPRequestHandler):
    server_version = "Checkin/2.0"

    def log_message(self, fmt, *args):
        line = args[0] if args and isinstance(args[0], str) else ""
        if not any(p in line for p in ("/api/state", "/healthz", "/static/")):
            print("[%s] %s" % (self.client_address[0], fmt % args), flush=True)

    def _send(self, code, body, ctype="application/json; charset=utf-8", cache="no-store", extra=None):
        if not isinstance(body, bytes):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", cache)
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        u = urlparse(self.path)
        try:
            if u.path in ("/", "/index.html"):
                with open(INDEX_FILE, "rb") as f:
                    return self._send(200, f.read(), "text/html; charset=utf-8")
            if u.path == "/healthz":
                ok = APP.store.healthy()
                return self._send(200 if ok else 503, {"ok": ok, "guests": len(APP.guests), "storage": APP.store.kind})
            if u.path.startswith("/static/"):
                name = os.path.basename(u.path)
                path = os.path.join(STATIC_DIR, name)
                ext = os.path.splitext(name)[1].lower()
                if ext in STATIC_TYPES and os.path.isfile(path):
                    with open(path, "rb") as f:
                        return self._send(200, f.read(), STATIC_TYPES[ext], "no-cache")
                return self._send(404, {"error": "not found"})
            if u.path == "/api/state":
                return self._send(200, APP.state(full="full=1" in (u.query or "")))
            if u.path == "/api/export.csv":
                fname = "situatie-checkin-%s.csv" % datetime.now(APP.tz).strftime("%Y-%m-%d-%H%M")
                return self._send(200, APP.export_csv(), "text/csv; charset=utf-8",
                                  extra={"Content-Disposition": 'attachment; filename="%s"' % fname})
        except Exception as e:
            print("Eroare GET %s: %r" % (u.path, e), flush=True)
            return self._send(503, {"error": "server indisponibil"})
        self._send(404, {"error": "not found"})

    def do_HEAD(self):
        self.send_response(200 if urlparse(self.path).path in ("/", "/healthz") else 404)
        self.end_headers()

    def do_POST(self):
        u = urlparse(self.path)
        b = self._body()
        gid = b.get("id")
        st = APP.store
        try:
            if u.path in ("/api/checkin", "/api/count", "/api/undo") and gid not in APP.by_id:
                return self._send(400, {"error": "invitat necunoscut"})
            if u.path == "/api/checkin":
                n = b.get("n") or APP.by_id[gid]["count"]
                ok = st.checkin(gid, n)
            elif u.path == "/api/count":
                ok = st.set_count(gid, b.get("n"))
            elif u.path == "/api/undo":
                ok = st.undo(gid)
            elif u.path == "/api/reset":
                if b.get("confirm") != "RESET":
                    return self._send(400, {"error": "confirmare lipsa"})
                st.reset()
                ok = True
            else:
                return self._send(404, {"error": "not found"})
        except (TypeError, ValueError):
            return self._send(400, {"error": "date invalide"})
        except Exception as e:
            print("Eroare POST %s: %r" % (u.path, e), flush=True)
            return self._send(503, {"error": "server indisponibil"})
        self._send(200, {"ok": ok, **APP.state()})


def local_ips():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return [ip]
    except Exception:
        return []


def main():
    global APP
    ap = argparse.ArgumentParser(description="Check-in invitati")
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8080")))
    a = ap.parse_args()
    APP = App()
    v, checkins = APP.store.snapshot()
    print("  %s: %d invitatii, %d persoane; stocare: %s, %d deja bifate" % (
        APP.event.get("title", "eveniment"), len(APP.guests), sum(g["count"] for g in APP.guests),
        APP.store.kind, len(checkins)), flush=True)
    srv = ThreadingHTTPServer(("0.0.0.0", a.port), Handler)
    srv.daemon_threads = True

    # Render trimite SIGTERM la oprirea instantei vechi: terminam cererile in curs si iesim curat
    def stop(*_):
        threading.Thread(target=srv.shutdown, daemon=True).start()
    signal.signal(signal.SIGTERM, stop)

    if APP.sheet_sync:
        APP.sheet_sync.start()
        print("  Google Sheet: sincronizare la %d s in foaia %r" % (APP.sheet_sync.interval, APP.sheet_sync.sheet_name), flush=True)

    print("  Pornit pe portul %d. Local: %s" % (
        a.port, ", ".join("http://%s:%d" % (ip, a.port) for ip in local_ips()) or "-"), flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    if APP.sheet_sync:
        APP.sheet_sync.stop_event.set()
    print("Oprit.", flush=True)


if __name__ == "__main__":
    main()
