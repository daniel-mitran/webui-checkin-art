#!/usr/bin/env python3
"""
Genereaza guests.json (lista statica de invitati) dintr-un Excel sau CSV.

    python3 import_excel.py lista.xlsx                       # A = nume, B = nr. persoane
    python3 import_excel.py lista.xlsx --name A --count C    # alte coloane
    python3 import_excel.py lista.xlsx --name A --detail B   # fara numar de persoane (=1 fiecare)
    python3 import_excel.py lista.xlsx --sheet "Invitati"    # alta foaie decat prima
    python3 import_excel.py lista.csv                        # CSV cu ; sau ,

Optiuni:
  --name A       coloana cu numele (implicit A)
  --count B      coloana cu numarul de persoane; "-" = fara coloana, fiecare invitatie = 1 (implicit B)
  --detail X     coloana cu un detaliu afisat sub nume (firma, functie...); optional
  --sheet NUME   foaia din Excel (implicit prima)
  --out FISIER   implicit guests.json langa acest script

Curatare automata:
- randurile de antet ("Nume", "Nume si prenume"...) si randul "Total" sunt sarite;
- numerele de telefon din coloana de nume sunt eliminate (nu ajung pe tablete);
- textul din paranteze, ex. "Ion Pop (Firma X - manager)", devine detaliu sub nume;
- caracterele invizibile (copiate din WhatsApp etc.) sunt eliminate.
La final se afiseaza numele care apar de mai multe ori, de verificat cu clientul.
Doar biblioteca standard Python 3.8+.
"""
import argparse
import collections
import csv
import hashlib
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
HEADER_WORDS = {"nume", "name", "nume si prenume", "nume și prenume", "invitat", "invitati", "invitați", "total"}


def col_index(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref.upper()).group(0):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def xlsx_rows(path, sheet_name=None):
    """-> lista de (numar_rand, {index_coloana: valoare})"""
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.get("Id"): r.get("Target") for r in rels}
    sheets = list(wb.find("m:sheets", NS))
    if sheet_name:
        match = [s for s in sheets if s.get("name", "").strip().lower() == sheet_name.strip().lower()]
        if not match:
            raise SystemExit("Foaia %r nu exista. Foi: %s" % (sheet_name, ", ".join(s.get("name") for s in sheets)))
        sh = match[0]
    else:
        sh = sheets[0]
    target = relmap[sh.get("{%s}id" % NS["r"])].lstrip("/")
    root = ET.fromstring(z.read(target if target.startswith("xl/") else "xl/" + target))
    rows = []
    for row in root.iter("{%s}row" % NS["m"]):
        vals = {}
        for c in row.findall("m:c", NS):
            t, v = c.get("t"), c.find("m:v", NS)
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
            elif v is not None:
                val = v.text
            else:
                val = None
            vals[col_index(c.get("r"))] = val
        rows.append((int(row.get("r")), vals))
    return rows


def csv_rows(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        text = f.read()
    delim = ";" if text.count(";") >= text.count(",") else ","
    return [(i, dict(enumerate(r))) for i, r in enumerate(csv.reader(text.splitlines(), delimiter=delim), 1)]


def clean(raw):
    s = re.sub(r"[​-‏⁠﻿]", "", str(raw or ""))   # caractere invizibile
    return re.sub(r"\s+", " ", s).strip()


def split_name(raw):
    s = clean(raw)
    details = [d.strip(" -") for d in re.findall(r"\(([^)]*)\)", s)]
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"\+?\d[\d .\-]{5,}\d", " ", s)        # telefoane
    s = re.sub(r"\s+", " ", s).strip(" -,")
    detail = " · ".join(re.sub(r"\s+-\s*|\s*-\s+", " - ", d).strip() for d in details if d)
    return s, detail


def to_int(v):
    try:
        return max(1, int(float(str(v).strip())))
    except Exception:
        return 1


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description="Excel/CSV cu invitati -> guests.json")
    ap.add_argument("file")
    ap.add_argument("--name", default="A")
    ap.add_argument("--count", default="B")
    ap.add_argument("--detail", default=None)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--out", default=os.path.join(here, "guests.json"))
    a = ap.parse_args()

    rows = csv_rows(a.file) if a.file.lower().endswith(".csv") else xlsx_rows(a.file, a.sheet)
    ci_name = col_index(a.name)
    ci_count = None if a.count in ("-", "", "none") else col_index(a.count)
    ci_detail = col_index(a.detail) if a.detail else None

    guests = []
    for rownum, vals in rows:
        name, detail = split_name(vals.get(ci_name))
        if not name or name.lower() in HEADER_WORDS:
            continue
        if ci_detail is not None:
            extra = clean(vals.get(ci_detail))
            detail = " · ".join(x for x in (extra, detail) if x)
        count = to_int(vals.get(ci_count)) if ci_count is not None else 1
        gid = hashlib.sha1(("%d|%s" % (rownum, name)).encode("utf-8")).hexdigest()[:10]
        guests.append({"id": gid, "row": rownum, "name": name, "detail": detail, "count": count})

    if not guests:
        raise SystemExit("Nu am gasit niciun invitat. Verifica --name / --sheet.")
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump({"guests": guests}, f, ensure_ascii=False, indent=1)
    print("%d invitatii, %d persoane -> %s" % (len(guests), sum(g["count"] for g in guests), a.out))

    key = lambda n: " ".join(sorted(n.lower().replace("-", " ").split()))
    dup = [n for n, k in collections.Counter(key(g["name"]) for g in guests).items() if k > 1]
    if dup:
        print("De verificat - nume care apar de mai multe ori:")
        for g in guests:
            if key(g["name"]) in dup:
                print("  rand %d: %s%s (%d pers.)" % (g["row"], g["name"], " · " + g["detail"] if g["detail"] else "", g["count"]))


if __name__ == "__main__":
    main()
