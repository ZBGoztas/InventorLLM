"""
inventor_api.db üretici — Autodesk Inventor kurulumundan API dokümantasyon veritabanı oluşturur.

Autodesk'in dokümantasyonu bu repoda dağıtılmaz; veritabanı her kullanıcının kendi kurulumundan üretilir.
Şema server.py'nin beklediğiyle aynıdır (pages + pages_fts, members, enums, objects).

Kaynaklar:
  1. Tip kütüphanesi (RxInventor.tlb)        -> kesin imzalar, varsayılan değerler, enum sabitleri  (zorunlu)
  2. API yardım dosyası (admapi_*.chm)        -> nesne/üye sayfaları, örnek kodlar, kılavuz sayfaları (bulunursa)

Kullanım (Windows, Inventor kurulu makinede; Inventor'ın açık olması gerekmez):
    pip install pywin32 markdownify
    python build_api_db.py
    python build_api_db.py --tlb "C:\\Program Files\\Autodesk\\Inventor 2027\\Bin\\RxInventor.tlb"
    python build_api_db.py --chm "C:\\Users\\Public\\Documents\\Autodesk\\Inventor 2027\\Local Help\\admapi_31_0.chm"
"""

import argparse
import glob
import html
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
from html.parser import HTMLParser
from pathlib import Path

import pythoncom

try:
    from markdownify import markdownify as _md
except ImportError:  # markdownify yoksa düz metne düşer
    _md = None


# ---------------------------------------------------------------------------
# Kurulum bulma
# ---------------------------------------------------------------------------
def _program_files():
    seen = []
    for pf in (os.environ.get("ProgramFiles"), r"C:\Program Files", r"D:\Program Files"):
        if pf and pf not in seen:
            seen.append(pf)
    return seen


def find_tlb():
    hits = []
    for pf in _program_files():
        hits += glob.glob(os.path.join(pf, "Autodesk", "Inventor *", "Bin", "RxInventor.tlb"))
    return sorted(hits, reverse=True)[0] if hits else None  # en yeni sürüm önce


def find_chm(tlb):
    roots = [os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), "Documents", "Autodesk")]
    if tlb:
        roots.append(str(Path(tlb).parents[1]))  # ...\Inventor 20xx
    hits = []
    for r in roots:
        hits += glob.glob(os.path.join(r, "**", "admapi*.chm"), recursive=True)
    if tlb:  # tlb ile aynı sürümü tercih et
        ver = re.search(r"Inventor (\d{4})", tlb)
        if ver:
            same = [h for h in hits if ver.group(1) in h]
            hits = same or hits
    return sorted(hits, reverse=True)[0] if hits else None


# ---------------------------------------------------------------------------
# Tip kütüphanesi
# ---------------------------------------------------------------------------
VT = {
    0: "Empty", 1: "Null", 2: "Integer", 3: "Long", 4: "Single", 5: "Double", 6: "Currency", 7: "Date",
    8: "String", 9: "Object", 10: "Error", 11: "Boolean", 12: "Variant", 13: "IUnknown", 14: "Decimal",
    16: "Byte", 17: "Byte", 18: "Integer", 19: "Long", 20: "LongLong", 21: "LongLong", 22: "Long",
    23: "Long", 24: "void", 25: "HRESULT", 30: "String", 31: "String",
}
VT_PTR, VT_SAFEARRAY, VT_CARRAY, VT_USERDEFINED = 26, 27, 28, 29
PARAMFLAG_FOUT, PARAMFLAG_FRETVAL, PARAMFLAG_FOPT, PARAMFLAG_FHASDEFAULT = 0x2, 0x8, 0x10, 0x20
IUNKNOWN_IDISPATCH = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo",
                      "GetIDsOfNames", "Invoke"}
INV = {1: "method", 2: "get", 4: "put", 8: "putref"}


def type_str(ti, td):
    try:
        if isinstance(td, int):
            return VT.get(td & 0xFFF, f"VT{td}") + ("()" if td & 0x2000 else "")
        vt, inner = td[0], td[1]
        if vt == VT_PTR:
            return type_str(ti, inner)
        if vt == VT_SAFEARRAY:
            return type_str(ti, inner) + "()"
        if vt == VT_CARRAY:
            return type_str(ti, inner[0] if isinstance(inner, tuple) else inner) + "()"
        if vt == VT_USERDEFINED:
            return ti.GetRefTypeInfo(inner).GetDocumentation(-1)[0]
        return VT.get(vt, f"VT{vt}")
    except Exception:
        return "?"


def _default_str(v):
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, str):
        return f'"{v}"'
    return str(v)


def read_typelib(path):
    objects, members, enums = {}, {}, []
    tl = pythoncom.LoadTypeLib(path)
    lib = tl.GetDocumentation(-1)[0]
    for i in range(tl.GetTypeInfoCount()):
        ti = tl.GetTypeInfo(i)
        attr = ti.GetTypeAttr()
        name, doc = ti.GetDocumentation(-1)[:2]
        kind = attr.typekind
        if kind == pythoncom.TKIND_ENUM:
            for v in range(attr.cVars):
                vd = ti.GetVarDesc(v)
                enums.append((name, ti.GetNames(vd.memid)[0], int(vd.value)))
            objects.setdefault(name, ("enum", doc or ""))
            continue
        if kind == pythoncom.TKIND_COCLASS:
            objects.setdefault(name, ("coclass", doc or ""))
            continue
        if kind not in (pythoncom.TKIND_DISPATCH, pythoncom.TKIND_INTERFACE):
            continue
        if kind == pythoncom.TKIND_DISPATCH and attr.wTypeFlags & 0x40:  # TYPEFLAG_FDUAL
            try:  # dual: arayüz tarafına geç -> kendi üyeleri + gerçek üst sınıf
                ti = ti.GetRefTypeInfo(ti.GetRefTypeOfImplType(-1))
                attr = ti.GetTypeAttr()
            except Exception:
                pass
        if objects.get(name, ("",))[0] == "object":
            continue
        objects[name] = ("object", doc or "")
        for f in range(attr.cFuncs):
            fd = ti.GetFuncDesc(f)
            names = ti.GetNames(fd.memid)
            mname = names[0]
            if mname in IUNKNOWN_IDISPATCH:
                continue
            mdoc = ti.GetDocumentation(fd.memid)[1] or ""
            ret = type_str(ti, fd.rettype[0])
            params = []
            for k, a in enumerate(fd.args):
                td, flags = a[0], a[1]
                pt = type_str(ti, td)
                if flags & PARAMFLAG_FRETVAL:
                    ret = pt
                    continue
                p = {"name": names[k + 1] if k + 1 < len(names) else f"arg{k}", "type": pt,
                     "optional": bool(flags & (PARAMFLAG_FOPT | PARAMFLAG_FHASDEFAULT))}
                if flags & PARAMFLAG_FOUT:
                    p["out"] = True
                if flags & PARAMFLAG_FHASDEFAULT and len(a) > 2 and a[2] is not None:
                    p["default"] = _default_str(a[2])
                params.append(p)
            key = (name, mname)
            inv = INV.get(fd.invkind, str(fd.invkind))
            hidden = bool(fd.wFuncFlags & 0x40)  # FUNCFLAG_FHIDDEN
            if key in members:
                m = members[key]
                m["access"].add(inv)
                if inv == "get" or (inv == "method" and not m["params"]):
                    m["returns"], m["params"] = ret, params
                m["doc"] = m["doc"] or mdoc
            else:
                members[key] = {"access": {inv},
                                "returns": ret if inv != "put" else (params[-1]["type"] if params else ret),
                                "params": params if inv not in ("put", "putref") else params[:-1],
                                "doc": mdoc, "hidden": hidden}
    return lib, objects, members, enums


def access_str(acc):
    if "method" in acc:
        return "method"
    if "get" in acc and ("put" in acc or "putref" in acc):
        return "get/set"
    if "get" in acc:
        return "get (read-only)"
    return "set"


# ---------------------------------------------------------------------------
# CHM -> markdown sayfaları
# ---------------------------------------------------------------------------
class _Text(HTMLParser):
    BLOCK = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "pre", "table", "dt", "dd"}

    def __init__(self):
        super().__init__()
        self.title, self._in_title, self.parts, self._skip = "", False, [], 0

    def handle_starttag(self, tag, attrs):
        if tag == "title":
            self._in_title = True
        elif tag in ("script", "style"):
            self._skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag == "title":
            self._in_title = False
        elif tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if self._in_title:
            self.title += data
        elif not self._skip:
            self.parts.append(data)

    def text(self):
        t = html.unescape("".join(self.parts))
        t = re.sub(r"[ \t\xa0]+", " ", t)
        return re.sub(r"\n\s*\n+", "\n\n", t).strip()


def html_to_page(raw):
    """(title, markdown body) — markdownify varsa tablolar/başlıklar/kod blokları korunur."""
    p = _Text()
    p.feed(raw)
    title = re.sub(r"\s+", " ", p.title).strip()
    if _md is None:
        return title, p.text()
    body_html = re.sub(r"(?is)<(script|style|head)\b.*?</\1>", "", raw)
    md = _md(body_html, heading_style="ATX", code_language="", strip=["a", "img"])
    md = md.replace("\xa0", " ")
    md = re.sub(r"[ \t]+\n", "\n", md)
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    return title, md


def decompile_chm(chm):
    out = tempfile.mkdtemp(prefix="inventor_chm_")
    hh = shutil.which("hh.exe") or r"C:\Windows\hh.exe"
    subprocess.run([hh, "-decompile", out, chm], check=False)
    last, stable = -1, 0  # hh.exe işi arka planda bitirebilir: dosya sayısı sabitlenene kadar bekle
    for _ in range(300):
        n = sum(len(f) for _, _, f in os.walk(out))
        stable = stable + 1 if n == last and n > 0 else 0
        if stable >= 3:
            break
        last = n
        time.sleep(1)
    return out


def read_chm_pages(chm_dir):
    pages = []
    for root, _, files in os.walk(chm_dir):
        for f in files:
            if not f.lower().endswith((".htm", ".html")):
                continue
            path = os.path.join(root, f)
            raw = Path(path).read_bytes()
            for enc in ("utf-8", "cp1252"):
                try:
                    s = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            try:
                title, body = html_to_page(s)
            except Exception:
                continue
            if len(body) < 40:
                continue
            pages.append({"stem": Path(f).stem, "title": title or Path(f).stem, "body": body})
    return pages


TITLE_RE = re.compile(r"^(?P<obj>[A-Za-z_]\w*)(\.(?P<mem>[A-Za-z_]\w*))?\s+"
                      r"(?P<kind>Object|Method|Property|Event|Enumerator|Enumeration)$")
KIND = {"Object": "object", "Method": "method", "Property": "property", "Event": "event",
        "Enumerator": "enum", "Enumeration": "enum"}


def classify(pg):
    """Inventor yardım başlıkları: 'Rows Object', 'Rows.Add Method', 'XxxEnum Enumerator' ..."""
    m = TITLE_RE.match(pg["title"])
    if m:
        kind = KIND[m.group("kind")]
        obj, mem = m.group("obj"), m.group("mem")
        if mem and kind in ("method", "property", "event"):
            return f"{obj}.{mem}", kind, obj, mem
        if not mem and kind in ("object", "enum"):
            return obj, kind, obj, None
    is_sample = bool(re.search(r"sample", pg["stem"], re.I) or re.search(r"^## Code Samples", pg["body"], re.M))
    return pg["stem"], ("sample" if is_sample else "manual"), None, None


# ---------------------------------------------------------------------------
# DB
# ---------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE pages(id INTEGER PRIMARY KEY, name TEXT UNIQUE, title TEXT, kind TEXT,
                   object TEXT, member TEXT, summary TEXT, body TEXT);
CREATE TABLE members(object TEXT, member TEXT, access TEXT, returns TEXT, params TEXT, doc TEXT, hidden INTEGER,
                     PRIMARY KEY(object, member));
CREATE TABLE enums(enum TEXT, name TEXT, value INTEGER, PRIMARY KEY(enum, name));
CREATE INDEX enums_name ON enums(name);
CREATE TABLE objects(name TEXT PRIMARY KEY, kind TEXT, doc TEXT);
CREATE VIRTUAL TABLE pages_fts USING fts5(name, title, summary, body, content='pages', content_rowid='id',
                                          tokenize='unicode61 tokenchars ''_''');
"""


def first_sentence(text):
    text = re.sub(r"[#*|`>-]+", " ", text or "")
    text = re.sub(r"\s+", " ", text).strip()
    m = re.match(r"(.{20,240}?[.!?])(\s|$)", text)
    return m.group(1) if m else text[:240]


def description(body):
    """'## Description' bölümünün ilk paragrafı (yoksa gövdenin başı)."""
    m = re.search(r"^## Description\s*\n+(.+?)(\n\n|\Z)", body, re.M | re.S)
    return first_sentence(m.group(1) if m else body)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tlb", help="RxInventor.tlb yolu (otomatik bulunur)")
    ap.add_argument("--chm", help="admapi_*.chm yolu (otomatik aranır)")
    ap.add_argument("--no-chm", action="store_true", help="Yardım dosyasını atla, sadece tip kütüphanesi")
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent / "inventor_api.db"))
    a = ap.parse_args()

    tlb = a.tlb or find_tlb()
    if not tlb or not os.path.exists(tlb):
        sys.exit("RxInventor.tlb bulunamadı. --tlb ile ver (Inventor kurulumunun Bin klasöründe).")
    print(f"Tip kütüphanesi: {tlb}")
    lib, objects, members, enums = read_typelib(tlb)
    print(f"  {lib}: {sum(1 for v in objects.values() if v[0] == 'object')} nesne, "
          f"{len(members)} üye, {len(enums)} enum sabiti")

    chm_pages = []
    if not a.no_chm:
        chm = a.chm or find_chm(tlb)
        if chm and os.path.exists(chm):
            print(f"Yardım: {chm} (açılıyor...)")
            chm_pages = read_chm_pages(decompile_chm(chm))
            print(f"  {len(chm_pages)} sayfa" + ("" if _md else "  (markdownify yok: düz metin)"))
        else:
            print("admapi*.chm bulunamadı; sadece tip kütüphanesi kullanılacak (--chm ile verebilirsin).")

    out = Path(a.out)
    if out.exists():
        out.unlink()
    db = sqlite3.connect(str(out))
    db.executescript(SCHEMA)
    db.executemany("INSERT OR IGNORE INTO enums VALUES (?,?,?)", enums)
    db.executemany("INSERT OR REPLACE INTO objects VALUES (?,?,?)", [(n, k, d) for n, (k, d) in objects.items()])
    by_obj = {}
    for (o, m), v in members.items():
        by_obj.setdefault(o, []).append(m)
        db.execute("INSERT OR REPLACE INTO members VALUES (?,?,?,?,?,?,?)",
                   (o, m, access_str(v["access"]), v["returns"], json.dumps(v["params"]), v["doc"],
                    int(v["hidden"])))

    # 1) Yardım sayfaları
    pages, used = [], set()
    for pg in chm_pages:
        name, kind, obj, mem = classify(pg)
        if name in used:
            continue
        used.add(name)
        pages.append((name, pg["title"], kind, obj, mem, description(pg["body"]), pg["body"]))

    # 2) Yardımda sayfası olmayan nesne / üye / enum'lar için tip kütüphanesinden üret
    for name, (kind, odoc) in objects.items():
        if kind == "enum" and name not in used:
            vals = [(n, v) for e, n, v in enums if e == name]
            body = f"# {name} Enumerator\n\n{odoc}\n\n" + "\n".join(f"{n} = {v}" for n, v in vals)
            pages.append((name, f"{name} Enumerator", "enum", name, None,
                          first_sentence(odoc) or f"{len(vals)} sabit", body))
            used.add(name)
        if kind != "object":
            continue
        ms = sorted(by_obj.get(name, []))
        if name not in used:
            lines = [f"# {name} Object", "", odoc, "", "## Members (type library)", ""]
            for m in ms:
                v = members[(name, m)]
                sig = ", ".join(p["name"] + ": " + p["type"] for p in v["params"])
                lines.append(f"- {m}({sig}) -> {v['returns']}   [{access_str(v['access'])}]")
            pages.append((name, f"{name} Object", "object", name, None,
                          first_sentence(odoc) or f"{len(ms)} üye", "\n".join(lines)))
            used.add(name)
        for m in ms:
            full = f"{name}.{m}"
            if full in used:
                continue
            v = members[(name, m)]
            sig = ", ".join(p["name"] + ": " + p["type"] + (" [out]" if p.get("out") else "") for p in v["params"])
            kind_m = "method" if "method" in v["access"] else "property"
            body = (f"# {full}\n\nParent Object: {name}\n\n{v['doc']}\n\n"
                    f"## Syntax\n\n{full}({sig}) -> {v['returns']}   [{access_str(v['access'])}]")
            pages.append((full, full, kind_m, name, m, first_sentence(v["doc"]) or sig, body))
            used.add(full)

    db.executemany("INSERT OR IGNORE INTO pages(name,title,kind,object,member,summary,body) VALUES (?,?,?,?,?,?,?)",
                   pages)
    db.execute("INSERT INTO pages_fts(pages_fts) VALUES ('rebuild')")
    db.commit()
    counts = dict(db.execute("SELECT kind, count(*) FROM pages GROUP BY kind").fetchall())
    db.execute("VACUUM")
    db.close()
    print(f"\nYazıldı: {out}  ({out.stat().st_size / 1e6:.1f} MB)")
    print("Sayfalar:", counts)


if __name__ == "__main__":
    main()
