"""
Inventor LLM — MCP sunucusu (v0.2)
Claude <-> Autodesk Inventor 2027 köprüsü. Python + COM (pywin32, dynamic dispatch).

Katmanlar:
  1. api_search / api_get / api_enum / api_guide  -> Inventor API dokümantasyonu (inventor_api.db)
  2. execute                                      -> Claude'un yazdığı Python kodunu Inventor'da çalıştırır
  3. Hazır tool'lar                               -> sık işler (parametre, feature ağacı, ekran görüntüsü...)

Kurallar:
- Tüm COM çağrıları tek bir işçi thread'inde yapılır.
- stdout MCP protokolüne ayrılmıştır: print() KULLANMA.
"""

import json
import math
import os
import queue
import re
import sqlite3
import tempfile
import threading
import time
import traceback
from pathlib import Path

import pythoncom
from win32com.client import dynamic
from mcp.server.fastmcp import FastMCP, Image

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "calls.jsonl"


def _find_db() -> Path:
    candidates = [
        os.environ.get("INVENTOR_API_DB", ""),
        BASE_DIR / "inventor_api.db",
        Path.home() / "Desktop" / "Inventor" / "inventor_api.db",
    ]
    for c in candidates:
        if c and Path(c).exists():
            return Path(c)
    raise FileNotFoundError("inventor_api.db bulunamadı: önce `python build_api_db.py` ile üret "
        "(veya dosyayı server.py yanına koy / INVENTOR_API_DB ayarla).")


DB_PATH = _find_db()
_db = sqlite3.connect(str(DB_PATH), check_same_thread=False)
_db_lock = threading.Lock()

# Tüm enum sabitleri belleğe: kPartDocumentObject -> 12290 ...
with _db_lock:
    ENUM_VALUES = {name: value for name, value in _db.execute("SELECT name, value FROM enums")}

INSTRUCTIONS = (
    "Autodesk Inventor 2027'yi Python+COM ile kontrol eden sunucu. "
    "İlk kullanımda api_guide'ı oku. Kod yazmadan önce api_search/api_get ile ilgili nesne ve "
    "metotların imzasına bak; ezbere API adı uydurma. Değişiklikleri küçük adımlarla execute ile yap, "
    "her adımdan sonra sonucu kontrol et (get_active_document, capture_view)."
)
mcp = FastMCP("inventor", instructions=INSTRUCTIONS)

# ---------------------------------------------------------------------------
# Birimler
# ---------------------------------------------------------------------------
UNIT_FACTORS = {"mm": 10.0, "cm": 1.0, "m": 0.01, "in": 1 / 2.54,
                "deg": 180.0 / math.pi, "rad": 1.0, "ul": 1.0}


def to_display(value_internal, unit):
    f = UNIT_FACTORS.get(unit)
    return None if f is None else round(value_internal * f, 6)


def to_internal(value_display, unit):
    f = UNIT_FACTORS.get(unit)
    if f is None:
        raise ValueError(f"Desteklenmeyen birim: {unit}. Desteklenen: {list(UNIT_FACTORS)}")
    return value_display / f


# ---------------------------------------------------------------------------
# COM işçi thread'i
# ---------------------------------------------------------------------------
_jobs: "queue.Queue" = queue.Queue()
_state = {"app": None}


def _connect():
    unk = pythoncom.GetActiveObject("Inventor.Application")
    return dynamic.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))


def _get_app():
    app = _state["app"]
    if app is not None:
        try:
            _ = app.Caption
            return app
        except Exception:
            _state["app"] = None
    _state["app"] = _connect()
    return _state["app"]


def _com_worker():
    pythoncom.CoInitialize()
    while True:
        fn, box, done = _jobs.get()
        try:
            box["ok"] = fn(_get_app())
        except Exception as e:  # noqa: BLE001
            box["err"] = f"{type(e).__name__}: {e}"
            box["tb"] = traceback.format_exc()
        finally:
            done.set()


threading.Thread(target=_com_worker, daemon=True, name="com-worker").start()


def run(fn, timeout: float = 120.0):
    box, done = {}, threading.Event()
    _jobs.put((fn, box, done))
    if not done.wait(timeout):
        raise TimeoutError("Inventor yanıt vermedi (açık bir diyalog olabilir).")
    if "err" in box:
        raise RuntimeError(box["err"])
    return box.get("ok")


def log_call(tool, args, result=None, error=None):
    rec = {"t": time.strftime("%Y-%m-%d %H:%M:%S"), "tool": tool, "args": args}
    if error:
        rec["error"] = error
    else:
        rec["result"] = result
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")


# ---------------------------------------------------------------------------
# Yardımcılar (işçi thread'inde)
# ---------------------------------------------------------------------------
def com_type_name(obj) -> str:
    try:
        return obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    except Exception:
        return "?"


def active_doc(app):
    doc = app.ActiveDocument
    if doc is None:
        raise RuntimeError("Inventor'da açık doküman yok.")
    return doc


def active_part(app):
    doc = active_doc(app)
    try:
        return doc, doc.ComponentDefinition
    except Exception:
        raise RuntimeError(f"Aktif doküman ({doc.DisplayName}) bir parça/montaj değil.")


def param_owner_map(cd) -> dict:
    owners = {}
    try:
        for feat in cd.Features:
            try:
                for p in feat.Parameters:
                    owners.setdefault(p.Name, f"feature:{feat.Name} ({com_type_name(feat)})")
            except Exception:
                pass
    except Exception:
        pass
    try:
        for sk in cd.Sketches:
            try:
                for dc in sk.DimensionConstraints:
                    owners.setdefault(dc.Parameter.Name, f"sketch:{sk.Name}")
            except Exception:
                pass
    except Exception:
        pass
    return owners


def summary(cd) -> dict:
    out = {}
    try:
        mp = cd.MassProperties
        out["mass_kg"] = round(mp.Mass, 4)
        out["volume_mm3"] = round(mp.Volume * 1000, 2)
    except Exception:
        pass
    try:
        rb = cd.RangeBox
        mn, mx = rb.MinPoint, rb.MaxPoint
        out["bbox_mm"] = {"x": round((mx.X - mn.X) * 10, 3),
                          "y": round((mx.Y - mn.Y) * 10, 3),
                          "z": round((mx.Z - mn.Z) * 10, 3)}
    except Exception:
        pass
    return out


def _jsonable(v, depth=0):
    """execute sonucunu JSON'a çevrilebilir hale getirir; COM nesneleri tip adıyla gösterilir."""
    if v is None or isinstance(v, (bool, int, float, str)):
        return v
    if isinstance(v, (list, tuple)):
        return [_jsonable(x, depth + 1) for x in v][:500] if depth < 4 else str(v)
    if isinstance(v, dict):
        return {str(k): _jsonable(x, depth + 1) for k, x in list(v.items())[:500]} if depth < 4 else str(v)
    if hasattr(v, "_oleobj_"):
        name = com_type_name(v)
        try:
            return f"<{name} '{v.Name}'>"
        except Exception:
            return f"<{name}>"
    return repr(v)


class _Constants:
    """Inventor enum sabitleri: C.kPartDocumentObject, C.kJoinOperation ..."""

    def __getattr__(self, name):
        try:
            return ENUM_VALUES[name]
        except KeyError:
            close = [k for k in ENUM_VALUES if name.lower()[1:] in k.lower()][:8]
            raise AttributeError(f"Bilinmeyen sabit: {name}. Yakın adlar: {close}") from None


C = _Constants()

# ---------------------------------------------------------------------------
# KATMAN 1 — DOKÜMANTASYON
# ---------------------------------------------------------------------------
API_GUIDE = r"""
# Inventor API — Python (COM, dynamic dispatch) rehberi

## execute ortamında hazır isimler
- app  : Inventor.Application
- doc  : app.ActiveDocument (yoksa None)
- cd   : doc.ComponentDefinition (parça/montajda)
- tg   : app.TransientGeometry   (CreatePoint2d, CreatePoint, CreateVector, CreateMatrix ...)
- to   : app.TransientObjects    (CreateObjectCollection, CreateNameValueMap ...)
- C    : enum sabitleri -> C.kJoinOperation, C.kPositiveExtentDirection, C.kPartDocumentObject
- mm(x), deg(x)  : mm -> cm, derece -> radyan (API'ye vermek için)
- to_mm(x), to_deg(x) : API'den gelen cm/radyan -> mm/derece
- print(...) : çıktıyı yakalar (sonuçta 'output' alanında döner)
- result = ... : bu değişkene atanan değer sonuç olarak döner
- com_type(obj) : COM nesnesinin tip adı (ör. 'ExtrudeFeature')

## Birimler
API iç birimi her zaman **cm** ve **radyan**dır (kütle kg). Sayısal değer verirken mm()/deg() kullan.
Değer yerine metin ifadesi verilebilen yerlerde ("20 mm") Türkçe yerel ayar ondalık VİRGÜL bekler;
bu yüzden sayısal (cm/rad) değer vermek daha güvenlidir.

## Python sözdizimi (VBA örneklerini çevirirken)
- `Set x = ...` -> `x = ...`;  `ThisApplication` -> `app`
- Koleksiyonlar 1 tabanlıdır: `cd.Sketches.Item(1)`; ada göre: `cd.Features.Item("Fillet1")`
- `for f in cd.Features:` ile gezilebilir.
- Opsiyonel parametreler atlanabilir; sıradaki bir opsiyoneli vermek için önceki opsiyonelleri de ver.
- Out (ByRef) parametreli metotlar Python'da tuple döndürebilir; emin değilsen sonucu print et.
- Enum'lar: `C.kXxx` (api_enum ile ara).

## Sık kalıplar
- Yeni sketch:  `sk = cd.Sketches.Add(cd.WorkPlanes.Item(3))`   (1=YZ, 2=XZ, 3=XY)
- Daire:        `sk.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(0,0), mm(10))`
- Profil:       `prof = sk.Profiles.AddForSolid()`
- Extrude:      `ed = cd.Features.ExtrudeFeatures.CreateExtrudeDefinition(prof, C.kJoinOperation)`
                `ed.SetDistanceExtent(mm(20), C.kPositiveExtentDirection)`
                `f = cd.Features.ExtrudeFeatures.Add(ed)`
- Yeni feature'larda genelde `CreateXxxDefinition` + `Add` kullanılır; `AddBy...` metotlarının çoğu
  "Hidden method" (eski) olarak işaretlidir.
- Sketch geometrisini döndürme: `Sketch.RotateSketchObjects`, taşıma: `Sketch.MoveSketchObjects`
- Feature bastırma: `feat.Suppressed = True`
- Model güncelleme: `doc.Update()`
- Kaydetme: `doc.Save()`, farklı kaydet: `doc.SaveAs(path, False)`

## Çalışma yöntemi
1. api_search ile ilgili nesne/metotları bul, api_get ile imzayı ve açıklamayı oku.
   Örnek kod gerekiyorsa api_search(..., kind="sample").
2. Önce okuyarak modeli anla (get_feature_tree, get_parameters, execute ile sorgu).
3. Değişikliği küçük adımlarla yap. execute her çağrıyı tek transaction'da çalıştırır;
   hata olursa otomatik geri alınır. Başarılı adım undo_last ile geri alınabilir.
4. Her adımdan sonra kontrol et: get_active_document (kütle/bbox), capture_view.
5. Kullanıcının orijinal dosyasını kaydetmeden önce sor; büyük değişikliklerde önce SaveAs ile kopya al.
"""


@mcp.tool()
def api_guide() -> str:
    """Inventor API'yi bu sunucuda Python ile kullanma rehberi. İlk kullanımda bir kez oku."""
    return API_GUIDE


@mcp.tool()
def api_search(query: str, kind: str = "", limit: int = 12) -> list[dict]:
    """
    Inventor API dokümantasyonunda arama (Reference Manual, User's Manual, örnekler).
    query: İngilizce anahtar kelimeler (ör. "extrude distance", "sketch circle", "export STEP").
    kind : filtre — object | method | property | event | enum | sample | manual (boş = hepsi).
    Sonuçtaki 'name' değerini api_get'e ver.
    """
    terms = re.findall(r"\w+", query)
    if not terms:
        return []
    fts = " OR ".join(f'"{t}"*' for t in terms)
    sql = ("SELECT p.name, p.kind, p.summary FROM pages_fts f JOIN pages p ON p.id = f.rowid "
           "WHERE pages_fts MATCH ? ")
    args = [fts]
    if kind:
        sql += "AND p.kind = ? "
        args.append(kind)
    sql += "ORDER BY bm25(pages_fts, 10, 5, 2, 1) LIMIT ?"
    args.append(max(limit * 4, 40))
    with _db_lock:
        rows = _db.execute(sql, args).fetchall()
    want_proxy = "proxy" in query.lower()
    out = []
    for name, k, summ in rows:
        if not want_proxy and "Proxy" in name:
            continue
        out.append({"name": name, "kind": k, "summary": (summ or "")[:200]})
        if len(out) >= limit:
            break
    return out


def _signature(obj, member):
    with _db_lock:
        r = _db.execute("SELECT access, returns, params, doc, hidden FROM members WHERE object=? AND member=?",
                        (obj, member)).fetchone()
    if not r:
        return None
    access, returns, params, doc, hidden = r
    ps = json.loads(params)
    sig = ", ".join(
        f"{p['name']}: {p['type']}" + (" = opsiyonel" if p.get("optional") else "") + (" [out]" if p.get("out") else "")
        for p in ps)
    return {"signature": f"{obj}.{member}({sig}) -> {returns}", "access": access,
            "hidden": bool(hidden), "doc": doc}


@mcp.tool()
def api_get(name: str, max_chars: int = 12000) -> dict:
    """
    Bir API sayfasının tamamı: "ExtrudeFeatures" (nesne), "ExtrudeFeatures.Add" (metot/özellik),
    "PartFeatureOperationEnum" (enum), örnek adı (ör. "ExtrudeFeature_Sample") veya kılavuz sayfası.
    Metot/özellikler için tip kütüphanesinden gelen kesin imza da eklenir.
    """
    with _db_lock:
        r = _db.execute("SELECT name, kind, body FROM pages WHERE name = ?", (name,)).fetchone()
        if not r:
            r = _db.execute("SELECT name, kind, body FROM pages WHERE lower(name) = lower(?)", (name,)).fetchone()
    out = {}
    if r:
        out = {"name": r[0], "kind": r[1], "body": r[2][:max_chars]}
        if len(r[2]) > max_chars:
            out["truncated"] = f"{len(r[2])} karakterin {max_chars} kadarı gösterildi"
        name = r[0]
    if "." in name:
        obj, member = name.split(".", 1)
        sig = _signature(obj, member)
        if sig:
            out["typelib"] = sig
    if not out:
        hits = api_search(name, limit=8)
        return {"error": f"'{name}' bulunamadı.", "suggestions": [h["name"] for h in hits]}
    return out


@mcp.tool()
def api_enum(query: str) -> list[dict]:
    """
    Enum sabiti arama. query: enum adı ("PartFeatureOperationEnum"), sabit adı/parçası ("kJoin")
    veya sayısal değer ("20481"). execute içinde sabitler C.kXxx ile kullanılır.
    """
    with _db_lock:
        if query.strip().lstrip("-").isdigit():
            rows = _db.execute("SELECT enum, name, value FROM enums WHERE value = ? LIMIT 50", (int(query),)).fetchall()
        else:
            rows = _db.execute("SELECT enum, name, value FROM enums WHERE enum = ? ORDER BY value", (query,)).fetchall()
            if not rows:
                rows = _db.execute("SELECT enum, name, value FROM enums WHERE name LIKE ? LIMIT 50",
                                   (f"%{query}%",)).fetchall()
    return [{"enum": e, "name": n, "value": v} for e, n, v in rows]


# ---------------------------------------------------------------------------
# KATMAN 2 — EXECUTE
# ---------------------------------------------------------------------------
@mcp.tool()
def execute(code: str, description: str, transaction: bool = True, timeout_s: int = 180) -> dict:
    """
    Inventor'da Python kodu çalıştırır (COM, dynamic dispatch). Hazır isimler: app, doc, cd, tg, to, C,
    mm(), deg(), to_mm(), to_deg(), com_type(), print(). Sonuç için `result = ...` ata.
    description: değişikliğin kısa açıklaması (Undo listesinde "LLM: <description>" olarak görünür).
    transaction: True ise kod tek transaction'da çalışır; hata olursa tüm değişiklikler geri alınır.
    Sadece okuma yapan kodlarda transaction=False verilebilir.
    Kod yazmadan önce api_guide / api_search / api_get kullan.
    """
    args = {"description": description, "transaction": transaction, "code": code}

    def job(app):
        out_lines = []

        def _print(*a, sep=" ", end="\n", **_):
            out_lines.append(sep.join(str(x) for x in a) + end)

        doc = app.ActiveDocument
        cd = None
        if doc is not None:
            try:
                cd = doc.ComponentDefinition
            except Exception:
                cd = None
        ns = {
            "app": app, "doc": doc, "cd": cd,
            "tg": app.TransientGeometry, "to": app.TransientObjects,
            "C": C, "math": math,
            "mm": lambda x: x / 10.0, "deg": lambda x: math.radians(x),
            "to_mm": lambda x: x * 10.0, "to_deg": lambda x: math.degrees(x),
            "com_type": com_type_name, "print": _print, "result": None,
        }
        tr = None
        if transaction and doc is not None:
            tr = app.TransactionManager.StartTransaction(doc, f"LLM: {description}"[:120])
        try:
            exec(compile(code, "<llm>", "exec"), ns)
            if tr is not None:
                tr.End()
        except Exception as e:  # noqa: BLE001
            if tr is not None:
                try:
                    tr.Abort()
                    doc.Update()
                except Exception:
                    pass
            tb = traceback.format_exc().splitlines()
            tb = "\n".join(l for l in tb if "server.py" not in l)[-3000:]
            return {"ok": False, "rolled_back": tr is not None, "error": f"{type(e).__name__}: {e}",
                    "traceback": tb, "output": "".join(out_lines)[-8000:]}
        res = {"ok": True, "output": "".join(out_lines)[-8000:], "result": _jsonable(ns.get("result"))}
        if cd is not None and transaction:
            res["after"] = summary(cd)
        return res

    try:
        res = run(job, timeout=timeout_s)
    except Exception as e:
        log_call("execute", args, error=str(e))
        raise
    log_call("execute", args, res)
    return res


# ---------------------------------------------------------------------------
# KATMAN 3 — HAZIR TOOL'LAR
# ---------------------------------------------------------------------------
@mcp.tool()
def connect() -> dict:
    """Çalışan Inventor oturumuna bağlanır; sürüm ve açık doküman sayısını döndürür."""
    def job(app):
        return {"version": app.SoftwareVersion.DisplayVersion,
                "open_documents": app.Documents.Count,
                "active": app.ActiveDocument.DisplayName if app.ActiveDocument else None}
    return run(job)


@mcp.tool()
def list_documents() -> list[dict]:
    """Inventor'da açık tüm dokümanlar (ad, yol, tip, kaydedilmemiş değişiklik)."""
    def job(app):
        return [{"name": d.DisplayName, "path": d.FullFileName, "type": com_type_name(d),
                 "dirty": bool(d.Dirty)} for d in app.Documents]
    return run(job)


@mcp.tool()
def get_active_document() -> dict:
    """Aktif doküman: ad, yol, tip, kaydedilmemiş değişiklik; parça/montajsa kütle ve sınır kutusu (mm)."""
    def job(app):
        doc = active_doc(app)
        out = {"name": doc.DisplayName, "path": doc.FullFileName, "type": com_type_name(doc),
               "dirty": bool(doc.Dirty)}
        try:
            out.update(summary(doc.ComponentDefinition))
        except Exception:
            pass
        return out
    return run(job)


@mcp.tool()
def get_parameters(name_filter: str = "") -> list[dict]:
    """
    Aktif parçanın parametreleri; değerler mm / deg / ul. 'owner' parametrenin ait olduğu
    feature veya sketch'tir. name_filter: ada göre kısmi filtre.
    """
    def job(app):
        _, cd = active_part(app)
        owners = param_owner_map(cd)
        try:
            user_names = {p.Name for p in cd.Parameters.UserParameters}
        except Exception:
            user_names = set()
        rows = []
        for p in cd.Parameters:
            if name_filter and name_filter.lower() not in p.Name.lower():
                continue
            rows.append({"name": p.Name, "value": to_display(p.Value, p.Units), "unit": p.Units,
                         "expression": p.Expression, "kind": "user" if p.Name in user_names else "model",
                         "owner": owners.get(p.Name), "comment": getattr(p, "Comment", "") or ""})
        return rows
    return run(job)


@mcp.tool()
def get_feature_tree() -> list[dict]:
    """Feature ağacı: sıra, ad, tip, bastırılmış mı."""
    def job(app):
        _, cd = active_part(app)
        out = []
        for i, f in enumerate(cd.Features, start=1):
            row = {"index": i, "name": f.Name, "type": com_type_name(f)}
            try:
                row["suppressed"] = bool(f.Suppressed)
            except Exception:
                pass
            out.append(row)
        return out
    return run(job)


@mcp.tool()
def capture_view(width: int = 1024, height: int = 768, fit: bool = False) -> Image:
    """Aktif görünümün ekran görüntüsü (PNG). fit=True ise önce modeli ekrana sığdırır."""
    path = os.path.join(tempfile.gettempdir(), f"inventor_view_{int(time.time() * 1000)}.png")

    def job(app):
        view = app.ActiveView
        if fit:
            view.Fit()
        view.SaveAsBitmap(path, width, height)
        return path
    run(job)
    return Image(path=path)


@mcp.tool()
def set_parameter(name: str, value: float, unit: str = "mm") -> dict:
    """
    Parametre değerini değiştirir ve modeli günceller (tek transaction; hata olursa geri alınır).
    unit: mm, cm, m, in, deg, rad, ul. Önce get_parameters ile doğru parametreyi bul.
    """
    args = {"name": name, "value": value, "unit": unit}

    def job(app):
        doc, cd = active_part(app)
        p = cd.Parameters.Item(name)
        before = {"value": to_display(p.Value, p.Units), "unit": p.Units, **summary(cd)}
        tr = app.TransactionManager.StartTransaction(doc, f"LLM: {name} = {value} {unit}")
        try:
            p.Value = to_internal(value, unit)
            doc.Update()
            after = {"value": to_display(p.Value, p.Units), "unit": p.Units, **summary(cd)}
            tr.End()
        except Exception:
            tr.Abort()
            doc.Update()
            raise
        return {"parameter": name, "before": before, "after": after}

    try:
        res = run(job)
        log_call("set_parameter", args, res)
        return res
    except Exception as e:
        log_call("set_parameter", args, error=str(e))
        raise


@mcp.tool()
def undo_last() -> str:
    """Inventor'daki son işlemi geri alır (Edit > Undo ile aynı)."""
    def job(app):
        app.TransactionManager.UndoTransaction()
        app.ActiveDocument.Update()
        return "Geri alındı."
    res = run(job)
    log_call("undo_last", {}, res)
    return res


if __name__ == "__main__":
    mcp.run()  # stdio
