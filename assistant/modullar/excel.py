"""Excel analizi modulu: istənilən (o cümlədən çox böyük) Excel faylını oxuyur və suallara cavab verir.

Böyük fayllar üçün (məs. 1C bazasının 300+ vərəqli, milyonlarla sətirli ixracı) fayl bir dəfə
lokal SQLite bazasına çevrilir və keşdə saxlanılır. Eyni fayl yenidən yüklənəndə çevirmə təkrarlanmır.

Çevirmə zamanı:
  - hər vərəqdə başlıq sətri avtomatik tapılır (ilk 15 sətir arasında);
  - sütun tipləri (rəqəm, tarix, mətn) nümunəyə baxaraq təyin olunur ("1 234,56" yazılışları da rəqəmdir);
  - hər sətir üçün axtarış mətni (_axtar) və adlar lüğəti (_lugat) qurulur ki, suallar tez cavablansın.
Bütün hesablamaları SQL/pandas edir; AI yalnız hansı aləti çağıracağını seçir.
"""
import datetime as dt
import difflib
import hashlib
import json
import pathlib
import re
import sqlite3

import pandas as pd

from uzlasma.config import az

TR = str.maketrans("əüöığşçı", "euoigsci")
NUMUNE = 2000          # tip təyini üçün nümunə sətir sayı
LUGAT_MAKS_SETIR = 60_000   # bundan böyük vərəqlərin dəyərləri lüğətə düşmür (jurnal/reyestrlər)
LUGAT_MAKS_DEYER = 20_000   # bir sütundan lüğətə düşən maksimum fərqli dəyər
STOP = {"true", "false", "none", "nan", "yoxdur", "var", "beli", "xeyr"}


def norm(s):
    return str(s).replace("İ", "i").replace("I", "ı").lower().translate(TR).strip()


def _sozler(s):
    return re.findall(r"\w+", norm(s))


# ======================= çevirmə: Excel -> SQLite =======================

def _bos(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _basliq(setirler):
    """İlk sətirlər arasından başlıq sətrini seçir: ən çox doldurulmuş, əsasən mətndən ibarət."""
    en, xal = 0, -1
    gen = max((len(r) for r in setirler), default=0)
    for i, r in enumerate(setirler):
        dolu = sum(not _bos(v) for v in r)
        metn = sum(isinstance(v, str) and v.strip() != "" for v in r)
        x = metn * 2 + dolu
        if dolu >= max(2, gen * 0.5) and metn >= dolu * 0.6 and x > xal:
            en, xal = i, x
    return en


def _unikal(adlar):
    gorulen, out = {}, []
    for i, a in enumerate(adlar):
        a = str(a).strip() if not _bos(a) else f"Sütun {i + 1}"
        if a in gorulen:
            gorulen[a] += 1
            a = f"{a} ({gorulen[a]})"
        else:
            gorulen[a] = 1
        out.append(a)
    return out


_TARIX_RE = re.compile(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$")


def _reqem(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    t = v.replace("\u00a0", "").replace(" ", "")
    t = re.sub(r"(AZN|₼|\$|€|%)$", "", t)
    if re.search(r"\d\.\d{3},", t):
        t = t.replace(".", "")
    elif re.search(r"\d,\d{3}\.", t):
        t = t.replace(",", "")
    t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def _tarix(v):
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        m = _TARIX_RE.match(v.strip())
        if m:
            d, mo, y = map(int, m.groups())
            y = y + 2000 if y < 100 else y
            try:
                return dt.date(y, mo, d).strftime("%Y-%m-%d")
            except ValueError:
                return None
    return None


def _tip(deyerler):
    dolu = [v for v in deyerler if not _bos(v)]
    if not dolu:
        return "boş"
    if sum(isinstance(v, (dt.date, dt.datetime)) for v in dolu) >= 0.85 * len(dolu):
        return "tarix"
    if sum(_reqem(v) is not None for v in dolu) >= 0.85 * len(dolu):
        return "rəqəm"
    if sum(isinstance(v, str) and bool(_TARIX_RE.match(v.strip())) for v in dolu) >= 0.85 * len(dolu):
        return "tarix"
    return "mətn"


def fayl_hash(fayl):
    h = hashlib.sha1()
    with open(fayl, "rb") as f:
        for blok in iter(lambda: f.read(1 << 20), b""):
            h.update(blok)
    return h.hexdigest()[:16]


def _cevir(fayl, db_yolu, progress=None):
    from openpyxl import load_workbook
    wb = load_workbook(fayl, read_only=True, data_only=True)
    db = sqlite3.connect(db_yolu)
    db.execute("PRAGMA journal_mode=OFF")
    db.execute("PRAGMA synchronous=OFF")
    db.execute("CREATE TABLE _lugat (acar TEXT, deyer TEXT, vereq TEXT, sutun TEXT)")
    meta = {}
    vereqler = wb.worksheets
    for idx, ws in enumerate(vereqler):
        if progress:
            progress(idx, len(vereqler), ws.title)
        it = ws.iter_rows(values_only=True)
        bas, nomre = [], 0
        for r in it:  # ilk 15 boş olmayan sətir
            nomre += 1
            if any(not _bos(v) for v in r):
                bas.append((nomre, r))
            if len(bas) >= 15:
                break
        if not bas:
            continue
        b = _basliq([r for _, r in bas])
        bnomre, brow = bas[b]
        gen = max(i + 1 for _, r in bas for i, v in enumerate(r) if not _bos(v))
        adlar = _unikal(list(brow[:gen]) + [None] * (gen - len(brow)))
        cedvel = f"v{len(meta)}"
        sql_cols = [f"c{i}" for i in range(gen)]

        def setirler():
            yield from bas[b + 1:]
            n = nomre
            for r in it:
                n += 1
                if any(not _bos(v) for v in r):
                    yield n, r

        gen_iter = setirler()
        numune = []
        for x in gen_iter:
            numune.append(x)
            if len(numune) >= NUMUNE:
                break
        tipler = [_tip([(r[i] if i < len(r) else None) for _, r in numune]) for i in range(gen)]
        db.execute(f'CREATE TABLE {cedvel} (_setir INTEGER, _axtar TEXT, {", ".join(sql_cols)})')
        metn_idx = [i for i, t in enumerate(tipler) if t == "mətn"]
        pozuq = [0] * gen
        lugat = {i: set() for i in metn_idx}
        say = 0

        def cevir_setir(n, r):
            out = [n, None]
            axtar = []
            for i in range(gen):
                v = r[i] if i < len(r) else None
                t = tipler[i]
                if _bos(v):
                    out.append(None)
                    continue
                if t == "rəqəm":
                    x = _reqem(v)
                    if x is None:
                        pozuq[i] += 1
                    out.append(x)
                elif t == "tarix":
                    out.append(_tarix(v) or str(v))
                else:
                    sv = str(v).strip()
                    out.append(sv)
                    nv = norm(sv)
                    axtar.append(nv)
                    if i in lugat and len(lugat[i]) < LUGAT_MAKS_DEYER and 3 <= len(nv) <= 80 and nv not in STOP \
                            and not re.fullmatch(r"[\d\s.,:/-]+", nv) and not nv.startswith("[binar"):
                        lugat[i].add((nv, sv))
            out[1] = " | ".join(axtar)
            return out

        sql = f'INSERT INTO {cedvel} VALUES ({",".join("?" * (gen + 2))})'
        bufer = [cevir_setir(n, r) for n, r in numune]
        say += len(bufer)
        db.executemany(sql, bufer)
        bufer = []
        for n, r in gen_iter:
            bufer.append(cevir_setir(n, r))
            if len(bufer) >= 20_000:
                db.executemany(sql, bufer)
                say += len(bufer)
                bufer = []
        if bufer:
            db.executemany(sql, bufer)
            say += len(bufer)
        if say <= LUGAT_MAKS_SETIR:
            for i, deyerler in lugat.items():
                rows = []
                for nv, sv in deyerler:
                    rows.append((nv, sv, ws.title, adlar[i]))
                    sz = nv.split()
                    if len(sz) >= 3:  # "Kür Avto MMC" -> "kur avto" ilə də tapılsın
                        rows.append((" ".join(sz[:2]), sv, ws.title, adlar[i]))
                db.executemany("INSERT INTO _lugat VALUES (?,?,?,?)", rows)
        meta[ws.title] = {"cedvel": cedvel, "setir": say, "basliq_setri": bnomre,
                          "sutunlar": adlar, "sql": sql_cols, "tipler": dict(zip(adlar, tipler)),
                          "metn_reqem": {adlar[i]: c for i, c in enumerate(pozuq) if c and tipler[i] == "rəqəm"}}
    db.execute("CREATE INDEX _lugat_acar ON _lugat(acar)")
    db.execute("CREATE TABLE _meta (j TEXT)")
    db.execute("INSERT INTO _meta VALUES (?)", (json.dumps(meta, ensure_ascii=False),))
    db.commit()
    db.close()
    wb.close()
    if not meta:
        raise ValueError("Faylda məlumat olan vərəq tapılmadı.")
    return meta


def ac(fayl, kes_qovlugu, ad=None, progress=None):
    """Faylı keşdən açır və ya çevirir. Qaytarır: {"ad", "db", "vereqler": meta}."""
    kes = pathlib.Path(kes_qovlugu)
    kes.mkdir(parents=True, exist_ok=True)
    db_yolu = kes / f"{fayl_hash(fayl)}.db"
    meta = None
    if db_yolu.exists():
        try:
            con = sqlite3.connect(db_yolu)
            meta = json.loads(con.execute("SELECT j FROM _meta").fetchone()[0])
            con.close()
        except Exception:
            meta = None
            db_yolu.unlink(missing_ok=True)
    if meta is None:
        tmp = db_yolu.with_suffix(".tmp")
        tmp.unlink(missing_ok=True)
        meta = _cevir(str(fayl), str(tmp), progress)
        tmp.replace(db_yolu)
    return {"ad": ad or pathlib.Path(fayl).name, "db": str(db_yolu), "vereqler": meta}


# ======================= köməkçilər =======================

def _con(s):
    con = sqlite3.connect(s.excel["db"])
    con.create_function("aznorm", 1, lambda x: None if x is None else norm(x), deterministic=True)
    return con


def _sorgu(s, sql, args=()):
    with _con(s) as con:
        return pd.read_sql_query(sql, con, params=args)


def _soz_uygun(k, w):
    """Vərəq adındakı söz (k) sorğudakı sözlə (w) uyğundur: kök, şəkilçi və ya kiçik yazı fərqi."""
    if w.startswith(k[:5]) or (len(w) >= 4 and k.startswith(w)):
        return True
    return len(k) >= 5 and len(w) >= 5 and difflib.SequenceMatcher(None, k[:8], w[:8]).ratio() >= 0.8


def _vereq_xali(ad, sozler):
    vs = [w for w in _sozler(ad) if len(w) >= 3]
    if not vs:
        return 0
    tutan = sum(any(_soz_uygun(k, w) for w in sozler) for k in vs)
    return tutan / len(vs) + (0.01 * tutan)


def _v(s, vereq):
    """Vərəqi ada görə tapır (dəqiq, hissəvi, sözlərə görə); tapılmasa, ən böyük vərəq."""
    vv = s.excel["vereqler"]
    if vereq:
        n = norm(vereq)
        for ad in vv:
            if norm(ad) == n:
                return ad, vv[ad]
        for ad in vv:
            if n in norm(ad):
                return ad, vv[ad]
        xal = max(vv, key=lambda a: _vereq_xali(a, _sozler(vereq)))
        if _vereq_xali(xal, _sozler(vereq)) >= 0.5:
            return xal, vv[xal]
    ad = max(vv, key=lambda a: vv[a]["setir"])
    return ad, vv[ad]


def _sutun(v, ad):
    """Sütunu təxmini ada görə tapır: (göstərilən ad, SQL adı)."""
    if not ad:
        return None, None
    n = norm(ad)
    adlar = v["sutunlar"]
    sec = next((c for c in adlar if norm(c) == n), None) or next((c for c in adlar if n in norm(c) or norm(c) in n), None)
    if sec is None:
        soz = set(_sozler(ad))
        xal = [(len(soz & set(_sozler(c))), c) for c in adlar]
        x, c = max(xal, key=lambda t: t[0]) if xal else (0, None)
        sec = c if x else None
    return (sec, v["sql"][adlar.index(sec)]) if sec else (None, None)


def _f(x, tip=None):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(x, str):
        if tip == "tarix" and re.match(r"^\d{4}-\d{2}-\d{2}", x):
            return f"{x[8:10]}.{x[5:7]}.{x[:4]}" + x[10:16]
        return x
    if isinstance(x, bool):
        return str(x)
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return str(x)
    return f"{int(xf):,}".replace(",", " ") if xf == int(xf) else az(xf)


def _df_goster(v, df):
    """SQL nəticəsini göstərilən adlara çevirir və formatlayır."""
    ad_map = dict(zip(v["sql"], v["sutunlar"]))
    df = df.rename(columns={**ad_map, "_setir": "Sətir"})
    return df


def _cedvel(basliq, df, tipler=None, n=50):
    tipler = tipler or {}
    return {"tip": "cedvel", "basliq": basliq, "sutunlar": [str(c) for c in df.columns],
            "setirler": [[_f(x, tipler.get(c)) for c, x in zip(df.columns, row)] for row in df.head(n).itertuples(index=False)]}


def _yoxdur(s):
    return {"metn": "Hələ Excel faylı yüklənməyib. 📎 düyməsi ilə istənilən Excel faylını əlavə edin.", "ui": None}


def _sutun_sec(v, maks=12):
    """Çox sütunlu vərəqlərdə göstərmək üçün ən mənalı sütunlar (boş və texniki sütunlar sona)."""
    texniki = re.compile(r"metadata|silinm|valideyn|^qrup$|^kod$|logotip|ссылка|версия|предопредел", re.I)
    sira = sorted(range(len(v["sutunlar"])), key=lambda i: (v["tipler"][v["sutunlar"][i]] == "boş",
                                                             bool(texniki.search(v["sutunlar"][i])), i))
    return [v["sql"][i] for i in sorted(sira[:maks])]


# ======================= alətlər =======================

def excel_xulase(s):
    if not s.excel:
        return _yoxdur(s)
    vv = s.excel["vereqler"]
    cem = sum(v["setir"] for v in vv.values())
    sirali = sorted(vv.items(), key=lambda kv: -kv[1]["setir"])
    setirler = []
    for ad, v in sirali[:60]:
        reqem = [c for c, t in v["tipler"].items() if t == "rəqəm"]
        setirler.append([ad, v["setir"], len(v["sutunlar"]), ", ".join(reqem[:3]) + ("…" if len(reqem) > 3 else "")])
    metn = [f"Fayl: {s.excel['ad']}, {len(vv)} vərəq, {cem} sətir."]
    if len(vv) <= 30:
        for ad, v in vv.items():
            metn.append(f"- «{ad}»: {v['setir']} sətir; sütunlar: " +
                        ", ".join(f"{c} ({v['tipler'][c]})" for c in v["sutunlar"][:40]))
    else:
        metn.append("Vərəqlər (sətir sayı): " + "; ".join(f"{a} ({v['setir']})" for a, v in sirali[:200]))
        metn.append("Vərəqin sütunlarını bilmək üçün 'sutunlar' alətini, vərəq adını tapmaq üçün 'vereq_tap' alətini çağır.")
    basliq = f"{s.excel['ad']} · {len(vv)} vərəq · {_f(cem)} sətir" + (" · ən böyük 60 vərəq" if len(vv) > 60 else "")
    return {"metn": "\n".join(metn), "ui": {"tip": "cedvel", "basliq": basliq,
            "sutunlar": ["Vərəq", "Sətir", "Sütun", "Rəqəm sütunları"], "setirler": [[a, _f(b), c, d] for a, b, c, d in setirler]}}


def vereq_tap(s, sorgu):
    if not s.excel:
        return _yoxdur(s)
    vv = s.excel["vereqler"]
    soz = _sozler(sorgu)
    xal = sorted(((_vereq_xali(a, soz), a) for a in vv), reverse=True)
    sec = [a for x, a in xal if x > 0][:20]
    if not sec:
        return {"metn": f"«{sorgu}» adına uyğun vərəq tapılmadı.", "ui": None}
    setirler = [[a, _f(vv[a]["setir"]), ", ".join(vv[a]["sutunlar"][:8])] for a in sec]
    return {"metn": "Uyğun vərəqlər:\n" + "\n".join(f"- {a} ({vv[a]['setir']} sətir): {', '.join(vv[a]['sutunlar'][:12])}" for a in sec),
            "ui": {"tip": "cedvel", "basliq": f"«{sorgu}» üzrə vərəqlər", "sutunlar": ["Vərəq", "Sətir", "Sütunlar"], "setirler": setirler}}


def sutunlar(s, vereq=None):
    if not s.excel:
        return _yoxdur(s)
    ad, v = _v(s, vereq)
    t = v["cedvel"]
    setirler, metn = [], []
    for c, q in zip(v["sutunlar"], v["sql"]):
        tip = v["tipler"][c]
        dolu, unikal = _sorgu(s, f"SELECT count({q}), count(DISTINCT {q}) FROM {t}").iloc[0]
        info = ""
        if tip == "rəqəm" and dolu:
            r = _sorgu(s, f"SELECT sum({q}), avg({q}), min({q}), max({q}) FROM {t}").iloc[0]
            info = f"cəm {_f(r.iloc[0])} · orta {_f(r.iloc[1])} · min {_f(r.iloc[2])} · maks {_f(r.iloc[3])}"
        elif tip == "tarix" and dolu:
            r = _sorgu(s, f"SELECT min({q}), max({q}) FROM {t}").iloc[0]
            info = f"{_f(r.iloc[0], 'tarix')} — {_f(r.iloc[1], 'tarix')}"
        elif tip == "mətn" and dolu:
            top = _sorgu(s, f"SELECT {q} d, count(*) n FROM {t} WHERE {q} IS NOT NULL GROUP BY {q} ORDER BY n DESC LIMIT 3")
            info = "; ".join(f"{a} ({b})" for a, b in top.itertuples(index=False))
        doluluq = f"{dolu / v['setir']:.0%}" if v["setir"] else "—"
        setirler.append([c, tip, doluluq, int(unikal), info[:120]])
        metn.append(f"- {c} [{tip}], dolu {doluluq}, unikal {int(unikal)}. {info[:150]}")
    return {"metn": f"«{ad}» vərəqi ({v['setir']} sətir), sütunlar:\n" + "\n".join(metn),
            "ui": {"tip": "cedvel", "basliq": f"«{ad}» sütunları · {_f(v['setir'])} sətir",
                   "sutunlar": ["Sütun", "Tip", "Dolu", "Unikal", "Qısa məlumat"], "setirler": setirler}}


def setirler(s, vereq=None, n=20, sutun=None, deyer=None):
    if not s.excel:
        return _yoxdur(s)
    ad, v = _v(s, vereq)
    n = int(n or 20)
    cols = _sutun_sec(v)
    where, args, basliq = "", (), f"«{ad}» · ilk {min(n, v['setir'])} sətir (cəmi {_f(v['setir'])})"
    xulase = f"«{ad}» vərəqində {_f(v['setir'])} sətir var" + (f", ilk {n}-i aşağıdadır." if v["setir"] > n else ".")
    if sutun and deyer is not None:
        c, q = _sutun(v, sutun)
        if c is None:
            return {"metn": f"«{sutun}» sütunu tapılmadı. Sütunlar: {', '.join(v['sutunlar'])}", "ui": None}
        where, args = f"WHERE aznorm({q}) LIKE ?", (f"%{norm(deyer)}%",)
        say = _sorgu(s, f"SELECT count(*) FROM {v['cedvel']} {where}", args).iloc[0, 0]
        basliq = f"«{ad}» · {c} ~ «{deyer}» · {_f(say)} sətir"
        xulase = f"«{ad}» vərəqində {c} sütununda «{deyer}» olan {_f(say)} sətir var."
    df = _sorgu(s, f"SELECT _setir, {', '.join(cols)} FROM {v['cedvel']} {where} LIMIT {n}", args)
    df = _df_goster(v, df)
    return {"metn": f"{xulase}\n" + df.head(15).to_string(index=False, max_colwidth=30)[:3000],
            "ui": _cedvel(basliq, df, v["tipler"], n)}


def qrupla(s, deyer_sutunu, qrup_sutunu=None, emeliyyat="cem", vereq=None, n=20):
    if not s.excel:
        return _yoxdur(s)
    ad, v = _v(s, vereq)
    dc, dq = _sutun(v, deyer_sutunu)
    if dc is None:
        return {"metn": f"«{deyer_sutunu}» sütunu «{ad}» vərəqində tapılmadı. Sütunlar: {', '.join(v['sutunlar'])}", "ui": None}
    em = {"cem": "sum", "cəm": "sum", "orta": "avg", "say": "count", "min": "min", "maks": "max", "max": "max"}.get(norm(emeliyyat), "sum")
    em_ad = {"sum": "Cəm", "avg": "Orta", "count": "Say", "min": "Min", "max": "Maks"}[em]
    if em != "count" and v["tipler"].get(dc) != "rəqəm":
        return {"metn": f"«{dc}» rəqəm sütunu deyil, ona {em_ad.lower()} tətbiq etmək olmaz.", "ui": None}
    t = v["cedvel"]
    if not qrup_sutunu:
        val = _sorgu(s, f"SELECT {em}({dq}) FROM {t}").iloc[0, 0]
        return {"metn": f"«{ad}» · {dc}: {em_ad} = {_f(val)}",
                "ui": {"tip": "kartlar", "kartlar": [{"ad": f"{em_ad} · {dc}", "deyer": _f(val), "alt": f"«{ad}» · {_f(v['setir'])} sətir"}]}}
    qc, qq = _sutun(v, qrup_sutunu)
    if qc is None:
        return {"metn": f"«{qrup_sutunu}» sütunu tapılmadı. Sütunlar: {', '.join(v['sutunlar'])}", "ui": None}
    qifade = f"substr({qq},1,7)" if v["tipler"].get(qc) == "tarix" else qq
    df = _sorgu(s, f"SELECT {qifade} AS q, {em}({dq}) AS d FROM {t} GROUP BY q ORDER BY d DESC LIMIT {int(n or 20)}")
    df.columns = [qc, f"{em_ad} · {dc}"]
    return {"metn": f"«{ad}»: {dc} {em_ad.lower()}, {qc} üzrə:\n" + "\n".join(f"- {a}: {_f(b)}" for a, b in df.itertuples(index=False)),
            "ui": _cedvel(f"{dc} · {em_ad.lower()} · {qc} üzrə", df, {}, int(n or 20))}


def en_boyuk(s, sutun, n=10, artan=False, vereq=None):
    if not s.excel:
        return _yoxdur(s)
    ad, v = _v(s, vereq)
    c, q = _sutun(v, sutun)
    if c is None:
        return {"metn": f"«{sutun}» sütunu «{ad}» vərəqində tapılmadı.", "ui": None}
    yon = "ASC" if artan else "DESC"
    cols = _sutun_sec(v)
    if q not in cols:
        cols = [q] + cols[:-1]
    df = _sorgu(s, f"SELECT _setir, {', '.join(cols)} FROM {v['cedvel']} WHERE {q} IS NOT NULL ORDER BY {q} {yon} LIMIT {int(n or 10)}")
    df = _df_goster(v, df)
    soz = "ən kiçik" if artan else "ən böyük"
    return {"metn": f"«{ad}»: {c} üzrə {soz} {len(df)} sətir:\n" + df.to_string(index=False, max_colwidth=25)[:3000],
            "ui": _cedvel(f"«{ad}» · {c} üzrə {soz} {len(df)}", df, v["tipler"], int(n or 10))}


def _hits(s, obyekt, yalniz=None, limit=30):
    """Obyektin adı keçən sətirlər, vərəqlər üzrə: [(vərəq, say, DataFrame nümunə)]."""
    n = f"%{norm(obyekt)}%"
    out = []
    with _con(s) as con:
        for ad, v in s.excel["vereqler"].items():
            if yalniz and ad != yalniz:
                continue
            say = con.execute(f"SELECT count(*) FROM {v['cedvel']} WHERE _axtar LIKE ?", (n,)).fetchone()[0]
            if say:
                out.append((ad, say))
    return out


def axtar(s, sorgu):
    if not s.excel:
        return _yoxdur(s)
    hits = sorted(_hits(s, sorgu), key=lambda x: -x[1])
    if not hits:
        return {"metn": f"«{sorgu}» heç bir vərəqdə tapılmadı.", "ui": None}
    ad, say = hits[0]
    v = s.excel["vereqler"][ad]
    df = _sorgu(s, f"SELECT _setir, {', '.join(_sutun_sec(v))} FROM {v['cedvel']} WHERE _axtar LIKE ? LIMIT 30", (f"%{norm(sorgu)}%",))
    df = _df_goster(v, df)
    metn = f"«{sorgu}» {len(hits)} vərəqdə tapıldı: " + "; ".join(f"{a} ({b})" for a, b in hits[:15]) + \
           f".\n«{ad}» vərəqindən nümunə:\n" + df.head(10).to_string(index=False, max_colwidth=25)[:1500]
    return {"metn": metn, "ui": _cedvel(f"«{ad}» · «{sorgu}» · {_f(say)} sətir" +
                                       (f" · daha {len(hits) - 1} vərəqdə də var" if len(hits) > 1 else ""), df, v["tipler"], 30)}


def problemler(s, vereq=None):
    if not s.excel:
        return _yoxdur(s)
    vv = s.excel["vereqler"]
    if not vereq and len(vv) > 10:
        en = sorted(vv, key=lambda a: -vv[a]["setir"])[:10]
        return {"metn": f"Faylda {len(vv)} vərəq var. Hansı vərəqi yoxlayım? Məsələn: " + ", ".join(en[:5]), "ui": None}
    secilen = [_v(s, vereq)[0]] if vereq else list(vv)
    out = []
    for ad in secilen:
        v = vv[ad]
        t, cols = v["cedvel"], v["sql"]
        dup = _sorgu(s, f"SELECT count(*) - count(DISTINCT {'||char(31)||'.join(f'ifnull({c},char(0))' for c in cols)}) FROM {t}").iloc[0, 0]
        if dup:
            nomreler = _sorgu(s, f"SELECT _setir FROM {t} WHERE rowid NOT IN (SELECT min(rowid) FROM {t} GROUP BY {', '.join(cols)}) LIMIT 8")
            out.append(["Orta", ad, "Təkrar sətirlər", f"{int(dup)} sətir tam təkrarlanır (Excel sətirləri: " +
                        ", ".join(str(x) for x in nomreler["_setir"]) + ")"])
        for c, q in zip(v["sutunlar"], cols):
            tip = v["tipler"][c]
            if tip == "boş" or not v["setir"]:
                continue
            bos = _sorgu(s, f"SELECT count(*) FROM {t} WHERE {q} IS NULL").iloc[0, 0]
            if 0 < bos <= v["setir"] * 0.3:
                out.append(["Aşağı", ad, f"Boş xanalar · {c}", f"{int(bos)} sətirdə boşdur (sütunun {bos / v['setir']:.0%}-i)"])
            if tip == "rəqəm":
                x = _sorgu(s, f"SELECT {q} x FROM {t} WHERE {q} IS NOT NULL LIMIT 300000")["x"]
                if len(x) >= 8:
                    q1, q3 = x.quantile(.25), x.quantile(.75)
                    iqr = q3 - q1
                    if iqr > 0:
                        kenar = x[(x < q1 - 3 * iqr) | (x > q3 + 3 * iqr)]
                        if len(kenar) and len(kenar) <= len(x) * 0.05:
                            out.append(["Orta", ad, f"Qeyri-adi dəyərlər · {c}",
                                        f"{len(kenar)} dəyər ümumi səviyyədən çox fərqlənir: " + ", ".join(_f(k) for k in kenar.head(5))])
                    menfi = x[x < 0]
                    if 0 < len(menfi) <= len(x) * 0.1:
                        out.append(["Orta", ad, f"Mənfi dəyərlər · {c}", f"{len(menfi)} mənfi dəyər var, qalanları müsbətdir"])
                pozuq = v.get("metn_reqem", {}).get(c)
                if pozuq:
                    out.append(["Orta", ad, f"Rəqəm sütununda mətn · {c}", f"{pozuq} xana rəqəm kimi oxunmadı (yazı xətası ola bilər)"])
            if tip == "tarix":
                gel = (dt.date.today() + dt.timedelta(days=366)).isoformat()
                r = _sorgu(s, f"SELECT sum({q} > ?), sum({q} < '1990-01-01' AND {q} >= '0002') FROM {t}", (gel,)).iloc[0]
                if (r.iloc[0] or 0) or (r.iloc[1] or 0):
                    out.append(["Orta", ad, f"Şübhəli tarixlər · {c}", f"{int(r.iloc[0] or 0)} çox uzaq gələcək, {int(r.iloc[1] or 0)} 1990-dan əvvəl"])
    sira = {"Yüksək": 0, "Orta": 1, "Aşağı": 2}
    out.sort(key=lambda r: sira.get(r[0], 9))
    if not out:
        return {"metn": "Struktur problemi tapılmadı: təkrar sətir, şübhəli dəyər və ya çox boş xana yoxdur.", "ui": None}
    return {"metn": f"{len(out)} problem tapıldı:\n" + "\n".join(f"- [{r[0]}] {r[1]} · {r[2]}: {r[3]}" for r in out[:30]),
            "ui": {"tip": "cedvel", "basliq": "Tapılan problemlər", "sutunlar": ["Risk", "Vərəq", "Problem", "İzah"], "setirler": out[:100]}}


HEREKET_SOZLERI = {"goster", "gostar", "ver", "siyahi", "siyahisi", "siyahisini", "bax", "hamisi", "butun", "list", "cixar",
                   "mene", "bize", "olan", "edin", "et", "ne", "var", "yox", "yaz", "tap", "zehmet", "olmasa", "lutfen", "acin",
                   "ac", "nece", "nedir", "necedir", "hansi", "harada", "ucun", "ile", "olaraq"}


# Gündəlik söz -> 1C/uçot terminləri (vərəq adlarında axtarılır)
MOVZU_SINONIM = {
    "maas": ["emek", "haqq"], "isci": ["isci", "emekdas"], "emekdas": ["emekdas", "isci"], "personal": ["isci", "emekdas", "kadr"],
    "musteri": ["kontragent", "alici"], "techizatci": ["kontragent", "malsatan"], "mal": ["nomenklatura", "mal"],
    "mehsul": ["nomenklatura", "mehsul"], "mezuniyyet": ["mezuniyyet"], "mukafat": ["mukafat"], "hesab": ["hesab"],
}


SEKILCILER = {"", "lar", "ler", "lari", "leri", "larin", "lerin", "i", "u", "in", "un", "nin", "nun", "a", "e", "da", "de",
              "dan", "den", "ya", "ye", "na", "ne", "si", "su", "sin", "sun", "lara", "lere", "larda", "lerde"}


def _movzu_uygun(k, w):
    """Ciddi uyğunluq: eyni kök (5 hərf) və ya qısa söz + Azərbaycan şəkilçisi ("bank" ~ "banklari", amma "emek" ≁ "emekdas")."""
    if min(len(k), len(w)) >= 5:
        return k[:5] == w[:5]
    qisa, uzun = sorted((k, w), key=len)
    return uzun.startswith(qisa) and uzun[len(qisa):] in SEKILCILER


def _movzu_sozleri(sorgu):
    return [w for w in _sozler(sorgu) if len(w) >= 3 and w not in HEREKET_SOZLERI]


def _movzu_vereqleri(vv, msoz):
    """Mövzu sözlərinə (və sinonimlərinə) adı uyğun gələn vərəqlər, ən uyğunu birinci."""
    qruplar = []  # hər mövzu sözü üçün: özü + sinonimləri
    for w in msoz:
        qruplar.append([w] + next((v for k, v in MOVZU_SINONIM.items() if w.startswith(k)), []))
    if not qruplar:
        return []

    def xal(ad):
        vs = [k for k in _sozler(ad) if len(k) >= 3]
        tutan = sum(any(_movzu_uygun(k, w) for k in vs for w in qrup) for qrup in qruplar)
        return tutan / len(qruplar)
    return sorted((a for a in vv if xal(a) > 0),
                  key=lambda a: (-xal(a), "›" in a, bool(re.search(r"\(\d+\)$", a)), len(a), -vv[a]["setir"]))


def movzu_goster(s, sorgu, n=20):
    """Mövzuya (xərclər, avanslar, məzuniyyət...) uyğun vərəqi tapıb göstərir; uyğun digər vərəqləri də sadalayır."""
    if not s.excel:
        return _yoxdur(s)
    vv = s.excel["vereqler"]
    msoz = _movzu_sozleri(sorgu)
    if not msoz:
        return {"metn": "Nəyi göstərim? Mövzunu yazın, məsələn: «avans hesabatları», «xərc maddələri», «banklar».", "ui": None}

    namized = _movzu_vereqleri(vv, msoz)
    if not namized:
        # vərəq adı yoxdur: dəyərlərdə axtar
        netice = axtar(s, " ".join(msoz))
        if netice["ui"]:
            return netice
        return {"metn": f"«{' '.join(msoz)}» mövzusunda vərəq və ya dəyər tapılmadı. Fayldakı vərəqləri görmək üçün «Fayl haqqında ümumi məlumat» yazın.", "ui": None}
    ad = namized[0]
    netice = setirler(s, vereq=ad, n=n)
    diger = namized[1:12]
    if diger:
        netice["metn"] = netice["metn"].split("\n")[0] + f" Bu mövzuda daha {len(namized) - 1} vərəq var: " + \
            ", ".join(f"«{a}»" for a in diger) + (" və s." if len(namized) > 12 else "") + \
            ". Başqasını görmək üçün adını yazın.\n" + "\n".join(netice["metn"].split("\n")[1:])
    return netice


# Göstərici sözlərinin sinonimləri: sorğudakı söz -> sütun adında axtarılan köklər
SINONIM = {
    "borc": ["borc", "qaliq", "debitor", "kreditor"], "qaliq": ["qaliq", "borc"], "mebleg": ["mebleg", "cem", "deyer"],
    "odenis": ["odenis", "odenilib", "odenen"], "maas": ["maas", "emek"], "emek": ["emek", "maas"],
    "miqdar": ["miqdar", "say", "eded"], "satis": ["satis", "mebleg", "gelir"], "gelir": ["gelir", "satis", "mebleg"],
    "xerc": ["xerc", "mebleg"], "qiymet": ["qiymet", "deyer"], "edv": ["edv"],
}


def _gosterici_sutunu(v, gosterici):
    kokler = set()
    for w in _sozler(gosterici or ""):
        if len(w) < 3:
            continue
        kokler.add(w[:5])
        for k, syn in SINONIM.items():
            if w.startswith(k):
                kokler.update(x[:5] for x in syn)
    en, xal = None, 0
    for c, t in v["tipler"].items():
        if t != "rəqəm":
            continue
        cw = _sozler(c)
        x = sum(2 for k in kokler if any(w.startswith(k) for w in cw))
        if x > xal:
            en, xal = c, x
    return en, xal


def obyekt_gostericisi(s, obyekt, gosterici=None, vereq=None):
    if not s.excel:
        return _yoxdur(s)
    hits = dict(_hits(s, obyekt))
    if not hits:
        return {"metn": f"«{obyekt}» faylda tapılmadı.", "ui": None}
    vv = s.excel["vereqler"]
    like = (f"%{norm(obyekt)}%",)
    if gosterici:
        namized = []
        for ad, say in hits.items():
            c, xal = _gosterici_sutunu(vv[ad], gosterici)
            if c:
                namized.append((xal, -say, ad, c))
        if namized:
            _, _, ad, c = max(namized)
            v = vv[ad]
            q = v["sql"][v["sutunlar"].index(c)]
            cem = _sorgu(s, f"SELECT sum({q}) FROM {v['cedvel']} WHERE _axtar LIKE ?", like).iloc[0, 0]
            cols = _sutun_sec(v)
            if q not in cols:
                cols = [q] + cols[:-1]
            df = _df_goster(v, _sorgu(s, f"SELECT _setir, {', '.join(cols)} FROM {v['cedvel']} WHERE _axtar LIKE ? LIMIT 30", like))
            diger = [f"«{a}»" for a in hits if a != ad][:8]
            metn = (f"«{obyekt}» üzrə {c}: **{_f(cem) or '0'}** («{ad}» vərəqi, {hits[ad]} sətir)." +
                    (f" Bu ad həmçinin {', '.join(diger)} vərəqlərində keçir." if diger else ""))
            return {"metn": metn, "ui": _cedvel(f"«{ad}» · «{obyekt}» · {c} = {_f(cem)}", df, v["tipler"], 30)}
    if vereq:
        ad_v, _ = _v(s, vereq)
        if ad_v in hits:
            hits = {ad_v: hits[ad_v]}
    # göstəricinin sözləri vərəq adına uyğun gəlirsə (məs. "müqavilələri"), o vərəq; yoxsa ən kiçik (kataloq) vərəq
    gsoz = _sozler(gosterici or "")
    ad = max(hits, key=lambda a: (_vereq_xali(a, gsoz), -vv[a]["setir"]))
    v = vv[ad]
    df = _df_goster(v, _sorgu(s, f"SELECT _setir, {', '.join(_sutun_sec(v))} FROM {v['cedvel']} WHERE _axtar LIKE ? LIMIT 30", like))
    metn = (f"«{obyekt}» «{ad}» vərəqində {hits[ad]} sətirdə keçir" +
            (f"; ümumilikdə {len(hits)} vərəqdə: " + ", ".join(f"«{a}» ({b})" for a, b in sorted(hits.items(), key=lambda x: -x[1])[:12]) if len(hits) > 1 else "") + ".")
    return {"metn": metn, "ui": _cedvel(f"«{ad}» · «{obyekt}» · {hits[ad]} sətir", df, v["tipler"], 30)}


# ======================= cəm (ümumi xərc, gəlir, məbləğ) =======================
CEM_SOZLERI = {"umumi", "cemi", "cem", "toplam", "qeder", "qederdir", "nece", "necedir", "sirket", "sirketin",
               "sirketimizin", "sirketimiz", "butun", "hamisi", "nedir", "hesabla", "hesablayin", "olub", "edib", "ili",
               "ilde", "ildə", "cemini", "umumilikde", "deyer"}
# Uçot hesablar planı (AR): 7-ci bölmə xərclər (Debet), 6-cı bölmə gəlirlər (Kredit)
HESAB_SINIFI = {"xerc": ("dt", "7", "xərci"), "gelir": ("ct", "6", "gəliri")}
TEXNIKI_REQEM = re.compile(r"setir|kod|nomre|№|sira|ardicil|novbe|prioritet|miqdar|kemiyyet|faiz|derece|splitter|lineno|valideyn|nov", re.I)


def _mebleg_sutunu(v, sorgu=""):
    """Cəmlənəcək rəqəm sütunu: əvvəl «Məbləğ/Cəm/Dəyər» tipli və sorğuya uyğun, sonra ilk məbləğ sütunu, sonra sorğuya uyğun istənilən."""
    reqem = [c for c in v["sutunlar"] if v["tipler"].get(c) == "rəqəm" and not TEXNIKI_REQEM.search(norm(c))]
    mebleg = [c for c in reqem if re.search(r"mebleg|cem|deyer|summa|сумма|netice|yekun|результат|итог", norm(c))]
    c, _ = _gosterici_sutunu(v, sorgu)
    if c in mebleg:
        return c
    if mebleg:
        return mebleg[0]
    return c if c in reqem else None


def _cemle(s, v, c, il=None):
    q = v["sql"][v["sutunlar"].index(c)]
    sert, args = "", ()
    tarix = next((x for x in v["sutunlar"] if v["tipler"].get(x) == "tarix"), None)
    if il and tarix:
        sert, args = f" WHERE substr({v['sql'][v['sutunlar'].index(tarix)]}, 1, 4) = ?", (il,)
    r = _sorgu(s, f"SELECT sum({q}), count({q}) FROM {v['cedvel']}{sert}", args).iloc[0]
    return (float(r.iloc[0]) if r.iloc[0] is not None and not pd.isna(r.iloc[0]) else 0.0), int(r.iloc[1]), bool(sert)


def _jurnal(s):
    """Mühasibat yazılışları vərəqi (1C «Müxabirləşmə jurnalı»): Hesab Dt, Hesab Ct, Məbləğ sütunları olan ən böyük vərəq."""
    vv = s.excel["vereqler"]
    for ad, v in sorted(vv.items(), key=lambda kv: -kv[1]["setir"]):
        cs = v["sutunlar"]
        dt = next((c for c in cs if re.fullmatch(r"(hesab|счет|schet) ?(dt|debet|дт)", norm(c))), None)
        ct = next((c for c in cs if re.fullmatch(r"(hesab|счет|schet) ?(ct|kt|kredit|кт)", norm(c))), None)
        m = next((c for c in cs if norm(c) in ("mebleg", "сумма", "summa")), None)
        if dt and ct and m:
            return ad, v, dt, ct, m
    return None


def _kod(x):
    """Hesab kodunu mətnə çevirir: 721.0 -> "721", 536.05 -> "536.05"."""
    if x is None:
        return ""
    try:
        f = float(x)
    except (TypeError, ValueError):
        return str(x).strip()
    if f != f:  # NaN
        return ""
    return str(int(f)) if f == int(f) else f"{f:.4f}".rstrip("0")


def _hesab_adlari(s):
    vv = s.excel["vereqler"]
    planlar = sorted((a for a in vv if "hesablar pla" in norm(a) and {"Kod", "Ad"} <= set(vv[a]["sutunlar"])),
                     key=lambda a: ("muhasibat" not in norm(a), len(a)))
    if not planlar:
        return {}
    v = vv[planlar[0]]
    k, a = (v["sql"][v["sutunlar"].index(x)] for x in ("Kod", "Ad"))
    out = {}
    for kod, ad in _sorgu(s, f"SELECT {k} AS k, {a} AS a FROM {v['cedvel']}").itertuples(index=False):
        out.setdefault(_kod(kod), ad)
    return out


# Mühasibin öyrətdiyi əlavə hesablar: məs. sığorta şirkətində 607 «Müstəqil ekspertlərlə hesablaşmalar» xərcə daxil edilir
AYAR_FAYLI = None  # state.py təyin edir: data/hesab_ayarlari.json
# xerc/gelir: bölmədən kənar əlavə hesablar; *_cixar: bölmədə olub daxil edilməyənlər; *_saxla: susmaya görə çıxarılanı saxla
_AYAR = {"xerc": [], "gelir": [], "xerc_cixar": [], "gelir_cixar": [], "xerc_saxla": [], "gelir_saxla": []}
KOD_RE = r"\d{1,4}(\.\d{1,3})?"


def _cixarilanlar(adlar, sinif, bol):
    """Bölmədə olub cəmə daxil edilməyən hesablar: mühasibin çıxardıqları + adı «hesablaşma» olanlar
    (məs. 607.05 «Müstəqil ekspertlərlə hesablaşmalar» gəlir deyil). Qaytarır: {kod: səbəb}."""
    ay = _ayar()
    out = {}
    for k, a in adlar.items():
        if k.startswith(sinif) and "hesablasma" in norm(a):
            out[k] = "hesablaşma hesabıdır, " + ("gəlir" if bol == "gelir" else "xərc") + " deyil"
            ust = k.split(".")[0]
            if "." in k and ust in adlar and re.match(r"(счет|schet|hesab)\b", norm(adlar[ust])):
                out[ust] = out[k]  # "Счет 607" kimi adsız ana hesab da
    for k in ay[bol + "_cixar"]:
        if re.fullmatch(KOD_RE, k):
            out[k] = "mühasibin göstərişi ilə çıxarılıb"
    for sx in ay[bol + "_saxla"]:
        for k in list(out):
            if k.startswith(sx) or sx.startswith(k):
                out.pop(k)
    # ana hesab çıxarılıbsa, alt hesabları ayrıca saymağa ehtiyac yoxdur
    return {k: v for k, v in out.items() if not any(k != u and k.startswith(u + ".") for u in out)}


def _ayar():
    if AYAR_FAYLI and pathlib.Path(AYAR_FAYLI).exists():
        try:
            d = json.loads(pathlib.Path(AYAR_FAYLI).read_text(encoding="utf-8"))
            return {k: d.get(k, []) for k in _AYAR}
        except (ValueError, OSError):
            pass
    if AYAR_FAYLI:  # fayl hələ yoxdur: təmiz ayarlar
        return {k: [] for k in _AYAR}
    return _AYAR


def _ayar_yaz(d):
    _AYAR.update(d)
    if AYAR_FAYLI:
        pathlib.Path(AYAR_FAYLI).parent.mkdir(parents=True, exist_ok=True)
        pathlib.Path(AYAR_FAYLI).write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def _kok(w):
    """Sözün kökü: "ekspertlərin" -> "ekspert" (axtarış üçün)."""
    for sek in sorted(SEKILCILER | {"lerin", "larin", "lerinin", "larinin", "leri", "lari"}, key=len, reverse=True):
        if sek and w.endswith(sek) and len(w) - len(sek) >= 4:
            return w[:len(w) - len(sek)]
    return w


FILTR_ATILAN = {"icinde", "icindeki", "daxilinde", "arasinda", "gostermir", "gorunmur", "gorunur", "niye", "neden",
                "hesabi", "hesablari", "hesablar", "terefinden", "uzerinde"}


def _jurnal_cemi(s, j, teref, sinif, ifade, il=None, filtr=()):
    """Uçot yazılışlarından ümumi xərc/gəlir: hesablar və xərc növləri (qarşı hesab) üzrə, istəsən sözlə süzülmüş."""
    ad, v, dt, ct, m = j
    sq = dict(zip(v["sutunlar"], v["sql"]))
    hes, qarsi = (sq[dt], sq[ct]) if teref == "dt" else (sq[ct], sq[dt])
    tarix = next((x for x in v["sutunlar"] if v["tipler"].get(x) == "tarix"), None)
    adlar = _hesab_adlari(s)
    bol = "xerc" if teref == "dt" else "gelir"
    elave = [h for h in _ayar()[bol] if re.fullmatch(KOD_RE, h) and not h.startswith(sinif)]
    cixar = _cixarilanlar(adlar, sinif, bol)

    # əsas şərt: 7xx (6xx) hesabları + mühasibin əlavə etdiyi hesablar (bağlanış yazılışları 8xx olmadan).
    # Əlavə hesab öz təbii işarəsi ilə girir: xərcdə Dt − Ct, gəlirdə Ct − Dt (məs. 603 «təkrarsığortaçıların payı»
    # kreditdə yığılır -> xərci azaldır).
    tbl = v["cedvel"]
    # sinif: "7"/"6" (sabit); qarşı hesab 8xx = bağlanış yazılışı (801-ə köçürmə və onun stornosu), sayılmır
    bolme_sert = f"CAST({hes} AS TEXT) LIKE '{sinif}%' AND CAST({qarsi} AS TEXT) NOT LIKE '8%'" + \
        "".join(f" AND CAST({hes} AS TEXT) NOT LIKE '{k}%'" for k in cixar)
    sertler, args = [f"({bolme_sert})"], []
    # Mühasibin "X xərcdir" göstərişi: hesabın əsas dövriyyəsi (hansı tərəfdə yığılırsa) müsbət, əks tərəfi
    # (storno/düzəliş) mənfi sayılır. Məs. 601 kreditdə yığılır -> kredit dövriyyəsi xərcə əlavə olunur.
    eks = []  # mənfi işarə ilə sayılan sətirlər
    for h in elave:
        r = _sorgu(s, f"SELECT sum(CASE WHEN CAST({hes} AS TEXT) LIKE ? THEN {sq[m]} ELSE 0 END), "
                      f"sum(CASE WHEN CAST({qarsi} AS TEXT) LIKE ? THEN {sq[m]} ELSE 0 END) FROM {tbl} "
                      f"WHERE CAST({hes} AS TEXT) NOT LIKE '8%' AND CAST({qarsi} AS TEXT) NOT LIKE '8%'", (f"{h}%", f"{h}%")).iloc[0]
        oz = f"(CAST({hes} AS TEXT) LIKE '{h}%' AND CAST({qarsi} AS TEXT) NOT LIKE '8%')"
        eks_t = f"(CAST({qarsi} AS TEXT) LIKE '{h}%' AND CAST({hes} AS TEXT) NOT LIKE '8%')"
        if float(r.iloc[1] or 0) > float(r.iloc[0] or 0):
            oz, eks_t = eks_t, oz  # əsas dövriyyə əks sütundadır
        sertler.append(oz)
        eks.append(eks_t)
    esas = "(" + " OR ".join(sertler) + ")"
    eks_sert = "(" + " OR ".join(eks) + ")" if eks else "0"
    M = f"(CASE WHEN {esas} THEN {sq[m]} ELSE -{sq[m]} END)"
    bolmede = f"(CAST({hes} AS TEXT) LIKE '{sinif}%')"
    elave_hes = " OR ".join(f"CAST({hes} AS TEXT) LIKE '{h}%'" for h in elave) or "0"
    h_ifade = f"CASE WHEN {bolmede} OR ({elave_hes}) THEN {hes} ELSE {qarsi} END"
    q_ifade = f"CASE WHEN {bolmede} OR ({elave_hes}) THEN {qarsi} ELSE {hes} END"
    sertler = [esas, eks_sert]
    sert = "(" + " OR ".join(sertler) + ")"
    if il and tarix:
        sert += f" AND substr({sq[tarix]}, 1, 4) = ?"
        args.append(il)
    else:
        il = None
    suzulmus = ""
    if filtr:
        kokler = [_kok(w) for w in filtr]
        qarsi_uygun = [k for k, a in adlar.items() if any(kk in norm(a) for kk in kokler)]
        alt = [f"_axtar LIKE ?" for _ in kokler] + [f"CAST({q_ifade} AS TEXT) LIKE ?" for _ in qarsi_uygun]
        sert += " AND (" + " OR ".join(alt) + ")"
        args += [f"%{k}%" for k in kokler] + [f"{k}%" if "." in k else f"{k}.%" for k in qarsi_uygun]
        suzulmus = " ".join(filtr)

    df = _sorgu(s, f"SELECT {h_ifade} AS h, {q_ifade} AS q, count(*) AS n, sum({M}) AS cem FROM {v['cedvel']} "
                   f"WHERE {sert} GROUP BY h, q ORDER BY cem DESC", args)
    il_metn = f" ({il}-ci il)" if il else ""

    # süzgəc sözünə uyğun, amma xərcə daxil olmayan hesablar (məs. 607.05 «Müstəqil ekspertlərlə hesablaşmalar»)
    kenar = ""
    if filtr:
        kokler = [_kok(w) for w in filtr]
        diger = [k for k, a in adlar.items() if any(kk in norm(a) for kk in kokler)
                 and not k.startswith(sinif) and not any(k.startswith(e) for e in elave)
                 and not k.startswith("5") and not k.startswith("2")]  # öhdəlik/debitor hesabları deyil
        for k in diger[:3]:
            r = _sorgu(s, f"SELECT sum(CASE WHEN CAST({hes} AS TEXT) LIKE ? THEN {sq[m]} ELSE 0 END), "
                          f"sum(CASE WHEN CAST({qarsi} AS TEXT) LIKE ? THEN {sq[m]} ELSE 0 END) FROM {v['cedvel']} "
                          f"WHERE (CAST({hes} AS TEXT) LIKE ? OR CAST({qarsi} AS TEXT) LIKE ?) AND CAST({hes} AS TEXT) NOT LIKE '8%' "
                          f"AND CAST({qarsi} AS TEXT) NOT LIKE '8%'", (f"{k}%",) * 4).iloc[0]
            dov = max(float(r.iloc[0] or 0), float(r.iloc[1] or 0))
            if not dov:
                continue
            # bu hesabın qarşı tərəfi artıq xərcin içindədirsə (məs. 607 <- 536.05 <- 721), əlavə etmək təkrar sayım olar
            qt = _sorgu(s, f"SELECT DISTINCT CASE WHEN CAST({hes} AS TEXT) LIKE ? THEN {qarsi} ELSE {hes} END FROM {v['cedvel']} "
                           f"WHERE (CAST({hes} AS TEXT) LIKE ? OR CAST({qarsi} AS TEXT) LIKE ?) AND CAST({hes} AS TEXT) NOT LIKE '8%' "
                           f"AND CAST({qarsi} AS TEXT) NOT LIKE '8%'", (f"{k}%",) * 3).iloc[:, 0].map(_kod)
            ortaq = sorted(set(qt) & set(df["q"].map(_kod))) if not df.empty else []
            kenar += f" Diqqət: «{adlar[k]}» ({k}) hesabında da {_f(dov)} dövriyyə var ({k[0]}-cı bölmə, cəmə daxil deyil)."
            if ortaq:
                kenar += (f" Onun qarşı hesabı ({', '.join(ortaq[:3])}) artıq xərclərin içindədir — yəni bu məbləğ əsasən eyni "
                          f"əməliyyatların o biri tərəfidir; xərcə əlavə etsəniz, təkrar sayıla bilər. Mühasiblə dəqiqləşdirin.")
            else:
                kenar += f" Xərcə daxil etmək lazımdırsa, yazın: «{k} hesabını {'xərclərə' if teref == 'dt' else 'gəlirlərə'} daxil et»."

    if df.empty:
        return {"metn": (f"«{suzulmus}» {'xərclərin' if teref == 'dt' else 'gəlirlərin'} içində tapılmadı{il_metn}." if filtr else
                         f"«{ad}» vərəqində {sinif}xx hesablarına yazılış tapılmadı{il_metn}.") + kenar +
                (f" Daxil edilməyən hesablar: " + ", ".join(f"{k} «{adlar.get(k, '')}»" for k in cixar) + "." if cixar else ""), "ui": None}
    df["h"], df["q"] = df["h"].map(_kod), df["q"].map(_kod)
    cem, say = float(df["cem"].sum()), int(df["n"].sum())
    dovr = ""
    if tarix:
        r = _sorgu(s, f"SELECT min({sq[tarix]}), max({sq[tarix]}) FROM {v['cedvel']} WHERE {sert}", args).iloc[0]
        dovr = f", {_f(r.iloc[0], 'tarix')[:10]} – {_f(r.iloc[1], 'tarix')[:10]}"
    illik = ""
    if tarix and not il:
        di = _sorgu(s, f"SELECT substr({sq[tarix]}, 1, 4) AS il, sum({M}) AS cem FROM {v['cedvel']} WHERE {sert} GROUP BY il ORDER BY il", args)
        if len(di) > 1:
            illik = " İllər üzrə: " + "; ".join(f"{a} — {_f(b)}" for a, b in di.itertuples(index=False)) + "."
    # xərc növü = qarşı hesab (721-in içində: əmək haqqı, vergilər, ekspertlər...); gəlir növü = gəlir hesabının özü
    # əlavə hesab (bölmədən kənar) özü ayrıca növdür: 603 «təkrarsığortaçıların payı» 213.01-ə qarışmasın
    df["nov"] = [q if (teref == "dt" and h.startswith(sinif)) else h for h, q in zip(df["h"], df["q"])]
    novler = df.groupby("nov", as_index=False).agg(n=("n", "sum"), cem=("cem", "sum")).sort_values("cem", ascending=False)
    hesablar = df.groupby("h", as_index=False)["cem"].sum().sort_values("cem", ascending=False)
    tab = pd.DataFrame([[f"{q}", adlar.get(q) or adlar.get(q.split(".")[0], ""), c, int(n)] for q, n, c in novler.itertuples(index=False)] +
                       [["Cəmi", "", cem, say]], columns=["Qarşı hesab" if teref == "dt" else "Gəlir hesabı", "Xərc növü" if teref == "dt" else "Gəlir növü", "Məbləğ", "Yazılış"])
    top = "; ".join(f"{adlar.get(q) or adlar.get(q.split('.')[0], q)} — {_f(c)}" for q, _, c in novler.head(5).itertuples(index=False))
    hes_metn = ", ".join(f"{h} {adlar.get(h, '')}".strip() + f" — {_f(c)}" for h, c in hesablar.head(6).itertuples(index=False))
    terefi = "Debet" if teref == "dt" else "Kredit"
    elave_metn = ""
    if elave:
        ec = df[df["h"].map(lambda h: not h.startswith(sinif))].groupby("h")["cem"].sum()
        elave_metn = " + mühasibin əlavə etdiyi " + ", ".join(
            f"{h} «{adlar.get(h, '')}» ({_f(c)})" for h, c in ec.items()) if len(ec) else ""
    if filtr:
        bas = f"{'Xərclərin' if teref == 'dt' else 'Gəlirlərin'} içində «{suzulmus}»{il_metn}: **{_f(cem)}** — {say} yazılış{dovr}."
    else:
        bas = (f"Şirkətin ümumi {ifade}{il_metn}: **{_f(cem)}** — «{ad}» vərəqində {terefi} {sinif}xx{elave_metn} "
               f"hesablarına {say} yazılışın cəmi{dovr}.")
    # cəmə daxil edilməyən bölmə hesabları: niyə və nə qədər
    xaric = ""
    if cixar and not filtr:
        hisseler = []
        for k, sebeb in cixar.items():
            x = _sorgu(s, f"SELECT sum({sq[m]}) FROM {v['cedvel']} WHERE CAST({hes} AS TEXT) LIKE ? AND CAST({qarsi} AS TEXT) NOT LIKE '8%'"
                          + (f" AND substr({sq[tarix]}, 1, 4) = ?" if il else ""), (f"{k}%",) + ((il,) if il else ())).iloc[0, 0]
            if x and not pd.isna(x):
                adi = adlar.get(k) if not re.match(r"(счет|schet)\b", norm(adlar.get(k, ""))) else None
                adi = adi or next((a for kk, a in adlar.items() if kk.startswith(k + ".")), "")
                hisseler.append(f"{k} «{adi}» — {_f(x)} ({sebeb})")
        if hisseler:
            xaric = (f" Daxil edilməyib: {'; '.join(hisseler)}. Daxil etmək lazımdırsa, yazın: "
                     f"«{next(iter(cixar))} hesabını {'xərclərə' if teref == 'dt' else 'gəlirə'} daxil et».")
    izah = (" Növ = qarşı hesab (məs. 721-in içində ekspertlər 536.05 ilə keçir)." if teref == "dt" else "")
    alt_setir = f"\nHesablar: {hes_metn}.{izah}" if teref == "dt" else ""
    metn = (bas + kenar + xaric + f" Ən böyük növlər: {top}.{illik}{alt_setir} Bağlanış (8xx) yazılışları sayılmır, "
            f"storno yazılışları daxildir — mühasiblə yoxlayın.")
    return {"metn": metn, "ui": _cedvel(f"{'Xərclər' if teref == 'dt' else 'Gəlirlər'} növlər üzrə · cəmi {_f(cem)}", tab, {"Məbləğ": "rəqəm"}, 60)}


def _hesab_tap(s, n):
    """Mesajda adı keçən hesabı hesablar planından tapır: "təkrarsığortaçıların payı xərcdir" -> 603."""
    if not s.excel:
        return None
    sozler = set(re.findall(r"\w+", n))
    en, xal = None, 0
    for k, a in _hesab_adlari(s).items():
        aw = [w for w in re.findall(r"\w+", norm(a)) if len(w) >= 3]
        if not aw or not k:
            continue
        tutan = sum(any(w == x or (len(w) >= 5 and x.startswith(w[:5])) for x in sozler) for w in aw) / len(aw)
        if tutan > xal or (tutan == xal and en and len(k) < len(en)):
            en, xal = k, tutan
    return en if xal >= 0.75 else None


def hesab_ayarla(s, hesab, bolme="xerc", emeliyyat="daxil"):
    """Mühasibin göstərişi: hesabı ümumi xərcə/gəlirə daxil edir və ya çıxarır (data/hesab_ayarlari.json-da saxlanır)."""
    d = {k: list(v) for k, v in _ayar().items()}
    bolme = "gelir" if bolme.startswith("gel") else "xerc"
    sinif = "7" if bolme == "xerc" else "6"
    hesab = _kod(hesab)
    if not re.fullmatch(KOD_RE, hesab):
        return {"metn": f"«{hesab}» hesab kodu deyil. Məsələn yazın: «607 hesabını gəlirdən çıxar».", "ui": None}
    adlar = _hesab_adlari(s) if s.excel else {}
    ad = adlar.get(hesab) or next((a for k, a in adlar.items() if k.startswith(hesab + ".")), "")
    sil = lambda lst: [h for h in lst if h != hesab]  # noqa: E731
    soz = "xərc" if bolme == "xerc" else "gəlir"
    if emeliyyat == "daxil":
        if hesab.startswith(sinif):
            d[bolme + "_cixar"] = sil(d[bolme + "_cixar"])
            d[bolme + "_saxla"] = sil(d[bolme + "_saxla"]) + [hesab]
        else:
            d[bolme] = sil(d[bolme]) + [hesab]
            diger = "gelir" if bolme == "xerc" else "xerc"
            if hesab.startswith("7" if diger == "xerc" else "6"):  # o biri bölmədədir: oradan çıxar
                d[diger + "_saxla"] = sil(d[diger + "_saxla"])
                d[diger + "_cixar"] = sil(d[diger + "_cixar"]) + [hesab]
            d[diger] = sil(d[diger])
        metn = (f"Oldu: {hesab} «{ad}» bundan sonra ümumi {soz}ə daxil edilir" +
                (f" və {'gəlirdən' if bolme == 'xerc' else 'xərcdən'} çıxarılır" if not hesab.startswith(sinif) else "") +
                ".")
    else:
        if hesab.startswith(sinif):
            d[bolme + "_saxla"] = sil(d[bolme + "_saxla"])
            d[bolme + "_cixar"] = sil(d[bolme + "_cixar"]) + [hesab]
        else:
            d[bolme] = sil(d[bolme])
            diger = "gelir" if bolme == "xerc" else "xerc"
            if hesab in d[diger + "_cixar"]:  # əvvəl bu bölməyə köçürülmüşdü: öz bölməsinə qaytar
                d[diger + "_cixar"] = sil(d[diger + "_cixar"])
                metn_q = f" və yenidən {'gəlirə' if diger == 'gelir' else 'xərcə'} qaytarıldı"
        metn = f"Oldu: {hesab} «{ad}» ümumi {soz}dən çıxarıldı" + (locals().get("metn_q") or "") + "."
    _ayar_yaz(d)
    xeber = ""
    j = _jurnal(s) if (s.excel and emeliyyat == "daxil" and not hesab.startswith(sinif)) else None
    if j:  # hesab əks tərəfdə işləyirsə, işarəni açıq de: məs. 601 kreditdə yığılır -> xərci azaldır
        _, v, dtc, ctc, mc = j
        sq = dict(zip(v["sutunlar"], v["sql"]))
        r = _sorgu(s, f"SELECT sum(CASE WHEN CAST({sq[dtc]} AS TEXT) LIKE ? AND CAST({sq[ctc]} AS TEXT) NOT LIKE '8%' THEN {sq[mc]} ELSE 0 END), "
                      f"sum(CASE WHEN CAST({sq[ctc]} AS TEXT) LIKE ? AND CAST({sq[dtc]} AS TEXT) NOT LIKE '8%' THEN {sq[mc]} ELSE 0 END) "
                      f"FROM {v['cedvel']}", (f"{hesab}%", f"{hesab}%")).iloc[0]
        dov_dt, dov_ct = float(r.iloc[0] or 0), float(r.iloc[1] or 0)
        xalis = dov_dt - dov_ct if bolme == "xerc" else dov_ct - dov_dt
        if xalis < 0:  # hesab 1C-də o biri bölmə kimi işləyir: göstərişə əməl edilir, amma fərq açıq deyilir
            teref = "kreditdə" if bolme == "xerc" else "debetdə"
            xeber = (f" Qeyd: 1C-də bu hesab {teref} yığılır və {'gəlir' if bolme == 'xerc' else 'xərc'} kimi 801-ə bağlanır; "
                     f"göstərişinizlə onun dövriyyəsi ({_f(-xalis)}) {soz}ə əlavə olunur. Ona görə nəticə 1C-nin mənfəət "
                     f"hesabatından fərqlənəcək. Geri qaytarmaq: «{hesab} hesabını {soz}dən çıxar».")
    cix = list(_cixarilanlar(adlar, sinif, bolme)) if adlar else d[bolme + "_cixar"]
    return {"metn": metn + xeber + f" Daxil edilməyən hesablar: {', '.join(cix) or 'yoxdur'}; bölmədən kənar əlavə: "
                           f"{', '.join(d[bolme]) or 'yoxdur'}. Yeni sual verin: «şirkətin ümumi {soz}i».", "ui": None}


def _niyyet(s, deyer):
    try:
        s.son_niyyet = deyer
    except AttributeError:
        pass


def cem_hesabla(s, sorgu, vereq=None):
    """Ümumi cəmi hesablayır: şirkətin ümumi xərci/gəliri (uçot yazılışlarından) və ya mövzu/vərəq üzrə məbləğ cəmi."""
    if not s.excel:
        return _yoxdur(s)
    vv = s.excel["vereqler"]
    il_m = re.search(r"\b(20\d\d)\b", sorgu)
    il = il_m.group(1) if il_m else None
    if vereq:
        ad, v = _v(s, vereq)
        c = _mebleg_sutunu(v, sorgu)
        if not c:
            reqem = [x for x in v["sutunlar"] if v["tipler"].get(x) == "rəqəm" and not TEXNIKI_REQEM.search(norm(x))]
            _niyyet(s, {"tip": "cem", "sorgu": sorgu, "vereq": ad})
            return {"metn": f"«{ad}» vərəqində məbləğ sütunu tapmadım." +
                    (f" Rəqəm sütunları: {', '.join(f'«{x}»' for x in reqem[:10])}. Hansını cəmləyim?" if reqem else
                     " Bu vərəqdə cəmlənəcək rəqəm yoxdur — başqa vərəqin adını yazın."), "ui": None}
        cem, say, suzulub = _cemle(s, v, c, il)
        _niyyet(s, None)
        metn = f"«{ad}» vərəqi üzrə {c} cəmi{f' ({il}-ci il)' if suzulub else ''}: **{_f(cem)}** ({say} sətir)."
        return {"metn": metn, "ui": {"tip": "kartlar", "kartlar": [{"ad": c, "deyer": _f(cem), "alt": f"«{ad}» · {say} sətir"}]}}

    n = norm(sorgu)
    msoz = [w for w in _movzu_sozleri(sorgu) if w not in CEM_SOZLERI and not w.isdigit()]
    for kok, (teref, sinif, ifade) in HESAB_SINIFI.items():
        if any(w.startswith(kok) for w in msoz) or (not msoz and kok in n):
            j = _jurnal(s)
            if j:
                _niyyet(s, None)
                filtr = [w for w in msoz if not w.startswith(("xerc", "gelir")) and w not in FILTR_ATILAN and len(w) >= 4]
                return _jurnal_cemi(s, j, teref, sinif, ifade, il, filtr)

    tapilan = []
    for ad in _movzu_vereqleri(vv, msoz)[:15]:
        c = _mebleg_sutunu(vv[ad], sorgu)
        if c:
            cem, say, _ = _cemle(s, vv[ad], c, il)
            tapilan.append((ad, c, cem, say))
    if not tapilan:
        # vərəq adı uyğun gəlmir: sütun adında mövzu sözü olan vərəq (məs. «Avans» sütunu)
        kokler = [w[:4] for w in msoz if not w.startswith("mebleg")]
        for ad in sorted(vv, key=lambda a: -vv[a]["setir"]):
            v = vv[ad]
            c = next((x for x in v["sutunlar"] if v["tipler"].get(x) == "rəqəm" and not TEXNIKI_REQEM.search(norm(x))
                      and any(k in norm(x) for k in kokler)), None) if kokler else None
            if c:
                cem, say, _ = _cemle(s, v, c, il)
                tapilan.append((ad, c, cem, say))
                if len(tapilan) >= 8:
                    break
    if tapilan:
        _niyyet(s, None)
        ad, c, cem, say = tapilan[0]
        metn = (f"«{ad}» vərəqi üzrə {c} cəmi{f' ({il}-ci il)' if il else ''}: **{_f(cem)}** ({say} sətir)." +
                (f" Bu mövzuda daha {len(tapilan) - 1} vərəq var — cəmləri cədvəldədir (bir-birini təkrarlaya bilər, "
                 f"ona görə toplanmayıb). Başqasını əsas götürmək üçün vərəqin adını yazın." if len(tapilan) > 1 else ""))
        df = pd.DataFrame([[a, cc, x, y] for a, cc, x, y in tapilan], columns=["Vərəq", "Sütun", "Cəm", "Sətir"])
        return {"metn": metn, "ui": _cedvel("Mövzu üzrə cəmlər", df, {"Cəm": "rəqəm"}, 15)}

    # haradan hesablayacağını bilmir: soruşur və cavabı gözləyir (növbəti mesajda vərəq adı gələcək)
    _niyyet(s, {"tip": "cem", "sorgu": sorgu})
    namized = [(a, _mebleg_sutunu(vv[a])) for a in sorted(vv, key=lambda a: -vv[a]["setir"])]
    namized = [(a, c) for a, c in namized if c][:15]
    mov = " ".join(msoz) or "bu sual"
    df = pd.DataFrame([[a, c, vv[a]["setir"]] for a, c in namized], columns=["Vərəq", "Məbləğ sütunu", "Sətir"])
    return {"metn": f"«{mov}» üçün hansı vərəqi cəmləyim? Aşağıdakı vərəqlərdən birinin adını yazın — onun məbləğ cəmini hesablayacağam.",
            "ui": _cedvel("Məbləğ sütunu olan vərəqlər", df, None, 15) if namized else None}


# ======================= sxemlər =======================
_vereq = {"type": "string", "description": "Vərəqin adı (verilməsə, ən böyük vərəq)"}
ALETLER = {
    "excel_xulase": (excel_xulase, "Yüklənmiş Excel faylının strukturu: vərəqlər, sətir sayı, sütunlar. "
                     "Excel haqqında suallarda ilk bunu çağır.", {}),
    "vereq_tap": (vereq_tap, "Ada görə uyğun vərəqləri tapır (çox vərəqli fayllarda, məs. 'bank', 'işçi', 'kontragent').", {
        "sorgu": {"type": "string"}}),
    "sutunlar": (sutunlar, "Vərəqin sütunları üzrə statistika: tip, doluluq, cəm/orta/min/maks, ən çox rast gələn dəyərlər.", {"vereq": _vereq}),
    "setirler": (setirler, "Vərəqdən sətirləri göstərir; istəsən sütun dəyərinə görə süzür (hissəvi uyğunluq).", {
        "vereq": _vereq, "n": {"type": "integer"}, "sutun": {"type": "string"}, "deyer": {"type": "string"}}),
    "qrupla": (qrupla, "Rəqəm sütunu üzrə cəm/orta/say/min/maks; qrup sütunu verilsə, ona görə qruplaşdırır (pivot).", {
        "deyer_sutunu": {"type": "string"}, "qrup_sutunu": {"type": "string"},
        "emeliyyat": {"type": "string", "enum": ["cem", "orta", "say", "min", "maks"]}, "vereq": _vereq, "n": {"type": "integer"}}),
    "en_boyuk": (en_boyuk, "Sütun üzrə ən böyük (və ya artan=true ilə ən kiçik) N sətir.", {
        "sutun": {"type": "string"}, "n": {"type": "integer"}, "artan": {"type": "boolean"}, "vereq": _vereq}),
    "axtar": (axtar, "Bütün vərəqlərdə mətn və ya rəqəm axtarır.", {"sorgu": {"type": "string"}}),
    "obyekt_gostericisi": (obyekt_gostericisi, "Konkret obyekt (kontragent, məhsul, şəxs) üzrə göstəricini tapır: "
                           "məs. 'Kür Avto-nun borcu', 'Aynurun maaşı'. Göstərici verilməsə, obyektin keçdiyi sətirləri göstərir.", {
        "obyekt": {"type": "string"}, "gosterici": {"type": "string", "description": "borc, məbləğ, maaş, ödəniş və s."},
        "vereq": _vereq}),
    "movzu_goster": (movzu_goster, "Mövzuya uyğun vərəqi tapıb sətirlərini göstərir və uyğun digər vərəqləri sadalayır "
                     "(məs. 'xərclər', 'avans hesabatları', 'məzuniyyətlər').", {"sorgu": {"type": "string"}, "n": {"type": "integer"}}),
    "cem_hesabla": (cem_hesabla, "Ümumi cəmi hesablayır: 'şirkətin ümumi xərci/gəliri' (uçot yazılışlarından, Dt 7xx / Ct 6xx), "
                    "mövzu üzrə (məs. 'avansların cəmi') və ya konkret vərəqin məbləğ sütununun cəmi. İl yazılsa, o il üzrə.", {
        "sorgu": {"type": "string"}, "vereq": _vereq}),
    "hesab_ayarla": (hesab_ayarla, "Mühasibin göstərişi ilə hesabı ümumi xərcə/gəlirə daxil edir və ya çıxarır "
                     "(məs. '607 hesabını xərclərə daxil et').", {
        "hesab": {"type": "string"}, "bolme": {"type": "string", "enum": ["xerc", "gelir"]},
        "emeliyyat": {"type": "string", "enum": ["daxil", "cixar"]}}),
    "problemler": (problemler, "Vərəqdə problemləri tapır: təkrar sətirlər, boş xanalar, qeyri-adi və mənfi dəyərlər, şübhəli tarixlər.",
                   {"vereq": _vereq}),
}
MECBURI = {"hesab_ayarla": ["hesab"], "cem_hesabla": ["sorgu"], "movzu_goster": ["sorgu"], "qrupla": ["deyer_sutunu"], "en_boyuk": ["sutun"], "axtar": ["sorgu"], "obyekt_gostericisi": ["obyekt"], "vereq_tap": ["sorgu"]}

MODUL = {
    "kod": "excel",
    "ad": "Excel analizi",
    "tesvir": "İstənilən (o cümlədən 1C ixracı kimi çox böyük) Excel faylını oxuyur: struktur, statistika, qruplaşdırma, axtarış, problemlər",
    "status": "aktiv",
    "teklifler": [
        {"s": "Fayl haqqında ümumi məlumat ver", "b": "Faylın strukturu", "a": "vərəqlər, sütunlar, sətir sayı"},
        {"s": "Şirkətin ümumi xərci nə qədərdir?", "b": "Ümumi xərc", "a": "uçot yazılışlarından hesablanır"},
        {"s": "Fayldakı problemləri tap", "b": "Problemləri tap", "a": "təkrarlar, boş xanalar, qeyri-adi dəyərlər"},
    ],
}


# ======================= lokal (AI-sız) rejim =======================

def _kecen_sutunlar(v, n):
    sozler = re.findall(r"\w+", n)
    out = []
    for c in sorted(v["sutunlar"], key=lambda c: -len(str(c))):
        esas = re.sub(r"\(.*?\)", "", norm(c)).strip()
        acar = [k for k in re.findall(r"\w+", esas) if len(k) >= 3]
        tam = re.search(rf"(?<!\w){re.escape(esas)}", n) if esas else None
        hisse = acar and all(any(w.startswith(k) or (len(w) >= 4 and k.startswith(w)) for w in sozler) for k in acar)
        if (tam or hisse) and c not in out:
            out.append(c)

    def yer(c):
        ks = [k for k in re.findall(r"\w+", re.sub(r"\(.*?\)", "", norm(c))) if len(k) >= 3]
        pos = [m.start() for k in ks for m in re.finditer(re.escape(k[:4]), n)]
        return min(pos) if pos else 999
    return sorted(out, key=yer)


SUAL_SOZLERI = {"ne", "qeder", "necedir", "nece", "hansi", "kim", "harada", "nedir", "var", "yox", "olan", "ucun", "ile",
                "uzre", "gore", "goster", "ver", "de", "da", "bu", "o", "mene", "bize", "cemi", "say", "sayi", "borcu", "borc"}


def _obyekt_tap(s, n):
    """Mesajda adı keçən obyekti adlar lüğətindən (_lugat) tapır; şəkilçili yazılışlar da tutulur."""
    sozler = re.findall(r"\w+", n)
    namizedler = set()
    for i in range(len(sozler)):
        for j in range(i + 1, min(i + 6, len(sozler) + 1)):
            parca = sozler[i:j]
            if j - i == 1 and (len(parca[0]) < 4 or parca[0] in SUAL_SOZLERI):
                continue
            son = parca[-1]
            for k in range(0, 5):  # şəkilçini at: tikintinin -> tikinti, aynurun -> aynur
                if len(son) - k < 3:
                    break
                namizedler.add(" ".join(parca[:-1] + [son[:len(son) - k]]))
    if not namizedler:
        return None
    with _con(s) as con:
        q = ",".join("?" * len(namizedler))
        rows = con.execute(f"SELECT acar, deyer FROM _lugat WHERE acar IN ({q})", list(namizedler)).fetchall()
    # vərəq adları (və onların başlanğıcı) obyekt deyil: "bank" -> «Banklar» vərəqi, "kontragentlər" -> «Kontragentlər»
    vereq_adlari = [norm(a) for a in s.excel["vereqler"]]
    rows = [r for r in rows if not any(va == r[0] or va.startswith(r[0]) or r[0] in va for va in vereq_adlari)]
    if not rows:
        return None
    acar, deyer = max(rows, key=lambda r: (len(r[0]), -len(r[1])))
    return acar, deyer


def qayda(s, mesaj):
    """AI-sız rejimdə Excel suallarını alətlərə çevirir. (alət, args) və ya None qaytarır."""
    if not s.excel:
        return None
    n = norm(mesaj)
    sozler = re.findall(r"\w+", n)
    vv = s.excel["vereqler"]
    say = re.search(r"\b(\d{1,3})\b", n)
    sitat = re.search(r"[«\"']([^»\"']+)[»\"']", mesaj)

    # 1) vərəq seçimi: adı keçən sütun sayı + vərəq adı uyğunluğu
    def xal(a):
        return (_vereq_xali(a, sozler) >= 0.99, len(_kecen_sutunlar(vv[a], n)), _vereq_xali(a, sozler), vv[a]["setir"])
    secim = max(vv, key=xal)
    v = vv[secim]
    adi_kecen = _kecen_sutunlar(v, n)
    vq = {"vereq": secim} if len(vv) > 1 else {}
    reqem = [c for c in adi_kecen if v["tipler"][c] == "rəqəm"]
    diger = [c for c in adi_kecen if v["tipler"][c] != "rəqəm"]
    vereq_adi = _vereq_xali(secim, sozler) >= 0.99

    # əvvəlki sualın davamı: "hansı vərəqi cəmləyim?" -> istifadəçi vərəqin adını yazdı
    gozle = getattr(s, "son_niyyet", None)
    if gozle:
        _niyyet(s, None)
        if gozle.get("tip") == "cem" and gozle.get("vereq") in vv and len(sozler) <= 6:
            gv = vv[gozle["vereq"]]
            if any(gv["tipler"].get(c) == "rəqəm" for c in _kecen_sutunlar(gv, n)):
                return "cem_hesabla", {"sorgu": mesaj, "vereq": gozle["vereq"]}
        if gozle.get("tip") == "cem" and vereq_adi and len(sozler) <= 6:
            uygun = sorted((a for a in vv if _vereq_xali(a, sozler) >= 0.99), key=lambda a: -vv[a]["setir"])
            ad = next((a for a in uygun if _mebleg_sutunu(vv[a], mesaj)), secim)
            return "cem_hesabla", {"sorgu": gozle["sorgu"] + " " + mesaj, "vereq": ad}

    # mühasibin göstərişi: "607 hesabını xərclərə daxil et", "607 gəlir deyil", "Təkrarsığortaçıların payı xərcdir"
    emr_daxil = re.search(r"\b(xerc|gelir)(dir|dur)\b|\b(xerc|gelir)\w* (daxil et|elave et|qos)", n)
    emr_cixar = re.search(r"\b(xerc|gelir)\w* deyil|\b(xerc|gelir)\w* (cixar|cixart|sil\b)", n)
    if emr_daxil or emr_cixar:
        hm = re.search(r"\b(\d{2,3}(?:\.\d{1,2})?)\b", n)
        kod = hm.group(1) if hm else _hesab_tap(s, n)
        if kod:
            mm = emr_cixar or emr_daxil
            bolme = next(g for g in mm.groups() if g in ("xerc", "gelir"))
            return "hesab_ayarla", {"hesab": kod, "bolme": bolme, "emeliyyat": "cixar" if emr_cixar else "daxil"}
    if re.search(r"problem|sehv|anomal|dublikat|\btekrar\b|tekrar setir|bos xana|\byoxla(?!nmali)", n):
        return "problemler", (vq if (vereq_adi or len(vv) == 1) else {})
    if re.search(r"vereq", n) and re.search(r"hansi|siyahi|tap|var", n) and not vereq_adi:
        return ("vereq_tap", {"sorgu": sitat.group(1)}) if sitat else ("excel_xulase", {})
    if re.search(r"sutun|struktur", n) and not reqem:
        return ("sutunlar", vq) if "sutun" in n else ("excel_xulase", {})
    if re.search(r"axtar|tap\b|harada", n) and sitat:
        return "axtar", {"sorgu": sitat.group(1)}

    # uçot yazılışları olan faylda xərc/gəlir sualları: "ekspert xərcləri", "xərclərin içində ekspertlər"
    if re.search(r"\bxerc|\bgelir", n) and not re.search(r"uzre|gore|her bir|qrupla|setir|sened", n) and _jurnal(s):
        return "cem_hesabla", {"sorgu": mesaj}
    # ümumi cəm: "şirkətin ümumi xərci nə qədərdir", "2025-ci ildə cəmi gəlir", "avansların cəmi"
    cem_movzu = r"xerc|gelir|mebleg|satis|odenis|maas|emek haqq|vergi|avans|medaxil|mexaric|tutulma|hesablama"
    if (re.search(r"\b(umumi|cemi|toplam|umumilikde|cemini)\b", n) and re.search(cem_movzu, n)
            and not re.search(r"uzre|gore|her bir|qrupla", n) and (len(vv) > 1 or not reqem)):
        return "cem_hesabla", {"sorgu": mesaj, **({"vereq": secim} if vereq_adi else {})}

    if re.search(r"fayl haqqinda|umumi melumat|faylin strukturu|ne var bu faylda", n):
        return "excel_xulase", {}
    if vereq_adi and not reqem and re.search(r"nece|siyahi|goster|bax|hamisi|list|ne var|\bsay|nefer", n):
        return "setirler", {"n": int(say.group(1)) if say else 20, **vq}

    qrup_niyyeti = diger and re.search(r"uzre|gore|her bir", n) and re.search(r"\b(cem|toplam|orta|say)", n)
    obj = None if qrup_niyyeti or re.search(r"en boyuk|en kicik|en cox|en az", n) else _obyekt_tap(s, n)
    if obj:
        k, val = obj
        qalan = n.replace(k, " ")
        gosterici = " ".join(w for w in re.findall(r"\w+", qalan) if len(w) >= 3 and w not in SUAL_SOZLERI - {"borc", "borcu"}) or None
        args = {"obyekt": val, "gosterici": gosterici}
        if vereq_adi:
            args["vereq"] = secim
        return "obyekt_gostericisi", args

    if re.search(r"ne qeder|necedir|nece manat", n) and re.search(r"xerc|gelir", n) and not reqem:
        return "cem_hesabla", {"sorgu": mesaj}
    if re.search(r"en boyuk|en cox|en yuksek|\btop\b|en kicik|en az|en asagi", n) and (reqem or adi_kecen):
        return "en_boyuk", {"sutun": (reqem or adi_kecen)[0], "n": int(say.group(1)) if say else 10,
                            "artan": bool(re.search(r"en kicik|en az|en asagi", n)), **vq}
    if re.search(r"\b(cem|toplam|orta|say|nece eded|min\b|maks)", n) and (reqem or re.search(r"\bsay|nece eded", n)):
        em = ("orta" if re.search(r"\borta", n) else "say" if re.search(r"\bsay|nece eded", n)
              else "min" if re.search(r"\bmin\b", n) else "maks" if "maks" in n else "cem")
        qrup = diger[0] if diger and re.search(r"uzre|gore|her bir|qrupla", n) else None
        deyer = (reqem or ([qrup] if qrup else adi_kecen) or [v["sutunlar"][0]])[0]
        return "qrupla", {"deyer_sutunu": deyer, "qrup_sutunu": qrup, "emeliyyat": em, **vq}
    if vereq_adi and re.search(r"nece|siyahi|goster|bax|hamisi|list|ne var|\bsay", n):
        return "setirler", {"n": int(say.group(1)) if say else 20, **vq}
    if re.search(r"xulase|umumi|fayl haqqinda|veziyyet", n):
        return "excel_xulase", {}
    if diger and sitat:
        return "setirler", {"sutun": diger[0], "deyer": sitat.group(1), **vq}
    if sitat:
        return "axtar", {"sorgu": sitat.group(1)}
    if vereq_adi or len(vv) == 1:
        if re.search(r"ilk|setir|goster|bax|siyahi", n) or vereq_adi:
            return "setirler", {"n": int(say.group(1)) if say else 20, **vq}
    # çox vərəqli faylda: mövzuya görə vərəq seç (ixtiyari vərəqə düşmə)
    if len(vv) > 1 and _movzu_sozleri(mesaj) and not re.search(r"borc|istisna|baglanti|qaime", n):
        return "movzu_goster", {"sorgu": mesaj, "n": int(say.group(1)) if say else 20}
    return None
