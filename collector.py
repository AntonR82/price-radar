#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Price Radar Collector — узел Салала
Собирает: розницу с маркетплейсов (BigHaat/Jumia/Jiji/Daraz), курсы валют,
бенчмарки World Bank Pink Sheet. Пишет prices.json + history/.
Запуск: python collector.py   (планировщик — GitHub Actions, см. .github/workflows/daily.yml)
"""
import json, re, sys, os, datetime, statistics, time
import requests
from bs4 import BeautifulSoup

try:
    from curl_cffi import requests as cffi  # эмуляция TLS-отпечатка Chrome — против анти-бот 403
except ImportError:
    cffi = None

UA = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
TODAY = datetime.date.today().isoformat()

# ---------- утилиты ----------
def get(url, extra=None, tries=3):
    """GET с эмуляцией Chrome (curl_cffi) и ретраями на 403/429/5xx; None при неудаче."""
    last = None
    for i in range(tries):
        try:
            h = {**UA, **(extra or {})}
            if cffi:
                r = cffi.get(url, headers=h, impersonate="chrome", timeout=25)
            else:
                r = requests.get(url, headers=h, timeout=25)
            if r.status_code == 200:
                return r
            last = f"HTTP {r.status_code}"
            if r.status_code not in (403, 408, 429, 500, 502, 503, 504):
                break  # 404 и прочее ретраить бессмысленно
        except Exception as e:
            last = e
        time.sleep(3 * (i + 1))
    print(f"  ! {url[:80]} -> {last}", file=sys.stderr, flush=True)
    return None

def get_json(url, referer=None):
    """GET, ожидающий JSON; если пришёл HTML (бот-заглушка) — печатает первые байты."""
    extra = {"Accept": "application/json, text/plain, */*"}
    if referer:
        extra["Referer"] = referer
    r = get(url, extra=extra)
    if not r:
        return None
    try:
        return r.json()
    except Exception:
        head = re.sub(r"\s+", " ", (r.text or "")[:70])
        print(f"  ! не-JSON от {url[:60]} -> «{head}»", file=sys.stderr, flush=True)
        return None

PACK_RE = re.compile(r'(\d+(?:[.,]\d+)?)\s*(kg|кг|kgs|kilo)', re.I)
def price_per_ton(price_local, title):
    """цена лота -> цена за тонну (по упаковке из названия; дефолт 1 кг)"""
    m = PACK_RE.search(title or "")
    pack = float(m.group(1).replace(",", ".")) if m else 1.0
    if pack <= 0: pack = 1.0
    return price_local / pack * 1000.0

def summarize(vals):
    vals = [v for v in vals if v and v > 0]
    if not vals: return None
    return {"median": round(statistics.median(vals)), "min": round(min(vals)),
            "max": round(max(vals)), "n": len(vals)}

def ldjson_products(html):
    """[(title, price)] из <script type="application/ld+json"> — общий запасной парсер."""
    out = []
    for m in re.finditer(r'<script[^>]*application/ld\+json[^>]*>(.*?)</script>', html or "", re.S | re.I):
        try:
            data = json.loads(m.group(1).strip())
        except Exception:
            continue
        stack = [data]
        while stack:
            node = stack.pop()
            if isinstance(node, dict):
                if str(node.get("@type", "")).lower() == "product":
                    offers = node.get("offers") or {}
                    if isinstance(offers, list):
                        offers = offers[0] if offers else {}
                    price = offers.get("price") or offers.get("lowPrice")
                    if price:
                        try:
                            out.append((node.get("name", ""), float(str(price).replace(",", ""))))
                        except Exception:
                            pass
                stack.extend(node.values())
            elif isinstance(node, list):
                stack.extend(node)
    return out

# ---------- курсы ----------
def fx_rates():
    out = {}
    data = get_json("https://open.er-api.com/v6/latest/USD")
    if data:
        rates = data.get("rates", {})
        for c in ["INR","PKR","KES","TZS","BDT","LKR","UGX","MZN","OMR","AED","CNY"]:
            if c in rates: out[c] = rates[c]
    return out

# ---------- бенчмарк: World Bank Pink Sheet (месячный) ----------
PINK_URLS = [
 "https://thedocs.worldbank.org/en/doc/5d903e848db1d1b83e0ec8f744e55570-0350012021/related/CMO-Historical-Data-Monthly.xlsx",
]
def pinksheet():
    """тянем xlsx, берём последнюю строку по Urea/DAP/TSP/MOP; при провале — None"""
    try:
        import openpyxl, io
        for u in PINK_URLS:
            r = get(u)
            if not r: continue
            wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
            ws = wb["Monthly Prices"]
            rows = list(ws.iter_rows(values_only=True))
            header_i = next(i for i,row in enumerate(rows) if row and "Urea" in str(row))
            hdr = [str(x) for x in rows[header_i]]
            last = next(r for r in reversed(rows) if r and r[0] and str(r[0])[:2].isdigit())
            def col(sub):
                for j,h in enumerate(hdr):
                    if sub.lower() in h.lower(): return last[j]
                return None
            return {"date": str(last[0]), "urea_fob": col("Urea"), "dap_fob": col("DAP"),
                    "tsp_fob": col("TSP"), "mop_fob": col("Potassium")}
    except Exception as e:
        print(f"  ! pinksheet -> {e}", file=sys.stderr, flush=True)
    return None

# ---------- адаптеры маркетплейсов ----------
def bighaat(query):
    """Индия. Ступень 1: Shopify-автоподсказка; ступень 2: JSON-LD со страницы поиска."""
    vals = []
    q = requests.utils.quote(query)
    data = get_json(f"https://www.bighaat.com/search/suggest.json?q={q}&resources[type]=product&resources[limit]=10",
                    referer="https://www.bighaat.com/")
    if data:
        try:
            for p in data["resources"]["results"].get("products", []):
                vals.append(price_per_ton(float(p.get("price", 0) or 0), p.get("title", "")))
        except Exception as e:
            print(f"  ! bighaat parse -> {e}", file=sys.stderr, flush=True)
    if not vals:
        r = get(f"https://www.bighaat.com/search?q={q}", extra={"Referer": "https://www.bighaat.com/"})
        if r:
            for title, price in ldjson_products(r.text)[:12]:
                vals.append(price_per_ton(price, title))
    return vals  # INR/t

def jumia(query, domain="jumia.co.ke"):
    """Кения. Карточки листинга; фолбэк — JSON-LD."""
    vals = []
    r = get(f"https://www.{domain}/catalog/?q={requests.utils.quote(query)}",
            extra={"Referer": f"https://www.{domain}/"})
    if r:
        soup = BeautifulSoup(r.text, "html.parser")
        for art in soup.select("article.prd")[:12]:
            name = art.select_one(".name"); prc = art.select_one(".prc")
            if not (name and prc): continue
            digits = re.sub(r"[^\d.]", "", prc.get_text().split("-")[0])
            if digits:
                vals.append(price_per_ton(float(digits), name.get_text()))
        if not vals:
            for title, price in ldjson_products(r.text)[:12]:
                vals.append(price_per_ton(price, title))
    return vals  # KES/t

def jiji(query, host="jiji.co.ke"):
    """Jiji (Кения/Танзания/Уганда...) — полупубличный JSON-эндпоинт листинга."""
    vals = []
    q = requests.utils.quote(query)
    data = get_json(f"https://{host}/api_web/v1/listing?query={q}&page=1",
                    referer=f"https://{host}/search?query={q}")
    if data:
        try:
            for ad in data.get("adverts_list", {}).get("adverts", [])[:15]:
                price = ad.get("price_obj", {}).get("value") or ad.get("price")
                title = ad.get("title","")
                if price: vals.append(price_per_ton(float(price), title))
        except Exception as e:
            print(f"  ! jiji parse -> {e}", file=sys.stderr, flush=True)
    return vals

def daraz(query, host="www.daraz.pk"):
    """Пакистан/Бангладеш/Шри-Ланка. Цены в window.pageData (JSON внутри страницы)."""
    vals = []
    r = get(f"https://{host}/catalog/?q={requests.utils.quote(query)}",
            extra={"Referer": f"https://{host}/"})
    if not r:
        return vals
    m = re.search(r'window\.pageData\s*=\s*(\{.*?\})\s*</script>', r.text, re.S)
    if m:
        try:
            for it in (json.loads(m.group(1)).get("mods", {}).get("listItems", []) or [])[:15]:
                price = it.get("price") or re.sub(r"[^\d.]", "", str(it.get("priceShow", "")))
                if price:
                    vals.append(price_per_ton(float(str(price).replace(",", "")), it.get("name", "")))
        except Exception as e:
            print(f"  ! daraz parse -> {e}", file=sys.stderr, flush=True)
    if not vals:  # старый запасной regex по сырому HTML
        for mm in re.finditer(r'"price":"?([\d.]+)"?.{0,400}?"name":"(.*?)"', r.text):
            try: vals.append(price_per_ton(float(mm.group(1)), mm.group(2)))
            except Exception: pass
    return vals

# ---------- конфиг целей ----------
TARGETS = {
  "Индия":     {"ccy":"INR","adapter":lambda q: bighaat(q),
    "grades":{"WSF 19-19-19 (Cl)":"19:19:19 water soluble","WSF 18-18-18+5S":"18:18:18 water soluble","WSF 13-0-45":"13:0:45","MKP 0-52-34":"0:52:34 MKP","SOP 0-0-50":"sulphate of potash soluble"}},
  "Пакистан":  {"ccy":"PKR","adapter":lambda q: daraz(q,"www.daraz.pk"),
    "grades":{"WSF 19-19-19 (Cl)":"npk 19 19 19 water soluble","WSF 18-18-18+5S":"npk 18 18 18","DAP 18-46":"DAP fertilizer bag","SOP 0-0-50":"sulphate of potash","WSF 13-0-45":"potassium nitrate fertilizer"}},
  "Бангладеш": {"ccy":"BDT","adapter":lambda q: daraz(q,"www.daraz.com.bd"),
    "grades":{"WSF 19-19-19 (Cl)":"npk 19 19 19","SOP 0-0-50":"potassium sulphate fertilizer"}},
  "Кения":     {"ccy":"KES","adapter":lambda q: (jumia(q,"jumia.co.ke") + jiji(q,"jiji.co.ke")),
    "grades":{"WSF 19-19-19 (Cl)":"npk 19 19 19 water soluble","WSF 18-18-18+5S":"npk 18 18 18","DAP 18-46":"DAP fertilizer","WSF 13-0-45":"potassium nitrate","SOP 0-0-50":"sulphate of potash"}},
  "Танзания":  {"ccy":"TZS","adapter":lambda q: jiji(q,"jiji.co.tz"),
    "grades":{"WSF 19-19-19 (Cl)":"npk 19 19 19","DAP 18-46":"DAP fertilizer","Карбамид":"urea fertilizer"}},
  "Уганда":    {"ccy":"UGX","adapter":lambda q: jiji(q,"jiji.ug"),
    "grades":{"WSF 19-19-19 (Cl)":"npk 19 19 19"}},
  "Шри-Ланка": {"ccy":"LKR","adapter":lambda q: daraz(q,"www.daraz.lk"),
    "grades":{"WSF 19-19-19 (Cl)":"npk 19 19 19 water soluble","SOP 0-0-50":"potassium sulphate"}},
}

def main():
    if not cffi:
        print("  ! curl_cffi не установлен — работаю без эмуляции Chrome (выше риск 403)",
              file=sys.stderr, flush=True)
    fx = fx_rates()
    out = {"updated": TODAY, "fx": fx, "benchmark": pinksheet(), "markets": {}}
    total = 0
    for market, cfg in TARGETS.items():
        print(f"== {market}", flush=True)
        mres = {}
        rate = fx.get(cfg["ccy"])
        for grade, q in cfg["grades"].items():
            vals_local = cfg["adapter"](q)
            s = summarize(vals_local)
            if s:
                s["ccy"] = cfg["ccy"]
                if rate: s["usd_t_median"] = round(s["median"] / rate)
                mres[grade] = s
                total += s["n"]
                print(f"   {grade}: n={s['n']} median {s['median']} {cfg['ccy']}/t" + (f" ≈ ${s.get('usd_t_median')}/t" if rate else ""), flush=True)
            time.sleep(1.5)
        if not mres:
            print("   (пусто — источник не отдал ни одной цены)", flush=True)
        out["markets"][market] = mres
    os.makedirs("history", exist_ok=True)
    json.dump(out, open("prices.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump(out, open(f"history/{TODAY}.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"OK -> prices.json (собрано позиций: {total})", flush=True)

if __name__ == "__main__":
    main()
