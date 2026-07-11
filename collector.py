#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Price Radar Collector — узел Салала
Собирает: розницу с маркетплейсов (BigHaat/Jumia/Jiji/Daraz), курсы валют,
бенчмарки World Bank Pink Sheet. Пишет prices.json + history/.
Запуск: python collector.py   (планировщик — GitHub Actions, см. .github/workflows/daily.yml)
"""
import json, re, sys, datetime, statistics, time
import requests
from bs4 import BeautifulSoup

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36",
      "Accept-Language": "en"}
TODAY = datetime.date.today().isoformat()

# ---------- утилиты ----------
def get(url, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=25, **kw)
        r.raise_for_status()
        return r
    except Exception as e:
        print(f"  ! {url[:80]} -> {e}", file=sys.stderr)
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

# ---------- курсы ----------
def fx_rates():
    out = {}
    r = get("https://open.er-api.com/v6/latest/USD")
    if r:
        rates = r.json().get("rates", {})
        for c in ["INR","PKR","KES","TZS","BDT","LKR","MZN","OMR","AED","CNY"]:
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
        print(f"  ! pinksheet -> {e}", file=sys.stderr)
    return None

# ---------- адаптеры маркетплейсов ----------
def bighaat(query):
    """Индия. BigHaat = Shopify -> публичный /products.json (самый надёжный адаптер)"""
    vals = []
    r = get(f"https://www.bighaat.com/search/suggest.json?q={requests.utils.quote(query)}&resources[type]=product&resources[limit]=10")
    if r:
        try:
            for p in r.json()["resources"]["results"].get("products", []):
                price = float(p.get("price", 0))
                vals.append(price_per_ton(price, p.get("title","")))
        except Exception as e:
            print(f"  ! bighaat parse -> {e}", file=sys.stderr)
    return vals  # INR/t

def jumia(query, domain="jumia.co.ke"):
    """Кения (и .co.tz нет — Танзания через jiji). Парсим карточки листинга."""
    vals = []
    r = get(f"https://www.{domain}/catalog/?q={requests.utils.quote(query)}")
    if r:
        soup = BeautifulSoup(r.text, "html.parser")
        for art in soup.select("article.prd")[:12]:
            name = art.select_one(".name"); prc = art.select_one(".prc")
            if not (name and prc): continue
            digits = re.sub(r"[^\d.]", "", prc.get_text().split("-")[0])
            if digits:
                vals.append(price_per_ton(float(digits), name.get_text()))
    return vals  # KES/t

def jiji(query, host="jiji.co.ke"):
    """Jiji (Кения/Танзания/Уганда...) — полупубличный JSON-эндпоинт листинга."""
    vals = []
    r = get(f"https://{host}/api_web/v1/listing?query={requests.utils.quote(query)}&page=1")
    if r:
        try:
            for ad in r.json().get("adverts_list", {}).get("adverts", [])[:15]:
                price = ad.get("price_obj", {}).get("value") or ad.get("price")
                title = ad.get("title","")
                if price: vals.append(price_per_ton(float(price), title))
        except Exception as e:
            print(f"  ! jiji parse -> {e}", file=sys.stderr)
    return vals

def daraz(query, host="www.daraz.pk"):
    """Пакистан/Бангладеш(daraz.com.bd). Цены сидят в JSON внутри страницы."""
    vals = []
    r = get(f"https://{host}/catalog/?q={requests.utils.quote(query)}")
    if r:
        for m in re.finditer(r'"price":"?([\d.]+)"?.{0,400}?"name":"(.*?)"', r.text):
            try: vals.append(price_per_ton(float(m.group(1)), m.group(2)))
            except: pass
        if not vals:
            for m in re.finditer(r'"priceShow":"Rs\.\s*([\d,]+)".{0,400}?"name":"(.*?)"', r.text):
                vals.append(price_per_ton(float(m.group(1).replace(",","")), m.group(2)))
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
    fx = fx_rates()
    out = {"updated": TODAY, "fx": fx, "benchmark": pinksheet(), "markets": {}}
    for market, cfg in TARGETS.items():
        print(f"== {market}")
        mres = {}
        rate = fx.get(cfg["ccy"])
        for grade, q in cfg["grades"].items():
            vals_local = cfg["adapter"](q)
            s = summarize(vals_local)
            if s:
                s["ccy"] = cfg["ccy"]
                if rate: s["usd_t_median"] = round(s["median"] / rate)
                mres[grade] = s
                print(f"   {grade}: n={s['n']} median {s['median']} {cfg['ccy']}/t" + (f" ≈ ${s.get('usd_t_median')}/t" if rate else ""))
            time.sleep(1.5)
        out["markets"][market] = mres
    json.dump(out, open("prices.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump(out, open(f"history/{TODAY}.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
    print("OK -> prices.json")

if __name__ == "__main__":
    main()
