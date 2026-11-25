#!/usr/bin/env python3
import os, re, json, html, hashlib, time
from datetime import datetime, timedelta
from urllib.parse import quote
from zoneinfo import ZoneInfo
import requests, feedparser
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

CAIRO = ZoneInfo("Africa/Cairo")
MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
DATA_FILE, XLSX_FILE = "data.json", "banks_report.xlsx"
MAX_ITEMS = 3000
NBE = "البنك الأهلي المصري"

BANKS = {
    NBE: ['"البنك الأهلي المصري"', '"National Bank of Egypt"'],
    "بنك مصر": ['"بنك مصر"', '"Banque Misr"'],
    "CIB": ['"البنك التجاري الدولي"', '"CIB Egypt"'],
    "QNB": ['"QNB مصر"', '"QNB Egypt"'],
    "بنك القاهرة": ['"بنك القاهرة"'],
    "بنك الإسكندرية": ['"بنك الإسكندرية"', '"Alexbank"'],
    "ADIB": ['"مصرف أبوظبي الإسلامي مصر"', '"ADIB Egypt"'],
    "AAIB": ['"البنك العربي الأفريقي الدولي"', '"AAIB"'],
    "HSBC": ['"HSBC مصر"', '"HSBC Egypt"'],
    "التعمير والإسكان": ['"بنك التعمير والإسكان"', '"HDB Egypt"'],
}
UA = {"User-Agent": "Mozilla/5.0 (compatible; EgyBankMonitor/1.0)"}

def now():
    return datetime.now(CAIRO)

def clean(t, n=400):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", t or ""))).strip()[:n]

def mk(bank, platform, title, text, url, ts, engagement=0):
    return {
        "id": hashlib.md5(url.encode()).hexdigest()[:16],
        "bank": bank, "platform": platform, "title": clean(title, 200),
        "snippet": clean(text), "url": url, "ts": ts.astimezone(CAIRO).isoformat(),
        "engagement": int(engagement),
    }

def fetch_news():
    out = []
    host = "".join(["https://", "news.", "google.", "com/rss/search?q="])
    for bank, queries in BANKS.items():
        for q in queries:
            url = host + quote(q + " when:2d") + "&hl=ar&gl=EG&ceid=EG:ar"
            try:
                for e in feedparser.parse(url).entries[:15]:
                    ts = datetime(*e.published_parsed[:6], tzinfo=ZoneInfo("UTC")) if getattr(e, "published_parsed", None) else now()
                    src = clean(getattr(e, "source", {}).get("title", "")) if hasattr(e, "source") else ""
                    out.append(mk(bank, "أخبار" + (f" - {src}" if src else ""), e.title, e.get("summary", ""), e.link, ts))
            except Exception as ex:
                print("news error", bank, ex)
    return out
def fetch_reddit():
    out = []
    host = "".join(["https://", "www.", "reddit.", "com"])
    for bank, queries in BANKS.items():
        q = queries[0].strip('"')
        try:
            r = requests.get(host + "/search.json", params={"q": q, "sort": "new", "t": "week", "limit": 10}, headers=UA, timeout=20)
            if r.status_code != 200:
                continue
            for c in r.json()["data"]["children"]:
                d = c["data"]
                out.append(mk(bank, "Reddit", d["title"], d.get("selftext", ""), host + d["permalink"],
                              datetime.fromtimestamp(d["created_utc"], ZoneInfo("UTC")), d.get("score", 0) + d.get("num_comments", 0)))
            time.sleep(1)
        except Exception as ex:
            print("reddit error", bank, ex)
    return out

def fetch_youtube():
    key = os.getenv("YOUTUBE_API_KEY")
    if not key:
        return []
    out = []
    after = (datetime.utcnow() - timedelta(days=3)).strftime("%Y-%m-%dT%H:%M:%SZ")
    host = "".join(["https://", "www.", "googleapis.", "com/youtube/v3"])
    watch = "".join(["https://", "www.", "youtube.", "com/watch?v="])
    for bank, queries in BANKS.items():
        try:
            r = requests.get(host + "/search", params={"part": "snippet", "q": queries[0].strip('"'), "type": "video", "order": "date",
                "publishedAfter": after, "maxResults": 8, "regionCode": "EG", "key": key}, timeout=20).json()
            ids = [i["id"]["videoId"] for i in r.get("items", [])]
            if not ids:
                continue
            stats = requests.get(host + "/videos", params={"part": "statistics", "id": ",".join(ids), "key": key}, timeout=20).json()
            views = {v["id"]: int(v["statistics"].get("viewCount", 0)) for v in stats.get("items", [])}
            for i in r["items"]:
                s, vid = i["snippet"], i["id"]["videoId"]
                ts = datetime.fromisoformat(s["publishedAt"].replace("Z", "+00:00"))
                out.append(mk(bank, "YouTube", s["title"], s.get("description", ""), watch + vid, ts, views.get(vid, 0)))
        except Exception as ex:
            print("youtube error", bank, ex)
    return out

def ask_groq(prompt, system, max_tokens=4000):
    api_url = "".join(["https://", "api.", "groq.", "com/openai/v1/chat/completions"])
    r = requests.post(api_url, timeout=120, headers={
        "Authorization": f"Bearer {os.environ['GROQ_API_KEY']}",
        "Content-Type": "application/json"},
        json={"model": MODEL, "max_tokens": max_tokens, "temperature": 0.2,
              "messages": [{"role": "system", "content": system}, {"role": "user", "content": prompt}]})
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]
def analyze(items):
    system = "أنت محلل رصد إعلامي للقطاع المصرفي المصري. أجب بمصفوفة JSON صالحة فقط بدون أي نص آخر أو علامات Markdown."
    kept = []
    for i in range(0, len(items), 15):
        batch = items[i:i + 15]
        payload = [{"id": b["id"], "bank": b["bank"], "platform": b["platform"], "title": b["title"], "text": b["snippet"]} for b in batch]
        prompt = (
            "لكل عنصر أرجع كائناً بالحقول: id, relevant (true إذا كان المحتوى يخص البنك المذكور فعلاً), "
            "summary (ملخص عربي في جملة أو جملتين), type (خبر | إعلان/حملة | منتج/خدمة | مالي/نتائج | شكوى/خدمة عملاء | تنظيمي | أخرى), "
            "sentiment (إيجابي | سلبي | محايد).\nأرجع مصفوفة JSON فقط.\n" + json.dumps(payload, ensure_ascii=False))
        try:
            txt = re.sub(r"^```(?:json)?|```$", "", ask_groq(prompt, system).strip(), flags=re.M).strip()
            res = {x["id"]: x for x in json.loads(txt)}
            time.sleep(2)
        except Exception as ex:
            print("analysis error", ex)
            continue
        for b in batch:
            b["_done"] = True
            x = res.get(b["id"])
            if x and x.get("relevant"):
                b.update(summary=x.get("summary", b["title"]), type=x.get("type", "أخرى"), sentiment=x.get("sentiment", "محايد"))
                kept.append(b)
    return kept

def benchmark(items, prev):
    last = prev and prev.get("generated")
    if last and now() - datetime.fromisoformat(last) < timedelta(hours=71):
        return prev
    since = now() - timedelta(days=3)
    recent = [i for i in items if datetime.fromisoformat(i["ts"]) >= since]
    if not recent:
        return prev
    stats = {b: {"count": len(rows := [i for i in recent if i["bank"] == b]),
                 "positive": sum(r["sentiment"] == "إيجابي" for r in rows),
                 "negative": sum(r["sentiment"] == "سلبي" for r in rows),
                 "neutral": sum(r["sentiment"] == "محايد" for r in rows)} for b in BANKS}
    nbe_top = [{"summary": r["summary"], "sentiment": r["sentiment"]} for r in recent if r["bank"] == NBE][:15]
    prompt = (f"إحصاءات آخر 3 أيام:\n{json.dumps(stats, ensure_ascii=False)}\nأبرز أخبار {NBE}:\n{json.dumps(nbe_top, ensure_ascii=False)}\n\n"
              f"اكتب تقرير مقارنة عربياً مختصراً (حتى 10 أسطر) يقارن {NBE} بالمنافسين وتوصيتين عمليتين.")
    try:
        return {"generated": now().isoformat(), "text": ask_groq(prompt, "أنت محلل استراتيجي للقطاع المصرفي المصري.").strip()}
    except Exception as ex:
        print("benchmark error", ex)
        return prev

def export_excel(items):
    wb = Workbook()
    ws = wb.active
    ws.title = "رصد البنوك"
    ws.sheet_view.rightToLeft = True
    headers = ["التاريخ", "الوقت", "البنك", "المنصة", "النوع", "الملخص", "التقييم", "التفاعل", "الرابط"]
    ws.append(headers)
    fill, font = PatternFill("solid", fgColor="0F172A"), Font(bold=True, color="FFFFFF")
    for col in range(1, len(headers) + 1):
        c = ws.cell(1, col)
        c.fill, c.font, c.alignment = fill, font, Alignment(horizontal="center")
    for it in items:
        dt = datetime.fromisoformat(it["ts"])
        ws.append([dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M"), it.get("bank", ""), it.get("platform", ""),
                   it.get("type", ""), it.get("summary", it.get("title", "")), it.get("sentiment", ""),
                   it.get("engagement", 0), it.get("url", "")])
    wb.save(XLSX_FILE)

def notify_whatsapp(nbe_items):
    sid, token = os.getenv("TWILIO_SID"), os.getenv("TWILIO_TOKEN")
    w_from, w_to = os.getenv("TWILIO_FROM"), os.getenv("WHATSAPP_TO")
    if not all([sid, token, w_from, w_to]) or not nbe_items:
        return
    msg = f"تنبيه رصد {NBE} ({len(nbe_items)} جديد):\n\n" + "".join(f"- [{i.get('platform')}] {i.get('summary')}\n{i.get('url')}\n\n" for i in nbe_items[:5])
    try:
        tw_url = "".join(["https://", "api.", "twilio.", "com/2010-04-01/Accounts/", sid, "/Messages.json"])
        requests.post(tw_url, auth=(sid, token), data={"From": w_from, "To": w_to, "Body": msg[:1500]}, timeout=20)
    except Exception as ex:
        print("whatsapp error", ex)

def main():
    prev = {"items": [], "seen": [], "benchmark": None}
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                if isinstance(loaded := json.load(f), dict):
                    prev.update(loaded)
        except Exception as ex:
            print("load error", ex)
    seen = set(prev.get("seen", []))
    raw = fetch_news() + fetch_reddit() + fetch_youtube()
    fresh = [x for x in raw if x["id"] not in seen]
    print(f"Fetched: {len(raw)}, Fresh: {len(fresh)}")
    kept = analyze(fresh)
    for x in fresh:
        if x.get("_done"):
            seen.add(x["id"])
            x.pop("_done", None)
    all_items = sorted(kept + prev.get("items", []), key=lambda x: x.get("ts", ""), reverse=True)[:MAX_ITEMS]
    bench = benchmark(all_items, prev.get("benchmark"))
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump({"updated": now().isoformat(), "items": all_items, "seen": list(seen)[-10000:], "benchmark": bench}, f, ensure_ascii=False, indent=2)
    export_excel(all_items)
    notify_whatsapp([x for x in kept if x.get("bank") == NBE])
    print(f"Done. Kept: {len(kept)}, Total: {len(all_items)}")

if __name__ == "__main__":
    main()
