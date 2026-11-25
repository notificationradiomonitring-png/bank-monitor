#!/usr/bin/env python3
"""محرك الرصد الإعلامي للبنوك المصرية - يعمل كل ساعة عبر GitHub Actions."""
import os, re, json, html, hashlib, time
from datetime import datetime, timedelta
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests, feedparser
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

CAIRO = ZoneInfo("Africa/Cairo")
MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-5-5")
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


# ---------------------------------------------------------------- المصادر
def fetch_news():
    out = []
    for bank, queries in BANKS.items():
        for q in queries:
            url = ("https://news.google.com/rss/search?q=" + quote(q + " when:2d")
                   + "&hl=ar&gl=EG&ceid=EG:ar")
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
    for bank, queries in BANKS.items():
        q = queries[0].strip('"')
        try:
            r = requests.get("https://www.reddit.com/search.json",
                             params={"q": q, "sort": "new", "t": "week", "limit": 10},
                             headers=UA, timeout=20)
            if r.status_code != 200:
                continue
            for c in r.json()["data"]["children"]:
                d = c["data"]
                out.append(mk(bank, "Reddit", d["title"], d.get("selftext", ""),
                              "https://www.reddit.com" + d["permalink"],
                              datetime.fromtimestamp(d["created_utc"], ZoneInfo("UTC")),
                              d.get("score", 0) + d.get("num_comments", 0)))
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
    for bank, queries in BANKS.items():
        try:
            r = requests.get("https://www.googleapis.com/youtube/v3/search", params={
                "part": "snippet", "q": queries[0].strip('"'), "type": "video", "order": "date",
                "publishedAfter": after, "maxResults": 8, "regionCode": "EG", "key": key}, timeout=20).json()
            ids = [i["id"]["videoId"] for i in r.get("items", [])]
            if not ids:
                continue
            stats = requests.get("https://www.googleapis.com/youtube/v3/videos",
                                 params={"part": "statistics", "id": ",".join(ids), "key": key}, timeout=20).json()
            views = {v["id"]: int(v["statistics"].get("viewCount", 0)) for v in stats.get("items", [])}
            for i in r["items"]:
                s, vid = i["snippet"], i["id"]["videoId"]
                ts = datetime.fromisoformat(s["publishedAt"].replace("Z", "+00:00"))
                out.append(mk(bank, "YouTube", s["title"], s.get("description", ""),
                              f"https://www.youtube.com/watch?v={vid}", ts, views.get(vid, 0)))
        except Exception as ex:
            print("youtube error", bank, ex)
    return out


# ---------------------------------------------------------------- Claude
def claude(prompt, system, max_tokens=4000):
    r = requests.post("https://api.anthropic.com/v1/messages", timeout=120, headers={
        "x-api-key": os.environ["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01",
        "content-type": "application/json"},
        json={"model": MODEL, "max_tokens": max_tokens, "system": system,
              "messages": [{"role": "user", "content": prompt}]})
    r.raise_for_status()
    return "".join(b.get("text", "") for b in r.json()["content"])


def analyze(items):
    """يصنّف العناصر الجديدة. يعيد قائمة العناصر ذات الصلة فقط."""
    system = ("أنت محلل رصد إعلامي للقطاع المصرفي المصري. أجب بـ JSON فقط بدون أي نص آخر أو علامات Markdown.")
    kept = []
    for i in range(0, len(items), 20):
        batch = items[i:i + 20]
        payload = [{"id": b["id"], "bank": b["bank"], "platform": b["platform"],
                    "title": b["title"], "text": b["snippet"]} for b in batch]
        prompt = (
            "لكل عنصر أرجع كائناً بالحقول: id, relevant (true إذا كان المحتوى يخص البنك المذكور فعلاً "
            "وليس مجرد ذكر عابر أو اسم مشابه), summary (ملخص عربي في جملة أو جملتين), "
            "type (خبر | إعلان/حملة | منتج/خدمة | مالي/نتائج | شكوى/خدمة عملاء | تنظيمي | أخرى), "
            "sentiment (إيجابي | سلبي | محايد) تجاه البنك المذكور.\n"
            "أرجع مصفوفة JSON فقط.\n" + json.dumps(payload, ensure_ascii=False))
        try:
            txt = re.sub(r"^```(?:json)?|```$", "", claude(prompt, system).strip(), flags=re.M).strip()
            res = {x["id"]: x for x in json.loads(txt)}
        except Exception as ex:
            print("analysis error", ex)
            continue  # لن تُسجَّل كمرئية، وستُعاد محاولتها في الساعة القادمة
        for b in batch:
            b["_done"] = True
            x = res.get(b["id"])
            if x and x.get("relevant"):
                b.update(summary=x.get("summary", b["title"]), type=x.get("type", "أخرى"),
                         sentiment=x.get("sentiment", "محايد"))
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
    stats = {}
    for b in BANKS:
        rows = [i for i in recent if i["bank"] == b]
        stats[b] = {"count": len(rows),
                    "positive": sum(r["sentiment"] == "إيجابي" for r in rows),
                    "negative": sum(r["sentiment"] == "سلبي" for r in rows),
                    "neutral": sum(r["sentiment"] == "محايد" for r in rows)}
    nbe_top = [{"summary": r["summary"], "sentiment": r["sentiment"]} for r in recent if r["bank"] == NBE][:15]
    prompt = (f"هذه إحصاءات الرصد الإعلامي آخر 3 أيام للبنوك المصرية:\n{json.dumps(stats, ensure_ascii=False)}\n"
              f"أبرز أخبار {NBE}:\n{json.dumps(nbe_top, ensure_ascii=False)}\n\n"
              f"اكتب تقرير مقارنة عربياً مختصراً (حتى 10 أسطر) يقارن {NBE} بالمنافسين: حصة الظهور، "
              "توازن الانطباع، أبرز الحملات أو الأخبار، وتوصيتين عمليتين. لا تخترع أرقاماً غير موجودة.")
    try:
        text = claude(prompt, "أنت محلل استراتيجي في العلاقات العامة المصرفية.", 1500)
    except Exception as ex:
        print("benchmark error", ex)
        return prev
    return {"generated": now().isoformat(), "period_days": 3, "stats": stats, "analysis": text.strip()}


# ---------------------------------------------------------------- المخرجات
def save_excel(items):
    wb = Workbook(); ws = wb.active; ws.title = "الرصد"; ws.sheet_view.rightToLeft = True
    head = ["التاريخ", "الوقت", "البنك", "المنصة", "الملخص", "الرابط", "التقييم"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0E6B4F")
        c.alignment = Alignment(horizontal="center")
    for i in sorted(items, key=lambda x: x["ts"], reverse=True):
        d = datetime.fromisoformat(i["ts"])
        ws.append([d.strftime("%Y-%m-%d"), d.strftime("%H:%M"), i["bank"], i["platform"],
                   i["summary"], i["url"], i["sentiment"]])
    for col, w in zip("ABCDEFG", [12, 8, 20, 22, 70, 50, 10]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        row[4].alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    wb.save(XLSX_FILE)


def notify_whatsapp(new_nbe):
    sid, tok = os.getenv("TWILIO_SID"), os.getenv("TWILIO_TOKEN")
    to, frm = os.getenv("WHATSAPP_TO"), os.getenv("TWILIO_FROM", "+14155238886")
    if not (sid and tok and to) or not new_nbe:
        return
    for i in new_nbe[:5]:
        body = f"🔔 {NBE}\n{i['summary']}\nالانطباع: {i['sentiment']} | {i['platform']}\n{i['url']}"
        try:
            requests.post(f"https://api.twilio.com/2010-04-01/Accounts/{sid}/Messages.json",
                          auth=(sid, tok), timeout=20,
                          data={"From": f"whatsapp:{frm}", "To": f"whatsapp:{to}", "Body": body})
        except Exception as ex:
            print("whatsapp error", ex)


# ---------------------------------------------------------------- التشغيل
def main():
    data = {"items": [], "seen": [], "benchmark": None}
    if os.path.exists(DATA_FILE):
        data.update(json.load(open(DATA_FILE, encoding="utf-8")))
    seen = set(data["seen"])

    fresh, ids = [], set()
    for i in fetch_news() + fetch_reddit() + fetch_youtube():
        if i["id"] not in seen and i["id"] not in ids:
            ids.add(i["id"]); fresh.append(i)
    print(f"عناصر جديدة: {len(fresh)}")

    kept = analyze(fresh) if fresh else []
    for i in fresh:
        if i.pop("_done", False):
            seen.add(i["id"])
    for i in kept:
        i.pop("snippet", None); i.pop("_done", None)
        d = datetime.fromisoformat(i["ts"])
        i["date"], i["time"] = d.strftime("%Y-%m-%d"), d.strftime("%H:%M")

    items = sorted(data["items"] + kept, key=lambda x: x["ts"], reverse=True)[:MAX_ITEMS]
    data.update(items=items, seen=list(seen)[-8000:], updated=now().isoformat(),
                benchmark=benchmark(items, data.get("benchmark")))
    json.dump(data, open(DATA_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    save_excel(items)
    notify_whatsapp([i for i in kept if i["bank"] == NBE])


if __name__ == "__main__":
    main()
