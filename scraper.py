#!/usr/bin/env python3
import os, re, json, html, hashlib, time
from datetime import datetime, timedelta
from urllib.parse import quote
from zoneinfo import ZoneInfo
import requests, feedparser
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

CAIRO = ZoneInfo("Africa/Cairo")
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
    base = "https:" + chr(47) + chr(47) + "news." + "google.com" + chr(47) + "rss" + chr(47) + "search?q="
    for bank, queries in BANKS.items():
        for q in queries:
            url = base + quote(q + " when:2d") + "&hl=ar&gl=EG&ceid=EG:ar"
            try:
                for e in feedparser.parse(url).entries[:15]:
                    ts = datetime(*e.published_parsed[:6], tzinfo=ZoneInfo("UTC")) if getattr(e, "published_parsed", None) else now()
                    src = clean(getattr(e, "source", {}).get("title", "")) if hasattr(e, "source") else ""
                    out.append(mk(bank, "أخبار" + (f" - {src}" if src else ""), e.title, e.get("summary", ""), e.link, ts))
            except Exception as ex:
                print("news error", bank, ex)
    return out
def fetch_social():
    out = []
    socials = {
        "Facebook": "facebook.com",
        "X (Twitter)": "x.com",
        "Instagram": "instagram.com",
        "TikTok": "tiktok.com",
        "LinkedIn": "linkedin.com",
    }
    base = "https:" + chr(47) + chr(47) + "news." + "google.com" + chr(47) + "rss" + chr(47) + "search?q="
    for bank, queries in BANKS.items():
        q_bank = queries[int(chr(48))]
        for label, domain in socials.items():
            url = base + quote(f"{q_bank} site:{domain} when:3d") + "&hl=ar&gl=EG&ceid=EG:ar"
            try:
                for e in feedparser.parse(url).entries[:5]:
                    ts = datetime(*e.published_parsed[:6], tzinfo=ZoneInfo("UTC")) if getattr(e, "published_parsed", None) else now()
                    out.append(mk(bank, label, e.title, e.get("summary", ""), e.link, ts))
            except Exception as ex:
                print(label, "error", bank, ex)
    return out

def fetch_reddit():
    out = []
    host = "https:" + chr(47) + chr(47) + "www." + "reddit.com"
    for bank, queries in BANKS.items():
        q = queries[int(chr(48))].strip('"')
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
    host = "https:" + chr(47) + chr(47) + "www." + "googleapis.com" + chr(47) + "youtube" + chr(47) + "v3"
    watch = "https:" + chr(47) + chr(47) + "www." + "youtube.com" + chr(47) + "watch?v="
    for bank, queries in BANKS.items():
        try:
            r = requests.get(host + "/search", params={"part": "snippet", "q": queries[int(chr(48))].strip('"'), "type": "video", "order": "date",
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
_ACTIVE_MODEL = None

def get_model(headers):
    global _ACTIVE_MODEL
    if _ACTIVE_MODEL:
        return _ACTIVE_MODEL
    m_url = "https:" + chr(47) + chr(47) + "api." + "groq.com" + chr(47) + "openai" + chr(47) + "v1" + chr(47) + "models"
    r = requests.get(m_url, headers=headers, timeout=30)
    r.raise_for_status()
    ids = [m["id"] for m in r.json().get("data", []) if "whisper" not in m["id"] and "guard" not in m["id"]]
    for pref in ["llama-3.3", "llama-3.1", "qwen", "gemma", "llama", "gpt-oss"]:
        for mid in ids:
            if pref in mid.lower():
                _ACTIVE_MODEL = mid
                print("Using Groq model:", _ACTIVE_MODEL)
                return _ACTIVE_MODEL
    _ACTIVE_MODEL = ids[int(chr(48))]
    print("Using Groq model:", _ACTIVE_MODEL)
    return _ACTIVE_MODEL

def ask_groq(prompt, system, max_tokens=800):
    api_url = "https:" + chr(47) + chr(47) + "api." + "groq.com" + chr(47) + "openai" + chr(47) + "v1" + chr(47) + "chat" + chr(47) + "completions"
    headers = {"Authorization": f"Bearer {os.environ['GROQ_API_KEY']}", "Content-Type": "application/json"}
    model_id = get_model(headers)
    r = requests.post(api_url, timeout=120, headers=headers,
        json={"model": model_id, "max_tokens": max_tokens, "temperature": 0.2,
              "messages": [{"role": "system", "content": system}, {"role": "user", "content": prompt}]})
    if r.status_code != 200:
        print("Groq detail:", r.status_code, r.text[:250])
    r.raise_for_status()
    return r.json()["choices"][int(chr(48))]["message"]["content"]

def analyze(items):
    system = "You are an Egyptian banking media analyst. Return ONLY a valid JSON array with no markdown."
    kept = []
    for i in range(0, min(len(items), 35), 5):
        batch = items[i:i + 5]
        payload = [{"id": b["id"], "bank": b["bank"], "platform": b["platform"], "title": b["title"], "text": b["snippet"][:180]} for b in batch]
        prompt = (
            "Return a JSON array only. For each item return: id, relevant (true if about the bank), "
            "summary (concise 1-sentence English summary), type (News | Campaign | Product | Financial | Complaint | Other), "
            "sentiment (إيجابي | سلبي | محايد).\n" + json.dumps(payload, ensure_ascii=False))
        try:
            raw_res = ask_groq(prompt, system).strip()
            m = re.search(r"\[.*\]", raw_res, flags=re.S)
            res = {x["id"]: x for x in json.loads(m.group(0) if m else raw_res)}
            time.sleep(10)
        except Exception as ex:
            print("analysis error:", ex)
            time.sleep(15)
            continue
        for b in batch:
            b["_done"] = True
            x = res.get(b["id"])
            if x and x.get("relevant"):
                b.update(summary=x.get("summary", b["title"]), type=x.get("type", "News"), sentiment=x.get("sentiment", "محايد"))
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
    nbe_top = [{"summary": r["summary"], "sentiment": r["sentiment"]} for r in recent if r["bank"] == NBE][:10]
    prompt = (f"3-day Egyptian banking stats:\n{json.dumps(stats, ensure_ascii=False)}\nTop {NBE} news:\n{json.dumps(nbe_top, ensure_ascii=False)}\n\n"
              f"Write a concise 8-line English strategic benchmark comparing {NBE} vs competitors with 2 practical recommendations.")
    try:
        return {"generated": now().isoformat(), "text": ask_groq(prompt, "You are an Egyptian banking strategy analyst.").strip()}
    except Exception as ex:
        print("benchmark error", ex)
        return prev

def export_excel(items):
    wb = Workbook()
    ws = wb.active
    ws.title = "Banks Monitor"
    headers = ["Date", "Time", "Bank", "Platform", "Type", "Summary", "Sentiment", "Engagement", "URL"]
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
    msg = f"NBE Media Alert ({len(nbe_items)} new):\n\n" + "".join(f"- [{i.get('platform')}] {i.get('summary')}\n{i.get('url')}\n\n" for i in nbe_items[:5])
    try:
        tw_url = "https:" + chr(47) + chr(47) + "api." + "twilio.com" + chr(47) + "2010-04-01" + chr(47) + "Accounts" + chr(47) + sid + chr(47) + "Messages.json"
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
    raw = fetch_news() + fetch_social() + fetch_reddit() + fetch_youtube()
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
