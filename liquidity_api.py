"""Small Flask API for the Q Options liquidity page."""
from datetime import date, datetime, time as dtime, timedelta, timezone
from html import unescape
from html.parser import HTMLParser
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import yfinance as yf
from flask import Flask, abort, jsonify, request, send_from_directory
from flask_cors import CORS

APP_DIR = Path(__file__).resolve().parent
app = Flask(__name__, static_folder=None)
CORS(app, origins=os.getenv("LIQUIDITY_CORS_ORIGIN", "*"))

BLS_ICS_URL = "https://www.bls.gov/schedule/news_release/bls.ics"
BEA_RELEASES_URL = "https://apps.bea.gov/API/signup/release_dates.json"
FED_FOMC_URL = "https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm"
MIZAN_SCREEN_URL = "https://askmizan.com/api/v1/screen/{}"
# Curated names frequently followed by US options traders. The live screen below
# decides which of these can appear; this list is never treated as a halal list.
POPULAR_WATCHLIST = [
    "AAPL", "NVDA", "MSFT", "AMZN", "GOOGL", "META", "TSLA", "AMD",
    "AVGO", "TSM", "MU", "PLTR", "NFLX", "ORCL", "QCOM", "INTC",
    "CRM", "ADBE", "SHOP", "UBER", "COIN", "MSTR", "HOOD", "SOFI",
    "LLY", "COST", "WMT", "JPM", "XOM", "DIS", "ARM", "SMCI",
]
SCREEN_CACHE = {"expires": 0, "reports": [], "failed": 0}
SCREEN_CACHE_SECONDS = 6 * 60 * 60
PUBLIC_SUFFIXES = {".html", ".css", ".js", ".png", ".jpg", ".jpeg", ".svg", ".ico", ".webmanifest"}


@app.get("/")
def home():
    return send_from_directory(APP_DIR, "index.html")


@app.get("/<path:filename>")
def public_files(filename):
    path = Path(filename)
    if path.suffix.lower() not in PUBLIC_SUFFIXES or any(part.startswith(".") for part in path.parts):
        abort(404)
    return send_from_directory(APP_DIR, filename)


def _screen_one(ticker):
    response = requests.get(MIZAN_SCREEN_URL.format(ticker), timeout=20,
                            headers={"User-Agent": "QOptionsLiquidity/1.0"})
    response.raise_for_status()
    report = response.json()
    standards = {row.get("standard"): row for row in report.get("screens", [])}
    required = ("AAOIFI", "DJIM", "SP", "MSCI")
    business = report.get("business", {})
    closer_look = business.get("needs_a_closer_look_under", [])
    eligible = (
        all(standards.get(name, {}).get("compliant") is True for name in required)
        and not closer_look
        and report.get("summary", {}).get("business_screen") == "pass"
    )
    return {"ticker": ticker, "eligible": eligible, "as_of": report.get("as_of"),
            "company": report.get("company"), "borderline": any(standards.get(name, {}).get("borderline") for name in required)}


def get_screened_watchlist():
    now = time.time()
    if SCREEN_CACHE["expires"] > now:
        return SCREEN_CACHE["reports"]
    reports, errors = [], []
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {pool.submit(_screen_one, ticker): ticker for ticker in POPULAR_WATCHLIST}
        for future in as_completed(futures):
            ticker = futures[future]
            try:
                reports.append(future.result())
            except Exception as exc:
                errors.append(f"{ticker}: {exc}")
    if not reports:
        raise RuntimeError("مصدر فحص التوافق الشرعي غير متاح الآن؛ لم يتم إجراء فحص الأسهم. أعيدي المحاولة لاحقًا.")
    SCREEN_CACHE.update(expires=now + SCREEN_CACHE_SECONDS, reports=reports, failed=len(errors))
    return reports


def batch_metrics(tickers, window=30):
    frame = yf.download(tickers, period="2mo", interval="1d", group_by="ticker",
                        auto_adjust=True, threads=True, progress=False, timeout=30)
    if frame is None or frame.empty:
        return []
    rows = []
    multi = isinstance(frame.columns, pd.MultiIndex)
    for ticker in tickers:
        try:
            data = frame[ticker] if multi else frame
            data = data.dropna(subset=["Close"]).tail(window)
            if len(data) < 15:
                continue
            dollar = data["Close"] * data["Volume"]
            prior_close = float(data["Close"].iloc[-2])
            prior_avg_volume = float(data["Volume"].iloc[:-1].mean())
            returns = data["Close"].pct_change().abs()
            valid = dollar > 0
            amihud = float((returns[valid] / dollar[valid]).mean() * 1e6) if valid.any() else None
            rows.append({
                "ticker": ticker,
                "price": float(data["Close"].iloc[-1]),
                "daily_change_pct": ((float(data["Close"].iloc[-1]) / prior_close) - 1) * 100 if prior_close else None,
                "relative_volume": float(data["Volume"].iloc[-1]) / prior_avg_volume if prior_avg_volume else None,
                "last_dollar_flow": float(dollar.iloc[-1]),
                "avg_volume": float(data["Volume"].mean()),
                "avg_dollar_vol": float(dollar.mean()),
                "median_dollar_vol": float(dollar.median()),
                "zero_vol_days": int((data["Volume"] == 0).sum()),
                "amihud": amihud,
                "latest_session": data.index[-1].strftime("%Y-%m-%d"),
            })
        except (KeyError, TypeError, ValueError, IndexError):
            continue
    return rows


def market_pulse():
    symbols = ["SPY", "QQQ", "DIA", "IWM"]
    frame = yf.download(symbols, period="5d", interval="1d", group_by="ticker",
                        auto_adjust=True, threads=True, progress=False, timeout=30)
    pulse = {}
    for ticker in symbols:
        try:
            data = frame[ticker]["Close"].dropna()
            pulse[ticker] = {"change_pct": (float(data.iloc[-1]) / float(data.iloc[-2]) - 1) * 100}
        except (KeyError, TypeError, ValueError, IndexError):
            pulse[ticker] = {"change_pct": None}
    return pulse


def classify(row):
    median = row["median_dollar_vol"]
    if row["zero_vol_days"] >= 3 or median < 1_000_000:
        return "illiquid"
    if median >= 50_000_000:
        return "high"
    if median >= 5_000_000:
        return "medium"
    return "low"


@app.get("/api/health")
def health():
    return jsonify({"ok": True})


def _parse_ics_datetime(value, tz_name=None):
    value = value.strip()
    if value.endswith("Z"):
        return datetime.strptime(value, "%Y%m%dT%H%M%SZ").replace(tzinfo=timezone.utc)
    if len(value) == 8:
        local_date = datetime.strptime(value, "%Y%m%d").date()
        return datetime.combine(local_date, dtime(8, 30), tzinfo=ZoneInfo(tz_name or "America/New_York"))
    zone = ZoneInfo(tz_name or "America/New_York")
    return datetime.strptime(value, "%Y%m%dT%H%M%S").replace(tzinfo=zone)


def _bls_ics_events(text):
    lines = []
    for line in text.replace("\r\n", "\n").split("\n"):
        if line.startswith((" ", "\t")) and lines:
            lines[-1] += line[1:]
        else:
            lines.append(line)
    events, current = [], None
    for line in lines:
        if line == "BEGIN:VEVENT":
            current = {}
        elif line == "END:VEVENT" and current is not None:
            title = current.get("SUMMARY", "").replace("\\,", ",").replace("\\;", ";").replace("\\n", " ").strip()
            if title and current.get("DTSTART"):
                events.append({"title": unescape(title), "datetime": current["DTSTART"], "source": "BLS"})
            current = None
        elif current is not None and ":" in line:
            key, value = line.split(":", 1)
            if key.startswith("DTSTART"):
                tz_match = re.search(r"TZID=([^;:]+)", key)
                current["DTSTART"] = _parse_ics_datetime(value, tz_match.group(1) if tz_match else None)
            elif key.startswith("SUMMARY"):
                current["SUMMARY"] = value
    return events


class _PlainText(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        if data.strip():
            self.parts.append(data.strip())


def _fomc_events(year):
    response = requests.get(FED_FOMC_URL, timeout=30, headers={"User-Agent": "QOptionsMarketCalendar/1.0"})
    response.raise_for_status()
    parser = _PlainText()
    parser.feed(response.text)
    text = "\n".join(parser.parts)
    match = re.search(rf"{year} FOMC Meetings(.*?)(?:\d{{4}} FOMC Meetings|Future Year:)", text, re.S)
    if not match:
        match = re.search(rf"{year} FOMC Meetings(.*?)(?:\d{{4}} FOMC Meetings|Last Update:)", text, re.S)
    if not match:
        return []
    months = {name: number for number, name in enumerate(("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"), 1)}
    lines = [line.strip() for line in match.group(1).splitlines() if line.strip()]
    found = []
    for index, line in enumerate(lines[:-1]):
        if line in months:
            day_match = re.match(r"(\d{1,2})(?:\s*[-–]\s*(\d{1,2}))?\*?$", lines[index + 1])
            if day_match:
                first = date(year, months[line], int(day_match.group(1)))
                last_day = day_match.group(2) or day_match.group(1)
                found.append({"title": f"اجتماع اللجنة الفيدرالية للسوق المفتوحة FOMC ({day_match.group(1)}–{last_day})", "datetime": datetime.combine(first, dtime(12), tzinfo=ZoneInfo("Asia/Riyadh")), "source": "Federal Reserve", "source_url": FED_FOMC_URL, "impact": "مرتفع", "all_day": True})
    return found


def _arabic_event_title(title):
    translations = (
        ("Consumer Price Index", "مؤشر أسعار المستهلك (CPI)"),
        ("Employment Situation", "تقرير الوظائف الأمريكي"),
        ("Producer Price Index", "مؤشر أسعار المنتجين (PPI)"),
        ("Job Openings and Labor Turnover", "فرص العمل ودوران العمالة (JOLTS)"),
        ("Employment Cost Index", "مؤشر تكلفة التوظيف"),
        ("Real Earnings", "الأجور الحقيقية"),
        ("U.S. Import and Export Price Indexes", "مؤشرات أسعار الواردات والصادرات الأمريكية"),
        ("Gross Domestic Product", "الناتج المحلي الإجمالي (GDP)"),
        ("GDP", "الناتج المحلي الإجمالي (GDP)"),
        ("Personal Income and Outlays", "الدخل الشخصي والإنفاق"),
        ("U.S. International Trade in Goods and Services", "التجارة الأمريكية في السلع والخدمات"),
    )
    for english, arabic in translations:
        if title.startswith(english):
            remainder = title[len(english):].lstrip(" ,—–-:")
            return f"{arabic} — {remainder}" if remainder else arabic
    return title


@app.get("/api/events")
def events():
    riyadh = ZoneInfo("Asia/Riyadh")
    now = datetime.now(riyadh)
    anchor_raw = request.args.get("anchor", "")
    try:
        anchor = date.fromisoformat(anchor_raw) if anchor_raw else now.date()
    except ValueError:
        return jsonify({"error": "تاريخ التقويم غير صالح."}), 400
    week_start = anchor - timedelta(days=anchor.weekday())
    week_end = week_start + timedelta(days=4)
    end = datetime.combine(week_end, dtime(23, 59, 59), tzinfo=riyadh)
    start = datetime.combine(week_start, dtime.min, tzinfo=riyadh)
    candidates, successful_sources, errors = [], [], []

    try:
        response = requests.get(BLS_ICS_URL, timeout=30, headers={"User-Agent": "QOptionsMarketCalendar/1.0"})
        response.raise_for_status()
        candidates.extend({**event, "source_url": "https://www.bls.gov/schedule/"} for event in _bls_ics_events(response.text))
        successful_sources.append("BLS")
    except Exception as exc:
        errors.append(f"BLS: {exc}")

    try:
        response = requests.get(BEA_RELEASES_URL, timeout=30, headers={"User-Agent": "QOptionsMarketCalendar/1.0"})
        response.raise_for_status()
        for title, record in response.json().items():
            for timestamp in record.get("release_dates", []):
                value = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
                candidates.append({"title": title, "datetime": value, "source": "BEA", "source_url": "https://www.bea.gov/news/schedule/"})
        successful_sources.append("BEA")
    except Exception as exc:
        errors.append(f"BEA: {exc}")

    try:
        candidates.extend(_fomc_events(now.year))
        successful_sources.append("Federal Reserve")
    except Exception as exc:
        errors.append(f"Federal Reserve: {exc}")

    if not successful_sources:
        app.logger.warning("Event sources unavailable: %s", "; ".join(errors))
        return jsonify({"error": "تعذر الوصول إلى جداول الجهات الرسمية الآن. أعيدي المحاولة لاحقًا."}), 502

    high_impact_terms = ("consumer price index", "employment situation", "gross domestic product", "personal income and outlays", "fomc", "employment cost index")
    selected, seen = [], set()
    for event in candidates:
        when = event["datetime"].astimezone(riyadh)
        if start <= when <= end:
            key = (event["source"], event["title"], when.isoformat())
            if key in seen:
                continue
            seen.add(key)
            title_low = event["title"].lower()
            event["impact"] = event.get("impact") or ("مرتفع" if any(term in title_low for term in high_impact_terms) else "مهم")
            selected.append({"title": _arabic_event_title(event["title"]), "datetime": when.isoformat(), "source": event["source"], "source_url": event["source_url"], "impact": event["impact"], "all_day": bool(event.get("all_day"))})
    selected.sort(key=lambda item: item["datetime"])
    return jsonify({"week_start": week_start.isoformat(), "week_end": week_end.isoformat(), "updated_at": now.isoformat(), "sources_checked": successful_sources, "events": selected})


@app.post("/api/liquidity")
def liquidity():
    body = request.get_json(silent=True) or {}
    try:
        position = max(1.0, float(body.get("position_usd", 50_000)))
        min_dollar_vol = max(0.0, float(body.get("min_dollar_vol", 5_000_000)))
    except (TypeError, ValueError):
        return jsonify({"error": "تحققي من قيمة المركز والحد الأدنى للتداول."}), 400
    try:
        reports = get_screened_watchlist()
        eligible = [report for report in reports if report["eligible"] and not report["borderline"]]
        symbols = [report["ticker"] for report in eligible]
        if not symbols:
            return jsonify({"error": "لم يظهر أي سهم من قائمة المتابعة مستوفيًا الشروط الأربعة حاليًا. لم تُعرض أسهم غير متحقق منها."}), 422
        rows = batch_metrics(symbols)
        if not rows:
            return jsonify({"error": "لم تصل بيانات من مصدر الأسعار. أعيدي المحاولة لاحقًا."}), 502
        frame = pd.DataFrame(rows)
        frame = frame[frame["median_dollar_vol"] >= min_dollar_vol].copy()
        frame["tier"] = frame.apply(classify, axis=1)
        frame["days_to_liquidate"] = position / (0.10 * frame["median_dollar_vol"])
        frame = frame.sort_values("median_dollar_vol", ascending=False)
        result = frame.where(pd.notna(frame), None).to_dict(orient="records")
        tiers = frame["tier"].value_counts().to_dict()
        return jsonify({
            "scanned": len(rows),
            "watchlist_size": len(POPULAR_WATCHLIST),
            "screened": len(reports),
            "eligible": len(eligible),
            "screen_checks_failed": SCREEN_CACHE["failed"],
            "screen_source": "Mizan · أربعة معايير منشورة",
            "screen_as_of": max((item.get("as_of", "") for item in eligible), default=""),
            "returned": len(result),
            "total_last_dollar_flow": float(frame["last_dollar_flow"].sum()) if len(frame) else 0,
            "tiers": tiers,
            "latest_session": max(row["latest_session"] for row in rows),
            "market_pulse": market_pulse(),
            "updated_at": datetime.now(timezone.utc).isoformat(),
            "results": result,
        })
    except Exception as exc:
        app.logger.exception("Liquidity scan failed")
        return jsonify({"error": f"تعذر إكمال الفحص: {str(exc)[:180]}"}), 502


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
