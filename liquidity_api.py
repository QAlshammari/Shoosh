"""Small Flask API for the Q Options liquidity page."""
from datetime import date, datetime, time as dtime, timedelta, timezone
from html import unescape
from html.parser import HTMLParser
import io
import os
import re
import time
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

NASDAQ_URL = "https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqlisted.txt"
OTHER_URL = "https://www.nasdaqtrader.com/dynamic/SymDir/otherlisted.txt"
BLS_ICS_URL = "https://www.bls.gov/schedule/news_release/bls.ics"
BEA_RELEASES_URL = "https://apps.bea.gov/API/signup/release_dates.json"
FED_FOMC_URL = "https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm"
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


def get_universe(include_etfs=False):
    def load(url, symbol_col):
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        frame = pd.read_csv(io.StringIO(response.text), sep="|")
        frame = frame[~frame[symbol_col].astype(str).str.startswith("File Creation")]
        if "Test Issue" in frame:
            frame = frame[frame["Test Issue"] == "N"]
        if not include_etfs and "ETF" in frame:
            frame = frame[frame["ETF"] == "N"]
        return frame[symbol_col].dropna().astype(str).tolist()

    symbols = load(NASDAQ_URL, "Symbol") + load(OTHER_URL, "ACT Symbol")
    symbols = [s.strip().replace(".", "-") for s in symbols]
    symbols = [s for s in symbols if s.isascii() and "$" not in s and len(s) <= 5]
    return sorted(set(symbols))


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
    days_since_sunday = (now.weekday() + 1) % 7
    week_start = now.date() - timedelta(days=days_since_sunday)
    week_end = week_start + timedelta(days=6)
    end = datetime.combine(week_end, dtime(23, 59, 59), tzinfo=riyadh)
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
        if now <= when <= end:
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
    include_etfs = bool(body.get("include_etfs", False))
    try:
        symbols = get_universe(include_etfs)
        rows = []
        batch_size = 150
        for start in range(0, len(symbols), batch_size):
            rows.extend(batch_metrics(symbols[start:start + batch_size]))
            if start + batch_size < len(symbols):
                time.sleep(0.5)
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
