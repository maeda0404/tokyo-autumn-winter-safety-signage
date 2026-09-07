import io
import json
import re
import time
import urllib.request
from datetime import date, datetime, timedelta, timezone
from html import unescape
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import load_workbook


# =========================================================
# 基本設定
# =========================================================

JST = timezone(timedelta(hours=9))
ROOT_DIR = Path(__file__).resolve().parents[1]
STATUS_FILE = ROOT_DIR / "data" / "status.json"

WARNING_URL = "https://www.jma.go.jp/bosai/warning/data/warning/130000.json"
FORECAST_URL = "https://www.jma.go.jp/bosai/forecast/data/forecast/130000.json"
FLU_TOP_URL = "https://idsc.tmiph.metro.tokyo.lg.jp/diseases/flu/flu/"

TOKYO_POINT_CODE = "44132"
TARGET_WARNING_AREA_CODES = {
    "130010",
    "1311300",
    "13113000",
}


# =========================================================
# 共通処理
# =========================================================

def fetch_bytes(url, retries=3):
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            request = urllib.request.Request(
                url,
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 "
                        "Tokyo-Autumn-Winter-Safety-Signage/2.0"
                    ),
                    "Accept": "*/*",
                },
            )
            with urllib.request.urlopen(request, timeout=30) as response:
                return response.read()
        except Exception as error:
            last_error = error
            print(f"通信失敗 {attempt}/{retries}: {url}: {error}")
            if attempt < retries:
                time.sleep(5 * attempt)
    raise last_error


def fetch_text(url):
    return fetch_bytes(url).decode("utf-8", errors="replace")


def fetch_json(url):
    return json.loads(fetch_bytes(url).decode("utf-8"))


def html_to_text(html):
    text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.IGNORECASE)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def load_previous_status():
    if not STATUS_FILE.exists():
        return {"influenza": {}, "weather": {}}
    try:
        return json.loads(STATUS_FILE.read_text(encoding="utf-8"))
    except Exception as error:
        print("前回データ読込エラー:", error)
        return {"influenza": {}, "weather": {}}


# =========================================================
# 気象庁データ
# =========================================================

def get_active_warning_names():
    warning_data = fetch_json(WARNING_URL)
    warning_names = []
    for area_type in warning_data.get("areaTypes", []):
        for area in area_type.get("areas", []):
            if str(area.get("code", "")) not in TARGET_WARNING_AREA_CODES:
                continue
            for warning in area.get("warnings", []):
                status = str(warning.get("status", ""))
                name = str(warning.get("name", ""))
                if name and "解除" not in status and "なし" not in status:
                    warning_names.append(name)
    return list(dict.fromkeys(warning_names))


def get_minimum_temperature():
    forecast_data = fetch_json(FORECAST_URL)
    temperatures = []
    for report in forecast_data:
        for time_series in report.get("timeSeries", []):
            for area_data in time_series.get("areas", []):
                code = str(area_data.get("area", {}).get("code", ""))
                if code != TOKYO_POINT_CODE:
                    continue
                for value in area_data.get("temps", []):
                    try:
                        temperatures.append(float(value))
                    except (TypeError, ValueError):
                        pass
    return min(temperatures) if temperatures else None


# =========================================================
# インフルエンザデータ
# =========================================================

def determine_flu_level(per_sentinel):
    if per_sentinel is None:
        return "確認中"
    if per_sentinel >= 30:
        return "警報レベル"
    if per_sentinel >= 10:
        return "注意報レベル"
    if per_sentinel >= 1:
        return "流行中"
    return "非流行"


def find_flu_excel_url(top_html):
    links = re.findall(
        r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>([\s\S]*?)</a>",
        top_html,
        flags=re.IGNORECASE,
    )
    candidates = []
    for href, label_html in links:
        complete_url = urljoin(FLU_TOP_URL, unescape(href))
        label = html_to_text(label_html)
        lower_url = complete_url.lower()
        if (".xlsx" in lower_url or ".xls" in lower_url) and (
            "hasseidoko" in lower_url or "患者報告数" in label
        ):
            candidates.append(complete_url)
    if not candidates:
        raise ValueError("東京都公式ページから患者報告数ExcelのURLを取得できません")
    return candidates[0]


def normalize(value):
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def number(value):
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = normalize(value).replace(",", "").replace("人", "")
    if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        return float(text)
    return None


def week_from_value(value, default_year):
    if isinstance(value, datetime):
        y, w, _ = value.date().isocalendar()
        return int(y), int(w)
    if isinstance(value, date):
        y, w, _ = value.isocalendar()
        return int(y), int(w)

    text = normalize(value)
    match = re.search(r"(20\d{2})年?第?(\d{1,2})週", text)
    if match:
        return int(match.group(1)), int(match.group(2))
    match = re.fullmatch(r"第?(\d{1,2})週", text)
    if match:
        return default_year, int(match.group(1))
    return None


def row_text(sheet, row_number, end_column=None):
    limit = end_column or sheet.max_column
    return " ".join(
        normalize(sheet.cell(row=row_number, column=column).value)
        for column in range(1, limit + 1)
    )


def find_latest_record(workbook):
    now = datetime.now(JST)
    candidates = []

    for sheet in workbook.worksheets:
        max_row = sheet.max_row
        max_column = sheet.max_column

        # 縦型: 各行が週、列のどこかに「定点当たり」の値がある表
        header_cells = []
        for row in range(1, min(max_row, 50) + 1):
            for column in range(1, max_column + 1):
                text = normalize(sheet.cell(row=row, column=column).value)
                score = 0
                if "定点" in text:
                    score += 5
                if "当たり" in text:
                    score += 5
                if "報告数" in text:
                    score += 3
                if score >= 8:
                    header_cells.append((score, row, column))

        for _, header_row, value_column in sorted(header_cells, reverse=True):
            current_year = now.year
            for row in range(header_row + 1, max_row + 1):
                value = number(sheet.cell(row=row, column=value_column).value)
                if value is None or not 0 <= value <= 500:
                    continue

                week_info = None
                context = row_text(sheet, row, min(max_column, max(value_column, 12)))
                explicit = re.search(r"(20\d{2})年?第?(\d{1,2})週", context)
                if explicit:
                    week_info = (int(explicit.group(1)), int(explicit.group(2)))
                else:
                    simple = re.search(r"第(\d{1,2})週", context)
                    if simple:
                        week_info = (current_year, int(simple.group(1)))

                if not week_info:
                    continue
                year, week = week_info
                if 1 <= week <= 53 and now.year - 1 <= year <= now.year:
                    candidates.append({
                        "year": year,
                        "week": week,
                        "value": value,
                        "sheet": sheet.title,
                    })

        # 横型: 列見出しが週で、行見出しが「定点当たり」の表
        for header_row in range(1, min(max_row, 50) + 1):
            week_columns = {}
            for column in range(1, max_column + 1):
                parsed = week_from_value(
                    sheet.cell(row=header_row, column=column).value,
                    now.year,
                )
                if parsed and 1 <= parsed[1] <= 53:
                    week_columns[column] = parsed
            if not week_columns:
                continue

            for row in range(header_row + 1, max_row + 1):
                label = row_text(sheet, row, min(max_column, 12))
                if "定点" not in label or "当たり" not in label:
                    continue
                for column, (year, week) in week_columns.items():
                    value = number(sheet.cell(row=row, column=column).value)
                    if value is not None and 0 <= value <= 500:
                        candidates.append({
                            "year": year,
                            "week": week,
                            "value": value,
                            "sheet": sheet.title,
                        })

    if not candidates:
        raise ValueError("Excelから最新週の定点当たり患者報告数を判定できません")

    candidates.sort(key=lambda item: (item["year"], item["week"]))
    latest_period = (candidates[-1]["year"], candidates[-1]["week"])
    latest_candidates = [
        item for item in candidates
        if (item["year"], item["week"]) == latest_period
    ]

    # 同じ最新週に複数候補がある場合、東京都全体として妥当な候補を優先する。
    # 候補値とシート名をログへ残し、取得元の検証を可能にする。
    print("最新週候補:", latest_candidates)
    return latest_candidates[0]


def iso_week_period(year, week):
    monday = date.fromisocalendar(year, week, 1)
    sunday = date.fromisocalendar(year, week, 7)
    return f"{monday.month}月{monday.day}日～{sunday.month}月{sunday.day}日"


def get_influenza_data(previous_status):
    top_html = fetch_text(FLU_TOP_URL)
    excel_url = find_flu_excel_url(top_html)
    print("インフルエンザExcel:", excel_url)

    workbook = load_workbook(
        io.BytesIO(fetch_bytes(excel_url)),
        data_only=True,
        read_only=True,
    )
    try:
        latest = find_latest_record(workbook)
    finally:
        workbook.close()

    previous_value = previous_status.get("influenza", {}).get("perSentinel")
    difference = None
    if previous_value is not None:
        try:
            difference = round(latest["value"] - float(previous_value), 2)
        except (TypeError, ValueError):
            pass

    return {
        "level": determine_flu_level(latest["value"]),
        "week": f"第{latest['week']}週",
        "period": iso_week_period(latest["year"], latest["week"]),
        "perSentinel": latest["value"],
        "previousPerSentinel": previous_value,
        "difference": difference,
        "reportedCases": None,
        "trend": "東京都公式週報",
        "provisional": True,
        "dataStatus": "取得成功",
        "sourceUrl": excel_url,
    }


# =========================================================
# status.json生成
# =========================================================

def main():
    previous_status = load_previous_status()
    previous_weather = previous_status.get("weather", {})
    errors = []
    now_text = datetime.now(JST).strftime("%Y-%m-%d %H:%M")

    try:
        warning_names = get_active_warning_names()
        warning_text = " ".join(warning_names)
        wind_active = any(word in warning_text for word in ["強風", "暴風", "風雪"])
        dry_active = "乾燥" in warning_text
    except Exception as error:
        print("警報・注意報取得エラー:", repr(error))
        errors.append("警報注意報")
        wind_active = previous_weather.get("wind", {}).get("active", False)
        dry_active = previous_weather.get("dry", {}).get("active", False)

    try:
        minimum_temperature = get_minimum_temperature()
        cold_active = minimum_temperature is not None and minimum_temperature <= 3
        cold_note = (
            f"予想最低気温 {minimum_temperature:g}℃"
            if minimum_temperature is not None
            else "予想最低気温を確認中"
        )
    except Exception as error:
        print("気温取得エラー:", repr(error))
        errors.append("気温予報")
        cold_active = previous_weather.get("cold", {}).get("active", False)
        cold_note = "予想最低気温を確認中"

    try:
        influenza_data = get_influenza_data(previous_status)
    except Exception as error:
        print("インフルエンザ取得エラー:", repr(error))
        errors.append("感染症情報")
        previous_influenza = previous_status.get("influenza", {})
        if previous_influenza.get("perSentinel") is not None:
            influenza_data = {
                **previous_influenza,
                "dataStatus": "取得失敗・前回値を維持",
            }
        else:
            influenza_data = {
                "level": "確認中",
                "week": "最新発表",
                "period": "東京都",
                "perSentinel": None,
                "previousPerSentinel": None,
                "difference": None,
                "reportedCases": None,
                "trend": "公式情報を確認中",
                "provisional": True,
                "dataStatus": "取得失敗",
                "sourceUrl": FLU_TOP_URL,
            }

    status = {
        "updated": now_text,
        "sourceStatus": "正常" if not errors else "一部取得失敗：" + "・".join(errors),
        "influenza": influenza_data,
        "weather": {
            "wind": {
                "active": wind_active,
                "normalText": "情報なし",
                "alertText": "強風注意",
                "note": "飛散・揚重確認",
            },
            "dry": {
                "active": dry_active,
                "normalText": "情報なし",
                "alertText": "火気注意",
                "note": "消火確認を徹底",
            },
            "cold": {
                "active": cold_active,
                "normalText": "情報なし",
                "alertText": "凍結注意",
                "note": cold_note,
            },
        },
    }

    STATUS_FILE.parent.mkdir(parents=True, exist_ok=True)
    temporary_file = STATUS_FILE.with_suffix(".tmp")
    temporary_file.write_text(
        json.dumps(status, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    json.loads(temporary_file.read_text(encoding="utf-8"))
    temporary_file.replace(STATUS_FILE)

    print(json.dumps(status, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
