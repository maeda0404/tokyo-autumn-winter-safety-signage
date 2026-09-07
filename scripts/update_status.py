import io
import json
import re
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

WARNING_URL = (
    "https://www.jma.go.jp/bosai/warning/data/warning/130000.json"
)

FORECAST_URL = (
    "https://www.jma.go.jp/bosai/forecast/data/forecast/130000.json"
)

FLU_TOP_URL = (
    "https://idsc.tmiph.metro.tokyo.lg.jp/diseases/flu/flu/"
)

TOKYO_POINT_CODE = "44132"

TARGET_WARNING_AREA_CODES = {
    "130010",    # 東京地方
    "1311300",   # 世田谷区として返る場合への予備
    "13113000",  # 世田谷区として返る場合への予備
}


# =========================================================
# 共通処理
# =========================================================

def fetch_bytes(url, retries=3):
    last_error = None

    for attempt in range(retries):
        try:
            request = urllib.request.Request(
                url,
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 "
                        "Tokyo-Autumn-Winter-Safety-Signage/1.0"
                    )
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=30,
            ) as response:
                return response.read()

        except Exception as error:
            last_error = error

            if attempt < retries - 1:
                print(
                    f"通信失敗。再試行します："
                    f"{attempt + 1}/{retries}"
                )

    raise last_error


def fetch_text(url):
    return fetch_bytes(url).decode(
        "utf-8",
        errors="replace",
    )


def fetch_json(url):
    return json.loads(
        fetch_bytes(url).decode("utf-8")
    )


def html_to_text(html):
    text = re.sub(
        r"<script[\s\S]*?</script>",
        " ",
        html,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"<style[\s\S]*?</style>",
        " ",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"<[^>]+>",
        " ",
        text,
    )

    text = unescape(text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def load_previous_status():
    if not STATUS_FILE.exists():
        return {
            "influenza": {},
            "weather": {},
        }

    try:
        return json.loads(
            STATUS_FILE.read_text(encoding="utf-8")
        )

    except Exception as error:
        print("前回データ読込エラー:", error)

        return {
            "influenza": {},
            "weather": {},
        }


# =========================================================
# 気象庁データ
# =========================================================

def get_active_warning_names():
    warning_data = fetch_json(WARNING_URL)
    warning_names = []

    for area_type in warning_data.get(
        "areaTypes",
        [],
    ):
        for area in area_type.get("areas", []):
            area_code = str(area.get("code", ""))

            if area_code not in TARGET_WARNING_AREA_CODES:
                continue

            for warning in area.get("warnings", []):
                status = str(
                    warning.get("status", "")
                )
                name = str(
                    warning.get("name", "")
                )

                if not name:
                    continue

                if "解除" in status or "なし" in status:
                    continue

                warning_names.append(name)

    return list(dict.fromkeys(warning_names))


def get_minimum_temperature():
    forecast_data = fetch_json(FORECAST_URL)
    temperatures = []

    for report in forecast_data:
        for time_series in report.get(
            "timeSeries",
            [],
        ):
            for area_data in time_series.get(
                "areas",
                [],
            ):
                area_code = str(
                    area_data.get(
                        "area",
                        {},
                    ).get("code", "")
                )

                if area_code != TOKYO_POINT_CODE:
                    continue

                for value in area_data.get(
                    "temps",
                    [],
                ):
                    try:
                        temperatures.append(
                            float(value)
                        )
                    except (TypeError, ValueError):
                        continue

    if not temperatures:
        return None

    return min(temperatures)


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
    anchor_pattern = re.compile(
        r"<a\b[^\"']+[\"'][^>]*>"
        r"([\s\S]*?)</a>",
        flags=re.IGNORECASE,
    )

    candidates = []

    for href, label_html in anchor_pattern.findall(
        top_html
    ):
        label = html_to_text(label_html)
        complete_url = urljoin(FLU_TOP_URL, href)

        href_lower = complete_url.lower()

        is_excel = (
            ".xlsx" in href_lower
            or ".xls" in href_lower
        )

        is_patient_report = (
            "患者報告数" in label
            or "hasseidoko" in href_lower
        )

        if is_excel and is_patient_report:
            candidates.append(complete_url)

    if not candidates:
        raise ValueError(
            "東京都公式の患者報告数Excelを取得できません"
        )

    return candidates[0]


def normalize_cell(value):
    if value is None:
        return ""

    return str(value).strip()


def to_float(value):
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    text = text.replace(",", "")
    text = text.replace("人", "")

    match = re.fullmatch(
        r"-?[0-9]+(?:\.[0-9]+)?",
        text,
    )

    if not match:
        return None

    return float(text)


def extract_year_week_from_value(value):
    if isinstance(value, datetime):
        iso_year, iso_week, _ = value.date().isocalendar()
        return iso_year, iso_week

    if isinstance(value, date):
        iso_year, iso_week, _ = value.isocalendar()
        return iso_year, iso_week

    text = normalize_cell(value)

    patterns = [
        r"(\d{4})\s*年\s*第?\s*(\d{1,2})\s*週",
        r"(\d{4})\s*[-/]\s*(\d{1,2})",
        r"第?\s*(\d{1,2})\s*週",
    ]

    for index, pattern in enumerate(patterns):
        match = re.search(pattern, text)

        if not match:
            continue

        if index < 2:
            return (
                int(match.group(1)),
                int(match.group(2)),
            )

        return (
            datetime.now(JST).year,
            int(match.group(1)),
        )

    if isinstance(value, int):
        if 1 <= value <= 53:
            return (
                datetime.now(JST).year,
                int(value),
            )

    if isinstance(value, float):
        if value.is_integer() and 1 <= value <= 53:
            return (
                datetime.now(JST).year,
                int(value),
            )

    return None


def iso_week_period(year, week_number):
    monday = date.fromisocalendar(
        year,
        week_number,
        1,
    )

    sunday = date.fromisocalendar(
        year,
        week_number,
        7,
    )

    return (
        f"{monday.month}月{monday.day}日"
        f"～"
        f"{sunday.month}月{sunday.day}日"
    )


def score_header(text):
    compact = text.replace(" ", "")

    score = 0

    if "定点" in compact:
        score += 5

    if "当たり" in compact:
        score += 5

    if "報告数" in compact:
        score += 3

    if "患者" in compact:
        score += 1

    return score


def find_best_header_column(sheet):
    best = None

    max_header_row = min(sheet.max_row, 30)

    for row_number in range(1, max_header_row + 1):
        for column_number in range(
            1,
            sheet.max_column + 1,
        ):
            text = normalize_cell(
                sheet.cell(
                    row=row_number,
                    column=column_number,
                ).value
            )

            if not text:
                continue

            score = score_header(text)

            if score <= 0:
                continue

            candidate = {
                "score": score,
                "row": row_number,
                "column": column_number,
                "text": text,
            }

            if best is None:
                best = candidate
                continue

            if candidate["score"] > best["score"]:
                best = candidate

    return best


def extract_latest_vertical_record(sheet):
    header = find_best_header_column(sheet)

    if not header:
        return None

    start_row = header["row"] + 1
    value_column = header["column"]

    records = []

    for row_number in range(
        start_row,
        sheet.max_row + 1,
    ):
        per_sentinel = to_float(
            sheet.cell(
                row=row_number,
                column=value_column,
            ).value
        )

        if per_sentinel is None:
            continue

        week_information = None

        search_start = max(1, value_column - 6)

        for column_number in range(
            search_start,
            value_column,
        ):
            value = sheet.cell(
                row=row_number,
                column=column_number,
            ).value

            parsed = extract_year_week_from_value(
                value
            )

            if parsed:
                week_information = parsed
                break

        if not week_information:
            for column_number in range(
                1,
                min(sheet.max_column, 8) + 1,
            ):
                value = sheet.cell(
                    row=row_number,
                    column=column_number,
                ).value

                parsed = extract_year_week_from_value(
                    value
                )

                if parsed:
                    week_information = parsed
                    break

        if not week_information:
            continue

        year, week_number = week_information

        if not 1 <= week_number <= 53:
            continue

        records.append(
            {
                "year": year,
                "weekNumber": week_number,
                "perSentinel": per_sentinel,
                "sheet": sheet.title,
            }
        )

    if not records:
        return None

    records.sort(
        key=lambda item: (
            item["year"],
            item["weekNumber"],
        )
    )

    return records[-1]


def extract_latest_horizontal_record(sheet):
    max_header_row = min(sheet.max_row, 30)
    records = []

    for row_number in range(1, max_header_row + 1):
        week_columns = {}

        for column_number in range(
            1,
            sheet.max_column + 1,
        ):
            value = sheet.cell(
                row=row_number,
                column=column_number,
            ).value

            parsed = extract_year_week_from_value(
                value
            )

            if parsed:
                week_columns[column_number] = parsed

        if not week_columns:
            continue

        for data_row in range(
            row_number + 1,
            sheet.max_row + 1,
        ):
            label_parts = []

            for column_number in range(
                1,
                min(sheet.max_column, 8) + 1,
            ):
                label_parts.append(
                    normalize_cell(
                        sheet.cell(
                            row=data_row,
                            column=column_number,
                        ).value
                    )
                )

            label = " ".join(label_parts)
            label_score = score_header(label)

            if label_score <= 0:
                continue

            for column_number, (
                year,
                week_number,
            ) in week_columns.items():
                per_sentinel = to_float(
                    sheet.cell(
                        row=data_row,
                        column=column_number,
                    ).value
                )

                if per_sentinel is None:
                    continue

                records.append(
                    {
                        "year": year,
                        "weekNumber": week_number,
                        "perSentinel": per_sentinel,
                        "sheet": sheet.title,
                    }
                )

    if not records:
        return None

    records.sort(
        key=lambda item: (
            item["year"],
            item["weekNumber"],
        )
    )

    return records[-1]


def extract_latest_flu_record(excel_bytes):
    workbook = load_workbook(
        filename=io.BytesIO(excel_bytes),
        data_only=True,
        read_only=True,
    )

    candidates = []

    for sheet in workbook.worksheets:
        vertical_record = extract_latest_vertical_record(
            sheet
        )

        if vertical_record:
            candidates.append(vertical_record)

        horizontal_record = (
            extract_latest_horizontal_record(sheet)
        )

        if horizontal_record:
            candidates.append(horizontal_record)

    workbook.close()

    if not candidates:
        raise ValueError(
            "Excelから最新週の定点当たり報告数を"
            "判定できません"
        )

    current_year = datetime.now(JST).year

    valid_candidates = [
        item
        for item in candidates
        if current_year - 1
        <= item["year"]
        <= current_year
        and 1 <= item["weekNumber"] <= 53
        and 0 <= item["perSentinel"] <= 500
    ]

    if not valid_candidates:
        raise ValueError(
            "Excel内に有効な最新週データがありません"
        )

    valid_candidates.sort(
        key=lambda item: (
            item["year"],
            item["weekNumber"],
        )
    )

    latest = valid_candidates[-1]

    same_week = [
        item
        for item in valid_candidates
        if item["year"] == latest["year"]
        and item["weekNumber"]
        == latest["weekNumber"]
    ]

    if same_week:
        latest = same_week[0]

    return latest


def get_influenza_data(previous_status):
    top_html = fetch_text(FLU_TOP_URL)
    excel_url = find_flu_excel_url(top_html)
    excel_bytes = fetch_bytes(excel_url)

    latest = extract_latest_flu_record(excel_bytes)

    year = latest["year"]
    week_number = latest["weekNumber"]
    per_sentinel = latest["perSentinel"]

    previous_influenza = previous_status.get(
        "influenza",
        {},
    )

    previous_value = previous_influenza.get(
        "perSentinel"
    )

    difference = None

    if previous_value is not None:
        try:
            difference = round(
                per_sentinel - float(previous_value),
                2,
            )
        except (TypeError, ValueError):
            difference = None

    return {
        "level": determine_flu_level(per_sentinel),
        "week": f"第{week_number}週",
        "period": iso_week_period(
            year,
            week_number,
        ),
        "perSentinel": per_sentinel,
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
    errors = []

    now_text = datetime.now(JST).strftime(
        "%Y-%m-%d %H:%M"
    )

    try:
        warning_names = get_active_warning_names()
        warning_text = " ".join(warning_names)

        wind_active = any(
            keyword in warning_text
            for keyword in [
                "強風",
                "暴風",
                "風雪",
            ]
        )

        dry_active = "乾燥" in warning_text

    except Exception as error:
        print(
            "警報・注意報取得エラー:",
            error,
        )

        errors.append("警報注意報")

        previous_weather = previous_status.get(
            "weather",
            {},
        )

        wind_active = (
            previous_weather
            .get("wind", {})
            .get("active", False)
        )

        dry_active = (
            previous_weather
            .get("dry", {})
            .get("active", False)
        )

    try:
        minimum_temperature = (
            get_minimum_temperature()
        )

        cold_active = (
            minimum_temperature is not None
            and minimum_temperature <= 3
        )

        cold_note = (
            f"予想最低気温 "
            f"{minimum_temperature:g}℃"
            if minimum_temperature is not None
            else "予想最低気温を確認中"
        )

    except Exception as error:
        print(
            "気温取得エラー:",
            error,
        )

        errors.append("気温予報")

        previous_weather = previous_status.get(
            "weather",
            {},
        )

        cold_active = (
            previous_weather
            .get("cold", {})
            .get("active", False)
        )

        cold_note = "予想最低気温を確認中"

    try:
        influenza_data = get_influenza_data(
            previous_status
        )

    except Exception as error:
        print(
            "インフルエンザ取得エラー:",
            error,
        )

   
