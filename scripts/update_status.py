#!/usr/bin/env python3
"""
京王笹塚 秋冬安全サイネージ用 status.json 生成スクリプト

取得情報
気象庁 VPWS50（集約定時通報）: 東京地方の強風注意報・暴風警報・乾燥注意報
気象庁 府県天気予報: 今夜～明朝の最低気温
東京都感染症情報センター: インフルエンザ定点当たり患者報告数

VPWS50を警報・注意報の主データとし、取得成功時はその結果を確定する。
旧 warning/130000.json は使用しない。
"""

import gzip
import io
import json
import re
import time
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta, timezone
from html import unescape
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import load_workbook


JST = timezone(timedelta(hours=9))

ROOT_DIR = Path(__file__).resolve().parents[1]
STATUS_FILE = ROOT_DIR / "data" / "status.json"

REGULAR_FEED_URL = (
    "https://www.data.jma.go.jp/developer/xml/feed/regular.xml"
)

FORECAST_URL = (
    "https://www.jma.go.jp/bosai/forecast/data/forecast/130000.json"
)

FLU_TOP_URL = (
    "https://idsc.tmiph.metro.tokyo.lg.jp/diseases/flu/flu/"
)


# VPWS50は全国集約電文。東京地方を対象とする。
VPWS50_TYPE = "VPWS50"
VPWS50_AREA = "010000"

TARGET_AREA_CODES = {
    "130010",
}

TARGET_AREA_NAMES = {
    "東京地方",
}


# 東京地点の気温コード
TOKYO_POINT_CODE = "44132"


DATA_URL_RE = re.compile(
    r"https://www.data.jma.go.jp/developer/xml/data/"
    r"(\d{14})_\d+_([A-Z0-9]+)_(\d{6})\.xml"
)


INACTIVE_STATUSES = {
    "解除",
    "なし",
    "発表警報・注意報はなし",
    "警報・注意報はなし",
}


def fetch_bytes(url, retries=3):
    """指定URLからバイトデータを取得する。"""

    last_error = None

    for attempt in range(1, retries + 1):
        try:
            separator = "&" if "?" in url else "?"

            cache_buster = int(
                datetime.now(timezone.utc).timestamp()
            )

            request = urllib.request.Request(
                f"{url}{separator}_={cache_buster}",
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 "
                        "Tokyo-Autumn-Winter-Safety-Signage/3.0"
                    ),
                    "Accept": (
                        "application/xml, "
                        "application/json, "
                        "text/html, */*"
                    ),
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=30,
            ) as response:
                raw = response.read()

                if raw[:2] == b"\x1f\x8b":
                    return gzip.decompress(raw)

                return raw

        except Exception as error:
            last_error = error

            print(
                f"通信失敗 {attempt}/{retries}: "
                f"{url}: {error}"
            )

            if attempt < retries:
                time.sleep(5 * attempt)

    raise last_error


def fetch_text(url):
    """指定URLからUTF-8テキストを取得する。"""

    return fetch_bytes(url).decode(
        "utf-8",
        errors="replace",
    )


def fetch_json(url):
    """指定URLからJSONデータを取得する。"""

    return json.loads(
        fetch_bytes(url).decode("utf-8")
    )


def html_to_text(html):
    """HTMLから表示文字列だけを取得する。"""

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

    return re.sub(
        r"\s+",
        " ",
        unescape(text),
    ).strip()


def load_previous_status():
    """前回のstatus.jsonを読み込む。"""

    if not STATUS_FILE.exists():
        return {
            "influenza": {},
            "weather": {},
        }

    try:
        value = json.loads(
            STATUS_FILE.read_text(
                encoding="utf-8"
            )
        )

        if isinstance(value, dict):
            return value

        return {
            "influenza": {},
            "weather": {},
        }

    except Exception as error:
        print(
            "前回データ読込エラー:",
            error,
        )

        return {
            "influenza": {},
            "weather": {},
        }


# ---------------------------------------------------------------------------
# VPWS50
# ---------------------------------------------------------------------------

def local_name(element):
    """XMLタグの名前空間を除去する。"""

    return element.tag.rsplit("}", 1)[-1]


def direct_text(element, name):
    """指定した直接の子要素から文字列を取得する。"""

    for child in element:
        if local_name(child) == name:
            return (child.text or "").strip()

    return ""


def select_latest_vpws50_url(feed_bytes):
    """regular.xmlから最新のVPWS50電文を選択する。"""

    text = feed_bytes.decode(
        "utf-8",
        errors="replace",
    )

    candidates = []

    for timestamp, message_type, area in DATA_URL_RE.findall(text):
        if (
            message_type == VPWS50_TYPE
            and area == VPWS50_AREA
        ):
            candidates.append(
                (
                    timestamp,
                    (
                        "https://www.data.jma.go.jp/"
                        "developer/xml/data/"
                        f"{timestamp}_0_{message_type}_{area}.xml"
                    ),
                )
            )

    if not candidates:
        return None

    return max(
        candidates,
        key=lambda item: item[0],
    )


def item_matches_tokyo(item):
    """XMLのItemが東京地方を対象にしているか判定する。"""

    for element in item.iter():
        if local_name(element) != "Area":
            continue

        area_name = direct_text(
            element,
            "Name",
        )

        area_code = direct_text(
            element,
            "Code",
        )

        if (
            area_code in TARGET_AREA_CODES
            or area_name in TARGET_AREA_NAMES
        ):
            return True

    return False


def parse_vpws50(xml_bytes):
    """
    東京地方の強風・暴風・乾燥を解析する。

    VPWS50の空Statusは有効として扱う。
    """

    root = ET.fromstring(xml_bytes)

    report_datetime = ""

    for element in root.iter():
        if (
            local_name(element) == "ReportDateTime"
            and element.text
        ):
            report_datetime = element.text.strip()
            break

    active_names = []
    wind_active = False
    dry_active = False

    items = (
        element
        for element in root.iter()
        if local_name(element) == "Item"
    )

    for item in items:
        if not item_matches_tokyo(item):
            continue

        kinds = (
            element
            for element in item.iter()
            if local_name(element) == "Kind"
        )

        for kind in kinds:
            name = direct_text(
                kind,
                "Name",
            )

            status = direct_text(
                kind,
                "Status",
            )

            # VPWS50は現在状況のスナップショットなので、
            # 空Statusも有効として扱う。
            if (
                not name
                or status in INACTIVE_STATUSES
            ):
                continue

            wind_names = (
                "強風注意報",
                "暴風警報",
                "暴風特別警報",
                "暴風雪警報",
                "暴風雪特別警報",
                "風雪注意報",
            )

            if any(
                word in name
                for word in wind_names
            ):
                wind_active = True
                active_names.append(name)

            if "乾燥注意報" in name:
                dry_active = True
                active_names.append(name)

    return {
        "windActive": wind_active,
        "dryActive": dry_active,
        "activeNames": list(
            dict.fromkeys(active_names)
        ),
        "reportDateTime": report_datetime,
    }


def get_vpws50_weather_alerts():
    """気象庁の最新VPWS50から東京地方の注意報を取得する。"""

    feed_bytes = fetch_bytes(
        REGULAR_FEED_URL
    )

    selected = select_latest_vpws50_url(
        feed_bytes
    )

    if not selected:
        raise ValueError(
            "regular.xmlからVPWS50を取得できません"
        )

    _, xml_url = selected

    result = parse_vpws50(
        fetch_bytes(xml_url)
    )

    result["sourceFile"] = (
        xml_url.rsplit("/", 1)[-1]
    )

    result["sourceType"] = VPWS50_TYPE
    result["emptyStatusAccepted"] = True
    result["resultLocked"] = True

    return result


# ---------------------------------------------------------------------------
# 最低気温
# ---------------------------------------------------------------------------

def parse_jma_datetime(value):
    """気象庁の日付文字列を日本時間に変換する。"""

    return datetime.fromisoformat(
        value
    ).astimezone(JST)


def get_overnight_minimum_temperature():
    """
    翌日0時の東京地点気温を優先する。

    翌日0時の気温が取得できない場合は、
    翌日の週間最低気温を使用する。
    """

    data = fetch_json(
        FORECAST_URL
    )

    target_date = (
        datetime.now(JST).date()
        + timedelta(days=1)
    )

    # 短期予報の翌日0時気温を優先
    for report in data:
        for series in report.get(
            "timeSeries",
            [],
        ):
            times = series.get(
                "timeDefines",
                [],
            )

            for area_data in series.get(
                "areas",
                [],
            ):
                area_code = str(
                    area_data.get(
                        "area",
                        {},
                    ).get(
                        "code",
                        "",
                    )
                )

                if area_code != TOKYO_POINT_CODE:
                    continue

                temperatures = area_data.get(
                    "temps",
                    [],
                )

                for time_text, value in zip(
                    times,
                    temperatures,
                ):
                    try:
                        forecast_time = (
                            parse_jma_datetime(
                                time_text
                            )
                        )

                        temperature = float(
                            value
                        )

                    except (TypeError, ValueError):
                        continue

                    if (
                        forecast_time.date()
                        == target_date
                        and forecast_time.hour == 0
                    ):
                        return (
                            temperature,
                            target_date,
                            "shortForecast00",
                        )

    # 翌日0時の値がない場合は週間最低気温を使用
    for report in data:
        for series in report.get(
            "timeSeries",
            [],
        ):
            times = series.get(
                "timeDefines",
                [],
            )

            for area_data in series.get(
                "areas",
                [],
            ):
                area = area_data.get(
                    "area",
                    {},
                )

                code = str(
                    area.get(
                        "code",
                        "",
                    )
                )

                name = str(
                    area.get(
                        "name",
                        "",
                    )
                )

                if (
                    code
                    not in {
                        TOKYO_POINT_CODE,
                        "130010",
                    }
                    and "東京" not in name
                ):
                    continue

                minimum_temperatures = (
                    area_data.get(
                        "tempsMin",
                        [],
                    )
                )

                for time_text, value in zip(
                    times,
                    minimum_temperatures,
                ):
                    try:
                        forecast_time = (
                            parse_jma_datetime(
                                time_text
                            )
                        )

                        temperature = float(
                            value
                        )

                    except (TypeError, ValueError):
                        continue

                    if (
                        forecast_time.date()
                        == target_date
                    ):
                        return (
                            temperature,
                            target_date,
                            "weeklyTempsMin",
                        )

    return (
        None,
        target_date,
        "notFound",
    )


# ---------------------------------------------------------------------------
# インフルエンザ
# ---------------------------------------------------------------------------

def determine_flu_level(value):
    """定点当たり報告数から表示レベルを決定する。"""

    if value is None:
        return "確認中"

    if value >= 30:
        return "警報レベル"

    if value >= 10:
        return "注意報レベル"

    if value >= 1:
        return "流行中"

    return "非流行"


def find_flu_excel_url(top_html):
    """東京都公式ページから患者報告数ExcelのURLを取得する。"""

    links = re.findall(
        r'<a\b[^"\']+["\'][^>]*>'
        r'([\s\S]*?)</a>',
        top_html,
        flags=re.IGNORECASE,
    )

    for href, label_html in links:
        url = urljoin(
            FLU_TOP_URL,
            unescape(href),
        )

        label = html_to_text(
            label_html
        )

        lower_url = url.lower()

        is_excel = (
            ".xlsx" in lower_url
            or ".xls" in lower_url
        )

        is_report_file = (
            "hasseidoko" in lower_url
            or "患者報告数" in label
        )

        if is_excel and is_report_file:
            return url

    raise ValueError(
        "東京都公式ページから患者報告数ExcelのURLを取得できません"
    )


def normalize(value):
    """セル値から空白を除去して文字列化する。"""

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        "",
        str(value),
    ).strip()


def number(value):
    """セル値を数値に変換する。"""

    if (
        isinstance(value, bool)
        or value is None
    ):
        return None

    if isinstance(
        value,
        (int, float),
    ):
        return float(value)

    text = (
        normalize(value)
        .replace(",", "")
        .replace("人", "")
    )

    if re.fullmatch(
        r"-?\d+(?:\.\d+)?",
        text,
    ):
        return float(text)

    return None


def parse_week_header(value):
    """週番号の見出しを解析する。"""

    match = re.fullmatch(
        r"(?:第)?(\d{1,2})(?:w|週)",
        normalize(value).lower(),
    )

    if not match:
        return None

    week = int(
        match.group(1)
    )

    if 1 <= week <= 53:
        return week

    return None


def find_latest_record(workbook):
    """Excelから最新週の合計・定点数を取得する。"""

    now = datetime.now(JST)
    candidates = []

    for sheet in workbook.worksheets:
        header_row = None
        week_columns = {}

        # 先頭30行から週番号の見出し行を探す
        for row in range(
            1,
            min(sheet.max_row, 30) + 1,
        ):
            found = {}

            for column in range(
                1,
                sheet.max_column + 1,
            ):
                week = parse_week_header(
                    sheet.cell(
                        row=row,
                        column=column,
                    ).value
                )

                if week is not None:
                    found[column] = week

            if len(found) > len(week_columns):
                header_row = row
                week_columns = found

        if (
            not header_row
            or not week_columns
        ):
            continue

        total_row = None
        sentinel_row = None

        # 合計行と定点数行を探す
        for row in range(
            header_row + 1,
            sheet.max_row + 1,
        ):
            row_labels = [
                normalize(
                    sheet.cell(
                        row=row,
                        column=column,
                    ).value
                ).lower()
                for column in range(
                    1,
                    min(
                        sheet.max_column,
                        8,
                    ) + 1,
                )
            ]

            if any(
                label in {
                    "total",
                    "合計",
                }
                for label in row_labels
            ):
                total_row = row

            if any(
                (
                    "定点数" in label
                    or "テイテンスウ" in label
                )
                for label in row_labels
            ):
                sentinel_row = row

        if (
            total_row is None
            or sentinel_row is None
        ):
            continue

        sheet_text = " ".join(
            normalize(
                sheet.cell(
                    row=row,
                    column=column,
                ).value
            )
            for row in range(
                1,
                min(
                    sheet.max_row,
                    8,
                ) + 1,
            )
            for column in range(
                1,
                min(
                    sheet.max_column,
                    12,
                ) + 1,
            )
        )

        season_match = re.search(
            r"(20\d{2})[-－]",
            sheet_text,
        )

        if season_match:
            season_start_year = int(
                season_match.group(1)
            )
        else:
            season_start_year = (
                now.year - 1
            )

        for column, week in week_columns.items():
            total = number(
                sheet.cell(
                    row=total_row,
                    column=column,
                ).value
            )

            points = number(
                sheet.cell(
                    row=sentinel_row,
                    column=column,
                ).value
            )

            if (
                total is None
                or points is None
                or points <= 0
            ):
                continue

            if week >= 36:
                year = season_start_year
            else:
                year = season_start_year + 1

            try:
                week_end = date.fromisocalendar(
                    year,
                    week,
                    7,
                )

            except ValueError:
                continue

            # 未来の週は採用しない
            if week_end > now.date():
                continue

            candidates.append(
                {
                    "year": year,
                    "week": week,
                    "value": round(
                        total / points,
                        2,
                    ),
                    "reportedCases": int(
                        total
                    ),
                    "sentinelCount": int(
                        points
                    ),
                    "sheet": sheet.title,
                }
            )

    if not candidates:
        raise ValueError(
            "Excelから合計行・定点数行・最新週を判定できません"
        )

    candidates.sort(
        key=lambda item: (
            item["year"],
            item["week"],
        )
    )

    latest = candidates[-1]

    print(
        "採用した最新週:",
        latest,
    )

    return latest


def iso_week_period(year, week):
    """ISO週番号から月曜日～日曜日の期間を生成する。"""

    monday = date.fromisocalendar(
        year,
        week,
        1,
    )

    sunday = date.fromisocalendar(
        year,
        week,
        7,
    )

    return (
        f"{monday.month}月{monday.day}日～"
        f"{sunday.month}月{sunday.day}日"
    )


def get_influenza_data(previous_status):
    """東京都公式Excelから最新のインフルエンザ情報を取得する。"""

    top_html = fetch_text(
        FLU_TOP_URL
    )

    excel_url = find_flu_excel_url(
        top_html
    )

    print(
        "インフルエンザExcel:",
        excel_url,
    )

    workbook = load_workbook(
        io.BytesIO(
            fetch_bytes(excel_url)
        ),
        data_only=True,
        read_only=True,
    )

    try:
        latest = find_latest_record(
            workbook
        )

    finally:
        workbook.close()

    previous_value = (
        previous_status
        .get("influenza", {})
        .get("perSentinel")
    )

    difference = None

    if previous_value is not None:
        try:
            difference = round(
                latest["value"]
                - float(previous_value),
                2,
            )

        except (TypeError, ValueError):
            pass

    return {
        "level": determine_flu_level(
            latest["value"]
        ),
        "week": (
            f"第{latest['week']}週"
        ),
        "period": iso_week_period(
            latest["year"],
            latest["week"],
        ),
        "perSentinel": latest["value"],
        "previousPerSentinel": previous_value,
        "difference": difference,
        "reportedCases": latest["reportedCases"],
        "sentinelCount": latest["sentinelCount"],
        "trend": "東京都公式週報",
        "provisional": True,
        "dataStatus": "取得成功",
        "sourceUrl": excel_url,
    }


def main():
    """各情報を取得し、status.jsonを生成する。"""

    previous = load_previous_status()

    previous_weather = previous.get(
        "weather",
        {},
    )

    errors = []

    now_text = datetime.now(
        JST
    ).strftime(
        "%Y-%m-%d %H:%M"
    )

    # -----------------------------------------------------------------------
    # 強風・乾燥
    # -----------------------------------------------------------------------

    try:
        warning_debug = (
            get_vpws50_weather_alerts()
        )

        wind_active = warning_debug[
            "windActive"
        ]

        dry_active = warning_debug[
            "dryActive"
        ]

    except Exception as error:
        print(
            "VPWS50取得エラー:",
            repr(error),
        )

        errors.append(
            "警報注意報"
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

        warning_debug = {
            "sourceType": VPWS50_TYPE,
            "error": type(error).__name__,
            "fallback": "previousStatus",
        }

    # -----------------------------------------------------------------------
    # 最低気温
    # -----------------------------------------------------------------------

    try:
        (
            minimum_temperature,
            minimum_date,
            temperature_source,
        ) = get_overnight_minimum_temperature()

        # 最低気温が5℃以下の場合に注意表示を有効化
        cold_active = (
            minimum_temperature is not None
            and minimum_temperature <= 5
        )

        if minimum_temperature is not None:
            cold_note = (
                "今夜～明朝 "
                f"{minimum_temperature:g}℃"
            )

        else:
            cold_note = (
                "今夜～明朝の最低気温を確認中"
            )

        cold_debug = {
            "targetDate": minimum_date.isoformat(),
            "minimumTemperature": minimum_temperature,
            "source": temperature_source,
        }

    except Exception as error:
        print(
            "気温取得エラー:",
            repr(error),
        )

        errors.append(
            "気温予報"
        )

        cold_active = (
            previous_weather
            .get("cold", {})
            .get("active", False)
        )

        cold_note = (
            "今夜～明朝の最低気温を確認中"
        )

        cold_debug = {
            "error": type(error).__name__,
            "fallback": "previousStatus",
        }

    # -----------------------------------------------------------------------
    # インフルエンザ
    # -----------------------------------------------------------------------

    try:
        influenza = get_influenza_data(
            previous
        )

    except Exception as error:
        print(
            "インフルエンザ取得エラー:",
            repr(error),
        )

        errors.append(
            "感染症情報"
        )

        old = previous.get(
            "influenza",
            {},
        )

        if old.get("perSentinel") is not None:
            influenza = {
                **old,
                "dataStatus": (
                    "取得失敗・前回値を維持"
                ),
            }

        else:
            influenza = {
                "level": "確認中",
                "week": "最新発表",
                "period": "東京都",
                "perSentinel": None,
                "previousPerSentinel": None,
                "difference": None,
                "reportedCases": None,
                "sentinelCount": None,
                "trend": "公式情報を確認中",
                "provisional": True,
                "dataStatus": "取得失敗",
                "sourceUrl": FLU_TOP_URL,
            }

    # -----------------------------------------------------------------------
    # status.json
    # -----------------------------------------------------------------------

    status = {
        "updated": now_text,
        "sourceStatus": (
            "正常"
            if not errors
            else (
                "一部取得失敗："
                + "・".join(errors)
            )
        ),
        "influenza": influenza,
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
                "alertText": "低温・凍結注意",
                "note": cold_note,
            },
        },
        "debug": {
            "warning": warning_debug,
            "cold": cold_debug,
        },
    }

    STATUS_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary_file = (
        STATUS_FILE.with_suffix(".tmp")
    )

    temporary_file.write_text(
        json.dumps(
            status,
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    # 作成した一時ファイルが正常なJSONか確認
    json.loads(
        temporary_file.read_text(
            encoding="utf-8"
        )
    )

    # 正常な場合のみstatus.jsonへ置き換える
    temporary_file.replace(
        STATUS_FILE
    )

    print(
        json.dumps(
            status,
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
