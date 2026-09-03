import json
import re
import urllib.request
from datetime import datetime, timedelta, timezone
from html import unescape
from pathlib import Path
from urllib.parse import urljoin


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
    "130010",   # 東京地方
    "1311300",  # 世田谷区として返る場合への予備
    "13113000", # 世田谷区として返る場合への予備
}


# =========================================================
# 共通処理
# =========================================================

def fetch_bytes(url):
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "Tokyo-Autumn-Winter-Safety-Signage/1.0"
            )
        },
    )

    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read()


def fetch_text(url):
    return fetch_bytes(url).decode("utf-8", errors="replace")


def fetch_json(url):
    return json.loads(fetch_bytes(url).decode("utf-8"))


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
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


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
            area_code = str(area.get("code", ""))

            if area_code not in TARGET_WARNING_AREA_CODES:
                continue

            for warning in area.get("warnings", []):
                status = str(warning.get("status", ""))
                name = str(warning.get("name", ""))

                if not name:
                    continue
                if "解除" in status or "なし" in status:
                    continue

                warning_names.append(name)

    return warning_names


def get_minimum_temperature():
    forecast_data = fetch_json(FORECAST_URL)
    temperatures = []

    for report in forecast_data:
        for time_series in report.get("timeSeries", []):
            for area_data in time_series.get("areas", []):
                area_code = str(area_data.get("area", {}).get("code", ""))

                if area_code != TOKYO_POINT_CODE:
                    continue

                for value in area_data.get("temps", []):
                    try:
                        temperatures.append(float(value))
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


def find_latest_flu_press_url(top_html):
    anchor_pattern = re.compile(
        r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>([\s\S]*?)</a>",
        flags=re.IGNORECASE,
    )

    candidates = []

    for href, label_html in anchor_pattern.findall(top_html):
        label = html_to_text(label_html)
        complete_url = urljoin(FLU_TOP_URL, href)

        is_press_release = "/information/press/" in complete_url
        mentions_flu = "インフルエンザ" in label

        if is_press_release and mentions_flu:
            candidates.append(complete_url)

    if not candidates:
        raise ValueError("最新のインフルエンザ報道発表URLを取得できません")

    return candidates[0]


def extract_flu_value(press_text):
    patterns = [
        (
            r"インフルエンザ(?:患者)?報告数"
            r"(?:は|が|：|:)?\s*"
            r"([0-9]+(?:\.[0-9]+)?)"
        ),
        (
            r"定点(?:医療機関)?(?:当たり|からの)"
            r".{0,100}?"
            r"(?:報告数)?(?:は|が|：|:)?\s*"
            r"([0-9]+(?:\.[0-9]+)?)"
        ),
    ]

    for pattern in patterns:
        match = re.search(pattern, press_text)
        if match:
            return float(match.group(1))

    return None


def extract_week_information(press_text):
    patterns = [
        (
            r"第\s*(\d+)\s*週\s*[（(]"
            r"\s*(\d+)\s*月\s*(\d+)\s*日\s*から\s*"
            r"(\d+)\s*月\s*(\d+)\s*日\s*まで\s*[）)]"
        ),
        (
            r"第\s*(\d+)\s*週.{0,25}?"
            r"(\d+)\s*月\s*(\d+)\s*日.{0,15}?"
            r"(\d+)\s*月\s*(\d+)\s*日"
        ),
    ]

    for pattern in patterns:
        match = re.search(pattern, press_text)
        if not match:
            continue

        week_number = int(match.group(1))
        start_month = int(match.group(2))
        start_day = int(match.group(3))
        end_month = int(match.group(4))
        end_day = int(match.group(5))

        return {
            "week": f"第{week_number}週",
            "period": f"{start_month}月{start_day}日～{end_month}月{end_day}日",
        }

    return {
        "week": "最新発表",
        "period": "東京都",
    }


def get_influenza_data(previous_status):
    top_html = fetch_text(FLU_TOP_URL)
    press_url = find_latest_flu_press_url(top_html)
    press_text = html_to_text(fetch_text(press_url))

    per_sentinel = extract_flu_value(press_text)
    week_information = extract_week_information(press_text)

    if per_sentinel is None:
        previous_influenza = previous_status.get("influenza", {})
        previous_value = previous_influenza.get("perSentinel")

        if previous_value is not None:
            return {
                **previous_influenza,
                "dataStatus": "今回の数値取得失敗・前回値を維持",
                "sourceUrl": press_url,
            }

        raise ValueError("定点当たり患者報告数を取得できません")

    return {
        "level": determine_flu_level(per_sentinel),
        "week": week_information["week"],
        "period": week_information["period"],
        "perSentinel": per_sentinel,
        "previousPerSentinel": None,
        "difference": None,
        "reportedCases": None,
        "trend": "東京都公式発表",
        "provisional": True,
        "dataStatus": "取得成功",
        "sourceUrl": press_url,
    }


# =========================================================
# status.json生成
# =========================================================

def main():
    previous_status = load_previous_status()
    errors = []
    now_text = datetime.now(JST).strftime("%Y-%m-%d %H:%M")

    try:
        warning_names = get_active_warning_names()
        warning_text = " ".join(warning_names)
        wind_active = any(
            keyword in warning_text
            for keyword in ["強風", "暴風", "風雪"]
        )
        dry_active = "乾燥" in warning_text
    except Exception as error:
        print("警報・注意報取得エラー:", error)
        errors.append("警報注意報")
        previous_weather = previous_status.get("weather", {})
        wind_active = previous_weather.get("wind", {}).get("active", False)
        dry_active = previous_weather.get("dry", {}).get("active", False)

    try:
        minimum_temperature = get_minimum_temperature()
        cold_active = (
            minimum_temperature is not None
            and minimum_temperature <= 3
        )
        cold_note = (
            f"予想最低気温 {minimum_temperature:g}℃"
            if minimum_temperature is not None
            else "予想最低気温を確認中"
        )
    except Exception as error:
        print("気温取得エラー:", error)
        errors.append("気温予報")
        previous_weather = previous_status.get("weather", {})
        cold_active = previous_weather.get("cold", {}).get("active", False)
        cold_note = "予想最低気温を確認中"

    try:
        influenza_data = get_influenza_data(previous_status)
    except Exception as error:
        print("インフルエンザ取得エラー:", error)
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
        "sourceStatus": (
            "正常"
            if not errors
            else "一部取得失敗：" + "・".join(errors)
        ),
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
    STATUS_FILE.write_text(
        json.dumps(status, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    print(json.dumps(status, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
