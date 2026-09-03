import json
import re
import urllib.request
from datetime import datetime, timezone, timedelta
from html import unescape
from pathlib import Path
from urllib.parse import urljoin


# 日本時間
JST = timezone(timedelta(hours=9))

# status.jsonの保存先
OUTPUT_FILE = (
    Path(__file__).resolve().parents[1]
    / "data"
    / "status.json"
)

# 気象庁
WARNING_URL = (
    "https://www.jma.go.jp/"
    "bosai/warning/data/warning/130000.json"
)

FORECAST_URL = (
    "https://www.jma.go.jp/"
    "bosai/forecast/data/forecast/130000.json"
)

# 東京都感染症情報センター
FLU_TOP_URL = (
    "https://idsc.tmiph.metro.tokyo.lg.jp/"
    "diseases/flu/flu/"
)

# 東京地点
TOKYO_POINT_CODE = "44132"


def fetch_bytes(url):
    """
    URLからデータを取得します。
    """

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "Tokyo-Autumn-Winter-Safety-Signage/1.0"
            )
        }
    )

    with urllib.request.urlopen(
        request,
        timeout=30
    ) as response:
        return response.read()


def fetch_json(url):
    """
    URLからJSONを取得します。
    """

    raw_data = fetch_bytes(url)

    return json.loads(
        raw_data.decode("utf-8")
    )


def fetch_text(url):
    """
    URLからHTMLを取得します。
    """

    raw_data = fetch_bytes(url)

    return raw_data.decode(
        "utf-8",
        errors="replace"
    )


def html_to_text(html):
    """
    HTMLタグなどを除去して、
    検索しやすい文字列に変換します。
    """

    text = re.sub(
        r"<script[\s\S]*?</script>",
        " ",
        html,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"<style[\s\S]*?</style>",
        " ",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"<[^>]+>",
        " ",
        text
    )

    text = unescape(text)

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def load_previous_status():
    """
    前回のstatus.jsonを読み込みます。

    外部データの取得に失敗した際、
    気象情報の前回値を維持するために使います。
    """

    if not OUTPUT_FILE.exists():
        return {
            "influenza": {},
            "weather": {}
        }

    try:
        return json.loads(
            OUTPUT_FILE.read_text(
                encoding="utf-8"
            )
        )

    except Exception:
        return {
            "influenza": {},
            "weather": {}
        }


def get_warning_text():
    """
    東京都の警報・注意報JSONを、
    検索用の文字列に変換します。
    """

    warning_data = fetch_json(
        WARNING_URL
    )

    return json.dumps(
        warning_data,
        ensure_ascii=False
    )


def get_minimum_temperature():
    """
    気象庁の東京都予報から、
    東京地点44132の最低気温を取得します。
    """

    forecast_data = fetch_json(
        FORECAST_URL
    )

    temperatures = []

    for report in forecast_data:

        time_series_list = report.get(
            "timeSeries",
            []
        )

        for time_series in time_series_list:

            areas = time_series.get(
                "areas",
                []
            )

            for area_data in areas:

                area = area_data.get(
                    "area",
                    {}
                )

                area_code = str(
                    area.get("code", "")
                )

                if area_code != TOKYO_POINT_CODE:
                    continue

                for temperature_text in area_data.get(
                    "temps",
                    []
                ):

                    try:
                        temperature = float(
                            temperature_text
                        )

                        temperatures.append(
                            temperature
                        )

                    except (
                        TypeError,
                        ValueError
                    ):
                        continue

    if not temperatures:
        return None

    return min(temperatures)


def get_latest_flu_press_url():
    """
    東京都感染症情報センターのページから、
    最新のインフルエンザ報道発表URLを取得します。

    過去の記事本文全体を判定に使用せず、
    一番上の報道発表リンクだけを取得します。
    """

    top_html = fetch_text(
        FLU_TOP_URL
    )

    link_pattern = re.compile(
        r'href=[^"\']+["\'][^>]*>'
        r'[^<]*(?:都内|東京都内)[^<]*'
        r'インフルエンザ',
        re.IGNORECASE
    )

    match = link_pattern.search(
        top_html
    )

    if match:
        return urljoin(
            FLU_TOP_URL,
            match.group(1)
        )

    # HTML構造が変わった場合の予備検索
    fallback_pattern = re.compile(
        r'href=[^"\']*'
        r'metro\.tokyo\.lg\.jp'
        r'[^"\']*["\']',
        re.IGNORECASE
    )

    fallback_match = fallback_pattern.search(
        top_html
    )

    if fallback_match:
        return urljoin(
            FLU_TOP_URL,
            fallback_match.group(1)
        )

    raise ValueError(
        "最新のインフルエンザ報道発表URLを取得できません"
    )


def determine_flu_level(per_sentinel):
    """
    定点当たり患者報告数によって、
    サイネージ上の表示レベルを判定します。
    """

    if per_sentinel is None:
        return "確認中"

    if per_sentinel >= 30:
        return "警報レベル"

    if per_sentinel >= 10:
        return "注意報レベル"

    if per_sentinel >= 1:
        return "流行中"

    return "非流行"


def get_influenza_data():
    """
    東京都の最新報道発表から、
    対象週、対象期間、定点当たり報告数を取得します。
    """

    press_url = get_latest_flu_press_url()

    press_html = fetch_text(
        press_url
    )

    press_text = html_to_text(
        press_html
    )

    # 例：
    # 第34週（8月17日から8月23日まで）
    period_match = re.search(
        r"第\s*(\d+)\s*週"
        r"（\s*"
        r"(\d+)\s*月\s*(\d+)\s*日"
        r"から"
        r"(\d+)\s*月\s*(\d+)\s*日"
        r"まで\s*）",
        press_text
    )

    # 例：
    # インフルエンザ患者報告数は1.49
    value_match = re.search(
        r"インフルエンザ"
        r"(?:患者)?報告数"
        r"は\s*"
        r"([0-9]+(?:\.[0-9]+)?)",
        press_text
    )

    if not value_match:
        # 表現が変わった場合の予備検索
        value_match = re.search(
            r"定点(?:医療機関)?"
            r"(?:当たり|からの)"
            r".{0,40}?"
            r"([0-9]+(?:\.[0-9]+)?)"
            r"\s*人",
            press_text
        )

    if not value_match:
        raise ValueError(
            "定点当たり患者報告数を取得できません"
        )

    per_sentinel = float(
        value_match.group(1)
    )

    if period_match:

        week_number = int(
            period_match.group(1)
        )

        start_month = int(
            period_match.group(2)
        )

        start_day = int(
            period_match.group(3)
        )

        end_month = int(
            period_match.group(4)
        )

        end_day = int(
            period_match.group(5)
        )

        week_text = (
            f"第{week_number}週"
        )

        period_text = (
            f"{start_month}月{start_day}日"
            f"～"
            f"{end_month}月{end_day}日"
        )

    else:
        week_text = "最新発表"
        
