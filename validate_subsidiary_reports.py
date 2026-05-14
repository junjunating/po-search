# -*- coding: utf-8 -*-

from pathlib import Path
import re
import pandas as pd
from playwright.sync_api import sync_playwright

from main import (
    parse_report_html,
    clean_text,
    html_to_text_cell,
)


OUT_FILE = Path("output/subsidiary_validation_report.xlsx")


LOG_CANDIDATES = [
    Path("run_log_20260512_131112 (full scanned) .txt"),
    Path("run_log_.txt"),
    Path("logs/run_log_.txt"),
]


def find_log_file():
    """
    검증에 사용할 로그 파일을 찾는다.
    우선 지정된 후보를 찾고, 없으면 현재 폴더/logs 폴더의 최신 txt 파일을 사용한다.
    """
    for p in LOG_CANDIDATES:
        if p.exists():
            return p

    txt_files = []

    txt_files.extend(Path(".").glob("*.txt"))

    logs_dir = Path("logs")
    if logs_dir.exists():
        txt_files.extend(logs_dir.glob("*.txt"))

    if txt_files:
        return max(txt_files, key=lambda x: x.stat().st_mtime)

    raise FileNotFoundError("검증에 사용할 run_log txt 파일을 찾지 못했습니다.")


def extract_report_urls_from_log(log_file):
    """
    run_log에서 itemReportTerm.do URL이 남아 있는 기관만 추출한다.
    """
    print("사용 로그 파일:", log_file)

    text = log_file.read_text(encoding="utf-8", errors="ignore")

    pattern = re.compile(
        r"(?:Report parsed|No subsidiary data, skipped):\s*(.*?)\s*/.*?"
        r"url=(https://www\.alio\.go\.kr/item/itemReportTerm\.do\?[^\s]+)"
    )

    items = []
    seen = set()

    for m in pattern.finditer(text):
        agency = clean_text(m.group(1))
        url = m.group(2).strip()

        key = (agency, url)

        if key in seen:
            continue

        seen.add(key)
        items.append((agency, url))

    return items


def get_all_tables(html):
    return re.findall(r"(?is)<table\b[^>]*>.*?</table>", str(html or ""))


def table_rows(table_html):
    rows = []

    trs = re.findall(r"(?is)<tr\b[^>]*>.*?</tr>", str(table_html or ""))

    for tr in trs:
        cells_html = re.findall(r"(?is)<t[dh]\b[^>]*>.*?</t[dh]>", tr)

        cells = []
        for c in cells_html:
            cell_text = clean_text(html_to_text_cell(c))
            if cell_text:
                cells.append(cell_text)

        if cells:
            rows.append(cells)

    return rows


def parse_int(value):
    """
    '70', '1,234', '70개' 같은 값을 int로 변환한다.
    """
    s = clean_text(value)

    if not s:
        return None

    s = s.replace(",", "")
    m = re.search(r"\d+", s)

    if not m:
        return None

    try:
        return int(m.group(0))
    except Exception:
        return None


def parse_expected_subsidiary_count_from_summary(html):
    """
    '1) 타법인 투·출자 현황 총괄표'의 최신연도 자회사 수를 추출한다.

    개선 기준:
    - 제목 table이나 header 위치에 의존하지 않는다.
    - 모든 table row를 한 번만 순회한다.
    - 총괄표의 '자회사' row는 다음 구조라는 점을 이용한다.

      자회사 | 0 | 0 | 0 | 1 | 2

    - 연도별 표는 다음 구조라 제외된다.

      자회사 | 법인명 | 보유/미보유 | ...

    - 해당연도 상세표는 다음 구조라 제외된다.

      자회사 | 법인명 | 설립일자 | 주요사업 | ...

    반환값:
    - 총괄표 최신연도 자회사 수 int
    - 찾지 못하면 None
    """
    tables = get_all_tables(html)

    if not tables:
        return None

    def _parse_int_strict(value):
        s = clean_text(value)

        if not s:
            return None

        s = s.replace(",", "").strip()

        # '-', 공백, 비숫자 제외
        if s in {"-", "–", "—"}:
            return 0

        # 순수 숫자 또는 소수 형태만 허용
        # 총괄표 개수는 보통 정수지만, 혹시 '2.0'처럼 들어오면 int 처리
        if re.fullmatch(r"\d+(\.0+)?", s):
            try:
                return int(float(s))
            except Exception:
                return None

        return None

    def _is_summary_subsidiary_row(row):
        if len(row) < 3:
            return False

        first = clean_text(row[0])

        # 총괄표의 첫 셀은 정확히 자회사
        # 헤더의 '자회사 출자회사 재출자회사'는 제외됨
        if first != "자회사":
            return False

        joined = " ".join(clean_text(x) for x in row)

        # 연도별 표 제외
        if "보유" in row or "미보유" in row:
            return False

        if "보유" in joined or "미보유" in joined:
            return False

        # 상세표 제외:
        # 두 번째 셀이 회사명이면 숫자로 파싱되지 않음
        numeric_values = [_parse_int_strict(x) for x in row[1:]]
        numeric_count = sum(v is not None for v in numeric_values)

        # 총괄표는 자회사 | 연도별 숫자... 구조라 숫자가 2개 이상이어야 함
        if numeric_count < 2:
            return False

        # row[1:] 중 숫자가 아닌 값이 너무 많으면 총괄표가 아닐 가능성이 큼
        # 예: 상세표는 회사명/설립일자/주요사업 등이 많음
        ratio = numeric_count / max(len(row[1:]), 1)

        if ratio < 0.6:
            return False

        return True

    candidates = []

    for table_index, table_html in enumerate(tables):
        rows = table_rows(table_html)

        for row_index, row in enumerate(rows):
            if not _is_summary_subsidiary_row(row):
                continue

            nums = []

            for cell in row[1:]:
                v = _parse_int_strict(cell)
                if v is not None:
                    nums.append(v)

            if not nums:
                continue

            # 총괄표는 보통 2021, 2022, 2023, 2024, 2025 순서이므로
            # 가장 오른쪽 숫자를 최신연도 값으로 사용
            expected = nums[-1]

            candidates.append({
                "expected": expected,
                "table_index": table_index,
                "row_index": row_index,
                "numeric_count": len(nums),
                "row_len": len(row),
                "row": row,
            })

    if not candidates:
        return None

    # 후보가 여러 개면 숫자 셀이 많은 row를 우선.
    # 그래도 같으면 문서 앞쪽 row를 우선.
    # 총괄표는 보통 연도별/상세표보다 앞에 있다.
    candidates.sort(
        key=lambda x: (
            -x["numeric_count"],
            x["table_index"],
            x["row_index"],
        )
    )

    return candidates[0]["expected"]


def validate_one_report(page, agency, url):
    """
    단일 ALIO 보고서 URL을 열어 해당연도 자회사 추출 결과를 검증한다.

    검증 기준:
    - expected_count: 총괄표 최신연도 자회사 수
    - extracted_count: parse_report_html()로 추출한 자회사 수
    - OK: expected_count == extracted_count
    - MISMATCH: expected_count != extracted_count
    - NO_SUMMARY_COUNT: 총괄표 expected를 못 찾음
    - NO_SUBSIDIARY_ROWS: extracted_count == 0
    """
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2500)

    html = page.content()

    expected_count = parse_expected_subsidiary_count_from_summary(html)

    parsed_rows = parse_report_html(html, agency)
    companies = [clean_text(r.get("company_name", "")) for r in parsed_rows]
    companies = [c for c in companies if c]

    # 순서 유지 중복 제거
    seen = set()
    unique_companies = []

    for c in companies:
        key = re.sub(r"\s+", "", c).lower()
        if key not in seen:
            seen.add(key)
            unique_companies.append(c)

    extracted_count = len(unique_companies)

    if expected_count is None:
        if extracted_count > 0:
            status = "NO_SUMMARY_COUNT"
        else:
            status = "NO_SUBSIDIARY_ROWS"
    else:
        if expected_count == extracted_count:
            status = "OK"
        else:
            status = "MISMATCH"

    return {
        "agency": agency,
        "expected_count_from_summary": expected_count,
        "extracted_count": extracted_count,
        "status": status,
        "companies": " | ".join(unique_companies),
        "url": url,
        "html_length": len(html),
        "error": "",
    }


def main():
    print("자회사 검증 시작")

    log_file = find_log_file()
    items = extract_report_urls_from_log(log_file)

    print("검증 대상 URL 수:", len(items))
    print("※ 로그에 itemReportTerm.do URL이 남아 있는 기관만 검증합니다.")
    print("※ expected_count_from_summary는 총괄표 최신연도 자회사 수입니다.")
    print()

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for i, (agency, url) in enumerate(items, start=1):
            print(f"[{i}/{len(items)}] {agency}")

            try:
                result = validate_one_report(page, agency, url)
                results.append(result)

            except Exception as e:
                results.append({
                    "agency": agency,
                    "expected_count_from_summary": None,
                    "extracted_count": None,
                    "status": "ERROR",
                    "companies": "",
                    "url": url,
                    "html_length": None,
                    "error": str(e),
                })

        browser.close()

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(results)

    columns = [
        "agency",
        "expected_count_from_summary",
        "extracted_count",
        "status",
        "companies",
        "url",
        "html_length",
        "error",
    ]

    df = df[columns]

    df.to_excel(OUT_FILE, index=False, engine="openpyxl")

    print()
    print("저장 완료:", OUT_FILE)
    print()
    print("상태별 개수:")
    print(df["status"].value_counts(dropna=False))

    print()
    print("요약:")
    ok_count = (df["status"] == "OK").sum()
    mismatch_count = (df["status"] == "MISMATCH").sum()
    no_summary_count = (df["status"] == "NO_SUMMARY_COUNT").sum()
    no_rows_count = (df["status"] == "NO_SUBSIDIARY_ROWS").sum()
    error_count = (df["status"] == "ERROR").sum()

    total_companies = pd.to_numeric(
        df["extracted_count"],
        errors="coerce"
    ).fillna(0).sum()

    print(f"OK 기관 수: {ok_count}")
    print(f"MISMATCH 기관 수: {mismatch_count}")
    print(f"NO_SUMMARY_COUNT 기관 수: {no_summary_count}")
    print(f"NO_SUBSIDIARY_ROWS 기관 수: {no_rows_count}")
    print(f"ERROR 기관 수: {error_count}")
    print(f"추출 자회사 수 합계: {int(total_companies)}")

    if mismatch_count > 0:
        print()
        print("MISMATCH 기관:")
        mismatch_df = df[df["status"] == "MISMATCH"].copy()
        print(
            mismatch_df[
                ["agency", "expected_count_from_summary", "extracted_count", "url"]
            ].to_string(index=False)
        )

    if no_summary_count > 0:
        print()
        print("NO_SUMMARY_COUNT 기관:")
        no_summary_df = df[df["status"] == "NO_SUMMARY_COUNT"].copy()
        print(
            no_summary_df[
                ["agency", "expected_count_from_summary", "extracted_count", "url"]
            ].to_string(index=False)
        )

    if error_count > 0:
        print()
        print("ERROR 기관:")
        error_df = df[df["status"] == "ERROR"]
        print(error_df[["agency", "url", "error"]].to_string(index=False))

    print()
    print("검증 완료")


if __name__ == "__main__":
    main()