# -*- coding: utf-8 -*-
"""
update.py

- 공공기관 CSV 다운로드
- 공직유관단체 다운로드
- 언론사 다운로드
- 공공기관 자회사 다운로드
- 4개 결과를 merged_list.csv로 통합
- last_update.txt 생성
"""

from pathlib import Path
import json
import shutil
from datetime import datetime

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
DOWNLOADS = BASE / "downloads"
META = BASE / "update_meta.json"

from main import (
    download_affiliated_org,
    download_media_company,
    collect_public_agency_subsidiaries,
)


# --------------------------------------------------
# 공통
# --------------------------------------------------

def load_meta():
    if META.exists():
        return json.loads(META.read_text(encoding="utf-8"))
    return {}


def save_meta(meta):
    meta["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    META.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# --------------------------------------------------
# 공공기관
# --------------------------------------------------

def update_public_agency():
    target_dir = DOWNLOADS / "public_agency"
    shutil.rmtree(target_dir, ignore_errors=True)
    target_dir.mkdir(parents=True, exist_ok=True)

    url = "https://www.alio.go.kr/guide/publicAgencyStatus.do"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto(url, timeout=30000)
        page.wait_for_load_state("domcontentloaded")

        download = None

        candidates = [
            "text=CSV",
            "text=csv",
            "text=CSV 다운로드",
            "text=CSV 파일",
        ]

        for selector in candidates:
            try:
                with page.expect_download(timeout=10000) as d:
                    page.click(selector)
                download = d.value
                break
            except Exception:
                continue

        if download is None:
            browser.close()
            raise RuntimeError("공공기관 CSV 다운로드 버튼을 찾지 못했습니다.")

        save_path = target_dir / download.suggested_filename
        download.save_as(save_path)

        browser.close()

    print("공공기관 업데이트 완료")


# --------------------------------------------------
# 공직유관단체
# --------------------------------------------------

def update_affiliated():
    folder = DOWNLOADS / "affiliated_org"
    shutil.rmtree(folder, ignore_errors=True)
    folder.mkdir(parents=True, exist_ok=True)

    download_affiliated_org()
    print("공직유관단체 업데이트 완료")


# --------------------------------------------------
# 언론사
# --------------------------------------------------

def update_media():
    """
    언론사 업데이트.

    정책:
    - 다운로드 시작 전에 기존 downloads/media_company 폴더를 삭제하지 않는다.
    - 다운로드 실패 시 기존 MainServiceExcel.xls를 유지한다.
    - 다운로드 성공 시 새 파일을 MainServiceExcel.xls로 표준화한다.
    - 성공 후 media_company 폴더에는 최신 MainServiceExcel.xls 하나만 남긴다.
    """
    folder = DOWNLOADS / "media_company"
    folder.mkdir(parents=True, exist_ok=True)

    canonical = folder / "MainServiceExcel.xls"

    before_files = set()
    before_files.update(folder.glob("*.xls"))
    before_files.update(folder.glob("*.xlsx"))
    before_files.update(folder.glob("*.csv"))

    result = download_media_company()

    after_files = set()
    after_files.update(folder.glob("*.xls"))
    after_files.update(folder.glob("*.xlsx"))
    after_files.update(folder.glob("*.csv"))

    status = str(result.get("status", "")).lower() if result else ""
    downloaded_file = str(result.get("downloaded_file", "") or "") if result else ""

    if status == "success" and downloaded_file:
        downloaded_path = Path(downloaded_file)

        if not downloaded_path.exists():
            raise RuntimeError(f"언론사 다운로드 성공으로 기록됐지만 파일이 없습니다: {downloaded_path}")

        # 새 파일 확장자가 xlsx인 경우도 대비
        if downloaded_path.suffix.lower() == ".xlsx":
            canonical = folder / "MainServiceExcel.xlsx"
        else:
            canonical = folder / "MainServiceExcel.xls"

        # 기존 표준 파일이 있고, 새 파일과 다른 파일이면 삭제
        if canonical.exists() and canonical.resolve() != downloaded_path.resolve():
            canonical.unlink()

        # 새 다운로드 파일을 표준 파일명으로 변경
        if downloaded_path.resolve() != canonical.resolve():
            downloaded_path.rename(canonical)

        # 성공 후 같은 폴더의 다른 언론사 파일 정리
        # 예: MainServiceExcel_114647.xls 제거
        for p in list(folder.glob("*.xls")) + list(folder.glob("*.xlsx")) + list(folder.glob("*.csv")):
            if p.resolve() != canonical.resolve():
                try:
                    p.unlink()
                except Exception as e:
                    print(f"⚠️ 언론사 중복 파일 삭제 실패: {p} / {e}")

        print(f"언론사 업데이트 완료: {canonical}")
        return

    # 다운로드 실패했지만 기존 파일이 있으면 유지
    existing_files = []
    existing_files.extend(folder.glob("*.xls"))
    existing_files.extend(folder.glob("*.xlsx"))
    existing_files.extend(folder.glob("*.csv"))

    if existing_files:
        print("⚠️ 언론사 다운로드 실패. 기존 언론사 파일을 유지합니다.")
        print("   유지 파일:")
        for p in sorted(existing_files):
            print(f"   - {p}")
        return

    # 기존 파일도 없고 새 다운로드도 실패한 경우
    raise RuntimeError(
        "언론사 다운로드 실패 및 기존 언론사 파일 없음. "
        "downloads/media_company 폴더에 MainServiceExcel.xls를 복구한 뒤 다시 실행하세요."
    )


# --------------------------------------------------
# 공공기관 자회사
# --------------------------------------------------

def update_subsidiary():
    collect_public_agency_subsidiaries()
    print("자회사 업데이트 완료")


# --------------------------------------------------
# 통합 리스트 생성
# --------------------------------------------------

def build_merged_list():
    import pandas as pd
    import re

    base_downloads = DOWNLOADS
    base_output = BASE / "output"

    merged = {}

    # 언론사는 제일 뒤
    category_order = {
        "공공기관": 1,
        "공직유관단체": 2,
        "공공기관 자회사": 3,
        "언론사": 4,
    }

    category_counts = {
        "공공기관": 0,
        "공직유관단체": 0,
        "언론사": 0,
        "공공기관 자회사": 0,
    }

    def clean_name(value):
        s = str(value or "").strip()
        s = re.sub(r"\s+", " ", s)
        if not s or s.lower() == "nan":
            return ""
        return s

    def norm_name(value):
        return re.sub(r"\s+", "", str(value or "")).lower()

    def add(name, category):
        name = clean_name(name)
        if not name:
            return

        key = norm_name(name)
        if not key:
            return

        if key not in merged:
            merged[key] = {
                "display_name": name,
                "categories": set(),
            }

        if category not in merged[key]["categories"]:
            merged[key]["categories"].add(category)
            category_counts[category] += 1

    def read_csv_flexible(path):
        for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                pass
        return None

    def read_excel_flexible(path):
        suffix = path.suffix.lower()

        if suffix == ".xls":
            try:
                return pd.read_excel(path, sheet_name=None, engine="xlrd")
            except Exception as e:
                print(f"⚠️ xls 읽기 실패: {path} / {e}")
                return {}

        if suffix == ".xlsx":
            try:
                return pd.read_excel(path, sheet_name=None, engine="openpyxl")
            except Exception as e:
                print(f"⚠️ xlsx 읽기 실패(openpyxl): {path} / {e}")
                try:
                    return pd.read_excel(path, sheet_name=None)
                except Exception as e2:
                    print(f"⚠️ xlsx 읽기 최종 실패: {path} / {e2}")
                    return {}

        return {}

    def pick_name_column(df, candidates):
        cols = [str(c).strip() for c in df.columns]
        df = df.copy()
        df.columns = cols

        for c in candidates:
            if c in df.columns:
                return c

        keyword_candidates = [
            "언론사", "매체", "제호", "신문", "방송",
            "회사", "법인", "기관", "단체", "상호", "명칭", "이름"
        ]

        for col in df.columns:
            col_text = str(col).strip()
            if any(k in col_text for k in keyword_candidates):
                return col

        best_col = None
        best_score = -1

        for c in df.columns:
            try:
                series = df[c].dropna().astype(str).str.strip()
                series = series[series != ""]

                if len(series) == 0:
                    continue

                sample = series.head(200)

                numeric_like = sample.str.fullmatch(r"[0-9]+").mean()
                short_like = sample.str.len().le(2).mean()
                text_like = sample.str.contains(r"[가-힣A-Za-z]", regex=True).mean()
                avg_len = sample.str.len().mean()

                score = 0
                score += len(series)
                score += text_like * 1000
                score += min(avg_len, 30) * 20
                score -= numeric_like * 3000
                score -= short_like * 1000

                if score > best_score:
                    best_score = score
                    best_col = c

            except Exception:
                pass

        return best_col

    def load_from_files(folder, category, candidate_columns):
        if not folder.exists():
            print(f"⚠️ {category} 폴더 없음: {folder}")
            return

        files = []
        files.extend(folder.glob("*.csv"))
        files.extend(folder.glob("*.xlsx"))
        files.extend(folder.glob("*.xls"))

        if not files:
            print(f"⚠️ {category} 파일 없음: {folder}")
            return

        before = category_counts[category]

        for path in files:
            suffix = path.suffix.lower()

            if suffix == ".csv":
                df = read_csv_flexible(path)
                if df is None or df.empty:
                    continue

                col = pick_name_column(df, candidate_columns)
                if not col:
                    continue

                for v in df[col].dropna():
                    add(v, category)

            elif suffix in [".xlsx", ".xls"]:
                sheets = read_excel_flexible(path)

                for _, df in sheets.items():
                    if df is None or df.empty:
                        continue

                    col = pick_name_column(df, candidate_columns)
                    if not col:
                        continue

                    for v in df[col].dropna():
                        add(v, category)

        after = category_counts[category]
        print(f"✅ {category} 반영: {after - before}개")

    def load_affiliated_from_column_c():
        """
        공직유관단체 파일 처리 규칙:
        - C열(세 번째 컬럼)에 실제 기관·단체명이 있음
        - 숫자, 헤더, 고시문구는 제외
        """
        folder = base_downloads / "affiliated_org"
        category = "공직유관단체"

        if not folder.exists():
            print(f"⚠️ {category} 폴더 없음: {folder}")
            return

        files = []
        files.extend(folder.glob("*.csv"))
        files.extend(folder.glob("*.xlsx"))
        files.extend(folder.glob("*.xls"))

        if not files:
            print(f"⚠️ {category} 파일 없음: {folder}")
            return

        before = category_counts[category]

        skip_values = {
            "기관·단체명",
            "기관ㆍ단체명",
            "기관단체명",
            "공직유관단체명",
            "인사혁신처장",
            "구분",
        }

        for path in files:
            suffix = path.suffix.lower()

            if suffix == ".csv":
                df = read_csv_flexible(path)
                if df is None or df.empty:
                    continue

                if df.shape[1] < 3:
                    print(f"⚠️ 공직유관단체 CSV C열 없음: {path}")
                    continue

                values = df.iloc[:, 2].dropna()

                for v in values:
                    name = clean_name(v)

                    if not name:
                        continue
                    if name.lower() == "nan":
                        continue
                    if name in skip_values:
                        continue
                    if name.isdigit():
                        continue

                    add(name, category)

            elif suffix in [".xlsx", ".xls"]:
                sheets = read_excel_flexible(path)

                for _, df in sheets.items():
                    if df is None or df.empty:
                        continue

                    if df.shape[1] < 3:
                        print(f"⚠️ 공직유관단체 Excel C열 없음: {path}")
                        continue

                    values = df.iloc[:, 2].dropna()

                    for v in values:
                        name = clean_name(v)

                        if not name:
                            continue
                        if name.lower() == "nan":
                            continue
                        if name in skip_values:
                            continue
                        if name.isdigit():
                            continue

                        add(name, category)

        after = category_counts[category]
        print(f"✅ 공직유관단체 반영(C열 기준): {after - before}개")

    def load_media_from_column_c():
        """
        언론사 파일 처리 규칙:
        - C열(세 번째 컬럼): 제호
        - K열(열한 번째 컬럼): 법인명
        - 표시 형식: 제호 (법인명)
        - 제호와 법인명이 모두 같은 경우만 중복 제거
        """
        folder = base_downloads / "media_company"
        category = "언론사"

        if not folder.exists():
            print(f"⚠️ {category} 폴더 없음: {folder}")
            return

        files = []
        files.extend(folder.glob("*.csv"))
        files.extend(folder.glob("*.xlsx"))
        files.extend(folder.glob("*.xls"))

        if not files:
            print(f"⚠️ {category} 파일 없음: {folder}")
            return

        before = category_counts[category]

        for path in files:
            suffix = path.suffix.lower()

            if suffix == ".csv":
                df = read_csv_flexible(path)
                if df is None or df.empty:
                    continue

                if df.shape[1] < 11:
                    print(f"⚠️ 언론사 CSV C/K열 없음: {path}")
                    continue

                title_values = df.iloc[:, 2]   # C열: 제호
                corp_values = df.iloc[:, 10]   # K열: 법인명

                for title, corp in zip(title_values, corp_values):
                    title = clean_name(title)
                    corp = clean_name(corp)

                    if not title:
                        continue
                    if title.lower() == "nan":
                        continue
                    if title in ["언론사명", "매체명", "회사명", "법인명", "제호"]:
                        continue

                    if corp and corp.lower() != "nan":
                        display_name = f"{title} ({corp})"
                    else:
                        display_name = title

                    add(display_name, category)

            elif suffix in [".xlsx", ".xls"]:
                sheets = read_excel_flexible(path)

                for _, df in sheets.items():
                    if df is None or df.empty:
                        continue

                    if df.shape[1] < 11:
                        print(f"⚠️ 언론사 Excel C/K열 없음: {path}")
                        continue

                    title_values = df.iloc[:, 2]   # C열: 제호
                    corp_values = df.iloc[:, 10]   # K열: 법인명

                    for title, corp in zip(title_values, corp_values):
                        title = clean_name(title)
                        corp = clean_name(corp)

                        if not title:
                            continue
                        if title.lower() == "nan":
                            continue
                        if title in ["언론사명", "매체명", "회사명", "법인명", "제호"]:
                            continue

                        if corp and corp.lower() != "nan":
                            display_name = f"{title} ({corp})"
                        else:
                            display_name = title

                        add(display_name, category)

        after = category_counts[category]
        print(f"✅ 언론사 반영(C열 제호 + K열 법인명 기준): {after - before}개")


    # 1. 공공기관
    load_from_files(
        base_downloads / "public_agency",
        "공공기관",
        ["기관명", "법인명", "기관", "공공기관명", "기관명(한글)"]
    )

    # 2. 공직유관단체
    # 공직유관단체 파일은 C열에 기관·단체명이 있으므로 C열 고정 사용
    load_affiliated_from_column_c()

    # 3. 언론사
    load_media_from_column_c()

    # 4. 공공기관 자회사
    subsidiary_sources = [
        base_output / "public_agency_subsidiaries.xlsx",
        base_output / "public_agency_subsidiaries_raw_fast.xlsx",
        base_downloads / "public_agency_subsidiaries.xlsx",
        base_downloads / "subsidiary" / "subsidiaries.csv",
    ]

    before = category_counts["공공기관 자회사"]

    for path in subsidiary_sources:
        if not path.exists():
            continue

        if path.suffix.lower() == ".csv":
            df = read_csv_flexible(path)
            if df is None or df.empty:
                continue

            col = pick_name_column(df, ["법인명", "company_name", "회사명", "name"])
            if col:
                for v in df[col].dropna():
                    add(v, "공공기관 자회사")

        else:
            sheets = read_excel_flexible(path)

            for _, df in sheets.items():
                if df is None or df.empty:
                    continue

                col = pick_name_column(df, ["법인명", "company_name", "회사명", "name"])
                if col:
                    for v in df[col].dropna():
                        add(v, "공공기관 자회사")

    after = category_counts["공공기관 자회사"]
    print(f"✅ 공공기관 자회사 반영: {after - before}개")

    final_rows = []

    for _, item in merged.items():
        cats = sorted(
            item["categories"],
            key=lambda c: category_order.get(c, 99)
        )

        category_text = " / ".join(cats)

        final_rows.append({
            "name": item["display_name"],
            "category": category_text,
        })

    def sort_key(row):
        first_category = row["category"].split(" / ")[0]
        return (
            category_order.get(first_category, 99),
            row["name"]
        )

    final_rows.sort(key=sort_key)

    numbered_rows = []
    for i, row in enumerate(final_rows, start=1):
        numbered_rows.append({
            "no": i,
            "name": row["name"],
            "category": row["category"],
        })

    out = base_downloads / "merged_list.csv"
    pd.DataFrame(numbered_rows).to_csv(out, index=False, encoding="utf-8-sig")

    last_update = datetime.now().strftime("%Y-%m-%d %H:%M")
    (base_downloads / "last_update.txt").write_text(last_update, encoding="utf-8")

    multi_category_count = sum(
        1 for item in merged.values()
        if len(item["categories"]) >= 2
    )

    print("=" * 60)
    print("✅ 통합 리스트 생성 완료")
    print(f"- 공공기관: {category_counts['공공기관']}개")
    print(f"- 공직유관단체: {category_counts['공직유관단체']}개")
    print(f"- 언론사: {category_counts['언론사']}개")
    print(f"- 공공기관 자회사: {category_counts['공공기관 자회사']}개")
    print(f"- 복수 분류 법인: {multi_category_count}개")
    print(f"- 최종 표시 행 수: {len(numbered_rows)}개")
    print(f"✅ 최근 업데이트 시간 저장: {last_update}")
    print("=" * 60)


# --------------------------------------------------
# main
# --------------------------------------------------

def main():
    print("✅ 업데이트 시작")
    meta = load_meta()

    update_public_agency()
    update_affiliated()
    update_media()
    update_subsidiary()

    build_merged_list()

    save_meta(meta)
    print("✅ 전체 업데이트 완료")


if __name__ == "__main__":
    main()
