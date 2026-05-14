from pathlib import Path
import json
import pandas as pd


ROOT = Path(__file__).resolve().parent

DOWNLOADS_DIR = ROOT / "downloads"
OUTPUT_DIR = ROOT / "output"
DOCS_DIR = ROOT / "docs"
DATA_DIR = DOCS_DIR / "data"
DOWNLOAD_DIR = DOCS_DIR / "download"

MERGED_CSV = DOWNLOADS_DIR / "merged_list.csv"
LAST_UPDATE_TXT = DOWNLOADS_DIR / "last_update.txt"


def read_last_update():
    if LAST_UPDATE_TXT.exists():
        return LAST_UPDATE_TXT.read_text(encoding="utf-8").strip()
    return ""


def main():
    if not MERGED_CSV.exists():
        raise FileNotFoundError(f"merged_list.csv 파일이 없습니다: {MERGED_CSV}")

    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(MERGED_CSV, encoding="utf-8-sig")
    df = df.fillna("")

    required_cols = ["no", "name", "category"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"merged_list.csv 필수 컬럼 누락: {missing}")

    # 웹 검색용 JSON 생성
    json_rows = (
        df[required_cols]
        .astype({"no": "str", "name": "str", "category": "str"})
        .to_dict(orient="records")
    )

    (DATA_DIR / "merged_list.json").write_text(
        json.dumps(json_rows, ensure_ascii=False),
        encoding="utf-8",
    )

    # 요약 정보 생성
    summary = {
        "last_update": read_last_update(),
        "total_count": int(len(df)),
    }

    (DATA_DIR / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # 다운로드용 파일 생성
    df.to_csv(DOWNLOAD_DIR / "merged_list.csv", index=False, encoding="utf-8-sig")
    df.to_excel(DOWNLOAD_DIR / "merged_list.xlsx", index=False, engine="openpyxl")

    print("정적 사이트 데이터 생성 완료")
    print(f"- JSON: {DATA_DIR / 'merged_list.json'}")
    print(f"- Summary: {DATA_DIR / 'summary.json'}")
    print(f"- CSV Download: {DOWNLOAD_DIR / 'merged_list.csv'}")
    print(f"- Excel Download: {DOWNLOAD_DIR / 'merged_list.xlsx'}")


if __name__ == "__main__":
    main()