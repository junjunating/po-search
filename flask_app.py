# -*- coding: utf-8 -*-

from flask import Flask, request, jsonify, render_template_string, send_file
from pathlib import Path
import csv
import math
from io import BytesIO

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DOWNLOADS = BASE_DIR / "downloads"
MERGED_FILE = DOWNLOADS / "merged_list.csv"
LAST_UPDATE_FILE = DOWNLOADS / "last_update.txt"

PER_PAGE = 1000


def load_last_update():
    if LAST_UPDATE_FILE.exists():
        return LAST_UPDATE_FILE.read_text(encoding="utf-8").strip()
    return "업데이트 정보 없음"


def load_rows():
    rows = []
    if not MERGED_FILE.exists():
        return rows
    with open(MERGED_FILE, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append({
                "no": row.get("no", ""),
                "name": row.get("name", ""),
                "category": row.get("category", ""),
            })
    return rows


TEMPLATE = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>PO 통합 리스트</title>
    <link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@400;600;700&family=Noto+Sans+KR:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: "Nunito Sans", "Noto Sans KR", "Malgun Gothic", "맑은 고딕", Arial, sans-serif;
            margin: 30px;
            background-color: #white;
            color: #222;
        }

        h1,
        input,
        button,
        table,
        th,
        td {
            font-family: "Nunito Sans", "Noto Sans KR", "Malgun Gothic", "맑은 고딕", Arial, sans-serif;
        }

        h1 {
            margin-bottom: 5px;
            color: #ff6200;
            font-weight: 700;
        }

        .meta {
            margin-bottom: 10px;
            color: #555;
            font-size: 14px;
        }

        .summary {
            margin-bottom: 15px;
            font-weight: 600;
            color: #333;
        }

        .search-box {
            background: white;
            padding: 15px;
            border: 1px solid #ffd1b3;
            border-radius: 8px;
            margin-bottom: 20px;
        }

        input {
            width: 420px;
            padding: 10px;
            font-size: 15px;
            border: 2px solid #ff7a26;
            border-radius: 6px;
            outline: none;
        }

        input:focus {
            border-color: #ff6200;
            box-shadow: 0 0 0 3px rgba(255, 98, 0, 0.15);
        }

        table {
            border-collapse: collapse;
            width: 100%;
            background-color: white;
            border: 1px solid #f0d2bf;
        }

        th,
        td {
            border: 1px solid #ead8cc;
            padding: 8px;
            text-align: left;
            font-size: 14px;
        }

        th {
            background-color: #ff6200;
            color: white;
            font-weight: 700;
        }

        tr:nth-child(even) {
            background-color: #fff7f0;
        }

        tr:hover {
            background-color: #ffe8d8;
        }

        .category {
            white-space: nowrap;
        }

        .notice {
            margin: 10px 0;
            color: #b00020;
            font-size: 14px;
        }

        .pagination {
            margin: 20px 0;
        }

        .pagination button {
            margin-right: 5px;
            margin-bottom: 5px;
            padding: 6px 10px;
            border: 1px solid #ff7a26;
            background: white;
            color: #ff6200;
            border-radius: 5px;
            cursor: pointer;
        }

        .pagination button.current {
            background: #ff6200;
            color: white;
            font-weight: 700;
        }

        .pagination button:disabled {
            color: #aaa;
            border-color: #ddd;
            cursor: not-allowed;
            background: #eee;
        }

        .table-tools {
            text-align: right;
            margin-bottom: 6px;
        }

        .download-link {
            font-size: 12px;
            color: #ff6200;
            text-decoration: none;
            font-weight: 600;
        }

        .download-link:hover {
            text-decoration: underline;
        }


        .footer {
            margin-top: 12px;
            margin-bottom: 12px;
	    padding-left: 18px;
            text-align: left;
            color: #999;
            font-size: 12px;
            font-weight: 400;
        }


    </style>
</head>
<body>
    <h1>PO 통합 리스트</h1>

    <div class="meta">최근 업데이트: {{ last_update }}</div>

    <div class="summary">
        전체 {{ "{:,}".format(total_count) }}개
    </div>

    <div class="search-box">
        <input
            type="text"
            id="searchInput"
            placeholder="법인명 검색"
            autocomplete="off"
        >
    </div>

    <div id="notice" class="notice"></div>

    <div id="paginationTop" class="pagination"></div>

    <div class="table-tools">
        <a href="/download" class="download-link">Excel 다운로드 </a>
    </div>

    <table>
        <thead>
            <tr>
                <th style="width:80px;">번호</th>
                <th>법인명</th>
                <th style="width:300px;">분류</th>
            </tr>
        </thead>
        <tbody id="listBody">
            {% for row in initial_rows %}
            <tr>
                <td>{{ row.no }}</td>
                <td>{{ row.name }}</td>
                <td class="category">{{ row.category }}</td>
            </tr>
            {% endfor %}
        </tbody>
    </table>

    <div id="paginationBottom" class="pagination"></div>

    <div style="
        margin-top: 40px;
        padding: 16px 18px;
        border-top: 1px solid #e5e7eb;
        color: #666;
        font-size: 13px;
        line-height: 1.6;
        background: #fafafa;
        border-radius: 8px;
    ">
        ※ 본 검색기는 ALIO, 인사혁신처 및 문화체육관광부 정기간행물 등록관리 시스템의 공개 데이터를 자동 수집·정리한 참고용 도구입니다.
        원천 사이트의 게시 오류, 갱신 지연, 형식 변경, 중복 등록, 기관별 공시 오류 등으로 실제 현황과 차이가 발생할 수 있습니다.
        대외 제출, 계약, 제재, 평가, 법적 판단 등 공식·확정적 용도로 사용하는 경우에는 원천 사이트의 최신 자료를 확인하시기 바랍니다.
    </div>

    <div class="footer">
        제작: 정준혁
    </div>

    <script>
        const searchInput = document.getElementById("searchInput");
        const listBody = document.getElementById("listBody");
        const notice = document.getElementById("notice");
        const paginationTop = document.getElementById("paginationTop");
        const paginationBottom = document.getElementById("paginationBottom");

        let timer = null;

        function escapeHtml(value) {
            return String(value)
                .replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll('"', "&quot;")
                .replaceAll("'", "&#039;");
        }

        function renderRows(data) {
            listBody.innerHTML = "";

            if (!data.rows || data.rows.length === 0) {
                listBody.innerHTML = `
                    <tr>
                        <td colspan="3">검색 결과가 없습니다.</td>
                    </tr>
                `;
                notice.textContent = "";
                renderPagination(1, 1);
                return;
            }

            const htmlRows = data.rows.map(row => `
                <tr>
                    <td>${escapeHtml(row.no)}</td>
                    <td>${escapeHtml(row.name)}</td>
                    <td class="category">${escapeHtml(row.category)}</td>
                </tr>
            `).join("");

            listBody.innerHTML = htmlRows;
            notice.textContent = "";
            renderPagination(data.page, data.total_pages);
        }

        function makeButton(label, page, disabled=false, isCurrent=false) {
            const btn = document.createElement("button");
            btn.textContent = label;

            if (disabled) {
                btn.disabled = true;
            }

            if (isCurrent) {
                btn.classList.add("current");
            }

            if (!disabled && !isCurrent) {
                btn.addEventListener("click", () => {
                    doSearch(page);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                });
            }

            return btn;
        }

        function renderPagination(page, totalPages) {
            paginationTop.innerHTML = "";
            paginationBottom.innerHTML = "";

            const containers = [paginationTop, paginationBottom];
            const PAGE_BUTTON_COUNT = 6;

            containers.forEach(container => {
                if (totalPages <= 1) {
                    return;
                }

                container.appendChild(makeButton("처음", 1, page === 1));
                container.appendChild(makeButton("이전", page - 1, page === 1));

                let start = Math.floor((page - 1) / PAGE_BUTTON_COUNT) * PAGE_BUTTON_COUNT + 1;
                let end = Math.min(start + PAGE_BUTTON_COUNT - 1, totalPages);

                if (start > 1) {
                    const span = document.createElement("span");
                    span.textContent = "...";
                    span.style.marginRight = "5px";
                    container.appendChild(span);
                }

                for (let p = start; p <= end; p++) {
                    container.appendChild(makeButton(String(p), p, false, p === page));
                }

                if (end < totalPages) {
                    const span = document.createElement("span");
                    span.textContent = "...";
                    span.style.marginRight = "5px";
                    container.appendChild(span);
                }

                container.appendChild(makeButton("다음", page + 1, page === totalPages));
                container.appendChild(makeButton("끝", totalPages, page === totalPages));
            });
        }

        function doSearch(page=1) {
            const q = searchInput.value.trim();

            fetch(`/api/search?q=${encodeURIComponent(q)}&page=${encodeURIComponent(page)}`)
                .then(response => response.json())
                .then(data => renderRows(data))
                .catch(() => {
                    notice.textContent = "검색 중 오류가 발생했습니다.";
                });
        }

        searchInput.addEventListener("input", () => {
            clearTimeout(timer);
            timer = setTimeout(() => {
                doSearch(1);
            }, 200);
        });

        renderPagination(1, {{ total_pages }});
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    rows = load_rows()
    last_update = load_last_update()
    total_count = len(rows)
    initial_rows = rows[:PER_PAGE]
    total_pages = max(1, math.ceil(total_count / PER_PAGE))

    return render_template_string(
        TEMPLATE,
        initial_rows=initial_rows,
        last_update=last_update,
        total_count=total_count,
        total_pages=total_pages,
    )

@app.route("/download")
def download_excel():
    rows = load_rows()

    output = BytesIO()

    try:
        import pandas as pd

        df = pd.DataFrame(rows, columns=["no", "name", "category"])
        df = df.rename(columns={
            "no": "번호",
            "name": "법인명",
            "category": "분류",
        })

        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="PO_통합리스트.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return f"Excel 다운로드 생성 중 오류가 발생했습니다: {e}", 500

@app.route("/api/search")
def api_search():
    q = request.args.get("q", "").strip().lower()

    try:
        page = int(request.args.get("page", "1"))
    except Exception:
        page = 1

    if page < 1:
        page = 1

    rows = load_rows()

    if q:
        filtered = [
            r for r in rows
            if q in str(r["name"]).lower()
        ]
    else:
        filtered = rows

    total = len(filtered)
    total_pages = max(1, math.ceil(total / PER_PAGE))

    if page > total_pages:
        page = total_pages

    start = (page - 1) * PER_PAGE
    end = start + PER_PAGE
    page_rows = filtered[start:end]

    return jsonify({
        "total": total,
        "page": page,
        "per_page": PER_PAGE,
        "total_pages": total_pages,
        "rows": page_rows,
    })


if __name__ == "__main__":
    app.run(debug=True)
