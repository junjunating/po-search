# -*- coding: utf-8 -*-
"""
main.py
- 공공기관 / 공공기관 자회사 / 공직유관단체 / 언론사 수집
- 공공기관 자회사(ALIO)는 이제 '보고서 버튼 클릭 -> 새 창 URL/itemReportTerm.do' 기준으로 처리합니다.
- 전체 DOM/row-scope fallback 파싱은 제거했습니다. 같은 42개 반복 오탐 방지 목적입니다.
"""

import os
import re
import sys
import time
import json
import zipfile
import hashlib
import traceback
from io import StringIO
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parent
DOWNLOAD_DIR = BASE_DIR / "downloads"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"
DEBUG_DIR = BASE_DIR / "debug"
for d in [DOWNLOAD_DIR, OUTPUT_DIR, LOG_DIR, DEBUG_DIR]:
    d.mkdir(exist_ok=True)

RUN_TIME = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = LOG_DIR / f"run_log_{RUN_TIME}.txt"

PUBLIC_AGENCY_URL = "https://www.alio.go.kr/mobile/guide/publicAgencyStatus.do"
PUBLIC_SUBSIDIARY_ALIO_URL = "https://www.alio.go.kr/item/itemOrganList.do?reportFormRootNo=31901"
AFFILIATED_ORG_URL = "https://www.mpm.go.kr/mpm/lawStat/infoLaw/lawAnwei/lawAnwei05/"
MEDIA_COMPANY_URL = "https://pds.mcst.go.kr/pds/main/press/selectPressList.do"

TEXT_KEYWORDS = ["자회사", "타법인", "출자", "법인명", "회사명", "지분", "지분율", "출자비율", "출자금액", "투자회사", "투자 및 출자"]
NO_DATA_WORDS = ["보고된 내역이 없습니다", "자료가 없습니다", "해당사항 없음", "조회된 데이터가 없습니다"]


def log(message):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def safe_filename(name):
    name = unquote(str(name or "downloaded_file"))
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return name or "downloaded_file"


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def validate_file(path, expected_exts=None):
    p = Path(path)
    if not p.exists():
        return False, "file_not_found"
    if p.stat().st_size == 0:
        return False, "empty_file"
    if expected_exts and p.suffix.lower() not in [e.lower() for e in expected_exts]:
        return False, f"unexpected_extension: {p.suffix}"
    return True, "ok"


def make_result(task, url, status="failed", file_path="", message="", extracted_files=None):
    extracted_files = extracted_files or []
    size = ""
    digest = ""
    if file_path and Path(file_path).exists():
        size = Path(file_path).stat().st_size
        digest = file_sha256(file_path)
    return {
        "task": task,
        "source_url": url,
        "status": status,
        "downloaded_file": str(file_path) if file_path else "",
        "file_size": size,
        "sha256": digest,
        "extracted_files": json.dumps(extracted_files, ensure_ascii=False),
        "message": message,
        "downloaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def launch_browser(p):
    """
    로컬 Windows에서는 Microsoft Edge 우선 사용.
    GitHub Actions/Linux 서버에서는 headless Chromium 사용.
    """
    import os
    import platform

    is_github_actions = os.getenv("GITHUB_ACTIONS", "").lower() == "true"
    system_name = platform.system().lower()

    if is_github_actions or system_name != "windows":
        log("Trying headless Chromium browser for CI/Linux environment...")
        return p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
            ],
        )

    try:
        log("Trying Microsoft Edge browser...")
        return p.chromium.launch(channel="msedge", headless=False)
    except Exception as e:
        log(f"Microsoft Edge launch failed, trying bundled Chromium: {e}")
        return p.chromium.launch(headless=False)


def robust_goto(page, url, timeout=30000, sleep_after=2):
    try:
        page.goto(url, wait_until="commit", timeout=timeout)
        time.sleep(sleep_after)
        return True
    except Exception as e:
        log(f"Page goto commit failed. Retry domcontentloaded: {e}")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            time.sleep(sleep_after)
            return True
        except Exception as e2:
            log(f"Page goto failed: {e2}")
            return False


def save_download(download, folder, fallback_name=None):
    name = safe_filename(download.suggested_filename or fallback_name or "downloaded_file")
    path = folder / name
    if path.exists():
        path = folder / f"{path.stem}_{datetime.now().strftime('%H%M%S')}{path.suffix}"
    download.save_as(path)
    return path


def click_download(page, locator, folder, exts, timeout=12000):
    try:
        with page.expect_download(timeout=timeout) as info:
            locator.click(timeout=5000, no_wait_after=True)
        path = save_download(info.value, folder)
        ok, msg = validate_file(path, exts)
        if ok:
            return path, "ok"
        try:
            path.unlink()
        except Exception:
            pass
        return None, msg
    except PlaywrightTimeoutError:
        return None, "download_timeout"
    except Exception as e:
        return None, str(e)


def direct_download(url, folder, exts, referer=None, fallback=None):
    headers = {"User-Agent": "Mozilla/5.0"}
    if referer:
        headers["Referer"] = referer
    r = requests.get(url, timeout=30, headers=headers)
    r.raise_for_status()
    cd = r.headers.get("content-disposition", "")
    filename = ""
    m = re.search(r"filename\*=UTF-8''([^;]+)", cd, re.I)
    if m:
        filename = unquote(m.group(1))
    if not filename:
        m = re.search(r'filename="?([^";]+)"?', cd, re.I)
        if m:
            filename = m.group(1)
    if not filename:
        filename = fallback or os.path.basename(urlparse(url).path) or "downloaded_file"
    filename = safe_filename(filename)
    if "." not in filename and fallback and "." in fallback:
        filename = safe_filename(fallback)
    path = folder / filename
    if path.exists():
        path = folder / f"{path.stem}_{datetime.now().strftime('%H%M%S')}{path.suffix}"
    with open(path, "wb") as f:
        f.write(r.content)
    ok, msg = validate_file(path, exts)
    if ok:
        return path, "ok"
    try:
        path.unlink()
    except Exception:
        pass
    return None, msg


def extract_zip_if_needed(path, folder):
    p = Path(path)
    if p.suffix.lower() != ".zip":
        return []
    out_dir = folder / f"extracted_{p.stem}"
    out_dir.mkdir(exist_ok=True)
    try:
        with zipfile.ZipFile(p, "r") as z:
            z.extractall(out_dir)
        files = [str(x) for x in out_dir.rglob("*") if x.is_file()]
        log(f"ZIP extracted: {p.name} -> {out_dir}")
        return files
    except Exception as e:
        log(f"ZIP extraction failed: {e}")
        return []


# -----------------------------------------------------------------------------
# 공통 클릭 후보 수집: 공직유관단체/언론사 기존 방식 유지
# -----------------------------------------------------------------------------
def get_clickables(page):
    js = """
    () => Array.from(document.querySelectorAll('a,button,input[type="button"],input[type="submit"],[role="button"],[onclick]')).map((el,i)=>{
        const r=el.getBoundingClientRect();
        const id='crawler_'+i+'_'+Math.random().toString(36).slice(2);
        el.setAttribute('data-crawler-id', id);
        return {crawler_id:id,text:el.innerText||el.value||el.textContent||'',href:el.href||el.getAttribute('href')||'',onclick:el.getAttribute('onclick')||'',title:el.getAttribute('title')||'',aria:el.getAttribute('aria-label')||'',class_name:el.getAttribute('class')||'',id_value:el.getAttribute('id')||'',visible:!!(r.width&&r.height),x:r.x,y:r.y,width:r.width,height:r.height};
    })
    """
    items = []
    def collect(frame_like, idx, frame_url):
        try:
            for item in frame_like.evaluate(js) or []:
                item["frame_index"] = idx
                item["frame_url"] = frame_url
                items.append(item)
        except Exception:
            pass
    collect(page, -1, page.url)
    try:
        for idx, frame in enumerate(page.frames):
            if frame == page.main_frame or frame.url == page.url:
                continue
            collect(frame, idx, frame.url)
    except Exception:
        pass
    return items


# -----------------------------------------------------------------------------
# 1. 공공기관
# -----------------------------------------------------------------------------
def download_public_agency():
    task = "공공기관"
    url = PUBLIC_AGENCY_URL
    folder = DOWNLOAD_DIR / "public_agency"
    folder.mkdir(parents=True, exist_ok=True)
    log("=" * 80)
    log(f"Starting task: {task}")
    log(f"URL: {url}")
    with sync_playwright() as p:
        browser = launch_browser(p)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            if not robust_goto(page, url, 30000, 2):
                browser.close()
                return make_result(task, url, "failed", "", "page_load_failed")
            locators = [
                page.get_by_text("다운로드").first,
                page.locator("a", has_text="다운로드").first,
                page.locator("button", has_text="다운로드").first,
                page.locator("[onclick]", has_text="다운로드").first,
            ]
            for idx, loc in enumerate(locators, 1):
                try:
                    if loc.count() == 0:
                        continue
                    log(f"Trying ALIO download locator {idx}")
                    saved, msg = click_download(page, loc, folder, [".zip", ".pdf"], timeout=10000)
                    if saved:
                        extracted = extract_zip_if_needed(saved, folder)
                        browser.close()
                        return make_result(task, url, "success", saved, "public_agency_download_success", extracted)
                    log(f"ALIO locator {idx} failed: {msg}")
                except Exception as e:
                    log(f"ALIO locator {idx} exception: {e}")
            browser.close()
            return make_result(task, url, "failed", "", "public_agency_download_not_found")
        except Exception as e:
            log(traceback.format_exc())
            browser.close()
            return make_result(task, url, "failed", "", str(e))



# -----------------------------------------------------------------------------
# 2. 공직유관단체 - 기존 방식 유지
# -----------------------------------------------------------------------------
def is_affiliated_title(text):
    text = clean_text(text)
    return "공직유관단체 지정 고시" in text and not any(w in text for w in ["임원", "재산공개", "재산 공개", "부동산", "취업심사", "퇴직공직자"])


def find_affiliated_link(page):
    links = page.locator("a")
    candidates = []
    for i in range(links.count()):
        try:
            link = links.nth(i)
            text = clean_text(link.inner_text(timeout=1000))
            if not is_affiliated_title(text):
                continue
            years = re.findall(r"(20\d{2})", text)
            year = max([int(y) for y in years], default=0)
            half = 2 if "하반기" in text else (1 if "상반기" in text else 0)
            candidates.append((year * 10 + half, link, text))
        except Exception:
            pass
    candidates.sort(key=lambda x: x[0], reverse=True)
    if candidates:
        log(f"Latest affiliated org title selected: score={candidates[0][0]} title={candidates[0][2]}")
        return candidates[0][1], candidates[0][2]
    return None, ""


def attachment_targets(page):
    targets = []
    for kw in ["첨부파일", "첨부 파일", "첨부"]:
        try:
            loc = page.get_by_text(kw).first
            if loc.count() > 0:
                targets.append(("main", kw, loc))
        except Exception:
            pass
    try:
        for i, frame in enumerate(page.frames):
            for kw in ["첨부파일", "첨부 파일", "첨부"]:
                try:
                    loc = frame.get_by_text(kw).first
                    if loc.count() > 0:
                        targets.append((f"frame_{i}", kw, loc))
                except Exception:
                    pass
    except Exception:
        pass
    return targets


def excel_links(page):
    out = []
    for item in get_clickables(page):
        text = str(item.get("text", "") or "").strip()
        href = str(item.get("href", "") or "").strip()
        blob = f"{text} {href}".lower()
        if not href or any(x in blob for x in [".pdf", ".hwp", ".hwpx", ".doc", ".docx"]):
            continue
        if ".xlsx" in blob or ".xls" in blob or "fldownload" in href.lower():
            score = (100 if "공직유관단체 지정 고시" in text else 0) + (80 if ".xlsx" in blob else 0) + (70 if ".xls" in blob else 0) + (50 if "fldownload" in href.lower() else 0)
            out.append((score, item))
    out.sort(key=lambda x: x[0], reverse=True)
    return [x[1] for x in out]


def download_affiliated_excel(detail, folder):
    targets = attachment_targets(detail)
    log(f"Attachment click targets found: {len(targets)}")
    clicked = False
    for name, kw, loc in targets[:4]:
        try:
            log(f"Clicking attachment target only: {name} / {kw}")
            loc.click(timeout=5000, no_wait_after=True)
            clicked = True
            time.sleep(2)
            break
        except Exception as e:
            log(f"Attachment target click failed: {name} / {kw} / {e}")
    if not clicked:
        log("No attachment target clicked. Trying direct Excel link search anyway.")
    links = excel_links(detail)
    log(f"Excel file links found after attachment step: {len(links)}")
    for idx, item in enumerate(links[:8], 1):
        text = str(item.get("text", "") or "").strip()
        href = str(item.get("href", "") or "").strip()
        log(f"Trying direct Excel link {idx}: text={text[:120]} href={href[:150]}")
        try:
            fallback = text if (".xlsx" in text.lower() or ".xls" in text.lower()) else "affiliated_org.xlsx"
            saved, msg = direct_download(urljoin(detail.url, href), folder, [".xlsx", ".xls"], referer=detail.url, fallback=fallback)
            if saved:
                return saved, "affiliated_excel_direct_download_success"
            log(f"Direct Excel link {idx} failed: {msg}")
        except Exception as e:
            log(f"Direct Excel link {idx} exception: {e}")
    return None, "excel_file_not_found_after_attachment"


def download_affiliated_org():
    task = "공직유관단체"
    url = AFFILIATED_ORG_URL
    folder = DOWNLOAD_DIR / "affiliated_org"
    folder.mkdir(parents=True, exist_ok=True)
    log("=" * 80)
    log(f"Starting task: {task}")
    log(f"URL: {url}")
    with sync_playwright() as p:
        browser = launch_browser(p)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            if not robust_goto(page, url, 30000, 2):
                browser.close()
                return make_result(task, url, "failed", "", "page_load_failed")
            link, title = find_affiliated_link(page)
            if link is None:
                browser.close()
                return make_result(task, url, "failed", "", "target_title_not_found")
            try:
                with context.expect_page(timeout=5000) as pop:
                    link.click(timeout=5000, no_wait_after=True)
                detail = pop.value
                try:
                    detail.wait_for_load_state("domcontentloaded", timeout=10000)
                except Exception:
                    pass
            except PlaywrightTimeoutError:
                link.click(timeout=5000, no_wait_after=True)
                detail = page
            time.sleep(2)
            saved, msg = download_affiliated_excel(detail, folder)
            browser.close()
            if saved:
                return make_result(task, url, "success", saved, f"{msg} / title={title}", [])
            return make_result(task, url, "failed", "", f"excel_not_downloaded / title={title} / reason={msg}")
        except Exception as e:
            log(traceback.format_exc())
            browser.close()
            return make_result(task, url, "failed", "", str(e))


# -----------------------------------------------------------------------------
# 3. 언론사 - 기존 방식 유지 (최종 안정화 버전, 통째 교체용)
# -----------------------------------------------------------------------------
def js_click_download(page, item, folder, exts, timeout=12000):
    cid = item.get("crawler_id")
    fi = item.get("frame_index", -1)
    script = """(id)=>{const el=document.querySelector(`[data-crawler-id="${id}"]`); if(el) el.click();}"""
    try:
        with page.expect_download(timeout=timeout) as info:
            if fi == -1:
                page.evaluate(script, cid)
            else:
                frames = page.frames
                if 0 <= fi < len(frames):
                    frames[fi].evaluate(script, cid)
                else:
                    page.evaluate(script, cid)
        path = save_download(info.value, folder)
        ok, msg = validate_file(path, exts)
        if ok:
            return path, "ok"
        try:
            path.unlink()
        except Exception:
            pass
        return None, msg
    except PlaywrightTimeoutError:
        return None, "download_timeout"
    except Exception as e:
        return None, str(e)


def download_media_company():
    task = "언론사"
    url = MEDIA_COMPANY_URL
    folder = DOWNLOAD_DIR / "media_company"
    folder.mkdir(parents=True, exist_ok=True)

    log("=" * 80)
    log(f"Starting task: {task}")
    log(f"URL: {url}")

    with sync_playwright() as p:
        browser = launch_browser(p)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            if not robust_goto(page, url, 45000, 8):
                browser.close()
                return make_result(task, url, "failed", "", "page_load_failed")

            # 1차 안정화: DOM / network / 추가 대기
            try:
                page.wait_for_load_state("domcontentloaded", timeout=30000)
            except Exception:
                pass

            try:
                page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                log("Media page networkidle timeout, continuing with extended waits...")

            # 사이트 JS가 늦게 붙는 경우 대비
            page.wait_for_timeout(8000)

            # 2차 안정화: 테이블 또는 다운로드 관련 텍스트 대기
            ready = False

            for attempt in range(1, 9):
                try:
                    ready = page.evaluate(
                        """
                        () => {
                            const bodyText = (document.body && document.body.innerText || "").toLowerCase();
                            const rows = document.querySelectorAll("table tr");
                            const clickables = Array.from(document.querySelectorAll("a, button, input, [onclick], [role='button']"));
                            const hasDownloadLike = clickables.some(el => {
                                const blob = [
                                    el.innerText || "",
                                    el.value || "",
                                    el.textContent || "",
                                    el.getAttribute("title") || "",
                                    el.getAttribute("aria-label") || "",
                                    el.getAttribute("href") || "",
                                    el.getAttribute("onclick") || "",
                                    el.getAttribute("class") || "",
                                    el.getAttribute("id") || ""
                                ].join(" ").toLowerCase();

                                return (
                                    blob.includes("엑셀") ||
                                    blob.includes("excel") ||
                                    blob.includes("xlsx") ||
                                    blob.includes("xls") ||
                                    blob.includes("download") ||
                                    blob.includes("다운로드")
                                );
                            });

                            return (
                                rows.length > 1 ||
                                hasDownloadLike ||
                                bodyText.includes("등록") ||
                                bodyText.includes("신고") ||
                                bodyText.includes("제호")
                            );
                        }
                        """
                    )
                except Exception:
                    ready = False

                if ready:
                    log(f"Media page readiness detected on attempt {attempt}")
                    break

                log(f"Media page not ready yet, waiting... attempt={attempt}/8")
                page.wait_for_timeout(5000)

            # 그래도 준비가 안 된 것 같으면 1회 reload 후 재대기
            if not ready:
                log("Media page still not ready. Reloading once...")
                try:
                    page.reload(wait_until="domcontentloaded", timeout=45000)
                    page.wait_for_timeout(10000)
                    try:
                        page.wait_for_load_state("networkidle", timeout=30000)
                    except Exception:
                        pass
                except Exception as e:
                    log(f"Media page reload failed, continuing anyway: {e}")

            # 3차: 명시적 selector 클릭 시도
            selector_candidates = [
                "a:has-text('엑셀')",
                "button:has-text('엑셀')",
                "a:has-text('Excel')",
                "button:has-text('Excel')",
                "a:has-text('EXCEL')",
                "button:has-text('EXCEL')",
                "a:has-text('다운로드')",
                "button:has-text('다운로드')",
                "[title*='엑셀']",
                "[title*='Excel']",
                "[aria-label*='엑셀']",
                "[aria-label*='Excel']",
                "[href*='.xls']",
                "[href*='download']",
                "[onclick*='excel']",
                "[onclick*='Excel']",
                "[onclick*='xls']",
                "[onclick*='download']",
            ]

            for idx, selector in enumerate(selector_candidates, 1):
                try:
                    loc = page.locator(selector).first
                    if loc.count() == 0:
                        continue

                    log(f"Trying media explicit selector {idx}: {selector}")

                    saved, msg = click_download(
                        page,
                        loc,
                        folder,
                        [".xlsx", ".xls"],
                        timeout=25000
                    )

                    if saved:
                        browser.close()
                        return make_result(
                            task,
                            url,
                            "success",
                            saved,
                            f"media_excel_download_success_by_selector: {selector}",
                            []
                        )

                    log(f"Media explicit selector {idx} failed: {msg}")

                except Exception as e:
                    log(f"Media explicit selector {idx} exception: {selector} / {e}")

            # 4차: get_clickables 기반 후보 탐색
            log("Trying media candidate search...")

            candidates = []

            for attempt in range(1, 7):
                candidates = []

                for item in get_clickables(page):
                    blob = (
                        f"{item.get('text','')} "
                        f"{item.get('href','')} "
                        f"{item.get('onclick','')} "
                        f"{item.get('title','')} "
                        f"{item.get('aria','')} "
                        f"{item.get('class_name','')} "
                        f"{item.get('id_value','')}"
                    ).lower()

                    score = 0

                    if "엑셀" in blob or "excel" in blob:
                        score += 200
                    if ".xlsx" in blob or ".xls" in blob:
                        score += 220
                    if "download" in blob or "다운로드" in blob:
                        score += 100
                    if "file" in blob:
                        score += 30
                    if "export" in blob:
                        score += 80
                    if item.get("href"):
                        score += 20
                    if item.get("onclick"):
                        score += 80

                    # 눈에 보이지 않아도 onclick/href가 있으면 후보로 둔다.
                    if not item.get("visible") and not item.get("onclick") and not item.get("href"):
                        score -= 100

                    for bad in [
                        "검색", "목록", "인쇄", "공유", "로그인", "회원가입",
                        "facebook", "twitter", "kakao", "개인정보", "사이트맵"
                    ]:
                        if bad.lower() in blob:
                            score -= 150

                    if score > 0:
                        item["score"] = score
                        candidates.append(item)

                candidates.sort(key=lambda x: x["score"], reverse=True)

                log(f"Media download candidates found on attempt {attempt}: {len(candidates)}")

                if candidates:
                    break

                page.wait_for_timeout(5000)

            # 5차: href가 직접 파일/다운로드면 direct_download 먼저 시도
            for idx, item in enumerate(candidates[:15], 1):
                text = str(item.get("text", "") or "").strip()
                href = str(item.get("href", "") or "").strip()
                blob = f"{text} {href} {item.get('onclick','')}".lower()

                if href and (
                    ".xlsx" in blob
                    or ".xls" in blob
                    or "download" in blob
                    or "fldownload" in blob
                ):
                    try:
                        log(f"Trying media direct href {idx}: text={text[:80]} href={href[:150]}")
                        fallback = "MainServiceExcel.xls"

                        saved, msg = direct_download(
                            urljoin(page.url, href),
                            folder,
                            [".xlsx", ".xls"],
                            referer=page.url,
                            fallback=fallback
                        )

                        if saved:
                            browser.close()
                            return make_result(
                                task,
                                url,
                                "success",
                                saved,
                                "media_excel_download_success_by_direct_href",
                                []
                            )

                        log(f"Media direct href {idx} failed: {msg}")

                    except Exception as e:
                        log(f"Media direct href {idx} exception: {e}")

            # 6차: JS click 후보 시도
            for idx, item in enumerate(candidates[:15], 1):
                log(
                    f"Trying media candidate {idx}: "
                    f"score={item.get('score')} "
                    f"visible={item.get('visible')} "
                    f"text={str(item.get('text','')).strip().replace(chr(10), ' ')[:80]} "
                    f"href={str(item.get('href',''))[:100]} "
                    f"onclick={str(item.get('onclick',''))[:120]}"
                )

                saved, msg = js_click_download(
                    page,
                    item,
                    folder,
                    [".xlsx", ".xls"],
                    timeout=25000
                )

                if saved:
                    browser.close()
                    return make_result(
                        task,
                        url,
                        "success",
                        saved,
                        "media_excel_download_success_by_js_candidate",
                        []
                    )

                log(f"Media candidate {idx} JS click failed: {msg}")

            # 실패 시 디버그 저장
            try:
                debug_html = DEBUG_DIR / f"media_page_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                debug_html.write_text(page.content(), encoding="utf-8")
                log(f"Media failure HTML saved: {debug_html}")
            except Exception as e:
                log(f"Media failure HTML save failed: {e}")

            try:
                debug_png = DEBUG_DIR / f"media_page_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                page.screenshot(path=str(debug_png), full_page=True)
                log(f"Media failure screenshot saved: {debug_png}")
            except Exception as e:
                log(f"Media failure screenshot save failed: {e}")

            log("⚠️ Media download failed: no valid Excel file downloaded")
            browser.close()
            return make_result(task, url, "failed", "", "media_excel_download_not_found")

        except Exception:
            log(traceback.format_exc())

            try:
                debug_html = DEBUG_DIR / f"media_exception_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                debug_html.write_text(page.content(), encoding="utf-8")
                log(f"Media exception HTML saved: {debug_html}")
            except Exception:
                pass

            browser.close()
            return make_result(task, url, "failed", "", "exception_occurred")

# -----------------------------------------------------------------------------
# 4. 공공기관 자회사 - 새 창 URL(itemReportTerm.do) 기준
# -----------------------------------------------------------------------------
def prepare_alio_list(page):
    for _ in range(3):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        except Exception:
            pass
        time.sleep(1.2)
    try:
        page.evaluate("window.scrollTo(0, 0)")
    except Exception:
        pass
    time.sleep(1)



def clean_agency_name_from_scope(scope_text, fallback=""):
    s = clean_text(scope_text).replace("보고서", " ")
    s = clean_text(s)
    type_markers = ["공기업(시장형)", "공기업(준시장형)", "준정부기관(위탁집행형)", "준정부기관(기금관리형)", "기타공공기관", "부설기관"]
    cut_positions = [s.find(m) for m in type_markers if m in s]
    if cut_positions:
        s = s[:min(cut_positions)].strip()
    return s or fallback

def collect_report_buttons(page):
    js = r"""
    () => {
      const out=[];
      function clean(s){return (s||'').replace(/\s+/g,' ').trim();}
      const buttons=Array.from(document.querySelectorAll('a,button,input[type="button"],input[type="submit"],[role="button"],[onclick]'));
      buttons.forEach((el,i)=>{
        const r=el.getBoundingClientRect();
        if(!r.width || !r.height) return;
        const text=clean(el.innerText||el.value||el.textContent||'');
        const title=clean(el.getAttribute('title')||'');
        const aria=clean(el.getAttribute('aria-label')||'');
        const href=el.href||el.getAttribute('href')||'';
        const onclick=el.getAttribute('onclick')||'';
        const blob=`${text} ${title} ${aria} ${href} ${onclick}`.toLowerCase();
        const isReport = text==='보고서' || title==='보고서' || aria==='보고서' || blob.includes('itemreport');
        if(!isReport) return;
        if(blob.includes('pdf') || blob.includes('download') || blob.includes('첨부') || blob.includes('다운로드')) return;
        const scope = el.closest('tr') || el.closest('li') || el.closest('div') || el.parentElement;
        const scopeText = scope ? clean(scope.innerText || scope.textContent || '') : '';
        const contextHtml = scope ? (scope.outerHTML || '').slice(0, 30000) : '';
        const cid='report_btn_'+i+'_'+Math.random().toString(36).slice(2);
        el.setAttribute('data-report-button-id', cid);
        out.push({id:cid,text:text||title||aria||'보고서',href,onclick,scopeText,contextHtml,x:r.x,y:r.y,width:r.width,height:r.height});
      });
      return out;
    }
    """
    try:
        raw = page.evaluate(js) or []
    except Exception as e:
        log(f"Report button scan failed: {e}")
        return []
    raw.sort(key=lambda x: (float(x.get("y", 0) or 0), float(x.get("x", 0) or 0)))
    groups = []
    for item in raw:
        y = float(item.get("y", 0) or 0)
        for g in groups:
            if abs(g["y"] - y) < 12:
                g["items"].append(item)
                break
        else:
            groups.append({"y": y, "items": [item]})
    selected = []
    for g in groups:
        g["items"].sort(key=lambda x: float(x.get("x", 0) or 0))
        selected.append(g["items"][0])
    if len(selected) >= 700 and len(selected) % 2 == 0:
        selected = selected[:len(selected)//2]
    for idx, item in enumerate(selected, 1):
        name = clean_text(item.get("scopeText", "")).replace("보고서", " ")
        name = re.sub(r"\s+", " ", name).strip()
        item["agencyName"] = clean_agency_name_from_scope(name, f"agency_{idx}")
    log(f"Report raw candidates found: {len(raw)}")
    log(f"Report buttons after grouping: {len(selected)}")
    for idx, item in enumerate(selected[:5], 1):
        log(f"Report candidate sample {idx}: text={item.get('text')} agency={item.get('agencyName')} href={str(item.get('href'))[:60]} onclick={str(item.get('onclick'))[:80]}")
    return selected


def is_list_like_text(text):
    markers = ["Total 355", "전체기관", "기관 유형 선택", "주무 부처 선택", "지역 선택"]
    return bool(text) and len(text) > 30000 and sum(1 for m in markers if m in text) >= 3


def parse_companies_from_text(text, agency):
    text = clean_text(text)
    if not text or is_list_like_text(text):
        return []
    if any(x in text for x in NO_DATA_WORDS):
        return []
    if not any(k in text for k in TEXT_KEYWORDS):
        return []
    candidates = set()
    patterns = [
        r"[가-힣A-Za-z0-9·&.,\- ]{2,40}\(주\)",
        r"주식회사\s*[가-힣A-Za-z0-9·&.,\- ]{2,40}",
        r"[가-힣A-Za-z0-9·&.,\- ]{2,40}주식회사",
        r"[가-힣A-Za-z0-9·&.,\- ]{2,40}\(유\)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            name = clean_text(m.group(0))
            name = re.sub(r"^(법인명|회사명|자회사|타법인|출자회사)\s*", "", name)
            bads = ["보고서", "다운로드", "첨부", "출자", "비율", "금액", "합계", "계 ", "기관"]
            if name and 2 <= len(name) <= 50 and not any(b in name for b in bads):
                candidates.add(name)
    return [{"agency_name": agency, "company_name": n, "classification": "자회사", "year": None} for n in sorted(candidates)]



def clean_agency_name_from_scope(scope_text, fallback=""):
    s = clean_text(scope_text).replace("보고서", " ")
    s = clean_text(s)
    type_markers = ["공기업(시장형)", "공기업(준시장형)", "준정부기관(위탁집행형)", "준정부기관(기금관리형)", "기타공공기관", "부설기관"]
    cut_positions = [s.find(m) for m in type_markers if m in s]
    if cut_positions:
        s = s[:min(cut_positions)].strip()
    return s or fallback


def flatten_columns(df):
    cols = []
    for c in df.columns:
        if isinstance(c, tuple):
            cols.append(" ".join(clean_text(x) for x in c if clean_text(x) and not str(x).startswith("Unnamed")))
        else:
            s = clean_text(c)
            cols.append("" if s.startswith("Unnamed") else s)
    df = df.copy()
    df.columns = cols
    return df


def normalize_company_cell(value):
    s = clean_text(value)
    if not s or s.lower() == "nan":
        return ""
    s = re.sub(r"\.(pdf|hwp|hwpx|docx?|xlsx?)$", "", s, flags=re.I)
    s = re.sub(r"[_\-\s]*(20\d{2})(년)?$", "", s).strip()
    s = re.sub(r"^(법인명|회사명|출자회사|출자법인|자회사|타법인)\s*[:：]?\s*", "", s)
    return s.strip(" _-[]")


def is_valid_company_name(name, agency=""):
    n = clean_text(name)
    if len(n) < 2 or len(n) > 80:
        return False
    bad = ["기관명", "보고서", "다운로드", "첨부", "파일", "선택하세요", "구분", "합계", "계 ", "비율", "금액", "지분", "출자", "소계", "합 계"]
    if any(b in n for b in bad):
        return False
    agency_key = re.sub(r"\s+", "", clean_agency_name_from_scope(agency, agency)).lower()
    name_key = re.sub(r"\s+", "", n).lower()
    if agency_key and agency_key == name_key:
        return False
    # 회사명 형태가 어느 정도 있는 값만 허용. 단, 영문/약어 자회사도 있어 완전 제한하지는 않음.
    company_markers = ["(주)", "주식회사", "(유)", "유한회사", "㈜", "corp", "co.", "ltd", "KOR", "kaz", "KAZ"]
    if any(m.lower() in n.lower() for m in company_markers):
        return True
    # 한글 3자 이상 단독 법인명도 허용하되, 숫자/분기/부처명 같은 값은 제외
    if re.fullmatch(r"[가-힣A-Za-z0-9·&.,\- ]{3,40}", n):
        return True
    return False


def rows_from_company_names(names, agency):
    out = []
    seen = set()
    for name in names:
        n = normalize_company_cell(name)
        if not is_valid_company_name(n, agency):
            continue
        key = re.sub(r"\s+", "", n).lower()
        if key not in seen:
            seen.add(key)
            out.append({"agency_name": agency, "company_name": n, "classification": "자회사", "year": None})
    return out


def extract_names_from_table_text(table_text):
    names = []
    # 표 내부에서만 정규식 보조 추출. 첨부파일명 전체 영역은 여기로 들어오지 않음.
    for pat in [
        r"[가-힣A-Za-z0-9·&.,\- ]{2,50}\(주\)",
        r"\(주\)[가-힣A-Za-z0-9·&.,\- ]{2,50}",
        r"주식회사\s*[가-힣A-Za-z0-9·&.,\- ]{2,50}",
        r"[가-힣A-Za-z0-9·&.,\- ]{2,50}주식회사",
        r"[가-힣A-Za-z0-9·&.,\- ]{2,50}\(유\)",
    ]:
        for m in re.finditer(pat, table_text):
            names.append(m.group(0))
    return names


def extract_report_snapshots(detail):
    """보고서 새 창의 main frame + iframe HTML/text를 모두 수집합니다."""
    snapshots = []
    try:
        try:
            detail.wait_for_load_state("domcontentloaded", timeout=12000)
        except Exception:
            pass
        try:
            detail.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        time.sleep(1.5)
        snapshots.append(("main", detail.url, detail.content()))
        for i, frame in enumerate(detail.frames):
            try:
                if frame == detail.main_frame:
                    continue
                snapshots.append((f"frame_{i}", frame.url, frame.content()))
            except Exception:
                pass
    except Exception:
        pass
    return snapshots

def normalize_alio_company_name(value):
    s = clean_text(value)
    if not s or s.lower() == "nan":
        return ""
    s = re.sub(r"\.(pdf|hwp|hwpx|docx?|xlsx?)$", "", s, flags=re.I)
    s = re.sub(r"[_\-\s]*(20\d{2})(년)?$", "", s).strip()
    s = re.sub(r"^(법인명|회사명|출자회사|출자법인|자회사|타법인)\s*[:：]?\s*", "", s).strip()
    # ALIO HTML의 <br> 때문에 들어간 이름 내부 공백 정리
    if s.startswith("주식회사"):
        rest = s.replace("주식회사", "", 1).strip()
        rest = re.sub(r"\s+", "", rest)
        s = "주식회사 " + rest if rest else "주식회사"
    else:
        s = re.sub(r"(?<=[가-힣A-Za-z0-9\)])\s+(?=[가-힣A-Za-z0-9\(])", "", s)
    return s.strip(" _-[]")


def is_company_like(name, agency=""):
    n = clean_text(name)
    if len(n) < 2 or len(n) > 80:
        return False
    bad = ["기관명", "보고서", "다운로드", "첨부", "파일", "선택하세요", "구분", "합계", "소계", "지분율", "취득가액", "장부가액", "해당연도", "설립일자", "주요사업"]
    if any(b in n for b in bad):
        return False
    agency_key = re.sub(r"\s+", "", clean_agency_name_from_scope(agency, agency)).lower() if "clean_agency_name_from_scope" in globals() else re.sub(r"\s+", "", agency).lower()
    name_key = re.sub(r"\s+", "", n).lower()
    if agency_key and agency_key == name_key:
        return False
    if any(m.lower() in n.lower() for m in ["(주)", "주식회사", "(유)", "유한회사", "㈜", "corp", "co.", "ltd"]):
        return True
    # 영문 합작법인/약어 등 예외 허용
    if re.fullmatch(r"[A-Za-z0-9&.,\- ]{3,40}", n):
        return True
    return False


def html_to_text_cell(cell_html):
    import html as _html
    s = re.sub(r"(?i)<br\s*/?>", " ", str(cell_html))
    s = re.sub(r"(?is)<script.*?</script>", " ", s)
    s = re.sub(r"(?is)<style.*?</style>", " ", s)
    s = re.sub(r"(?is)<[^>]+>", " ", s)
    s = _html.unescape(s)
    return clean_text(s)


def normalize_alio_company_name(value):
    s = clean_text(value)
    if not s or s.lower() == "nan":
        return ""
    s = re.sub(r"\.(pdf|hwp|hwpx|docx?|xlsx?)$", "", s, flags=re.I)
    s = re.sub(r"[_\-\s]*(20\d{2})(년)?$", "", s).strip()
    s = re.sub(r"^(법인명|회사명|출자회사|출자법인|자회사|타법인)\s*[:：]?\s*", "", s).strip()
    # ALIO <br> 때문에 생기는 회사명 내부 공백 정리
    if s.startswith("주식회사"):
        rest = s.replace("주식회사", "", 1).strip()
        rest = re.sub(r"\s+", "", rest)
        s = "주식회사 " + rest if rest else "주식회사"
    else:
        s = re.sub(r"(?<=[가-힣A-Za-z0-9\)])\s+(?=[가-힣A-Za-z0-9\(])", "", s)
    return s.strip(" _-[]")


def is_alio_company_name(name, agency=""):
    n = clean_text(name)
    if len(n) < 2 or len(n) > 80:
        return False
    bad = ["기관명", "보고서", "다운로드", "첨부", "파일", "선택하세요", "구분", "합계", "소계", "지분율", "취득가액", "장부가액", "해당연도", "설립일자", "주요사업", "사전협의"]
    if any(b in n for b in bad):
        return False
    try:
        agency_key = re.sub(r"\s+", "", clean_agency_name_from_scope(agency, agency)).lower()
    except Exception:
        agency_key = re.sub(r"\s+", "", agency).lower()
    name_key = re.sub(r"\s+", "", n).lower()
    if agency_key and agency_key == name_key:
        return False
    if any(m.lower() in n.lower() for m in ["(주)", "주식회사", "(유)", "유한회사", "㈜", "corp", "co.", "ltd"]):
        return True
    # 영문 합작법인/약어 등 예외 허용
    if re.fullmatch(r"[A-Za-z0-9&.,\- ]{3,40}", n):
        return True
    return False


def parse_tables_from_html(html, agency):
    """
    ALIO 보고서 HTML에서 공공기관 자회사 정보를 추출한다.

    최종 추출 기준:
    - 모든 table의 tr/td를 순회한다.
    - 데이터 row의 첫 번째 td가 정확히 "자회사"인 경우만 선택한다.
    - 선택된 row의 두 번째 td를 자회사명으로 사용한다.

    제외 기준:
    - 총괄표 row 제외:
      예) 자회사 | 3 | 3 | 2 | 2 | 2
      → 두 번째 td가 숫자면 제외한다.
    - 연도별 표 row 제외:
      예) 자회사 | 회사명 | 보유/미보유 | 2021년 ...
      → 세 번째 td가 "보유" 또는 "미보유"면 제외한다.
    - 헤더 row 제외:
      예) 자회사 출자회사 재출자회사 | 법인명 | ...
      → 첫 번째 td가 정확히 "자회사"가 아니므로 제외된다.

    주의:
    - 회사명에 "합자회사", "투자회사", "선박투자회사", "부동산투자회사"가 들어갈 수 있으므로
      "자회사"라는 문자열을 부분 포함 기준으로 제거하면 안 된다.
    """
    html = str(html or "")
    rows_out = []

    def get_all_tables(full_html):
        return re.findall(r"(?is)<table\b[^>]*>.*?</table>", str(full_html or ""))

    def table_rows(table_html):
        rows = []

        trs = re.findall(r"(?is)<tr\b[^>]*>.*?</tr>", str(table_html or ""))

        for tr in trs:
            cells_html = re.findall(r"(?is)<t[dh]\b[^>]*>.*?</t[dh]>", tr)

            cells = []
            for c in cells_html:
                cell_text = html_to_text_cell(c)
                cell_text = clean_text(cell_text)

                if cell_text:
                    cells.append(cell_text)

            if cells:
                rows.append(cells)

        return rows

    def is_numeric_like(value):
        """
        총괄표의 숫자 셀, 금액/비율 셀 등을 걸러내기 위한 판별 함수.
        """
        s = clean_text(value)

        if not s:
            return True

        # '-', '0', '1,234', '99.9', '100.0' 같은 값 처리
        if s in {"-", "–", "—"}:
            return True

        tmp = re.sub(r"[\s,.\-%원억원백만원천원]", "", s)

        if tmp.isdigit():
            return True

        if re.fullmatch(r"[\d,.\s]+%?", s):
            return True

        return False

    def clean_company_name(value):
        """
        td[1]에서 가져온 회사명을 정리한다.

        중요:
        - "자회사"를 부분 포함 금지어로 쓰면 안 된다.
        - "사모투자합자회사", "선박투자회사", "부동산투자회사"가 모두 삭제되는 문제가 생긴다.
        """
        name = normalize_alio_company_name(value)

        if not name:
            return ""

        if is_numeric_like(name):
            return ""

        # 정확히 헤더 그 자체인 경우만 제거
        exact_bad_words = {
            "구분",
            "법인명",
            "회사명",
            "자회사",
            "출자회사",
            "재출자회사",
            "설립일자",
            "취득일자",
            "주요사업",
            "합계",
            "소계",
            "해당사항 없음",
        }

        if name in exact_bad_words:
            return ""

        # 회사명 내부에 들어갈 가능성이 거의 없는 설명성 문구만 부분 포함 제거
        # 절대 넣으면 안 되는 것:
        # "자회사", "출자회사", "재출자회사", "출자", "투자"
        contains_bad_words = [
            "재무현황",
            "취득가액",
            "장부가액",
            "사전협의",
            "수행여부",
            "다운로드",
            "보고서",
            "선택하세요",
        ]

        if any(w in name for w in contains_bad_words):
            return ""

        agency_key = re.sub(r"\s+", "", clean_text(agency)).lower()
        name_key = re.sub(r"\s+", "", name).lower()

        if agency_key and agency_key == name_key:
            return ""

        return name

    def company_dedupe_key(name):
        """
        중복 제거용 key.
        너무 강한 정규화는 피한다.
        단, 영문명에서 마침표/쉼표/공백 차이로 중복되는 경우는 완화한다.
        """
        s = clean_text(name).lower()
        s = re.sub(r"\s+", "", s)

        s = s.rstrip(".")
        s = s.replace(".,", ",")
        s = s.replace(",.", ",")
        s = s.replace(".", "")
        s = s.replace(",", "")

        return s

    def is_detail_subsidiary_row(cells):
        """
        실제 '해당연도 타법인 투·출자 현황'의 자회사 데이터 row인지 판별한다.

        통과해야 하는 구조:
        - 자회사 | 법인명 | 설립일자/취득일자 또는 주요사업 관련 셀 | ...

        제외되는 구조:
        - 헤더:
          자회사 출자회사 재출자회사 | 법인명 | ...
        - 총괄표:
          자회사 | 3 | 3 | 2 | 2 | 2
        - 연도별 표:
          자회사 | 법인명 | 보유/미보유 | 2021년 ...
        """
        if len(cells) < 2:
            return False

        first = clean_text(cells[0])

        # 핵심: 부분 포함이 아니라 정확히 일치해야 함
        # 헤더의 "자회사 출자회사 재출자회사"는 여기서 제외됨
        if first != "자회사":
            return False

        second = clean_text(cells[1])

        # 총괄표 제외: 자회사 | 숫자 | 숫자 ...
        if not second or is_numeric_like(second):
            return False

        # 연도별 표 제외: 자회사 | 법인명 | 보유/미보유 | ...
        if len(cells) >= 3:
            third = clean_text(cells[2])
            if third in {"보유", "미보유"}:
                return False

        return True

    tables = get_all_tables(html)

    if not tables:
        return rows_out

    companies = []
    seen = set()

    # 모든 table을 훑는다.
    # 단, row 구조 기준으로 총괄표/연도별표/헤더는 제외된다.
    for table_html in tables:
        for cells in table_rows(table_html):
            if not is_detail_subsidiary_row(cells):
                continue

            company_name = clean_company_name(cells[1])

            if not company_name:
                continue

            key = company_dedupe_key(company_name)

            if key in seen:
                continue

            seen.add(key)
            companies.append(company_name)

    for name in companies:
        rows_out.append({
            "agency_name": agency,
            "company_name": name,
            "classification": "자회사",
            "year": 2025,
        })

    return rows_out


def parse_report_html(html, agency):
    """
    ALIO 보고서 본문 표에서 자회사 정보를 추출한다.

    문서목차, 총괄표, 연도별 표, 첨부파일명은 자회사로 간주하지 않는다.
    실제 데이터 row의 구조:
    자회사 | 법인명 | ...
    만 사용한다.
    """
    return parse_tables_from_html(html, agency)

def extract_direct_url_from_text(blob):
    if not blob:
        return ""
    blob = unquote(str(blob)).replace("&amp;", "&")
    m = re.search(r"https?://www\.alio\.go\.kr/item/itemReportTerm\.do\?[^\s'\"<>]+", blob)
    if m:
        return m.group(0)
    m = re.search(r"itemReportTerm\.do\?[^\s'\"<>]+", blob)
    if m:
        return urljoin("https://www.alio.go.kr/item/", m.group(0))
    apba = re.search(r"apbaId[=:'\"&\s]+([A-Za-z0-9]+)", blob)
    disc = re.search(r"disclosureNo[=:'\"&\s]+([0-9A-Za-z_-]+)", blob)
    root = re.search(r"reportFormRootNo[=:'\"&\s]+([0-9A-Za-z]+)", blob)
    if apba:
        report_root = root.group(1) if root else "31901"
        disclosure = disc.group(1) if disc else ""
        return f"https://www.alio.go.kr/item/itemReportTerm.do?apbaId={apba.group(1)}&reportFormRootNo={report_root}&disclosureNo={disclosure}"
    return ""


def click_report_and_parse(page, candidate, agency, dialog_state, idx):
    dialog_state["message"] = ""
    selector = f'[data-report-button-id="{candidate.get("id")}"]'
    report_url = extract_direct_url_from_text(" ".join([str(candidate.get("href", "")), str(candidate.get("onclick", "")), str(candidate.get("contextHtml", ""))]))

    popup = None
    if not report_url:
        try:
            with page.context.expect_page(timeout=2000) as pop:
                page.locator(selector).click(timeout=7000, no_wait_after=True)
            popup = pop.value
            try:
                popup.wait_for_load_state("domcontentloaded", timeout=15000)
            except Exception:
                pass
            report_url = popup.url
        except PlaywrightTimeoutError:
            if dialog_state.get("message"):
                return [], f"dialog_skip: {dialog_state.get('message')}"
            return [], "no_popup_or_report_url"
        except Exception as e:
            if dialog_state.get("message"):
                return [], f"dialog_skip: {dialog_state.get('message')}"
            return [], f"popup_click_failed: {e}"
    else:
        try:
            page.locator(selector).click(timeout=7000, no_wait_after=True)
            time.sleep(0.5)
        except Exception:
            pass

    if dialog_state.get("message"):
        try:
            if popup:
                popup.close()
        except Exception:
            pass
        return [], f"dialog_skip: {dialog_state.get('message')}"

    detail = popup
    if report_url and (not popup or popup.is_closed()):
        try:
            detail = page.context.new_page()
            detail.goto(report_url, wait_until="commit", timeout=25000)
        except Exception as e:
            return [], f"open_report_url_failed: {e} url={report_url}"

    if not detail:
        return [], "no_popup_or_report_url"

    try:
        snapshots = extract_report_snapshots(detail) if "extract_report_snapshots" in globals() else [("main", detail.url, detail.content())]
        rows = []
        html_len = 0
        for source, frame_url, html in snapshots:
            html_len += len(html or "")
            parsed = parse_report_html(html or "", agency)
            if parsed:
                log(f"Report frame parsed: agency={agency} source={source} url={frame_url} rows={len(parsed)}")
            rows.extend(parsed)

        seen = set()
        unique = []
        for r in rows:
            key = re.sub(r"\s+", "", clean_text(r.get("company_name"))).lower()
            if key and key not in seen:
                seen.add(key)
                unique.append(r)

        if idx <= 20 or "강원랜드" in agency or "공영홈쇼핑" in agency:
            try:
                dbg = DEBUG_DIR / f"alio_report_html_{idx:03d}_{safe_filename(agency)}.html"
                dbg.write_text("\n\n<!-- SNAPSHOT SPLIT -->\n\n".join(h for _, _, h in snapshots)[:500000], encoding="utf-8")
            except Exception:
                pass

        try:
            detail.close()
        except Exception:
            pass

        if unique:
            return unique, f"popup_report_rows_found={len(unique)} url={report_url or ''} snapshots={len(snapshots)} html_len={html_len}"
        return [], f"popup_report_no_rows url={report_url or ''} snapshots={len(snapshots)} html_len={html_len}"
    except Exception as e:
        try:
            if detail:
                detail.close()
        except Exception:
            pass
        return [], f"report_parse_failed: {e}"

def latest_names(raw_rows):
    latest = {}
    for row in raw_rows:
        name = clean_text(row.get("company_name"))
        if not name:
            continue
        year = row.get("year") or 0
        if name not in latest or year >= latest[name].get("year", 0):
            latest[name] = {"year": year, "name": name}
    return sorted(latest.keys())


def collect_public_agency_subsidiaries():
    task = "공공기관 자회사"
    url = PUBLIC_SUBSIDIARY_ALIO_URL
    log("=" * 80)
    log(f"Starting task: {task}")
    log(f"URL: {url}")
    raw_rows = []
    processed = 0
    no_data = 0
    errors = 0
    with sync_playwright() as p:
        browser = launch_browser(p)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        dialog_state = {"message": ""}
        def on_dialog(dialog):
            msg = dialog.message
            dialog_state["message"] = msg
            log(f"Dialog detected: {msg}")
            try:
                dialog.accept()
            except Exception:
                pass
        page.on("dialog", on_dialog)
        try:
            if not robust_goto(page, url, 30000, 3):
                browser.close()
                return make_result(task, url, "failed", "", "page_load_failed")
            log("ALIO list page opened")
            prepare_alio_list(page)
            buttons = collect_report_buttons(page)
            log(f"Report buttons found: {len(buttons)}")
            if not buttons:
                browser.close()
                return make_result(task, url, "failed", "", "report_buttons_not_found")
            for idx, candidate in enumerate(buttons, 1):
                if page.is_closed():
                    errors += 1
                    log("Listing page closed. Stop public agency subsidiary loop.")
                    break
                agency = clean_text(candidate.get("agencyName")) or f"agency_{idx}"
                log(f"Trying report {idx}/{len(buttons)}: agency={agency}")
                rows, msg = click_report_and_parse(page, candidate, agency, dialog_state, idx)
                if rows:
                    processed += 1
                    raw_rows.extend(rows)
                    log(f"Report parsed: {agency} / {msg}")
                else:
                    no_data += 1
                    log(f"No subsidiary data, skipped: {agency} / {msg}")
            names = latest_names(raw_rows)
            output = OUTPUT_DIR / "public_agency_subsidiaries.xlsx"
            pd.DataFrame({"법인명": names}).to_excel(output, index=False, engine="openpyxl")
            browser.close()
            msg = f"public_agency_subsidiaries_saved / subsidiaries={len(names)} / reports_processed={processed} / no_data={no_data} / errors={errors}"
            return make_result(task, url, "success", output, msg, [])
        except Exception as e:
            log(traceback.format_exc())
            browser.close()
            return make_result(task, url, "failed", "", str(e))




def save_summary(results):
    df = pd.DataFrame(results)
    out = OUTPUT_DIR / f"download_summary_{RUN_TIME}.xlsx"
    latest = OUTPUT_DIR / "latest_download_summary.xlsx"
    df.to_excel(out, index=False, engine="openpyxl")
    df.to_excel(latest, index=False, engine="openpyxl")
    log(f"Summary saved: {out}")
    log(f"Latest summary saved: {latest}")


def main():
    log("Program started")
    log(f"Base directory: {BASE_DIR}")
    tasks = [
        ("공공기관", PUBLIC_AGENCY_URL, download_public_agency),
        ("공직유관단체", AFFILIATED_ORG_URL, download_affiliated_org),
        ("언론사", MEDIA_COMPANY_URL, download_media_company),
	("공공기관 자회사", PUBLIC_SUBSIDIARY_ALIO_URL, collect_public_agency_subsidiaries)
    ]
    results = []
    for name, url, func in tasks:
        try:
            r = func()
            results.append(r)
            log(f"Task result: {name} -> {r.get('status')} / {r.get('message')}")
        except Exception as e:
            log(f"Unexpected error in {name} task: {e}")
            r = make_result(name, url, "failed", "", str(e))
            results.append(r)
            log(f"Task result: {name} -> failed / {e}")
    save_summary(results)
    success = sum(1 for r in results if str(r.get("status", "")).lower() == "success")
    log("=" * 80)
    log(f"Success count: {success} / {len(results)}")
    log("All tasks completed successfully" if success == len(results) else "Some tasks failed or need review")
    for r in results:
        log(f"Final result - {r.get('task')}: {r.get('status')} / {r.get('downloaded_file')} / {r.get('message')}")
    log("Program finished")
    print()
    print("Done.")
    print(f"Success count: {success} / {len(results)}")
    print("All tasks completed successfully" if success == len(results) else "Some tasks failed or need review")
    print(f"Check downloads folder: {DOWNLOAD_DIR}")
    print(f"Check output folder: {OUTPUT_DIR}")
    print(f"Check log file: {LOG_FILE}")


if __name__ == "__main__":
    main()
