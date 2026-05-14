PO_Crawler 운영 메모

1. 주요 실행 파일
- main.py: 데이터 수집/크롤링 본체
- update.py: 전체 업데이트 및 merged_list.csv 생성
- flask_app.py: 검색 웹앱 실행
- validate_subsidiary_reports.py: ALIO 자회사 보고서 검증

2. 주요 산출물
- downloads/merged_list.csv
  검색 웹앱이 사용하는 최종 통합 리스트

- downloads/last_update.txt
  최근 업데이트 시간

- output/public_agency_subsidiaries.xlsx
  ALIO 타법인 투·출자 현황 상세표 기준 공공기관 자회사 목록

- output/subsidiary_validation_report.xlsx
  자회사 검증 리포트

3. 전체 업데이트 방법
CMD에서 프로젝트 폴더로 이동 후 실행:

python update.py

4. 자회사 검증 방법

python validate_subsidiary_reports.py

정상 기준:
- ERROR: 0
- NO_SUMMARY_COUNT: 0
- MISMATCH: 한국탄소산업진흥원 1건만 존재 가능
  사유: ALIO 총괄표와 상세표가 불일치하는 원천 보고서 오류 케이스.
  최종 리스트는 상세표 기준 0개 유지.

5. 웹앱 실행 방법

python flask_app.py

접속 주소:
http://127.0.0.1:5000

6. 병합만 다시 생성하는 방법

python -c "from update import build_merged_list; build_merged_list()"

7. 현재 정상 기준값
- 공공기관: 355개
- 공직유관단체: 1557개
- 언론사: 약 23957개
- 공공기관 자회사: 651개
- 최종 merged_list.csv: 약 26206행

단, 언론사 원천 파일이 갱신되면 언론사 수와 최종 행 수는 변동 가능.

8. 언론사 다운로드 관련 주의
- update.py는 언론사 다운로드 실패 시 기존 MainServiceExcel.xls를 보존하도록 수정됨.
- 다운로드 성공 시 새 파일을 MainServiceExcel.xls로 표준화하고 중복 파일은 정리하는 것이 원칙.
- downloads/media_company 폴더에는 MainServiceExcel.xls 하나만 남는 상태가 이상적.

9. ALIO 자회사 추출 기준
- ALIO 타법인 투·출자 현황 보고서의 상세표 기준.
- 첫 번째 셀이 정확히 "자회사"인 행만 추출.
- 총괄표와 상세표가 다를 경우 상세표 기준 우선.

10. 화면 고지문
검색기 하단에는 다음 고지문을 표시함:

본 검색기는 ALIO, 인사혁신처 및 문화체육관광부 정기간행물 등록관리 시스템의 공개 데이터를 자동 수집·정리한 참고용 도구입니다. 원천 사이트의 게시 오류, 갱신 지연, 형식 변경, 중복 등록, 기관별 공시 오류 등으로 실제 현황과 차이가 발생할 수 있습니다. 대외 제출, 계약, 제재, 평가, 법적 판단 등 공식·확정적 용도로 사용하는 경우에는 원천 사이트의 최신 자료를 확인하시기 바랍니다.

11. 문제 발생 시 확인 순서
1) logs/run_log_*.txt 확인
2) downloads/media_company/MainServiceExcel.xls 존재 여부 확인
3) output/public_agency_subsidiaries.xlsx 행 수 확인
4) output/subsidiary_validation_report.xlsx 상태별 개수 확인
5) downloads/merged_list.csv category별 개수 확인
6) flask_app.py 실행 후 검색 테스트

12. 대표 검색 테스트어
- 중소기업은행
- 한국산업은행
- 한국자산관리공사
- 한국해양진흥공사
- 한국벤처투자
- 한국남부발전
- 코스포서비스
- 불교신문
- 레이디경향
- 인천항보안공사