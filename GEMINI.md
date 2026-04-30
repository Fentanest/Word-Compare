# Word Compare Guide

이 문서는 이 저장소에서 작업할 때 기준이 되는 운영 문서입니다.  
현재 프로젝트 구조, 보존해야 할 동작, 각 모듈의 책임을 최신 상태로 정리합니다.

## 목적

이 프로그램은 Microsoft Word의 `문서 비교` 기능을 이용해 두 문서를 비교하고:

- Word 비교 결과 문서를 저장하고
- 필요하면 차이점을 Excel 보고서로 정리하고
- 여러 쌍의 문서를 한 번에 처리할 수 있게 하는 데스크톱 앱입니다.

## 절대 바뀌면 안 되는 동작

1. `listViewbefore`와 `listViewafter`의 같은 순번 파일끼리 비교한다.
2. 두 리스트의 파일 개수가 다르면 작업을 시작하지 않는다.
3. 사용자는 파일을 드래그 앤 드롭으로 추가할 수 있어야 한다.
4. 사용자는 리스트 안에서 순서를 바꿀 수 있어야 한다.
5. 선택된 항목은 `Delete` 키로 삭제할 수 있어야 한다.
6. 각 리스트 아이템은 `ItemIsDropEnabled`를 제거해 덮어쓰기식 드롭을 막아야 한다.
7. Word 비교 결과는 그대로 저장되어야 한다.
8. Excel 옵션이 켜져 있으면 `위치 / 수정 전 / 수정 후` 형식의 보고서를 만든다.
9. native Rust 추출기가 없어도 Python fallback 으로 정상 동작해야 한다.

## 현재 구조

### Entry

- [main.py](/home/better0101/projects/word-compare/main.py)
  앱 진입점만 담당한다.

### UI

- [app/ui/main_window.py](/home/better0101/projects/word-compare/app/ui/main_window.py)
  메인 윈도우, 버튼 연결, 로그 출력, 비교 시작 흐름 연결
- [app/ui/file_list_manager.py](/home/better0101/projects/word-compare/app/ui/file_list_manager.py)
  리스트 아이템 생성, 정렬, before/after 파일쌍 조립
- [main.ui](/home/better0101/projects/word-compare/main.ui)
  Qt Designer 원본
- [main_ui.py](/home/better0101/projects/word-compare/main_ui.py)
  생성 파일. 직접 수정하지 않는다.

### Domain / Shared Models

- [app/models.py](/home/better0101/projects/word-compare/app/models.py)
  `AppSettings`, `FilePair`, `CompareOptions`, `CompareResult`
  `ParagraphData`, `RunData`, `TableCellData`, `ExtractedDocument`
- [app/resources.py](/home/better0101/projects/word-compare/app/resources.py)
  번들 환경과 개발 환경을 모두 고려한 리소스 경로 해석

### Services

- [app/services/settings_service.py](/home/better0101/projects/word-compare/app/services/settings_service.py)
  `QSettings` 저장/복원
- [app/services/word_session.py](/home/better0101/projects/word-compare/app/services/word_session.py)
  Word COM 세션 생성/종료, 숨김 상태 유지
- [app/services/word_compare_service.py](/home/better0101/projects/word-compare/app/services/word_compare_service.py)
  문서 열기, revision 정리, Word 비교, 결과 저장, Excel 보고서 호출
- [app/services/docx_extractor.py](/home/better0101/projects/word-compare/app/services/docx_extractor.py)
  Word 문서를 임시 `.docx`로 저장하고 구조화 데이터로 추출
- [app/services/native_docx_extractor.py](/home/better0101/projects/word-compare/app/services/native_docx_extractor.py)
  Rust CLI 추출기 탐색, 실행, JSON 결과 복원

### Reports

- [app/reports/models.py](/home/better0101/projects/word-compare/app/reports/models.py)
  `ExcelReportInput`, `ExcelDiffPlan`, `TableDiffPlan`
- [app/reports/diff_engine.py](/home/better0101/projects/word-compare/app/reports/diff_engine.py)
  본문/표 정렬 및 opcode 계산
  표는 전체 행/열 시그니처를 이용하고, 정렬용 시그니처와 실제 변경 시그니처를 분리한다.
- [app/reports/excel_writer.py](/home/better0101/projects/word-compare/app/reports/excel_writer.py)
  `xlsxwriter`로 메인 시트와 표 시트 출력
- [app/reports/excel_report_service.py](/home/better0101/projects/word-compare/app/reports/excel_report_service.py)
  추출기, diff engine, writer를 연결
- [excel_generator.py](/home/better0101/projects/word-compare/excel_generator.py)
  기존 호출 경로와의 호환성 래퍼

### Native

- [native/docx-structure-extractor/Cargo.toml](/home/better0101/projects/word-compare/native/docx-structure-extractor/Cargo.toml)
- [native/docx-structure-extractor/src/main.rs](/home/better0101/projects/word-compare/native/docx-structure-extractor/src/main.rs)

Rust 추출기는 DOCX 패키지 XML을 직접 읽어:

- 본문 문단
- 표
- 머리말/꼬리말
- 각주/미주
- 주석
- revision
- section/page 설정
- drawing/textbox 관련 메타데이터

를 구조화 JSON으로 내보낸다.

### Tools

- [tools/analyze_files.py](/home/better0101/projects/word-compare/tools/analyze_files.py)
- [tools/debug_pos.py](/home/better0101/projects/word-compare/tools/debug_pos.py)
- [tools/run_simulation.py](/home/better0101/projects/word-compare/tools/run_simulation.py)

루트의 동명 파일은 당분간 호환성 래퍼로만 유지한다.

### Tests

- [tests/test_diff_engine.py](/home/better0101/projects/word-compare/tests/test_diff_engine.py)
  문단/표 정렬, 병합/너비/행높이/숫자 변화 회귀 테스트
- [tests/test_docx_metadata_extractor.py](/home/better0101/projects/word-compare/tests/test_docx_metadata_extractor.py)
  header/footer, notes, comments, revision, section, shape 메타데이터 추출 테스트
- [tests/test_excel_report_pipeline.py](/home/better0101/projects/word-compare/tests/test_excel_report_pipeline.py)
  xlsx 시트/문자열 출력 테스트
- [tests/test_file_list_manager.py](/home/better0101/projects/word-compare/tests/test_file_list_manager.py)
  UI 파일 목록 조작 테스트
- [tests/test_native_docx_extractor.py](/home/better0101/projects/word-compare/tests/test_native_docx_extractor.py)
  Rust 추출기 탐색/실행/복원 테스트
- [tests/test_native_extractor_binary.py](/home/better0101/projects/word-compare/tests/test_native_extractor_binary.py)
  실제 빌드된 native 바이너리 스모크 테스트
- [tests/test_word_compare_service.py](/home/better0101/projects/word-compare/tests/test_word_compare_service.py)
  Excel 보고서용 원본 재열기, 종료 시 close 예외 무시 테스트
- [tests/test_word_session.py](/home/better0101/projects/word-compare/tests/test_word_session.py)
  Word 세션 종료 예외 무시 테스트

### Docs

- [docs/refactoring-log.md](/home/better0101/projects/word-compare/docs/refactoring-log.md)
  리팩터링 배치 기록

## 핵심 동작 메모

### Word 비교 흐름

1. before/after 문서를 Word COM으로 연다.
2. 비교용 문서에서는 revision 을 받아들여 비교 전 정리한다.
3. Word `CompareDocuments`로 결과 문서를 만든다.
4. 결과를 `비교_결과_...docx`로 저장한다.
5. Excel 옵션이 켜져 있으면 원본 before/after 문서를 다시 열어 추출에 사용한다.

비교용 문서와 Excel 추출용 문서를 분리하는 이유는, 주석/변경 추적/revision 메타데이터를 Excel 보고서에서 보존하기 위해서다.

### Excel 보고서 흐름

1. `DocxExtractor`가 Word 문서를 임시 `.docx`로 저장한다.
2. native Rust 추출기를 우선 시도한다.
3. 실패하면 `python-docx + zip/xml` 경로로 fallback 한다.
4. `ExcelDiffEngine`이 본문/표 diff 계획을 만든다.
5. `ExcelReportWriter`가 `변경 내용(일반)` 시트와 `표 n` 시트를 출력한다.

### 표 비교 메모

- 표는 단순히 같은 row index끼리 비교하지 않는다.
- 행/열 정렬은 전체 행/열 시그니처를 이용한다.
- 정렬용 시그니처에서는 숫자만 바뀐 셀을 완화해서, 같은 항목이 아래로 밀린 경우를 더 잘 맞춘다.
- 실제 셀 변경 여부 판단은 full signature로 해서 숫자/서식/병합/너비 변화는 그대로 변경으로 표시한다.

## 리팩터링 원칙

1. 동작 보존이 구조 개선보다 우선이다.
2. UI 코드는 UI 이벤트와 모델 반영만 담당한다.
3. Word COM 직접 호출은 서비스 계층에만 둔다.
4. 문서 추출, diff 계산, Excel 출력은 분리한다.
5. 생성 파일(`main_ui.py`)은 수정하지 않는다.
6. 실험 스크립트와 배포 코드는 분리한다.
7. 성능 로그는 사용자가 병목을 파악할 수 있을 만큼만 남긴다.

## 수동 검증 체크리스트

1. 앱 실행이 정상인지
2. before/after 드래그 앤 드롭이 되는지
3. 리스트 순서 변경이 되는지
4. `Delete` 삭제가 되는지
5. 파일 개수 불일치 시 시작이 막히는지
6. 비교 결과 docx가 저장되는지
7. Excel 옵션 켠 상태에서 xlsx가 저장되는지
8. 표에서 행/열 밀림이 삭제+재삽입이 아니라 같은 항목 변경으로 보이는지
9. native 추출기가 있을 때 자동 사용되는지
10. native 추출기가 없어도 Python fallback 으로 정상 동작하는지
