# Word Compare Guide

이 문서는 이 저장소에서 작업할 때 기준이 되는 운영 문서입니다.  
기존의 짧은 요구사항 메모를 확장해, 현재 구조와 리팩터링 방향까지 함께 정리합니다.

## 목적

이 프로그램은 Microsoft Word의 `문서 비교` 기능을 이용해 두 문서를 비교하고:

- Word 비교 결과 문서를 저장하고
- 필요하면 차이점을 Excel 보고서로 정리해
- 사용자가 여러 쌍의 문서를 한 번에 처리할 수 있게 하는 데스크톱 앱입니다.

## 절대 바뀌면 안 되는 동작

아래 항목은 리팩터링 중에도 보존해야 합니다.

1. `listViewbefore`와 `listViewafter`의 같은 순번 파일끼리 비교한다.
2. 두 리스트의 파일 개수가 다르면 작업을 시작하지 않는다.
3. 사용자는 파일을 드래그 앤 드롭으로 추가할 수 있어야 한다.
4. 사용자는 리스트 안에서 순서를 바꿀 수 있어야 한다.
5. 선택된 항목은 `Delete` 키로 삭제할 수 있어야 한다.
6. 각 리스트 아이템은 `ItemIsDropEnabled`를 제거해 덮어쓰기식 드롭을 막아야 한다.
7. Word 비교 결과는 그대로 저장되어야 한다.
8. Excel 옵션이 켜져 있으면 `위치 / 수정 전 / 수정 후` 형식의 보고서를 만든다.

## 현재 구조

### Entry

- [main.py](/home/better0101/projects/word-compare/main.py)
  앱 진입점만 담당한다.

### UI

- [app/ui/main_window.py](/home/better0101/projects/word-compare/app/ui/main_window.py)
  메인 윈도우, 리스트 모델 연결, 사용자 이벤트 처리
- [main.ui](/home/better0101/projects/word-compare/main.ui)
  Qt Designer 원본
- [main_ui.py](/home/better0101/projects/word-compare/main_ui.py)
  생성 파일. 직접 수정하지 않는다.

### Services

- [app/services/settings_service.py](/home/better0101/projects/word-compare/app/services/settings_service.py)
  `QSettings` 저장/복원
- [app/services/word_session.py](/home/better0101/projects/word-compare/app/services/word_session.py)
  Word COM 세션 생성/종료
- [app/services/word_compare_service.py](/home/better0101/projects/word-compare/app/services/word_compare_service.py)
  문서 비교 작업 오케스트레이션
- [app/services/docx_extractor.py](/home/better0101/projects/word-compare/app/services/docx_extractor.py)
  Word 문서를 임시 docx로 저장한 뒤 `python-docx`로 문단/표를 추출

### Reports

- [app/reports/models.py](/home/better0101/projects/word-compare/app/reports/models.py)
  Excel 보고서 입력/계획 모델
- [app/reports/diff_engine.py](/home/better0101/projects/word-compare/app/reports/diff_engine.py)
  본문/표 비교를 위한 opcode 계산
- [app/reports/excel_writer.py](/home/better0101/projects/word-compare/app/reports/excel_writer.py)
  `xlsxwriter` 출력 담당
- [app/reports/excel_report_service.py](/home/better0101/projects/word-compare/app/reports/excel_report_service.py)
  추출 데이터와 writer를 연결
- [excel_generator.py](/home/better0101/projects/word-compare/excel_generator.py)
  호환성 유지용 얇은 래퍼

### Tools

- [tools/analyze_files.py](/home/better0101/projects/word-compare/tools/analyze_files.py)
- [tools/debug_pos.py](/home/better0101/projects/word-compare/tools/debug_pos.py)
- [tools/run_simulation.py](/home/better0101/projects/word-compare/tools/run_simulation.py)

루트의 동명 파일은 당분간 호환성 래퍼로만 유지한다.

## 리팩터링 원칙

1. 동작 보존이 구조 개선보다 우선이다.
2. UI 코드는 UI 이벤트와 모델 반영만 담당한다.
3. Word COM 직접 호출은 서비스 계층에만 둔다.
4. 문서 추출, diff 계산, Excel 출력은 분리한다.
5. 생성 파일(`main_ui.py`)은 수정하지 않는다.
6. 실험 스크립트와 배포 코드는 분리한다.

## 리팩터링 진행 상태

현재까지 완료:

- `main.py` 얇은 엔트리포인트화
- UI / Services / Reports 분리
- Excel 보고서의 계산 계층과 출력 계층 분리
- 작업 기록 문서 추가: [docs/refactoring-log.md](/home/better0101/projects/word-compare/docs/refactoring-log.md)

다음 후보:

- 루트 래퍼 스크립트 정리
- 테스트 보강
- 패키징 구조 정리
- 샘플/산출물 정리

## 수동 검증 체크리스트

작업 후 아래를 확인한다.

1. 앱 실행이 정상인지
2. before/after 드래그 앤 드롭이 되는지
3. 리스트 순서 변경이 되는지
4. `Delete` 삭제가 되는지
5. 파일 개수 불일치 시 시작이 막히는지
6. 비교 결과 docx가 저장되는지
7. Excel 옵션 켠 상태에서 xlsx가 저장되는지
