# SCD Business Report

팀원별 개인 일일업무 엑셀을 취합하여 다음 문서를 자동 작성하는 Windows 프로그램입니다.

- 팀 일일업무
- 일일공수
- 주간공수
- 주간보고 리스트

## 사용 방법

1. Releases 또는 Actions의 빌드 결과에서 `팀업무_자동취합.exe`를 받습니다.
2. 팀원별 개인업무 엑셀 파일을 한 폴더에 모읍니다.
3. 프로그램에서 팀원업무 폴더, 대표님용 엑셀, 결과 저장 위치를 선택합니다.
4. `자동 취합 실행`을 누릅니다.

## EXE 빌드

GitHub의 **Actions → Build Windows EXE → Run workflow**를 실행하면 Windows 실행파일이 생성됩니다.
빌드가 완료되면 해당 실행 기록의 **Artifacts**에서 다운로드할 수 있습니다.
