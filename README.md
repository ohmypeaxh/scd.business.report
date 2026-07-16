# SCD Business Report

팀원들이 작성한 여러 개인 일일업무 엑셀을 읽어 기존 전체 업무보고 양식에 자동 반영하는 Windows 프로그램입니다.

## 생성 내용

- `팀 일일업무`: 취합된 데이터 중 가장 최근 작업일의 업무를 기존 영역에 반영
- `일일공수`: 개인별 상세 업무 시간을 날짜별·업무별로 합산
- `주간공수`: 기존 수식과 서식을 유지하여 일일공수 반영 결과를 계산
- `자동취합_원본`: 취합 근거가 되는 모든 상세 업무 기록
- `주간보고 리스트`: 일자·이름·구분·업무·시간·업무 내용 목록

원본 전체용 엑셀은 직접 수정하지 않습니다. 결과는 사용자가 지정한 새 `.xlsx` 파일로 저장합니다.

## 사용 방법

1. 팀원별 개인 일일업무 `.xlsx` 파일을 하나의 폴더에 모읍니다.
2. 프로그램에서 **팀원업무 폴더**를 선택합니다.
3. 기존 양식이 들어 있는 **전체용 엑셀**을 선택합니다.
4. **결과 저장위치**를 선택합니다.
5. **자동취합 실행**을 누릅니다.

개인 파일에는 `개인일일업무` 시트가 있어야 합니다. 열려 있는 엑셀의 임시 파일(`~$...xlsx`)은 자동 제외됩니다.

## Windows EXE 받기

GitHub의 **Actions → Build Windows EXE**에서 성공한 실행 기록을 열고, **Artifacts**의 `SCD-Business-Report-Windows`를 내려받습니다. `main` 브랜치에 커밋할 때마다 테스트와 EXE 빌드가 자동 실행됩니다.

## 로컬 개발

Python 3.11 환경에서 실행합니다.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:PYTHONPATH = "src"
python -m unittest discover -s tests -v
python main.py
```

EXE 빌드:

```powershell
pyinstaller "SCD Business Report.spec" --noconfirm
```

결과 파일은 `dist/SCD Business Report.exe`에 생성됩니다.

## 처리 원칙

- 전체용 엑셀의 기존 시트, 수식, 병합 셀, 인쇄 설정과 서식을 유지합니다.
- 자동 생성하는 두 시트만 매 실행 시 최신 데이터로 다시 만듭니다.
- 일부 팀원 파일을 읽지 못해도 나머지는 계속 처리하고 완료 창에 경고를 표시합니다.
- 일일공수 시트에 없는 날짜는 기존 양식을 변경하지 않기 위해 자동으로 행을 삽입하지 않습니다. 해당 기록은 `자동취합_원본`과 `주간보고 리스트`에는 남습니다.
