# 전대리스 엑셀 파일 처리 프로그램

전대리스 엑셀 파일을 자동으로 처리하고 필요한 시트를 생성하는 Python 프로그램입니다.

## 주요 기능

- 엑셀 파일에서 전대 관련 데이터 자동 필터링
- 계약별 상각표 시트 생성
- 전대_Lease_Data 시트 생성 및 데이터 입력
- 리스채권 시트 생성 및 데이터 입력
- 임대보증금 시트 생성 및 데이터 입력
- 전대사용권자산 시트 생성 및 데이터 입력
- 케이스별 자동 판단 및 처리

## 요구사항

- Python 3.7 이상
- openpyxl

## 설치 방법

### 1. 저장소 클론 또는 다운로드

```bash
git clone <repository-url>
cd 전대리스코드
```

### 2. 가상환경 생성 (권장)

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. 필요한 패키지 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### 기본 실행

```bash
python A_SubLease.py
```

1. 프로그램 실행 시 파일 선택 다이얼로그가 열립니다.
2. 처리할 엑셀 파일(.xlsx, .xlsm, .xls)을 선택합니다.
3. 프로그램이 자동으로:
   - Input_data 시트에서 '전대' 포함 데이터 필터링
   - 케이스 판단 및 처리
   - 필요한 시트 생성 및 데이터 입력
4. 처리 완료 후 동일 폴더에 `_전대반영후` 접미사가 추가된 새 파일로 저장됩니다.

### 출력 파일

- 원본 파일명: `example.xlsx`
- 출력 파일명: `example_전대반영후.xlsx`
- 기존 파일이 있으면: `example_전대반영후_1.xlsx` (번호 자동 증가)

## 프로젝트 구조

```
전대리스코드/
├── A_SubLease.py              # 메인 실행 파일
├── file_handler.py            # 파일 선택, 로드, 저장
├── excel_utils.py             # 엑셀 작업 유틸리티
├── case_processor.py          # 케이스 판단 및 처리
├── data_manager.py            # 데이터 관리
├── sheet_creator.py           # 시트 생성 모듈
├── amortization_sheet.py      # 상각표 시트 생성
├── lease_data_sheet.py        # 전대_Lease_Data 시트
├── lease_receivable_sheet.py  # 리스채권 시트
├── lease_deposit_sheet.py     # 임대보증금 시트
├── right_of_use_asset_sheet.py # 전대사용권자산 시트
├── requirements.txt           # 필요한 패키지 목록
└── README.md                  # 이 파일
```

## 입력 파일 형식

프로그램은 다음 형식의 엑셀 파일을 기대합니다:

- **Input_data** 시트가 있어야 합니다.
- **input_table** 테이블이 있어야 합니다.
- 테이블에는 다음 컬럼이 포함되어야 합니다:
  - 리스명 (또는 유사한 이름)
  - Ref no. (또는 계약번호)
  - 선급/후급
  - 리스개시기준

## 처리 로직

1. **데이터 필터링**: Input_data 시트의 input_table에서 '전대'가 포함된 리스명 데이터만 추출
2. **유형 분류**: 선급/후급과 리스개시기준(월초/월말) 조합으로 4가지 유형으로 분류
   - 선급_월초
   - 선급_월말
   - 후급_월초
   - 후급_월말
3. **시트 생성**: 각 계약별 상각표 시트 및 통합 시트 생성
4. **데이터 입력**: 필터링된 데이터를 각 시트에 자동 입력

## 문제 해결

### 파일 로드 실패
- 파일이 다른 프로그램에서 열려있는지 확인
- 파일 경로에 특수문자가 없는지 확인
- 파일 형식(.xlsx, .xlsm, .xls) 확인

### 시트를 찾을 수 없음
- Input_data 시트가 존재하는지 확인
- input_table 테이블이 존재하는지 확인

### 패키지 설치 오류
- Python 버전 확인 (3.7 이상 필요)
- 가상환경이 활성화되어 있는지 확인
- 인터넷 연결 확인

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.

## 문의

문제가 발생하거나 개선 사항이 있으면 이슈를 등록해주세요.

