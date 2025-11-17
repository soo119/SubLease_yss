"""
케이스별 처리 모듈
- 케이스 판단 로직
- 케이스별 작업 수행 함수들
"""

from excel_utils import get_sheet_by_name, get_table_data
from data_manager import DataManager


def determine_case(wb):
    """
    워크북을 분석하여 처리할 케이스 판단
    
    Args:
        wb: openpyxl Workbook 객체
        
    Returns:
        str: 케이스 식별자
    """
    # TODO: 케이스 판단 로직 구현
    # 예시:
    # - 시트 개수 확인
    # - 특정 시트 존재 여부 확인
    # - 특정 데이터 패턴 확인
    
    # 기본 케이스 판단 예시
    sheet_count = len(wb.sheetnames)
    
    if sheet_count == 0:
        return "empty"
    elif sheet_count == 1:
        return "single_sheet"
    else:
        return "multiple_sheets"


def process_all_cases(wb):
    """
    모든 케이스에 대한 기본 처리
    - Input_data 시트의 input_table에서 '전대' 포함 데이터 필터링
    - 선급/후급과 리스개시기준 조합으로 4가지 유형 분류
    - 계약번호를 키로 데이터 저장
    
    Args:
        wb: openpyxl Workbook 객체
        
    Returns:
        DataManager: 처리된 데이터를 담은 DataManager 객체
    """
    print("\n[전체 케이스 처리 시작]")
    
    # 1. Input_data 시트 확인
    input_sheet = get_sheet_by_name(wb, "Input_data")
    if input_sheet is None:
        raise ValueError("'Input_data' 시트를 찾을 수 없습니다.")
    
    print("✓ Input_data 시트 확인 완료")
    
    # 2. input_table 테이블에서 데이터 읽기
    try:
        table_data = get_table_data(input_sheet, "input_table")
        print(f"✓ input_table 데이터 읽기 완료: 총 {len(table_data)}행")
    except Exception as e:
        raise ValueError(f"input_table 데이터 읽기 실패: {e}")
    
    # 3. '전대' 포함 라인 필터링
    filtered_data = []
    lease_name_col = None
    
    # 리스명 컬럼 찾기 (대소문자 무시, 공백 무시)
    if table_data:
        for col_name in table_data[0].keys():
            if '리스명' in str(col_name).replace(' ', '').replace('_', ''):
                lease_name_col = col_name
                break
    
    if lease_name_col is None:
        raise ValueError("'리스명' 컬럼을 찾을 수 없습니다.")
    
    print(f"✓ 리스명 컬럼 확인: '{lease_name_col}'")
    
    for row in table_data:
        lease_name = str(row.get(lease_name_col, '') or '')
        if '전대' in lease_name:
            filtered_data.append(row)
    
    print(f"✓ '전대' 포함 필터링 완료: {len(filtered_data)}행")
    
    if len(filtered_data) == 0:
        print("경고: 필터링된 데이터가 없습니다.")
        return DataManager()
    
    # 4. DataManager 생성 및 데이터 추가
    data_manager = DataManager()
    
    # Ref no. 컬럼 찾기
    ref_no_col = None
    for col_name in table_data[0].keys():
        col_lower = str(col_name).lower().replace(' ', '').replace('_', '').replace('.', '')
        if 'refno' in col_lower or '계약번호' in str(col_name):
            ref_no_col = col_name
            break
    
    if ref_no_col is None:
        raise ValueError("'Ref no.' 또는 '계약번호' 컬럼을 찾을 수 없습니다.")
    
    print(f"✓ 계약번호 컬럼 확인: '{ref_no_col}'")
    
    # 각 행을 DataManager에 추가
    for row in filtered_data:
        contract_no = row.get(ref_no_col)
        if contract_no is None:
            print(f"경고: 계약번호가 없는 행을 건너뜁니다: {row}")
            continue
        
        data_manager.add_contract(contract_no, row)
    
    print(f"✓ 데이터 매니저에 추가 완료: {len(data_manager.contract_data)}개 계약")
    
    # 5. 디버그 출력
    data_manager.debug_print()
    
    return data_manager


def process_case1(wb):
    """
    케이스 1 처리 함수
    (다음 챗에서 구체적인 작업 내용 추가 예정)
    
    Args:
        wb: openpyxl Workbook 객체
    """
    print("케이스 1 처리 중...")
    # TODO: 케이스 1 작업 내용
    pass


def process_case2(wb):
    """
    케이스 2 처리 함수
    (다음 챗에서 구체적인 작업 내용 추가 예정)
    
    Args:
        wb: openpyxl Workbook 객체
    """
    print("케이스 2 처리 중...")
    # TODO: 케이스 2 작업 내용
    pass


def process_case_default(wb):
    """
    기본 케이스 처리 함수
    
    Args:
        wb: openpyxl Workbook 객체
    """
    print("기본 케이스 처리 중...")
    # TODO: 기본 케이스 작업 내용
    pass

