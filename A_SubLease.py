"""
전대리스 엑셀 파일 처리 메인 모듈
- 파일 선택 및 케이스별 작업 수행
- 작업 완료 후 동일폴더에 '_전대반영후' 로 저장
    
모듈 구조:
- file_handler: 파일 선택, 로드, 저장
- excel_utils: 엑셀 작업 유틸리티 (시트, 테이블, 수식 등)
- case_processor: 케이스 판단 및 케이스별 처리
"""

# ============================================================================
# 모듈 Import
# ============================================================================

from pathlib import Path

from file_handler import (
    select_excel_file,
    load_workbook,
    save_workbook,
    generate_output_path
)

from case_processor import (
    determine_case,
    process_all_cases,
    process_case1,
    process_case2,
    process_case_default
)

from sheet_creator import (
    create_sheets_for_all_contracts,
    create_sublease_lease_data_sheet,
    create_lease_receivable_sheet,
    create_lease_deposit_sheet
)


# ============================================================================
# 메인 실행 함수
# ============================================================================

def main():
    """
    메인 실행 함수
    - 엑셀 파일 선택
    - 케이스별 작업 수행
    - 결과 파일 저장
    """
    # 1. 엑셀 파일 선택
    excel_file_path = select_excel_file()
    if not excel_file_path:
        print("파일 선택이 취소되었습니다.")
        return
    
    print(f"선택된 파일: {excel_file_path}")
    
    # 2. 엑셀 파일 로드
    try:
        wb = load_workbook(excel_file_path)
    except Exception as e:
        print(f"파일 로드 실패: {e}")
        return
    
    # 3. 케이스 판단 및 작업 수행
    data_manager = None
    try:
        case = determine_case(wb)
        print(f"판단된 케이스: {case}")
        
        # 케이스별 작업 수행
        if case == "케이스1":
            process_case1(wb)
        elif case == "케이스2":
            process_case2(wb)
        elif case == "default":
            process_case_default(wb)
        else:
            # 기본 처리 또는 모든 케이스 처리
            data_manager = process_all_cases(wb)
            
            # 시트 생성 작업
            if data_manager and len(data_manager.contract_data) > 0:
                create_sheets_for_all_contracts(wb, data_manager)
                
                # 전대_Lease_Data 시트 생성
                create_sublease_lease_data_sheet(wb)
                
                # 전대_Lease_Data 시트에 데이터 채우기
                from sheet_creator import fill_sublease_lease_data_sheet
                fill_sublease_lease_data_sheet(wb, data_manager)
                
                # 리스채권 시트 생성
                create_lease_receivable_sheet(wb)
                
                # 리스채권 시트에 데이터 채우기
                from sheet_creator import fill_lease_receivable_sheet
                fill_lease_receivable_sheet(wb, data_manager)
                
                # 임대보증금 시트 생성
                create_lease_deposit_sheet(wb)
                
                # 임대보증금 시트에 데이터 채우기
                from sheet_creator import fill_lease_deposit_sheet
                fill_lease_deposit_sheet(wb, data_manager)
                
                # 전대사용권자산 시트 생성
                from right_of_use_asset_sheet import create_right_of_use_asset_sheet
                create_right_of_use_asset_sheet(wb)
                
                # 전대사용권자산 시트에 데이터 추가
                from right_of_use_asset_sheet import fill_right_of_use_asset_sheet
                fill_right_of_use_asset_sheet(wb, data_manager)
            
    except Exception as e:
        print(f"작업 수행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        wb.close()
        return
    
    # 4. 결과 파일 저장
    output_path = generate_output_path(excel_file_path, avoid_overwrite=True)
    
    # 기존 파일이 있는지 확인하여 메시지 출력
    output_path_obj = Path(output_path)
    if output_path_obj.exists():
        print(f"기존 파일이 존재합니다. 새 파일로 저장합니다: {output_path_obj.name}")
    else:
        print(f"저장할 파일 경로: {output_path_obj.name}")
    
    try:
        save_workbook(wb, output_path)
    except Exception as e:
        print(f"파일 저장 실패: {e}")


# ============================================================================
# 실행
# ============================================================================

if __name__ == "__main__":
    main()

