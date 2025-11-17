"""
데이터 관리 모듈
- 필터링된 데이터 저장 및 관리
- 계약번호를 키로 하는 데이터 구조 관리
"""


class DataManager:
    """
    전대리스 데이터를 관리하는 클래스
    - 계약번호를 키로 데이터 저장
    - 유형별 데이터 분류
    """
    
    def __init__(self):
        """데이터 매니저 초기화"""
        # 계약번호를 키로 하는 전체 데이터 딕셔너리
        self.contract_data = {}
        
        # 유형별로 분류된 계약번호 리스트
        self.type_선급_월초 = []
        self.type_선급_월말 = []
        self.type_후급_월초 = []
        self.type_후급_월말 = []
    
    def add_contract(self, contract_no, row_data):
        """
        계약번호와 데이터를 추가하고 유형별로 분류
        
        Args:
            contract_no: 계약번호 (Ref no.)
            row_data: 해당 행의 모든 데이터 (딕셔너리)
        """
        # 전체 데이터에 추가
        self.contract_data[contract_no] = row_data
        
        # 선급/후급과 리스개시기준 추출
        payment_type = str(row_data.get('선급/후급', '')).strip()
        start_basis = str(row_data.get('리스개시기준', '')).strip()
        
        # 유형별 분류
        if payment_type == '선급' and start_basis == '월초':
            self.type_선급_월초.append(contract_no)
        elif payment_type == '선급' and start_basis == '월말':
            self.type_선급_월말.append(contract_no)
        elif payment_type == '후급' and start_basis == '월초':
            self.type_후급_월초.append(contract_no)
        elif payment_type == '후급' and start_basis == '월말':
            self.type_후급_월말.append(contract_no)
    
    def get_contract_data(self, contract_no):
        """
        계약번호로 데이터 가져오기
        
        Args:
            contract_no: 계약번호
            
        Returns:
            dict: 계약 데이터, 없으면 None
        """
        return self.contract_data.get(contract_no)
    
    def get_contracts_by_type(self, payment_type, start_basis):
        """
        유형별 계약번호 리스트 가져오기
        
        Args:
            payment_type: '선급' 또는 '후급'
            start_basis: '월초' 또는 '월말'
            
        Returns:
            list: 계약번호 리스트
        """
        if payment_type == '선급' and start_basis == '월초':
            return self.type_선급_월초
        elif payment_type == '선급' and start_basis == '월말':
            return self.type_선급_월말
        elif payment_type == '후급' and start_basis == '월초':
            return self.type_후급_월초
        elif payment_type == '후급' and start_basis == '월말':
            return self.type_후급_월말
        return []
    
    def get_all_types_summary(self):
        """
        모든 유형별 요약 정보 반환
        
        Returns:
            dict: 유형별 계약번호 리스트와 개수
        """
        return {
            '선급_월초': {
                'contracts': self.type_선급_월초,
                'count': len(self.type_선급_월초)
            },
            '선급_월말': {
                'contracts': self.type_선급_월말,
                'count': len(self.type_선급_월말)
            },
            '후급_월초': {
                'contracts': self.type_후급_월초,
                'count': len(self.type_후급_월초)
            },
            '후급_월말': {
                'contracts': self.type_후급_월말,
                'count': len(self.type_후급_월말)
            }
        }
    
    def debug_print(self):
        """디버그용 전체 데이터 출력"""
        print("\n" + "="*80)
        print("데이터 매니저 디버그 출력")
        print("="*80)
        
        print(f"\n전체 계약 개수: {len(self.contract_data)}")
        
        # 유형별 요약
        summary = self.get_all_types_summary()
        print("\n[유형별 분류]")
        for type_name, info in summary.items():
            print(f"  {type_name}: {info['count']}개")
            if info['count'] > 0:
                print(f"    계약번호: {info['contracts']}")
        
        # 각 계약별 상세 정보
        print("\n[계약별 상세 정보]")
        for contract_no, data in self.contract_data.items():
            print(f"\n  계약번호: {contract_no}")
            for key, value in data.items():
                print(f"    {key}: {value}")
        
        print("\n" + "="*80 + "\n")

