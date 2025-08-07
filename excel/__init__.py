"""
Excel 명세서 생성 모듈

이 모듈은 데이터베이스 메타데이터를 기반으로 Excel 형태의 명세서를 생성합니다.

주요 클래스:
- DBSpecExcelGenerator: Excel 명세서 생성기

사용 예시:
    from excel import excel_generator
    
    excel_path = excel_generator.generate_excel(metadata, "명세서.xlsx")
"""

from .excel_generator import DBSpecExcelGenerator, excel_generator

__all__ = [
    'DBSpecExcelGenerator',
    'excel_generator'
]