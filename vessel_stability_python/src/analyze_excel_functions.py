"""
Excel 파일의 모든 시트에서 사용된 함수를 분석하고 Python으로 구현
"""

import pandas as pd
import xlrd
from pathlib import Path
import re
from collections import defaultdict

def analyze_excel_functions(file_path: str):
    """Excel 파일의 모든 시트에서 함수를 분석"""
    print("=" * 60)
    print("📊 Excel 함수 분석")
    print("=" * 60)
    
    # .xls 파일 읽기
    xls_file = xlrd.open_workbook(file_path, on_demand=True)
    
    all_functions = defaultdict(list)
    sheet_data = {}
    sheet_names = xls_file.sheet_names()
    
    print(f"\n📄 파일: {Path(file_path).name}")
    print(f"📋 총 시트 수: {len(sheet_names)}\n")
    
    for sheet_name in sheet_names:
        print(f"🔍 시트 분석: {sheet_name}")
        try:
            sheet = xls_file.sheet_by_name(sheet_name)
            
            # DataFrame으로 읽기 (수식이 아닌 값만)
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # xlrd로 수식 추출
            formulas = []
            for row_idx in range(min(sheet.nrows, 100)):  # 처음 100행만
                for col_idx in range(min(sheet.ncols, 50)):  # 처음 50열만
                    try:
                        cell = sheet.cell(row_idx, col_idx)
                        if cell.ctype == xlrd.XL_CELL_FORMULA:
                            formula = xlrd.formula.xls_formula(formula_str=cell.value, book=xls_file)
                            formulas.append({
                                'row': row_idx + 1,
                                'col': col_idx + 1,
                                'formula': formula
                            })
                            
                            # 함수명 추출
                            func_matches = re.findall(r'([A-Z][A-Z0-9_]*)\s*\(', formula)
                            for func in func_matches:
                                all_functions[func].append({
                                    'sheet': sheet_name,
                                    'cell': f"{chr(64+col_idx+1)}{row_idx+1}",
                                    'formula': formula
                                })
                    except:
                        pass
            
            sheet_data[sheet_name] = {
                'rows': sheet.nrows,
                'cols': sheet.ncols,
                'formulas_count': len(formulas),
                'sample_formulas': formulas[:5]  # 처음 5개만
            }
            
            print(f"  ✓ {sheet.nrows}행 x {sheet.ncols}열, 수식 {len(formulas)}개")
            
        except Exception as e:
            print(f"  ⚠️  오류: {e}")
    
    xls_file.release_resources()
    
    return all_functions, sheet_data

def extract_sample_data(file_path: str, sheet_name: str):
    """시트의 샘플 데이터 추출"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
        return df
    except:
        return None

if __name__ == "__main__":
    file_path = "1.Vessel Stability Booklet.xls"
    
    functions, sheet_data = analyze_excel_functions(file_path)
    
    print("\n" + "=" * 60)
    print("📊 발견된 Excel 함수")
    print("=" * 60)
    
    for func_name, occurrences in sorted(functions.items()):
        print(f"\n{func_name} ({len(occurrences)}회 사용):")
        for occ in occurrences[:3]:  # 처음 3개만
            print(f"  - {occ['sheet']} / {occ['cell']}: {occ['formula'][:80]}")
    
    print("\n" + "=" * 60)
    print("📋 시트별 요약")
    print("=" * 60)
    
    for sheet_name, data in sheet_data.items():
        print(f"\n{sheet_name}:")
        print(f"  크기: {data['rows']}행 x {data['cols']}열")
        print(f"  수식: {data['formulas_count']}개")
        if data['sample_formulas']:
            print(f"  샘플 수식:")
            for f in data['sample_formulas'][:2]:
                print(f"    {f['formula'][:60]}...")

