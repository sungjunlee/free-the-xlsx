import os
import json
import click
import xlwings as xw
import pandas as pd
from pathlib import Path
from typing import Optional, Literal

def excel_to_df(excel_file):
    """엑셀 파일을 DataFrame으로 변환"""
    try:
        app = xw.App(visible=False)
        wb = app.books.open(excel_file)
        
        result = []
        for sheet in wb.sheets:
            # 데이터를 판다스 DataFrame으로 변환
            raw_data = sheet.used_range.options(pd.DataFrame, index=False).value
            
            # 헤더와 데이터 분리
            headers = raw_data.iloc[0] if not raw_data.empty else pd.Series()
            data = raw_data.iloc[1:] if not raw_data.empty else pd.DataFrame()
            
            # 컬럼명 처리
            if headers.dtype == 'int64' or headers.isnull().any():
                # 숫자형이거나 빈 컬럼명인 경우 기본 이름 생성
                columns = [f'Column_{i}' for i in range(len(headers))]
            else:
                # 중복된 컬럼명에 번호 붙이기
                columns = []
                seen = {}
                for col in headers:
                    if pd.isna(col):
                        col = 'Unnamed'
                    col = str(col)
                    if col in seen:
                        seen[col] += 1
                        columns.append(f'{col}_{seen[col]}')
                    else:
                        seen[col] = 0
                        columns.append(col)
            
            # 새로운 DataFrame 생성
            df = pd.DataFrame(data.values, columns=columns)
            df['Sheet_Name'] = sheet.name
            result.append(df)
        
        wb.close()
        app.quit()
        
        if result:
            return pd.concat(result, ignore_index=True)
        return None
    
    except Exception as e:
        print(f"파일 처리 중 오류 발생: {excel_file}")
        print(f"오류 내용: {str(e)}")
        return None

def convert_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame의 데이터 타입을 문자열로 안전하게 변환"""
    df = df.copy()
    
    # 날짜/시간 데이터 처리
    for col in df.select_dtypes(include=['datetime64', 'datetimetz']).columns:
        df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # 숫자형 데이터 중 NaN 처리
    for col in df.select_dtypes(include=['float64', 'float32']).columns:
        df[col] = df[col].apply(lambda x: f"{x:.0f}" if pd.notna(x) and x.is_integer() else x)
    
    # None, NaN 등의 널 데이터 처리
    df = df.fillna('')
    
    # 불리언 데이터 처리
    for col in df.select_dtypes(include=['bool']).columns:
        df[col] = df[col].map({True: '예', False: '아니오'})
    
    return df

def save_as_csv(df: pd.DataFrame, output_file: Path) -> None:
    """DataFrame을 CSV로 저장"""
    df = convert_dtypes(df)
    
    # 시트별로 파일 저장
    for sheet_name, group in df.groupby('Sheet_Name'):
        group_without_sheet = group.drop('Sheet_Name', axis=1)
        sheet_file = output_file.parent / f"{output_file.stem}_{sheet_name}{output_file.suffix}"
        group_without_sheet.to_csv(sheet_file, index=False, encoding='utf-8-sig')
        print(f"시트 '{sheet_name}' 저장 완료: {sheet_file}")

def save_as_json(df: pd.DataFrame, output_file: Path) -> None:
    """DataFrame을 JSON으로 저장"""
    df = convert_dtypes(df)
    
    # 시트별로 파일 저장
    for sheet_name, group in df.groupby('Sheet_Name'):
        group_without_sheet = group.drop('Sheet_Name', axis=1)
        sheet_file = output_file.parent / f"{output_file.stem}_{sheet_name}{output_file.suffix}"
        
        result = group_without_sheet.to_dict(orient='records')
        with open(sheet_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"시트 '{sheet_name}' 저장 완료: {sheet_file}")

def save_as_markdown(df: pd.DataFrame, output_file: Path) -> None:
    """DataFrame을 Markdown 테이블로 저장"""
    df = convert_dtypes(df)
    
    # 시트별로 파일 저장
    for sheet_name, group in df.groupby('Sheet_Name'):
        group_without_sheet = group.drop('Sheet_Name', axis=1)
        sheet_file = output_file.parent / f"{output_file.stem}_{sheet_name}{output_file.suffix}"
        
        with open(sheet_file, 'w', encoding='utf-8') as f:
            f.write(f"# {sheet_name}\n\n")
            
            if group_without_sheet.empty:
                f.write("*데이터가 없습니다.*\n\n")
            else:
                try:
                    markdown_table = group_without_sheet.to_markdown(index=False, tablefmt="pipe")
                    f.write(f"{markdown_table}\n\n")
                except Exception as e:
                    f.write(f"*테이블 변환 중 오류 발생: {str(e)}*\n\n")
                    f.write("```\n")
                    f.write(group_without_sheet.to_string(index=False))
                    f.write("\n```\n\n")
        print(f"시트 '{sheet_name}' 저장 완료: {sheet_file}")

@click.command()
@click.argument('input_dir', type=click.Path(exists=True))
@click.argument('output_dir', type=click.Path())
@click.option('--format', '-f', 
              type=click.Choice(['csv', 'json', 'markdown', 'all']),
              default='csv',
              help='출력 파일 형식 (기본값: csv)')
@click.option('--combine/--no-combine', 
              default=False,
              help='모든 시트를 하나의 파일로 합칠지 여부 (기본값: False)')
def process_excel_files(input_dir, output_dir, format, combine):
    """
    지정된 디렉토리의 모든 엑셀 파일을 처리하여 지정된 형식으로 변환합니다.
    
    FORMAT 옵션:
    - csv: CSV 형식으로 저장 (기본값)
    - json: JSON 형식으로 저장
    - markdown: Markdown 테이블 형식으로 저장
    - all: 모든 형식으로 저장
    
    COMBINE 옵션:
    --combine: 모든 시트를 하나의 파일로 저장
    --no-combine: 시트별로 별도의 파일로 저장 (기본값)
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    excel_files = list(input_path.glob('*.xlsx')) + list(input_path.glob('*.xls'))
    
    if not excel_files:
        print("처리할 엑셀 파일을 찾을 수 없습니다.")
        return
    
    # 저장 함수 매핑
    save_functions = {
        'csv': (save_as_csv, '.csv'),
        'json': (save_as_json, '.json'),
        'markdown': (save_as_markdown, '.md')
    }
    
    formats_to_process = list(save_functions.keys()) if format == 'all' else [format]
    
    for excel_file in excel_files:
        print(f"처리 중: {excel_file.name}")
        
        df = excel_to_df(str(excel_file))
        if df is not None:
            for fmt in formats_to_process:
                save_func, extension = save_functions[fmt]
                output_file = output_path / f"{excel_file.stem}{extension}"
                save_func(df, output_file)
                print(f"변환 완료: {output_file}")

if __name__ == '__main__':
    process_excel_files() 