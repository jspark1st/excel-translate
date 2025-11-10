"""
엑셀 파일 번역 프로그램
구글 번역 API를 사용하여 엑셀 파일의 내용을 한국어로 번역합니다.
"""

import pandas as pd
import os
import sys
from googletrans import Translator
from typing import Union
import time


class ExcelTranslator:
    """엑셀 파일 번역 클래스"""
    
    def __init__(self):
        self.translator = Translator()
        self.translated_count = 0
        self.error_count = 0
    
    def translate_text(self, text: Union[str, float, int]) -> str:
        """
        텍스트를 한국어로 번역
        
        Args:
            text: 번역할 텍스트
            
        Returns:
            번역된 텍스트 (한국어)
        """
        # 숫자나 None 값은 그대로 반환
        if pd.isna(text) or text == '':
            return text
        
        # 숫자 타입은 문자열로 변환하지 않고 그대로 반환
        if isinstance(text, (int, float)):
            return text
        
        text_str = str(text).strip()
        
        # 빈 문자열이면 그대로 반환
        if not text_str:
            return text
        
        try:
            # 구글 번역 API 호출
            result = self.translator.translate(text_str, dest='ko')
            self.translated_count += 1
            
            # API 제한을 고려한 딜레이
            time.sleep(0.1)
            
            return result.text
        except Exception as e:
            self.error_count += 1
            print(f"번역 오류 (원문: {text_str[:50]}...): {str(e)}")
            return text_str  # 오류 발생 시 원문 반환
    
    def translate_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        DataFrame의 모든 셀을 번역
        
        Args:
            df: 번역할 DataFrame
            
        Returns:
            번역된 DataFrame
        """
        translated_df = df.copy()
        
        print("번역 중...")
        for col in df.columns:
            print(f"  컬럼 '{col}' 번역 중...")
            translated_df[col] = df[col].apply(self.translate_text)
        
        return translated_df
    
    def translate_excel(self, input_file: str, output_file: str = None):
        """
        엑셀 파일을 번역하여 새 파일로 저장
        
        Args:
            input_file: 입력 엑셀 파일 경로
            output_file: 출력 엑셀 파일 경로 (None이면 자동 생성)
        """
        # 파일 존재 확인
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {input_file}")
        
        # 출력 파일명 자동 생성
        if output_file is None:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_translated.xlsx"
        
        print(f"입력 파일: {input_file}")
        print(f"출력 파일: {output_file}")
        print("-" * 50)
        
        # 엑셀 파일 읽기 (모든 시트 포함)
        excel_file = pd.ExcelFile(input_file)
        sheet_names = excel_file.sheet_names
        
        print(f"총 {len(sheet_names)}개의 시트를 발견했습니다.")
        
        # ExcelWriter로 모든 시트를 번역하여 저장
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                print(f"\n시트 '{sheet_name}' 처리 중...")
                
                # 시트 읽기
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                print(f"  행 수: {len(df)}, 열 수: {len(df.columns)}")
                
                # 번역
                translated_df = self.translate_dataframe(df)
                
                # 번역된 시트 저장
                translated_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  시트 '{sheet_name}' 번역 완료")
        
        print("\n" + "=" * 50)
        print("번역 완료!")
        print(f"총 번역된 셀 수: {self.translated_count}")
        if self.error_count > 0:
            print(f"번역 오류 수: {self.error_count}")
        print(f"출력 파일: {output_file}")


def main():
    """메인 함수"""
    if len(sys.argv) < 2:
        print("사용법: python translate_excel.py <엑셀파일경로> [출력파일경로]")
        print("\n예시:")
        print("  python translate_excel.py input.xlsx")
        print("  python translate_excel.py input.xlsx output.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        translator = ExcelTranslator()
        translator.translate_excel(input_file, output_file)
    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

