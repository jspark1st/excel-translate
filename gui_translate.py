"""
엑셀 파일 번역 GUI 프로그램
구글 번역 API를 사용하여 엑셀 파일의 내용을 한국어로 번역합니다.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
from translate_excel import ExcelTranslator


class ExcelTranslateGUI:
    """엑셀 번역 GUI 클래스"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 파일 번역 프로그램")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # 변수 초기화
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.is_translating = False
        
        # GUI 구성
        self.create_widgets()
        
        # 중앙 정렬
        self.center_window()
    
    def center_window(self):
        """창을 화면 중앙에 배치"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """GUI 위젯 생성"""
        # 제목
        title_label = tk.Label(
            self.root,
            text="엑셀 파일 번역 프로그램",
            font=("맑은 고딕", 18, "bold"),
            pady=20
        )
        title_label.pack()
        
        # 입력 파일 선택 프레임
        input_frame = tk.Frame(self.root, pady=10)
        input_frame.pack(fill=tk.X, padx=20)
        
        tk.Label(
            input_frame,
            text="입력 파일:",
            font=("맑은 고딕", 10),
            width=10,
            anchor="w"
        ).pack(side=tk.LEFT)
        
        input_entry = tk.Entry(
            input_frame,
            textvariable=self.input_file_path,
            font=("맑은 고딕", 9),
            width=50
        )
        input_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        input_btn = tk.Button(
            input_frame,
            text="파일 선택",
            command=self.select_input_file,
            font=("맑은 고딕", 9),
            bg="#4CAF50",
            fg="white",
            relief=tk.RAISED,
            padx=10
        )
        input_btn.pack(side=tk.LEFT)
        
        # 출력 파일 선택 프레임
        output_frame = tk.Frame(self.root, pady=10)
        output_frame.pack(fill=tk.X, padx=20)
        
        tk.Label(
            output_frame,
            text="출력 파일:",
            font=("맑은 고딕", 10),
            width=10,
            anchor="w"
        ).pack(side=tk.LEFT)
        
        output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_file_path,
            font=("맑은 고딕", 9),
            width=50
        )
        output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        output_btn = tk.Button(
            output_frame,
            text="경로 선택",
            command=self.select_output_file,
            font=("맑은 고딕", 9),
            bg="#2196F3",
            fg="white",
            relief=tk.RAISED,
            padx=10
        )
        output_btn.pack(side=tk.LEFT)
        
        # 안내 텍스트
        info_label = tk.Label(
            self.root,
            text="※ 출력 파일을 선택하지 않으면 입력 파일과 같은 위치에 '_translated'가 추가된 파일로 저장됩니다.",
            font=("맑은 고딕", 8),
            fg="gray",
            pady=5
        )
        info_label.pack()
        
        # 번역 버튼 프레임
        button_frame = tk.Frame(self.root, pady=20)
        button_frame.pack()
        
        self.translate_btn = tk.Button(
            button_frame,
            text="번역 시작",
            command=self.start_translation,
            font=("맑은 고딕", 12, "bold"),
            bg="#FF9800",
            fg="white",
            relief=tk.RAISED,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        self.translate_btn.pack()
        
        # 진행 상황 표시 영역
        progress_label = tk.Label(
            self.root,
            text="진행 상황:",
            font=("맑은 고딕", 10, "bold"),
            anchor="w"
        )
        progress_label.pack(anchor="w", padx=20, pady=(10, 5))
        
        # 스크롤 가능한 텍스트 영역
        self.progress_text = scrolledtext.ScrolledText(
            self.root,
            height=15,
            font=("맑은 고딕", 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg="#f5f5f5"
        )
        self.progress_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
    
    def select_input_file(self):
        """입력 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.input_file_path.set(file_path)
            # 출력 파일 경로 자동 설정
            if not self.output_file_path.get():
                base_name = os.path.splitext(file_path)[0]
                self.output_file_path.set(f"{base_name}_translated.xlsx")
            self.log_message(f"입력 파일 선택: {os.path.basename(file_path)}")
    
    def select_output_file(self):
        """출력 파일 경로 선택"""
        file_path = filedialog.asksaveasfilename(
            title="번역된 파일 저장 위치",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.output_file_path.set(file_path)
            self.log_message(f"출력 파일 경로 설정: {os.path.basename(file_path)}")
    
    def log_message(self, message):
        """진행 상황 메시지 추가"""
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.progress_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def clear_log(self):
        """로그 영역 초기화"""
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.delete(1.0, tk.END)
        self.progress_text.config(state=tk.DISABLED)
    
    def start_translation(self):
        """번역 시작"""
        # 입력 파일 확인
        input_file = self.input_file_path.get().strip()
        if not input_file:
            messagebox.showwarning("경고", "입력 파일을 선택해주세요.")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("오류", "입력 파일을 찾을 수 없습니다.")
            return
        
        # 이미 번역 중이면 중단
        if self.is_translating:
            messagebox.showinfo("알림", "이미 번역이 진행 중입니다.")
            return
        
        # 출력 파일 경로 설정
        output_file = self.output_file_path.get().strip()
        if not output_file:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_translated.xlsx"
            self.output_file_path.set(output_file)
        
        # 번역 시작
        self.is_translating = True
        self.translate_btn.config(state=tk.DISABLED, text="번역 중...")
        self.clear_log()
        
        # 별도 스레드에서 번역 실행 (UI가 멈추지 않도록)
        thread = threading.Thread(
            target=self.translate_in_thread,
            args=(input_file, output_file),
            daemon=True
        )
        thread.start()
    
    def translate_in_thread(self, input_file, output_file):
        """별도 스레드에서 번역 실행"""
        try:
            self.log_message("=" * 50)
            self.log_message("번역 시작")
            self.log_message(f"입력 파일: {os.path.basename(input_file)}")
            self.log_message(f"출력 파일: {os.path.basename(output_file)}")
            self.log_message("-" * 50)
            
            # 번역기 초기화 및 번역 실행
            translator = ExcelTranslator()
            
            # 엑셀 파일 정보 읽기
            import pandas as pd
            excel_file = pd.ExcelFile(input_file)
            sheet_names = excel_file.sheet_names
            
            self.log_message(f"총 {len(sheet_names)}개의 시트를 발견했습니다.")
            self.log_message("")
            
            # 각 시트 번역
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for idx, sheet_name in enumerate(sheet_names, 1):
                    self.log_message(f"[{idx}/{len(sheet_names)}] 시트 '{sheet_name}' 처리 중...")
                    
                    # 시트 읽기
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    self.log_message(f"  행 수: {len(df)}, 열 수: {len(df.columns)}")
                    
                    # 각 컬럼 번역
                    translated_df = df.copy()
                    for col_idx, col in enumerate(df.columns, 1):
                        self.log_message(f"  컬럼 '{col}' 번역 중... ({col_idx}/{len(df.columns)})")
                        translated_df[col] = df[col].apply(
                            lambda x: self.translate_text_with_log(translator, x)
                        )
                    
                    # 번역된 시트 저장
                    translated_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    self.log_message(f"  ✓ 시트 '{sheet_name}' 번역 완료")
                    self.log_message("")
            
            # 완료 메시지
            self.log_message("=" * 50)
            self.log_message("번역 완료!")
            self.log_message(f"총 번역된 셀 수: {translator.translated_count}")
            if translator.error_count > 0:
                self.log_message(f"번역 오류 수: {translator.error_count}")
            self.log_message(f"출력 파일: {output_file}")
            
            # 성공 메시지 박스
            self.root.after(0, lambda: messagebox.showinfo(
                "완료",
                f"번역이 완료되었습니다!\n\n출력 파일: {os.path.basename(output_file)}\n번역된 셀 수: {translator.translated_count}"
            ))
            
        except Exception as e:
            error_msg = f"오류 발생: {str(e)}"
            self.log_message("=" * 50)
            self.log_message("오류 발생!")
            self.log_message(error_msg)
            self.root.after(0, lambda: messagebox.showerror("오류", error_msg))
        
        finally:
            # 버튼 상태 복원
            self.is_translating = False
            self.root.after(0, lambda: self.translate_btn.config(
                state=tk.NORMAL,
                text="번역 시작"
            ))
    
    def translate_text_with_log(self, translator, text):
        """텍스트 번역 (로깅 포함)"""
        import pandas as pd
        from typing import Union
        
        # 숫자나 None 값은 그대로 반환
        if pd.isna(text) or text == '':
            return text
        
        # 숫자 타입은 그대로 반환
        if isinstance(text, (int, float)):
            return text
        
        text_str = str(text).strip()
        
        # 빈 문자열이면 그대로 반환
        if not text_str:
            return text
        
        try:
            # 구글 번역 API 호출
            result = translator.translator.translate(text_str, dest='ko')
            translator.translated_count += 1
            
            # API 제한을 고려한 딜레이
            import time
            time.sleep(0.1)
            
            return result.text
        except Exception as e:
            translator.error_count += 1
            return text_str  # 오류 발생 시 원문 반환


def main():
    """메인 함수"""
    root = tk.Tk()
    app = ExcelTranslateGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

