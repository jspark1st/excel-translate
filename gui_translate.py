"""
엑셀 파일 번역 GUI 프로그램
구글 번역 API를 사용하여 엑셀 파일의 내용을 한국어로 번역합니다.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import traceback
import logging
from datetime import datetime
from translate_excel import ExcelTranslator
from logger_config import setup_logger

# 로거 설정 (기본값: INFO 레벨로 설정하여 성능 향상)
logger = setup_logger("gui_translate", logging.INFO)


class ExcelTranslateGUI:
    """엑셀 번역 GUI 클래스"""
    
    def __init__(self, root):
        logger.info("GUI 프로그램 초기화 시작")
        self.root = root
        self.root.title("엑셀 파일 한국어 번역 프로그램")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # 변수 초기화
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.is_translating = False
        self.should_stop = False  # 번역 중지 플래그
        self.start_time = None
        self.total_cells = 0  # 전체 셀 수
        self.processed_cells = 0  # 처리된 셀 수
        self.progress_var = tk.DoubleVar()  # 진행률 변수 (0-100)
        self.translator_instance = None  # 번역기 인스턴스 참조 (중지용)
        self.last_progress_time = None  # 마지막 진행률 업데이트 시간
        self.last_progress_cells = 0  # 마지막 진행률 업데이트 시 셀 수
        
        # GUI 구성
        try:
            self.create_widgets()
            logger.debug("GUI 위젯 생성 완료")
        except Exception as e:
            logger.error(f"GUI 위젯 생성 실패: {str(e)}", exc_info=True)
            raise
        
        # 중앙 정렬
        self.center_window()
        logger.info("GUI 프로그램 초기화 완료")
    
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
        self.translate_btn.pack(side=tk.LEFT, padx=5)
        
        # 번역 중지 버튼
        self.stop_btn = tk.Button(
            button_frame,
            text="번역 중지",
            command=self.stop_translation,
            font=("맑은 고딕", 12, "bold"),
            bg="#F44336",
            fg="white",
            relief=tk.RAISED,
            padx=30,
            pady=10,
            cursor="hand2",
            state=tk.DISABLED  # 초기에는 비활성화
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 진행 상황 표시 영역
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.X, padx=20, pady=(10, 5))
        
        progress_label = tk.Label(
            progress_frame,
            text="진행 상황:",
            font=("맑은 고딕", 10, "bold"),
            anchor="w"
        )
        progress_label.pack(side=tk.LEFT)
        
        # 진행률 퍼센트 레이블
        self.progress_percent_label = tk.Label(
            progress_frame,
            text="0%",
            font=("맑은 고딕", 10, "bold"),
            fg="#2196F3",
            anchor="e"
        )
        self.progress_percent_label.pack(side=tk.RIGHT)
        
        # 진행률 바
        self.progress_bar = ttk.Progressbar(
            self.root,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))
        
        # 진행 상황 상세 정보 레이블
        self.progress_detail_label = tk.Label(
            self.root,
            text="대기 중...",
            font=("맑은 고딕", 9),
            fg="gray",
            anchor="w"
        )
        self.progress_detail_label.pack(anchor="w", padx=20, pady=(0, 5))
        
        # 스크롤 가능한 텍스트 영역
        self.progress_text = scrolledtext.ScrolledText(
            self.root,
            height=12,
            font=("맑은 고딕", 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg="#f5f5f5"
        )
        self.progress_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
    
    def select_input_file(self):
        """입력 파일 선택"""
        logger.debug("파일 선택 대화상자 열기")
        try:
            file_path = filedialog.askopenfilename(
                title="엑셀 파일 선택",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            if file_path:
                logger.info(f"입력 파일 선택: {file_path}")
                self.input_file_path.set(file_path)
                # 출력 파일 경로 자동 설정 (입력 파일과 같은 폴더에 '_translated' 추가)
                base_name = os.path.splitext(file_path)[0]
                output_path = f"{base_name}_translated.xlsx"
                self.output_file_path.set(output_path)
                logger.debug(f"출력 파일 경로 자동 설정: {output_path}")
                self.log_message(f"입력 파일 선택: {os.path.basename(file_path)}")
            else:
                logger.debug("파일 선택 취소됨")
        except Exception as e:
            logger.error(f"파일 선택 중 오류: {str(e)}", exc_info=True)
            messagebox.showerror("오류", f"파일 선택 중 오류가 발생했습니다: {str(e)}")
    
    def select_output_file(self):
        """출력 파일 경로 선택"""
        logger.debug("출력 파일 경로 선택 대화상자 열기")
        try:
            file_path = filedialog.asksaveasfilename(
                title="번역된 파일 저장 위치",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("All files", "*.*")
                ]
            )
            if file_path:
                logger.info(f"출력 파일 경로 설정: {file_path}")
                self.output_file_path.set(file_path)
                self.log_message(f"출력 파일 경로 설정: {os.path.basename(file_path)}")
            else:
                logger.debug("출력 파일 경로 선택 취소됨")
        except Exception as e:
            logger.error(f"출력 파일 경로 선택 중 오류: {str(e)}", exc_info=True)
            messagebox.showerror("오류", f"파일 경로 선택 중 오류가 발생했습니다: {str(e)}")
    
    def log_message(self, message, level=logging.INFO):
        """진행 상황 메시지 추가"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {message}"
            
            # 로거에도 기록
            if level == logging.DEBUG:
                logger.debug(message)
            elif level == logging.WARNING:
                logger.warning(message)
            elif level == logging.ERROR:
                logger.error(message)
            else:
                logger.info(message)
            
            # GUI에 표시
            self.progress_text.config(state=tk.NORMAL)
            self.progress_text.insert(tk.END, formatted_message + "\n")
            self.progress_text.see(tk.END)
            self.progress_text.config(state=tk.DISABLED)
            self.root.update_idletasks()
        except Exception as e:
            logger.error(f"로그 메시지 추가 중 오류: {str(e)}", exc_info=True)
    
    def clear_log(self):
        """로그 영역 초기화"""
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.delete(1.0, tk.END)
        self.progress_text.config(state=tk.DISABLED)
        # 진행률 초기화
        self.progress_var.set(0)
        self.progress_percent_label.config(text="0%")
        self.progress_detail_label.config(text="대기 중...")
        self.total_cells = 0
        self.processed_cells = 0
        self.should_stop = False  # 중지 플래그도 초기화
    
    def update_progress(self, current: int, total: int, detail: str = ""):
        """
        진행률 업데이트
        
        Args:
            current: 현재 처리된 셀 수
            total: 전체 셀 수
            detail: 상세 정보 텍스트
        """
        if total > 0:
            progress_percent = (current / total) * 100
            self.progress_var.set(progress_percent)
            self.progress_percent_label.config(text=f"{progress_percent:.1f}%")
            
            # 예상 남은 시간 계산
            remaining_time_str = ""
            if self.start_time and current > 0:
                elapsed_time = (datetime.now() - self.start_time).total_seconds()
                
                # 처리 속도 계산 (셀/초)
                if elapsed_time > 0:
                    cells_per_second = current / elapsed_time
                    remaining_cells = total - current
                    
                    if cells_per_second > 0 and remaining_cells > 0:
                        remaining_seconds = remaining_cells / cells_per_second
                        
                        # 시간 포맷팅
                        if remaining_seconds < 60:
                            remaining_time_str = f"약 {int(remaining_seconds)}초 남음"
                        elif remaining_seconds < 3600:
                            minutes = int(remaining_seconds // 60)
                            seconds = int(remaining_seconds % 60)
                            remaining_time_str = f"약 {minutes}분 {seconds}초 남음"
                        else:
                            hours = int(remaining_seconds // 3600)
                            minutes = int((remaining_seconds % 3600) // 60)
                            remaining_time_str = f"약 {hours}시간 {minutes}분 남음"
            
            # 상세 정보 표시
            if detail:
                if remaining_time_str:
                    self.progress_detail_label.config(
                        text=f"{detail} ({current}/{total} 셀 처리됨) - {remaining_time_str}"
                    )
                else:
                    self.progress_detail_label.config(
                        text=f"{detail} ({current}/{total} 셀 처리됨)"
                    )
            else:
                if remaining_time_str:
                    self.progress_detail_label.config(
                        text=f"진행 중... ({current}/{total} 셀 처리됨) - {remaining_time_str}"
                    )
                else:
                    self.progress_detail_label.config(
                        text=f"진행 중... ({current}/{total} 셀 처리됨)"
                    )
        else:
            self.progress_var.set(0)
            self.progress_percent_label.config(text="0%")
            self.progress_detail_label.config(text=detail if detail else "대기 중...")
        
        self.root.update_idletasks()
    
    def start_translation(self):
        """번역 시작"""
        logger.info("번역 시작 버튼 클릭")
        
        try:
            # 입력 파일 확인
            input_file = self.input_file_path.get().strip()
            if not input_file:
                logger.warning("입력 파일이 선택되지 않음")
                messagebox.showwarning("경고", "입력 파일을 선택해주세요.")
                return
            
            if not os.path.exists(input_file):
                logger.error(f"입력 파일을 찾을 수 없음: {input_file}")
                messagebox.showerror("오류", "입력 파일을 찾을 수 없습니다.")
                return
            
            # 이미 번역 중이면 중단
            if self.is_translating:
                logger.warning("이미 번역이 진행 중")
                messagebox.showinfo("알림", "이미 번역이 진행 중입니다.")
                return
            
            # 출력 파일 경로 설정
            output_file = self.output_file_path.get().strip()
            if not output_file:
                base_name = os.path.splitext(input_file)[0]
                output_file = f"{base_name}_translated.xlsx"
                self.output_file_path.set(output_file)
                logger.debug(f"출력 파일 경로 자동 생성: {output_file}")
            
            logger.info(f"번역 시작 - 입력: {input_file}, 출력: {output_file}")
            
            # 번역 시작
            self.is_translating = True
            self.should_stop = False  # 중지 플래그 초기화
            self.start_time = datetime.now()
            self.last_progress_time = datetime.now()
            self.last_progress_cells = 0
            self.translate_btn.config(state=tk.DISABLED, text="번역 중...")
            self.stop_btn.config(state=tk.NORMAL)  # 중지 버튼 활성화
            self.clear_log()
            
            # 별도 스레드에서 번역 실행 (UI가 멈추지 않도록)
            thread = threading.Thread(
                target=self.translate_in_thread,
                args=(input_file, output_file),
                daemon=True
            )
            thread.start()
            logger.debug("번역 스레드 시작됨")
            
        except Exception as e:
            logger.error(f"번역 시작 중 오류: {str(e)}", exc_info=True)
            messagebox.showerror("오류", f"번역 시작 중 오류가 발생했습니다: {str(e)}")
            self.is_translating = False
            self.should_stop = False
            self.translate_btn.config(state=tk.NORMAL, text="번역 시작")
            self.stop_btn.config(state=tk.DISABLED)  # 중지 버튼 비활성화
    
    def stop_translation(self):
        """번역 중지"""
        if not self.is_translating:
            return
        
        # 중지 확인
        if messagebox.askyesno("번역 중지", "번역을 중지하시겠습니까?\n\n진행 중인 작업은 저장되지 않습니다."):
            logger.warning("사용자가 번역 중지 요청")
            self.should_stop = True
            
            # 번역기 인스턴스에 중지 플래그 전달
            if self.translator_instance:
                self.translator_instance.should_stop = True
            
            self.log_message("번역 중지 요청됨...", logging.WARNING)
            self.stop_btn.config(state=tk.DISABLED, text="중지 중...")
    
    def translate_in_thread(self, input_file, output_file):
        """별도 스레드에서 번역 실행"""
        translator = None
        try:
            logger.info("번역 스레드 시작")
            self.log_message("=" * 50)
            self.log_message("번역 시작")
            self.log_message(f"입력 파일: {os.path.basename(input_file)}")
            self.log_message(f"출력 파일: {os.path.basename(output_file)}")
            self.log_message("-" * 50)
            
            # 진행률 콜백 함수 정의
            def update_progress_callback(current, total, detail=""):
                """진행률 업데이트 콜백"""
                self.root.after(0, lambda: self.update_progress(current, total, detail))
            
            # 번역기 초기화 (진행률 콜백 전달)
            logger.debug("번역기 인스턴스 생성 시작")
            translator = ExcelTranslator(
                debug_mode=False,  # 성능 향상을 위해 기본값 False
                progress_callback=update_progress_callback
            )
            translator.should_stop = False  # 중지 플래그 초기화
            self.translator_instance = translator  # 인스턴스 참조 저장
            logger.info("번역기 인스턴스 생성 완료")
            
            # 중지 플래그 확인
            if self.should_stop:
                logger.info("번역 시작 전 중지 요청됨")
                self.log_message("번역이 중지되었습니다.", logging.WARNING)
                return
            
            # 번역 실행
            try:
                translator.translate_excel(input_file, output_file)
            except InterruptedError:
                # 번역 중지 요청 처리
                logger.warning("번역이 중지되었습니다.")
                self.log_message("번역이 중지되었습니다.", logging.WARNING)
                self.root.after(0, lambda: messagebox.showwarning(
                    "중지됨",
                    "번역이 중지되었습니다.\n\n일부 셀만 번역되었을 수 있습니다."
                ))
                return
            
            # 완료 시 진행률 100%로 설정
            self.root.after(0, lambda: self.update_progress(
                translator.total_cells, 
                translator.total_cells_to_process if translator.total_cells_to_process > 0 else translator.total_cells,
                "번역 완료!"
            ))
            
            # 완료 메시지
            total_duration = (datetime.now() - self.start_time).total_seconds() if self.start_time else 0
            self.log_message("=" * 50)
            self.log_message("번역 완료!")
            self.log_message(f"총 소요 시간: {total_duration:.2f}초")
            self.log_message(f"총 처리 셀 수: {translator.total_cells}")
            self.log_message(f"번역된 셀 수: {translator.translated_count}")
            self.log_message(f"건너뛴 셀 수: {translator.skipped_count}")
            if translator.error_count > 0:
                self.log_message(f"번역 오류 수: {translator.error_count}", logging.WARNING)
            self.log_message(f"출력 파일: {output_file}")
            
            logger.info("번역 완료")
            logger.info(f"통계 - 총 셀: {translator.total_cells}, 번역: {translator.translated_count}, "
                       f"건너뜀: {translator.skipped_count}, 오류: {translator.error_count}")
            
            # 성공 메시지 박스
            self.root.after(0, lambda: messagebox.showinfo(
                "완료",
                f"번역이 완료되었습니다!\n\n"
                f"출력 파일: {os.path.basename(output_file)}\n"
                f"총 소요 시간: {total_duration:.2f}초\n"
                f"번역된 셀 수: {translator.translated_count}\n"
                f"건너뛴 셀 수: {translator.skipped_count}"
            ))
            
        except KeyboardInterrupt:
            logger.warning("사용자에 의해 번역 중단됨")
            self.log_message("번역이 중단되었습니다.", logging.WARNING)
            self.root.after(0, lambda: messagebox.showwarning("중단", "번역이 중단되었습니다."))
        
        except Exception as e:
            error_msg = f"오류 발생: {str(e)}"
            logger.error(f"번역 중 오류 발생: {str(e)}", exc_info=True)
            logger.debug(f"스택 트레이스:\n{traceback.format_exc()}")
            
            self.log_message("=" * 50, logging.ERROR)
            self.log_message("오류 발생!", logging.ERROR)
            self.log_message(error_msg, logging.ERROR)
            self.log_message(f"상세 로그는 logs 폴더를 확인하세요.", logging.ERROR)
            
            self.root.after(0, lambda: messagebox.showerror(
                "오류",
                f"번역 중 오류가 발생했습니다.\n\n{error_msg}\n\n상세 로그는 logs 폴더를 확인하세요."
            ))
        
        finally:
            # 버튼 상태 복원
            self.is_translating = False
            self.should_stop = False
            self.translator_instance = None
            self.root.after(0, lambda: (
                self.translate_btn.config(state=tk.NORMAL, text="번역 시작"),
                self.stop_btn.config(state=tk.DISABLED, text="번역 중지")
            ))
            logger.debug("번역 스레드 종료")
    
    def translate_text_with_log(self, translator, text):
        """텍스트 번역 (로깅 포함) - GUI에서는 translate_excel의 메서드 사용"""
        # translate_excel.py의 translate_text 메서드를 직접 사용
        return translator.translate_text(text)


def main():
    """메인 함수"""
    try:
        logger.info("=" * 80)
        logger.info("GUI 프로그램 시작")
        logger.debug(f"작업 디렉토리: {os.getcwd()}")
        
        root = tk.Tk()
        app = ExcelTranslateGUI(root)
        
        logger.info("GUI 메인 루프 시작")
        root.mainloop()
        
        logger.info("GUI 프로그램 종료")
    except KeyboardInterrupt:
        logger.warning("사용자에 의해 프로그램 중단됨")
    except Exception as e:
        logger.critical(f"GUI 프로그램 치명적 오류: {str(e)}", exc_info=True)
        logger.debug(f"스택 트레이스:\n{traceback.format_exc()}")
        messagebox.showerror("치명적 오류", f"프로그램 실행 중 치명적 오류가 발생했습니다.\n\n{str(e)}\n\n상세 로그는 logs 폴더를 확인하세요.")


if __name__ == "__main__":
    main()

