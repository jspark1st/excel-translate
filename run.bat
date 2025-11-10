@echo off
echo 가상 환경 활성화 중...
call venv\Scripts\activate.bat
echo GUI 프로그램 실행 중...
python gui_translate.py
pause

