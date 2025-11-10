# PowerShell 실행 스크립트
Write-Host "가상 환경 활성화 중..." -ForegroundColor Green
& .\venv\Scripts\Activate.ps1
Write-Host "GUI 프로그램 실행 중..." -ForegroundColor Green
python gui_translate.py

