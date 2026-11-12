@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ======================================================
echo    LQD - LOC BANG DIEM THEO SO BAO DANH (THPT LE QUY DON)
echo ======================================================
echo.
set /p INPUT=Nhap ten file nguon (.json/.csv/.xlsx), vi du K12-4-Danh-sach-phong-thi_va_Phieu-thu-bai.xlsx: 
if "%INPUT%"=="" (
  echo [LOI] Ban chua nhap file nguon.
  pause
  exit /b 1
)
echo.
echo Dang xu ly...
py -3 export_lqd_scores.py --input "%INPUT%" --sbd-file sbd.txt --output Bang_diem_loc_theo_SBD.xlsx 2>nul || python export_lqd_scores.py --input "%INPUT%" --sbd-file sbd.txt --output Bang_diem_loc_theo_SBD.xlsx
echo.
pause
