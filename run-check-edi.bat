@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===== Check EDI Various Forms (HAHM) =====
echo วางไฟล์ MANIFEST (.xls/.xlsx) และ ENTER (.pdf) ไว้ในโฟลเดอร์ input ก่อน
echo.

rem หา Python ที่รันได้จริง (ตัว python.exe ของ Microsoft Store บางเครื่องเป็นแค่ทางลัดที่รันไม่ได้)
set PY=
py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY python -c "import sys" >nul 2>nul && set PY=python
if not defined PY (
  echo !! ไม่พบ Python - ติดตั้งจาก https://www.python.org/downloads/ แล้วติ๊ก "Add python.exe to PATH"
  pause
  exit /b 1
)

%PY% -c "import openpyxl, xlrd, pymupdf" >nul 2>nul
if errorlevel 1 (
  echo กำลังติดตั้งไลบรารีที่ต้องใช้ ครั้งแรกครั้งเดียว...
  %PY% -m pip install -r requirements.txt
)

set SHEDL=0141
set SHEDO=0122
set /p SHEDL=SHED NO. ของ B/L ที่ STATUS เป็น CFS/LCL [Enter = 0141]:
set /p SHEDO=SHED NO. ของ B/L ที่เหลือ (CY) [Enter = 0122]:
if not exist output mkdir output

rem ไม่เขียนทับรายงานเดิม: EDI-HAHM.xlsx, EDI-HAHM-2.xlsx, EDI-HAHM-3.xlsx ...
set NAME=EDI-HAHM.xlsx
set N=1
:next
if exist "output\%NAME%" (
  set /a N+=1
  call set NAME=EDI-HAHM-%%N%%.xlsx
  goto next
)
%PY% scripts\hahm_edi.py --indir . --outdir output --shed-lcl %SHEDL% --shed-other %SHEDO% --name %NAME%
echo.
echo ตรวจย้อนกลับว่าทุกค่าในรายงานมาจากไฟล์จริง:
for %%M in (input\*.xls input\*.xlsx) do set MANF=%%M
for %%E in (input\*.pdf) do set ENTF=%%E
%PY% scripts\verify_trace.py "%MANF%" "%ENTF%"
echo.
echo ผลลัพธ์อยู่ที่ output\%NAME%
pause
