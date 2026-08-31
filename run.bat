@echo off
chcp 65001 > nul
title ระบบบริหารจัดการข้อมูลครู อาจารย์ ครูฝึก
color 0b

echo =====================================================================
echo       ระบบบริหารจัดการข้อมูลครู อาจารย์ ครูฝึก (บช.ตชด.)
echo =====================================================================
echo.

:: 1. Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] ไม่พบ Python ในเครื่อง กรุณาติดตั้ง Python 3.10 ขึ้นไป
    pause
    exit /b 1
)

:: 2. Check Virtual Environment
if not exist ".venv" (
    echo [*] กำลังสร้าง Virtual Environment (.venv)...
    python -m venv .venv
)

:: 3. Activate and Install Requirements
echo [*] กำลังตรวจสอบและติดตั้งส่วนประกอบที่จำเป็น...
call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet

:: 4. Launch Browser in background after 2 seconds
echo.
echo [*] กำลังเปิดระบบในเว็บเบราว์เซอร์ (http://localhost:8000)...
start "" cmd /c "timeout /t 2 >nul & start http://localhost:8000"

:: 5. Start FastAPI Server
echo =====================================================================echo  ระบบกำลังทำงานที่: http://localhost:8000
 echo  (สามารถปิดหน้านี้เมื่อต้องการหยุดการทำงานของระบบ)
 echo =====================================================================
 echo.
 echo  [หมายเหตุ] ตั้งค่า environment variables ก่อนรัน:
 echo    set ADMIN_API_KEY=your-secret-key    (รหัสผ่านสำหรับแก้ไขข้อมูล)
 echo    set ALLOWED_ORIGINS=https://yourdomain.com  (จำกัด CORS)
 echo.
 python main.py

pause
