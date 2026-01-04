@echo off
chcp 65001 >nul

echo.

REM conda init cmd.exe
REM 检查并激活 conda 环境
set CONDA_ENV_PATH=d:\env\spreadsheet

if exist "%CONDA_ENV_PATH%\python.exe" (
    echo 发现 conda 环境，正在激活...
    call conda activate "%CONDA_ENV_PATH%"
    set PYTHON_CMD=python
) else (
    echo 未找到 conda 环境，尝试使用系统 Python...
    set PYTHON_CMD=python
)


echo 检查端口8000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8000 ^| findstr LISTENING') do (
    echo 发现端口8000被占用，正在释放...
    taskkill /F /PID %%a >nul 2>&1
)

echo 安装依赖...
REM 优先使用本地离线包
if exist d:\env\offline_packages (
    echo 使用本地离线包安装...
    %PYTHON_CMD% -m pip install --no-index --find-links=d:\env\offline_packages -r requirements.txt -q
) else (
    echo 在线安装依赖...
    %PYTHON_CMD% -m pip install -r requirements.txt -q
)

cd /d "%~dp0backend"

REM 读取.env配置
set SERVER_IP=NONE
set SERVER_PORT=8000
set ADMIN_KEY=admin123456

for /f "tokens=1,2 delims==" %%a in ('type "%~dp0.env" ^| findstr /v "^#"') do (
    if "%%a"=="SERVER_IP" set SERVER_IP=%%b
    if "%%a"=="SERVER_PORT" set SERVER_PORT=%%b
    if "%%a"=="ADMIN_KEY" set ADMIN_KEY=%%b
)

REM 如果SERVER_IP为NONE，自动获取本机IP
if /i "%SERVER_IP%"=="NONE" (
    for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
        for /f "tokens=1" %%b in ("%%a") do (
            set SERVER_IP=%%b
            goto :got_ip
        )
    )
)
:got_ip

echo.
echo ================================================
echo 启动服务器...
echo ================================================
echo.
echo 服务器IP: %SERVER_IP%
echo 访问地址:
echo   - 管理后台: http://%SERVER_IP%:%SERVER_PORT%/admin/%ADMIN_KEY%
echo.
echo 请在管理后台创建表格并获取访问链接
echo.
echo 按 Ctrl+C 停止服务器
echo ================================================
echo.

REM 延迟2秒后自动打开浏览器（使用127.0.0.1访问本机）
start "" cmd /c "ping -n 3 127.0.0.1 >nul && start http://127.0.0.1:%SERVER_PORT%/admin/%ADMIN_KEY%"

%PYTHON_CMD% -m uvicorn main:app --host 0.0.0.0 --port %SERVER_PORT% --reload

pause
