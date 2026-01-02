@echo off
chcp 65001 >nul
echo ================================================
echo spreadsheet - Anaconda 环境配置脚本
echo ================================================
echo.

REM 检查conda是否已安装
where conda >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到 conda 命令
    echo 请确保已安装 Anaconda 或 Miniconda 并已添加到 PATH
    pause
    exit /b 1
)

echo 1. 创建 conda 环境 (d:\env\spreadsheet)
call conda create -p d:\env\spreadsheet python=3.10 -y
if %errorlevel% neq 0 (
    echo 创建环境失败
    pause
    exit /b 1
)

echo.
echo 2. 激活环境并配置中国镜像源
call conda activate d:\env\spreadsheet

REM 配置 conda 清华源
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/main
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/r
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/pro
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/cloud/conda-forge
conda config --set show_channel_urls yes

echo.
echo 3. 配置 pip 清华源
pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
pip config set install.trusted-host pypi.tuna.tsinghua.edu.cn

echo.
echo 4. 下载并缓存所有依赖包到本地
echo 正在下载 Python 依赖包...
cd /d "%~dp0backend"
call pip download -r requirements.txt -d d:\env\offline_packages

echo.
echo 5. 安装依赖到环境
call pip install --no-index --find-links=d:\env\offline_packages -r requirements.txt

echo.
echo 6. 安装前端依赖（如果需要）
cd /d "%~dp0"
if exist node_modules (
    echo node_modules 已存在，跳过安装
) else (
    echo 正在配置 npm 淘宝镜像...
    call npm config set registry https://registry.npmmirror.com
    echo 正在安装前端依赖...
    call npm install
)

echo.
echo ================================================
echo 环境配置完成！
echo ================================================
echo.
echo 环境路径: d:\env\spreadsheet
echo 离线包路径: d:\env\offline_packages
echo.
echo 使用说明:
echo   1. 激活环境: conda activate d:\env\spreadsheet
echo   2. 运行系统: run.bat
echo   3. 离线环境: 将整个项目复制到离线环境后，
echo      使用 conda activate d:\env\spreadsheet 激活环境即可运行
echo.
pause
