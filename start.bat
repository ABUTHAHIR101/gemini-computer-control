@echo off
chcp 65001 >nul
echo ========================================
echo   Gemini计算机控制系统 - 快速启动
echo ========================================
echo.

echo [1/3] 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到Python，请先安装Python 3.8+
    pause
    exit /b 1
)
echo ✅ Python环境正常
echo.

echo [2/3] 检查后端依赖...
cd backend
if not exist .env (
    echo ⚠️  未找到.env文件，正在创建...
    copy .env.example .env >nul
    echo ✅ 已创建.env文件，请编辑并填入你的GEMINI_API_KEY
    echo.
    echo 按任意键继续安装依赖...
    pause >nul
)

echo 正在安装Python依赖...
pip install -r requirements.txt >nul 2>&1
if errorlevel 1 (
    echo ⚠️  依赖安装可能有问题，尝试继续...
) else (
    echo ✅ 依赖安装完成
)
echo.

echo [3/3] 启动全栈服务...
echo.
echo ========================================
echo   服务启动中...
echo   访问地址: http://localhost:5000
echo ========================================
echo.
echo 💡 提示：
echo   - 确保已在 .env 文件中配置 GEMINI_API_KEY
echo   - 浏览器将自动打开 http://localhost:5000
echo   - 如果未自动打开，请手动输入地址
echo   - 按 Ctrl+C 停止服务
echo.

:: 自动在浏览器中打开前端页面
start http://localhost:5000

:: 启动后端（它现在也负责提供前端页面）
python main.py