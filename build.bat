@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m PyInstaller build.spec --noconfirm
echo.
echo 打包完成，输出目录: dist\销售订单合同生成工具\
pause
