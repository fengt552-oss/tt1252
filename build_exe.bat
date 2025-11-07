@echo off
REM Windows: 在当前目录下打包为单文件 exe（需要安装 pyinstaller）
REM 执行前请先在 cmd 中运行: pip install pyinstaller
pyinstaller --onefile --clean --name "批量发货匹配_v2" 批量发货匹配_v2.py
pause
