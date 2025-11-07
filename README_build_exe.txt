README - 打包为 Windows .exe 使用说明

目标：在 Windows 上把 Python 脚本打包为单文件可执行程序 (.exe)。

前提 (Windows 环境):
- Windows 10/11
- 已安装 Python 3.8+（请把 Python 添加到 PATH）
- pip 可用

步骤：
1. 把以下文件放在同一文件夹：
   - 批量发货匹配_v2.py
   - fahuo.xlsx
   - 订单.xlsx
   - 快递.txt
   - requirements.txt
   - build_exe.bat

2. 安装依赖（可选，但建议）：
   pip install -r requirements.txt

3. 安装 PyInstaller：
   pip install pyinstaller

4. 运行构建脚本（双击或在 cmd 中运行）：
   build_exe.bat

5. 成功后，查看 dist\\ 执行文件夹，里面会有 批量发货匹配_v2.exe
   把 fahuo.xlsx、订单.xlsx、快递.txt 放到 exe 同目录，双击运行即可。

常见问题：
- 打包后运行报缺少 dll：请保证在目标机器上也安装了 Microsoft Visual C++ Redistributable。
- 如果需要图形界面或更多日志，可以联系我帮你定制。

