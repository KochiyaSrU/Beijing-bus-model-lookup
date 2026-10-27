@echo off
setlocal
cd /d "%~dp0"
python generate_icon.py
if /i "%~1"=="folder" goto folder

python -m PyInstaller --noconfirm --clean --onefile --windowed --icon "北京公交车型查询器.ico" --add-data "北京公交车型查询器.ico;." --version-file "version_info.txt" --name "北京公交车型查询器" --exclude-module numpy --exclude-module unittest --exclude-module pydoc --exclude-module pdb --exclude-module doctest --exclude-module difflib --exclude-module PIL.ImageQt --exclude-module PIL.ImageShow --exclude-module test --exclude-module lib2to3 --exclude-module xmlrpc BPT_BusDigitalNumber.py
copy /y "dist\北京公交车型查询器.exe" ".\北京公交车型查询器.exe" >nul
rmdir /s /q build
rmdir /s /q dist
del /q "北京公交车型查询器.spec"
echo.
echo 构建完成：北京公交车型查询器.exe（单文件版，约 17.7 MB）
echo 请确保 BPT_BusDigitalNumber.xlsx 与 EXE 位于同一文件夹。
echo 想要启动更快的版本（实测约 1 秒 vs 2 秒），可以改用： build.bat folder
pause
exit /b 0

:folder
python -m PyInstaller --noconfirm --clean --onedir --windowed --icon "北京公交车型查询器.ico" --add-data "北京公交车型查询器.ico;." --version-file "version_info.txt" --name "北京公交车型查询器" --exclude-module numpy --exclude-module unittest --exclude-module pydoc --exclude-module pdb --exclude-module doctest --exclude-module difflib --exclude-module PIL.ImageQt --exclude-module PIL.ImageShow --exclude-module test --exclude-module lib2to3 --exclude-module xmlrpc BPT_BusDigitalNumber.py
rmdir /s /q build
del /q "北京公交车型查询器.spec"
echo.
echo 构建完成：dist\北京公交车型查询器\ 文件夹版（约 38 MB，启动更快）
echo 把数据文件放进 dist\北京公交车型查询器\ 后运行里面的 EXE 即可。
pause
exit /b 0