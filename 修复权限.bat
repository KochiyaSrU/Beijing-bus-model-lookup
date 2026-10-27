@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  修复本文件夹的沙箱权限标记
echo.
echo  本文件夹里的文件和 EXE 被打上了「低完整性级别」
echo  Low Mandatory Label，文件夹上还有一条
echo  「Everyone 拒绝删除子项」的规则。这两样都不是本程序
echo  产生的，是自动化工具体系留下的，会让 Windows 对
echo  双击运行 / 复制 / 删除 / 杀软的判断变得反常。
echo.
echo  下面把它们改回正常状态。
echo ============================================================
echo.

echo [1/2] 整棵目录树：完整性级别改为 Medium
icacls "%CD%" /setintegritylevel Medium /T

echo.
echo [2/2] 整棵目录树：移除 Everyone 的拒绝规则
icacls "%CD%" /remove:d *S-1-1-0 /T

echo.
echo ------------------------------------------------------------
echo  处理完毕。
echo  若上面出现「拒绝访问 / Access is denied」，
echo  请右键本文件，选择「以管理员身份运行」再来一次。
echo.
echo  之后建议双击 build.bat 重新生成 EXE。
echo ------------------------------------------------------------
pause
