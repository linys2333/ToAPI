@echo off

echo 卸载前请退出所有VS2008
pause

set addinpath="%userprofile%\Documents\Visual Studio 2008\Addins"
echo 目标目录：%addinpath%

echo ====================
echo 开始卸载...
echo ====================

del %addinpath%\ToAPI.dll
del %addinpath%\ToAPI.AddIn

echo ====================
echo 卸载完毕！
echo ====================

pause