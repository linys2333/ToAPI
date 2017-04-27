@echo off

echo 卸载前请退出所有VS2008
pause

set addinpath="%userprofile%\Documents\Visual Studio 2008\Addins"
echo 目录：%addinpath%

del %addinpath%\ToAPI.dll
del %addinpath%\ToAPI.AddIn

echo ====================
echo 卸载完毕！
echo ====================

pause