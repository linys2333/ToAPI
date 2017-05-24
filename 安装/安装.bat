@echo off

echo 安装前请退出所有VS2008
pause

set vspath="%userprofile%\Documents\Visual Studio 2008"
set addinpath=%vspath%\Addins
echo 目标目录：%addinpath%

echo ====================
echo 开始卸载旧版本...
echo ====================
	
del %addinpath%\ToAPI.dll
del %addinpath%\ToAPI.AddIn

echo ====================
echo 开始安装...
echo ====================
	
if exist %vspath% (
	if not exist %addinpath% (
		mkdir %addinpath%
	)

	copy ToAPI.AddIn %addinpath%
	copy ToAPI.dll %addinpath%
	
	echo ====================
	echo 安装成功！
	echo ====================
) else (
	echo ==============================
	echo 安装中止，%vspath% 不存在！
	echo ==============================
)

pause