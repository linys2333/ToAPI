@echo off

echo 安装前请退出所有VS2008
pause

set vspath="%userprofile%\Documents\Visual Studio 2008"
set addinpath=%vspath%\Addins
echo 目录：%addinpath%

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