@echo off

echo ��װǰ���˳�����VS2008
pause

set vspath="%userprofile%\Documents\Visual Studio 2008"
set addinpath=%vspath%\Addins
echo Ŀ��Ŀ¼��%addinpath%

echo ====================
echo ��ʼж�ؾɰ汾...
echo ====================
	
del %addinpath%\ToAPI.dll
del %addinpath%\ToAPI.AddIn

echo ====================
echo ��ʼ��װ...
echo ====================
	
if exist %vspath% (
	if not exist %addinpath% (
		mkdir %addinpath%
	)

	copy ToAPI.AddIn %addinpath%
	copy ToAPI.dll %addinpath%
	
	echo ====================
	echo ��װ�ɹ���
	echo ====================
) else (
	echo ==============================
	echo ��װ��ֹ��%vspath% �����ڣ�
	echo ==============================
)

pause