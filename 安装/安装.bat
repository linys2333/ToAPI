@echo off

echo ��װǰ���˳�����VS2008
pause

set vspath="%userprofile%\Documents\Visual Studio 2008"
set addinpath=%vspath%\Addins
echo Ŀ¼��%addinpath%

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