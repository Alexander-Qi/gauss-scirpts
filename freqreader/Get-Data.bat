@Echo Off
mode con cols=52 lines=27
Set Prog= ��������ȡ����
Set L1=���������������q�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�r
Set L2=���������������U��������������     �U
Set L3=�������q�T�T�T�T�T�T�T�� %Prog% ���T�T�T�T�T�T�r
Set L4=�������U�� �����U���������� ���������U �� ���U
Set L5=�������U�� �����t�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�s�������U
Set L6=�������U������ ���������� �����������������U
Set L7=�������c�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�f
Set L8=�������t���������������������������������������������������������������������s
Title %Prog%

CLS
Title %Prog%
Echo.
Echo %L1%
Echo %L2%
Echo %L3%
Echo %L4%
Echo %L5%
Echo %L6%
Echo �������U �����߱������3��Python�ű�ʹ�á��U
Echo �������U ���Զ�����ȡ����ļ��е���Ϣ���� �U
Echo %L6%
Echo �������U ����ѡ������������������������� �U
Echo %L6%
Echo �������U����[1] ������н��          �����U
Echo �������U����[2] ������ٽ��  ��        ���U
Echo       �U    [Q] �˳����򡡡�          �����U
Echo %L6% 
Echo %L7%
Echo %L8%
Echo.
Set Choice=
Set /P Choice=        ��ѡ����� (1/2/Q) ��Ȼ�󰴻س���
If "%Choice%"=="" Goto Start
If Not "%Choice%"=="" Set Choice=%Choice:~0,1%
If /I "%Choice%"=="1" Goto A
If /I "%Choice%"=="2" Goto B
If /I "%Choice%"=="Q" Exit

:A
cd /d E:\opt
python -u G-output-reader.py .
cd /d E:\zeropoint
python -u orca-zpreader.py .
cd /d E:\
python -u combine.py long
for /r E:\opt %%G in (*.xls) do move %%a E:\xinghuo\archive >nul
for /r E:\zeropoint %%G in (*.xls) do move %%a E:\xinghuo\archive >nul
goto end


:B
cd /d E:\opt
python -u G-output-reader.py .
cd /d E:\zeropoint
python -u orca-zpreader.py .
cd /d E:\
python -u combine.py short
for /r E:\opt %%G in (*.xls) do move %%a E:\xinghuo\archive >nul
for /r E:\zeropoint %%G in (*.xls) do move %%a E:\xinghuo\archive >nul
goto end

:end
Echo    �q�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�r
Echo    |  ���гɹ�����ȥ/archive �鿴��Ľ�� |
Echo    �t�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�s
Echo.
Echo      ��������˳�
pause>nul
exit

