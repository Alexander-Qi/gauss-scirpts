@Echo Off
mode con cols=52 lines=27
Set Prog= 计算结果提取工具
Set L1=　　　　　　　╭═══════════════════╮
Set L2=　　　　　　　║　　　　　　　     ║
Set L3=　　　╭═══════┤ %Prog% ├══════╮
Set L4=　　　║　 　　║　　　　　 　　　　║ 　 　║
Set L5=　　　║　 　　╰═══════════════════╯　　　║
Set L6=　　　║　　　 　　　　　 　　　　　　　　║
Set L7=　　　╟══════════════════════════════════╢
Set L8=　　　╰──────────────────────────────────╯
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
Echo 　　　║ 本工具必须配合3个Python脚本使用　║
Echo 　　　║ 可自动化提取输出文件中的信息　　 ║
Echo %L6%
Echo 　　　║ 　请选择操作：　　　　　　　　　 ║
Echo %L6%
Echo 　　　║　　[1] 输出所有结果          　　║
Echo 　　　║　　[2] 输出较少结果  　        　║
Echo       ║    [Q] 退出程序　　          　　║
Echo %L6% 
Echo %L7%
Echo %L8%
Echo.
Set Choice=
Set /P Choice=        请选择操作 (1/2/Q) ，然后按回车：
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
Echo    ╭═════════════════════════╮
Echo    |  运行成功！请去/archive 查看你的结果 |
Echo    ╰═════════════════════════╯
Echo.
Echo      按任意键退出
pause>nul
exit

