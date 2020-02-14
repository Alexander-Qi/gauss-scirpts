These scripts were tested under Python 3.7.4; the `.bat` file can only be used under Windows, its equivalent written by shell script can easily be found online. 

Usage:
1. Put `G-output-reader.py` in the folder that you store your Gaussian output files.
2. Put `orca-zpreader.py` in the folder that you store your ORCA input files.
3. Modify `combine.py` and `Get-Data.bat` according to your specfic folder names.
4. Double click on `Get-Data.bat`, type '1' for long output and '2' for short output.
5. The results will be presented in a `.xls` file.

By default, `G-output-reader.py` can only read `.log` files, but you can replace 'log' with 'out' in the script to read `.out` files.