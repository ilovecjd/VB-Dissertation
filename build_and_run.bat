@echo off

REM Change the code page to UTF-8 for proper handling of Korean characters
chcp 65001

echo Clearing the BuildLog.txt...
REM Clear the contents of BuildLog.txt
echo. > "%~dp0BuildLog.txt"

echo Find the existing executable...
if exist "%~dp0프로젝트스케쥴.exe" (
	echo Deleting the existing executable...
    del "%~dp0프로젝트스케쥴.exe"
)

echo Find the pdb file...
if exist "%~dp0프로젝트스케쥴.pdb" (

	echo Deleting the pdb file...
    del "%~dp0프로젝트스케쥴.pdb"
)

REM Restore the original code page (optional)
chcp 437

echo Please wait. Build the VB6 project
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make "%1" /out "%2"

echo Wait for a few seconds to ensure the build process completes
ping 127.0.0.1 -n 5 > nul

echo Run the VB6 project
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /run "%1"