@echo off

REM Change the code page to UTF-8 for proper handling of Korean characters
chcp 65001

echo Clearing the BuildLog.txt...
REM Clear the contents of BuildLog.txt
echo. > "%~dp0BuildLog.txt"

echo Find the existing executable...
if exist "%~dp0프로젝트스케쥴.exe" (
    del "%~dp0프로젝트스케쥴.exe"
	echo 프로젝트스케쥴.exe has been deleted.
)else (
    echo No existing ?봽濡쒖젥?듃?뒪耳?伊?.exe found.
)

echo Find the pdb file...
if exist "%~dp0프로젝트스케쥴.pdb" (
    del "%~dp0프로젝트스케쥴.pdb"
	echo 프로젝트스케쥴.pdb has been deleted.
)else (
    echo No existing 프로젝트스케쥴.pdb found.
)

REM Restore the original code page (optional)
chcp 437

echo Please wait. Build the VB6 project
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make "%1" /out "%2"
set ERRORLEVEL=%ERRORLEVEL%


if %ERRORLEVEL% equ 0 (
    echo Build was successful.
    echo Wait for a few seconds to ensure the build process completes
    ping 127.0.0.1 -n 5 > nul
	
	echo Run the VB6 project    
    "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /run "%1"
) else (
    echo Build failed with error code %ERRORLEVEL%.
)

echo =======================================
echo Process completed.
echo =======================================
