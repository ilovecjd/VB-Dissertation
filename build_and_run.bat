@echo off

REM Change the code page to UTF-8 for proper handling of Korean characters
chcp 65001

echo Clearing the BuildLog.txt...
REM Clear the contents of BuildLog.txt
echo. > "%~dp0BuildLog.txt"

echo Find the existing executable...
if exist "%~dp0ÇÁ·ÎÁ§Æ®½ºÄÉÁì.exe" (
    del "%~dp0ÇÁ·ÎÁ§Æ®½ºÄÉÁì.exe"
	echo ÇÁ·ÎÁ§Æ®½ºÄÉÁì.exe has been deleted.
)else (
    echo No existing ?”„ë¡œì ?Š¸?Š¤ì¼?ì¥?.exe found.
)

echo Find the pdb file...
if exist "%~dp0ÇÁ·ÎÁ§Æ®½ºÄÉÁì.pdb" (
    del "%~dp0ÇÁ·ÎÁ§Æ®½ºÄÉÁì.pdb"
	echo ÇÁ·ÎÁ§Æ®½ºÄÉÁì.pdb has been deleted.
)else (
    echo No existing ÇÁ·ÎÁ§Æ®½ºÄÉÁì.pdb found.
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
