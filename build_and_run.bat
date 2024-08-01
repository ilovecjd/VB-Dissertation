@echo off

REM Define variables for project path, log file path, and executable
set PROJECT_PATH=%~dp0Simulator.vbp
set LOG_FILE_PATH=%~dp0BuildLog.txt
set EXECUTABLE=%~dp0������Ʈ������.exe

REM Display paths (for debugging purposes)
REM echo ������Ʈ ���: %PROJECT_PATH%
REM echo �α� ���� ���: %LOG_FILE_PATH%
REM echo ���� ���� ���: %EXECUTABLE%

REM Set the code page to EUC-KR
chcp 949 >nul

echo =======================================
echo ���� �α׸� �ʱ�ȭ�մϴ�...

REM Clear the contents of the log file
echo. > %LOG_FILE_PATH%

echo ���� ���� ������ �����մϴ�...

REM Delete the existing executable if it exists
if exist "%EXECUTABLE%" (
    del "%EXECUTABLE%"
    echo %EXECUTABLE%�� �����Ǿ����ϴ�. >> %LOG_FILE_PATH%
    echo %EXECUTABLE%�� �����Ǿ����ϴ�. 
) else (
    echo ���� %EXECUTABLE% ������ �����ϴ�. >> %LOG_FILE_PATH%
    echo ���� %EXECUTABLE% ������ �����ϴ�.
)

echo.
echo =======================================
echo ��ø� ��ٷ� �ּ���. VB6 ������Ʈ�� ���� ���Դϴ�...

REM Build the VB6 project and redirect output to the log file
REM �ٹٲ�
echo. >> %LOG_FILE_PATH%
echo !!���� �߻����ο� +2 �Ͽ� ���� ������!!! >> %LOG_FILE_PATH%
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make %PROJECT_PATH% /out %LOG_FILE_PATH% >> %LOG_FILE_PATH% 2>&1
set ERRORLEVEL=%ERRORLEVEL%

REM Check if the build was successful and log the result
if %ERRORLEVEL% equ 0 (
    echo ���尡 �����߽��ϴ�. >> %LOG_FILE_PATH%
    echo ���尡 �����߽��ϴ�. 

    echo ���尡 �Ϸ�Ǳ⸦ ��ٸ��� ���Դϴ�...
    
    REM Wait for a few seconds
    ping 127.0.0.1 -n 5 > nul

    echo.
    echo =======================================
    echo VB6 ������Ʈ�� �����մϴ�
    
    REM Run the VB6 project
    "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /run %PROJECT_PATH%

    echo.
    echo =======================================
    echo ���μ����� �Ϸ�Ǿ����ϴ�.
) else (
    echo ���尡 �����߽��ϴ�. ���� �ڵ� %ERRORLEVEL%. >> %LOG_FILE_PATH%
    echo ���尡 �����߽��ϴ�. ���� �ڵ� %ERRORLEVEL%. 
)

