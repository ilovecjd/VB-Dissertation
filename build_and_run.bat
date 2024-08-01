@echo off

REM Define variables for project path, log file path, and executable
set PROJECT_PATH=%~dp0Simulator.vbp
set LOG_FILE_PATH=%~dp0BuildLog.txt
set EXECUTABLE=%~dp0프로젝트스케쥴.exe

REM Display paths (for debugging purposes)
REM echo 프로젝트 경로: %PROJECT_PATH%
REM echo 로그 파일 경로: %LOG_FILE_PATH%
REM echo 실행 파일 경로: %EXECUTABLE%

REM Set the code page to EUC-KR
chcp 949 >nul

echo =======================================
echo 빌드 로그를 초기화합니다...

REM Clear the contents of the log file
echo. > %LOG_FILE_PATH%

echo 기존 실행 파일을 삭제합니다...

REM Delete the existing executable if it exists
if exist "%EXECUTABLE%" (
    del "%EXECUTABLE%"
    echo %EXECUTABLE%가 삭제되었습니다. >> %LOG_FILE_PATH%
    echo %EXECUTABLE%가 삭제되었습니다. 
) else (
    echo 기존 %EXECUTABLE% 파일이 없습니다. >> %LOG_FILE_PATH%
    echo 기존 %EXECUTABLE% 파일이 없습니다.
)

echo.
echo =======================================
echo 잠시만 기다려 주세요. VB6 프로젝트를 빌드 중입니다...

REM Build the VB6 project and redirect output to the log file
REM 줄바꿈
echo. >> %LOG_FILE_PATH%
echo !!에러 발생라인에 +2 하여 살펴 보세요!!! >> %LOG_FILE_PATH%
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make %PROJECT_PATH% /out %LOG_FILE_PATH% >> %LOG_FILE_PATH% 2>&1
set ERRORLEVEL=%ERRORLEVEL%

REM Check if the build was successful and log the result
if %ERRORLEVEL% equ 0 (
    echo 빌드가 성공했습니다. >> %LOG_FILE_PATH%
    echo 빌드가 성공했습니다. 

    echo 빌드가 완료되기를 기다리는 중입니다...
    
    REM Wait for a few seconds
    ping 127.0.0.1 -n 5 > nul

    echo.
    echo =======================================
    echo VB6 프로젝트를 실행합니다
    
    REM Run the VB6 project
    "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /run %PROJECT_PATH%

    echo.
    echo =======================================
    echo 프로세스가 완료되었습니다.
) else (
    echo 빌드가 실패했습니다. 오류 코드 %ERRORLEVEL%. >> %LOG_FILE_PATH%
    echo 빌드가 실패했습니다. 오류 코드 %ERRORLEVEL%. 
)

