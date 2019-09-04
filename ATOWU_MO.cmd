::ATOWU for Microsoft Office v1.0

::STATUS : [SCRIPT RUNNING WELL & WORKING ON INTERFACE... & NEED IMPROVEMENTS]
::[UPDATE!!] From Now ATOWU will have a new Similliar Script Called "ATOWU for Microsoft Office"


::Turning On Debug Mode if Specified File is Exist
if exist "%temp%\ATOWU-MO.DEBUG" (
    set ATOWUDEBUG=1
    del /Q %temp%\ATOWU-MO.DEBUG
    goto TITLE
) else (
    set ATOWUDEBUG=0
    del /Q %temp%\ATOWU-MO.DEBUG
    @echo off
    cls
)
echo [Status: Starting] ATOWU Engine is Starting...
::This Command to Prevent Starting with Local Run (not Running as Administrator)
sc config bits start= disabled>NUL
if errorlevel 1 goto error
echo [Status: Running] ATOWU Engine Successfully Running
goto TITLE

:error
title Error
echo [Status: Error] Please, Run As Administrator to start the process
pause>NUL
exit


:TITLE
echo [Status: Running] Checking Microsoft Office App and Service...
::Title ATOWU in Command-Line
title ATOWU for Microsoft Office v1.0

:Engine
::Input the Variable to Return the Result
set RESULT_ATOWU=NOT_FOUND
set PID_MICROSOFT_EXCEL=NOT_FOUND
set PID_MICROSOFT_ACCESS=NOT_FOUND
set PID_MICROSOFT_ONENOTE=NOT_FOUND
set PID_MICROSOFT_OUTLOOK=NOT_FOUND
set PID_MICROSOFT_POWERPOINT=NOT_FOUND
set PID_MICROSOFT_PUBLISHER=NOT_FOUND
set PID_MICROSOFT_WORD=NOT_FOUND
set STATUS_SERVICE=NOT_FOUND
set STATUS_SERVICE_IN_ENGINE=NOT_FOUND


::Checking Microsoft Office Tasklist
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set STATUS_SERVICE=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr EXCEL.EXE') do set PID_MICROSOFT_EXCEL=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr MSACCESS.EXE') do set PID_MICROSOFT_ACCESS=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr ONENOTE.EXE') do set PID_MICROSOFT_ONENOTE=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr OUTLOOK.EXE') do set PID_MICROSOFT_OUTLOOK=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr POWERPNT.EXE') do set PID_MICROSOFT_POWERPOINT=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr MSPUB.EXE') do set PID_MICROSOFT_PUBLISHER=%%b
for /f "tokens=2" %%b in ('tasklist ^| findstr WINWORD.EXE') do set PID_MICROSOFT_WORD=%%b
goto CHECK_MICROSOFT_APP_STATE_EXCEL

:CHECK_MICROSOFT_APP_STATE_EXCEL
if %PID_MICROSOFT_EXCEL%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto CHECK_MICROSOFT_APP_STATE_ACCESS
) else (
    goto START_MICROSOFT_SERVICE_EXCEL
)
:CHECK_MICROSOFT_APP_STATE_ACCESS
if %PID_MICROSOFT_ACCESS%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto CHECK_MICROSOFT_APP_STATE_ONENOTE
) else (
    goto START_MICROSOFT_SERVICE_ACCESS
)
:CHECK_MICROSOFT_APP_STATE_ONENOTE
if %PID_MICROSOFT_ONENOTE%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto CHECK_MICROSOFT_APP_STATE_OUTLOOK
) else (
    goto START_MICROSOFT_SERVICE_ONENOTE
)
:CHECK_MICROSOFT_APP_STATE_OUTLOOK
if %PID_MICROSOFT_OUTLOOK%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto CHECK_MICROSOFT_APP_STATE_POWERPOINT
) else (
    goto START_MICROSOFT_SERVICE_OUTLOOK
)
:CHECK_MICROSOFT_APP_STATE_POWERPOINT
if %PID_MICROSOFT_POWERPOINT%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto CHECK_MICROSOFT_APP_STATE_PUBLISHER
) else (
    goto START_MICROSOFT_SERVICE_POWERPOINT
)
:CHECK_MICROSOFT_APP_STATE_PUBLISHER
if %PID_MICROSOFT_PUBLISHER%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto CHECK_MICROSOFT_APP_STATE_WORD
) else (
    goto START_MICROSOFT_SERVICE_PUBLISHER
)
:CHECK_MICROSOFT_APP_STATE_WORD
if %PID_MICROSOFT_WORD%==NOT_FOUND (
    if %STATUS_SERVICE%==RUNNING goto SHUTDOWN_SERVICE
    if %STATUS_SERVICE%==STOPPED goto Engine
) else (
    goto START_MICROSOFT_SERVICE_WORD
)
echo Loop Engine
goto Engine

:START_MICROSOFT_SERVICE_ACCESS
echo [Status: FOUND!!] Microsoft Access is Running, Starting Microsoft Office Service...
::if ATOWU found an Microsoft Office is Running, Start the Service
set RESULT_STATUS_SERVICE_ACCESS=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_ACCESS=%%b
goto CHECK_SERVICE_ACCESS

:CHECK_SERVICE_ACCESS
::Checking if Service is starting or Stopping 
::ATOWU Will Prevent to Stop Service, if the Service is Starting either Stopping
if %RESULT_STATUS_SERVICE_ACCESS%==Please (
    echo [Status: Queued] Service is Starting or Stopping, Please Wait...
    goto CHECK_SERVICE_ACCESS
) else (
    goto CHECK_SERVICE_STATE_ACCESS
)

:CHECK_SERVICE_STATE_ACCESS
::Checking Service Status
set RESULT_STATE_SERVICE_IN_ACCESS=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_ACCESS=%%b
if %RESULT_STATE_SERVICE_IN_ACCESS%==RUNNING goto CHECK_APP_ACCESS
if %RESULT_STATE_SERVICE_IN_ACCESS%==STOPPED goto ATTEMPT_1_START_SERVICE_ACCESS
if %RESULT_STATE_SERVICE_IN_ACCESS%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_ACCESS

:CHECK_APP_ACCESS
echo [Status: Waiting] The Service is Turned on, Waiting Microsoft Access to Shutting Down
goto CHECK_APP_ACCESS_2

:CHECK_APP_ACCESS_2
::Checking App, if Stopped Automatically Stop the service
::if Still Running ATOWU will keep Checking until it stopped (Loop Check)
set PID_MICROSOFT_ACCESS=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr MSACCESS.EXE') do set PID_MICROSOFT_ACCESS=%%b
if %PID_MICROSOFT_ACCESS%==NOT_FOUND (
    echo [Status: Shutting Down] Microsoft Access is Closing, Shutting Down the Service
    goto Engine
) else (
    goto CHECK_APP_ACCESS_2
)


:START_MICROSOFT_SERVICE_EXCEL
echo [Status: FOUND!!] Microsoft Excel is Running, Starting Microsoft Office Service...
set RESULT_STATUS_SERVICE_EXCEL=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_EXCEL=%%b
goto CHECK_SERVICE_EXCEL

:CHECK_SERVICE_EXCEL
if %RESULT_STATUS_SERVICE_EXCEL%==Please (
    echo [Status: Queued] Service is Starting or Stopping, Please Wait...
    goto CHECK_SERVICE_EXCEL
) else (
    goto CHECK_SERVICE_STATE_EXCEL
)

:CHECK_SERVICE_STATE_EXCEL
set RESULT_STATE_SERVICE_IN_EXCEL=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_EXCEL=%%b
if %RESULT_STATE_SERVICE_IN_EXCEL%==RUNNING goto CHECK_APP_EXCEL
if %RESULT_STATE_SERVICE_IN_EXCEL%==STOPPED goto ATTEMPT_1_START_SERVICE_EXCEL
if %RESULT_STATE_SERVICE_IN_EXCEL%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_EXCEL

:CHECK_APP_EXCEL
echo [Status: Waiting] The Service is Turned on, Waiting Microsoft Excel to Shutting Down
goto CHECK_APP_EXCEL_2

:CHECK_APP_EXCEL_2
set PID_MICROSOFT_EXCEL=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr EXCEL.EXE') do set PID_MICROSOFT_EXCEL=%%b
if %PID_MICROSOFT_EXCEL%==NOT_FOUND (
    echo [Status: Shutting Down] Microsoft Excel is Closing, Shutting Down the Service
    goto Engine
) else (
    goto CHECK_APP_EXCEL_2
)

:START_MICROSOFT_SERVICE_ONENOTE
set RESULT_STATUS_SERVICE_ONENOTE=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_ONENOTE=%%b
goto CHECK_SERVICE_ONENOTE

:CHECK_SERVICE_ONENOTE
if %RESULT_STATUS_SERVICE_ONENOTE%==Please (
    goto CHECK_SERVICE_ONENOTE
) else (
    goto CHECK_SERVICE_STATE_ONENOTE
)

:CHECK_SERVICE_STATE_ONENOTE
set RESULT_STATE_SERVICE_IN_ONENOTE=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_ONENOTE=%%b
if %RESULT_STATE_SERVICE_IN_ONENOTE%==RUNNING goto CHECK_APP_ONENOTE
if %RESULT_STATE_SERVICE_IN_ONENOTE%==STOPPED goto ATTEMPT_1_START_SERVICE_ONENOTE
if %RESULT_STATE_SERVICE_IN_ONENOTE%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_ONENOTE

:CHECK_APP_ONENOTE
set PID_MICROSOFT_ONENOTE=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr ONENOTE.EXE') do set PID_MICROSOFT_ONENOTE=%%b
if %PID_MICROSOFT_ONENOTE%==NOT_FOUND (
    goto Engine
) else (
    goto CHECK_APP_ONENOTE
)

:START_MICROSOFT_SERVICE_OUTLOOK
set RESULT_STATUS_SERVICE_OUTLOOK=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_OUTLOOK=%%b
goto CHECK_SERVICE_OUTLOOK

:CHECK_SERVICE_OUTLOOK
if %RESULT_STATUS_SERVICE_OUTLOOK%==Please (
    goto CHECK_SERVICE_OUTLOOK
) else (
    goto CHECK_SERVICE_STATE_OUTLOOK
)

:CHECK_SERVICE_STATE_OUTLOOK
set RESULT_STATE_SERVICE_IN_OUTLOOK=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_OUTLOOK=%%b
if %RESULT_STATE_SERVICE_IN_OUTLOOK%==RUNNING goto CHECK_APP_OUTLOOK
if %RESULT_STATE_SERVICE_IN_OUTLOOK%==STOPPED goto ATTEMPT_1_START_SERVICE_OUTLOOK
if %RESULT_STATE_SERVICE_IN_OUTLOOK%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_OUTLOOK

:CHECK_APP_OUTLOOK
set PID_MICROSOFT_OUTLOOK=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr OUTLOOK.EXE') do set PID_MICROSOFT_OUTLOOK=%%b
if %PID_MICROSOFT_OUTLOOK%==NOT_FOUND (
    goto Engine
) else (
    goto CHECK_APP_OUTLOOK
)

:START_MICROSOFT_SERVICE_POWERPOINT
echo [Status: FOUND!!] Microsoft PowerPoint is Running, Starting Microsoft Office Service...
set RESULT_STATUS_SERVICE_POWERPOINT=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_POWERPOINT=%%b
goto CHECK_SERVICE_POWERPOINT

:CHECK_SERVICE_POWERPOINT
if %RESULT_STATUS_SERVICE_POWERPOINT%==Please (
    echo [Status: Queued] Service is Starting or Stopping, Please Wait...
    goto CHECK_SERVICE_POWERPOINT
) else (
    goto CHECK_SERVICE_STATE_POWERPOINT
)

:CHECK_SERVICE_STATE_POWERPOINT
set RESULT_STATE_SERVICE_IN_POWERPOINT=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_POWERPOINT=%%b
if %RESULT_STATE_SERVICE_IN_POWERPOINT%==RUNNING goto CHECK_APP_POWERPOINT
if %RESULT_STATE_SERVICE_IN_POWERPOINT%==STOPPED goto ATTEMPT_1_START_SERVICE_POWERPOINT
if %RESULT_STATE_SERVICE_IN_POWERPOINT%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_POWERPOINT

:CHECK_APP_POWERPOINT
echo [Status: Waiting] The Service is Turned on, Waiting Microsoft PowerPoint to Shutting Down
goto CHECK_APP_POWERPOINT_2

:CHECK_APP_POWERPOINT_2
set PID_MICROSOFT_POWERPOINT=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr POWERPNT.EXE') do set PID_MICROSOFT_POWERPOINT=%%b
if %PID_MICROSOFT_POWERPOINT%==NOT_FOUND (
    echo [Status: Shutting Down] Microsoft PowerPoint is Closing, Shutting Down the Service
    goto Engine
) else (
    goto CHECK_APP_POWERPOINT_2
)

:START_MICROSOFT_SERVICE_PUBLISHER
echo [Status: FOUND!!] Microsoft Publisher is Running, Starting Microsoft Office Service...
set RESULT_STATUS_SERVICE_PUBLISHER=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_PUBLISHER=%%b
goto CHECK_SERVICE_PUBLISHER

:CHECK_SERVICE_PUBLISHER
if %RESULT_STATUS_SERVICE_PUBLISHER%==Please (
    echo [Status: Queued] Service is Starting or Stopping, Please Wait...
    goto CHECK_SERVICE_PUBLISHER
) else (
    goto CHECK_SERVICE_STATE_PUBLISHER
)

:CHECK_SERVICE_STATE_PUBLISHER
set RESULT_STATE_SERVICE_IN_PUBLISHER=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_PUBLISHER=%%b
if %RESULT_STATE_SERVICE_IN_PUBLISHER%==RUNNING goto CHECK_APP_PUBLISHER
if %RESULT_STATE_SERVICE_IN_PUBLISHER%==STOPPED goto ATTEMPT_1_START_SERVICE_PUBLISHER
if %RESULT_STATE_SERVICE_IN_PUBLISHER%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_PUBLISHER

:CHECK_APP_PUBLISHER
echo [Status: Waiting] The Service is Turned on, Waiting Microsoft Publisher to Shutting Down
goto CHECK_APP_PUBLISHER_2

:CHECK_APP_PUBLISHER_2
set PID_MICROSOFT_PUBLISHER=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr MSPUB.EXE') do set PID_MICROSOFT_PUBLISHER=%%b
if %PID_MICROSOFT_PUBLISHER%==NOT_FOUND (
    echo [Status: Shutting Down] Microsoft Excel is Closing, Shutting Down the Service
    goto Engine
) else (
    goto CHECK_APP_PUBLISHER_2
)

:START_MICROSOFT_SERVICE_WORD
echo [Status: FOUND!!] Microsoft Word is Running, Starting Microsoft Office Service...
set RESULT_STATUS_SERVICE_WORD=NOT_FOUND
sc config ClickToRunSvc start=auto>NUL
for /f "tokens=7" %%b in ('net start ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE_WORD=%%b
goto CHECK_SERVICE_WORD

:CHECK_SERVICE_WORD
if %RESULT_STATUS_SERVICE_WORD%==Please (
    echo [Status: Queued] Service is Starting or Stopping, Please Wait...
    goto CHECK_SERVICE_WORD
) else (
    goto CHECK_SERVICE_STATE_WORD
)

:CHECK_SERVICE_STATE_WORD
set RESULT_STATE_SERVICE_IN_WORD=NOT_FOUND
for /f "tokens=4" %%b in ('sc query ClickToRunSvc ^| findstr STATE') do set RESULT_STATE_SERVICE_IN_WORD=%%b
if %RESULT_STATE_SERVICE_IN_WORD%==RUNNING goto CHECK_APP_WORD
if %RESULT_STATE_SERVICE_IN_WORD%==STOPPED goto ATTEMPT_1_START_SERVICE_WORD
if %RESULT_STATE_SERVICE_IN_WORD%==NOT_FOUND goto ATTEMPT_1_START_SERVICE_WORD

:CHECK_APP_WORD
echo [Status: Waiting] The Service is Turned on, Waiting Microsoft Word to Shutting Down
goto CHECK_APP_WORD_2

:CHECK_APP_WORD_2
set PID_MICROSOFT_WORD=NOT_FOUND
for /f "tokens=2" %%b in ('tasklist ^| findstr WINWORD.EXE') do set PID_MICROSOFT_WORD=%%b
if %PID_MICROSOFT_WORD%==NOT_FOUND (
    echo [Status: Shutting Down] Microsoft Word is Closing, Shutting Down the Service
    goto Engine
) else (
    goto CHECK_APP_WORD_2
)

:SHUTDOWN_SERVICE
::Shutting Down Service Command
set RESULT_STATUS_SERVICE=NOT_FOUND
sc config ClickToRunSvc start=disabled>NUL
for /f "tokens=7" %%b in ('net stop ClickToRunSvc ^| findstr service') do set RESULT_STATUS_SERVICE=%%b
echo [Status: Turned Off] The Service is Turned Off
goto Engine







