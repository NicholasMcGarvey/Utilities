    @echo off
    SETLOCAL ENABLEDELAYEDEXPANSION

    IF "%~1" == "" GOTO :NoFileDropped

    SET "dbPath=%~dpnx1"
    SET "accessExePath=C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" REM Adjust this path to your Access installation

    IF NOT EXIST "%accessExePath%" (
        ECHO Error: MSACCESS.EXE not found at "%accessExePath%".
        GOTO :EOF
    )


	IF NOT EXIST "%CD%\deCompileBackups\" (
		mkdir "%CD%\deCompileBackups\	
	)


	set "dt=%date:~10,4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
	set "buDestPath=%CD%\deCompileBackups\%~n1_%dt%%~x1" 

	ECHO.

	ECHO Copying: %dbPath% 
	ECHO to: %buDestPath%...
	COPY "%~1" "%buDestPath%"

	ECHO.

    ECHO Decompiling and Compacting "%dbPath%"...

    "%accessExePath%" "%dbPath%" /decompile /compact
    ECHO Decompilation and Compact complete.

    ::"%accessExePath%" "%dbPath%" /compact
    ::ECHO Compaction complete.

	ECHO.
    ECHO Process finished for "%dbPath%".
	ECHO.
    PAUSE
    GOTO :EOF

    :NoFileDropped
    ECHO Please drag and drop an Access database file onto this batch file.
    PAUSE
	
	:EOF