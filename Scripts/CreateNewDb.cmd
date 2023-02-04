@ECHO OFF

SET SourceDir=%~1
IF "%SourceDir%"=="" (SET /P SourceDir="Input the source path: ")

SET OutFilePath=%~2
IF "%OutFilePath%"=="" (SET /P OutFilePath="Input the output JSON file path: ")

SET PS1_PATH=%~dp0Run.ps1
@ECHO ON
powershell -ExecutionPolicy Bypass -File "%PS1_PATH%" -SourceDir %SourceDir% -OutFilePath "%OutFilePath%"

@ECHO OFF
SET OutFilePath=
SET SourceDir=
SET PS1_PATH=

@PAUSE
