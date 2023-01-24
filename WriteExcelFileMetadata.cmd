@ECHO OFF

SET SourcePath=%~1
IF "%SourcePath%"=="" (SET /P SourcePath="Input the source path: ")

SET OutFilePath=%~2
IF "%OutFilePath%"=="" (SET /P OutFilePath="Input the output JSON file path: ")

SET PS1_PATH=%~dp0Run.ps1
@ECHO ON
powershell -ExecutionPolicy Bypass -File "%PS1_PATH%" -SourcePath %SourcePath% -OutFilePath "%OutFilePath%"

@ECHO OFF
SET OutFilePath=
SET SourcePath=
SET PS1_PATH=

@PAUSE
