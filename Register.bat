@echo off
cls
echo Trying to register "SESE.dll" (%PROCESSOR_ARCHITECTURE%)
echo ------------------------------------------------------

rem 
rem Get the name of the WINDOWS directory 
rem 
if "%WINDIR%"=="" goto NOWINDIR 
rem 
rem See if the Net Framework v2.0.50727 regasm exists 
rem 

if "%PROCESSOR_ARCHITECTURE%" == "x86" (
	set _NET_FRAMEWORK_=%WINDIR%\microsoft.net\framework\v2.0.50727
) else (
 	set _NET_FRAMEWORK_=%WINDIR%\microsoft.net\framework64\v2.0.50727
)

if NOT EXIST %_NET_FRAMEWORK_%\ goto NO_NET_FRAMEWORK


if NOT EXIST %_NET_FRAMEWORK_%\regasm.exe goto NOFILE 

rem 
rem See if the dll exists 
rem 

if NOT EXIST "SESE.dll" goto NODLLFILE 

rem 
rem Register dll
rem 
%_NET_FRAMEWORK_%\regasm "SESE.dll" -codebase /nologo
goto end 

:NOFILE 
echo ERROR: No regasm.exe file found in %_NET_FRAMEWORK_%\ 
goto end 

:NODLLFILE
echo ERROR: No "SESE.dll" found in
goto END

:NOWINDIR 
echo ERROR: No WINDIR environment variable found 
goto end 

:NO_NET_FRAMEWORK
echo ERROR: %_NET_FRAMEWORK_% not found
goto end 

:end 

pause