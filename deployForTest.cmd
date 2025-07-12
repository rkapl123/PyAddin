@echo off
Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
	set source=bin\Release
)
echo copying from %source%
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\PyAddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\PyAddin.xll"
	copy /Y %source%\PyAddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\PyAddin.dll.config "%appdata%\Microsoft\AddIns\PyAddin.xll.config"
	copy /Y PyAddinCentral.config "%appdata%\Microsoft\AddIns\PyAddinCentral.config"
) else (
	echo 32bit office
	copy /Y %source%\PyAddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns\PyAddin.xll"
	copy /Y %source%\PyAddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\PyAddin.dll.config "%appdata%\Microsoft\AddIns\PyAddin.xll.config"
	copy /Y PyAddinCentral.config "%appdata%\Microsoft\AddIns\PyAddinCentral.config"
)
set source=bin\Release
If "%answr%"=="r" (
	copy /Y %source%\PyAddin-AddIn64-packed.xll Distribution\PyAddin64.xll"
	copy /Y %source%\PyAddin-AddIn-packed.xll Distribution\PyAddin32.xll"
	copy /Y %source%\PyAddin.dll.config Distribution\PyAddin.xll.config
	copy /Y PyAddinCentral.config Distribution
	copy /Y PyAddinUser.config Distribution
)
pause