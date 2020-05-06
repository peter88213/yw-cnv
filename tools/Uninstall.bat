@echo off
REM Removes the PyWriter Tools software package.
REM 
REM See: https://github.com/peter88213/PyWriter
REM License: The MIT License (https://opensource.org/licenses/mit-license.php)
REM Copyright: (c) 2020, Peter Triesberger
REM 
REM Note: This script is to be executed manually.
REM 
REM Preconditions:
REM * PyWriter Tools are installed.
REM * LibreOffice 5.x or 6.x is installed.
REM
REM Postconditions:
REM * Previously auto-installed items of PyWriter Tools are removed.
REM * The Explorer context menu entry "PyWriter Tools" is removed.

set _release=0.10.0

pushd setup

set _LibreOffice5_w64=c:\Program Files (x86)\LibreOffice 5
set _LibreOffice5_w32=c:\Program Files\LibreOffice 5
set _LibreOffice6_w64=c:\Program Files (x86)\LibreOffice
set _LibreOffice6_w32=c:\Program Files\LibreOffice

set _LibreOffice_Userprofile=AppData\Roaming\LibreOffice\4\user

echo -----------------------------------------------------------------
echo PyWriter (yWriter to OpenOffice/LibreOffice) v%_release%
echo Installing software package ...
echo -----------------------------------------------------------------

rem Detect combination of Windows and Office 

if exist "%_LibreOffice5_w64%\program\swriter.exe" goto LibreOffice5-Win64
if exist "%_LibreOffice5_w32%\program\swriter.exe" goto LibreOffice5-Win32
if exist "%_LibreOffice6_w64%\program\swriter.exe" goto LibreOffice6-Win64
if exist "%_LibreOffice6_w32%\program\swriter.exe" goto LibreOffice6-Win32
echo ERROR: No supported version of LibreOffice found!
echo Installation aborted
goto end

:LibreOffice5-Win64
set _writer=%_LibreOffice5_w64%
set _user=%USERPROFILE%\%_LibreOffice_Userprofile%
echo LibreOffice 5 found ...
goto settings_done

:LibreOffice5-Win32
set _writer=%_LibreOffice5_w32%
set _user=%USERPROFILE%\%_LibreOffice_Userprofile%
echo LibreOffice 5 found ...
goto settings_done

:LibreOffice6-Win64
set _writer=%_LibreOffice6_w64%
set _user=%USERPROFILE%\%_LibreOffice_Userprofile%
echo LibreOffice found ...
goto settings_done

:LibreOffice6-Win32
set _writer=%_LibreOffice6_w32%
set _user=%USERPROFILE%\%_LibreOffice_Userprofile%
echo LibreOffice found ...
goto settings_done

:settings_done

echo Deleting program components in %_user%\Scripts\python ...

del /q "%_user%\Scripts\python\saveyw7.py"
del /q c:\pywriter\openyw7.bat

echo Removing LibreOffice extension ...

"%_writer%\program\unopkg" remove -f SaveYw7-L-%_release%.oxt

echo Removing Explorer context menu entry (You may be asked for approval) ...

del_cm.reg

popd

echo -----------------------------------------------------------------
echo #
echo # PyWriter v%_release% is removed from your PC.
echo #
echo -----------------------------------------------------------------
pause



:end
