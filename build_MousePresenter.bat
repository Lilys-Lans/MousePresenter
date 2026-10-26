@echo off
cd /d "%~dp0"

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" goto nocsc

echo Building MousePresenter.exe ...
"%CSC%" /nologo /target:winexe /optimize+ /codepage:65001 /win32icon:mouse_presenter.ico /out:MousePresenter.exe /reference:System.Windows.Forms.dll /reference:System.Drawing.dll MousePresenter.cs
if errorlevel 1 goto error

echo.
echo DONE: MousePresenter.exe is ready in this folder.
echo Double-click it to start. A mouse icon will appear in the system tray.
pause
exit /b 0

:nocsc
echo ERROR: .NET Framework 4 compiler (csc.exe) was not found.
echo Install .NET Framework 4.5 or newer from Windows Update and try again.
pause
exit /b 1

:error
echo.
echo BUILD FAILED. See the messages above.
pause
exit /b 1
