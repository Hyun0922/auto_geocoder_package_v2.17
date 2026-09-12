@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Auto Geocoder EXE Builder
echo ============================================================
echo  Auto Geocoder EXE Builder
echo ============================================================
echo.
if not exist "auto_geocoder.py" goto missing_source
if not exist "build_release.py" goto missing_builder
if not exist "requirements.txt" goto missing_requirements
where py >nul 2>nul && goto use_py
where python >nul 2>nul && goto use_python
goto no_python

:use_py
echo [INFO] Python launcher found.
py -3 build_release.py
goto finished

:use_python
echo [INFO] Python command found.
python build_release.py
goto finished

:missing_source
echo [ERROR] auto_geocoder.py was not found.
goto failed

:missing_builder
echo [ERROR] build_release.py was not found.
goto failed

:missing_requirements
echo [ERROR] requirements.txt was not found.
goto failed

:no_python
echo [ERROR] Python was not found.
echo Install Python 3.11 or later, then run this file again.
goto failed

:finished
echo.
echo ============================================================
echo  Build process finished.
echo  Check the release folder.
echo ============================================================
if exist "release" start "" "release"
pause
exit /b 0

:failed
echo.
echo Build failed. Check the message above.
pause
exit /b 1
