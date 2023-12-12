@echo off
call :Main
exit /b 0


:Main
echo "------------ Main ------------"
call :Install_Choco status
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_Resource_Kit_Tools status)
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_Smartmontools status)
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_Python status)
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_RobotFramework status)
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_CSV_Library status) 
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_Selenium_Library status)
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (call :Install_Ocular_Library status)
if %status%==FAIL (echo "Setup environment for robotframework is not successful" && exit /b 0) else (echo "Setup environment for robotframework is successful" && echo "Done")
exit /b 0

::Installing Choco
:Install_Choco
echo "Check choco is installed or not"
choco --version
if %errorlevel%==9009 (echo "Choco is not installed") else (echo "choco is available" && set status=PASS && exit /b 0)
echo "Installing choco..."
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin" 
if %errorlevel%==0 (echo Installing choco is successful && set status=PASS) else (echo "Installing choco is failed" && set status=FAIL && exit /b 0)
timeout /t 3
choco --version
if not %errorlevel%==9009 (echo "Choco is actived" && set status=PASS) else (echo "Choco is not actived. Please turn off and turn on command line to refresh it, then run again" && set status=FAIL)
exit /b 0



::Installing Resource kit tools
:Install_Resource_Kit_Tools
echo "Check resource kit tools is installed or not"
pathman
if %errorlevel%==9009 (echo "Resource kit tools is not installed") else (echo "Resource kit tools is available" && set status=PASS && exit /b 0)
echo "Installing resource kit tools..."
choco install rktools.2003 --yes
if %errorlevel%==0 (echo "Installing resource kit tools is successful" && set status=PASS) else (echo "Installing resource kit tools is failed" && set status=FAIL && exit /b 0)
call "refreshenv"
timeout /t 3
pathman
echo errorlevel: %errorlevel%
if not %errorlevel%==9009 (echo "Resource kit tools is actived" && set status=PASS) else (echo "Resource kit tools is not actived. Please turn off and turn on command line to refresh it, then run again" && set status=FAIL)
exit /b 0

::Installing python
:Install_Python
echo "Check python is installed or not"
python --version
if %errorlevel%==9009 (echo "Python is not installed") else (echo "Python is available" && set status=PASS && exit /b 0)
echo "Installing python..." 
choco install python --yes
if %errorlevel%==0 (echo "Installing python is successful" && set status=PASS && cmd /c "pathman /as C:\Python37\Lib\site-packages\") else (echo "Installing python is failed" && set status=FAIL && exit /b 0)
call "refreshenv"
timeout /t 3
python --version
if not %errorlevel%==9009 (echo "Python is actived" && set status=PASS) else (echo "Python is not actived. Please turn off and turn on command line to refresh it, then run again" && set status=FAIL)
exit /B 0

::Installing robotframework
:Install_RobotFramework
echo "Check robotframework is installed or not"
robot --version
if %errorlevel%==9009 (echo "RobotFramework is not installed") else (echo "RobotFramework is available" && set status=PASS && exit /b 0)
echo "Installing robotframework..." 
pip install robotframework
if %errorlevel%==0 (echo "Installing robotframework is successful" && set status=PASS) else (echo "Installing robotframework is failed" && set status=FAIL && exit /b 0)
call "refreshenv"
timeout /t 3
robot --version
if not %errorlevel%==9009 (echo "RobotFramework is actived" && set status=PASS) else (echo "RobotFramework is not actived. Please turn off and turn on command line to refresh it, then run again" && set status=FAIL)
exit /B 0

::Installing Smartmontools
:Install_Smartmontools
echo "Check smartmontools is installed or not"
smartctl --version
if %errorlevel%==9009 (echo "Smartmontools is not installed") else (echo "Smartmontools is available" && set status=PASS && exit /b 0)
echo "Installing smartmontools..."
choco install smartmontools --yes
if %errorlevel%==0 (echo "Installing smartmontools is successful" && set status=PASS) else (echo "Installing smartmontools is failed" && set status=FAIL && exit /b 0)
call "refreshenv"
timeout /t 3
smartctl --version
if not %errorlevel%==9009 (echo "Smartmontools is actived" && set status=PASS) else (echo "Smartmontools is not actived. Please turn off and turn on command line to refresh it, then run again" && set status=FAIL)
exit /B 0


::Installing csv library
:Install_CSV_Library
echo "Installing csv library..."
pip install robotframework-csvlib==1.0.0
if %errorlevel%==0 (echo "Installing csv library is successful" && set status=PASS) else (echo "Installing csv library is failed" && set status=FAIL)

::Installing selenium library
:Install_Selenium_Library
echo "Installing selenium2 library..."
pip install robotframework-seleniumlibrary
if %errorlevel%==0 (echo "Installing selenium2 library is successful" && set status=PASS) else (echo "Installing selenium2 library is failed" && set status=FAIL)

::Installing ocular library
:Install_Ocular_Library
echo "Installing ocular library..."
pip install ocular
if %errorlevel%==0 (echo "Installing ocular library is successful" && set status=PASS) else (echo "Installing ocular library is failed" && set status=FAIL)
exit /B 0


