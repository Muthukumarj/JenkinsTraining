REM ******************** SET PATHS *******************
set PROJECTPATH=C:\AutomationFramework\eclipseWorkspace\myINSEAD\selenium
set SELENIUMPATH=C:\AutomationFramework\selenium
set POIPATH=C:\AutomationFramework\apachepoi\poi-3.13
set TESTNGPATH=C:\AutomationFramework\testng
set JDKPATH="C:\Program Files\Java\jdk1.8.0_65"
set CLASSPATH=%PROJECTPATH%\bin;%POIPATH%\*;%POIPATH%\ooxml-lib\*;%SELENIUMPATH%\*;%TESTNGPATH%\*

REM ******************** COMPILE JAVA FILES *******************
cd %PROJECTPATH%
mkdir bin
%JDKPATH%\bin\javac src\inseadTesting\*.java -d bin

REM ******************** RUN TESTNG SUITE(S) *******************
java org.testng.TestNG src\inseadTesting\seleniumConfig.xml
