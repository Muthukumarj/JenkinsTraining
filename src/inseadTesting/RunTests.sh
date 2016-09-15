#!/bin/sh

# ******************** SET PATHS *******************
PROJECTPATH=~/testing/eclipseWorkspace/myINSEAD/selenium
SELENIUMPATH=~/testing/selenium
POIPATH=~/testing/apachepoi/poi-3.13
TESTNGPATH=~/testing/testng
export CLASSPATH=$PROJECTPATH/bin:$POIPATH/*:$POIPATH/ooxml-lib/*:$SELENIUMPATH/*:$TESTNGPATH/*

# ******************** COMPILE JAVA FILES *******************
cd $PROJECTPATH
mkdir bin
javac src/inseadTesting/*.java -d bin

# ******************** RUN TESTNG SUITE(S) *******************
java org.testng.TestNG src/inseadTesting/seleniumConfig.xml
