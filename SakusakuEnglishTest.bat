@echo off
cd /d %~dp0
setlocal enabledelayedexpansion

del /Q %~dp0_VocabularyTest

for %%i in (VocabularyBook\*.csv) do (

	python Sakusaku_VocabularyTest.py -s %%i -c 50

)
