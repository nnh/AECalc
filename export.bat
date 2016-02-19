@echo off

cscript vbac.wsf decombine /binary . /source macros
bin\nkf32.exe -Lu -w80 --overwrite macros\AECalc.xlsm\*
