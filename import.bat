@echo off
mkdir  macros\tmp\AECalc.xlsm
for %%i in (macros\AECalc.xlsm\*) do (
  bin\nkf32.exe -Lw --oc=Shift_JIS %%i >  macros\tmp\AECalc.xlsm\%%~nxi
)
cscript vbac.wsf combine /binary . /source macros\tmp
rmdir /Q /S macros\tmp
