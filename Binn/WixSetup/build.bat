@echo off

rem ------ init local vars
setlocal
if "%wix_dir%"=="" set wix_dir=C:\WiX
if "%wix_scripts%"=="" set wix_scripts=C:\Work\BuildTools\WixScripts
pushd %~dp0\.
rem goto :compile

rem ------ get executable product version
echo Get product version...
cscript /nologo "%wix_scripts%\GetVersion.vbs" /i:..\UcsFP10.dll > Version.wxi


rem ------ extract registry info
call "%wix_scripts%\extract_reg.bat" ..\UcsFP10.dll

rem ------ compile and link
:compile
echo Compile setup...
%wix_dir%\candle.exe -nologo UcsFiscalPrinter.wxs
if errorlevel 1 goto :end
echo Link setup...
%wix_dir%\light.exe -nologo -out UcsFP10.msm UcsFiscalPrinter.wixobj
popd
:end
pause