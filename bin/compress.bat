@echo off
setlocal
set copy_dbg=C:\Work\BuildTools\CopyDebugInfo\CopyDebugInfo.exe
pushd %~dp0\.
md Original 2> nul
copy *.dll Original > nul
upx --lzma -qq -9 --compress-resource=0 *.dll
for %%i in (Original\*.dll) do %copy_dbg% %%i %%~nxi
pause