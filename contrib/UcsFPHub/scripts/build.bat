@echo off
setlocal
for %%i in ("%~dp0.") do set "file_dir=%%~dpnxi"
for %%i in ("%file_dir%\..\src\.") do set "src_dir=%%~dpnxi"
for %%i in ("%file_dir%\..\..\..\bin\.") do set "bin_dir=%%~dpnxi"
set "output_exe=%bin_dir%\UcsFPHub.exe"
set "VbCodeLines=%bin_dir%\VbCodeLines.exe"
set "Ummm=%bin_dir%\UMMM.exe"
set "mt=C:\Work\BuildTools\UMMM\mt.exe"
set "replace=cscript //nologo C:\Work\BuildTools\WixScripts\Replace.vbs"
set "codesign=call C:\Work\BuildTools\Certificates\codesign.bat"
if not exist "%mt%" set "mt=mt.exe"
set "Vb6=%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist "%Vb6%" set "Vb6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
set "log_file=%file_dir%\compile.out"
for /f %%i in ('git rev-parse --short HEAD') do set "HeadCommit=%%i"

echo Cleanup %file_dir%...
for %%i in ("%file_dir%\*.*") do (if not "%%~nxi"=="build.bat" if not "%%~nxi"=="UcsFPHub.ini" del "%%i" > nul)
rd /s /q "%file_dir%\Shared" 2>&1
mkdir "%file_dir%\Shared"

echo Copy sources from %src_dir%...
for %%i in ("%src_dir%\*.bas";"%src_dir%\*.cls";"%src_dir%\*.frm";"%src_dir%\*.frx";"%src_dir%\*.vbp") do (copy "%%i" "%file_dir%" > nul)
for %%i in ("%src_dir%\Shared\*.bas";"%src_dir%\Shared\*.cls";"%src_dir%\Shared\*.ctl") do (copy "%%i" "%file_dir%\Shared" > nul)

echo Put lines to sources in %file_dir%...
for %%i in ("%file_dir%\*.vbp") do (start "" /w "%VbCodeLines%" %%i)

echo Embed commit %HeadCommit% as STR_LATEST_COMMIT in mdStartup.bas...
%replace% /f:%file_dir%\mdStartup.bas /s:"Private Const STR_LATEST_COMMIT         As String = ^q^q" /r:"Private Const STR_LATEST_COMMIT         As String = ^q-%HeadCommit%^q"

echo Compiling to %bin_dir%...
for %%i in ("%file_dir%\*.vbp") do (
    del "%log_file%" > nul 2>&1
    "%Vb6%" /make "%%i" /out "%log_file%" /outdir "%bin_dir%"
    findstr /r /c:"Build of '.*' succeeded" "%log_file%" || (type "%log_file%" 1>&2 & exit /b 1)
)

echo Embedding manifest in %output_exe%...
"%Ummm%" "%file_dir%\UcsFPHub.ini"
"%mt%" -nologo -manifest "%file_dir%\UcsFPHub.ini.manifest" -outputresource:"%output_exe%;1"


echo Code-signing %output_exe%...
%codesign% /d "Unicontsoft Fiscal Printers Hub" %output_exe%

git checkout -- "%bin_dir%\..\src\Shared\mdJson.bas"

echo Done.
