@echo off
setlocal
for %%i in ("%~dp0.") do set "file_dir=%%~dpnxi"
for %%i in ("%file_dir%\..\..\src\UcsFP20\.") do set "src_dir=%%~dpnxi"
for %%i in ("%file_dir%\..\..\bin\.") do set "bin_dir=%%~dpnxi"
set "VbCodeLines=%bin_dir%\VbCodeLines.exe"
set "Vb6=%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist "%Vb6%" set "Vb6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
set "log_file=%file_dir%\compile.out"

echo Cleanup %file_dir%...
if exist "%bin_dir%\UcsFP20.dll" start "" /w regsvr32 /u /s "%bin_dir%\UcsFP20.dll"
for %%i in ("%file_dir%\*.*") do (if not "%%~nxi"=="build.bat" del "%%i" > nul)
rd /s /q "%file_dir%\Connectors" 2>&1
mkdir "%file_dir%\Connectors"
rd /s /q "%file_dir%\Protocols" 2>&1
mkdir "%file_dir%\Protocols"
rd /s /q "%file_dir%\Shared" 2>&1
mkdir "%file_dir%\Shared"
echo Copy sources from %src_dir%...
for %%i in ("%src_dir%\*.bas";"%src_dir%\*.cls";"%src_dir%\*.frm";"%src_dir%\*.frx";"%src_dir%\*.vbp") do (copy "%%i" "%file_dir%" > nul)
echo Copy sources from %src_dir%\Connectors...
for %%i in ("%src_dir%\Connectors\*.bas";"%src_dir%\Connectors\*.cls") do (copy "%%i" "%file_dir%\Connectors" > nul)
echo Copy sources from %src_dir%\Protocols...
for %%i in ("%src_dir%\Protocols\*.bas";"%src_dir%\Protocols\*.cls") do (copy "%%i" "%file_dir%\Protocols" > nul)
echo Copy sources from %src_dir%\Shared...
for %%i in ("%src_dir%\Shared\*.bas";"%src_dir%\Shared\*.cls") do (copy "%%i" "%file_dir%\Shared" > nul)
echo Put lines to sources in %file_dir%...
for %%i in ("%file_dir%\*.vbp") do (start "" /w "%VbCodeLines%" %%i)
echo Compiling to %bin_dir%...
attrib -r "%bin_dir%\*.*"
for %%i in ("%file_dir%\*.vbp") do (
    del "%log_file%" > nul 2>&1
    "%Vb6%" /make "%%i" /out "%log_file%" /outdir "%bin_dir%"
    findstr /r /c:"Build of '.*' succeeded" "%log_file%" || (type "%log_file%" 1>&2 & exit /b 1)
)
echo Done.
