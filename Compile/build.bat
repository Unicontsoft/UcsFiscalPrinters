@echo off
setlocal
set VbCodeLines="C:\work\BuildTools\VBCodeLines\VbCodeLines.exe"
for %%i in ("%~dp0.") do set file_dir=%%~dpnxi
for %%i in ("%file_dir%\..\Src\.") do set src_dir=%%~dpnxi
for %%i in ("%file_dir%\..\Binn\.") do set bin_dir=%%~dpnxi
set Vb6="%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist %Vb6% set Vb6="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"

echo Copy sources from %src_dir%...
for %%i in (%file_dir%\*.*) do (if not "%%~nxi"=="build.bat" del %%i > nul)
for %%i in (%src_dir%\*.bas;%src_dir%\*.cls;%src_dir%\*.frm;%src_dir%\*.frx;%src_dir%\*.vbp) do (copy %%i > nul)
echo Put lines to sources in %file_dir%...
for %%i in (%file_dir%\*.vbp) do (%VbCodeLines% %%i)
echo Compiling to %bin_dir%...
attrib -r "%bin_dir%\*.*"
for %%i in (%file_dir%\*.vbp) do (%Vb6% /m %%i)
echo Done.