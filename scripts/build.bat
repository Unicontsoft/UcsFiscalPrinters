@echo off
setlocal
for %%i in ("%~dp0.") do set file_dir=%%~dpnxi
for %%i in ("%file_dir%\..\src\.") do set src_dir=%%~dpnxi
for %%i in ("%file_dir%\..\bin\.") do set bin_dir=%%~dpnxi
set VbCodeLines="%bin_dir%\VbCodeLines.exe"
set Vb6="%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist %Vb6% set Vb6="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"

echo Cleanup %file_dir%...
for %%i in (%file_dir%\*.*) do (if not "%%~nxi"=="build.bat" del %%i > nul)
rd /s /q %file_dir%\Connectors
mkdir %file_dir%\Connectors
rd /s /q %file_dir%\Protocols
mkdir %file_dir%\Protocols
rd /s /q %file_dir%\Setup
mkdir %file_dir%\Setup
rd /s /q %file_dir%\Shared
mkdir %file_dir%\Shared
echo Copy sources from %src_dir%...
for %%i in (%src_dir%\*.bas;%src_dir%\*.cls;%src_dir%\*.frm;%src_dir%\*.frx;%src_dir%\*.vbp) do (copy %%i %file_dir% > nul)
echo Copy sources from %src_dir%\Connectors...
for %%i in (%src_dir%\Connectors\*.bas;%src_dir%\Connectors\*.cls) do (copy %%i %file_dir%\Connectors > nul)
echo Copy sources from %src_dir%\Protocols...
for %%i in (%src_dir%\Protocols\*.bas;%src_dir%\Protocols\*.cls) do (copy %%i %file_dir%\Protocols > nul)
echo Copy sources from %src_dir%\Setup...
for %%i in (%src_dir%\Setup\*.frm;%src_dir%\Setup\*.frx) do (copy %%i %file_dir%\Setup > nul)
echo Copy sources from %src_dir%\Shared...
for %%i in (%src_dir%\Shared\*.bas;%src_dir%\Shared\*.cls) do (copy %%i %file_dir%\Shared > nul)
echo Put lines to sources in %file_dir%...
for %%i in (%file_dir%\*.vbp) do (%VbCodeLines% %%i)
echo Compiling to %bin_dir%...
attrib -r "%bin_dir%\*.*"
for %%i in (%file_dir%\*.vbp) do (%Vb6% /m %%i)
echo Done.