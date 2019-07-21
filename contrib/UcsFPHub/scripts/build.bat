@echo off
setlocal
for %%i in ("%~dp0.") do set file_dir=%%~dpnxi
for %%i in ("%file_dir%\..\src\.") do set src_dir=%%~dpnxi
for %%i in ("%file_dir%\..\..\..\bin\.") do set bin_dir=%%~dpnxi
set output_exe=%bin_dir%\UcsFPHub.exe
set VbCodeLines="%bin_dir%\VbCodeLines.exe"
set Ummm="%bin_dir%\UMMM.exe"
set mt=C:\Work\BuildTools\UMMM\mt.exe -nologo
if not exist %mt% set mt=mt.exe -nologo
set Vb6="%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist %Vb6% set Vb6="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"

echo Cleanup %file_dir%...
for %%i in (%file_dir%\*.*) do (if not "%%~nxi"=="build.bat" if not "%%~nxi"=="UcsFPHub.ini" del %%i > nul)
rd /s /q %file_dir%\Shared
mkdir %file_dir%\Shared

echo Copy sources from %src_dir%...
for %%i in (%src_dir%\*.bas;%src_dir%\*.cls;%src_dir%\*.frm;%src_dir%\*.frx;%src_dir%\*.vbp) do (copy %%i %file_dir% > nul)
for %%i in (%src_dir%\Shared\*.bas;%src_dir%\Shared\*.cls) do (copy %%i %file_dir%\Shared > nul)

echo Put lines to sources in %file_dir%...
for %%i in (%file_dir%\*.vbp) do (%VbCodeLines% %%i)

echo Compiling to %bin_dir%...
for %%i in (%file_dir%\*.vbp) do (%Vb6% /m %%i)

echo Embedding manifest in %output_exe%...
%Ummm% %file_dir%\UcsFPHub.ini
%mt%  -manifest "%file_dir%\UcsFPHub.ini.manifest" -outputresource:"%output_exe%;1"

echo Done.
