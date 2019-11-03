set Project=GetCert2
set  ZipEXE=7za.exe

for /f "tokens=1-9 delims=\" %%a in ("%cd%") do set a=%%a %%b %%c %%d %%e %%f %%g %%h %%i
for %%i in (%a%) do set ProjectFolder=%%i
if not "%ProjectFolder%"=="Windows" goto Continue

@echo off
echo.
echo Can't run this from a UNC path.
pause
goto EOF


:Continue
if "%Project%"=="" set Project=%ProjectFolder%

del                                    %Project%.zip
echo %ZipEXE% -bb a -r -xr!bin -xr!obj %Project%.zip *.*
     %ZipEXE% -bb a -r -xr!bin -xr!obj %Project%.zip *.*

if not exist bin\Release\*.* goto EOF

cd bin\Release

del Setup.zip
del Setup.zzz
%ZipEXE% a Setup.zip %Project%.exe
copy Setup.zip *.zzz
