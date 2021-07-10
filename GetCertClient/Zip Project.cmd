@echo off
for %%I in (.) do set ProjectApp=%%~nxI
if not "%ProjectApp%"=="Windows" goto Continue

echo.
echo Can't run this from a UNC path.
pause
goto EOF


:Continue
cd ..
for %%I in (.) do set Project=%%~nxI
cd %ProjectApp%
set ZipEXE=7za.exe

del                                    "%Project%.zip"
if exist                               "%Project%.zip" pause

echo on
echo %ZipEXE% -bb a -r -xr!bin -xr!obj "%Project%.zip" *.*
     %ZipEXE% -bb a -r -xr!bin -xr!obj "%Project%.zip" *.*

if not exist bin\Release\*.* goto EOF

cd bin\Release

del Setup.zip
del Setup.zzz
%ZipEXE% a Setup.zip "%Project%.exe"
copy Setup.zip *.zzz
