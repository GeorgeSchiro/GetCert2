@echo off
for %%I in (.) do set ProjectApp=%%~nxI
if not "%ProjectApp%"=="Windows" goto Continue

echo.
echo Can't run this from a UNC path.
pause
goto EOF


:Continue
echo on
del *.user
del *.suo
del /ah *.suo
del bin\*.cmd         /s/q
del bin\*.config      /s/q
del bin\*.dll         /s/q
del bin\*.application /s/q
del bin\*.manifest    /s/q
del bin\*.txt         /s/q
del bin\*.xml         /s/q
del bin\*.pdb         /s/q
del bin\*.vshost.*    /s/q
del bin\Release\*.z*  /s/q
del bin\Release\Setup*.exe /s/q
rd  bin\Debug         /s/q
rd  obj               /s/q
rd  .vs               /s/q
