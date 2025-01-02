@echo off
echo test
call test
echo test cd
call test %cd%
echo test cd desktop
call test %cd% C:\Users\syenl\Desktop
echo.

echo cd
echo %cd%
echo cd\..
for %%I in (%cd%\..) do echo %%~fI
echo dp0
echo %~dp0
echo dp0..
for %%I in (%~dp0..) do echo %%~fI
echo f0
echo %~f0
echo f0\..
for %%I in (%~f0\..) do echo %%~fI
echo.

for %%I in (%~f0) do echo I: %%~I
echo Expands %I which removes any surrounding quotation marks.
for %%I in (%~f0) do echo fI: %%~fI
echo Expands %I to a fully qualified path name.
for %%I in (%~f0) do echo dI: %%~dI
echo Expands %I to a drive letter only.
for %%I in (%~f0) do echo pI: %%~pI
echo Expands %I to a path only.
for %%I in (%~f0) do echo nI: %%~nI
echo Expands %I to a file name only.
for %%I in (%~f0) do echo xI: %%~xI
echo Expands %I to a file name extension only.
for %%I in (%~f0) do echo sI: %%~sI
echo Expands path to contain short names only.
for %%I in (%~f0) do echo aI: %%~aI
echo Expands %I to the file attributes of file.
for %%I in (%~f0) do echo tI: %%~tI
echo Expands %I to the date and time of file.
for %%I in (%~f0) do echo zI: %%~zI
echo Expands %I to the size of the file.
echo.

for %%I in (%~f0) do echo dpI: %%~dpI
echo Expands %I to a drive letter and path only.
for %%I in (%~f0) do echo nxI: %%~nxI
echo Expands %I to a file name and extension only.
for %%I in (%~f0) do echo fsI: %%~fsI
echo Expands %I to a full path name with short names only.
for %%I in (%~f0) do echo ftzaI: %%~ftzaI
echo Expands %I to an output line that is like dir.
pause