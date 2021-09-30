REM @echo off
REM @echo Processing: %1

cscript.exe "%~dp0"\WiRunSQL.vbs "Modify.msi" "INSERT INTO `Property` (`Property`, `Value`) VALUES ('MYPROPERTY', 'PropertyValue')"

pause
