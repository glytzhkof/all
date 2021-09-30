REM @echo off
REM @echo Processing: %1

cscript.exe "%~dp0"\WiRunSQL.vbs "MySetup.msi" "INSERT INTO `ControlCondition` (`Dialog_`, `Control_`, `Action`, `Condition`) VALUES ('LicenseAgreementDlg', 'Back', 'Hide', '1')"

pause
