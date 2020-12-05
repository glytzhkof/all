# https://stackoverflow.com/questions/65150110/powershell-how-to-use-windowsinstaller-installer-to-insert-a-property-value/65152502#65152502

$WIObject = new-object -comobject WindowsInstaller.Installer
$MSIOpenDatabaseModeTransact = 1
$MSIPath = "C:\Users\User\Desktop\MyTest.msi"

$MSIDB = $WIObject.GetType().InvokeMember(
    	"OpenDatabase", 
    	"InvokeMethod", 
    	$Null, 
    	$WIObject, 
    	@($MSIPath, $MSIOpenDatabaseModeTransact)
    )

$Query1 = "INSERT INTO ``Property`` (``Property``,``Value``) VALUES ('REBOOT','Force')"

$Insert = $MSIDB.GetType().InvokeMember(
    	"OpenView",
    	"InvokeMethod",
    	$Null,
    	$MSIDB,
    	($Query1)
    )

$Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null)		
$Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null)

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

$MSIDB.GetType().InvokeMember("Commit", "InvokeMethod", $Null, $MSIDB, $Null)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($MSIDB) | Out-Null
