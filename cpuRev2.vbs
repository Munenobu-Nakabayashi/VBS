Option Explicit

Dim oClassSet, oClass, oLocator, oService
Dim szMsg
Dim strFile
Dim objFS
Dim objText

strFile = "C:\cpuusage\cpuusage2.log"
Set objFS = CreateObject("Scripting.FileSystemObject")


Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer
Set oClassSet = oService.ExecQuery("SELECT * FROM Win32_Processor")

Set objText = objFS.OpenTextFile(strFile, 8)

For Each oClass In oClassSet
	'MsgBox "CPU Name: " & oClass.Name & vbCrLF & _
	'"CPU Usage[%]: " & oClass.LoadPercentage & vbCrLf & _
	'"CPU Clock : " & oClass.CurrentClockSpeed
	objText.WriteLine(Now & " CPU Usage[%]: " & oClass.LoadPercentage)
Next

Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing

objText.Close
