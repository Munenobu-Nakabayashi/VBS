Option Explicit

Dim oClassSet, oClass, oLocator, oService
Dim szMsg
Dim strFile
Dim objFS
Dim objText

strFile = "C:\cpuusage\cpuusage.log"
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objText = objFS.OpenTextFile(strFile, 8)

Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer
Set oClassSet = oService.ExecQuery("SELECT * FROM Win32_Processor")
For Each oClass In oClassSet
  'MsgBox "CPU Name: " & oClass.Name & vbCrLF & _
  '       "CPU Usage[%]: " & oClass.LoadPercentage & vbCrLf & _
  '       "CPU Clock : " & oClass.CurrentClockSpeed
  objText.WriteLine("CPU Usage[%]: " & oClass.LoadPercentage)
Next

Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing

objText.Close