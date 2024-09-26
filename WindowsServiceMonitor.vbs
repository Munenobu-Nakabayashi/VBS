Option Explicit

Dim objLocator
Dim objConSvr
Dim objService
Dim objSvc
Dim objWSH
Dim objFS
Dim objText
Dim strServiceName
Dim strState
Dim strStartMode
Dim workFile
Dim strNow

Dim modeFLG
Dim statFLG
Dim retStat

Dim arrayDaemon(9)		'対象のデーモンを限定し配列を使用する
arrayDaemon(0) = "gupdate"
arrayDaemon(1) = "vmicrdv"
arrayDaemon(2) = "vmcompute"
arrayDaemon(3) = "VSS"
arrayDaemon(4) = "SNMPTRAP"
arrayDaemon(5) = "ssh-agent"
arrayDaemon(6) = "W32Time"
arrayDaemon(7) = "XblGameSave"
arrayDaemon(8) = "XboxGipSvc"
arrayDaemon(9) = "XboxNetApiSvc"
Dim arrayService

Function createLogFile()
	workFile = ".\ServiceMonitoring.log"
	Set objFS = CreateObject("Scripting.FileSystemObject")

	Set objText = objFS.CreateTextFile(workFile)
	WScript.Sleep 500
	objText.Close
	Set objText = objFS.OpenTextFile(workFile, 8)
End Function

Function main()

	createLogFile()


	For Each arrayService In arrayDaemon

		strServiceName=arrayService

		Set objLocator=WScript.CreateObject("WbemScripting.SWbemLocator")
		Set objConSvr=objLocator.ConnectServer
		'Set objService=objConSvr.ExecQuery("Select * From Win32_Service Where DisplayName='" & strServiceName & "'")	'Del 2024.09.26
		Set objService=objConSvr.ExecQuery("Select * From Win32_Service Where Name='" & strServiceName & "'")			'Mod 2024.09.26

		statFLG = -1
		modeFLG = -1

		strNow = Now
		If objService.Count=0 Then
			objText.WriteLine(strNow & ": " & arrayService & ": is Unknown Daemon !!")
		Else
			For Each objSvc In objService

				If objSvc.State="Running" Then
					objText.WriteLine(strNow & ": " & arrayService & ":	Running Status.")
				ElseIf objSvc.State="Stopped" Then
					objText.WriteLine(strNow & ": " & arrayService & ":	Stopped Status.")
					statFLG = 0
				ElseIf objSvc.State="Paused" Then
					objText.WriteLine(strNow & ": " & arrayService & ":	Paused Status.")		'中断やて。聞いた事がないステータスやで
				Else
					objText.WriteLine(strNow & ": " & arrayService & ":	Unknown Status ?!")
				End If

				If objSvc.StartMode="Auto" Then
					objText.WriteLine(strNow & ": " & arrayService & ":	Start Mode is Auto.")
					modeFLG = 0
				ElseIf objSvc.StartMode="Manual" Then
					objText.WriteLine(strNow & ": " & arrayService & ":	Start Mode is Manual.")
					'modeFLG = 0
				ElseIf objSvc.StartMode="Disabled" Then
					objText.WriteLine(strNow & ": " & arrayService & ":	Start Mode is Disabled.")
				Else
					objText.WriteLine(strNow & ": " & arrayService & ":	Start Mode is Unknown !!")
				End If

				'Msgbox("サービス名：" & strServiceName & vbCrLf & "サービス状態：" & strState & vbCrLf & "スタートアップの種類：" & strStartMode)

				retStat = -1
				If statFLG = 0 Then			'ステータスが停止で、かつ、
					If modeFLG = 0 Then		'開始モードが自動の場合
						retStat = objSvc.StartService()
						strNow = Now
						If retStat = 0 Then
							objText.WriteLine(strNow & ": " & arrayService & ":	Restarting is Succeeded !!")
						Else
							objText.WriteLine(strNow & ": " & arrayService & ":	Restarting is Failed...")
						End If
					End If
				End If

			Next
		End If

		set objLocator=Nothing
		set objConSvr=Nothing
		set objService=Nothing
		set objSvc=Nothing
		set objWSH=Nothing
	
	Next

End Function

main()

set objText = Nothing
