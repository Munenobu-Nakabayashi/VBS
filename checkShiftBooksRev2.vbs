rem Option Explicit

Dim fileName
Dim workFile
Dim strPath

Dim workFileName

Dim kyou
Dim asu
Dim asatte
Dim shiasatte	'ADD NEW 2024.04.01 --- 金曜日に翌週の月曜日のチェックをする為
Dim kyouFlg
Dim asuFlg
Dim asatteFlg
dim shiasatteFlg	'ADD NEW 2024.04.01
rem //曜日はいちオリジンで1～7になる
Dim arrayWeekDay(7)
arrayWeekDay(0) = ""
arrayWeekDay(1) = "(Sun)"
arrayWeekDay(2) = "(Mon)"
arrayWeekDay(3) = "(Tue)"
arrayWeekDay(4) = "(Wed)"
arrayWeekDay(5) = "(Thu)"
arrayWeekDay(6) = "(Fri)"
arrayWeekDay(7) = "(Sat)"

kyou = Date()
asu = Date() + 1
asatte = Date() + 2
shiasatte = Date() + 3
kyou = Replace(kyou, "/", "")
asu = Replace(asu, "/", "")
asatte = Replace(asatte, "/", "")
shiasatte = Replace(shiasatte, "/", "")		'ADD NEW 2024.04.01

workFile = "I:\Systems\Public\休日出勤・深夜残業届\旧\CheckList\" & kyou & ".txt"
strPath = "I:\Systems\Public\休日出勤・深夜残業届"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(strPath)

rem //ファイルはその都度初期化: Open Output -> Close -> Open I-Oする
Set objText = objFS.CreateTextFile(workFile)
WScript.Sleep 500
objText.Close
Set objText = objFS.OpenTextFile(workFile, 8)

rem ADD 2024.04.01 Start --- //当VBSの実行タイムスタンプを1行目に明記する 
objText.WriteLine("★ タイムスタンプ ★")
objText.WriteLine("チェック日時:" & now())
objText.WriteLine()
rem ADD 2024.04.01 End

rem //本日日付(yyyymmdd)

objText.WriteLine("本　日の日付:" & Date() & arrayWeekDay(WeekDay(Date())))
objText.WriteLine("明　日の日付:" & Date() + 1 & arrayWeekDay(WeekDay(Date() + 1)))
objText.WriteLine("明後日の日付:" & Date() + 2 & arrayWeekDay(WeekDay(Date() + 2)))
objText.WriteLine("明々後日の日付:" & Date() + 3 & arrayWeekDay(WeekDay(Date() + 3)))	'ADD NEW 2024.04.01
objText.WriteLine()
objText.WriteLine("-----●該　当●-----")

kyouFlg = -1
asuFlg = -1
asatteFlg = -1
shiasatteFlg = -1
For Each objFile In objFolder.Files

	rem objText.WriteLine(objFile.Name)
	workFileName = objFile.Name
	If InStr(workFileName, kyou) > 0 Then
		objText.WriteLine("本　日日付のファイル: " & objFile.Name)
		kyouFlg = 0
	'End if
	ElseIf InStr(workFileName, asu) > 0 Then
		objText.WriteLine("明　日日付のファイル: " & objFile.Name)
		asuFlg = 0
	'End if
	ElseIf InStr(workFileName, asatte) > 0 Then
		objText.WriteLine("明後日日付のファイル: " & objFile.Name)
		asatteFlg = 0
	ElseIf instr(workFileName, shiasatte) > 0 Then	'ADD NEW 2024.04.01
		objText.WriteLine("明々後日日付のファイル: " & objFile.Name)
		shiasatteFlg = 0
	Else
		if instr(workFileName, ".ods") > 0 Then	'拡張子によって識別。無関係ファイルを除外
			'今日、明日、明後日、明々後日以外を対象外日付とし、全数を出力する(ADD NEW --- 2024.04.01)
			objText.WriteLine("対象外日付のファイル: " & objFile.Name)
		end if
	End if

Next

rem //金曜日対応（6である）
IF WeekDay(Date()) = 6 Then
	if asuFlg = 0 Then
		objText.WriteLine("★本日、明　日の土曜日分の対応が必要★")
	End if
	if asatteFlg = 0 Then
		objText.WriteLine("★本日、明後日の日曜日分の対応が必要★")
	End if
	if shiasatteFlg = 0 Then		'
		objText.WriteLine("★明々後日の月曜日分の勤務時間チェックが必要。未明の場合は対応が必要★")
	End if
Else
	if asuFlg = 0 Then
		objText.WriteLine("☆本日、明　日分の対応は不要（※明日が祝日ではない事）☆")
	End if
	if asatteFlg = 0 Then
		objText.WriteLine("☆本日、明後日分の対応は不要（※明後日に対応する事）☆")
	End if
	if shiasatteFlg = 0 Then
		objText.WriteLine("☆本日、明々後日分の対応は不要（※明々後日に対応する事）☆")
	End if
End if 

objText.WriteLine()
objText.WriteLine("-----■非該当■-----")

If kyouFlg <> 0 Then
	objText.WriteLine("本　日日付のファイル: 該当なし")
End if
If asuFlg <> 0 Then
	objText.WriteLine("明　日日付のファイル: 該当なし")
End if
If asatteFlg <> 0 Then
	objText.WriteLine("明後日日付のファイル: 該当なし")
End if
If shiasatteFlg <> 0 Then
	objText.WriteLine("明々後日日付のファイル: 該当なし")
End if

objText.Close

rem End