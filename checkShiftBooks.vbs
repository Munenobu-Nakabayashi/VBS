'Option Explicit

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
dim teleworkFlg		'ADD NEW 2024.08.06
dim searchStr		'【在宅】文字列
'曜日はいちオリジンで1～7になる
Dim arrayWeekDay(8)
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

'//ファイルはその都度初期化: Open Output -> Close -> Open I-Oする
Set objText = objFS.CreateTextFile(workFile)
WScript.Sleep 500
objText.Close
Set objText = objFS.OpenTextFile(workFile, 8)

'ADD 2024.04.01 Start --- //当VBSの実行タイムスタンプを1行目に明記する 
objText.WriteLine("★ タイムスタンプ ★")
objText.WriteLine("チェック日時:" & now())
objText.WriteLine()
'ADD 2024.04.01 End

'//本日日付(yyyymmdd)

objText.WriteLine("本　日の日付:" & Date() & arrayWeekDay(WeekDay(Date())))
objText.WriteLine("明　日の日付:" & Date() + 1 & arrayWeekDay(WeekDay(Date() + 1)))
objText.WriteLine("明後日の日付:" & Date() + 2 & arrayWeekDay(WeekDay(Date() + 2)))
objText.WriteLine("明々後日の日付:" & Date() + 3 & arrayWeekDay(WeekDay(Date() + 3)))	'ADD NEW 2024.04.01
objText.WriteLine("☆範囲外の日付:" & Date() + 4 & arrayWeekDay(WeekDay(Date() + 4)) & " 以降の日付☆")
objText.WriteLine()
objText.WriteLine("-----●該　当●-----")

kyouFlg = -1
asuFlg = -1
asatteFlg = -1
shiasatteFlg = -1
taishougaiFlg = -1		'UPDATE --- 2024.05.08
For Each objFile In objFolder.Files

	'objText.WriteLine(objFile.Name)
	workFileName = objFile.Name
	If InStr(workFileName, kyou) > 0 Then
		objText.WriteLine("本　日日付のファイル:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel内シートに【在宅】文字列があるかチェック
		If teleworkFlg = 0 Then
			objText.WriteLine("◆本　日日付のファイル（" & objFile.Name & "）にテレワークする旨の記載あり。記載事項を確認すること！◆")
		End If
		' ADD NEW 2024.08.06 --- End
		kyouFlg = 0
	'End if
	ElseIf InStr(workFileName, asu) > 0 Then
		objText.WriteLine("明　日日付のファイル:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel内シートに【在宅】文字列があるかチェック
		If teleworkFlg = 0 Then
			objText.WriteLine("◆明　日日付のファイル（" & objFile.Name & "）にテレワークする旨の記載あり。記載事項を確認すること！◆")
		End If
		' ADD NEW 2024.08.06 --- End
		asuFlg = 0
	'End if
	ElseIf InStr(workFileName, asatte) > 0 Then
		objText.WriteLine("明後日日付のファイル:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel内シートに【在宅】文字列があるかチェック
		If teleworkFlg = 0 Then
			objText.WriteLine("◆明後日日付のファイル（" & objFile.Name & "）にテレワークする旨の記載あり。記載事項を確認すること！◆")
		End If
		' ADD NEW 2024.08.06 --- End
		asatteFlg = 0
	ElseIf InStr(workFileName, shiasatte) > 0 Then	'ADD NEW 2024.04.01
		objText.WriteLine("明々後日日付のファイル:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel内シートに【在宅】文字列があるかチェック
		If teleworkFlg = 0 Then
			objText.WriteLine("◆明々後日日付のファイル（" & objFile.Name & "）にテレワークする旨の記載あり。記載事項を確認すること！◆")
		End If
		' ADD NEW 2024.08.06 --- End
		shiasatteFlg = 0
	Else	'本日、明日、明後日、明々後日以降のExcelファイルを範囲外日付と見なす
		If InStr(workFileName, ".xlsx") > 0 Then
			'今日、明日、明後日、明々後日以外を範囲外日付と位置付け、全数を出力する(ADD NEW --- 2024.04.01)
			objText.WriteLine("範囲外日付のファイル:" & objFile.Name)	'UPDATE --- 2024.05.06
			' ADD NEW 2024.08.06 --- Start
			teleworkFlg = -1
			teleworkFlg = findTelework(objFile.Name)	'Excel内シートに【在宅】文字列があるかチェック
			If teleworkFlg = 0 Then
				objText.WriteLine("◆範囲外日付のファイル（" & objFile.Name & "）にテレワークする旨の記載あり。記載事項を確認すること！◆")
			End If
		' 	ADD NEW 2024.08.06 --- End
			taishougaiFlg = 0	'UPDATE --- 2024.05.08
		End If
	End if

Next

if kyouFlg = 0 Then
	objText.WriteLine("★★本日、本日分の対応が必要★★")		'UPDATE --- 2024.05.08 表示を追加
End if

'金曜日対応（6である）
If WeekDay(Date()) = 6 Then
	If asuFlg = 0 Then
		objText.WriteLine("★本日、明　日の土曜日分の対応が必要★")
	End If
	If asatteFlg = 0 Then
		objText.WriteLine("★本日、明後日の日曜日分の対応が必要★")
	End If
	if shiasatteFlg = 0 Then
		objText.WriteLine("★明々後日の月曜日分の勤務時間チェックが必要。未明開始の場合は本日に対応が必要★")
	End If
Else	'金曜日(6)以外である
	If asuFlg = 0 Then		'UPDATE --- 2024.05.21 未明勤務開始有無につき文言を追加
		objText.WriteLine("☆本日、明　日分の対応は不要（※明日が祝日ではない事。および未明勤務開始有無を確認する事）☆")
	End If
	if asatteFlg = 0 Then	'明後日（木曜日の明後日は土曜日）
		If WeekDay(Date()) = 5 Then	'UPDATE --- 2024.06.13 木曜日時点において明後日である土曜日のファイルが既に格納済みの場合の対応
			objText.WriteLine("☆本日、明後日の土曜日分の対応は不要（★明日金曜日に対応する事★）☆")
		Else
			objText.WriteLine("☆本日、明後日分の対応は不要（※明後日に対応する事）☆")
		End If
	End if
	if shiasatteFlg = 0 Then	'木曜の明々後日は日曜日、水曜日の明々後日は水曜日である
		If WeekDay(Date()) = 5 Then	'UPDATE --- 2024.06.13 木曜日時点において明々後日である日曜日のファイルが既に格納済みの場合の対応
			objText.WriteLine("☆本日、明々後日の日曜日分の対応は不要（★明日金曜日に対応する事★）☆")
		ElseIf WeekDay(Date()) = 4 Then	'UPDATE --- 2024.06.13 水曜日時点において明々後日である土曜日のファイルが既に格納済みの場合の対応
			objText.WriteLine("☆本日、明々後日の土曜日分の対応は不要（★明後日金曜日に対応する事★）☆")
		Else
			objText.WriteLine("☆本日、明々後日分の対応は不要（明々後日に対応する事）☆")
		End If
	End if
End if 
'UPDATE --- 場所を移動（範囲外日付を最後に置く）
If taishougaiFlg = 0 Then
	objText.WriteLine("☆本日、範囲外日分の対応は不要☆")	'UPDATE --- 2024.05.08 表示を追加
End If

objText.WriteLine()
objText.WriteLine("-----■非該当■-----")

If kyouFlg <> 0 Then
	objText.WriteLine("本　日日付のファイル: 該当なし")
End If
If asuFlg <> 0 Then
	objText.WriteLine("明　日日付のファイル: 該当なし")
End If
If asatteFlg <> 0 Then
	objText.WriteLine("明後日日付のファイル: 該当なし")
End If
If shiasatteFlg <> 0 Then
	objText.WriteLine("明々後日日付のファイル: 該当なし")
End If
If taishougaiFlg <> 0 Then
	objText.WriteLine("範囲外日付のファイル: 該当なし")		'ADD NEW 2024.05.08
End If

Set objFS = Nothing
Set objFolder = Nothing
Set objText = Nothing

'objText.Close

Function findTelework(ByVal fileName)		
	'ADD NEW 2024.08.06 --- Start
	Dim i
	Dim ret

	'Const xlValues = -4163	'意味不明なるも転用する

	Set objExcelApp = CreateObject("Excel.Application")
	'objExcelApp.Visible = True
	objExcelApp.Visible = False
	objExcelApp.Workbooks.Open(strPath & "\" & fileName)
	WScript.Sleep 5000	'5秒待機

	searchStr = "在宅"

	For i = 7 to 31
		ret = 0
		ret = InStr(objExcelApp.WorkSheets("届出").Range("I" & i & ":I" & i), searchStr)
		If ret > 0 Then
			findTelework = 0	'【在宅】文字列があった
			Exit For
		End if
	Next

	If i >= 31 Then
		findTelework = -1		'【在宅】文字列はなかった
	End If

	objExcelApp.Workbooks.Close
	Set objExcelApp = Nothing
	'ADD NEW 2024.08.06 --- End
End Function
'End