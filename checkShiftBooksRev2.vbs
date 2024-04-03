rem Option Explicit

Dim fileName
Dim workFile
Dim strPath

Dim workFileName

Dim kyou
Dim asu
Dim asatte
Dim shiasatte	'ADD NEW 2024.04.01 --- ���j���ɗ��T�̌��j���̃`�F�b�N�������
Dim kyouFlg
Dim asuFlg
Dim asatteFlg
dim shiasatteFlg	'ADD NEW 2024.04.01
rem //�j���͂����I���W����1�`7�ɂȂ�
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

workFile = "I:\Systems\Public\�x���o�΁E�[��c�Ɠ�\��\CheckList\" & kyou & ".txt"
strPath = "I:\Systems\Public\�x���o�΁E�[��c�Ɠ�"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(strPath)

rem //�t�@�C���͂��̓s�x������: Open Output -> Close -> Open I-O����
Set objText = objFS.CreateTextFile(workFile)
WScript.Sleep 500
objText.Close
Set objText = objFS.OpenTextFile(workFile, 8)

rem ADD 2024.04.01 Start --- //��VBS�̎��s�^�C���X�^���v��1�s�ڂɖ��L���� 
objText.WriteLine("�� �^�C���X�^���v ��")
objText.WriteLine("�`�F�b�N����:" & now())
objText.WriteLine()
rem ADD 2024.04.01 End

rem //�{�����t(yyyymmdd)

objText.WriteLine("�{�@���̓��t:" & Date() & arrayWeekDay(WeekDay(Date())))
objText.WriteLine("���@���̓��t:" & Date() + 1 & arrayWeekDay(WeekDay(Date() + 1)))
objText.WriteLine("������̓��t:" & Date() + 2 & arrayWeekDay(WeekDay(Date() + 2)))
objText.WriteLine("���X����̓��t:" & Date() + 3 & arrayWeekDay(WeekDay(Date() + 3)))	'ADD NEW 2024.04.01
objText.WriteLine()
objText.WriteLine("-----���Y�@����-----")

kyouFlg = -1
asuFlg = -1
asatteFlg = -1
shiasatteFlg = -1
For Each objFile In objFolder.Files

	rem objText.WriteLine(objFile.Name)
	workFileName = objFile.Name
	If InStr(workFileName, kyou) > 0 Then
		objText.WriteLine("�{�@�����t�̃t�@�C��: " & objFile.Name)
		kyouFlg = 0
	'End if
	ElseIf InStr(workFileName, asu) > 0 Then
		objText.WriteLine("���@�����t�̃t�@�C��: " & objFile.Name)
		asuFlg = 0
	'End if
	ElseIf InStr(workFileName, asatte) > 0 Then
		objText.WriteLine("��������t�̃t�@�C��: " & objFile.Name)
		asatteFlg = 0
	ElseIf instr(workFileName, shiasatte) > 0 Then	'ADD NEW 2024.04.01
		objText.WriteLine("���X������t�̃t�@�C��: " & objFile.Name)
		shiasatteFlg = 0
	Else
		if instr(workFileName, ".ods") > 0 Then	'�g���q�ɂ���Ď��ʁB���֌W�t�@�C�������O
			'�����A�����A������A���X����ȊO��ΏۊO���t�Ƃ��A�S�����o�͂���(ADD NEW --- 2024.04.01)
			objText.WriteLine("�ΏۊO���t�̃t�@�C��: " & objFile.Name)
		end if
	End if

Next

rem //���j���Ή��i6�ł���j
IF WeekDay(Date()) = 6 Then
	if asuFlg = 0 Then
		objText.WriteLine("���{���A���@���̓y�j�����̑Ή����K�v��")
	End if
	if asatteFlg = 0 Then
		objText.WriteLine("���{���A������̓��j�����̑Ή����K�v��")
	End if
	if shiasatteFlg = 0 Then		'
		objText.WriteLine("�����X����̌��j�����̋Ζ����ԃ`�F�b�N���K�v�B�����̏ꍇ�͑Ή����K�v��")
	End if
Else
	if asuFlg = 0 Then
		objText.WriteLine("���{���A���@�����̑Ή��͕s�v�i���������j���ł͂Ȃ����j��")
	End if
	if asatteFlg = 0 Then
		objText.WriteLine("���{���A��������̑Ή��͕s�v�i��������ɑΉ����鎖�j��")
	End if
	if shiasatteFlg = 0 Then
		objText.WriteLine("���{���A���X������̑Ή��͕s�v�i�����X����ɑΉ����鎖�j��")
	End if
End if 

objText.WriteLine()
objText.WriteLine("-----����Y����-----")

If kyouFlg <> 0 Then
	objText.WriteLine("�{�@�����t�̃t�@�C��: �Y���Ȃ�")
End if
If asuFlg <> 0 Then
	objText.WriteLine("���@�����t�̃t�@�C��: �Y���Ȃ�")
End if
If asatteFlg <> 0 Then
	objText.WriteLine("��������t�̃t�@�C��: �Y���Ȃ�")
End if
If shiasatteFlg <> 0 Then
	objText.WriteLine("���X������t�̃t�@�C��: �Y���Ȃ�")
End if

objText.Close

rem End