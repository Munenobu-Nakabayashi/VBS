'Option Explicit

Dim fileName
Dim workFile
Dim strPath

Dim workFileName

Dim kyou
Dim asu
Dim asatte
Dim shiasatte		'ADD NEW 2024.04.01 --- ���j���ɗ��T�̌��j���̃`�F�b�N�������
Dim kyouFlg
Dim asuFlg
Dim asatteFlg
Dim shiasatteFlg	'ADD NEW 2024.04.01	<--- �G�C�v�����t�[�����
Dim teleworkFlg		'ADD NEW 2024.08.06 <--- �L���̓����
Dim mimeiFlg		'ADD NEW 2024.08.16 <--- ��{��S�̓���Łi����҂̌�ꓰ�₪�ȁB�Ӗ��[�������邪�ȁB����ڂ��������� �ڂ��͂���݁[�j
Dim searchStr		'�y�ݑ�z������
'�j���͂����I���W����1�`7�ɂȂ�
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

workFile = "I:\Systems\Public\�x���o�΁E�[��c�Ɠ�\��\CheckList\" & kyou & ".txt"
strPath = "I:\Systems\Public\�x���o�΁E�[��c�Ɠ�"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(strPath)

'//�t�@�C���͂��̓s�x������: Open Output -> Close -> Open I-O����
Set objText = objFS.CreateTextFile(workFile)
WScript.Sleep 500
objText.Close
Set objText = objFS.OpenTextFile(workFile, 8)

'ADD 2024.04.01 Start --- //��VBS�̎��s�^�C���X�^���v��1�s�ڂɖ��L���� 
objText.WriteLine("�� �^�C���X�^���v ��")
objText.WriteLine("�`�F�b�N����:" & now())
objText.WriteLine()
'ADD 2024.04.01 End

objText.WriteLine("�� �֘A���t�ꗗ ��")
objText.WriteLine("�{�@���̓��t:" & Date() & arrayWeekDay(WeekDay(Date())))
objText.WriteLine("���@���̓��t:" & Date() + 1 & arrayWeekDay(WeekDay(Date() + 1)))
objText.WriteLine("������̓��t:" & Date() + 2 & arrayWeekDay(WeekDay(Date() + 2)))
objText.WriteLine("���X����̓��t:" & Date() + 3 & arrayWeekDay(WeekDay(Date() + 3)))	'ADD NEW 2024.04.01
objText.WriteLine("���͈͊O�̓��t:" & Date() + 4 & arrayWeekDay(WeekDay(Date() + 4)) & " �ȍ~�̓��t��")
objText.WriteLine()
objText.WriteLine("-----���Y�@����-----")

kyouFlg = -1
asuFlg = -1
asatteFlg = -1
shiasatteFlg = -1
taishougaiFlg = -1		'UPDATE --- 2024.05.08
For Each objFile In objFolder.Files

	'objText.WriteLine(objFile.Name)
	workFileName = objFile.Name
	If InStr(workFileName, kyou) > 0 Then
		objText.WriteLine("�{�@�����t�̃t�@�C��:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel���V�[�g�Ɂy�ݑ�z�����񂪂��邩�`�F�b�N
		If teleworkFlg = 0 Then
			objText.WriteLine("���{�@�����t�̃t�@�C���i" & objFile.Name & "�j�Ƀe�����[�N����|�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.06 --- End
		' ���{�����t�t�@�C���ɂ����関���J�n�`�F�b�N�͎��{���Ȃ��B���R�͎��ۂ̃`�F�b�N���{������16:00�ł���A���ɉߋ��������ԑтł��邽��
		kyouFlg = 0
	'End if
	ElseIf InStr(workFileName, asu) > 0 Then
		objText.WriteLine("���@�����t�̃t�@�C��:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel���V�[�g�Ɂy�ݑ�z�����񂪂��邩�`�F�b�N
		If teleworkFlg = 0 Then
			objText.WriteLine("�����@�����t�̃t�@�C���i" & objFile.Name & "�j�Ƀe�����[�N����|�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.06 --- End
		' ADD NEW 2024.08.16 --- Start
		mimeiFlg = -1
		mimeiFlg = findMimeiStart(objFile.Name)		'�J�n�����������̂��̂����邩�`�F�b�N
		if mimeiFlg = 0 Then						'���������J�n�͖{���Ή���v����i�e�����[�N�łȂ�����j
			objText.WriteLine("�������@�����t�̃t�@�C���i" & objFile.Name & "�j�ɖ����J�n�̋L�ڂ���B�{���Ή���v����\����ɂ��L�ڎ������m�F���邱�ƁI����")
		End If
		' ADD NEW 2024.08.16 --- End
		asuFlg = 0
	'End if
	ElseIf InStr(workFileName, asatte) > 0 Then
		objText.WriteLine("��������t�̃t�@�C��:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel���V�[�g�Ɂy�ݑ�z�����񂪂��邩�`�F�b�N
		If teleworkFlg = 0 Then
			objText.WriteLine("����������t�̃t�@�C���i" & objFile.Name & "�j�Ƀe�����[�N����|�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.06 --- End
		' ADD NEW 2024.08.16 --- Start
		mimeiFlg = -1
		mimeiFlg = findMimeiStart(objFile.Name)		'�J�n�����������̂��̂����邩�`�F�b�N
		if mimeiFlg = 0 Then
			objText.WriteLine("����������t�̃t�@�C���i" & objFile.Name & "�j�ɖ����J�n�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.16 --- End
		asatteFlg = 0
	ElseIf InStr(workFileName, shiasatte) > 0 Then	'ADD NEW 2024.04.01
		objText.WriteLine("���X������t�̃t�@�C��:" & objFile.Name)
		' ADD NEW 2024.08.06 --- Start
		teleworkFlg = -1
		teleworkFlg = findTelework(objFile.Name)	'Excel���V�[�g�Ɂy�ݑ�z�����񂪂��邩�`�F�b�N
		If teleworkFlg = 0 Then
			objText.WriteLine("�����X������t�̃t�@�C���i" & objFile.Name & "�j�Ƀe�����[�N����|�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.06 --- End
		' ADD NEW 2024.08.16 --- Start
		mimeiFlg = -1
		mimeiFlg = findMimeiStart(objFile.Name)		'�J�n�����������̂��̂����邩�`�F�b�N
		if mimeiFlg = 0 Then
			objText.WriteLine("�����X������t�̃t�@�C���i" & objFile.Name & "�j�ɖ����J�n�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.16 --- End
		shiasatteFlg = 0
	Else	'�{���A�����A������A���X����ȍ~��Excel�t�@�C����͈͊O���t�ƌ��Ȃ�
		If InStr(workFileName, ".xlsx") > 0 Then
			'�����A�����A������A���X����ȊO��͈͊O���t�ƈʒu�t���A�S�����o�͂���(ADD NEW --- 2024.04.01)
			objText.WriteLine("�͈͊O���t�̃t�@�C��:" & objFile.Name)	'UPDATE --- 2024.05.06
			' ADD NEW 2024.08.06 --- Start
			teleworkFlg = -1
			teleworkFlg = findTelework(objFile.Name)	'Excel���V�[�g�Ɂy�ݑ�z�����񂪂��邩�`�F�b�N
			If teleworkFlg = 0 Then
				objText.WriteLine("���͈͊O���t�̃t�@�C���i" & objFile.Name & "�j�Ƀe�����[�N����|�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
			End If
		' ADD NEW 2024.08.06 --- End
		' ADD NEW 2024.08.16 --- Start
		mimeiFlg = -1
		mimeiFlg = findMimeiStart(objFile.Name)		'�J�n�����������̂��̂����邩�`�F�b�N
		if mimeiFlg = 0 Then
			objText.WriteLine("���͈͊O���t�̃t�@�C���i" & objFile.Name & "�j�ɖ����J�n�̋L�ڂ���B�L�ڎ������m�F���邱�ƁI��")
		End If
		' ADD NEW 2024.08.16 --- End
			taishougaiFlg = 0	'UPDATE --- 2024.05.08
		End If
	End if

Next

objText.WriteLine()
objText.WriteLine("-----���Ή��v�ہ�-----")
if kyouFlg = 0 Then
	objText.WriteLine("�������{���A�{�����̑Ή����K�v������")		'UPDATE --- 2024.05.08 �\����ǉ�
End if

'���j���Ή��i6�ł���j
If WeekDay(Date()) = 6 Then
	If asuFlg = 0 Then
		objText.WriteLine("�������{���A���@���̓y�j�����̑Ή����K�v������")
	End If
	If asatteFlg = 0 Then
		objText.WriteLine("�������{���A������̓��j�����̑Ή����K�v������")
	End If
	if shiasatteFlg = 0 Then
		objText.WriteLine("�������{���A���X����̌��j�����̋Ζ����ԃ`�F�b�N���K�v�B�����J�n�̏ꍇ�͖{���ɑΉ����K�v������")
	End If
Else	'���j��(6)�ȊO�ł���
	If asuFlg = 0 Then		'UPDATE --- 2024.05.21 �����Ζ��J�n�L���ɂ�������ǉ�
		objText.WriteLine("���{���A���@�����̑Ή��͕s�v�i���������j���ł͂Ȃ����B����і����Ζ��J�n�L�����m�F���鎖�j��")
	End If
	if asatteFlg = 0 Then	'������i�ؗj���̖�����͓y�j���j
		If WeekDay(Date()) = 5 Then	'UPDATE --- 2024.06.13 �ؗj�����_�ɂ����Ė�����ł���y�j���̃t�@�C�������Ɋi�[�ς݂̏ꍇ�̑Ή�
			objText.WriteLine("���{���A������̓y�j�����̑Ή��͕s�v�i���������j���ɑΉ����鎖���j��")
		Else
			objText.WriteLine("���{���A��������̑Ή��͕s�v�i��������ɑΉ����鎖�j��")
		End If
	End if
	if shiasatteFlg = 0 Then	'�ؗj�̖��X����͓��j���A���j���̖��X����͐��j���ł���
		If WeekDay(Date()) = 5 Then	'UPDATE --- 2024.06.13 �ؗj�����_�ɂ����Ė��X����ł�����j���̃t�@�C�������Ɋi�[�ς݂̏ꍇ�̑Ή�
			objText.WriteLine("���{���A���X����̓��j�����̑Ή��͕s�v�i���������j���ɑΉ����鎖���j��")
		ElseIf WeekDay(Date()) = 4 Then	'UPDATE --- 2024.06.13 ���j�����_�ɂ����Ė��X����ł���y�j���̃t�@�C�������Ɋi�[�ς݂̏ꍇ�̑Ή�
			objText.WriteLine("���{���A���X����̓y�j�����̑Ή��͕s�v�i����������j���ɑΉ����鎖���j��")
		Else
			objText.WriteLine("���{���A���X������̑Ή��͕s�v�i���X����ɑΉ����鎖�j��")
		End If
	End if
End if 
'UPDATE --- �ꏊ���ړ��i�͈͊O���t���Ō�ɒu���j
If taishougaiFlg = 0 Then
	objText.WriteLine("���{���A�͈͊O�����̑Ή��͕s�v��")	'UPDATE --- 2024.05.08 �\����ǉ�
End If

objText.WriteLine()
objText.WriteLine("-----����Y����-----")

If kyouFlg <> 0 Then
	objText.WriteLine("�{�@�����t�̃t�@�C��: �Y���Ȃ�")
End If
If asuFlg <> 0 Then
	objText.WriteLine("���@�����t�̃t�@�C��: �Y���Ȃ�")
End If
If asatteFlg <> 0 Then
	objText.WriteLine("��������t�̃t�@�C��: �Y���Ȃ�")
End If
If shiasatteFlg <> 0 Then
	objText.WriteLine("���X������t�̃t�@�C��: �Y���Ȃ�")
End If
If taishougaiFlg <> 0 Then
	objText.WriteLine("�͈͊O���t�̃t�@�C��: �Y���Ȃ�")		'ADD NEW 2024.05.08
End If

Set objFS = Nothing
Set objFolder = Nothing
Set objText = Nothing

'objText.Close

Function findTelework(ByVal fileName)		
	'ADD NEW 2024.08.06 --- Start
	Dim i
	Dim ret

	'Const xlValues = -4163	'�Ӗ��s���Ȃ���]�p����

	Set objExcelApp = CreateObject("Excel.Application")
	'objExcelApp.Visible = True
	objExcelApp.Visible = False
	objExcelApp.Workbooks.Open(strPath & "\" & fileName)
	WScript.Sleep 3000	'3�b�ҋ@

	searchStr = "�ݑ�"

	For i = 7 to 31	'[�͏o]�y��[�͏o_2]�V�[�gI7�`I31�Z���ɂ����āy�ݑ�z����������m����d�g��
		ret = 0
		ret = InStr(objExcelApp.WorkSheets("�͏o").Range("I" & i & ":I" & i), searchStr)
		If ret > 0 Then
			findTelework = 0	'�y�ݑ�z�����񂪂�����
			Exit For			'�o�m������A�����𒆒f������
		End if
		'UPDATE 2024.08.14 [�͏o_2]�V�[�g�Ή� --- Start
		ret = InStr(objExcelApp.WorkSheets("�͏o_2").Range("I" & i & ":I" & i), searchStr)
		If ret > 0 Then
			findTelework = 0	'�y�ݑ�z�����񂪂�����
			Exit For			'�o�m������A�����𒆒f������
		End if	
		'UPDATE 2024.08.14 [�͏o_2]�V�[�g�Ή� --- End
	Next

	If i >= 31 Then
		findTelework = -1		'�y�ݑ�z������͂Ȃ�����
	End If

	objExcelApp.Workbooks.Close
	Set objExcelApp = Nothing
	'ADD NEW 2024.08.06 --- End
End Function

Function findMimeiStart(ByVal fileName)
	'ADD NEW 2024.08.16 --- Start
	Dim i

	'�őP�ƌ����Ȃ����iyield�����Ȃ��߂������o����l�������ł��邩��j�y�ݑ�z�������o�Ƃ͈قȂ�֐���p���Ė����J�n�����o������
	Set objExcelApp = CreateObject("Excel.Application")
	'objExcelApp.Visible = True
	objExcelApp.Visible = False
	objExcelApp.Workbooks.Open(strPath & "\" & fileName)
	WScript.Sleep 3000	'3�b�ҋ@

	i = 7	'�J�n�����Z���ł���ԒnE7����E9�AE11�cE31�܂ł��`�F�b�N���A������9�����ł���ꍇ�͖����J�n�ƌ��Ȃ��i�y�ߑO�z�������񂪂���ꍇ�͌��o�ΏۊO�P�[�X�ŗǂ��j
	Do Until i > 31
		if IsNumeric(objExcelApp.WorkSheets("�͏o").Range("E" & i & ":E" & i)) = True And Trim(objExcelApp.WorkSheets("�͏o").Range("E" & i & ":E" & i)) <> "" Then
			if CInt(objExcelApp.WorkSheets("�͏o").Range("E" & i & ":E" & i)) < 9 Then		'09:00�ȑO�������Ė����J�n�ƌ��Ȃ�
				findMimeiStart = 0
				Exit Do
			End If
		ElseIf IsNumeric(objExcelApp.WorkSheets("�͏o_2").Range("E" & i & ":E" & i)) = True And Trim(objExcelApp.WorkSheets("�͏o_2").Range("E" & i & ":E" & i)) <> "" Then
			if CInt(objExcelApp.WorkSheets("�͏o_2").Range("E" & i & ":E" & i)) < 9 Then	'09:00�ȑO�������Ė����J�n�ƌ��Ȃ�
				findMimeiStart = 0
				Exit Do
			End If
		End If
		i = i + 2	'2���J�E���g�A�b�v����i7, 9, 11...31�j
	Loop

	if i >= 31 Then
		findMimeiStart = -1		'�����J�n�͂Ȃ�����
	End If

	objExcelApp.Workbooks.Close
	Set objExcelApp = Nothing
	'ADD NEW 2024.08.16 --- End

End Function
'End