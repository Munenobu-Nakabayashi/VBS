Option Explicit
On Error Resume Next

Dim objWshNetwork   ' WshNetwork �I�u�W�F�N�g
Dim objDrives       ' �h���C�u���
Dim lngLoop         ' ���[�v�J�E���^

Set objWshNetwork = WScript.CreateObject("WScript.Network")
If Err.Number = 0 Then
    Set objDrives = objWshNetwork.EnumNetworkDrives
    If Err.Number = 0 Then
        If objDrives.Count > 0 Then
            WScript.Echo "�l�b�g���[�N�h���C�u�F"
            For lngLoop = 0 To objDrives.Count - 1 Step 2
                WScript.Echo "�@" & objDrives.Item(lngLoop) & " => " & objDrives.Item(lngLoop + 1)
            Next
        Else
            WScript.Echo "�ڑ�����Ă��܂���B"
        End If
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If
Else
    WScript.Echo "�G���[: " & Err.Description
End If

Set objWshNetwork = Nothing
