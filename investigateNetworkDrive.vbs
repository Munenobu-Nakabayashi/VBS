Option Explicit
On Error Resume Next

Dim objWshNetwork   ' WshNetwork オブジェクト
Dim objDrives       ' ドライブ情報
Dim lngLoop         ' ループカウンタ

Set objWshNetwork = WScript.CreateObject("WScript.Network")
If Err.Number = 0 Then
    Set objDrives = objWshNetwork.EnumNetworkDrives
    If Err.Number = 0 Then
        If objDrives.Count > 0 Then
            WScript.Echo "ネットワークドライブ："
            For lngLoop = 0 To objDrives.Count - 1 Step 2
                WScript.Echo "　" & objDrives.Item(lngLoop) & " => " & objDrives.Item(lngLoop + 1)
            Next
        Else
            WScript.Echo "接続されていません。"
        End If
    Else
        WScript.Echo "エラー: " & Err.Description
    End If
Else
    WScript.Echo "エラー: " & Err.Description
End If

Set objWshNetwork = Nothing
