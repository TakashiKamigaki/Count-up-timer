Option Explicit

'定数：自動起動許可ユーザーリスト（カンマ区切りで指定）
Private Const ALLOWED_USER_LIST As String = "山田 太郎,佐藤 花子,Suzuki Ichiro"

Private Sub Workbook_Open()
    Dim currentUser As String
    Dim allowedUsers() As String
    Dim user As Variant
    Dim isAllowed As Boolean
    
    currentUser = Application.UserName
    allowedUsers = Split(ALLOWED_USER_LIST, ",")
    isAllowed = False
    For Each user In allowedUsers
        If LCase(Trim(currentUser)) = LCase(Trim(CStr(user))) Then
            isAllowed = True
            Exit For
        End If
    Next user
    
    If isAllowed Then
        Call LaunchTimerProcess(IsManual:=False, UserName:=currentUser)
    Else
        Application.StatusBar = "【タイマー】ユーザー名「" & currentUser & "」は未登録です。タイマーの起動はマクロ一覧 [Alt] + [F8] から実行してください。"
    End If
End Sub
