VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Dim currentUser As String
    Dim allowedUsers As Variant
    Dim user As Variant
    Dim isAllowed As Boolean
    
    ' ★ 1. 現在Excelにサインインしているユーザー名を「一度だけ」取得
    currentUser = Application.UserName
    
    ' 2. 自動起動を「許可」する特定のユーザー名を配列で指定
    allowedUsers = Array("神垣貴誌", "上田花子")
    
    ' 3. 現在のユーザーが許可リストに含まれているかチェック
    isAllowed = False
    For Each user In allowedUsers
        If LCase(Trim(currentUser)) = LCase(Trim(CStr(user))) Then
            isAllowed = True
            Exit For
        End If
    Next user
    
    ' 4. 許可されている場合のみ、ユーザー名を渡して起動処理を呼び出す
    If isAllowed Then
        ' 自動起動なので IsManual:=False
        Call LaunchTimerProcess(IsManual:=False, UserName:=currentUser)
    End If
End Sub
