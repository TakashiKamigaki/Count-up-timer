VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'© 2026 神垣貴誌
Option Explicit

Private Sub Workbook_Open()
    Dim currentUser As String
    Dim allowedUsers As Variant
    Dim user As Variant
    Dim isAllowed As Boolean
    
    currentUser = Application.UserName
    allowedUsers = Array("")  '許可リスト
    isAllowed = False

    For Each user In allowedUsers
        If LCase(Trim(currentUser)) = LCase(Trim(CStr(user))) Then
            isAllowed = True
            Exit For
        End If
    Next user
    
    If isAllowed Then
        Call LaunchTimerProcess(IsManual:=False, UserName:=currentUser)
    End If
End Sub
