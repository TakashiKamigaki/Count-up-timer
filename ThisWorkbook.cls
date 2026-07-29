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
    Dim targetSheetIndex As Integer
    Dim ws As Worksheet
    Dim cleanWSName As String
    Dim msg As String
    
    targetSheetIndex = -1
    ThisWorkbook.Protect Structure:=True, Windows:=False
    
    Call Module1.LoadConfig
    
    If Not Module1.UserDict Is Nothing Then
        currentUser = Replace(Replace(Application.userName, " ", ""), "　", "")
        If Module1.UserDict.exists(currentUser) Then
            For Each ws In ThisWorkbook.Sheets
                cleanWSName = Replace(Replace(ws.Name, " ", ""), "　", "")
                If cleanWSName = currentUser Then
                    targetSheetIndex = ws.Index
                    Exit For
                End If
            Next ws
            Set Module1.UserDict = Nothing
            Call LaunchTimerProcess(IsManual:=False, Index:=targetSheetIndex)
        Else
            Application.StatusBar = "【タイマー " & Module1.PROGRAM_VERSION & "】ユーザーは未登録です。タイマーの起動はマクロ一覧(Alt + F8)から行ってください。"
            Call ClearGlobalVariables
        End If
    End If
End Sub
