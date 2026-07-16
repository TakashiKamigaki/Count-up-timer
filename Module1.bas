Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

' 手動起動用マクロ
Sub ShowTimerForm()
    Application.StatusBar = False
    Call LaunchTimerProcess(IsManual:=True)
End Sub

' 実際の起動処理
Public Sub LaunchTimerProcess(ByVal IsManual As Boolean, Optional ByVal UserName As String = "")
    Dim frm As Object
    Dim b1Value As String
    Dim initMs As Long
    Dim initText As String
    Dim ws As Worksheet
    Dim foundSheet As Worksheet
    Dim targetWS As Worksheet
    
    For Each frm In UserForms
        If frm.Name = "UserForm1" Then Exit Sub
    Next frm
    
    If IsManual Then
        Set targetWS = ActiveSheet
    Else
        For Each ws In ThisWorkbook.Worksheets
            If InStr(1, ws.Name, UserName, vbTextCompare) > 0 Then
                Set foundSheet = ws
                Exit For
            End If
        Next ws
        
        If Not foundSheet Is Nothing Then
            Set targetWS = foundSheet
            On Error Resume Next
            targetWS.Activate
            On Error GoTo 0
        Else
            Set targetWS = ActiveSheet
        End If
    End If
    
    b1Value = Trim(targetWS.Range("B1").Text)
    
    If Not IsManual And b1Value Like "##:##:##.###" Then
        Application.StatusBar = "【タイマー】自動起動をスキップしました。"
        Exit Sub
    End If
    
    If b1Value Like "##:##:##.###" Then
        initMs = ParseToMillisec(b1Value)
        initText = b1Value
    Else
        initMs = 0
        initText = "00:00:00.000"
    End If
    
    Load UserForm1
    Call UserForm1.SetInitialTime(initMs, initText)
    UserForm1.Show vbModeless
End Sub

' ユーザーフォームのループを起動するための中継マクロ
Sub TriggerFormLoop(Optional Dummy As Byte = 0)
    On Error Resume Next
    UserForm1.StartTimer
    On Error GoTo 0
End Sub

' 文字列をミリ秒に逆変換する関数
Private Function ParseToMillisec(ByVal timeStr As String) As Long
    On Error GoTo ErrorHandler
    Dim parts() As String, timeParts() As String
    Dim h As Long, m As Long, s As Long, ms As Long
    parts = Split(timeStr, ".")
    ms = CLng(parts(1))
    timeParts = Split(parts(0), ":")
    h = CLng(timeParts(0))
    m = CLng(timeParts(1))
    s = CLng(timeParts(2))
    ParseToMillisec = (h * 3600000) + (m * 60000) + (s * 1000) + ms
    Exit Function
ErrorHandler:
    ParseToMillisec = 0
End Function

' ステータスバーのメッセージをクリアするマクロ
Sub ClearStatusBar(Optional Dummy As Byte = 0)
    Application.StatusBar = False
End Sub
