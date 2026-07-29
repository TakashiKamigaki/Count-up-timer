Attribute VB_Name = "Module1"
Option Explicit

'システムのバージョン番号
Public Const PROGRAM_VERSION As String = "1.0"

Public TargetCellAddress As String
Public LimitMilliseconds As Long
Public UserDict As Object

' 手動起動用マクロ
Sub ShowTimerForm()
    Application.StatusBar = False
    Call Module1.LoadConfig(IsManual:=True)
    Call LaunchTimerProcess(IsManual:=True)
End Sub

' 実際の起動処理（ThisWorkbookからUserNameを受け取ります）
Public Sub LaunchTimerProcess(ByVal IsManual As Boolean, Optional ByVal Index As Integer)
    Dim frm As Object
    Dim timeValue As String
    Dim initMs As Long
    Dim initText As String
    
    For Each frm In UserForms
        If frm.Name = "UserForm1" Then Exit Sub
    Next frm
    
    If Not IsManual Then
        ThisWorkbook.Sheets(Index).Activate
    End If
    
    timeValue = Trim(ActiveSheet.Range(TargetCellAddress).Text)
    
    If Not IsManual And timeValue Like "##:##:##.###" Then
        Application.StatusBar = "自動起動をスキップしました。"
        Exit Sub
    End If
    
    If timeValue Like "##:##:##.###" Then
        initMs = ParseToMillisec(timeValue)
        initText = timeValue
    Else
        initMs = 0
        initText = "00:00:00.000"
    End If
    
    Load UserForm1
    Call UserForm1.SetInitialTime(initMs, initText, Index)
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

Public Sub ClearGlobalVariables(Optional Dummy As Byte = 0)
    Set UserDict = Nothing

    TargetCellAddress = ""
    LimitMilliseconds = 0
End Sub

'設定ファイルとユーザーリストを一括で読み込む共通関数
Public Sub LoadConfig(Optional ByVal IsManual As Boolean)
    Dim dataArray As Variant
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim isUserList As Boolean
    Dim key As String
    Dim testRange As Range
    
    TargetCellAddress = "S13"
    LimitMilliseconds = 600000
    
    On Error Resume Next
    Set ws = Worksheets("設定")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "設定シートが見つかりません。" & vbCrLf & "デフォルト値を使います。", vbCritical, "エラー"
        Application.StatusBar = "【タイマー " & Module1.PROGRAM_VERSION & "】設定シートが見つかりませんでした。"
    Else
        dataArray = ws.Range(ws.Cells(1.1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, 2)).Value
        If Not IsManual Then
            Set UserDict = CreateObject("Scripting.Dictionary")
        End If
    
        For i = LBound(dataArray, 1) + 1 To UBound(dataArray, 1)
            If dataArray(i, 1) <> "" Then
                If dataArray(i, 1) = "[ユーザーリスト]" Then
                    If IsManual Then Exit For
                    isUserList = True
                ElseIf isUserList Then
                    key = Replace(Replace(dataArray(i, 1), " ", ""), "　", "")
                    
                    If Not UserDict.exists(key) Then
                        UserDict.Add key, True
                    End If
                Else
                    Select Case dataArray(i, 1)
                    Case "タイムを書き込むセル"
                        On Error Resume Next
                        Set testRange = Range(dataArray(i, 2))
                        On Error GoTo 0
                        
                        If Not testRange Is Nothing Then
                            TargetCellAddress = dataArray(i, 2)
                        End If
                    Case "目標タイム(分)"
                        If IsNumeric(dataArray(i, 2)) Then
                            LimitMilliseconds = dataArray(i, 2) * 60 * 1000&
                        End If
                    End Select
                End If
            End If
        Next i
    End If
End Sub
