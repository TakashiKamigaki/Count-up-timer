Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

' 記録先シートを保持するグローバル変数
Public TargetSheet As Worksheet

Public Const TEAMS_NOTE_URL As String = "https://kamigaki.sharepoint.com/sites/manu/_layouts/15/Doc.aspx?sourcedoc={0584c5e2-1fe2-4f83-a09f-82a73683b124}&action=edit&wd=target%28%E5%AE%89%E5%85%A8%E6%B4%BB%E5%8B%95.one%7C6e79c070-1b84-4f84-8b10-0cd6de682ec2%2F%E9%AB%98%E7%B2%BE%E5%BA%A6%E3%83%9F%E3%83%AA%E7%A7%92%E3%82%BF%E3%82%A4%E3%83%9E%E3%83%BC%E3%82%B7%E3%82%B9%E3%83%86%E3%83%A0%20%E6%93%8D%E4%BD%9C%E3%83%9E%E3%83%8B%E3%83%A5%E3%82%A2%E3%83%AB%7C28daf8a1-c7d8-4e27-aed7-ef41a04fd9ed%2F%29&wdorigin=NavigationUrl"

' ★ 手動起動用マクロ
Sub ShowTimerForm()
    ' 手動起動なので IsManual:=True、UserNameは渡さない
    Call LaunchTimerProcess(IsManual:=True)
End Sub

' 実際の起動処理（ThisWorkbookからUserNameを受け取ります）
Public Sub LaunchTimerProcess(ByVal IsManual As Boolean, Optional ByVal UserName As String = "")
    Dim frm As Object
    Dim b1Value As String
    Dim initMs As Long
    Dim initText As String
    Dim ws As Worksheet
    Dim foundSheet As Worksheet
    
    ' 既に開いている場合は、二重起動しないよう無視する
    For Each frm In UserForms
        If frm.Name = "UserForm1" Then Exit Sub
    Next frm
    
    ' ★【仕様変更】手動起動か、自動起動かでシートの決定ロジックを分岐
    If IsManual Then
        ' 手動起動時は、現在の「アクティブシート」に記録を固定
        Set TargetSheet = ActiveSheet
    Else
        ' 自動起動時は、渡されたユーザー名が含まれるシートを検索
        For Each ws In ThisWorkbook.Worksheets
            If InStr(1, ws.Name, UserName, vbTextCompare) > 0 Then
                Set foundSheet = ws
                Exit For
            End If
        Next ws
        
        ' 該当シートが見つかればアクティブにし、無ければアクティブシートにする
        If Not foundSheet Is Nothing Then
            Set TargetSheet = foundSheet
            On Error Resume Next
            TargetSheet.Activate
            On Error GoTo 0
        Else
            Set TargetSheet = ActiveSheet
        End If
    End If
    
    ' 特定したTargetSheetのS13セルをチェック
    b1Value = Trim(TargetSheet.Range("S13").Text)
    
    ' ★【仕様変更】自動起動時、B1セルにすでにタイムが入力されていたら起動しない
    If Not IsManual And b1Value Like "##:##:##.###" Then
        Exit Sub
    End If
    
    ' 初期値の設定（B1に正しい形式があればそれを引き継ぎ、無ければ0リセット）
    If b1Value Like "##:##:##.###" Then
        initMs = ParseToMillisec(b1Value)
        initText = b1Value
    Else
        initMs = 0
        initText = "00:00:00.000"
    End If
    
    ' ユーザーフォームを表示
    Load UserForm1
    Call UserForm1.SetInitialTime(initMs, initText)
    UserForm1.Show vbModeless
End Sub

''' <summary>
''' TeamsチャネルのNoteを既定のブラウザで開く
''' </summary>
Sub OpenTeamsNote(Optional Dummy As Byte = 0)
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ' 既定のブラウザでURLを起動
    wsh.Run TEAMS_NOTE_URL
    Set wsh = Nothing
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
