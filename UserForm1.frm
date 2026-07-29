VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "タイマー (v1.0)"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private isRunning As Boolean   ' タイマーが作動中かどうか
Private isLooping As Boolean   ' ループ処理自体が動いているかどうか
Private totalElapsed As Long   ' これまでに蓄積された経過時間（ミリ秒）
Private startTime As Long      ' スタートした瞬間の timeGetTime の値
Private closeApproved As Boolean ' 閉じることが許可されたかどうかのフラグ
Private savedSheetIndex As Integer ' インデックス番号を保持する変数
Private ManualURL As Worksheet

' 外部から初期値（ミリ秒）と表示用文字列を受け取るメソッド
Public Sub SetInitialTime(ByVal initialMs As Long, ByVal initialText As String, ByRef sheetIndex As Integer)
    totalElapsed = initialMs
    lblTime.Caption = initialText
    savedSheetIndex = sheetIndex
    
    If totalElapsed >= Module1.LimitMilliseconds Then
        lblTime.ForeColor = vbRed
    Else
        lblTime.ForeColor = vbBlack
    End If
End Sub

' タイマーのメインループ
Public Sub StartTimer()
    Dim nowTime As Long
    Dim currentSession As Long
    Dim displayTime As Long
    Dim lastUpdate As Long
    
    On Error GoTo ErrorHandler
    
    isRunning = True
    isLooping = True
    btnStartStop.Caption = "ストップ"
    startTime = timeGetTime()
    lastUpdate = timeGetTime()

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Do While isLooping
        If Not isRunning Then Exit Do
        
        nowTime = timeGetTime()
        currentSession = nowTime - startTime
        displayTime = totalElapsed + currentSession
        
        If timeGetTime() - lastUpdate >= 20 Then
            lblTime.Caption = FormatMillisec(displayTime)
            
            If displayTime >= Module1.LimitMilliseconds Then
                lblTime.ForeColor = vbRed
            Else
                lblTime.ForeColor = vbBlack
            End If
            lastUpdate = timeGetTime()
        End If
        
        Sleep 1
        DoEvents
    Loop

CleanUp:
    isLooping = False
    btnStartStop.Enabled = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "タイマー処理中に予期せぬエラーが発生しました。", vbCritical
    Resume CleanUp
End Sub

'
Private Function FormatMillisec(ByVal ms As Long) As String
    Dim h As Long, m As Long, s As Long
    h = ms \ 3600000
    ms = ms Mod 3600000
    m = ms \ 60000
    ms = ms Mod 60000
    s = ms \ 1000
    ms = ms Mod 1000
    FormatMillisec = Format(h, "00") & ":" & Format(m, "00") & ":" & Format(s, "00") & "." & Format(ms, "000")
End Function

' スタート / ストップ ボタン
Private Sub btnStartStop_Click()
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub
    isProcessing = True
    
    If Not isRunning Then
        startTime = timeGetTime()
        isRunning = True
        btnStartStop.Caption = "ストップ"
        
        btnStartStop.SetFocus
        
        isProcessing = False

        If Not isLooping Then Call StartTimer
    Else
        totalElapsed = totalElapsed + (timeGetTime() - startTime)
        isRunning = False
        btnStartStop.Caption = "スタート"
        btnStartStop.SetFocus
        isProcessing = False
    End If
End Sub

' リセットボタン
Private Sub btnReset_Click()
    isRunning = False
    isLooping = False
    totalElapsed = 0
    lblTime.Caption = "00:00:00.000"
    lblTime.ForeColor = vbBlack
    btnStartStop.Caption = "スタート"
End Sub

Private Sub lblManualLink_Click()
    If isRunning Then
        btnStartStop_Click
    End If
    
    ManualURL.Activate
End Sub

' 記録して終了ボタン
Private Sub btnSaveEnd_Click()
    Dim wasRunning As Boolean
    Dim suspendStart As Long
    Dim targetWS As Worksheet
    
    btnSaveEnd.Enabled = False
    wasRunning = isRunning
    suspendStart = timeGetTime()
    isRunning = False
    
    If MsgBox("現在の時間を記録してタイマーを終了しますか？", vbQuestion + vbYesNo, "確認") = vbYes Then
        If savedSheetIndex <= 0 Then
            Set targetWS = ActiveSheet
        Else
            Set targetWS = ThisWorkbook.Sheets(savedSheetIndex)
        End If
        
        targetWS.Range(Module1.TargetCellAddress).NumberFormatLocal = "@"
        targetWS.Range(Module1.TargetCellAddress).Value = lblTime.Caption
        
        Application.StatusBar = "【タイマー " & Module1.PROGRAM_VERSION & "】タイムを記録しました。(" & Format(Now, "HH:nn:ss") & ")"
        Application.OnTime Now + timeValue("00:00:03"), "ClearStatusBar"
        
        closeApproved = True
        Unload Me
    Else
        If wasRunning Then
            startTime = startTime + (timeGetTime() - suspendStart)
            isRunning = True
            If Not isLooping Then Call StartTimer
        End If
        btnSaveEnd.Enabled = True
    End If
End Sub

' フォームが閉じられる直前の処理（右上の「×」ボタン対策）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If closeApproved Then
        Call ClearGlobalVariables
        Exit Sub
    End If
    
    Dim wasRunning As Boolean
    Dim suspendStart As Long
    
    wasRunning = isRunning
    suspendStart = timeGetTime()
    isRunning = False
    
    If MsgBox("記録せずにタイマーを終了しますか？", vbQuestion + vbYesNo, "確認") = vbYes Then
        isLooping = False
        Call ClearGlobalVariables
    Else
        Cancel = True
        If wasRunning Then
            startTime = startTime + (timeGetTime() - suspendStart)
            isRunning = True
            If Not isLooping Then Call StartTimer
        End If
    End If
End Sub

' フォーム読み込み時
Private Sub UserForm_Initialize()
    closeApproved = False
    totalElapsed = 0
    isRunning = False
    isLooping = False
    Me.Caption = "タイマー (v" & Module1.PROGRAM_VERSION & ")"
    
    On Error Resume Next
    Set ManualURL = ThisWorkbook.Worksheets("マニュアル")
    
    btnStartStop.TabIndex = 0
    btnReset.TabIndex = 1
    btnSaveEnd.TabIndex = 2
    
    lblTime.TabStop = False
    On Error GoTo 0
    If ManualURL Is Nothing Then
        lblManualLink.Visible = False
    Else
        lblManualLink.Visible = True
    End If
    
    Application.OnTime Now, "TriggerFormLoop"
End Sub
