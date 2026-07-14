VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "タイマー"
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

'© 2026 神垣貴誌
Option Explicit

Private isRunning As Boolean
Private isLooping As Boolean
Private totalElapsed As Long
Private startTime As Long
Private closeApproved As Boolean

Public Sub SetInitialTime(ByVal initialMs As Long, ByVal initialText As String)
    totalElapsed = initialMs
    lblTime.Caption = initialText
    
    If totalElapsed >= 600000 Then
        lblTime.ForeColor = vbRed
    Else
        lblTime.ForeColor = vbBlack
    End If
End Sub

Public Sub StartTimer()
    Dim nowTime As Long
    Dim currentSession As Long
    Dim displayTime As Long
    
    On Error GoTo ErrorHandler
    
    isRunning = True
    isLooping = True
    btnStartStop.Caption = "ストップ"
    startTime = timeGetTime()

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Do While isLooping
        If Not isRunning Then Exit Do
        
        nowTime = timeGetTime()
        currentSession = nowTime - startTime
        displayTime = totalElapsed + currentSession
        
        lblTime.Caption = FormatMillisec(displayTime)
        
        If displayTime >= 600000 Then
            lblTime.ForeColor = vbRed
        Else
            lblTime.ForeColor = vbBlack
        End If
        
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

Private Sub btnSaveEnd_Click()
    Dim wasRunning As Boolean
    Dim suspendStart As Long
    
    btnSaveEnd.Enabled = False
    wasRunning = isRunning
    suspendStart = timeGetTime()
    isRunning = False
    
    If MsgBox("現在の時間を記録してタイマーを終了しますか？", vbQuestion + vbYesNo, "確認") = vbYes Then
        On Error Resume Next
        TargetSheet.Range("B1").NumberFormatLocal = "@"
        TargetSheet.Range("B1").Value = lblTime.Caption
        On Error GoTo 0
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If closeApproved Then Exit Sub
    
    Dim wasRunning As Boolean
    Dim suspendStart As Long
    
    wasRunning = isRunning
    suspendStart = timeGetTime()
    isRunning = False
    
    If MsgBox("記録せずにタイマーを終了しますか？", vbQuestion + vbYesNo, "確認") = vbYes Then
        isLooping = False
    Else
        Cancel = True
        If wasRunning Then
            startTime = startTime + (timeGetTime() - suspendStart)
            isRunning = True
            If Not isLooping Then Call StartTimer
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    closeApproved = False
    totalElapsed = 0
    isRunning = False
    isLooping = False
    
    On Error Resume Next
    btnStartStop.TabIndex = 0
    btnReset.TabIndex = 1
    btnSaveEnd.TabIndex = 2
    
    lblTime.TabStop = False
    On Error GoTo 0
    
    Application.OnTime Now, "TriggerFormLoop"
End Sub
