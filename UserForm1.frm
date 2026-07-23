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
Option Explicit

Private isRunning As Boolean   ' タイマーが作動中かどうか
Private isLooping As Boolean   ' ループ処理自体が動いているかどうか
Private totalElapsed As Long   ' これまでに蓄積された経過時間（ミリ秒）
Private startTime As Long      ' スタートした瞬間の timeGetTime の値
Private closeApproved As Boolean ' 閉じることが許可されたかどうかのフラグ

' 外部から初期値（ミリ秒）と表示用文字列を受け取るメソッド
Public Sub SetInitialTime(ByVal initialMs As Long, ByVal initialText As String)
    totalElapsed = initialMs
    lblTime.Caption = initialText
    
    ' 10分以上なら文字色を赤にする
    If totalElapsed >= 600000 Then
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
    
    On Error GoTo ErrorHandler
    
    isRunning = True
    isLooping = True
    btnStartStop.Caption = "ストップ"
    startTime = timeGetTime()

    ' ★【修正】画面更新とイベントを常に有効（True）にします。
    ' これにより、タイマー作動中もシートのセルを自由にクリック・編集できるようになります。
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
        
        ' ★Excelにシート操作の隙を与えるための処理
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

' ★ここに新しく（正しく）配置されている必要があります
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
    ' 多重起動を防ぐため、一瞬だけガードをかける
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub
    isProcessing = True
    
    If Not isRunning Then
        ' --- 再開・スタート処理 ---
        startTime = timeGetTime()
        isRunning = True
        btnStartStop.Caption = "ストップ"
        
        ' ボタンにフォーカスを当てておく（Spaceキーで止めるため）
        btnStartStop.SetFocus
        
        isProcessing = False
        ' ループ処理へ突入
        If Not isLooping Then Call StartTimer
    Else
        ' --- 一時停止処理 ---
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
    ' ユーザーがマニュアルを読んでいる間に時間が進まないよう、安全のため一時停止させる
    If isRunning Then
        btnStartStop_Click
    End If
    
    ' 標準モジュールのTeams Note起動マクロを呼び出し
    Call OpenTeamsNote
End Sub

' 記録して終了ボタン
Private Sub btnSaveEnd_Click()
    Dim wasRunning As Boolean
    Dim suspendStart As Long
    
    btnSaveEnd.Enabled = False
    wasRunning = isRunning
    suspendStart = timeGetTime() ' ダイアログを開いた時間を一時記録
    isRunning = False            ' カウントを一時保留
    
    If MsgBox("現在の時間を記録してタイマーを終了しますか？", vbQuestion + vbYesNo, "確認") = vbYes Then
        ' 固定されたTargetSheetに対して安全に書き込み
        On Error Resume Next
        TargetSheet.Range("S13").NumberFormatLocal = "@"
        TargetSheet.Range("S13").Value = lblTime.Caption
        On Error GoTo 0
        closeApproved = True
        Unload Me
    Else
        ' キャンセルされた場合
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

' フォーム読み込み時
Private Sub UserForm_Initialize()
    closeApproved = False
    totalElapsed = 0
    isRunning = False
    isLooping = False
    
    ' ★【追加】タブキーを押したときの移動順序を厳密に固定する
    ' (0番から順番にフォーカスが移動します)
    On Error Resume Next
    btnStartStop.TabIndex = 0
    btnReset.TabIndex = 1
    btnSaveEnd.TabIndex = 2
    
    ' ラベルなどの他のコントロールは Tab キーで選択されないように除外する
    lblTime.TabStop = False
    On Error GoTo 0
    
    ' 画面描画完了後に自動起動させる予約
    Application.OnTime Now, "TriggerFormLoop"
End Sub
