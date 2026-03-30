VERSION 5.00
Begin VB.Form Plst_DirR 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "部品表読み込みﾌｧｲﾙ指定"
   ClientHeight    =   5790
   ClientLeft      =   1935
   ClientTop       =   2130
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Plst_DirR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5790
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optAUX 
      BackColor       =   &H00004000&
      Caption         =   "共有ﾌｫﾙﾀﾞ-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton optStandard 
      BackColor       =   &H00004000&
      Caption         =   "個別ﾌｫﾙﾀﾞ-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdShinki 
      Caption         =   "新規(&N)"
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全部(&A)"
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "決定(&G)"
      Height          =   735
      Left            =   4320
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txt_file 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "中止(&Q)"
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lbl_folder 
      BackColor       =   &H00004000&
      Caption         =   "ﾌｧｲﾙの場所 ："
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblnamae 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00004000&
      Caption         =   "読み込み ﾌｧｲﾙ名 (&N) ："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lbl_dir 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ファイル名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   5655
   End
End
Attribute VB_Name = "Plst_DirR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*******************
'*** ファイル選択 ***
'*******************
'
Option Explicit
'
    Dim FLGback As Integer
    Dim flag_Busy As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (6705 - 960) * HyoujiBairitu + 480
    Height = 480 + (6270 - 960) * HyoujiBairitu + 480
'
    If Eeos2_mainMDI.ScaleHeight > Height Then
        Top = Eeos2_mainMDI.Top + 480 + (Eeos2_mainMDI.ScaleHeight - Height) \ 2
    Else
        Top = (Screen.Height - Height) \ 2
    End If
'
    If Eeos2_mainMDI.ScaleWidth > Width Then
        Left = Eeos2_mainMDI.Left + (Eeos2_mainMDI.ScaleWidth - Width) \ 2
    Else
        Left = (Screen.Width - Width) \ 2
    End If
'
    Call GamenSettei    '*** 画面部品諸元指定 ***
'
    Me.Caption = STATUS & " (RD)"
'
    flag_Busy = True
'
    If Xcont0(17) = "0" Then    '*** 共有ﾌｫﾙﾀﾞ ***
        optAUX.Value = True
    Else    '= "1"              '*** 個別ﾌｫﾙﾀﾞ ***
        optStandard.Value = True
    End If
'
    Call SET_FLbox      '*** ファイルリストボックスの設定 ***
    Call SET_cmdAll     '*** 全部コマンドボタン設定 ***
    Call SET_shinki     '*** 新規コマンドボタン設定 ***
    FLGesc = 0
    flag_Busy = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1
    End If
End Sub

Private Sub SET_cmdAll()
    Select Case FLGlevel
    Case 1 To 4
        cmdAll.Enabled = False
    Case 5 To 8
        cmdAll.Enabled = True
    End Select
'
    FLGall = 0
End Sub

Private Sub SET_FLbox()
    Dim i As Integer
'
    FLGback = 0     '*** エラーでバックフラグリセット ***
kaisi:
On Error GoTo NoFile
'
    Select Case FLGlevel
    Case 2      '*** OrCADデータの変換 ***
        lbl_dir.Caption = TMPdir3 & "\"
        File1.Path = TMPdir3 & "\"
        File1.Pattern = "X*.*"
'
        flag_Busy = True
        optStandard.Value = True
        optAUX.Enabled = False
'        optAUX.Visible = False
        flag_Busy = False
'
    Case 4      '*** 部品表作成データ編集 ***
        lbl_dir.Caption = TMPdir1 & "\"
        File1.Path = TMPdir1 & "\"
        File1.Pattern = "*.dat"
        txt_file.Text = "PLSTWORK.DAT"
'
        flag_Busy = True
        optStandard.Value = True
        optAUX.Enabled = False
'        optAUX.Visible = False
        flag_Busy = False
'
    Case Else
        If optStandard.Value = True Then
            lbl_dir.Caption = TMPdir1 & "\"
            File1.Path = TMPdir1 & "\"
        Else
            lbl_dir.Caption = TMPdir2 & "\"
            File1.Path = TMPdir2 & "\"
        End If
'
        File1.Pattern = "B*.*; B*.PLT"
    End Select
Exit Sub
'
NoFile: Beep
    i = MsgBox("部品表が収容されているフォルダーを準備しましたか？", vbQuestion Or vbRetryCancel, STATUS)
    If i = vbRetry Then
        Resume kaisi
    Else    '*** キャンセル ***
        FLGback = 1             '*** エラーでバックフラグセット、MsgBoxの動作が終了できないので。 ***
        Timer1.Enabled = True   '*** タイマー動作開始 → タイマーＵＰで Unload ***
        Resume Next
    End If
End Sub

Private Sub SET_shinki()
    Select Case FLGjob
    Case 1  '*** 構成表 ***
            cmdShinki.Enabled = True
    Case 2  '*** 部品表 ***
        Select Case FLGlevel
        Case 1, 3
            cmdShinki.Enabled = True
        Case 2, 4 To 8
            cmdShinki.Enabled = False
        End Select
    End Select
End Sub

Private Sub cmdAll_Click()
    Dim i As Integer
'
    Beep
    i = MsgBox("構成表に従ってすべてを処理します。", vbExclamation Or vbOKCancel, STATUS2)
    If i = vbOK Then
        FLGall = 1
        If optStandard.Value = True Then
            TMPplst = TMPdir1
        Else
            TMPplst = TMPdir2
        End If
'
        Select Case FLGlevel
        Case 5  '*** 標準部品表 ***
            Unload Me
'
        Case 6  '*** 一覧表 ***
            Unload Me
'
        Case 7  '*** 数量表 ***
            Unload Me
'
        Case 8  '*** ラベル印刷 ***
'            Plst_PRNlbl.Show
            Unload Me
'
        End Select
'
    Else
        FLGall = 0
    End If
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer
    Dim Tmp1 As String
'
    If txt_file.Text = "" Then
        Beep
        i = MsgBox("ﾌｧｲﾙ名が指定されていません。", vbCritical, STATUS)
'
    Else
        DRVplstWork = TMPdir1 & "\PLSTWORK.DAT"
'
        If optStandard.Value = True Then
            TMPplst = TMPdir1
        Else
            TMPplst = TMPdir2
        End If
'
        Select Case FLGlevel
        Case 1      '*** 部品表更新・確認 ***
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
            Tmp1 = Dir(DRVpartlistT)
            If Tmp1 = "" Then
                i = MsgBox("ファイル " & DRVpartlistT & " が見つかりません｡ ", vbCritical, STATUS)
                Exit Sub
'
            End If
'
            Me.MousePointer = vbHourglass
            Unload Me
'
        Case 2      '*** OrCADからの変換 ***
            DRVcadplst = TMPdir3 & "\" & txt_file.Text
            PFLnameT = "B" & Mid(txt_file.Text, 2)
            DRVpartlistT = TMPplst & "\" & PFLnameT
            Tmp1 = Dir(DRVcadplst)
            If Tmp1 = "" Then
                i = MsgBox("ファイル " & DRVcadplst & " が見つかりません｡ ", vbCritical, STATUS)
                Exit Sub
'
            End If
'
            Unload Me
'
        Case 3      '*** 部品表新規作成 ***
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
            Unload Me
'
        Case 4      '*** 部品表作成データ編集 ***
            FLGesc = 0
'            Parts_work.Show
            Unload Me
'
        Case 5      '*** 標準部品表印刷 ***
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
            FLGall = 0
'
            Tmp1 = Dir(DRVpartlistT)
            If Tmp1 = "" Then
                i = MsgBox("ファイル " & DRVpartlistT & " が見つかりません｡ ", vbCritical, STATUS)
                Exit Sub
'
            End If
            Unload Me
'
        Case 6      '*** 部品一覧表印刷 ***
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
            FLGall = 0
'
            Tmp1 = Dir(DRVpartlistT)
            If Tmp1 = "" Then
                i = MsgBox("ファイル " & DRVpartlistT & " が見つかりません｡ ", vbCritical, STATUS)
                Exit Sub
'
            End If
            Unload Me
'
        Case 7      '*** 部品数量表印刷 ***
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
            FLGall = 0
'
            Tmp1 = Dir(DRVpartlistT)
            If Tmp1 = "" Then
                i = MsgBox("ファイル " & DRVpartlistT & " が見つかりません｡ ", vbCritical, STATUS)
                Exit Sub
'
            End If
            Unload Me
'
        Case 8      '*** 部品ラベル印刷 ***
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
            FLGall = 0
'
            Tmp1 = Dir(DRVpartlistT)
            If Tmp1 = "" Then
                i = MsgBox("ファイル " & DRVpartlistT & " が見つかりません｡ ", vbCritical, STATUS)
                Exit Sub
'
            End If
'           Plst_PRNlbl.Show
            Unload Me
'
        Case Else
            PFLnameT = txt_file.Text
            DRVpartlistT = lbl_dir.Caption & PFLnameT
        End Select
    End If
End Sub

Private Sub cmdQuit_Click()
    FLGesc = 1
    Unload Me
End Sub

Private Sub cmdShinki_Click()
    FLGlevel = 3        '*** 新規作成 ***
    txt_file.Text = "BSHINKI.PLT"
End Sub

Private Sub File1_Click()
    txt_file.Text = File1.FileName
End Sub

Private Sub lbl_dir_Change()
'    Debug.Print lbl_dir.Caption
End Sub

Private Sub optAUX_Click()
    If flag_Busy = True Then Exit Sub
'
    optStandard.Value = False
    optAUX.Value = True
'
    lbl_dir.Caption = TMPdir2 & "\"
    File1.Path = TMPdir2 & "\"
End Sub

Private Sub optstandard_Click()
    If flag_Busy = True Then Exit Sub
'
    optStandard.Value = True
    optAUX.Value = False
'
    lbl_dir.Caption = TMPdir1 & "\"
    File1.Path = TMPdir1 & "\"
End Sub

Private Sub Timer1_Timer()
    If FLGback = 1 Then
        FLGesc = 1
        Unload Me
    End If
End Sub

Private Sub GamenSettei()
    lbl_folder.Top = 480
    lbl_folder.Left = 480
    lbl_folder.Width = 1335 * HyoujiBairitu
    lbl_folder.Height = 255 * HyoujiBairitu
    lbl_folder.FontSize = 10 * HyoujiBairitu
'
    optStandard.Top = 480
    optStandard.Left = 480 + (3120 - 480) * HyoujiBairitu
    optStandard.Width = 1455 * HyoujiBairitu
    optStandard.Height = 255 * HyoujiBairitu
    optStandard.FontSize = 10 * HyoujiBairitu
'
    optAUX.Top = 480
    optAUX.Left = 480 + (4680 - 480) * HyoujiBairitu
    optAUX.Width = 1455 * HyoujiBairitu
    optAUX.Height = 255 * HyoujiBairitu
    optAUX.FontSize = 10 * HyoujiBairitu
'
    lbl_dir.Top = 480 + (840 - 480) * HyoujiBairitu
    lbl_dir.Left = 480
    lbl_dir.Width = 5655 * HyoujiBairitu
    lbl_dir.Height = 495 * HyoujiBairitu
    lbl_dir.FontSize = 10 * HyoujiBairitu
'
    lblnamae.Top = 480 + (1560 - 480) * HyoujiBairitu
    lblnamae.Left = 480
    lblnamae.Width = 2175 * HyoujiBairitu
    lblnamae.Height = 255 * HyoujiBairitu
    lblnamae.FontSize = 10 * HyoujiBairitu
'
    txt_file.Top = 480 + (1920 - 480) * HyoujiBairitu
    txt_file.Left = 480 + (1080 - 480) * HyoujiBairitu
    txt_file.Width = 2415 * HyoujiBairitu
    txt_file.Height = 285 * HyoujiBairitu
    txt_file.FontSize = 10 * HyoujiBairitu
'
    File1.Top = 480 + (2280 - 480) * HyoujiBairitu
    File1.Left = 480 + (1080 - 480) * HyoujiBairitu
    File1.Width = 2655 * HyoujiBairitu
    File1.FontSize = 10 * HyoujiBairitu
    File1.Height = 3015 * HyoujiBairitu
'
    cmdGo.Top = 480 + (1800 - 480) * HyoujiBairitu
    cmdGo.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmdGo.Width = 1215 * HyoujiBairitu
    cmdGo.Height = 735 * HyoujiBairitu
    cmdGo.FontSize = 10 * HyoujiBairitu
'
    cmdAll.Top = 480 + (2760 - 480) * HyoujiBairitu
    cmdAll.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmdAll.Width = 1215 * HyoujiBairitu
    cmdAll.Height = 615 * HyoujiBairitu
    cmdAll.FontSize = 10 * HyoujiBairitu
'
    cmdShinki.Top = 480 + (3600 - 480) * HyoujiBairitu
    cmdShinki.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmdShinki.Width = 1215 * HyoujiBairitu
    cmdShinki.Height = 615 * HyoujiBairitu
    cmdShinki.FontSize = 10 * HyoujiBairitu
'
    cmdQuit.Top = 480 + (4560 - 480) * HyoujiBairitu
    cmdQuit.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmdQuit.Width = 1215 * HyoujiBairitu
    cmdQuit.Height = 615 * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
End Sub

