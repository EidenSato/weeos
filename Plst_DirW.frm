VERSION 5.00
Begin VB.Form Plst_DirW 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "部品表書き込みﾌｧｲﾙ指定"
   ClientHeight    =   5835
   ClientLeft      =   810
   ClientTop       =   2115
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
   Icon            =   "Plst_DirW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5835
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optAUX 
      BackColor       =   &H00004000&
      Caption         =   "共有ﾌｫﾙﾀﾞ-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton optStandard 
      BackColor       =   &H00004000&
      Caption         =   "個別ﾌｫﾙﾀﾞ-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSetPLT 
      Caption         =   "[ .PLT]にする"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "実行(&G)"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblnamae 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00004000&
      Caption         =   "新しい ﾌｧｲﾙ名 (&N) ："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
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
Attribute VB_Name = "Plst_DirW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'***********************
'*** ファイル選択(WR) ***
'***********************
'
Option Explicit
'
    Dim flag_Busy As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (6715 - 960) * HyoujiBairitu + 480
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
    Me.Caption = STATUS & "(WR)"
'
    flag_Busy = True
    optStandard.Value = True
    optAUX.Value = False
'
    SET_FLbox       '*** ファイルリストボックスの設定 ***
    FLGesc = 0      '*** エスケープフラグクリアー ***
    flag_Busy = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1      '*** エスケープフラグセット ***
    End If
End Sub

Private Sub SET_FLbox()
    Dim i As Integer, j As Integer
'
    lblnamae.Caption = "新しい ﾌｧｲﾙ名 (&N) ："
    txt_file.Text = PFLnameT
'
On Error GoTo Wr_inhibit
'
kaisi:
    If optStandard.Value = True Then
        lbl_dir.Caption = TMPdir1 & "\"
        File1.Path = TMPdir1 & "\"
    Else
        lbl_dir.Caption = TMPdir2 & "\"
        File1.Path = TMPdir2 & "\"
    End If
'
    File1.Pattern = "B*.*; B*.PLT"
Exit Sub
'
Wr_inhibit:
    Beep
    i = MsgBox("部品表データを収容するフォルダーが見つかりません！", vbOKCancel, STATUS)
    If i = vbOK Then
        FLG_job_error_end = 0   '*** フラグりセット ***
        Resume kaisi
    End If
'
    Resume Utikiri
Utikiri:
    FLG_job_error_end = 1   '*** フラグセット ***
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer
'
    If txt_file.Text = "" Then
        Beep
        i = MsgBox("ﾌｧｲﾙ名が指定されていません。", vbCritical, STATUS)
'
    Else
        Me.MousePointer = vbHourglass
'
        PFLnameT = txt_file.Text
        DRVpartlistT = lbl_dir.Caption & PFLnameT
'
        Call WRpartlist(DRVpartlistT, PlistnameT, PlistdateT, RemarksT, PLSTT(), PtotalT, PdimT)  '*** 部品表セーブ ***
        FLGchange = 0
'
'        Call WRplstWork      '*** 部品表作成データ書き込み ***
'
        FLGesc = 0
        Me.MousePointer = vbDefault
        Unload Me
    End If
End Sub

Private Sub cmdQuit_Click()
    FLGesc = 1      '*** エスケープフラグセット ***
    Unload Me
End Sub

Private Sub cmdSetPLT_Click()
    Dim temp As String
    Dim ichi As Integer
'
    temp = txt_file.Text
    ichi = InStr(1, temp, ".")
    If ichi = 0 Then
        temp = temp & ".PLT"
    Else
        If Right(UCase(temp), 4) <> ".PLT" Then
            temp = Left(temp, ichi - 1) & Mid(temp, ichi + 1) & ".PLT"
        Else
            Beep
        End If
    End If
'
    txt_file.Text = temp
End Sub

Private Sub File1_Click()
    txt_file.Text = File1.FileName
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
    lbl_dir.FontSize = 10 * HyoujiBairitu
    lbl_dir.Height = 495 * HyoujiBairitu
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
    cmdSetPLT.Top = 480 + (2760 - 480) * HyoujiBairitu
    cmdSetPLT.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmdSetPLT.Width = 1215 * HyoujiBairitu
    cmdSetPLT.Height = 495 * HyoujiBairitu
    cmdSetPLT.FontSize = 9 * HyoujiBairitu
'
    cmdQuit.Top = 480 + (4560 - 480) * HyoujiBairitu
    cmdQuit.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmdQuit.Width = 1215 * HyoujiBairitu
    cmdQuit.Height = 615 * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
End Sub

