VERSION 5.00
Begin VB.Form Const_DirW 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "EＥOS 電気 構成表 (WR)"
   ClientHeight    =   5790
   ClientLeft      =   1710
   ClientTop       =   1875
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
   Icon            =   "Const_DirW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5790
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmd_Quit 
      Caption         =   "中止(&Q)"
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_std 
      Caption         =   "標準(&S)"
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "決定(&G)"
      Default         =   -1  'True
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
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
      Text            =   "Text1"
      Top             =   1920
      Width           =   2415
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
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lbl_folder 
      BackColor       =   &H00004000&
      Caption         =   "ﾌｧｲﾙの場所　："
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
      Height          =   225
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1455
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
   Begin VB.Label lblnamae 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00004000&
      Caption         =   "書き込み ﾌｧｲﾙ名 (&N)  ："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "Const_DirW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'***************************
'*** 新ファイル名でセーブ ***
'***************************
'
Option Explicit
'
    Private FLGback As Integer
'
Private Sub Form_Load()
    Call GamenSettei
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
    Me.Caption = STATUS & "(WR)"
'
    Call SET_FLbox       '*** ファイルリストボックスの設定 ***
    FLGesc = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1
    End If
End Sub

Private Sub cmd_Quit_Click()
    FLGesc = 1
    Unload Me
End Sub

Private Sub cmd_OK_Click()
    Me.Caption = "  構成表を保存しています。ちょっと待ってね！"
    Me.MousePointer = vbHourglass
    DoEvents
'
    DRVconstT = TMPdir1 & "\" & txt_file.Text
    CURR_file = txt_file.Text
'
    RevdateT = Format(Date, "yy/mm/dd")    '*** 日付更新 ***
'
    Call WRconst_lst(DRVconstT, CATnoT, CATnameT, ZubanT, PersonT, OrgdateT, RevdateT _
            , CheckdateT, OutdateT, KLSTT(), KtotalT, KdimT, KoubanT, DaisuuT, KbikouT _
                , KyobiAT, KyobiBT)                   '*** 構成表セーブ ***
    DoEvents
'
    Me.MousePointer = vbDefault
    Me.Caption = STATUS
    Unload Me
End Sub

Private Sub cmd_std_Click()
    txt_file.Text = "constlst.cod"
End Sub

Private Sub File1_Click()
    txt_file.Text = File1.FileName
End Sub

Private Sub Timer1_Timer()
    If FLGback = 1 Then
        FLGesc = 1
        Unload Me
    End If
End Sub

Private Sub SET_FLbox()
    Dim i As Integer
'
    FLGback = 0     '*** エラーで戻りフラグリセット ***
'
kaisi:
On Error GoTo NoFile
'
    txt_file.Text = CURR_file
    lbl_dir.Caption = TMPdir1 & "\"
    File1.Path = TMPdir1 & "\"
    File1.Pattern = "*.cod"
Exit Sub
'
NoFile: Beep
    i = MsgBox("構成表を収容するフォルダーを準備しましたか？", vbQuestion Or vbRetryCancel, STATUS)
    If i = vbRetry Then
        Resume kaisi
'
    Else    '*** キャンセル ***
        FLGback = 1             '*** エラーでバックフラグセット、MsgBoxの動作が終了できないので。 ***
        Timer1.Enabled = True   '*** タイマー動作開始 → タイマーＵＰで Unload ***
        Resume Next
    End If
End Sub

Private Sub GamenSettei()
    Me.Width = 480 + (6705 - 960) * HyoujiBairitu + 480
    Me.Height = 480 + (6270 - 960) * HyoujiBairitu + 480
'
    lbl_folder.Top = 480
    lbl_folder.Left = 480
    lbl_folder.Width = 1335 * HyoujiBairitu
    lbl_folder.Height = 255 * HyoujiBairitu
    lbl_folder.FontSize = 10 * HyoujiBairitu
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
    cmd_OK.Top = 480 + (1800 - 480) * HyoujiBairitu
    cmd_OK.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmd_OK.Width = 1215 * HyoujiBairitu
    cmd_OK.Height = 735 * HyoujiBairitu
    cmd_OK.FontSize = 10 * HyoujiBairitu
'
    cmd_std.Top = 480 + (2760 - 480) * HyoujiBairitu
    cmd_std.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmd_std.Width = 1215 * HyoujiBairitu
    cmd_std.Height = 615 * HyoujiBairitu
    cmd_std.FontSize = 10 * HyoujiBairitu
'
    cmd_Quit.Top = 480 + (4560 - 480) * HyoujiBairitu
    cmd_Quit.Left = 480 + (4320 - 480) * HyoujiBairitu
    cmd_Quit.Width = 1215 * HyoujiBairitu
    cmd_Quit.Height = 615 * HyoujiBairitu
    cmd_Quit.FontSize = 10 * HyoujiBairitu
End Sub

