VERSION 5.00
Begin VB.Form Const_DirR 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "構成表読み込みﾌｧｲﾙ指定"
   ClientHeight    =   4335
   ClientLeft      =   945
   ClientTop       =   2325
   ClientWidth     =   7095
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
   Icon            =   "Const_DirR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4335
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   2400
   End
   Begin VB.CommandButton cmd_New 
      Caption         =   "新規(&N)"
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Quit 
      Caption         =   "中止(&Q)"
      Height          =   615
      Left            =   5400
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Std 
      Caption         =   "標準(&S)"
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "決定(&G)"
      Default         =   -1  'True
      Height          =   735
      Left            =   5400
      TabIndex        =   6
      Top             =   480
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
      Left            =   2880
      TabIndex        =   5
      Top             =   480
      Width           =   1935
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
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblTadasigaki 
      Alignment       =   1  '右揃え
      BackColor       =   &H00004000&
      Caption         =   "（環境設定による）"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
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
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lbl_dir 
      Alignment       =   1  '右揃え
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
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblnamae 
      Alignment       =   1  '右揃え
      BackColor       =   &H00004000&
      Caption         =   "読み込み ﾌｧｲﾙ名 (&N) ："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Const_DirR"
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

Private Sub cmd_new_Click()
    txt_file.Text = "shinki01.cod"
    FLGshinki = 1       '*** 新規フラグセット ***
'
'   Call cmd_OK_Click
End Sub

Private Sub cmd_OK_Click()
    Dim i As Integer
    Dim Tmp1 As String
'
    If txt_file.Text = "" Then
        Beep
        i = MsgBox("ﾌｧｲﾙ名が指定されていません。", vbCritical, STATUS)
    Else
        DRVconst = TMPdir1 & "\" & txt_file.Text
        CURR_file = txt_file.Text
        Tmp1 = Dir(DRVconst)
        If FLGshinki = 0 And Tmp1 = "" Then
            i = MsgBox("ファイル " & DRVconst & " が見つかりません｡ ", vbCritical, STATUS)
            Exit Sub
'
        End If
'
        If FLGshinki = 0 Then
            Me.MousePointer = vbHourglass
            DoEvents
            Call RDconst_lst(DRVconst, CATno, CATname, Zuban, Person, Orgdate, Revdate, Checkdate, Outdate, _
                KLST(), Ktotal, Kdim, Kouban, Daisuu, Kbikou, Kyobi1, Kyobi2)
            DoEvents
            Me.MousePointer = vbDefault
            DoEvents
        End If
'
        Unload Me
    End If
End Sub

Private Sub cmd_Quit_Click()
    FLGesc = 1
    Unload Me
End Sub

Private Sub cmd_std_Click()
    txt_file.Text = "constlst.cod"
    FLGshinki = 0   '*** 新規フラグクリアー ***
'
'   Call cmd_OK_Click
End Sub

Private Sub File1_Click()
    txt_file.Text = File1.FileName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Width = 480 + (7185 - 960) * HyoujiBairitu + 480
    Height = 480 + (4710 - 960) * HyoujiBairitu + 480
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
    Me.Caption = STATUS & "(RD)"
'
    If FLGconst = 0 Then
        cmd_New.Enabled = True
    Else
        cmd_New.Enabled = False
    End If
    Call SET_FLbox       '*** ファイルリストボックスの設定 ***
    FLGesc = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1
    End If
End Sub

Private Sub Timer1_Timer()
    If FLGback = 1 Then
        FLGesc = 1
        Unload Me
    End If
End Sub

Private Sub txt_file_Change()
    FLGshinki = 0   '*** 新規フラグクリアー ***
End Sub

Private Sub SET_FLbox()
    Dim i As Integer, j As Integer
'
    FLGback = 0     '*** エラーで戻りフラグリセット ***
'
kaisi:
On Error GoTo NoFile
'
    If FLGlevel = 0 Then
        txt_file.Text = "constlst.cod"
    Else
        txt_file.Text = CURR_file
    End If
'
    lbl_dir.Caption = TMPdir1 & "\"
    File1.Path = TMPdir1 & "\"
    File1.Pattern = "*.cod"
Exit Sub
'
NoFile: Beep
    i = MsgBox("構成表が収容されているフォルダーを準備しましたか？", vbQuestion Or vbRetryCancel, STATUS)
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
    lblnamae.Top = 480
    lblnamae.Left = 480
    lblnamae.Width = 2295 * HyoujiBairitu
    lblnamae.Height = 255 * HyoujiBairitu
    lblnamae.FontSize = 10 * HyoujiBairitu
'
    lbl_folder.Top = 480 + (960 - 480) * HyoujiBairitu
    lbl_folder.Left = 480
    lbl_folder.Width = 1335 * HyoujiBairitu
    lbl_folder.Height = 195 * HyoujiBairitu
    lbl_folder.FontSize = 9 * HyoujiBairitu
'
    lbl_dir.Top = 480 + (1200 - 480) * HyoujiBairitu
    lbl_dir.Left = 480
    lbl_dir.Width = 2295 * HyoujiBairitu
    lbl_dir.Height = 495 * HyoujiBairitu
    lbl_dir.FontSize = 10 * HyoujiBairitu
'
    lblTadasigaki.Top = 480 + (1800 - 480) * HyoujiBairitu
    lblTadasigaki.Left = 480
    lblTadasigaki.Width = 2175 * HyoujiBairitu
    lblTadasigaki.Height = 255 * HyoujiBairitu
    lblTadasigaki.FontSize = 9 * HyoujiBairitu
'
    txt_file.Top = 480
    txt_file.Left = 480 + (2880 - 480) * HyoujiBairitu
    txt_file.Width = 1935 * HyoujiBairitu
    txt_file.Height = 285 * HyoujiBairitu
    txt_file.FontSize = 10 * HyoujiBairitu
'
    File1.Top = 480 + (840 - 480) * HyoujiBairitu
    File1.Left = 480 + (2880 - 480) * HyoujiBairitu
    File1.Width = 2175 * HyoujiBairitu
    File1.Height = 3015 * HyoujiBairitu
    File1.FontSize = 10 * HyoujiBairitu
'
    cmd_OK.Top = 480
    cmd_OK.Left = 480 + (5400 - 480) * HyoujiBairitu
    cmd_OK.Width = 1215 * HyoujiBairitu
    cmd_OK.Height = 735 * HyoujiBairitu
    cmd_OK.FontSize = 10 * HyoujiBairitu
'
    cmd_std.Top = 480 + (1440 - 480) * HyoujiBairitu
    cmd_std.Left = 480 + (5400 - 480) * HyoujiBairitu
    cmd_std.Width = 1215 * HyoujiBairitu
    cmd_std.Height = 615 * HyoujiBairitu
    cmd_std.FontSize = 10 * HyoujiBairitu
'
    cmd_New.Top = 480 + (2280 - 480) * HyoujiBairitu
    cmd_New.Left = 480 + (5400 - 480) * HyoujiBairitu
    cmd_New.Width = 1215 * HyoujiBairitu
    cmd_New.Height = 615 * HyoujiBairitu
    cmd_New.FontSize = 10 * HyoujiBairitu
'
    cmd_Quit.Top = 480 + (3240 - 480) * HyoujiBairitu
    cmd_Quit.Left = 480 + (5400 - 480) * HyoujiBairitu
    cmd_Quit.Width = 1215 * HyoujiBairitu
    cmd_Quit.Height = 615 * HyoujiBairitu
    cmd_Quit.FontSize = 10 * HyoujiBairitu
End Sub

