VERSION 5.00
Begin VB.Form Kankyow_Itiran 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ＥＥＯＳ２ 環境設定"
   ClientHeight    =   4935
   ClientLeft      =   840
   ClientTop       =   1575
   ClientWidth     =   8415
   FillColor       =   &H00FFFFFF&
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
   Icon            =   "Kankyow_Itiran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4935
   ScaleWidth      =   8415
   Begin VB.TextBox Text3 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1440
      Width           =   4935
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Kankyow_Itiran.frx":030A
      Left            =   3000
      List            =   "Kankyow_Itiran.frx":030C
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1920
      Width           =   4935
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'ｵﾌ固定
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   13
      Text            =   "**********"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   960
      Width           =   4935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Kankyow_Itiran.frx":030E
      Left            =   3000
      List            =   "Kankyow_Itiran.frx":0310
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ｷｬﾝｾﾙ(&E)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "更新(&U)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Text            =   "Text13"
      Top             =   2400
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "閉じる(&Q)"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "構成/部品表(個別)ﾌｫﾙﾀﾞ-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   4
      ToolTipText     =   "部品表フォルダ「PARTLIST」の場所を指定します。 「Default設定」になります。"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ｽﾀｰﾄｱｯﾌﾟ 画面"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   0
      ToolTipText     =   "立ち上げた時に表示する画面を設定します。"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部品表 (共有)ﾌｫﾙﾀﾞｰ "
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   6
      ToolTipText     =   "部品表フォルダ「PARTLIST」の場所を指定します。 「Default設定」になります。"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblpass 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "パスワード"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   12
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label13 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾜｰｸﾌｫﾙﾀﾞｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   8
      ToolTipText     =   "EEOSが内部で作業するフォルダ名を指定します。"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "利用者グループ"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   10
      ToolTipText     =   "所属部署を設定します。"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ｺｰﾄﾞ表 ﾌｫﾙﾀﾞｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   2
      ToolTipText     =   "部品ｺｰﾄﾞ表のあるフォルダ名を設定します。"
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Kankyow_Itiran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*****************
'*  環境設定画面  *
'*****************
'
Option Explicit
'

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (8505 - 960) * HyoujiBairitu + 480
    Height = 480 + (5415 - 840) * HyoujiBairitu + 360
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
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    Kankyow_Itiran.Caption = "環境設定"
'
    With Combo2
        .AddItem "電気 構成表"
        .AddItem "電気 部品表"
        .AddItem "部品ｺｰﾄﾞ表"
        .AddItem "ﾒｰｶｰｺｰﾄﾞ表"
        .AddItem "商社ｺｰﾄﾞ表"
        .AddItem "環境設定"
    End With
'
    With Combo1
        .AddItem "生産部門"
        .AddItem "資材部門"
        .AddItem "技術部門"
        .AddItem "ビジター"
        .AddItem "ｽｰﾊﾟﾊﾞｲｻﾞｰ"
    End With
'
    DSPkankyow1
'
    Command2.Enabled = False
    Command3.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub DSPkankyow1()
    Dim busyo$, a%
'
    If Val(Xcont0(1)) = 0 Then
        Xcont0(1) = 3
    End If
    Combo2.ListIndex = Val(Xcont0(1) - 1)
'
    Text2.Text = Xcont0(2)
    Text3.Text = Xcont0(3)
    Text4.Text = Xcont0(4)
    'Text5.Text = Xcont0(5)
'
    busyo$ = Xcont0(8)
        SELgroup busyo$, a%
        Combo1.ListIndex = a%
'
    Text13.Text = Xcont0(13)
'
    SETkoumoku
End Sub
Private Sub GETgroup(a%, busyo$)
    Select Case a%
    Case 4
        busyo$ = "_"
    Case 3
        busyo$ = "*"
    Case 2
        busyo$ = "2"
    Case 1
        busyo$ = "1"
    Case 0
        busyo$ = "0"
    End Select
End Sub

Private Sub SELgroup(busyo$, a%)
    Select Case busyo$
    Case "_"
        a% = 4
    Case "*"
        a% = 3
    Case "2"
        a% = 2
    Case "1"
        a% = 1
    Case "0"
        a% = 0
    End Select
End Sub

Private Sub SETkankyow1()
    Dim busyo$, a%
'
    Xcont0(1) = Combo2.ListIndex + 1
    Xcont0(2) = Text2.Text
    Xcont0(3) = Text3.Text
    Xcont0(4) = Text4.Text
    Xcont0(5) = "*" 'Text5.Text
    Xcont0(6) = "*"
    Xcont0(7) = "*"
'
        a% = Combo1.ListIndex
        GETgroup a%, busyo$
    Xcont0(8) = busyo$
'
    Xcont0(9) = "*"
    Xcont0(10) = "*"
    Xcont0(11) = "*"
    Xcont0(12) = "*"
    Xcont0(13) = Text13.Text
    Xcont0(14) = "*"
'   Xcont0(15) = "*"    *** オプション設定で使用 ***
'   Xcont0(16) = "*"    *** オプション設定で使用 ***
End Sub

Private Sub SETkoumoku()
    Select Case Combo1.ListIndex
    Case 0        '*** 生産部 ***
        lblpass.Visible = False
        txtpass.Visible = False
        txtpass.Text = "**********"
    Case 1        '*** 資材部 ***
        lblpass.Visible = False
        txtpass.Visible = False
        txtpass.Text = "**********"
    Case 2        '*** 技術部 ***
        lblpass.Visible = False
        txtpass.Visible = False
        txtpass.Text = "**********"
    Case 3        '*** ビジター ***
        lblpass.Visible = False
        txtpass.Visible = False
        txtpass.Text = "**********"
    Case 4        '*** ｽｰﾊﾟﾊﾞｲｻﾞｰ ***
        lblpass.Visible = True
        txtpass.Visible = True
    End Select
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 4 Then
        lblpass.Visible = True
        txtpass.Visible = True
        txtpass.Text = "****--****"
    Else
        lblpass.Visible = False
        txtpass.Visible = False
        txtpass.Text = "**********"
    End If
'
    Command2.Enabled = True
    Command3.Enabled = True
End Sub

Private Sub Combo2_Click()
    Command2.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If txtpass.Text = "**********" Then
        SETkankyow1
        WRcont
    Else
        Beep
    End If
'
    DSPkankyow1
'
    Command2.Enabled = False
    Command3.Enabled = False
End Sub

Private Sub Command3_Click()
    RDcont
    DSPkankyow1
'
    Command3.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub Text2_Click()
    TMPdir1 = Xcont0(2)
    Tmpleft = Text2.Left
    Tmptop = Text2.Top
    TMPiti = 1
    Kankyow_Dir.Show 1
End Sub

Private Sub Text3_Click()
    TMPdir1 = Xcont0(3)
    Tmpleft = Text3.Left
    Tmptop = Text3.Top
    TMPiti = 2
    Kankyow_Dir.Show 1
End Sub

Private Sub Text4_Click()
    TMPdir1 = Xcont0(4)
    Tmpleft = Text4.Left
    Tmptop = Text4.Top
    TMPiti = 3
    Kankyow_Dir.Show 1
End Sub

'Private Sub Text5_Click()
'    TMPdir1 = Xcont0(5)
'    Tmpleft = Text5.Left
'    Tmptop = Text5.Top
'    TMPiti = 4
'    Kankyow_Dir.Show 1
'End Sub

Private Sub Text13_Click()
    TMPdir1 = Xcont0(13)
    Tmpleft = Text13.Left
    Tmptop = Text13.Top
    TMPiti = 4
    Kankyow_Dir.Show 1
End Sub

Private Sub txtpass_LostFocus()
    If txtpass.Text = "Version0.4" Then
        txtpass.Text = "**********"
    Else
        txtpass.Text = "ちがいます！"
    End If
End Sub

Private Sub DSPgamenBuhin()
    Combo2.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    Combo2.Top = 480
    Combo2.FontSize = 10 * HyoujiBairitu
    Combo2.Width = 2055 * HyoujiBairitu
'   Combo2.Height = 315 * HyoujiBairitu
'
    Label1.Left = 480
    Label1.Top = 480
    Label1.FontSize = 10 * HyoujiBairitu
    Label1.Width = 2535 * HyoujiBairitu
    Label1.Height = Combo2.Height           '*** combo2に合わせる ***
'
    Text2.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    Text2.Top = 480 + (960 - 480) * HyoujiBairitu
    Text2.FontSize = 10 * HyoujiBairitu
    Text2.Width = 4935 * HyoujiBairitu
    Text2.Height = 315 * HyoujiBairitu
'
    Label2.Left = 480
    Label2.Top = 480 + (960 - 480) * HyoujiBairitu
    Label2.FontSize = 10 * HyoujiBairitu
    Label2.Width = 2535 * HyoujiBairitu
    Label2.Height = Text2.Height
'
    Text3.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    Text3.Top = 480 + (1440 - 480) * HyoujiBairitu
    Text3.FontSize = 10 * HyoujiBairitu
    Text3.Width = 4935 * HyoujiBairitu
    Text3.Height = 315 * HyoujiBairitu
'
    Label3.Left = 480
    Label3.Top = 480 + (1440 - 480) * HyoujiBairitu
    Label3.FontSize = 10 * HyoujiBairitu
    Label3.Width = 2535 * HyoujiBairitu
    Label3.Height = Text2.Height
'
    Text4.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    Text4.Top = 480 + (1920 - 480) * HyoujiBairitu
    Text4.FontSize = 10 * HyoujiBairitu
    Text4.Width = 4935 * HyoujiBairitu
    Text4.Height = 315 * HyoujiBairitu
'
    Label4.Left = 480
    Label4.Top = 480 + (1920 - 480) * HyoujiBairitu
    Label4.FontSize = 10 * HyoujiBairitu
    Label4.Width = 2535 * HyoujiBairitu
    Label4.Height = Text4.Height
'
'    Text5.Left = 480 + (3000 - 480) * HyoujiBairitu
'    Text5.Top = 480 + (3000 - 480) * HyoujiBairitu
'    Text5.FontSize = 10 * HyoujiBairitu
'    Text5.Width = 4935 * HyoujiBairitu
'    Text5.Height = 315 * HyoujiBairitu
'
'    Label5.Left = 480
'    Label5.Top = 480 + (3000 - 480) * HyoujiBairitu
'    Label5.FontSize = 10 * HyoujiBairitu
'    Label5.Width = 2535 * HyoujiBairitu
'    Label5.Height = Text4.Height
'
    Text13.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    Text13.Top = 480 + (2400 - 480) * HyoujiBairitu
    Text13.FontSize = 10 * HyoujiBairitu
    Text13.Width = 4935 * HyoujiBairitu
    Text13.Height = 285 * HyoujiBairitu
'
    Label13.Left = 480
    Label13.Top = 480 + (2400 - 480) * HyoujiBairitu
    Label13.FontSize = 10 * HyoujiBairitu
    Label13.Width = 2535 * HyoujiBairitu
    Label13.Height = Text13.Height
'
    Combo1.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    Combo1.Top = 480 + (2880 - 480) * HyoujiBairitu
    Combo1.FontSize = 10 * HyoujiBairitu
    Combo1.Width = 2055 * HyoujiBairitu
'   Combo1.Height = 315 * HyoujiBairitu
'
    Label8.Left = 480
    Label8.Top = 480 + (2880 - 480) * HyoujiBairitu
    Label8.FontSize = 10 * HyoujiBairitu
    Label8.Width = 2535 * HyoujiBairitu
    Label8.Height = Combo1.Height
'
    txtpass.Left = 480 + (3000 - 480) * HyoujiBairitu + 10
    txtpass.Top = 480 + (3360 - 480) * HyoujiBairitu
    txtpass.FontSize = 10 * HyoujiBairitu
    txtpass.Width = 1215 * HyoujiBairitu
    txtpass.Height = 285 * HyoujiBairitu
'
    lblpass.Left = 480
    lblpass.Top = 480 + (3360 - 480) * HyoujiBairitu
    lblpass.FontSize = 10 * HyoujiBairitu
    lblpass.Width = 2535 * HyoujiBairitu
    lblpass.Height = txtpass.Height
'
    Command3.Left = 480 + (1320 - 480) * HyoujiBairitu
    Command3.Top = 480 + (3960 - 480) * HyoujiBairitu
    Command3.FontSize = 10 * HyoujiBairitu
    Command3.Width = 1455 * HyoujiBairitu
    Command3.Height = 495 * HyoujiBairitu
'
    Command2.Left = 480 + (3240 - 480) * HyoujiBairitu
    Command2.Top = 480 + (3960 - 480) * HyoujiBairitu
    Command2.FontSize = 10 * HyoujiBairitu
    Command2.Width = 1455 * HyoujiBairitu
    Command2.Height = 495 * HyoujiBairitu
'
    Command1.Left = 480 + (5160 - 480) * HyoujiBairitu
    Command1.Top = 480 + (3960 - 480) * HyoujiBairitu
    Command1.FontSize = 10 * HyoujiBairitu
    Command1.Width = 1455 * HyoujiBairitu
    Command1.Height = 495 * HyoujiBairitu
End Sub

