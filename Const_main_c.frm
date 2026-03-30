VERSION 5.00
Begin VB.Form Const_main_c 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "EＥOS 電気 構成表"
   ClientHeight    =   1695
   ClientLeft      =   2580
   ClientTop       =   1980
   ClientWidth     =   4215
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
   Icon            =   "Const_main_c.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   1695
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancel 
      Caption         =   "中止(&C)"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdkousin 
      Caption         =   "決定 (&G)"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtnyuuryoku 
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
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblitimei 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Const_main_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************
'   構成表変更追加画面
'***********************
'
Option Explicit

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 360 + (4335 - 720) * HyoujiBairitu + 360
    Height = 360 + (2100 - 600) * HyoujiBairitu + 240
    Left = Const_main.Left + (Const_main.Width - Width) \ 2
    Top = Const_main.Top + (Const_main.Height - Height) \ 2
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    If FLGtuika = 1 Then
        If TMProw <= KtotalT Then
            Me.Caption = STATUS & "  >行挿入<"
        Else
            Me.Caption = STATUS & "  >行追加<"
            TMProw = KtotalT
        End If
        TMPcol = 1
    Else
        Me.Caption = STATUS & "  <内容変更>"
    End If
'
    Col_hyouji      '*** 記入位置の表示 ***
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Col_hyouji()
        Select Case TMPcol
        Case 1
            lblitimei.Caption = "名  称"
        Case 2
            lblitimei.Caption = "回路図"
        Case 3
            lblitimei.Caption = "電気部品表"
        Case 4
            lblitimei.Caption = "ﾊﾟﾀｰﾝ番号"
        Case 5
            lblitimei.Caption = "数"
        Case 6
            lblitimei.Caption = "備  考"
        End Select
'
        txtnyuuryoku.Text = KLSTT(TMProw, TMPcol - 1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdkousin_Click()
    If txtnyuuryoku.Text = "" And TMPcol = 6 Then
        txtnyuuryoku.Text = "*"
    ElseIf txtnyuuryoku.Text = "dou" Or txtnyuuryoku.Text = "DOU" Then
        txtnyuuryoku.Text = "〃"
    End If
'
    KLSTT(TMProw, TMPcol - 1) = Trim(txtnyuuryoku.Text)
    Const_main.MSFlexGrid1.Row = TMProw
    Const_main.MSFlexGrid1.Col = TMPcol
    Const_main.MSFlexGrid1.Text = " " & KLSTT(TMProw, TMPcol - 1)
'
    FLGchange = 1
'
    If FLGtuika = 1 Then
        Select Case TMPcol
        Case 1
            TMPcol = 2
            Col_hyouji      '*** 記入位置の表示 ***
            txtnyuuryoku.SetFocus
        Case 2
            TMPcol = 3
            Col_hyouji      '*** 記入位置の表示 ***
            txtnyuuryoku.SetFocus
        Case 3
            TMPcol = 4
            Col_hyouji      '*** 記入位置の表示 ***
            txtnyuuryoku.SetFocus
        Case 4
            TMPcol = 5
            Col_hyouji      '*** 記入位置の表示 ***
            txtnyuuryoku.SetFocus
        Case 5
            TMPcol = 6
            Col_hyouji      '*** 記入位置の表示 ***
            txtnyuuryoku.SetFocus
        Case 6
            FLGtuika = 0    '*** 追加フラグクリアー ***
            Unload Me
        End Select
    Else
        Unload Me
    End If
End Sub

Private Sub DSPgamenBuhin()
    txtnyuuryoku.Left = 360 + (1560 - 360) * HyoujiBairitu
    txtnyuuryoku.Top = 360
    txtnyuuryoku.FontSize = 10 * HyoujiBairitu
    txtnyuuryoku.Width = 2295 * HyoujiBairitu
    txtnyuuryoku.Height = 285 * HyoujiBairitu
'
    lblitimei.Left = 360
    lblitimei.Top = 360
    lblitimei.FontSize = 10 * HyoujiBairitu
    lblitimei.Width = 1215 * HyoujiBairitu
    lblitimei.Height = txtnyuuryoku.Height
'
    cmdkousin.Left = 360 + (720 - 360) * HyoujiBairitu
    cmdkousin.Top = 360 + (960 - 360) * HyoujiBairitu
    cmdkousin.FontSize = 10 * HyoujiBairitu
    cmdkousin.Width = 1215 * HyoujiBairitu
    cmdkousin.Height = 495 * HyoujiBairitu
'
    cmdCancel.Left = 360 + (2280 - 360) * HyoujiBairitu
    cmdCancel.Top = 360 + (960 - 360) * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1215 * HyoujiBairitu
    cmdCancel.Height = 495 * HyoujiBairitu
End Sub
