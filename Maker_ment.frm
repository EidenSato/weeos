VERSION 5.00
Begin VB.Form Maker_ment 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ﾒｰｶｰｺｰﾄﾞ ＜帳票形式＞"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Maker_ment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5430
   ScaleWidth      =   7215
   Begin VB.TextBox Text0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   "Text0"
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox cboTrader2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "cboTrader2"
      Top             =   2880
      Width           =   3855
   End
   Begin VB.ComboBox cboTrader1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "cboTrader1"
      Top             =   2520
      Width           =   3855
   End
   Begin VB.ComboBox cboTrader0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Maker_ment.frx":030A
      Left            =   2280
      List            =   "Maker_ment.frx":030C
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "cboTrader0"
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "決定保存(&G)"
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
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
      Left            =   5520
      TabIndex        =   24
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ﾒｰｶｰ追加(&A)"
      Enabled         =   0   'False
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
      Left            =   720
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "内容変更(&C)"
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   20
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   4560
      TabIndex        =   22
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UP"
      Height          =   495
      Left            =   3600
      TabIndex        =   23
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Text7"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Text6"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label0 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "コード"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "更新日"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "買先-Ｃ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "買先-Ｂ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "買先-Ａ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "名称 英"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "名称 和"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "略 称"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Maker_ment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'****************************
'*  メーカーコード細目一覧   *
'****************************
'
Option Explicit
'
    Dim FLG_CorA As Integer
    Dim FLG_Motonoiro As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If FLG_CorA = 1 Or FLG_CorA = 2 Then Exit Sub   '*** 変更追加画面の時はキーが効かない ***
'
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        If Makps > 10 Then
            Makps = Makps - 10
        Else
            Beep
            Makps = 1
        End If
'
        DSPlevel2      '*** 品種 細目表示 ***
'
    Case vbKeyUp        '*** ↑ ***
        If Makps > 1 Then
            Makps = Makps - 1
        Else
            Beep
        End If
'
        Call DSPlevel2  '*** 品種 細目表示 ***
'
    Case vbKeyPageUp    '*** Roll Down
        If Makps + 10 <= Maknum0 Then
            Makps = Makps + 10
        Else
            Beep
            Makps = Maknum0
        End If
'
        Call DSPlevel2  '*** 品種 細目表示 ***
'
    Case vbKeyDown      '*** ↓ ***
        If Makps + 1 <= Maknum0 Then
            Makps = Makps + 1
        Else
            Beep
        End If
'
        Call DSPlevel2  '*** 品種 細目表示 ***
'
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 720 + (7305 - 1440) * HyoujiBairitu + 720
    Height = 360 + (5805 - 720) * HyoujiBairitu + 360
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
    Call DSPgamenBuhin
'
    Call DSPlevel2       '*** メーカー 細目一覧 ***
'
    cboTrader0.Visible = False
    cboTrader1.Visible = False
    cboTrader2.Visible = False
'
    If Xcont0(8) = "_" Then
        Command3.Enabled = True
        Command4.Enabled = True
        Command9.Visible = True
        Command9.Enabled = False
    Else
        Command3.Enabled = False
        Command4.Enabled = False
        Command9.Visible = False
    End If
'
    Call DSPpointer(1)      '*** マウスポインタセット ***
'
    FLG_CorA = 0
    FLG_Motonoiro = Me.BackColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub
    
Private Sub DSPlevel2()    '*** メーカーコード 細目表示 ***
    Dim temp As String
'
    Text0.Text = Maker(Makps, 0)
    Text1.Text = Maker(Makps, 1)
    Text2.Text = Maker(Makps, 2)
    Text3.Text = Maker(Makps, 3)
    temp = Maker(Makps, 4)
        GETtrader1 temp
    Text4.Text = Maker(Makps, 4) & "  " & temp
    temp = Maker(Makps, 5)
        GETtrader1 temp
    Text5.Text = Maker(Makps, 5) & "  " & temp
    temp = Maker(Makps, 6)
        GETtrader1 temp
    Text6.Text = Maker(Makps, 6) & "  " & temp
    Text7.Text = Maker(Makps, 7)
End Sub

Private Sub Command1_Click()
    If Makps > 1 Then
        Makps = Makps - 1
    Else
        Beep
    End If
'
    DSPlevel2      '*** 品種 細目表示 ***
End Sub

Private Sub Command2_Click()
    If Makps < Maknum0 Then
        Makps = Makps + 1
    Else
        Beep
    End If
'
    DSPlevel2      '*** 品種 細目表示 ***
End Sub

Private Sub Command3_Click()    '*** 内容変更 ***
    Command3.Enabled = False
    FLG_CorA = 1
    Call Set_Gamen_CorA     '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)      '*** マウスポインターセット ***
'
    Text1.SetFocus
End Sub

Private Sub Command4_Click()    '*** 品種追加 ***
    Command4.Enabled = False
    FLG_CorA = 2
    Call Set_Gamen_CorA     '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)      '*** マウスポインターセット ***
'
    Text0.SetFocus
End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Command9_Click()    '*** 決定 ***
    Dim Tempcode As String
    Dim i As Integer, j As Integer
'
    Select Case FLG_CorA
    Case 1
        Call GET_Chg_Data   '*** 変更データ取り込み ***
'
    Case 2
        Tempcode = Trim(Text0.Text)
        i = Len(Tempcode)
        If i <> 3 Then
            j = MsgBox("正しいコード番号を記入してください。", vbCritical)
            Exit Sub
'
        End If
'
        For i = 1 To Maknum0
            If Maker(i, 0) = Tempcode Then
                j = MsgBox("同じコード番号がすでにあります。", vbExclamation)
                Exit Sub
'
            End If
        Next i
'
        Call Ins_maker(Tempcode)    '*** データを挿入する場所(Makps)を作る ***
'
        Maker(Makps, 0) = Tempcode
        Call GET_Chg_Data   '*** 変更データ取り込み ***
'
    End Select
'
    Text0.Enabled = True   '*** 表示を元に戻す ***
    Text0.TabStop = False
    Text1.TabStop = False
    Text2.TabStop = False
    Text3.TabStop = False
    Text4.TabStop = False
    Text5.TabStop = False
    Text6.TabStop = False
    Text7.TabStop = False
    Me.BackColor = FLG_Motonoiro
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command9.Enabled = False
'
    cboTrader0.Visible = False
    cboTrader1.Visible = False
    cboTrader2.Visible = False
    Text4.Visible = True
    Text5.Visible = True
    Text6.Visible = True
'
    Me.MousePointer = vbHourglass
'
    Call WRmaker     '*** メーカーデータセーブ ***
    Call RDmaker     '*** メーカーデータ再読み込み ***
    Call DSPlevel2   '*** メーカー 細目表示 ***
'
    Call DSPpointer(0)  '*** マウスポインターセット ***
'
    Me.MousePointer = vbDefault
    FLG_CorA = 0
'
    FLGmak_data_change = 1      '*** 変更有り ***
'
End Sub

Private Sub Ins_maker(Tempcode As String)  '*** データを挿入する場所を作る ***
    Dim i As Integer, j As Integer
    Dim DirStr As String
'
    For i = 1 To Maknum0
        If Tempcode < Maker(i, 0) Then
            Exit For
'
        End If
    Next i
'
    Makps = i
'
    If Makps <= Maknum0 Then
        For i = Maknum0 To Makps Step -1
            For j = 0 To Makdim0
                Maker(i + 1, j) = Maker(i, j)
            Next j
        Next i
    End If
'
    Maknum0 = Maknum0 + 1
End Sub

Private Sub GET_Chg_Data()      '*** 変更データ取り込み ***
    Dim i As Integer
'
    Maker(Makps, 1) = Trim(Text1.Text)
    Maker(Makps, 2) = Trim(Text2.Text)
    Maker(Makps, 3) = Trim(Text3.Text)
'
    i = cboTrader0.ListIndex
    If i = 0 Then
        Maker(Makps, 4) = "*"
    Else
        Maker(Makps, 4) = Trader(i, 0)
    End If
'
    i = cboTrader1.ListIndex
    If i = 0 Then
        Maker(Makps, 5) = "*"
    Else
        Maker(Makps, 5) = Trader(i, 0)
    End If
'
    i = cboTrader2.ListIndex
    If i = 0 Then
        Maker(Makps, 6) = "*"
    Else
        Maker(Makps, 6) = Trader(i, 0)
    End If
'
    Maker(Makps, 7) = Format(Date, "yy/mm/dd")
End Sub

Private Sub DSPgamenBuhin()
    Text0.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text0.Top = 360
    Text0.FontSize = 10 * HyoujiBairitu
    Text0.Width = 495 * HyoujiBairitu
    Text0.Height = 285 * HyoujiBairitu
'
    Label0.Left = 720
    Label0.Top = 360
    Label0.FontSize = 10 * HyoujiBairitu
    Label0.Width = 975 * HyoujiBairitu
    Label0.Height = Text0.Height
'
    Text1.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text1.Top = 360 + (720 - 360) * HyoujiBairitu
    Text1.FontSize = 10 * HyoujiBairitu
    Text1.Width = 855 * HyoujiBairitu
    Text1.Height = 285 * HyoujiBairitu
'
    Label1.Left = 720
    Label1.Top = 360 + (720 - 360) * HyoujiBairitu
    Label1.FontSize = 10 * HyoujiBairitu
    Label1.Width = 975 * HyoujiBairitu
    Label1.Height = Text1.Height
'
    Text2.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text2.Top = 360 + (1200 - 360) * HyoujiBairitu
    Text2.FontSize = 10 * HyoujiBairitu
    Text2.Width = 4815 * HyoujiBairitu
    Text2.Height = 285 * HyoujiBairitu
'
    Label2.Left = 720
    Label2.Top = 360 + (1200 - 360) * HyoujiBairitu
    Label2.FontSize = 10 * HyoujiBairitu
    Label2.Width = 975 * HyoujiBairitu
    Label2.Height = Text2.Height
'
    Text3.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text3.Top = 360 + (1560 - 360) * HyoujiBairitu
    Text3.FontSize = 10 * HyoujiBairitu
    Text3.Width = 4815 * HyoujiBairitu
    Text3.Height = 285 * HyoujiBairitu
'
    Label3.Left = 720
    Label3.Top = 360 + (1560 - 360) * HyoujiBairitu
    Label3.FontSize = 10 * HyoujiBairitu
    Label3.Width = 975 * HyoujiBairitu
    Label3.Height = Text3.Height
'
    Text4.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text4.Top = 360 + (2040 - 360) * HyoujiBairitu
    Text4.FontSize = 10 * HyoujiBairitu
    Text4.Width = 3855 * HyoujiBairitu
    Text4.Height = 285 * HyoujiBairitu
'
    cboTrader0.Left = 720 + (1695 - 720) * HyoujiBairitu
    cboTrader0.Top = 360 + (2040 - 360) * HyoujiBairitu
    cboTrader0.FontSize = 10 * HyoujiBairitu
    cboTrader0.Width = 3855 * HyoujiBairitu
'   cboTrader0.Height = 315 * HyoujiBairitu
'
    Label4.Left = 720
    Label4.Top = 360 + (2040 - 360) * HyoujiBairitu
    Label4.FontSize = 10 * HyoujiBairitu
    Label4.Width = 975 * HyoujiBairitu
    Label4.Height = Text4.Height
'
    Text5.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text5.Top = 360 + (2400 - 360) * HyoujiBairitu
    Text5.FontSize = 10 * HyoujiBairitu
    Text5.Width = 3855 * HyoujiBairitu
    Text5.Height = 285 * HyoujiBairitu
'
    cboTrader1.Left = 720 + (1695 - 720) * HyoujiBairitu
    cboTrader1.Top = 360 + (2400 - 360) * HyoujiBairitu
    cboTrader1.FontSize = 10 * HyoujiBairitu
    cboTrader1.Width = 3855 * HyoujiBairitu
'   cboTrader1.Height = 315 * HyoujiBairitu
'
    Label5.Left = 720
    Label5.Top = 360 + (2400 - 360) * HyoujiBairitu
    Label5.FontSize = 10 * HyoujiBairitu
    Label5.Width = 975 * HyoujiBairitu
    Label5.Height = Text5.Height
'
    Text6.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text6.Top = 360 + (2760 - 360) * HyoujiBairitu
    Text6.FontSize = 10 * HyoujiBairitu
    Text6.Width = 3855 * HyoujiBairitu
    Text6.Height = 285 * HyoujiBairitu
'
    cboTrader2.Left = 720 + (1695 - 720) * HyoujiBairitu
    cboTrader2.Top = 360 + (2760 - 360) * HyoujiBairitu
    cboTrader2.FontSize = 10 * HyoujiBairitu
    cboTrader2.Width = 3855 * HyoujiBairitu
'   cboTrader2.Height = 315 * HyoujiBairitu
'
    Label6.Left = 720
    Label6.Top = 360 + (2760 - 360) * HyoujiBairitu
    Label6.FontSize = 10 * HyoujiBairitu
    Label6.Width = 975 * HyoujiBairitu
    Label6.Height = Text6.Height
'
    Text7.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text7.Top = 360 + (3240 - 360) * HyoujiBairitu
    Text7.FontSize = 10 * HyoujiBairitu
    Text7.Width = 1095 * HyoujiBairitu
    Text7.Height = 285 * HyoujiBairitu
'
    Label7.Left = 720
    Label7.Top = 360 + (3240 - 360) * HyoujiBairitu
    Label7.FontSize = 10 * HyoujiBairitu
    Label7.Width = 975 * HyoujiBairitu
    Label7.Height = Text7.Height
'
    Command9.Left = 720 + (2160 - 720) * HyoujiBairitu
    Command9.Top = 360 + (3840 - 360) * HyoujiBairitu
    Command9.FontSize = 9 * HyoujiBairitu
    Command9.Width = 1215 * HyoujiBairitu
    Command9.Height = 495 * HyoujiBairitu
'
    Command4.Left = 720
    Command4.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command4.FontSize = 9 * HyoujiBairitu
    Command4.Width = 1215 * HyoujiBairitu
    Command4.Height = 495 * HyoujiBairitu
'
    Command3.Left = 720 + (2160 - 720) * HyoujiBairitu
    Command3.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command3.FontSize = 9 * HyoujiBairitu
    Command3.Width = 1215 * HyoujiBairitu
    Command3.Height = 495 * HyoujiBairitu
'
    Command1.Left = 720 + (3600 - 720) * HyoujiBairitu
    Command1.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command1.FontSize = 10 * HyoujiBairitu
    Command1.Width = 735 * HyoujiBairitu
    Command1.Height = 495 * HyoujiBairitu
'
    Command2.Left = 720 + (4560 - 720) * HyoujiBairitu
    Command2.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command2.FontSize = 10 * HyoujiBairitu
    Command2.Width = 735 * HyoujiBairitu
    Command2.Height = 495 * HyoujiBairitu
'
    Command6.Left = 720 + (5520 - 720) * HyoujiBairitu
    Command6.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command6.FontSize = 9 * HyoujiBairitu
    Command6.Width = 975 * HyoujiBairitu
    Command6.Height = 495 * HyoujiBairitu
End Sub

Private Sub Set_Gamen_CorA()    '*** 追加変更画面設定 ***
    Dim i As Integer
    Dim Mdata As String
'
    If FLG_CorA = 1 Then
        Text0.Enabled = False
    End If
    Text1.TabStop = True
    Text2.TabStop = True
    Text3.TabStop = True
    Text7.TabStop = True
'
    Me.BackColor = &H808080
    Command1.Enabled = False
    Command2.Enabled = False
    If FLG_CorA = 1 Then
        Command4.Enabled = False
    Else
        Command3.Enabled = False
    End If
    Command9.Enabled = True
'
    Text4.Visible = False
    Text5.Visible = False
    Text6.Visible = False
    cboTrader0.Visible = True
    cboTrader0.TabStop = True
    cboTrader0.Clear
    cboTrader1.Visible = True
    cboTrader1.TabStop = True
    cboTrader1.Clear
    cboTrader2.Visible = True
    cboTrader2.TabStop = True
    cboTrader2.Clear
'
    cboTrader0.AddItem "*"
    cboTrader1.AddItem "*"
    cboTrader2.AddItem "*"
    For i = 1 To Trdnum0
        cboTrader0.AddItem Trader(i, 0) & "  " & Trader(i, 1)
        cboTrader1.AddItem Trader(i, 0) & "  " & Trader(i, 1)
        cboTrader2.AddItem Trader(i, 0) & "  " & Trader(i, 1)
    Next i
'
    For i = 1 To Trdnum0
        If Maker(Makps, 4) = Trader(i, 0) Then
            cboTrader0.ListIndex = i
            Exit For
        End If
        cboTrader0.ListIndex = 0
    Next i
'
    For i = 1 To Trdnum0
        If Maker(Makps, 5) = Trader(i, 0) Then
            cboTrader1.ListIndex = i
            Exit For
        End If
        cboTrader1.ListIndex = 0
    Next i
'
    For i = 1 To Trdnum0
        If Maker(Makps, 6) = Trader(i, 0) Then
            cboTrader2.ListIndex = i
            Exit For
        End If
        cboTrader2.ListIndex = 0
    Next i
End Sub

Private Sub DSPpointer(X As Integer)
    If X = 1 Then
        Text0.MousePointer = 1
        Text1.MousePointer = 1
        Text2.MousePointer = 1
        Text3.MousePointer = 1
        Text4.MousePointer = 1
        Text5.MousePointer = 1
        Text6.MousePointer = 1
        Text7.MousePointer = 1
    Else
        Text0.MousePointer = 0
        Text1.MousePointer = 0
        Text2.MousePointer = 0
        Text3.MousePointer = 0
        Text4.MousePointer = 0
        Text5.MousePointer = 0
        Text6.MousePointer = 0
        Text7.MousePointer = 0
    End If
End Sub

Private Sub Text0_LostFocus()
    Dim Tempcode As String
    Dim i As Integer, j As Integer
'
    If FLG_CorA <> 2 Then       '*** 部品追加モード時のみ有効 ***
        Exit Sub
    End If
'
    Tempcode = Trim(Text0.Text)
    i = Len(Tempcode)
    If i <> 3 Then
        j = MsgBox("正しいコード番号を記入してください。", vbCritical)
'
        Text0.SetFocus
        Exit Sub
'
    End If
'
    For i = 1 To Maknum0
        If Maker(i, 0) = Tempcode Then
            j = MsgBox("同じコード番号がすでにあります。", vbExclamation)
            Exit Sub
'
        End If
    Next i
End Sub

