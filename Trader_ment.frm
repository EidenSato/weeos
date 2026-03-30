VERSION 5.00
Begin VB.Form Trader_ment 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "商社ｺｰﾄﾞ ＜帳票形式＞"
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
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Trader_ment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5430
   ScaleWidth      =   7215
   Begin VB.TextBox Text8 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Text8"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   "Text0"
      Top             =   360
      Width           =   615
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
      Left            =   2040
      TabIndex        =   23
      Top             =   3960
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
      Left            =   5640
      TabIndex        =   20
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "商社追加(&A)"
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
      Left            =   600
      TabIndex        =   22
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
      Left            =   2040
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UP"
      Height          =   495
      Left            =   3480
      TabIndex        =   18
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Text7"
      Top             =   3000
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "215-0033"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   2280
      Width           =   4815
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
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "044-988-4811/9887-7635"
      Top             =   1440
      Width           =   2415
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
      Top             =   1080
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
   Begin VB.Label Label8 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "更新日"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   16
      Top             =   3360
      Width           =   975
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
      Caption         =   "所在地"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "郵便番号"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "取扱ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "担当者"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "電話/FAX"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "名称"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   1080
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
Attribute VB_Name = "Trader_ment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'****************************
'*  商社コード細目一覧   *
'****************************
'
Option Explicit
'
    Dim FLG_CorA As Integer, FLG_Motonoiro As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If FLG_CorA = 1 Or FLG_CorA = 2 Then Exit Sub   '*** 変更追加画面の時はキーが効かない ***
'
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        If Trdps > 10 Then
            Trdps = Trdps - 10
        Else
            Beep
            Trdps = 1
        End If
'
        Call DSPlevel2  '*** 細目表示 ***
'
    Case vbKeyUp        '*** ↑ ***
        If Trdps > 1 Then
            Trdps = Trdps - 1
        Else
            Beep
        End If
'
        Call DSPlevel2  '*** 細目表示 ***
'
    Case vbKeyPageUp    '*** Roll Down
        If Trdps + 10 <= Trdnum0 Then
            Trdps = Trdps + 10
        Else
            Beep
            Trdps = Trdnum0
        End If
'
        Call DSPlevel2  '*** 細目表示 ***
'
    Case vbKeyDown      '*** ↓ ***
        If Trdps + 1 <= Trdnum0 Then
            Trdps = Trdps + 1
        Else
            Beep
        End If
'
        Call DSPlevel2  '*** 細目表示 ***
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
    Call DSPlevel2      '*** 商社 細目一覧 ***
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
    Call DSPpointer(1)  '*** マウスポインタセット ***
'
    FLG_CorA = 0
    FLG_Motonoiro = Me.BackColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub
    
Private Sub DSPlevel2()    '*** 商社コード 細目表示 ***
    Text0.Text = Trader(Trdps, 0)
    Text1.Text = Trader(Trdps, 2)
    Text2.Text = Trader(Trdps, 1)
    Text3.Text = Trader(Trdps, 3)
    Text4.Text = Trader(Trdps, 4)
    Text5.Text = Trader(Trdps, 7)
    Text6.Text = Trader(Trdps, 5)
    Text7.Text = Trader(Trdps, 6)
    Text8.Text = Trader(Trdps, 9)
End Sub

Private Sub Command1_Click()
    If Trdps > 1 Then
        Trdps = Trdps - 1
    Else
        Beep
    End If
'
    Call DSPlevel2  '*** 細目表示 ***
End Sub

Private Sub Command2_Click()
    If Trdps < Trdnum0 Then
        Trdps = Trdps + 1
    Else
        Beep
    End If
'
    Call DSPlevel2  '*** 細目表示 ***
End Sub

Private Sub Command3_Click()    '*** 内容変更 ***
    Command3.Enabled = False
    FLG_CorA = 1
    Call Set_Gamen_CorA '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)  '*** マウスポインターセット ***
'
    Text1.SetFocus
End Sub

Private Sub Command4_Click()    '*** 商社追加 ***
    Command4.Enabled = False
    FLG_CorA = 2
    Call Set_Gamen_CorA '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)  '*** マウスポインターセット ***
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
        If i <> 4 Then
            j = MsgBox("正しいコード番号を記入してください。", vbCritical)
'
            Text0.SetFocus
            Exit Sub
'
        End If
'
        For i = 1 To Trdnum0
            If Trader(i, 0) = Tempcode Then
                j = MsgBox("同じコード番号がすでにあります。", vbExclamation)
'
                Text0.SetFocus
                Exit Sub
'
            End If
        Next i
'
        Call Ins_trader(Tempcode)   '*** データを挿入する場所(Trdps)を作る ***
'
        Trader(Trdps, 0) = Tempcode
        Call GET_Chg_Data   '*** 変更データ取り込み ***
'
    End Select
'
    Text0.Enabled = True   '*** 表示を元に戻す ***
'
    Text0.TabStop = False
    Text1.TabStop = False
    Text2.TabStop = False
    Text3.TabStop = False
    Text4.TabStop = False
    Text5.TabStop = False
    Text6.TabStop = False
    Text7.TabStop = False
'
    Text8.Enabled = True
'
    Me.BackColor = FLG_Motonoiro
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command9.Enabled = False
'
    Me.MousePointer = vbHourglass
'
    Call WRtrader   '*** 商社データセーブ ***
    Call RDtrader   '*** 商社データ再読み込み ***
    Call DSPlevel2  '*** 商社 細目表示 ***
'
    Call DSPpointer(0)  '*** マウスポインターセット ***
'
    Me.MousePointer = vbDefault
    FLG_CorA = 0
    FLGtrd_data_change = 1   '*** 変更フラグ設定(メイン画面の表示を変更する)***
'
End Sub

Private Sub Ins_trader(Tempcode As String)  '*** データを挿入する場所を作る ***
    Dim i As Integer, j As Integer
    Dim DirStr As String
'
    For i = 1 To Trdnum0
        If Tempcode < Trader(i, 0) Then
            Exit For
'
        End If
    Next i
'
    Trdps = i
'
    If Trdps <= Trdnum0 Then
        For i = Trdnum0 To Trdps Step -1
            For j = 0 To Trddim0
                Trader(i + 1, j) = Trader(i, j)
            Next j
        Next i
    End If
'
    Trdnum0 = Trdnum0 + 1
End Sub

Private Sub GET_Chg_Data()      '*** 変更データ取り込み ***
    Trader(Trdps, 2) = Trim(Text1.Text)
    Trader(Trdps, 1) = Trim(Text2.Text)
    Trader(Trdps, 3) = Trim(Text3.Text)
    Trader(Trdps, 4) = Trim(Text4.Text)
    Trader(Trdps, 7) = Trim(Text5.Text)
    Trader(Trdps, 5) = Trim(Text6.Text)
    Trader(Trdps, 6) = Trim(Text7.Text)
    Trader(Trdps, 8) = "*"
    Trader(Trdps, 9) = Format(Date, "yy/mm/dd")
End Sub

Private Sub Set_Gamen_CorA()    '*** 追加変更画面設定 ***
    Dim i As Integer
    Dim Mdata As String
'
    If FLG_CorA = 1 Then
        Text0.Enabled = False
    Else
        Text0.TabStop = True
    End If
'
    Text1.TabStop = True
    Text2.TabStop = True
    Text3.TabStop = True
    Text4.TabStop = True
    Text5.TabStop = True
    Text6.TabStop = True
    Text7.TabStop = True
'
    Text8.Enabled = False
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
    If FLG_CorA <> 2 Then       '*** 追加モード時のみ有効 ***
        Exit Sub
    End If
'
    Tempcode = Trim(Text0.Text)
    i = Len(Tempcode)
    If i <> 4 Then
        j = MsgBox("正しいコード番号を記入してください。", vbCritical)
'
        Text0.SetFocus
        Exit Sub
'
    End If
'
    For i = 1 To Trdnum0
        If Trader(i, 0) = Tempcode Then
            j = MsgBox("同じコード番号がすでにあります。", vbExclamation)
            Exit Sub
'
        End If
    Next i
End Sub

Private Sub DSPgamenBuhin()
    Text0.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text0.Top = 360
    Text0.FontSize = 10 * HyoujiBairitu
    Text0.Width = 615 * HyoujiBairitu
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
    Text2.Top = 360 + (1080 - 360) * HyoujiBairitu
    Text2.FontSize = 10 * HyoujiBairitu
    Text2.Width = 4815 * HyoujiBairitu
    Text2.Height = 285 * HyoujiBairitu
'
    Label2.Left = 720
    Label2.Top = 360 + (1080 - 360) * HyoujiBairitu
    Label2.FontSize = 10 * HyoujiBairitu
    Label2.Width = 975 * HyoujiBairitu
    Label2.Height = Text2.Height
'
    Text3.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text3.Top = 360 + (1440 - 360) * HyoujiBairitu
    Text3.FontSize = 10 * HyoujiBairitu
    Text3.Width = 2415 * HyoujiBairitu
    Text3.Height = 285 * HyoujiBairitu
'
    Label3.Left = 720
    Label3.Top = 360 + (1440 - 360) * HyoujiBairitu
    Label3.FontSize = 10 * HyoujiBairitu
    Label3.Width = 975 * HyoujiBairitu
    Label3.Height = Text3.Height
'
    Text4.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text4.Top = 360 + (1800 - 360) * HyoujiBairitu
    Text4.FontSize = 10 * HyoujiBairitu
    Text4.Width = 1455 * HyoujiBairitu
    Text4.Height = 285 * HyoujiBairitu
'
    Label4.Left = 720
    Label4.Top = 360 + (1800 - 360) * HyoujiBairitu
    Label4.FontSize = 10 * HyoujiBairitu
    Label4.Width = 975 * HyoujiBairitu
    Label4.Height = Text4.Height
'
    Text5.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text5.Top = 360 + (2280 - 360) * HyoujiBairitu
    Text5.FontSize = 10 * HyoujiBairitu
    Text5.Width = 4815 * HyoujiBairitu
    Text5.Height = 285 * HyoujiBairitu
'
    Label5.Left = 720
    Label5.Top = 360 + (2280 - 360) * HyoujiBairitu
    Label5.FontSize = 10 * HyoujiBairitu
    Label5.Width = 975 * HyoujiBairitu
    Label5.Height = Text5.Height
'
    Text6.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text6.Top = 360 + (2640 - 360) * HyoujiBairitu
    Text6.FontSize = 10 * HyoujiBairitu
    Text6.Width = 975 * HyoujiBairitu
    Text6.Height = 285 * HyoujiBairitu
'
    Label6.Left = 720
    Label6.Top = 360 + (2640 - 360) * HyoujiBairitu
    Label6.FontSize = 10 * HyoujiBairitu
    Label6.Width = 975 * HyoujiBairitu
    Label6.Height = Text6.Height
'
    Text7.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text7.Top = 360 + (3000 - 360) * HyoujiBairitu
    Text7.FontSize = 10 * HyoujiBairitu
    Text7.Width = 4815 * HyoujiBairitu
    Text7.Height = 285 * HyoujiBairitu
'
    Label7.Left = 720
    Label7.Top = 360 + (3000 - 360) * HyoujiBairitu
    Label7.FontSize = 10 * HyoujiBairitu
    Label7.Width = 975 * HyoujiBairitu
    Label7.Height = Text7.Height
'
    Text8.Left = 720 + (1695 - 720) * HyoujiBairitu
    Text8.Top = 360 + (3360 - 360) * HyoujiBairitu
    Text8.FontSize = 10 * HyoujiBairitu
    Text8.Width = 1095 * HyoujiBairitu
    Text8.Height = 285 * HyoujiBairitu
'
    Label8.Left = 720
    Label8.Top = 360 + (3360 - 360) * HyoujiBairitu
    Label8.FontSize = 10 * HyoujiBairitu
    Label8.Width = 975 * HyoujiBairitu
    Label8.Height = Text8.Height
'
    Command9.Left = 720 + (2040 - 720) * HyoujiBairitu
    Command9.Top = 360 + (3960 - 360) * HyoujiBairitu
    Command9.FontSize = 9 * HyoujiBairitu
    Command9.Width = 1215 * HyoujiBairitu
    Command9.Height = 495 * HyoujiBairitu
'
    Command4.Left = 720 + (600 - 720) * HyoujiBairitu
    Command4.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command4.FontSize = 9 * HyoujiBairitu
    Command4.Width = 1215 * HyoujiBairitu
    Command4.Height = 495 * HyoujiBairitu
'
    Command3.Left = 720 + (2040 - 720) * HyoujiBairitu
    Command3.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command3.FontSize = 9 * HyoujiBairitu
    Command3.Width = 1215 * HyoujiBairitu
    Command3.Height = 495 * HyoujiBairitu
'
    Command1.Left = 720 + (3480 - 720) * HyoujiBairitu
    Command1.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command1.FontSize = 10 * HyoujiBairitu
    Command1.Width = 855 * HyoujiBairitu
    Command1.Height = 495 * HyoujiBairitu
'
    Command2.Left = 720 + (4560 - 720) * HyoujiBairitu
    Command2.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command2.FontSize = 10 * HyoujiBairitu
    Command2.Width = 855 * HyoujiBairitu
    Command2.Height = 495 * HyoujiBairitu
'
    Command6.Left = 720 + (5640 - 720) * HyoujiBairitu
    Command6.Top = 360 + (4560 - 360) * HyoujiBairitu
    Command6.FontSize = 9 * HyoujiBairitu
    Command6.Width = 975 * HyoujiBairitu
    Command6.Height = 495 * HyoujiBairitu
End Sub

