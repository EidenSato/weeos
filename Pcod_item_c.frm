VERSION 5.00
Begin VB.Form Pcod_item_c 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "部品ｺｰﾄﾞ <項目 帳票形式>"
   ClientHeight    =   4590
   ClientLeft      =   780
   ClientTop       =   1335
   ClientWidth     =   7815
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
   Icon            =   "Pcod_item_c.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4590
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdSakujyo 
      Caption         =   "項目削除(&D)"
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
      Left            =   840
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdKettei 
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
      Left            =   2280
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTuika 
      Caption         =   "項目追加(&A)"
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
      Left            =   840
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtGaiyou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox txtZumen3 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtZumen2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtZumen1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtDsk 
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtK_mei 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtKigou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
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
      Left            =   5880
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdKousin 
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
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblZumen3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "図面記号 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblZumen2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "図面記号 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblKigou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "項 目 記 号"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblGaiyou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "概    要"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblZumen1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "図面記号 １"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblDsk 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾊﾞｯｸｱｯﾌﾟﾃﾞｨｽｸ番号"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblK_mei 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "項  目  名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MousePointer    =   1  '矢印
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Pcod_item_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'************************
'*** 項目一覧 帳票形式 ***
'************************
'
Option Explicit
'
    Dim HeadTitle As String
    Dim FLG_CorA As Integer
    Dim FLG_Motonoiro As Long

Private Sub Form_Initialize()
    HeadTitle = "部品ｺｰﾄﾞ <項目 帳票形式>"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If FLG_CorA = 1 Or FLG_CorA = 2 Then Exit Sub   '*** 変更追加画面の時はキーが効かない ***
'
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        Hup             '*** 上へ ***
        DSPlevel11       '*** 項目内容表示 ***
'
    Case vbKeyUp        '*** ↑ ***
        H1up            '*** 一つ上へ ***
        DSPlevel11       '*** 項目内容表示 ***
'
    Case vbKeyPageUp    '*** Roll Down
        Hdown           '*** 下へ ***
        DSPlevel11       '*** 項目内容表示 ***
'
    Case vbKeyDown      '*** ↓ ***
        H1down          '*** 一つ下へ ***
        DSPlevel11       '*** 項目内容表示 ***
'
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()     '*** フォームの表示位置の設定
    Width = 720 + (7905 - 1440) * HyoujiBairitu + 720
    Height = 360 + (4965 - 720) * HyoujiBairitu + 360
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
    Me.Caption = HeadTitle
'
    Call DSPlevel11     '*** 項目内容表示 ***
'
    If Xcont0(8) = "_" Then     '*** ｽｰﾊﾟﾊﾞｲｻﾞｰ ***
        cmdSakujyo.Enabled = True
        cmdKettei.Enabled = False
        cmdTuika.Enabled = True
        cmdkousin.Enabled = True
    Else
        cmdSakujyo.Visible = False
        cmdKettei.Visible = False
        cmdTuika.Enabled = False
        cmdkousin.Enabled = False
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

Private Sub DSPlevel11()         '*** 項目内容表示 ***
    txtKigou.Text = Aitem0(ips0, 0)
    txtK_mei.Text = Aitem0(ips0, 1)
    txtDsk.Text = Aitem0(ips0, 2)
    txtZumen1.Text = Aitem0(ips0, 3)
    txtZumen2.Text = Aitem0(ips0, 4)
    txtZumen3.Text = Aitem0(ips0, 5)
    txtGaiyou.Text = Aitem0(ips0, 6)
End Sub

Private Sub cmdKettei_Click()    '*** 決定 ***
    Dim Tempcode As String
    Dim i As Integer, j As Integer
'
    Select Case FLG_CorA
    Case 1                  '*** 変更データ取り込み ***
        Call GET_Chg_Data
'
    Case 2                  '*** 追加 ***
        Tempcode = Trim(txtKigou.Text)
        i = Len(Tempcode)
        If i <> 2 Then
            j = MsgBox("２文字の記号を記入してください。", vbCritical)
            Exit Sub
'
        End If
'
        For i = 1 To Anum0
            If Aitem0(i, 0) = Tempcode Then
                j = MsgBox("同じ記号がすでにあります。", vbExclamation)
                Exit Sub
'
            End If
        Next i
'
        i = MsgBox("はい なら " & Aitem0(ips0, 0) & " の前に挿入します。" & vbCrLf & " いいえ なら後に挿入します。", vbYesNoCancel)
        Select Case i
        Case vbYes
            'ips0=ips0
        Case vbNo
            ips0 = ips0 + 1
        Case vbCancel
            Exit Sub
'
        End Select
'
        Call Ins_Aitem0(Tempcode)    '*** データを１つずらす ***
'
        Aitem0(ips0, 0) = Tempcode
        Call GET_Chg_Data           '*** 変更データ取り込み ***
'
    End Select
'
    Me.MousePointer = vbHourglass
    DoEvents
'
    txtKigou.Enabled = True     '*** 表示を元に戻す ***
'
    txtKigou.TabStop = False
    txtK_mei.TabStop = False
    txtDsk.TabStop = False
    txtZumen1.TabStop = False
    txtZumen2.TabStop = False
    txtZumen3.TabStop = False
    txtGaiyou.TabStop = False
'
    Me.BackColor = FLG_Motonoiro
'
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    cmdSakujyo.Enabled = True
    cmdTuika.Enabled = True
    cmdkousin.Enabled = True
    cmdKettei.Enabled = False
'
    Call WRitem(DRVitem0, Aitem0(), Anum0, Adim0) '*** 項目データセーブ ***
    DoEvents
    Call RDitem(DRVitem0, Aitem0(), Anum0, Adim0) '*** 項目データ再読み込み ***
    DoEvents
    Call DSPlevel11  '*** 項目 内容表示 ***
    DoEvents
'
    Call DSPpointer(0)  '*** マウスポインターセット ***
'
    Me.MousePointer = vbDefault
    FLG_CorA = 0
    FLGitem_data_change = 1 '*** 変更フラグ設定(メイン画面の表示を変更する)***
End Sub

Private Sub cmdkousin_Click()    '*** 内容変更 ***
    FLG_CorA = 1
    Call Set_Gamen_CorA     '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)      '*** マウスポインターセット ***
'
    txtKigou.SetFocus
End Sub

Private Sub cmdSakujyo_Click()
    Dim i As Integer, j As Integer
'
    If Anum0 = 1 Then
        i = MsgBox("最後の１項目は削除できません。", vbCritical)
    Else
        i = MsgBox("項目 < " & Aitem0(ips0, 0) & " > を削除して良ろしいですか？", vbYesNo)
        If i = vbYes Then
            Me.MousePointer = vbHourglass
            DoEvents
'
            For i = ips0 To Anum0 - 1
                For j = 0 To Adim0
                    Aitem0(i, j) = Aitem0(i + 1, j)
                Next j
            Next i
            If ips0 = Anum0 Then
                ips0 = Anum0 - 1
            End If
            Anum0 = Anum0 - 1
            DoEvents
'
            Call WRitem(DRVitem0, Aitem0(), Anum0, Adim0) '*** 項目データセーブ ***
            DoEvents
            Call RDitem(DRVitem0, Aitem0(), Anum0, Adim0) '*** 項目データ再読み込み ***
            DoEvents
            Call DSPlevel11
            FLGitem_data_change = 1 '*** 変更フラグ設定(メイン画面の表示を変更する)***
'
            Me.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub cmdTuika_Click()    '*** 品種追加 ***
    FLG_CorA = 2
    Call Set_Gamen_CorA     '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)      '*** マウスポインターセット ***
'
    txtKigou.SetFocus
End Sub

Private Sub cmdDown_Click()
    Call H1down         '*** 一つ下へ ***
    Call DSPlevel11     '*** 項目内容表示 ***
End Sub

Private Sub H1down()
    If ips0 < Anum0 Then
        ips0 = ips0 + 1
    Else
        Beep
    End If
End Sub

Private Sub H1up()
    If ips0 > 1 Then
        ips0 = ips0 - 1
    Else
        Beep
    End If
End Sub

Private Sub Hdown()
    If ips0 + 10 <= Anum0 Then
        ips0 = ips0 + 10
    Else
        Beep
        ips0 = Anum0
    End If
End Sub

Private Sub Hup()
    If ips0 > 10 Then
        ips0 = ips0 - 10
    Else
        Beep
        ips0 = 1
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdUp_Click()
    Call H1up           '*** 一つ上へ ***
    Call DSPlevel11     '*** 項目内容表示 ***
End Sub

Private Sub Ins_Aitem0(Tempcode As String)  '*** データを１つずらす ***
    Dim i As Integer, j As Integer
    Dim DirStr As String, DRVindexDIR As String
'
    For i = Anum0 To ips0 Step -1
        For j = 0 To Adim0
            Aitem0(i + 1, j) = Aitem0(i, j)
        Next j
    Next i
    Anum0 = Anum0 + 1
'
    DRVindexT = Xcont0(2) & "\" & Tempcode & "\" & Tempcode & "INDEX.COD"
    DRVindexDIR = Xcont0(2) & "\" & Tempcode
    DirStr = Dir(DRVindexT)
    If DirStr = "" Then
        BnumT = 1
        BdimT = cBdim0
        MkDir DRVindexDIR           '*** DIR作成 ***
        ReDim BindexT(BnumT, BdimT) '*** 仮データ作成 ***
        BindexT(1, 0) = "仮"
        Call WRindex(DRVindexT, BindexT(), BnumT, BdimT)    '*** コード表仮作成 ***
    End If
End Sub

Private Sub GET_Chg_Data()      '*** 変更データ取り込み ***
    Aitem0(ips0, 1) = Trim(txtK_mei.Text)
    Aitem0(ips0, 2) = "*" 'Trim(txtDsk.Text)
    Aitem0(ips0, 3) = Trim(txtZumen1.Text)
    Aitem0(ips0, 4) = Trim(txtZumen2.Text)
    If Aitem0(ips0, 4) = "" Then
        Aitem0(ips0, 4) = "*"
    End If
    Aitem0(ips0, 5) = Trim(txtZumen3.Text)
    If Aitem0(ips0, 5) = "" Then
        Aitem0(ips0, 5) = "*"
    End If
    Aitem0(ips0, 6) = Trim(txtGaiyou.Text)
End Sub

Private Sub Set_Gamen_CorA()    '*** 追加変更画面設定 ***
    txtKigou.TabStop = True
    txtK_mei.TabStop = True
'   txtDsk.TabStop = True
    txtZumen1.TabStop = True
    txtZumen2.TabStop = True
    txtZumen3.TabStop = True
    txtGaiyou.TabStop = True
    Me.BackColor = &H808080
'
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If FLG_CorA = 1 Then
        cmdkousin.Enabled = True
        cmdTuika.Enabled = False
    Else
        cmdkousin.Enabled = False
        cmdTuika.Enabled = True
    End If
    cmdSakujyo.Enabled = False
    cmdKettei.Enabled = True
End Sub

Private Sub DSPpointer(X As Integer)
    If X = 1 Then
        txtKigou.MousePointer = vbArrow
        txtK_mei.MousePointer = vbArrow
'       txtDsk.MousePointer = vbArrow
        txtZumen1.MousePointer = vbArrow
        txtZumen2.MousePointer = vbArrow
        txtZumen3.MousePointer = vbArrow
        txtGaiyou.MousePointer = vbArrow
    Else
        txtKigou.MousePointer = vbDefault
        txtK_mei.MousePointer = vbDefault
'       txtDsk.MousePointer = vbDefault
        txtZumen1.MousePointer = vbDefault
        txtZumen2.MousePointer = vbDefault
        txtZumen3.MousePointer = vbDefault
        txtGaiyou.MousePointer = vbDefault
    End If
End Sub

Private Sub txtKigou_LostFocus()
    Dim Tempcode As String
    Dim i As Integer, j As Integer
'
    If FLG_CorA <> 2 Then       '*** 部品追加モード時のみ有効 ***
        Exit Sub
    End If
'
    Tempcode = Trim(txtKigou.Text)
    i = Len(Tempcode)
    If i <> 2 Then
        j = MsgBox("正しい項目記号を記入してください。", vbCritical)
'
        txtKigou.SetFocus
        Exit Sub
'
    End If
'
    For i = 1 To Anum0
        If Aitem0(i, 0) = Tempcode Then
            j = MsgBox("同じ項目記号がすでにあります。", vbExclamation)
            Exit Sub
'
        End If
    Next i
End Sub

Private Sub DSPgamenBuhin()
    txtKigou.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtKigou.Top = 360
    txtKigou.FontSize = 10 * HyoujiBairitu
    txtKigou.Width = 375 * HyoujiBairitu
    txtKigou.Height = 285 * HyoujiBairitu
'
    lblKigou.Left = 720
    lblKigou.Top = 360
    lblKigou.FontSize = 10 * HyoujiBairitu
    lblKigou.Width = 1575 * HyoujiBairitu
    lblKigou.Height = txtKigou.Height
'
    txtK_mei.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtK_mei.Top = 360 + (720 - 360) * HyoujiBairitu
    txtK_mei.FontSize = 10 * HyoujiBairitu
    txtK_mei.Width = 1455 * HyoujiBairitu
    txtK_mei.Height = 285 * HyoujiBairitu
'
    lblK_mei.Left = 720
    lblK_mei.Top = 360 + (720 - 360) * HyoujiBairitu
    lblK_mei.FontSize = 10 * HyoujiBairitu
    lblK_mei.Width = 1575 * HyoujiBairitu
    lblK_mei.Height = txtK_mei.Height
'
    txtDsk.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtDsk.Top = 360 + (1080 - 360) * HyoujiBairitu
    txtDsk.FontSize = 10 * HyoujiBairitu
    txtDsk.Width = 255 * HyoujiBairitu
    txtDsk.Height = 285 * HyoujiBairitu
'
    lblDsk.Left = 720
    lblDsk.Top = 360 + (1080 - 360) * HyoujiBairitu
    lblDsk.FontSize = 9 * HyoujiBairitu
    lblDsk.Width = 1575 * HyoujiBairitu
    lblDsk.Height = txtDsk.Height
'
    txtZumen1.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtZumen1.Top = 360 + (1440 - 360) * HyoujiBairitu
    txtZumen1.FontSize = 10 * HyoujiBairitu
    txtZumen1.Width = 495 * HyoujiBairitu
    txtZumen1.Height = 285 * HyoujiBairitu
'
    lblZumen1.Left = 720
    lblZumen1.Top = 360 + (1440 - 360) * HyoujiBairitu
    lblZumen1.FontSize = 10 * HyoujiBairitu
    lblZumen1.Width = 1575 * HyoujiBairitu
    lblZumen1.Height = txtZumen1.Height
'
    txtZumen2.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtZumen2.Top = 360 + (1800 - 360) * HyoujiBairitu
    txtZumen2.FontSize = 10 * HyoujiBairitu
    txtZumen2.Width = 495 * HyoujiBairitu
    txtZumen2.Height = 285 * HyoujiBairitu
'
    lblZumen2.Left = 720
    lblZumen2.Top = 360 + (1800 - 360) * HyoujiBairitu
    lblZumen2.FontSize = 10 * HyoujiBairitu
    lblZumen2.Width = 1575 * HyoujiBairitu
    lblZumen2.Height = txtZumen2.Height
'
    txtZumen3.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtZumen3.Top = 360 + (2160 - 360) * HyoujiBairitu
    txtZumen3.FontSize = 10 * HyoujiBairitu
    txtZumen3.Width = 495 * HyoujiBairitu
    txtZumen3.Height = 285 * HyoujiBairitu
'
    lblZumen3.Left = 720
    lblZumen3.Top = 360 + (2160 - 360) * HyoujiBairitu
    lblZumen3.FontSize = 10 * HyoujiBairitu
    lblZumen3.Width = 1575 * HyoujiBairitu
    lblZumen3.Height = txtZumen3.Height
'
    txtGaiyou.Left = 720 + (2400 - 720) * HyoujiBairitu
    txtGaiyou.Top = 360 + (2520 - 360) * HyoujiBairitu
    txtGaiyou.FontSize = 10 * HyoujiBairitu
    txtGaiyou.Width = 4695 * HyoujiBairitu
    txtGaiyou.Height = 285 * HyoujiBairitu
'
    lblGaiyou.Left = 720
    lblGaiyou.Top = 360 + (2520 - 360) * HyoujiBairitu
    lblGaiyou.FontSize = 10 * HyoujiBairitu
    lblGaiyou.Width = 1575 * HyoujiBairitu
    lblGaiyou.Height = txtGaiyou.Height
'
    cmdSakujyo.Left = 720 + (840 - 720) * HyoujiBairitu
    cmdSakujyo.Top = 360 + (3120 - 360) * HyoujiBairitu
    cmdSakujyo.FontSize = 9 * HyoujiBairitu
    cmdSakujyo.Width = 1215 * HyoujiBairitu
    cmdSakujyo.Height = 495 * HyoujiBairitu
'
    cmdKettei.Left = 720 + (2280 - 720) * HyoujiBairitu
    cmdKettei.Top = 360 + (3120 - 360) * HyoujiBairitu
    cmdKettei.FontSize = 9 * HyoujiBairitu
    cmdKettei.Width = 1215 * HyoujiBairitu
    cmdKettei.Height = 495 * HyoujiBairitu
'
    cmdTuika.Left = 720 + (840 - 720) * HyoujiBairitu
    cmdTuika.Top = 360 + (3720 - 360) * HyoujiBairitu
    cmdTuika.FontSize = 9 * HyoujiBairitu
    cmdTuika.Width = 1215 * HyoujiBairitu
    cmdTuika.Height = 495 * HyoujiBairitu
'
    cmdkousin.Left = 720 + (2280 - 720) * HyoujiBairitu
    cmdkousin.Top = 360 + (3720 - 360) * HyoujiBairitu
    cmdkousin.FontSize = 9 * HyoujiBairitu
    cmdkousin.Width = 1215 * HyoujiBairitu
    cmdkousin.Height = 495 * HyoujiBairitu
'
    cmdUp.Left = 720 + (3720 - 720) * HyoujiBairitu
    cmdUp.Top = 360 + (3720 - 360) * HyoujiBairitu
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Width = 855 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdDown.Left = 720 + (4800 - 720) * HyoujiBairitu
    cmdDown.Top = 360 + (3720 - 360) * HyoujiBairitu
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Width = 855 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 720 + (5880 - 720) * HyoujiBairitu
    cmdQuit.Top = 360 + (3720 - 360) * HyoujiBairitu
    cmdQuit.FontSize = 9 * HyoujiBairitu
    cmdQuit.Width = 1095 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
End Sub
