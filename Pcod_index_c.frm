VERSION 5.00
Begin VB.Form Pcod_index_c 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ＥＥＯＳ 電気部品コード 単価表＜品種一覧 帳票形式＞"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10215
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
   Icon            =   "Pcod_index_c.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5895
   ScaleWidth      =   10215
   Begin VB.TextBox txtHyoujyun 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7320
      MousePointer    =   1  '矢印
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "Text11"
      Top             =   3720
      Width           =   375
   End
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
      Left            =   1440
      TabIndex        =   35
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox cboMaker2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6120
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "cboMaker2"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ComboBox cboMaker1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "cboMaker1"
      Top             =   720
      Width           =   3495
   End
   Begin VB.ComboBox cboMaker0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "cboMaker0"
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdKettei 
      Caption         =   "決定保存(&G)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   29
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtKigou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text12"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   7560
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdTuika 
      Caption         =   "品種追加(&A)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   34
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdkousin 
      Caption         =   "内容変更(&C)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   33
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   6360
      TabIndex        =   31
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   5280
      TabIndex        =   30
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtHBikou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "Text11"
      Top             =   3720
      Width           =   4575
   End
   Begin VB.TextBox txtHyouki 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Text10"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtMaker2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6120
      MousePointer    =   1  '矢印
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "Text9"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txtFamName2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "Text8"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtMaker1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6120
      MousePointer    =   1  '矢印
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Text7"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox txtFamName1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Text6"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtMaker0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6120
      MousePointer    =   1  '矢印
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txtFamName0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtEName 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox txtJName 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox txtMCode 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblHyoujyun 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "標準部品"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6360
      TabIndex        =   27
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblKigou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "項目名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblHBikou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種備考"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   25
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblHyouki 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "図面表記"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   23
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblMaker2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5520
      TabIndex        =   20
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblFamName2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "型名 ３"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblMaker1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5520
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblFamName1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "型名 ２"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblMaker0 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5520
      TabIndex        =   10
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblFamName0 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "型名 １"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblEName 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種名 英"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblJName 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種名 和"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblMCode 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "親コード"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Pcod_index_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*******************
'*  品種 細目一覧   *
'*******************
'
Option Explicit
'
    Dim HeadTitle As String
    Dim FLG_CorA As Integer
    Dim FLG_Motonoiro As Long

Private Sub DSPlevel21()    '*** 品種 細目表示 ***
    Dim makercode As String
    Dim makername As String
    Dim X As Integer
'
    txtMCode.Text = "L" & Bindex0(jps0, 0)
    txtJName.Text = Bindex0(jps0, 1)
    txtEName.Text = Bindex0(jps0, 2)
    txtFamName0.Text = Bindex0(jps0, 3) & "xxx" & Bindex0(jps0, 4)
'
    If Bindex0(jps0, 5) = "998" Then X = 8 Else X = 5
    makercode = Bindex0(jps0, X)
    makername = makercode
    Call Makerget1(makername)   '***ﾒｰｶｰ名取得 ***
    txtMaker0.Text = makercode & " " & makername
'
    If Bindex0(jps0, 5) = "998" Then
        lblFamName1.Enabled = True
        txtFamName1.Enabled = True
        txtFamName1.Text = Bindex0(jps0, 11) & "xxx" & Bindex0(jps0, 12)
        If Bindex0(jps0, 9) = "000" Then
            makercode = "***"
            makername = "*"
        Else
            makercode = Bindex0(jps0, 9)
            makername = makercode
            Call Makerget1(makername)   '***ﾒｰｶｰ名取得 ***
        End If
        lblMaker1.Enabled = True
        txtMaker1.Enabled = True
        txtMaker1.Text = makercode & " " & makername
'
        lblFamName2.Enabled = True
        txtFamName2.Enabled = True
        txtFamName2.Text = Bindex0(jps0, 13) & "xxx" & Bindex0(jps0, 14)
'
        If Bindex0(jps0, 10) = "000" Then
            makercode = "***"
            makername = "*"
        Else
            makercode = Bindex0(jps0, 10)
            makername = makercode
            Call Makerget1(makername)   '***ﾒｰｶｰ名取得 ***
        End If
        lblMaker2.Enabled = True
        txtMaker2.Enabled = True
        txtMaker2.Text = makercode & " " & makername
    Else
        lblFamName1.Enabled = False
        txtFamName1.Text = "*xxx*"
        txtFamName1.Enabled = False
        lblMaker1.Enabled = False
        txtMaker1.Text = "*"
        txtMaker1.Enabled = False
        lblFamName2.Enabled = False
        txtFamName2.Text = "*xxx*"
        txtFamName2.Enabled = False
        lblMaker2.Enabled = False
        txtMaker2.Text = "*"
        txtMaker2.Enabled = False
    End If
'
    txtHyouki.Text = Bindex0(jps0, 6)
    txtHBikou.Text = Bindex0(jps0, 7)
        If Bindex0(jps0, 15) = "" Then Bindex0(jps0, 15) = "*"
    txtHyoujyun.Text = Bindex0(jps0, 15)
End Sub

Private Sub cboMaker0_Click()
    Dim i As Integer
'
    i = cboMaker0.ListIndex
    If Maker(i, 0) = "998" Then
        lblFamName1.Enabled = True
        txtFamName1.Enabled = True
        lblMaker1.Enabled = True
        cboMaker1.Enabled = True
        lblFamName2.Enabled = True
        txtFamName2.Enabled = True
        cboMaker2.Enabled = True
'
        cboMaker0.ListIndex = 0
    End If
End Sub

Private Sub cmdDown_Click()
    Call H1down         '*** 一つ下へ ***
    Call DSPlevel21    '*** 項目内容表示 ***
'
    txtMCode.SetFocus
End Sub

Private Sub cmdkousin_Click()
    FLG_CorA = 1
    Call Set_Gamen_CorA     '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)      '*** マウスポインターセット ***
'
    txtJName.SetFocus
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSakujyo_Click()
    Dim i As Integer, j As Integer
'
    If Bnum0 = 1 Then
        i = MsgBox("最後の１品種は削除できません。", vbCritical)
    Else
        i = MsgBox("品種 < L" & Bindex0(jps0, 0) & " > を削除して良ろしいですか？", vbYesNo)
        If i = vbYes Then
            Me.MousePointer = vbHourglass
            DoEvents
'
            For i = jps0 To Bnum0 - 1
                For j = 0 To Bdim0
                    Bindex0(i, j) = Bindex0(i + 1, j)
                Next j
            Next i
            If jps0 = Bnum0 Then
                jps0 = Bnum0 - 1
            End If
            Bnum0 = Bnum0 - 1
            DoEvents
'
            Call WRindex(DRVindex0, Bindex0(), Bnum0, Bdim0)    '*** 品種データセーブ ***
            DoEvents
            Call RDindex(DRVindex0, Bindex0(), Bnum0, Bdim0)    '*** 品種データ再読み込み ***
            DoEvents
            Call DSPlevel21
            FLGindex_data_change = 1    '*** 変更フラグ設定(メイン画面の表示を変更する)***
'
            Me.MousePointer = vbDefault
        End If
    End If
'
    txtMCode.SetFocus
End Sub

Private Sub cmdTuika_Click()
    FLG_CorA = 2
    Call Set_Gamen_CorA     '*** 追加変更画面設定 ***
'
    Call DSPpointer(0)      '*** マウスポインターセット ***
'
    txtMCode.SetFocus
End Sub

Private Sub cmdUp_Click()
    Call H1up           '*** 一つ上へ ***
    Call DSPlevel21    '*** 項目内容表示 ***
'
    txtMCode.SetFocus
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
        Tempcode = Mid(Trim(txtMCode.Text), 2)
        i = Len(Tempcode)
        If i <> 4 Then
            j = MsgBox("正しいコード番号を記入してください。", vbCritical)
            Exit Sub
'
        End If
'
        For i = 1 To Bnum0
            If Bindex0(i, 0) = Tempcode Then
                j = MsgBox("同じコード番号がすでにあります。", vbExclamation)
                Exit Sub
'
            End If
        Next i
'
        i = MsgBox("はい なら L" & Bindex0(jps0, 0) & " の前に挿入します。" & vbCrLf & " いいえ なら後に挿入します。", vbYesNoCancel)
        Select Case i
        Case vbYes
            'jps0=jps0
        Case vbNo
            jps0 = jps0 + 1
        Case vbCancel
            Exit Sub
'
        End Select
'
        Call Ins_Bindex(Tempcode)   '*** データを１つずらす ***
'
        Bindex0(jps0, 0) = Tempcode
        Call GET_Chg_Data     '*** 変更データ取り込み ***
'
    End Select
'
    Me.MousePointer = vbHourglass
    DoEvents
'
    txtMCode.TabStop = False    '*** 表示を元に戻す ***
    txtJName.TabStop = False
    txtEName.TabStop = False
    txtFamName0.TabStop = False
    txtFamName1.TabStop = False
    txtFamName2.TabStop = False
    txtHyouki.TabStop = False
    txtHBikou.TabStop = False
    txtHyoujyun.TabStop = False
    Me.BackColor = FLG_Motonoiro
'
    cmdDown.Enabled = True
    cmdUp.Enabled = True
    cmdkousin.Enabled = True
    cmdTuika.Enabled = True
    cmdSakujyo.Enabled = True
    cmdKettei.Enabled = False
'
    cboMaker0.Visible = False
    cboMaker1.Visible = False
    cboMaker2.Visible = False
    txtMaker0.Visible = True
    txtMaker1.Visible = True
    txtMaker2.Visible = True
'
    Call WRindex(DRVindex0, Bindex0(), Bnum0, Bdim0)    '*** INDEXデータセーブ ***
    DoEvents
    Call RDindex(DRVindex0, Bindex0(), Bnum0, Bdim0)    '*** INDEXデータ再読み込み ***
    DoEvents
    Call DSPlevel21  '*** 品種 細目表示 ***
    DoEvents
'
    Call DSPpointer(0)  '*** マウスポインターセット ***
'
    Me.MousePointer = vbDefault
    FLG_CorA = 0
    FLGindex_data_change = 1    '*** 変更フラグ設定(メイン画面の表示を変更する)***
'
    txtMCode.SetFocus
End Sub

Private Sub Ins_Bindex(Tempcode As String)  '*** データを１つずらす ***
    Dim i As Integer
    Dim j As Integer
    Dim DirStr As String
'
    For i = Bnum0 To jps0 Step -1
        For j = 0 To Bdim0
            Bindex0(i + 1, j) = Bindex0(i, j)
        Next j
    Next i
    Bnum0 = Bnum0 + 1
'
    If Aitem0(ips0, 0) = "IC" Then
        DRVmainT = Xcont0(2) & "\IC\IC" & Left$(Tempcode, 1) _
        & "\L" & Tempcode & ".COD"
    Else
        DRVmainT = Xcont0(2) & "\" & Aitem0(ips0, 0) _
        & "\L" & Tempcode & ".COD"
    End If
    DirStr = Dir(DRVmainT)
    If DirStr = "" Then
        CnumT = 1
        CdimT = cCdim0
        ReDim CmainT(CnumT, CdimT)                      '*** 仮データ作成 ***
        CmainT(1, 0) = "仮"
        Call WRmain(DRVmainT, CmainT(), CnumT, CdimT)   '*** コード表仮作成 ***
    End If
End Sub

Private Sub Form_Initialize()
    HeadTitle = "部品ｺｰﾄﾞ <品種 帳票形式>"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If FLG_CorA = 1 Or FLG_CorA = 2 Then Exit Sub   '*** 変更追加画面の時はキーが効かない ***
'
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        Hup             '*** 上へ ***
        DSPlevel21      '*** 品種 細目表示 ***
'
    Case vbKeyUp        '*** ↑ ***
        H1up            '*** 一つ上へ ***
        DSPlevel21      '*** 品種 細目表示 ***
'
    Case vbKeyPageUp    '*** Roll Down
        Hdown           '*** 下へ ***
        DSPlevel21      '*** 品種 細目表示 ***
'
    Case vbKeyDown      '*** ↓ ***
        H1down          '*** 一つ下へ ***
        DSPlevel21      '*** 品種 細目表示 ***
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
                                '*** フォームの表示位置の設定
    Width = 600 + (10305 - 1200) * HyoujiBairitu + 600
    Height = 360 + (6270 - 600) * HyoujiBairitu + 360
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
    txtKigou.Text = Aitem0(ips0, 0)
'
    Call DSPlevel21     '*** 品種 細目一覧 ***
'
    cboMaker0.Visible = False
    cboMaker1.Visible = False
    cboMaker2.Visible = False
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

Private Sub GET_Chg_Data()      '*** 変更データ取り込み ***
    Dim i As Integer
    Dim TempAno As Integer
    Dim TempBno As Integer
    Dim TempCno As Integer
'
    Bindex0(jps0, 1) = Trim(txtJName.Text)
    Bindex0(jps0, 2) = Trim(txtEName.Text)
'
    i = InStr(1, Trim(txtFamName0.Text), "xxx")
    If i = 0 Then
        Bindex0(jps0, 3) = Trim(txtFamName0.Text)
        Bindex0(jps0, 4) = "*"
    Else
        Bindex0(jps0, 3) = Left(Trim(txtFamName0.Text), i - 1)
        Bindex0(jps0, 4) = Mid(Trim(txtFamName0.Text), i + 3)
    End If
'
    TempAno = cboMaker0.ListIndex
    TempBno = cboMaker1.ListIndex
    TempCno = cboMaker2.ListIndex
'
    If TempBno = 0 And TempCno = 0 Then
        Bindex0(jps0, 5) = Maker(TempAno, 0)
        Bindex0(jps0, 8) = "*"
        Bindex0(jps0, 9) = "*"
        Bindex0(jps0, 10) = "*"
        Bindex0(jps0, 11) = "*"
        Bindex0(jps0, 12) = "*"
        Bindex0(jps0, 13) = "*"
        Bindex0(jps0, 14) = "*"
    Else
        Bindex0(jps0, 5) = "998"
        Bindex0(jps0, 8) = Maker(TempAno, 0)
'
        Bindex0(jps0, 9) = Maker(TempBno, 0)
'
        i = InStr(1, Trim(txtFamName1.Text), "xxx")
        If i = 0 Then
            Bindex0(jps0, 11) = Trim(txtFamName1.Text)
            Bindex0(jps0, 12) = "*"
        Else
            Bindex0(jps0, 11) = Left(Trim(txtFamName1.Text), i - 1)
            Bindex0(jps0, 12) = Mid(Trim(txtFamName1.Text), i + 3)
        End If
'
        If TempCno = 0 Then
            Bindex0(jps0, 10) = "*"
            Bindex0(jps0, 13) = "*"
            Bindex0(jps0, 14) = "*"
        Else
            Bindex0(jps0, 10) = Maker(TempCno, 0)
'
            i = InStr(1, Trim(txtFamName2.Text), "xxx")
            If i = 0 Then
                Bindex0(jps0, 13) = Trim(txtFamName2.Text)
                Bindex0(jps0, 14) = "*"
            Else
                Bindex0(jps0, 13) = Left(Trim(txtFamName2.Text), i - 1)
                Bindex0(jps0, 14) = Mid(Trim(txtFamName2.Text), i + 3)
            End If
        End If
    End If
'
    Bindex0(jps0, 6) = Trim(txtHyouki.Text)
    Bindex0(jps0, 7) = Trim(txtHBikou.Text)
    Bindex0(jps0, 15) = Trim(txtHyoujyun.Text)
End Sub

Private Sub Set_Gamen_CorA()    '*** 追加変更画面設定 ***
    Dim i As Integer
    Dim Mdata As String
'
    If FLG_CorA = 1 Then
        txtMCode.Locked = True
    End If
    txtJName.TabStop = True
    txtEName.TabStop = True
    txtFamName0.TabStop = True
    txtFamName1.TabStop = True
    txtFamName2.TabStop = True
    txtHyouki.TabStop = True
    txtHBikou.TabStop = True
    txtHyoujyun.TabStop = True
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
'
    txtMaker0.Visible = False
    txtMaker1.Visible = False
    txtMaker2.Visible = False
    cboMaker0.Visible = True
    cboMaker0.Clear
    cboMaker1.Visible = True
    cboMaker1.Clear
    cboMaker2.Visible = True
    cboMaker2.Clear
'
    cboMaker0.AddItem "*"
    cboMaker1.AddItem "*"
    cboMaker2.AddItem "*"
    For i = 1 To Maknum0
        cboMaker0.AddItem Maker(i, 0) & " " & Maker(i, 2)
        cboMaker1.AddItem Maker(i, 0) & " " & Maker(i, 2)
        cboMaker2.AddItem Maker(i, 0) & " " & Maker(i, 2)
    Next i
'
    If Bindex0(jps0, 5) = "998" And FLG_CorA <> 2 Then
        cboMaker1.Enabled = True
        cboMaker2.Enabled = True
'
        For i = 1 To Maknum0
            If Bindex0(jps0, 8) = Maker(i, 0) Then
                cboMaker0.ListIndex = i
                Exit For
            End If
            cboMaker0.ListIndex = 0
        Next i
'
        For i = 1 To Maknum0
            If Bindex0(jps0, 9) = Maker(i, 0) Then
                cboMaker1.ListIndex = i
                Exit For
            End If
            cboMaker1.ListIndex = 0
        Next i
'
        For i = 1 To Maknum0
            If Bindex0(jps0, 10) = Maker(i, 0) Then
                cboMaker2.ListIndex = i
                Exit For
            End If
            cboMaker2.ListIndex = 0
        Next i
    Else
        cboMaker1.Enabled = False
        cboMaker2.Enabled = False
        For i = 1 To Maknum0
            If Bindex0(jps0, 5) = Maker(i, 0) Then
                cboMaker0.ListIndex = i
                Exit For
            End If
            cboMaker0.ListIndex = 0
        Next i
'
        cboMaker1.ListIndex = 0
        cboMaker2.ListIndex = 0
    End If
End Sub

Private Sub DSPpointer(X As Integer)
    If X = 1 Then
        txtMCode.MousePointer = 1
        txtMCode.MousePointer = 1
        txtJName.MousePointer = 1
        txtEName.MousePointer = 1
        txtFamName0.MousePointer = 1
        txtMaker0.MousePointer = 1
        txtFamName1.MousePointer = 1
        txtMaker1.MousePointer = 1
        txtFamName2.MousePointer = 1
        txtMaker2.MousePointer = 1
        txtJName.MousePointer = 1
        txtEName.MousePointer = 1
    Else
        txtMCode.MousePointer = 0
        txtJName.MousePointer = 0
        txtEName.MousePointer = 0
        txtFamName0.MousePointer = 0
        txtFamName1.MousePointer = 0
        txtFamName2.MousePointer = 0
        txtJName.MousePointer = 0
        txtEName.MousePointer = 0
    End If
End Sub

Private Sub Hup()
    If jps0 > 10 Then
        jps0 = jps0 - 10
    Else
        Beep
        jps0 = 1
    End If
End Sub

Private Sub Hdown()
    If jps0 + 10 <= Bnum0 Then
        jps0 = jps0 + 10
    Else
        Beep
        jps0 = Bnum0
    End If
End Sub

Private Sub H1down()
    If jps0 < Bnum0 Then
        jps0 = jps0 + 1
    Else
        Beep
    End If
End Sub

Private Sub H1up()
    If jps0 > 1 Then
        jps0 = jps0 - 1
    Else
        Beep
    End If
End Sub

Private Sub DSPgamenBuhin()
    txtKigou.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtKigou.Top = 360
    txtKigou.FontSize = 10 * HyoujiBairitu
    txtKigou.Width = 375 * HyoujiBairitu
    txtKigou.Height = 285 * HyoujiBairitu
'
    lblKigou.Left = 600
    lblKigou.Top = 360
    lblKigou.FontSize = 10 * HyoujiBairitu
    lblKigou.Width = 975 * HyoujiBairitu
    lblKigou.Height = txtKigou.Height
'
    txtMCode.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtMCode.Top = 360 + (720 - 360) * HyoujiBairitu
    txtMCode.FontSize = 10 * HyoujiBairitu
    txtMCode.Width = 735 * HyoujiBairitu
    txtMCode.Height = 285 * HyoujiBairitu
'
    lblMCode.Left = 600
    lblMCode.Top = 360 + (720 - 360) * HyoujiBairitu
    lblMCode.FontSize = 10 * HyoujiBairitu
    lblMCode.Width = 975 * HyoujiBairitu
    lblMCode.Height = txtMCode.Height
'
    txtJName.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtJName.Top = 360 + (1200 - 360) * HyoujiBairitu
    txtJName.FontSize = 10 * HyoujiBairitu
    txtJName.Width = 4575 * HyoujiBairitu
    txtJName.Height = 285 * HyoujiBairitu
'
    lblJName.Left = 600
    lblJName.Top = 360 + (1200 - 360) * HyoujiBairitu
    lblJName.FontSize = 10 * HyoujiBairitu
    lblJName.Width = 975 * HyoujiBairitu
    lblJName.Height = txtJName.Height
'
    txtEName.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtEName.Top = 360 + (1560 - 360) * HyoujiBairitu
    txtEName.FontSize = 10 * HyoujiBairitu
    txtEName.Width = 4575 * HyoujiBairitu
    txtEName.Height = 285 * HyoujiBairitu
'
    lblEName.Left = 600
    lblEName.Top = 360 + (1560 - 360) * HyoujiBairitu
    lblEName.FontSize = 10 * HyoujiBairitu
    lblEName.Width = 975 * HyoujiBairitu
    lblEName.Height = txtEName.Height
'
    txtFamName0.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtFamName0.Top = 360 + (2040 - 360) * HyoujiBairitu
    txtFamName0.FontSize = 10 * HyoujiBairitu
    txtFamName0.Width = 3615 * HyoujiBairitu
    txtFamName0.Height = 285 * HyoujiBairitu
'
    lblFamName0.Left = 600
    lblFamName0.Top = 360 + (2040 - 360) * HyoujiBairitu
    lblFamName0.FontSize = 10 * HyoujiBairitu
    lblFamName0.Width = 975 * HyoujiBairitu
    lblFamName0.Height = txtFamName0.Height
'
    txtFamName1.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtFamName1.Top = 360 + (2400 - 360) * HyoujiBairitu
    txtFamName1.FontSize = 10 * HyoujiBairitu
    txtFamName1.Width = 3615 * HyoujiBairitu
    txtFamName1.Height = 285 * HyoujiBairitu
'
    lblFamName1.Left = 600
    lblFamName1.Top = 360 + (2400 - 360) * HyoujiBairitu
    lblFamName1.FontSize = 10 * HyoujiBairitu
    lblFamName1.Width = 975 * HyoujiBairitu
    lblFamName1.Height = txtFamName1.Height
'
    txtFamName2.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtFamName2.Top = 360 + (2760 - 360) * HyoujiBairitu
    txtFamName2.FontSize = 10 * HyoujiBairitu
    txtFamName2.Width = 3615 * HyoujiBairitu
    txtFamName2.Height = 285 * HyoujiBairitu
'
    lblFamName2.Left = 600
    lblFamName2.Top = 360 + (2760 - 360) * HyoujiBairitu
    lblFamName2.FontSize = 10 * HyoujiBairitu
    lblFamName2.Width = 975 * HyoujiBairitu
    lblFamName2.Height = txtFamName2.Height
'
    txtMaker0.Left = 600 + (6135 - 600) * HyoujiBairitu
    txtMaker0.Top = 360 + (2040 - 360) * HyoujiBairitu
    txtMaker0.FontSize = 10 * HyoujiBairitu
    txtMaker0.Width = 3495 * HyoujiBairitu
    txtMaker0.Height = 285 * HyoujiBairitu
'
    lblMaker0.Left = 600 + (5520 - 600) * HyoujiBairitu
    lblMaker0.Top = 360 + (2040 - 360) * HyoujiBairitu
    lblMaker0.FontSize = 10 * HyoujiBairitu
    lblMaker0.Width = 615 * HyoujiBairitu
    lblMaker0.Height = txtMaker0.Height
'
    cboMaker0.Left = 600 + (6135 - 600) * HyoujiBairitu
    cboMaker0.Top = 360 + (2040 - 360) * HyoujiBairitu
    cboMaker0.FontSize = 10 * HyoujiBairitu
    cboMaker0.Width = 3495 * HyoujiBairitu
'   cboMaker0.Height = txtMaker0.Height
'
    txtMaker1.Left = 600 + (6135 - 600) * HyoujiBairitu
    txtMaker1.Top = 360 + (2400 - 360) * HyoujiBairitu
    txtMaker1.FontSize = 10 * HyoujiBairitu
    txtMaker1.Width = 3495 * HyoujiBairitu
    txtMaker1.Height = 285 * HyoujiBairitu
'
    lblMaker1.Left = 600 + (5520 - 600) * HyoujiBairitu
    lblMaker1.Top = 360 + (2400 - 360) * HyoujiBairitu
    lblMaker1.FontSize = 10 * HyoujiBairitu
    lblMaker1.Width = 615 * HyoujiBairitu
    lblMaker1.Height = txtMaker1.Height
'
    cboMaker1.Left = 600 + (6135 - 600) * HyoujiBairitu
    cboMaker1.Top = 360 + (2400 - 360) * HyoujiBairitu
    cboMaker1.FontSize = 10 * HyoujiBairitu
    cboMaker1.Width = 3495 * HyoujiBairitu
'   cboMaker1.Height = txtMaker1.Height
'
    txtMaker2.Left = 600 + (6135 - 600) * HyoujiBairitu
    txtMaker2.Top = 360 + (2760 - 360) * HyoujiBairitu
    txtMaker2.FontSize = 10 * HyoujiBairitu
    txtMaker2.Width = 3495 * HyoujiBairitu
    txtMaker2.Height = 285 * HyoujiBairitu
'
    lblMaker2.Left = 600 + (5520 - 600) * HyoujiBairitu
    lblMaker2.Top = 360 + (2760 - 360) * HyoujiBairitu
    lblMaker2.FontSize = 10 * HyoujiBairitu
    lblMaker2.Width = 615 * HyoujiBairitu
    lblMaker2.Height = txtMaker2.Height
'
    cboMaker2.Left = 600 + (6135 - 600) * HyoujiBairitu
    cboMaker2.Top = 360 + (2760 - 360) * HyoujiBairitu
    cboMaker2.FontSize = 10 * HyoujiBairitu
    cboMaker2.Width = 3495 * HyoujiBairitu
'   cboMaker2.Height = txtMaker2.Height
'
    txtHyouki.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtHyouki.Top = 360 + (3240 - 360) * HyoujiBairitu
    txtHyouki.FontSize = 10 * HyoujiBairitu
    txtHyouki.Width = 1815 * HyoujiBairitu
    txtHyouki.Height = 285 * HyoujiBairitu
'
    lblHyouki.Left = 600
    lblHyouki.Top = 360 + (3240 - 360) * HyoujiBairitu
    lblHyouki.FontSize = 10 * HyoujiBairitu
    lblHyouki.Width = 975 * HyoujiBairitu
    lblHyouki.Height = txtHyouki.Height
'
    txtHBikou.Left = 600 + (1575 - 600) * HyoujiBairitu
    txtHBikou.Top = 360 + (3720 - 360) * HyoujiBairitu
    txtHBikou.FontSize = 10 * HyoujiBairitu
    txtHBikou.Width = 4575 * HyoujiBairitu
    txtHBikou.Height = 285 * HyoujiBairitu
'
    lblHBikou.Left = 600
    lblHBikou.Top = 360 + (3720 - 360) * HyoujiBairitu
    lblHBikou.FontSize = 10 * HyoujiBairitu
    lblHBikou.Width = 975 * HyoujiBairitu
    lblHBikou.Height = txtHBikou.Height
'
    txtHyoujyun.Left = 600 + (7335 - 600) * HyoujiBairitu
    txtHyoujyun.Top = 360 + (3720 - 360) * HyoujiBairitu
    txtHyoujyun.FontSize = 10 * HyoujiBairitu
    txtHyoujyun.Width = 375 * HyoujiBairitu
    txtHyoujyun.Height = 285 * HyoujiBairitu
'
    lblHyoujyun.Left = 600 + (6360 - 600) * HyoujiBairitu
    lblHyoujyun.Top = 360 + (3720 - 360) * HyoujiBairitu
    lblHyoujyun.FontSize = 10 * HyoujiBairitu
    lblHyoujyun.Width = 975 * HyoujiBairitu
    lblHyoujyun.Height = txtHyoujyun.Height
'
    cmdSakujyo.Left = 360 + (1440 - 360) * HyoujiBairitu
    cmdSakujyo.Top = 360 + (4320 - 360) * HyoujiBairitu
    cmdSakujyo.FontSize = 10 * HyoujiBairitu
    cmdSakujyo.Width = 1335 * HyoujiBairitu
    cmdSakujyo.Height = 495 * HyoujiBairitu
'
    cmdKettei.Left = 360 + (3120 - 360) * HyoujiBairitu
    cmdKettei.Top = 360 + (4320 - 360) * HyoujiBairitu
    cmdKettei.FontSize = 9 * HyoujiBairitu
    cmdKettei.Width = 1335 * HyoujiBairitu
    cmdKettei.Height = 495 * HyoujiBairitu
'
    cmdTuika.Left = 360 + (1440 - 360) * HyoujiBairitu
    cmdTuika.Top = 360 + (5040 - 360) * HyoujiBairitu
    cmdTuika.FontSize = 10 * HyoujiBairitu
    cmdTuika.Width = 1335 * HyoujiBairitu
    cmdTuika.Height = 495 * HyoujiBairitu
'
    cmdkousin.Left = 360 + (3120 - 360) * HyoujiBairitu
    cmdkousin.Top = 360 + (5040 - 360) * HyoujiBairitu
    cmdkousin.FontSize = 10 * HyoujiBairitu
    cmdkousin.Width = 1335 * HyoujiBairitu
    cmdkousin.Height = 495 * HyoujiBairitu
'
    cmdUp.Left = 360 + (5280 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (5040 - 360) * HyoujiBairitu
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Width = 855 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdDown.Left = 360 + (6360 - 360) * HyoujiBairitu
    cmdDown.Top = 360 + (5040 - 360) * HyoujiBairitu
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Width = 855 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 360 + (7560 - 360) * HyoujiBairitu
    cmdQuit.Top = 360 + (5040 - 360) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1215 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
End Sub

