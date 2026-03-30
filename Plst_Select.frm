VERSION 5.00
Begin VB.Form Plst_Select 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "品種品目選択"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Plst_Select.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox txtBikou 
      BackColor       =   &H00008000&
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
      Height          =   315
      Left            =   5520
      MousePointer    =   1  '矢印
      TabIndex        =   23
      Text            =   "*"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtSpecial 
      BackColor       =   &H00008000&
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
      Height          =   270
      Left            =   2280
      MousePointer    =   1  '矢印
      TabIndex        =   21
      Text            =   "aaaaaaaa"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSiborikomi 
      Caption         =   "絞り込み(&R)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame fraKKey 
      BackColor       =   &H00008000&
      Caption         =   "絞り込み条件"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   480
      TabIndex        =   14
      Top             =   2400
      Width           =   3615
      Begin VB.OptionButton optStandard 
         BackColor       =   &H00008000&
         Caption         =   "標準部品を含む"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optSpecial 
         BackColor       =   &H00008000&
         Caption         =   "指定文字列一致"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00008000&
         Caption         =   "絞り込みなし"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optData 
         BackColor       =   &H00008000&
         Caption         =   "ＥＥＯＳ図面表記データ 一致"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdKettei 
      Caption         =   "決定(&F)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "中止(&C)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   17
      ToolTipText     =   "作業を中断します。 作成データは記憶されます。"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtTokki 
      BackColor       =   &H00008000&
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
      Height          =   315
      Left            =   5520
      MousePointer    =   1  '矢印
      TabIndex        =   13
      Text            =   "*"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox cbomaker 
      BackColor       =   &H00008000&
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
      Height          =   315
      ItemData        =   "Plst_Select.frx":000C
      Left            =   2520
      List            =   "Plst_Select.frx":000E
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cboCode1 
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
      Height          =   315
      Left            =   1560
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   9
      Top             =   1320
      Width           =   7575
   End
   Begin VB.ComboBox cboCode0 
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
      Height          =   315
      Left            =   1560
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   960
      Width           =   7575
   End
   Begin VB.TextBox txtMeisyou0 
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
      Height          =   270
      Left            =   3000
      MousePointer    =   1  '矢印
      TabIndex        =   5
      Text            =   "123456789012345678901234567890"
      Top             =   480
      Width           =   6135
   End
   Begin VB.TextBox txtCodeno 
      BackColor       =   &H00008000&
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
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtpno 
      BackColor       =   &H00008000&
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
      Height          =   270
      Left            =   480
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   " U12"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblBikou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "備考欄"
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
      Height          =   315
      Left            =   4560
      MousePointer    =   1  '矢印
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblTokki 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "特記事項"
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
      Height          =   315
      Left            =   4560
      MousePointer    =   1  '矢印
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblmaker 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ指定"
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
      Height          =   315
      Left            =   1560
      MousePointer    =   1  '矢印
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblcmain 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品目"
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
      Height          =   315
      Left            =   1080
      MousePointer    =   1  '矢印
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblbindex 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種"
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
      Height          =   315
      Left            =   1080
      MousePointer    =   1  '矢印
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblMeisyou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部 品 名 称"
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
      Height          =   255
      Left            =   3000
      MousePointer    =   1  '矢印
      TabIndex        =   4
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label lblCodeno 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ｺｰﾄﾞ番号"
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
      Height          =   255
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblPno 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部品番号"
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
      Height          =   255
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Plst_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*******************
'* 品種・品目選択 ***
'*******************
'
Option Explicit
'
Dim HeadTitle As String
'
Dim FLGdankai As Integer
Dim FLGsai_enable As Integer
Dim BTable() As Integer     '*** 対応表 ***
Dim BTnumber As Integer
Dim FLGhamidasi As Integer
'

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (9705 - 960) * HyoujiBairitu + 480
    Height = 240 + (4710 - 480) * HyoujiBairitu + 240
'
    Left = Eeos2_mainMDI.Left + (Eeos2_mainMDI.ScaleWidth - Width) \ 2 + 80 * HyoujiBairitu
    Top = Eeos2_mainMDI.Top + 480 + (Eeos2_mainMDI.ScaleHeight - Height) \ 2 _
                                            + (1980 - 1320) * HyoujiBairitu
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    Me.Caption = STATUS & "  ＜品種・品目選択＞"
'
    FLGesc = 0
    FLGsai_enable = 0
    FLGhamidasi = 0
'
    txtpno.Text = PLSTT(mps, 0)
    txtMeisyou0.Text = PLSTT(mps, 1)
'
    txtSpecial.Text = SearchKey '*** 検索文字のプリ設定 ***
    txtBikou.ToolTipText = "仕様書番号や指示項目などを記入します。"
    txtTokki.ToolTipText = "水晶の周波数やコイルの容量などを記入します。"
    optData.ToolTipText = "図面表記が回路図の品名に含まれている場合に選択され、" _
        & "該当無い場合は「絞り込み無し」ですべて表示します。"
    optSpecial.ToolTipText = "指定した文字列が項目名に含まれている部品を選択します。"
    optStandard.ToolTipText = "標準部品を含む項目名を表示します。"
'
    Select Case FLGoption
    Case 0
        optData.Value = True
        txtSpecial.Enabled = False
    Case 1
        optSpecial.Value = True
        txtSpecial.Enabled = True
    Case 2
        optAll.Value = True
        txtSpecial.Enabled = False
    Case 3
        optStandard.Value = True
        txtSpecial.Enabled = False
    End Select
'
    cmdSiborikomi_Click
'
End Sub

Private Sub cboCode0_Click()
    Dim Moji As String, Dtemp As String
'
    BTnumber = cboCode0.ListIndex
    jpsT = BTable(BTnumber)
    If jpsT = 0 Then
        Gdata3(0) = "未登録部品"
        FLGdankai = 9
        Call Check_Dankai
    Else
        Call GET_jps(BindexT(jpsT, 0), Aitem0(), ipsT, BindexT(), jpsT, jcpsT, _
                    DRVmainT, CmainT(), CnumT, CdimT, kcpsT)
        cboCode1.Clear
        For kpT = 1 To CnumT
            Moji = "       " & CmainT(kpT, 0) & "   xxx" & CmainT(kpT, 1) & "xxx"
            Call ADD_kuuhaku(Moji, 18)
            Dtemp = CmainT(kpT, 3)
            Call TRSsitei3(Dtemp)
            Moji = Moji & " " & Dtemp
            Call ADD_kuuhaku(Moji, 28)
            cboCode1.AddItem Moji & " " & CmainT(kpT, 2)
        Next kpT
'
        If MEM_ips = ipsT And MEM_jps = jpsT Then
            kpsT = MEM_kps
        Else
            kpsT = 1
        End If
        cboCode1.ListIndex = kpsT - 1
        Gdata3(0) = "L" & BindexT(jpsT, 0) & "-" & CmainT(kpsT, 0)
'
        FLGdankai = 1
        Call Check_Dankai
'
        Call SET_cbomaker
    End If
'
    txtCodeno.Text = Gdata3(0)
End Sub

Private Sub cmdCancel_Click()
    FLGesc = 1
    Unload Me
End Sub

Private Sub cmdKettei_Click()
    Dim SubTemp As String
'
    If FLGhamidasi = 1 Then
        '*** 何もしない ***
    Else
        'PLSTT(mps, 0)                '*** 部品番号(代入済み) ***
        If Left(Trim(Gdata3(0)), 1) <> "L" Then
            PLSTT(mps, 1) = "*" & Trim(txtMeisyou0.Text) '*** 未登録時は名称 ***
        Else
            PLSTT(mps, 1) = Gdata3(0)                       '*** コード番号 ***
        End If
'
        SubTemp = Trim(txtBikou.Text)   '*** 備考欄 ***
        If SubTemp = "" Then
            SubTemp = "*"
        End If
        PLSTT(mps, 2) = SubTemp
'
        PLSTT(mps, 3) = Gdata5(0)           '*** メーカー指定 ***
'
        SubTemp = Trim(txtTokki.Text)   '*** 特記事項 ***
        If SubTemp = "" Then
            SubTemp = "*"
        End If
        PLSTT(mps, 4) = SubTemp
'
        MEM_ips = ipsT
        MEM_jps = jpsT
        MEM_kps = kpsT
    End If
'
    FLGesc = 0
    Unload Me
End Sub

Private Sub cmdSiborikomi_Click()
'   mps: ｸﾞﾛｰﾊﾞﾙ変数, 視点フラグ
    Dim i As Integer
'
    Me.MousePointer = vbHourglass
'
    Gdata1(0) = PLSTT(mps, PdimT + 1)
    Call GET_koumoku(Gdata1(0), Aitem0(), Anum0)    '*** 項目名の検索 ***
'
    If Gdata1(0) = "**" Then   '[はみ出し部品]
        i = MsgBox("この項目番号は登録されていないので選択できません。" & vbCrLf & "このまま決定にしてください。", vbCritical, STATUS2)
        FLGhamidasi = 1
        Exit Sub
'
    End If
'
    Call GET_ips(Gdata1(0), Aitem0(), Anum0, Adim0, ipsT, icpsT, DRVindexT, BindexT(), BnumT, BdimT, jcpsT, kcpsT)  '*** ips の決定 ***
'
    ReDim BTable(BnumT)
    BTable(0) = 0
    FLGdankai = 0           '*** とりあえず指定 ***
'
    Select Case FLGoption
    Case 0
        Call DSPoya_Data        '*** 図面表記による絞り込み ***
    Case 1
        SearchKey = Trim(txtSpecial.Text)
        Call DSPoya_special     '*** 入力文字による絞り込み ***
    Case 2
        Call DSPoya_all         '*** 絞り込み無し ***
    Case 3
        Call DSPoya_std         '*** 標準部品による絞り込み ***
    End Select
'
    Call Check_Dankai
    Beep        '*** 終了の通知 ***
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCode1_Click()
'
    kpsT = cboCode1.ListIndex + 1
'
    If CmainT(kpsT, 16) = "1" Then    '*** 特記事項 ***
        lblTokki.Enabled = True
        txtTokki.Enabled = True
        txtTokki.Text = ""
    Else
        lblTokki.Enabled = False
        txtTokki.Enabled = False
    End If
'
    Gdata3(0) = "L" & BindexT(jpsT, 0) & "-" & CmainT(kpsT, 0)
    txtCodeno.Text = Gdata3(0)
'
    Gdata5(0) = "0"
End Sub

Private Sub cbomaker_Click()
    Select Case cbomaker.ListIndex
    Case 0
        Gdata5(0) = "0"
    Case 1
        Gdata5(0) = "8"
    Case 2
        Gdata5(0) = "9"
    Case 3
        Gdata5(0) = "10"
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1
    End If
End Sub

Private Sub optAll_Click()
    FLGoption = 2
    FLGsai_enable = 0
    txtSpecial.Enabled = False
    FLGdankai = 0
    Call Check_Dankai
    cmdSiborikomi_Click     '*** 絞り込み実行 ***
End Sub

Private Sub optData_Click()
    FLGoption = 0
    FLGsai_enable = 0
    txtSpecial.Enabled = False
    FLGdankai = 0
    Call Check_Dankai
    cmdSiborikomi_Click     '*** 絞り込み実行 ***
End Sub

Private Sub optSpecial_Click()
    FLGoption = 1
    FLGsai_enable = 1
    txtSpecial.Enabled = True
    FLGdankai = 0
    Call Check_Dankai
End Sub

Private Sub optstandard_Click()
    FLGoption = 3
    FLGsai_enable = 0
    txtSpecial.Enabled = False
    FLGdankai = 0
    Call Check_Dankai
    cmdSiborikomi_Click     '*** 絞り込み実行 ***
End Sub

Private Sub SET_cbomaker()
    Dim i As Integer
    Dim Mdata As String
'
    If BindexT(jpsT, 5) = "998" Then
        cbomaker.Enabled = True
        lblmaker.Enabled = True
'
            cbomaker.Clear
            cbomaker.AddItem "複数指"
        For i = 8 To 10
            Mdata = BindexT(jpsT, i)
            Call Makerget2(Mdata)
            cbomaker.AddItem Mdata
        Next i
'
        Select Case Val(Gdata5(0))
        Case 0
            cbomaker.ListIndex = 0
        Case 8
            cbomaker.ListIndex = 1
        Case 9
            cbomaker.ListIndex = 2
        Case 10
            cbomaker.ListIndex = 3
        End Select
    Else
        cbomaker.Enabled = False
        lblmaker.Enabled = False
    End If
End Sub

Private Sub DSPoya_Data()
                            '*** 図面表記と一致した親コード表示 ***
    Dim FLGmaker As String
    Dim Moji As String
    Dim Index_Moji As String
    Dim i As Integer
'
    BTnumber = 0
    cboCode0.BackColor = &H8000&
    cboCode0.Clear
    cboCode0.AddItem "未登録部品       (上の欄に直接記入してね！)"
    For jpT = 1 To BnumT                 '*** データー裏表示 ***
        Index_Moji = BindexT(jpT, 6)
        Call removeFullName(Index_Moji)
        
        If InStr(1, PLSTT(mps, 1), Index_Moji, vbTextCompare) > 0 Then
            BTnumber = BTnumber + 1
            BTable(BTnumber) = jpT
'
            FLGmaker = BindexT(jpT, 5)
            Call Makerget2(FLGmaker)
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
        End If
    Next jpT
'
    If BTnumber = 0 Then        '*** 該当なし => 選択なしで表示 ***
        cboCode0.BackColor = &H8080&
        For jpT = 1 To BnumT                 '*** データー裏表示 ***
            FLGmaker = BindexT(jpT, 5)
            Call Makerget2(FLGmaker)
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
'
            BTnumber = BTnumber + 1
            BTable(BTnumber) = jpT
        Next jpT
    End If
'
    If MEM_ips = ipsT Then
        For i = 1 To BTnumber
            If BTable(i) = jpsT Then
                cboCode0.ListIndex = i
                cboCode0_Click          '*** 子コードの設定 ***
'
                FLGdankai = 1
                Call Check_Dankai
                Exit Sub
            End If
        Next i
    End If
'
    If BTnumber = 0 Then
        cboCode0.ListIndex = 0
    Else
        cboCode0.ListIndex = 1
    End If
End Sub

Private Sub DSPoya_special()
                            '*** 記入文字と一致した親コード表示 ***
    Dim FLGmaker As String
    Dim Moji As String
    Dim i As Integer
'
    BTnumber = 0
    cboCode0.Clear
    cboCode0.AddItem "未登録部品       (上の欄に直接記入してね！)"
    For jpT = 1 To Bnum0                 '*** データー裏表示 ***
        If InStr(1, BindexT(jpT, 3), SearchKey, vbTextCompare) > 0 Then
            BTnumber = BTnumber + 1
            BTable(BTnumber) = jpT
'
            FLGmaker = BindexT(jpT, 5)
            Call Makerget2(FLGmaker)
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
        End If
    Next jpT
'
    If MEM_ips = ipsT Then
        For i = 1 To BTnumber
            If BTable(i) = MEM_jps Then
                cboCode0.ListIndex = i
                cboCode0_Click          '*** 子コードの設定 ***
'
                FLGdankai = 1
                Call Check_Dankai
                Exit Sub
            End If
        Next i
    End If
'
    If BTnumber = 0 Then
        cboCode0.ListIndex = 0
    Else
        cboCode0.ListIndex = 1
    End If
End Sub

Private Sub DSPoya_std()
                            '*** 標準部品を含む親コード表示 ***
    Dim FLGmaker As String
    Dim Moji As String
    Dim i As Integer
'
    BTnumber = 0
    cboCode0.BackColor = &H8000&
    cboCode0.Clear
    cboCode0.AddItem "未登録部品       (上の欄に直接記入してね！)"
    
    For jpT = 1 To BnumT                 '*** データー裏表示 ***
        If BindexT(jpT, 15) = "1" Then
            BTnumber = BTnumber + 1
            BTable(BTnumber) = jpT
'
            FLGmaker = BindexT(jpT, 5)
            Call Makerget2(FLGmaker)
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
        End If
    Next jpT
'
    If BTnumber = 0 Then        '*** 該当なし => 選択なしで表示 ***
        cboCode0.BackColor = &H8080&
        For jpT = 1 To BnumT                 '*** データー裏表示 ***
            FLGmaker = BindexT(jpT, 5)
            Call Makerget2(FLGmaker)
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
'
            BTnumber = BTnumber + 1
            BTable(BTnumber) = jpT
        Next jpT
    End If
'
    If MEM_ips = ipsT Then
        For i = 1 To BTnumber
            If BTable(i) = jpsT Then
                cboCode0.ListIndex = i
                cboCode0_Click          '*** 子コードの設定 ***
'
                FLGdankai = 1
                Call Check_Dankai
                Exit Sub
            End If
        Next i
    End If
'
    If BTnumber = 0 Then
        cboCode0.ListIndex = 0
    Else
        cboCode0.ListIndex = 1
    End If
End Sub

Private Sub DSPoya_all()
                            '*** 選択なしですべての親コード表示 ***
    Dim FLGmaker As String
    Dim Moji As String
'
    BTnumber = 0
    cboCode0.Clear
    cboCode0.AddItem "未登録部品       (上の欄に直接記入してね！)"
    For jpT = 1 To BnumT                 '*** データー裏表示 ***
        FLGmaker = BindexT(jpT, 5)
        Call Makerget2(FLGmaker)
        Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
        Call ADD_kuuhaku(Moji, 24)
        cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
'
        BTnumber = BTnumber + 1
        BTable(BTnumber) = jpT
    Next jpT
'
    If MEM_ips = ipsT Then
        cboCode0.ListIndex = MEM_jps
        cboCode0_Click          '*** 子コードの設定 ***
'
        FLGdankai = 1
        Call Check_Dankai
    Else
        If BTnumber = 0 Then
            cboCode0.ListIndex = 0
        Else
            cboCode0.ListIndex = 1
        End If
    End If
'
End Sub

Private Sub txtBikou_Click()
    txtBikou.MousePointer = vbIbeam
End Sub

Private Sub txtBikou_LostFocus()
    txtBikou.MousePointer = vbArrow
End Sub

Private Sub txtMeisyou0_GotFocus()
    If Left(Trim(Gdata3(0)), 1) = "L" Then
        txtMeisyou0.Locked = True
    End If
End Sub

Private Sub txtMeisyou0_LostFocus()
    txtMeisyou0.Locked = False
End Sub

Private Sub txtSpecial_Click()
    txtSpecial.MousePointer = vbIbeam
End Sub

Private Sub txtSpecial_LostFocus()
    txtSpecial.MousePointer = vbArrow
End Sub

Private Sub txtTokki_Click()
    txtTokki.MousePointer = vbIbeam
End Sub

Private Sub txtTokki_LostFocus()
    txtTokki.MousePointer = vbArrow
End Sub

Private Sub Check_Dankai()
                            '*** 段階に合わせてテキストボックスを許可する。 ***
    Select Case FLGdankai
    Case 0
        lblCodeno.Enabled = False
        txtCodeno.Enabled = False
        lblcmain.Enabled = False
        cboCode1.Enabled = False
        lblBikou.Enabled = False
        txtBikou.Enabled = False
        lblmaker.Enabled = False
        cbomaker.Enabled = False
        lblTokki.Enabled = False
        txtTokki.Enabled = False
    Case 1
        lblCodeno.Enabled = True
        txtCodeno.Enabled = True
        lblcmain.Enabled = True
        cboCode1.Enabled = True
        lblBikou.Enabled = True
        txtBikou.Enabled = True
        lblmaker.Enabled = False
        cbomaker.Enabled = False
        lblTokki.Enabled = False
        txtTokki.Enabled = False
    Case 9
        lblCodeno.Enabled = False
        txtCodeno.Enabled = False
        lblcmain.Enabled = False
        cboCode1.Enabled = False
        lblBikou.Enabled = True
        txtBikou.Enabled = True
        lblmaker.Enabled = False
        cbomaker.Enabled = False
        lblTokki.Enabled = False
        txtTokki.Enabled = False
    End Select
'
    If FLGsai_enable = 0 Then
        cmdSiborikomi.Enabled = False
    Else
        cmdSiborikomi.Enabled = True
    End If
End Sub

Private Sub DSPgamenBuhin()
    lblPno.Left = 480
    lblPno.Top = 240
    lblPno.FontSize = 10 * HyoujiBairitu
    lblPno.Width = 975 * HyoujiBairitu
    lblPno.Height = 255 * HyoujiBairitu
'
    txtpno.Left = 480
    txtpno.Top = 240 + (480 - 240) * HyoujiBairitu
    txtpno.FontSize = 10 * HyoujiBairitu
    txtpno.Width = 975 * HyoujiBairitu
    txtpno.Height = 270 * HyoujiBairitu
'
    lblCodeno.Left = 480 + (1680 - 480) * HyoujiBairitu
    lblCodeno.Top = 240
    lblCodeno.FontSize = 10 * HyoujiBairitu
    lblCodeno.Width = 1095 * HyoujiBairitu
    lblCodeno.Height = 255 * HyoujiBairitu
'
    txtCodeno.Left = 480 + (1680 - 480) * HyoujiBairitu
    txtCodeno.Top = 240 + (480 - 240) * HyoujiBairitu
    txtCodeno.FontSize = 10 * HyoujiBairitu
    txtCodeno.Width = 1095 * HyoujiBairitu
    txtCodeno.Height = 270 * HyoujiBairitu
'
    lblMeisyou.Left = 480 + (3000 - 480) * HyoujiBairitu
    lblMeisyou.Top = 240
    lblMeisyou.FontSize = 10 * HyoujiBairitu
    lblMeisyou.Width = 6135 * HyoujiBairitu
    lblMeisyou.Height = 255 * HyoujiBairitu
'
    txtMeisyou0.Left = 480 + (3000 - 480) * HyoujiBairitu
    txtMeisyou0.Top = 240 + (480 - 240) * HyoujiBairitu
    txtMeisyou0.FontSize = 10 * HyoujiBairitu
    txtMeisyou0.Width = 6135 * HyoujiBairitu
    txtMeisyou0.Height = 270 * HyoujiBairitu
'
    cboCode0.Left = 480 + (1575 - 480) * HyoujiBairitu
    cboCode0.Top = 240 + (960 - 240) * HyoujiBairitu
    cboCode0.FontSize = 10 * HyoujiBairitu
    cboCode0.Width = 7575 * HyoujiBairitu
'   cboCode0.Height = 315 * HyoujiBairitu
'
    lblbindex.Left = 480 + (1080 - 480) * HyoujiBairitu
    lblbindex.Top = 240 + (960 - 240) * HyoujiBairitu
    lblbindex.FontSize = 10 * HyoujiBairitu
    lblbindex.Width = 495 * HyoujiBairitu
    lblbindex.Height = cboCode0.Height
'
    cboCode1.Left = 480 + (1575 - 480) * HyoujiBairitu
    cboCode1.Top = 240 + (1320 - 240) * HyoujiBairitu
    cboCode1.FontSize = 10 * HyoujiBairitu
    cboCode1.Width = 7575 * HyoujiBairitu
'   cboCode1.Height = 315 * HyoujiBairitu
'
    lblcmain.Left = 480 + (1080 - 480) * HyoujiBairitu
    lblcmain.Top = 240 + (1320 - 240) * HyoujiBairitu
    lblcmain.FontSize = 10 * HyoujiBairitu
    lblcmain.Width = 495 * HyoujiBairitu
    lblcmain.Height = cboCode1.Height
'
    cbomaker.Left = 480 + (2535 - 480) * HyoujiBairitu
    cbomaker.Top = 240 + (1800 - 240) * HyoujiBairitu
    cbomaker.FontSize = 10 * HyoujiBairitu
    cbomaker.Width = 1455 * HyoujiBairitu
'   cbomaker.Height = 315 * HyoujiBairitu
'
    lblmaker.Left = 480 + (1560 - 480) * HyoujiBairitu
    lblmaker.Top = 240 + (1800 - 240) * HyoujiBairitu
    lblmaker.FontSize = 10 * HyoujiBairitu
    lblmaker.Width = 975 * HyoujiBairitu
    lblmaker.Height = cbomaker.Height
'
    txtTokki.Left = 480 + (5520 - 480) * HyoujiBairitu
    txtTokki.Top = 240 + (1800 - 240) * HyoujiBairitu
    txtTokki.FontSize = 10 * HyoujiBairitu
    txtTokki.Width = 2415 * HyoujiBairitu
    txtTokki.Height = 315 * HyoujiBairitu
'
    lblTokki.Left = 480 + (4560 - 480) * HyoujiBairitu
    lblTokki.Top = 240 + (1800 - 240) * HyoujiBairitu
    lblTokki.FontSize = 10 * HyoujiBairitu
    lblTokki.Width = 975 * HyoujiBairitu
    lblTokki.Height = txtTokki.Height
'
    txtBikou.Left = 480 + (5520 - 480) * HyoujiBairitu
    txtBikou.Top = 240 + (2280 - 240) * HyoujiBairitu
    txtBikou.FontSize = 10 * HyoujiBairitu
    txtBikou.Width = 2415 * HyoujiBairitu
    txtBikou.Height = 315 * HyoujiBairitu
'
    lblBikou.Left = 480 + (4560 - 480) * HyoujiBairitu
    lblBikou.Top = 240 + (2280 - 240) * HyoujiBairitu
    lblBikou.FontSize = 10 * HyoujiBairitu
    lblBikou.Width = 975 * HyoujiBairitu
    lblBikou.Height = txtBikou.Height
'
    fraKKey.Left = 480
    fraKKey.Top = 240 + (2400 - 240) * HyoujiBairitu
    fraKKey.FontSize = 9 * HyoujiBairitu
    fraKKey.Width = 3615 * HyoujiBairitu
    fraKKey.Height = 1695 * HyoujiBairitu
'
    optData.Left = 120 * HyoujiBairitu
    optData.Top = 240 * HyoujiBairitu
    optData.FontSize = 9 * HyoujiBairitu
    optData.Width = 2535 * HyoujiBairitu
    optData.Height = 255 * HyoujiBairitu
'
    optSpecial.Left = 120 * HyoujiBairitu
    optSpecial.Top = 600 * HyoujiBairitu
    optSpecial.FontSize = 9 * HyoujiBairitu
    optSpecial.Width = 1575 * HyoujiBairitu
    optSpecial.Height = 255 * HyoujiBairitu
'
    optStandard.Left = 120 * HyoujiBairitu
    optStandard.Top = 960 * HyoujiBairitu
    optStandard.FontSize = 9 * HyoujiBairitu
    optStandard.Width = 1695 * HyoujiBairitu
    optStandard.Height = 255 * HyoujiBairitu
'
    optAll.Left = 120 * HyoujiBairitu
    optAll.Top = 1320 * HyoujiBairitu
    optAll.FontSize = 9 * HyoujiBairitu
    optAll.Width = 1455 * HyoujiBairitu
    optAll.Height = 255 * HyoujiBairitu
'
    txtSpecial.Left = 480 + (2280 - 480) * HyoujiBairitu
    txtSpecial.Top = 240 + (3000 - 240) * HyoujiBairitu
    txtSpecial.FontSize = 10 * HyoujiBairitu
    txtSpecial.Width = 1095 * HyoujiBairitu
    txtSpecial.Height = 270 * HyoujiBairitu
'
    cmdSiborikomi.Left = 480 + (2520 - 480) * HyoujiBairitu
    cmdSiborikomi.Top = 240 + (3360 - 240) * HyoujiBairitu
    cmdSiborikomi.FontSize = 10 * HyoujiBairitu
    cmdSiborikomi.Width = 1215 * HyoujiBairitu
    cmdSiborikomi.Height = 375 * HyoujiBairitu
'
    cmdCancel.Left = 480 + (6120 - 480) * HyoujiBairitu
    cmdCancel.Top = 240 + (3240 - 240) * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1095 * HyoujiBairitu
    cmdCancel.Height = 615 * HyoujiBairitu
'
    cmdKettei.Left = 480 + (7800 - 480) * HyoujiBairitu
    cmdKettei.Top = 240 + (3240 - 240) * HyoujiBairitu
    cmdKettei.FontSize = 10 * HyoujiBairitu
    cmdKettei.Width = 1095 * HyoujiBairitu
    cmdKettei.Height = 615 * HyoujiBairitu
End Sub


