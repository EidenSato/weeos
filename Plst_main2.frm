VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Plst_main2 
   BackColor       =   &H00004000&
   Caption         =   "部品表内容表示画面"
   ClientHeight    =   7680
   ClientLeft      =   450
   ClientTop       =   1710
   ClientWidth     =   11400
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
   Icon            =   "Plst_main2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   7680
   ScaleWidth      =   11400
   Begin VB.CommandButton cmdHelp 
      Caption         =   "？"
      Height          =   255
      Left            =   10560
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optTanka 
      BackColor       =   &H00004000&
      Caption         =   "単価,,形状"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   6720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optOndo 
      BackColor       =   &H00004000&
      Caption         =   "MSL/耐熱,,ﾒｯｷ/RoHS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox txtFolder 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "\1234\"
      Top             =   240
      Width           =   8175
   End
   Begin VB.CommandButton cmdAppend 
      Caption         =   "部品追加(&A)"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtKiji 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "123456789012345678901234567890"
      Top             =   960
      Width           =   8535
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8040
      TabIndex        =   5
      Text            =   "1997/03/31"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtKmeisyou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
      MousePointer    =   1  '矢印
      TabIndex        =   3
      Text            =   "1234567890123456789012345"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtZuban 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "A1234-001AR12"
      Top             =   600
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   21
      Cols            =   9
      FixedCols       =   0
      BackColor       =   32768
      ForeColor       =   16777215
      BackColorFixed  =   8421376
      ForeColorFixed  =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFolder 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾌｧｲﾙの場所"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblPingoukei 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   " 足ピン数合計："
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label lblSoukei 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部品単価合計："
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblTyuui 
      BackColor       =   &H00004000&
      Caption         =   "！：使用禁止、？：在庫限り、￥：変更推奨"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   3240
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lblKiji 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "記事 (&3)"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblDate 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "日付 (&2)"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7200
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblKmeisyou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "小名称 (&1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblZuban 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾌｧｲﾙ名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "ﾌｧｲﾙ(&X)"
      Begin VB.Menu mnuSinki 
         Caption         =   "新規作成(&N)"
      End
      Begin VB.Menu mnu区切り線10 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileSV 
         Caption         =   "上書き保存(&S)"
      End
      Begin VB.Menu mnufileNW 
         Caption         =   "名前を付けて保存(&R)..."
      End
      Begin VB.Menu mnu区切り線11 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileRD 
         Caption         =   "同一ﾌｧｲﾙ再読込(&C)"
      End
      Begin VB.Menu mnufileNR 
         Caption         =   "別ﾌｧｲﾙ読込(&F)..."
      End
      Begin VB.Menu mnu区切り線12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRNstd 
         Caption         =   "標準部品表印刷(&P)..."
      End
      Begin VB.Menu mnu区切り線13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrn 
         Caption         =   "部品一覧表印刷(&L)..."
      End
      Begin VB.Menu mnuFileWR 
         Caption         =   "一覧表ﾌｧｲﾙ出力(&W)..."
      End
      Begin VB.Menu mnu区切り線14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "閉じる(&Q)"
      End
      Begin VB.Menu mnu区切り線15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAQuit 
         Caption         =   "EEOS２の終了(&X)"
      End
   End
   Begin VB.Menu mnuKouseihyou 
      Caption         =   "構成表(&K)"
      Begin VB.Menu mnuKousei 
         Caption         =   "電気 構成表(&C)..."
      End
   End
   Begin VB.Menu mnuBuhinhyou 
      Caption         =   "部品表(&P)"
      Begin VB.Menu mnuBuhin 
         Caption         =   "電気 部品表(&C)..."
      End
      Begin VB.Menu mnuBuhin2 
         Caption         =   "電気 部品表２(&D)..."
      End
      Begin VB.Menu mnuORCAD 
         Caption         =   "OrCAD変換(&O)..."
      End
      Begin VB.Menu mnuConvFile 
         Caption         =   "変換作業ﾌｧｲﾙ(&W)"
      End
      Begin VB.Menu mnu区切り線31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuhinPRN 
         Caption         =   "部品表印刷(&P)..."
      End
      Begin VB.Menu mnuFilePrnA 
         Caption         =   "一覧表印刷(&L)..."
      End
      Begin VB.Menu mnuSuuryo 
         Caption         =   "数量表印刷(&T)..."
      End
   End
   Begin VB.Menu mnuCodehyou 
      Caption         =   "ｺｰﾄﾞ表(&C)"
      Begin VB.Menu mnuCode 
         Caption         =   "項目一覧(&M)"
      End
      Begin VB.Menu mnuHinsyu 
         Caption         =   "品種一覧(&I)"
      End
      Begin VB.Menu mnuPmain 
         Caption         =   "品目一覧(&P)"
      End
      Begin VB.Menu mnuMakerment 
         Caption         =   "ﾒｰｶｰｺｰﾄﾞ表(&M)"
      End
      Begin VB.Menu mnuTraderment 
         Caption         =   "商社ｺｰﾄﾞ表(&T)"
      End
   End
   Begin VB.Menu mnuJump 
      Caption         =   "ｼﾞｬﾝﾌﾟ(&J)"
      Begin VB.Menu mnuJumpT 
         Caption         =   "先頭へｼﾞｬﾝﾌﾟ(&T)"
      End
      Begin VB.Menu mnuJumpB 
         Caption         =   "前の項目へｼﾞｬﾝﾌﾟ(&B)"
      End
      Begin VB.Menu mnuJumpC 
         Caption         =   "中心へｼﾞｬﾝﾌﾟ(&C)"
      End
      Begin VB.Menu mnuJumpN 
         Caption         =   "次項目へｼﾞｬﾝﾌﾟ(&N)"
      End
      Begin VB.Menu mnuJumpE 
         Caption         =   "最後部へｼﾞｬﾝﾌﾟ(&E)"
      End
   End
   Begin VB.Menu mnuWindou 
      Caption         =   "ｳｲﾝﾄﾞｳ(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileH 
         Caption         =   "上下に並べて表示(&H)"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "左右に並べて表示(&V)"
      End
      Begin VB.Menu mnuTileC 
         Caption         =   "重ねて表示(&C)"
      End
      Begin VB.Menu mnuReform 
         Caption         =   "初期位置に戻す(&S)"
      End
   End
   Begin VB.Menu mnuKnakyou 
      Caption         =   "環境(&O)"
      Begin VB.Menu mnuSettei 
         Caption         =   "環境設定(&K)"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "ｵﾌﾟｼｮﾝ(&O)"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuSetumei 
         Caption         =   "操作説明(&S)"
      End
      Begin VB.Menu mnuKaihan 
         Caption         =   "改版履歴(&H)"
      End
      Begin VB.Menu mnu区切り線81 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ(&V)"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ﾎﾟｯﾌﾟｱｯﾌﾟﾒﾆｭｰ"
      Visible         =   0   'False
      Begin VB.Menu mnuJumpP 
         Caption         =   "ｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpTP 
         Caption         =   "先頭へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpBP 
         Caption         =   "前の項目へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpCP 
         Caption         =   "中心へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpNP 
         Caption         =   "次項目へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpEP 
         Caption         =   "最後部へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnu区切り線91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSinkiP 
         Caption         =   "新規作成"
      End
      Begin VB.Menu mnu区切り線931 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileSVP 
         Caption         =   "上書き保存"
      End
      Begin VB.Menu mnufileNWP 
         Caption         =   "名前を付けて保存..."
      End
      Begin VB.Menu mnu区切り線932 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileRDP 
         Caption         =   "同一ﾌｧｲﾙ再読込"
      End
      Begin VB.Menu mnufileNRP 
         Caption         =   "別ﾌｧｲﾙ読込..."
      End
      Begin VB.Menu mnu区切り線933 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPuKouseihyou 
         Caption         =   "構成表"
         Begin VB.Menu mnuKouseiP 
            Caption         =   "電気 構成表(&C)..."
         End
      End
      Begin VB.Menu mnuPuBuhinhyou 
         Caption         =   "部品表"
         Begin VB.Menu mnuBuhinP 
            Caption         =   "電気 部品表..."
         End
         Begin VB.Menu mnuBuhin2P 
            Caption         =   "電気 部品表２..."
         End
         Begin VB.Menu mnuORCADP 
            Caption         =   "OrCAD変換..."
         End
         Begin VB.Menu mnuConvFileP 
            Caption         =   "変換作業ﾌｧｲﾙ"
         End
         Begin VB.Menu mnu区切り線951 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBuhinPRNP 
            Caption         =   "部品表印刷..."
         End
         Begin VB.Menu mnuFilePrnAP 
            Caption         =   "一覧表印刷..."
         End
         Begin VB.Menu mnuSuuryoP 
            Caption         =   "数量表印刷..."
         End
      End
      Begin VB.Menu mnuPuCodehyou 
         Caption         =   "ｺｰﾄﾞ表"
         Begin VB.Menu mnuCodeP 
            Caption         =   "項目一覧"
         End
         Begin VB.Menu mnuHinsyuP 
            Caption         =   "品種一覧"
         End
         Begin VB.Menu mnuPmainP 
            Caption         =   "品目一覧"
         End
         Begin VB.Menu mnuMakermentP 
            Caption         =   "ﾒｰｶｰｺｰﾄﾞ表"
         End
         Begin VB.Menu mnuTradermentP 
            Caption         =   "商社ｺｰﾄﾞ表"
         End
      End
      Begin VB.Menu mnu区切り線95 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackP 
         Caption         =   "閉じる"
      End
      Begin VB.Menu mnu区切り線96 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAQuitP 
         Caption         =   "EEOS２の終了"
      End
   End
End
Attribute VB_Name = "Plst_main2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 部品表の一覧２ ***
'**********************
'
Option Explicit
'
Dim HeadTitle As String
'
Dim FlagBusy As Boolean
'
Dim FLGoffsetX As Integer
Dim FLGoffsetY As Integer
'                                   567twip=10mm,1440twip=1inch
Private Const OrgWidth = 11640  '*** フォーム寸法初期値 ***
Private Const OrgHeight = 8040
Dim tempWidth As Integer
Dim tempHeight As Integer
'
Dim CHaiki As String
Dim Pmodu As String, Pitem As String, Pno As String
Dim G_row   As Integer
Dim Soukei As Double
Dim goukei_en As Boolean
Dim Pingoukei As Integer
'
Dim DRVpartlist As String    '*** 部品表ディレクトリ ***
Dim PFLname As String        '*** 表示ファイル名 ***
Dim Plistname As String      '*** 機種名 ***
Dim Plistdate As String      '*** 作成日 ***
Dim Remarks As String        '*** 備考欄 ***
Dim Ptotal As Integer        '*** 部品表配列数 ***
Dim Pdim0 As Integer         '*** 部品表次元数 ***
Dim PLST() As String         '*** 部品表データ配列 ***

Private Sub Form_Activate()
    Dim i As Integer, ipoint As Integer
    Dim j As Integer
    Dim k As Integer
    Dim kekka As Integer
'
'   PLST(i, Pdim0 + 1) = Pitem
'   PLST(i, Pdim0 + 2) = Pno
'   PLST(i, Pdim0 + 3) = Pmodu
'
    Me.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
'
    FLGplst2 = 1        '*** 部品表画面存在フラグ設定 ***
    FLGjob = 2
    FLGlevel = 1
    STATUS = HeadTitle  '*** 選択ウインドウのタイトル名称 ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    If FLGtuika = 1 Then        '*** 追加の時のみ以下を実行する。 ***
        Pitem = Gdata2(0)
        Call GETsymbol(Pmodu, Pitem, Pno) '*** 項目と番号を得る ***
'
        ipoint = 0      '*** 位置フラグリセット ***
'
        For i = 1 To Ptotal         '*** 項目/位置の特定 ***
            If PLST(i, Pdim0 + 1) = Pitem Then
                If Pmodu = "" And PLST(i, Pdim0 + 3) = "" Then
                    If Val(PLST(i, Pdim0 + 2)) < Val(Pno) Then
                        ipoint = i + 1  '*** フラグセット ***
'
                    ElseIf Val(PLST(i, Pdim0 + 2)) = Val(Pno) Then  '*** 数値が等しい ***
                        kekka = StrComp(PLST(i, 0), Gdata2(0), 1)   '*** 文字列で確認 ***
                        Select Case kekka
                        Case -1                                     '*** a < b
                            ipoint = i + 1  '*** フラグセット ***
                        Case 0, 1                                   '*** a >= b
                            ipoint = i      '*** フラグセット ***
                            Exit For        '*** ループから抜ける ***
'
                        End Select
                    Else
                        ipoint = i      '*** 追加位置決定 ***
                        Exit For        '*** ループから抜ける ***
'
                    End If
'
                ElseIf Pmodu = "" And PLST(i, Pdim0 + 3) <> "" Then
                    ipoint = i      '*** 追加位置決定 ***
                    Exit For        '*** ループから抜ける ***
'
                Else
                    If Val(Mid(PLST(i, Pdim0 + 3), 2)) < Val(Mid(Pmodu, 2)) Then
                        ipoint = i + 1  '*** フラグセット ***
                    Else
                        If Val(PLST(i, Pdim0 + 2)) < Val(Pno) Then
                            ipoint = i + 1  '*** フラグセット ***
'
                        ElseIf Val(PLST(i, Pdim0 + 2)) = Val(Pno) Then  '*** 数値が等しい ***
                            kekka = StrComp(PLST(i, 0), Gdata2(0), 1)   '*** 文字列で確認 ***
                            Select Case kekka
                            Case -1                                     '*** a < b
                                ipoint = i + 1  '*** フラグセット ***
                            Case 0, 1                                   '*** a >= b
                                ipoint = i      '*** フラグセット ***
                                Exit For        '*** ループから抜ける ***
'
                            End Select
                        Else
                            ipoint = i      '*** 追加位置決定 ***
                            Exit For        '*** ループから抜ける ***
'
                        End If
                    End If
                End If
            End If
        Next i
'
        If ipoint = 0 Then          '*** 同一部品番号記号が無い時 ***
            For i = 1 To Anum0
                If Aitem0(i, 0) = Gdata1(0) Then
                    Select Case Pitem
                    Case Aitem0(i, 3)
                        For j = 1 To Ptotal
                            If PLST(j, Pdim0 + 1) = Aitem0(i, 4) Or PLST(j, Pdim0 + 1) = Aitem0(i, 5) Then
                                ipoint = j
                                Exit For
'
                            End If
                        Next j
                        Exit For
'
                    Case Aitem0(i, 4)
                        For j = 1 To Ptotal
                            If PLST(j, Pdim0 + 1) = Aitem0(i, 5) Then
                                ipoint = j
                                Exit For
                            End If
                        Next j
'
                        For j = 1 To Ptotal
                            If PLST(j, Pdim0 + 1) = Aitem0(i, 3) Then
                                ipoint = j + 1
                            End If
                        Next j
                        Exit For
'
                    Case Aitem0(i, 5)
                        For j = 1 To Ptotal
                            If PLST(j, Pdim0 + 1) = Aitem0(i, 3) Or PLST(j, Pdim0 + 1) = Aitem0(i, 4) Then
                                ipoint = j + 1
                            End If
                        Next j
                        Exit For

                    End Select
                End If
            Next i      '*** 何かに該当するのでここで終わることは無い ***
'
            If ipoint = 0 Then          '*** 同一項目が無い時 ***
                For k = i To Anum0      '*** 次項目の前に入れる ***
                    For j = 1 To Ptotal
                        If Aitem0(k, 3) = PLST(j, Pdim0 + 1) Or Aitem0(k, 4) = PLST(j, Pdim0 + 1) Or Aitem0(k, 5) = PLST(j, Pdim0 + 1) Then
                            ipoint = j
                            Exit For
'
                        End If
                    Next j
'
                    If ipoint <> 0 Then
                        Exit For
                    End If
                Next k
'
                If ipoint = 0 Then      '*** 次が無いので最後に置きましょう ***
                    ipoint = Ptotal + 1
                End If
            End If
        End If
'
        Call INS1gyou(ipoint)   '*** １行追加 ***
'
        Call Gamen_settei       '*** 画面設定 内容表示 ***
        FLGchange = 1
'
        If Ptotal < G_row Then
            Beep
            MSFlexGrid1.TopRow = 1
        ElseIf ipoint <= 8 Then
            MSFlexGrid1.TopRow = 1
        ElseIf (Ptotal + 1) - (ipoint - 8) >= G_row Then
            MSFlexGrid1.TopRow = ipoint - 8
        Else
            Beep
            MSFlexGrid1.TopRow = (Ptotal + 1) - G_row + 2
        End If
'
        Call H1up       '*** 上から１つ下に表示する ***
        FLGtuika = 0
'
    Else                '*** 部品表保存したファイル名を表示 ***
        txtZuban.Text = PFLname
    End If
'
    MSFlexGrid1.MousePointer = flexDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
    HeadTitle = STATUS
    FLGplst2 = 1        '*** 部品表画面存在フラグ設定 ***
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        Call Hup        '*** 上へ ***
'
    Case vbKeyUp        '*** ↑ ***
        Call H1up       '*** 一つ上へ ***
'
    Case vbKeyPageUp    '*** Roll Down
        Call Hdown      '*** 下へ ***
'
    Case vbKeyDown      '*** ↓ ***
        Call H1down     '*** 一つ下へ ***
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
                            '*** フォームのサイズの設定
    tempWidth = 360 + (OrgWidth - 720) * HyoujiBairitu + 360
    tempHeight = 360 + (OrgHeight - 720) * HyoujiBairitu + 360
'
    Me.Width = tempWidth    '*** これで「Form_Resize」割り込みが発生する。 ***
    Me.Height = tempHeight
'
    FLGoffsetX = 0          '*** 初期化 ***
    FLGoffsetY = 0
'
    Call setFormArea        '*** フォームの表示位置の設定
'
    Me.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
    Me.Caption = HeadTitle
'
    CHaiki = "部品表は変更されています。「廃棄終了」をキャンセルしますか？"
'
    FlagBusy = True
    optTanka.Value = True
    optOndo.Value = False
    FlagBusy = False
'
    DRVmaker = Xcont0(2) & "\MAKER.COD"
    Call RDmaker        '*** メーカーコード 読み込み ***
'
    If FLG_job_error_end = 1 Then
        Kankyow_Itiran.Show 1
'
        FLG_job_error_end = 0
        Unload Me
    End If
'
    DRVitem0 = Xcont0(2) & "\ITEM.COD"
    Call RDitem(DRVitem0, Aitem0(), Anum0, Adim0) '*** ITEM.COD 読み込み ***
'
    icpsT = 0       '*** コード項目番号初期化 ***
    jcpsT = 0
    kcpsT = 0
    FLGesc = 0
    goukei_en = False
'
    If FLGesc = 0 Then
        Select Case FLGlevel
        Case 3          '*** 部品表新規作成 ***
            DRVpartlist = DRVpartlistT
            PFLname = PFLnameT
'
            Ptotal = 1
            Pdim0 = cPdim0  '*** 部品表項目数設定
            ReDim PLST(Ptotal, Pdim0 + 3)
'
            Plistname = "*"
            Plistdate = Format(Date, "yyyy/mm/dd")
            Remarks = "*"
'
            PLST(1, 0) = "Z1"
            PLST(1, 1) = "*削除してね"
            PLST(1, 2) = "*"
            PLST(1, 3) = "0"
            PLST(1, 4) = "*"
            PLST(1, Pdim0 + 1) = "Z"
            PLST(1, Pdim0 + 2) = "1"
            PLST(1, Pdim0 + 3) = ""
            FLGchange = 0
'
'            Call RDplstWork     '*** 部品表作成データ読み込み ***
'
        Case 1, 2
            DRVpartlist = DRVpartlistT
            PFLname = PFLnameT
            Call RDpartlist(DRVpartlist, Plistname, Plistdate, Remarks, PLST(), Ptotal, Pdim0)      '*** 部品表読み込み ***
            FLGchange = 0
'
'            Call RDplstWork      '*** 部品表作成データ読み込み ***
'
            For i = 1 To Ptotal
                Pitem = PLST(i, 0)
                Call GETsymbol(Pmodu, Pitem, Pno) '*** 項目と番号を得る ***
                PLST(i, Pdim0 + 1) = Pitem
                PLST(i, Pdim0 + 2) = Pno
                PLST(i, Pdim0 + 3) = Pmodu
            Next i
        End Select
'
        Call Gamen_settei   '*** 画面設定 内容表示 ***
    End If
    MSFlexGrid1.MousePointer = flexDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
'
    If UnloadMode = vbFormControlMenu Then
        If FLGchange = 1 Then
            Beep
            i = MsgBox(CHaiki, vbQuestion Or vbYesNo, STATUS)
            If i = vbYes Then
                Cancel = True   '*** Unload の停止 ***
                FLGowari = 1
            Else
                Unload Me
            End If
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Resize()
'                   フォーム構成部品の表示位置の設定
    If Me.Width > tempWidth Then
        FLGoffsetX = Me.Width - tempWidth
    Else
        FLGoffsetX = 0
    End If
'
    If Me.Height > tempHeight Then
        FLGoffsetY = (Me.Height - tempHeight) \ 2
    Else
        FLGoffsetY = 0
    End If
'
    With MSFlexGrid1          '***グリッドの幅の設定
        .Width = 10820 * HyoujiBairitu + FLGoffsetX
        .Font.Size = 10 * HyoujiBairitu
        .RowHeightMin = 245 * HyoujiBairitu  '245
        .Height = MSFlexGrid1.RowHeightMin * 21
        .Cols = 9
        .ColWidth(0) = 400 * HyoujiBairitu
        .ColWidth(1) = 450 * HyoujiBairitu
        .ColWidth(2) = 900 * HyoujiBairitu
        .ColWidth(3) = 1070 * HyoujiBairitu + (300 * HyoujiBairitu - 300)
        .ColWidth(4) = 2880 * HyoujiBairitu + FLGoffsetX * 0.8
        .ColWidth(5) = 840 * HyoujiBairitu
        .ColWidth(6) = 1600 * HyoujiBairitu + FLGoffsetX * 0.2
        .ColWidth(7) = 960 * HyoujiBairitu
        .ColWidth(8) = 1380 * HyoujiBairitu
    End With
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FLGplst2 = 0    '*** 部品表画面存在フラグ初期化 ***
End Sub

Private Sub Copy_plst_work2temp()
    PFLnameT = PFLname
    DRVpartlistT = DRVpartlist
    PlistnameT = Plistname
    PlistdateT = Plistdate
'
    PtotalT = Ptotal
    PdimT = Pdim0
    ReDim PLSTT(Ptotal, Pdim0)
'
    PLSTT() = PLST()
'
    RemarksT = Remarks
End Sub

Private Sub CHK7moji(Ddata As String)
    If Len(Ddata) < 8 Then
        Ddata = " " & Ddata
    End If
End Sub

Private Sub CHKmitou(Ddata As String, FLGmitou As Integer)
    If Left(Ddata, 1) = "*" Then
        Ddata = "未登録部品"
        FLGmitou = 1
    Else
        Ddata = " " & Ddata
        FLGmitou = 0
    End If
End Sub

Private Sub DEL1gyou(Mpoint As Integer)
    Dim i As Integer, j As Integer
    Dim Ddata As String
'
    For i = Mpoint To Ptotal - 1
        For j = 0 To Pdim0 + 3
            PLST(i, j) = PLST(i + 1, j)
        Next j
'
        With MSFlexGrid1
            .Col = 1                '*** 項目名 ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 2                '*** 部品番号 ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 3                '*** 部品ｺｰﾄﾞ ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 4                '*** 部品名称 ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 5                '*** ﾒｰｶｰ略称 ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 6                '*** 備考 ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 7                '*** 平均単価 ***
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
'
            .Col = 8
            .Row = i + 1
            Ddata = .Text
            .Row = i
            .Text = Ddata
        End With
    Next i
'
    With MSFlexGrid1
        .Row = Ptotal
        .Col = 0                '*** 行 ***
        .Text = ""
        .Col = 1                '*** 項目名 ***
        .Text = ""
        .Col = 2                '*** 部品番号 ***
        .Text = ""
        .Col = 3                '*** 部品ｺｰﾄﾞ ***
        .Text = ""
        .Col = 4                '*** 部品名称 ***
        .Text = ""
        .Col = 5                '*** ﾒｰｶｰ略称 ***
        .Text = ""
        .Col = 6                '*** 備考 ***
        .Text = ""
        .Col = 7                '*** 平均単価 ***
        .Text = ""
        .Col = 8
        .Text = ""
    End With
'
    Ptotal = Ptotal - 1
End Sub

Private Sub DSPgamenBuhin()
    txtFolder.FontSize = 10 * HyoujiBairitu
    txtFolder.Height = 285 * HyoujiBairitu
    txtFolder.Left = 360 + (1570 - 360) * HyoujiBairitu
    txtFolder.Top = 360 + FLGoffsetY
    txtFolder.Width = 8175 * HyoujiBairitu
'
    lblFolder.FontSize = 10 * HyoujiBairitu
    lblFolder.Height = txtFolder.Height
    lblFolder.Left = 360
    lblFolder.Top = 360 + FLGoffsetY
    lblFolder.Width = 1215 * HyoujiBairitu
'
    txtZuban.FontSize = 10 * HyoujiBairitu
    txtZuban.Height = 285 * HyoujiBairitu
    txtZuban.Left = 360 + (1210 - 360) * HyoujiBairitu
    txtZuban.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    txtZuban.Width = 1815 * HyoujiBairitu
'
    lblZuban.FontSize = 10 * HyoujiBairitu
    lblZuban.Height = txtZuban.Height
    lblZuban.Left = 360
    lblZuban.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    lblZuban.Width = 855 * HyoujiBairitu
'
    txtKmeisyou.FontSize = 10 * HyoujiBairitu
    txtKmeisyou.Height = 285 * HyoujiBairitu
    txtKmeisyou.Left = 360 + (4210 - 360) * HyoujiBairitu
    txtKmeisyou.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    txtKmeisyou.Width = 2775 * HyoujiBairitu
'
    lblKmeisyou.FontSize = 10 * HyoujiBairitu
    lblKmeisyou.Height = txtKmeisyou.Height
    lblKmeisyou.Left = 360 + (3240 - 360) * HyoujiBairitu
    lblKmeisyou.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    lblKmeisyou.Width = 975 * HyoujiBairitu
'
    txtDate.FontSize = 10 * HyoujiBairitu
    txtDate.Height = 285 * HyoujiBairitu
    txtDate.Left = 360 + (8050 - 360) * HyoujiBairitu
    txtDate.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    txtDate.Width = 1695 * HyoujiBairitu
'
    lblDate.FontSize = 10 * HyoujiBairitu
    lblDate.Height = txtDate.Height
    lblDate.Left = 360 + (7200 - 360) * HyoujiBairitu
    lblDate.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    lblDate.Width = 855 * HyoujiBairitu
'
    txtKiji.FontSize = 10 * HyoujiBairitu
    txtKiji.Height = 285 * HyoujiBairitu
    txtKiji.Left = 360 + (1210 - 360) * HyoujiBairitu
    txtKiji.Top = 360 + (1080 - 360) * HyoujiBairitu + FLGoffsetY
    txtKiji.Width = 8535 * HyoujiBairitu
'
    lblKiji.FontSize = 10 * HyoujiBairitu
    lblKiji.Height = txtKiji.Height
    lblKiji.Left = 360
    lblKiji.Top = 360 + (1080 - 360) * HyoujiBairitu + FLGoffsetY
    lblKiji.Width = 855 * HyoujiBairitu
'
    cmdHelp.Left = 360 + (10440 - 360) * HyoujiBairitu
    cmdHelp.Top = 360 + (1140 - 360) * HyoujiBairitu + FLGoffsetY
    cmdHelp.FontSize = 10 * HyoujiBairitu
    cmdHelp.Width = 255 * HyoujiBairitu
    cmdHelp.Height = 255 * HyoujiBairitu
'
    MSFlexGrid1.Left = 360
    MSFlexGrid1.Top = 360 + (1440 - 360) * HyoujiBairitu + FLGoffsetY
'
    lblSoukei.FontSize = 10 * HyoujiBairitu
    lblSoukei.Height = 285 * HyoujiBairitu
    lblSoukei.Left = 360
    lblSoukei.Top = 360 + (6720 - 360) * HyoujiBairitu + FLGoffsetY
    lblSoukei.Width = 2535 * HyoujiBairitu
'
    lblPingoukei.FontSize = 10 * HyoujiBairitu
    lblPingoukei.Height = 285 * HyoujiBairitu
    lblPingoukei.Left = 360
    lblPingoukei.Top = 360 + (7080 - 360) * HyoujiBairitu + FLGoffsetY
    lblPingoukei.Width = 2535 * HyoujiBairitu
'
    lblTyuui.FontSize = 9 * HyoujiBairitu
    lblTyuui.Height = 585 * HyoujiBairitu
    lblTyuui.Left = 360 + (3240 - 360) * HyoujiBairitu
    lblTyuui.Top = 360 + (6720 - 360) * HyoujiBairitu + FLGoffsetY
    lblTyuui.Width = 1215 * HyoujiBairitu
'
    cmdAppend.FontSize = 10 * HyoujiBairitu
    cmdAppend.Height = 495 * HyoujiBairitu
    cmdAppend.Left = 360 + (4560 - 360) * HyoujiBairitu
    cmdAppend.Top = 360 + (6840 - 360) * HyoujiBairitu + FLGoffsetY
    cmdAppend.Width = 1215 * HyoujiBairitu
'
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
    cmdUp.Left = 360 + (6000 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (6840 - 360) * HyoujiBairitu + FLGoffsetY
    cmdUp.Width = 735 * HyoujiBairitu
'
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
    cmdDown.Left = 360 + (6840 - 360) * HyoujiBairitu
    cmdDown.Top = 360 + (6840 - 360) * HyoujiBairitu + FLGoffsetY
    cmdDown.Width = 735 * HyoujiBairitu
'
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
    cmdQuit.Left = 360 + (7800 - 360) * HyoujiBairitu
    cmdQuit.Top = 360 + (6840 - 360) * HyoujiBairitu + FLGoffsetY
    cmdQuit.Width = 975 * HyoujiBairitu
'
    optTanka.FontSize = 10 * HyoujiBairitu
    optTanka.Height = 255 * HyoujiBairitu
    optTanka.Left = 360 + (9000 - 360) * HyoujiBairitu
    optTanka.Top = 360 + (6840 - 360) * HyoujiBairitu + FLGoffsetY
    optTanka.Width = 1815 * HyoujiBairitu
'
    optOndo.FontSize = 10 * HyoujiBairitu
    optOndo.Height = 255 * HyoujiBairitu
    optOndo.Left = 360 + (9000 - 360) * HyoujiBairitu
    optOndo.Top = 360 + (7080 - 360) * HyoujiBairitu + FLGoffsetY
    optOndo.Width = 2295 * HyoujiBairitu
End Sub

Private Sub INS1gyou(Mpoint As Integer)
    Dim i As Integer, j As Integer
    Dim Ddata As String
    Dim temp() As String
'
    ReDim temp(Ptotal, Pdim0 + 3)
'
    For i = 1 To Ptotal
        For j = 0 To Pdim0 + 3
            temp(i, j) = PLST(i, j)
        Next j
    Next i
'
    ReDim PLST(Ptotal + 1, Pdim0 + 3)   '*** 項目数１増加 ***
'
    For i = 1 To Mpoint - 1
        For j = 0 To Pdim0 + 3
            PLST(i, j) = temp(i, j)
        Next j
    Next i
'
    PLST(Mpoint, 0) = Gdata2(0)     '*** 部品番号 ***
    PLST(Mpoint, 1) = Gdata3(0)     '*** 部品コード ***
    PLST(Mpoint, 2) = "*"           '*** 備考 ***
    PLST(Mpoint, 3) = Gdata5(0)     '*** メーカー指定 ***
    PLST(Mpoint, 4) = Gdata6(0)     '*** 特記事項 ***
        Pitem = Gdata2(0)
        Call GETsymbol(Pmodu, Pitem, Pno)
    PLST(Mpoint, Pdim0 + 1) = Pitem
    PLST(Mpoint, Pdim0 + 2) = Pno
    PLST(Mpoint, Pdim0 + 3) = Pmodu
'
    For i = Mpoint To Ptotal   '*** ipoint = ptotal + 1 の時は何もしない ***
        For j = 0 To Pdim0 + 3
            PLST(i + 1, j) = temp(i, j)
        Next j
    Next i
'
    Ptotal = Ptotal + 1
End Sub

Private Sub Gamen_settei()
    Dim tempSt As String
    Dim i As Integer
'
    Soukei = 0
    Pingoukei = 0
    goukei_en = True
'
    i = InStr(1, DRVpartlist, PFLname)
    tempSt = Left(DRVpartlist, i - 1)
    txtFolder.Text = tempSt
    txtFolder.MousePointer = vbArrow
'
    Call remove_PLT(PFLname, tempSt)
'
    txtZuban.Text = tempSt
    txtZuban.MousePointer = vbArrow
    txtKmeisyou.Text = Trim(Plistname)
    txtKmeisyou.MousePointer = vbArrow
    txtDate.Text = Trim(Plistdate)
    txtDate.MousePointer = vbArrow
    txtKiji.Text = Trim(Remarks)
    txtKiji.MousePointer = vbArrow
'
    Call setMSFlexGrid1     '*** グリッドの諸元設定 ***
'
    Call set_Pname          '*** 部品名称の表示 ***
'
    lblSoukei.Caption = "部品単価合計： \ " & Format(Soukei, "#,###,##0.0")
    lblPingoukei.Caption = " 足ピン数合計：  " & Format(Pingoukei, "#,###,##0")
    goukei_en = False
End Sub

Private Sub GETpname(FLGkoumoku As String, FLGpname As String, FLGmaker As String, Bsitei As Integer, _
                    FLGtanka As Currency, FLGkeijou As String, FLGondo As String, FLGmekki As String)
    Dim i As Integer, j As Integer, k As Integer
    Dim FLGbcod As String, Dtemp As String
'
    For i = 1 To Anum0
        If FLGkoumoku = Aitem0(i, 0) Then
            Call GET_ips(FLGkoumoku, Aitem0(), Anum0, Adim0, ipsT, icpsT, DRVindexT, BindexT(), BnumT, BdimT, jcpsT, kcpsT)
'
            For j = 1 To BnumT
                If InStr(FLGpname, BindexT(j, 0)) <> 0 Then
                    FLGbcod = BindexT(j, 0)
                    Call GET_jps(FLGbcod, Aitem0(), ipsT, BindexT(), jpsT, jcpsT, DRVmainT, CmainT(), CnumT, CdimT, kcpsT)
'
                    For k = 1 To CnumT
                        FLGmaker = BindexT(j, 5)
                        If FLGpname = "L" & FLGbcod & "-" & CmainT(k, 0) Then
                            If FLGmaker = "998" And Bsitei = 8 Then
                                If BindexT(j, 4) = "*" Then
                                    FLGpname = BindexT(j, 3) & CmainT(k, 1)
                                Else
                                    FLGpname = BindexT(j, 3) & CmainT(k, 1) & BindexT(j, 4)
                                End If
                            ElseIf FLGmaker = "998" And Bsitei = 9 Then
                                If BindexT(j, 12) = "*" Then
                                    FLGpname = BindexT(j, 11) & CmainT(k, 1)
                                Else
                                    FLGpname = BindexT(j, 11) & CmainT(k, 1) & BindexT(j, 12)
                                End If
                            ElseIf FLGmaker = "998" And Bsitei = 10 Then
                                If BindexT(j, 14) = "*" Then
                                    FLGpname = BindexT(j, 13) & CmainT(k, 1)
                                Else
                                    FLGpname = BindexT(j, 13) & CmainT(k, 1) & BindexT(j, 14)
                                End If
                            Else
                                If BindexT(j, 4) = "*" Then
                                    FLGpname = BindexT(j, 3) & CmainT(k, 1)
                                Else
                                    FLGpname = BindexT(j, 3) & CmainT(k, 1) & BindexT(j, 4)
                                End If
                            End If
'
                            If FLGmaker = "000" Then
                                FLGmaker = CmainT(k, 13)
                            ElseIf FLGmaker = "998" And Bsitei <> 0 Then
                                FLGmaker = BindexT(j, Bsitei)
                            ElseIf FLGmaker = "998" And Bsitei = 0 Then
                                FLGpname = FLGpname & "相当"
                            End If
'
                            If CmainT(k, 3) = "0" Then   '*** 部品指定表示 ***
                                FLGpname = "!" & FLGpname
                                lblTyuui.Visible = True
                            ElseIf CmainT(k, 3) = "3" Then
                                FLGpname = "?" & FLGpname
                                lblTyuui.Visible = True
                            ElseIf CmainT(k, 3) = "4" Then
                                FLGpname = "\" & FLGpname
                                lblTyuui.Visible = True
                            Else
                                FLGpname = " " & FLGpname
                            End If
'
                            FLGtanka = Val(CmainT(k, 5))
                            FLGkeijou = CmainT(k, 18)
'
                            FLGondo = CmainT(k, 11)
                            If Len(FLGondo) < 3 Then FLGondo = "     "  '***3文字未満は無い
'
                            Dtemp = CmainT(k, 6)        'MSL表示
                            Call TRS_Mlevel2(Dtemp)
                            FLGondo = Dtemp & "/" & FLGondo
'
                            If 5 < Len(CmainT(k, 19)) Then    '*** ﾒｯｷ記入
                                FLGmekki = CmainT(k, 19)
                            ElseIf 5 = Len(CmainT(k, 19)) Then
                                FLGmekki = " " & CmainT(k, 19)
                            ElseIf 4 = Len(CmainT(k, 19)) Then
                                FLGmekki = " " & CmainT(k, 19) & " "
                            ElseIf 3 = Len(CmainT(k, 19)) Then
                                FLGmekki = "  " & CmainT(k, 19) & " "
                            ElseIf 2 = Len(CmainT(k, 19)) Then
                                FLGmekki = "  " & CmainT(k, 19) & "  "
                            Else
                                FLGmekki = "      "      '*** 1文字の金属メッキは無い ***
                            End If
'
                            If InStr(CmainT(k, 19), "SnPb") <> 0 Then
                                FLGmekki = FLGmekki & "/----"
                            ElseIf InStr(CmainT(k, 2), "#<RoHS>") <> 0 Then
                                FLGmekki = FLGmekki & "/=>>="
                            ElseIf InStr(CmainT(k, 2), "<RoHS>") <> 0 Then
                                FLGmekki = FLGmekki & "/RoHS"
                            ElseIf InStr(CmainT(k, 2), "#<Pbﾌﾘｰ>") <> 0 Then
                                FLGmekki = FLGmekki & "/=>>="
                            ElseIf InStr(CmainT(k, 2), "<Pbﾌﾘｰ>") <> 0 Then
                                FLGmekki = FLGmekki & "/NoPb"
                            ElseIf InStr(CmainT(k, 2), "<Green>") <> 0 Then
                                FLGmekki = FLGmekki & "/Gree"
                            ElseIf InStr(CmainT(k, 2), "<Ro2>") <> 0 Then   '*** Ver2.1ﾆﾃ追加
                                FLGmekki = FLGmekki & "/Ro2"
                            ElseIf InStr(CmainT(k, 2), "<R863>") <> 0 Then  '*** Ver2.3ﾆﾃ追加
                                FLGmekki = FLGmekki & "/R863"
                            ElseIf InStr(BindexT(j, 1), "#<RoHS>") <> 0 Then
                                FLGmekki = FLGmekki & "/=>>="
                            ElseIf InStr(BindexT(j, 1), "<RoHS>") <> 0 Then
                                FLGmekki = FLGmekki & "/RoHS"
                            ElseIf InStr(BindexT(j, 1), "#<Pbﾌﾘｰ>") <> 0 Then
                                FLGmekki = FLGmekki & "/=>>="
                            ElseIf InStr(BindexT(j, 1), "<Pbﾌﾘｰ>") <> 0 Then
                                FLGmekki = FLGmekki & "/NoPb"
                            ElseIf InStr(BindexT(j, 1), "<Green>") <> 0 Then
                                FLGmekki = FLGmekki & "/Gree"
                            ElseIf InStr(BindexT(j, 1), "<Ro2>") <> 0 Then  '*** Ver2.1ﾆﾃ追加
                                FLGmekki = FLGmekki & "/Ro2"
                            ElseIf InStr(BindexT(j, 1), "<R863>") <> 0 Then '*** Ver2.3ﾆﾃ追加
                                FLGmekki = FLGmekki & "/R863"
                            End If
'
                            If goukei_en = True Then
                                Soukei = Soukei + FLGtanka                  '*** 単価合計 ***
                                Pingoukei = Pingoukei + Val(CmainT(k, 21))  '*** ピン数合計 ***
                            End If
'
                            Makerget2 FLGmaker
                            Exit Sub
'
                        End If
                    Next k
                End If
            Next j
        End If
    Next i
'
    FLGpname = "!!! このｺｰﾄﾞ番号は見つかり見つかりません。 !!!"
    FLGmaker = "  ???"
End Sub

Private Sub H1down()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Ptotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf (Ptotal + 1) - j >= G_row Then
        MSFlexGrid1.TopRow = j + 1
    Else
        Beep
    End If
End Sub

Private Sub H1up()
    Dim j As Integer
    j = MSFlexGrid1.TopRow
'
    If j = 1 Then
        Beep
    ElseIf Ptotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 1
    End If
End Sub

Private Sub Hdown()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Ptotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf (Ptotal + 1) - (j + 10) >= G_row Then
        MSFlexGrid1.TopRow = j + 10
    Else
        Beep
        MSFlexGrid1.TopRow = (Ptotal + 1) - G_row + 1
    End If
End Sub

Private Sub Hup()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Ptotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j - 10 <= 0 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 10
    End If
End Sub

Private Sub set_Pname()
    Dim i As Integer
'
    lblTyuui.Visible = False
'
    For i = 1 To Ptotal
        Call set_G_line(i)  '*** １行表示 ***
    Next i
End Sub

Private Sub setMSFlexGrid1()
    Dim i As Integer, iend As Integer
'
    G_row = 20      '*** グリッドのデータ表示行数 ***
'
    If Ptotal < G_row Then
        MSFlexGrid1.Rows = G_row + 1
    Else
        MSFlexGrid1.Rows = Ptotal + 2
    End If
'                       *** 表題 ***
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Text = "行"
        .ColAlignment(0) = flexAlignCenterCenter    '*** 中央 ***
        .Col = 1
        .Text = "項目"
        .ColAlignment(1) = flexAlignCenterCenter    '*** 中央 ***
        .Col = 2
        .Text = "部品番号"
        .ColAlignment(2) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 3
        .Text = "  部品ｺｰﾄﾞ"
        .ColAlignment(3) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 4
        .Text = "                部 品 名 称"
        .ColAlignment(4) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 5
        .Text = "ﾒｰｶｰ名"
        .ColAlignment(5) = flexAlignCenterCenter    '*** 中央 ***
        .Col = 6
        .Text = "         備  考"
        .ColAlignment(6) = flexAlignLeftCenter      '*** 左詰め ***
'
            If optTanka.Value = True Then
        .Col = 7
        .Text = "単価   "
        .ColAlignment(7) = flexAlignRightCenter     '*** 右詰め ***
        .Col = 8
        .Text = "形状"
        .ColAlignment(8) = flexAlignCenterCenter    '*** 中央 ***
            Else
        .Col = 7
        .Text = "MSL/耐熱"
        .ColAlignment(7) = flexAlignCenterCenter    '*** 中央 ***
        .Col = 8
        .Text = " ﾒｯｷ/RoHS"
        .ColAlignment(8) = flexAlignCenterCenter    '*** 中央 ***
            End If
    End With
'
'               *** 表示内容クリアー ***
    For i = Ptotal + 1 To MSFlexGrid1.Rows - 1
        With MSFlexGrid1
            .Row = i
            .Col = 0
            .Text = ""
            .Col = 1
            .Text = ""
            .Col = 2
            .Text = ""
            .Col = 3
            .Text = ""
            .Col = 4
            .Text = ""
            .Col = 5
            .Text = ""
            .Col = 6
            .Text = ""
            .Col = 7
            .Text = ""
            .Col = 8
            .Text = ""
        End With
    Next i
End Sub

Private Sub setFormArea()   '*** フォームの表示位置の設定 ***
'    If Eeos2_mainMDI.ScaleHeight > Height Then
'        Top = Eeos2_mainMDI.ScaleHeight - Height
'    Else
        Me.Top = 420
'    End If
'
    If Eeos2_mainMDI.ScaleWidth > Me.Width Then
        Me.Left = Eeos2_mainMDI.ScaleWidth - Me.Width   '*** 右端に表示 ***
    Else
        Me.Left = 0
    End If
End Sub

Private Sub set_G_line(i As Integer)
    Dim Ddata As String, FLGmitou As Integer
    Dim FLGkoumoku As String, FLGpname As String, FLGmaker As String
    Dim Bsitei As Integer, FLGtanka As Currency, FLGkeijou As String
    Dim FLGondo As String, FLGmekki As String
'
    With MSFlexGrid1
        .Row = i
        .Col = 0                '*** 行番号 ***
        .Text = Trim(str(i))
'
        .Col = 1                '*** 項目名 ***
            FLGkoumoku = PLST(i, Pdim0 + 1)
            Call GET_koumoku(FLGkoumoku, Aitem0(), Anum0)
        .Text = " " & FLGkoumoku
'
        .Col = 2                '*** 部品番号 ***
            Ddata = PLST(i, 0)
            Call CHK7moji(Ddata)
        .Text = Ddata
'
        .Col = 3                '*** 部品ｺｰﾄﾞ ***
            Ddata = PLST(i, 1)
            Call CHKmitou(Ddata, FLGmitou)
        .Text = Ddata
'
        .Col = 4                '*** 部品名称 ***
            If FLGmitou = 1 Then
                FLGpname = " " & Mid(PLST(i, 1), 2)
                FLGtanka = 0
                FLGmaker = "******"
                FLGkeijou = "*"
                FLGondo = "*"
                FLGmekki = "*"
            Else
                FLGpname = PLST(i, 1)
                Bsitei = Val(PLST(i, 3))
                Call GETpname(FLGkoumoku, FLGpname, FLGmaker, Bsitei, FLGtanka, FLGkeijou, FLGondo, FLGmekki)
'
                If PLST(i, 4) = "*" Or PLST(i, 4) = "" Then     '*** 特記事項 ***
                '
                Else
                    FLGpname = FLGpname & PLST(i, 4)
                End If
            End If
        .Text = FLGpname
'
        .Col = 5                '*** ﾒｰｶｰ略称 ***
        .Text = FLGmaker
'
        .Col = 6                '*** 備考 ***
            Ddata = " " & PLST(i, 2)
        .Text = Ddata
'
            If optTanka.Value = True Then
        .Col = 7                '*** 平均単価/部品形状 ***
        .Text = Format(FLGtanka, "#,###,##0.0")
'
        .Col = 8
            Call TRSkeijou2(FLGkeijou)
        .Text = FLGkeijou
'
            Else
        .Col = 7                '*** 耐熱温度/ﾒｯｷ材料･RoHS対応状況 ***
        .Text = FLGondo
'
        .Col = 8
        .Text = FLGmekki
            End If
    End With
End Sub

Private Sub cmdHelp_Click()
    Call mnuSetumei_Click
End Sub

Private Sub cmdAppend_Click()
    Call Copy_plst_work2temp
'
    Plst_append.Show 1
'
    Call Form_Activate
End Sub

Private Sub cmdDown_Click()
    Call Hdown      '*** 下へ ***
End Sub

Private Sub cmdQuit_Click()
    Dim i As Integer
'
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbYesNo, HeadTitle)
        If i = vbNo Then
            FLGjob = 0
            FLGlevel = 0
            Unload Me
'
        Else
            Call Copy_plst_work2temp
'
            Plst_DirW.Show 1
'
            If FLGesc = 0 Then
                txtZuban.Text = PFLname
                FLGjob = 0
                FLGlevel = 0
                Unload Me
'
            End If
        End If
    Else
        Unload Me
'
    End If
End Sub

Private Sub cmdUp_Click()
    Call Hup        '*** 上へ ***
End Sub

Private Sub mnuAQuitP_Click()
    Call mnuAQuit_Click
End Sub

Private Sub mnuBackP_Click()
    Call cmdQuit_Click
End Sub

Private Sub mnuBuhin_Click()
    Call mnuDenkiBuhinhyou
End Sub

Private Sub mnuBuhinP_Click()
    Call mnuBuhin_Click
End Sub

Private Sub mnuBuhin2_Click()
    Call mnuDenkiBuhinhyou2
End Sub

Private Sub mnuBuhin2P_Click()
    Call mnuBuhin2_Click
End Sub

Private Sub mnuBuhinPRN_Click()
    Call mnuStandardBuhinhyouPrint
End Sub

Private Sub mnuBuhinPRNP_Click()
    Call mnuBuhinPRN_Click
End Sub

Private Sub mnuCode_Click()
    Call mnuCodeBuhinMaintenance
End Sub

Private Sub mnuCodeP_Click()
    Call mnuCode_Click
End Sub

Private Sub mnufileNR_Click()
    Dim i As Integer
'
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbQuestion Or vbYesNo, STATUS)
        If i = vbYes Then
            Exit Sub
'
        End If
    End If
'
    Plst_DirR.Show 1
'
    If FLGesc = 1 Then      '*** エスケープフラグセット ***
        Exit Sub
'
    End If
'
    DRVpartlist = DRVpartlistT
    PFLname = PFLnameT
'
    Me.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
    Call RDpartlist(DRVpartlist, Plistname, Plistdate, Remarks, PLST(), Ptotal, Pdim0)      '*** 部品表読み込み ***
    FLGchange = 0
'
'    Call RDplstWork     '*** 部品表作成データ読み込み ***
'
    For i = 1 To Ptotal
        Pitem = PLST(i, 0)
        Call GETsymbol(Pmodu, Pitem, Pno) '*** 項目と番号を得る ***
        PLST(i, Pdim0 + 1) = Pitem
        PLST(i, Pdim0 + 2) = Pno
        PLST(i, Pdim0 + 3) = Pmodu
    Next i
'
    Call Gamen_settei   '*** 画面設定と内容表示 ***
'
    icpsT = 0           '*** MAINコードメモリー クリアー ***
'
    MSFlexGrid1.MousePointer = flexDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub mnufileNRP_Click()
    Call mnufileNR_Click
End Sub

Private Sub mnufileNW_Click()
    Call Copy_plst_work2temp
'
    Plst_DirW.Show 1
'
    If FLGesc = 0 Then
        DRVpartlist = DRVpartlistT
        PFLname = PFLnameT
        txtZuban.Text = PFLname
        FLGlevel = 1    '*** 部品表 更新・確認 ***
    End If
End Sub

Private Sub mnufileNWP_Click()
    Call mnufileNW_Click
End Sub

Private Sub mnuFilePrn_Click()
    FLGfile = 0         '*** 印刷 ***
    FLGall = 0          '*** 個別 ***
    DRVpartlistT = DRVpartlist
    PFLnameT = PFLname
'
    Plst_PRNlst.Show 1  '*** 一覧表印刷/ファイル出力 ***
'
End Sub

Private Sub mnuFilePrnA_Click()
    Call mnuBuhinItiranhyouPrint
End Sub

Private Sub mnuFilePrnAP_Click()
    Call mnuFilePrnA_Click
End Sub

Private Sub mnufileRD_Click()
    Dim i As Integer
'
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbQuestion Or vbYesNo, STATUS)
        If i = vbYes Then
            Exit Sub
'
        End If
    End If
'
    Me.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
    Call RDpartlist(DRVpartlist, Plistname, Plistdate, Remarks, PLST(), Ptotal, Pdim0)  '*** 部品表読み込み ***
    FLGchange = 0
'
'    Call RDplstWork     '*** 部品表作成データ読み込み ***
'
    For i = 1 To Ptotal
        Pitem = PLST(i, 0)
        Call GETsymbol(Pmodu, Pitem, Pno) '*** 項目と番号を得る ***
        PLST(i, Pdim0 + 1) = Pitem
        PLST(i, Pdim0 + 2) = Pno
        PLST(i, Pdim0 + 3) = Pmodu
    Next i
'
    Call Gamen_settei   '*** 画面設定と内容表示 ***
'
    icpsT = 0           '*** MAINコードメモリー クリアー ***
'
    MSFlexGrid1.MousePointer = flexDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub mnufileRDP_Click()
    Call mnufileRD_Click
End Sub

Private Sub mnufileSV_Click()
    If FLGlevel = 3 Then   '*** ファイル名未定 ***
        Call Copy_plst_work2temp
'
        Plst_DirW.Show 1
'
        If FLGesc = 0 Then
            DRVpartlist = DRVpartlistT
            PFLname = PFLnameT
            txtZuban.Text = PFLname
        End If
    Else
        Me.MousePointer = vbHourglass
        MSFlexGrid1.MousePointer = flexHourglass
'
        Call WRpartlist(DRVpartlist, Plistname, Plistdate, Remarks, PLST(), Ptotal, Pdim0)  '*** 部品表セーブ ***
        FLGchange = 0
'
'        Call WRplstWork      '*** 部品表作成データ書き込み ***
'
        MSFlexGrid1.MousePointer = flexDefault
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub mnufileSVP_Click()
    Call mnufileSV_Click
End Sub

Private Sub mnuFileWR_Click()
    FLGfile = 1         '*** ファイル出力 ***
    FLGall = 0          '*** 個別 ***
    DRVpartlistT = DRVpartlist
    PFLnameT = PFLname
'
    Plst_PRNlst.Show 1  '*** 一覧表印刷/ファイル出力 ***
'
End Sub

Private Sub mnuHinsyu_Click()
    Call mnuCodeHinsyuMaintenance
End Sub

Private Sub mnuHinsyuP_Click()
    Call mnuHinsyu_Click
End Sub

Private Sub mnuJumpBP_Click()
    Call mnuJumpB_Click
End Sub

Private Sub mnuJumpC_Click()
    Dim Tyuusin As Integer
'
    Tyuusin = Ptotal \ 2
    If Tyuusin <= G_row \ 2 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Tyuusin - G_row \ 2
    End If
End Sub

Private Sub mnuJumpCP_Click()
    Call mnuJumpC_Click
End Sub

Private Sub mnuJumpEP_Click()
    Call mnuJumpE_Click
End Sub

Private Sub mnuJumpNP_Click()
    Call mnuJumpN_Click
End Sub

Private Sub mnuJumpTP_Click()
    Call mnuJumpT_Click
End Sub

Private Sub mnuKaihan_Click()
    Call mnuKaihanRireki
End Sub

Private Sub mnuKousei_Click()
    Call mnuDenkiKouseihyou
End Sub

Private Sub mnuKouseiP_Click()
    Call mnuKousei_Click
End Sub

Private Sub mnuMakerment_Click()
    Call mnuCodeMakerMaintenance
End Sub

Private Sub mnuMakermentP_Click()
    Call mnuMakerment_Click
End Sub

Private Sub mnuOption_Click()
    Call mnuOptionSettei
End Sub

Private Sub mnuORCAD_Click()
    Call mnuOrCAD_Henkan
End Sub

Private Sub mnuORCADP_Click()
    Call mnuORCAD_Click
End Sub

Private Sub mnuConvFile_Click()
    Call mnuConvFile_Edit
End Sub

Private Sub mnuConvFileP_Click()
    Call mnuConvFile_Edit
End Sub

Private Sub mnuPmain_Click()
    Call mnuCodePmainMaintenance
End Sub

Private Sub mnuPmainP_Click()
    Call mnuPmain_Click
End Sub

Private Sub mnuPRNstd_Click()
    FLGfile = 0         '*** 印刷 ***
    FLGall = 0          '*** 個別 ***
    STATUS = "標準部品表印刷"
    DRVpartlistT = DRVpartlist
    PFLnameT = PFLname
'
    Plst_PRNstd.Show 1  '*** 標準部品表印刷/ファイル出力 ***
'
End Sub

Private Sub mnuReform_Click()
    Width = tempWidth
    Height = tempHeight '*** これで「Form_Resize」割り込みが発生する。 ***
'
    Call setFormArea    '*** フォームの表示位置の設定 ***
End Sub

Private Sub mnuSettei_Click()
    Call mnuKankyouSettei
End Sub

Private Sub mnuSinki_Click()
    Dim i As Integer
'
    i = vbYes
    If FLGchange = 1 Then
        Beep
        i = MsgBox("変更内容を放棄しますか？", vbQuestion Or vbYesNo, STATUS)
    End If
'
    If i = vbYes Then
        Ptotal = 1
        ReDim PLST(Ptotal, Pdim0 + 3)
'
        PFLname = "BSHINKI.PLT"
        Plistname = "*"
        Plistdate = Format(Date, "yyyy/mm/dd")
        Remarks = "*"
'
        PLST(1, 0) = "Z1"
        PLST(1, 1) = "*削除してね"
        PLST(1, 2) = "*"
        PLST(1, 3) = "0"
        PLST(1, 4) = "*"
        PLST(1, Pdim0 + 1) = "Z"
        PLST(1, Pdim0 + 2) = "1"
        PLST(1, Pdim0 + 3) = ""
'
        FLGchange = 0
        icpsT = 0           '*** コード項目番号初期化 ***
        FLGlevel = 3        '*** 新規作成 ***
'
'        Call RDplstWork     '*** 部品表作成データ読み込み ***
'
        Call Gamen_settei   '*** 画面設定 内容表示 ***
    End If
End Sub

Private Sub mnuSinkiP_Click()
    Call mnuSinki_Click
End Sub

Private Sub mnuSuuryo_Click()
    Call mnuBuhinSuuryohyouPrint
End Sub

Private Sub mnuSuuryoP_Click()
    Call mnuSuuryo_Click
End Sub

Private Sub mnuTileC_Click()
    Eeos2_mainMDI.Arrange vbCascade         '*** 重ねて表示 ***
End Sub

Private Sub mnuTileH_Click()
    Eeos2_mainMDI.Arrange vbTileHorizontal  '*** 並べて表示 ***
End Sub

Private Sub mnuTileV_Click()
    Eeos2_mainMDI.Arrange vbTileVertical    '*** 並べて表示 ***
End Sub

Private Sub mnuTraderment_Click()
    Call mnuCodeTraderMaintenance
End Sub

Private Sub mnuTradermentP_Click()
    Call mnuTraderment_Click
End Sub

Private Sub mnuAQuit_Click()
    Dim i As Integer
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbQuestion Or vbYesNo, STATUS)
        If i = vbNo Then
            Unload Me
            End
'
        End If
'
    Else
        Unload Me
        End
'
    End If
End Sub

Private Sub mnuBack_Click()
    Call cmdQuit_Click
End Sub

Private Sub mnuJumpB_Click()
    Dim Adata As String
    Dim i As Integer, j As Integer
'
    j = MSFlexGrid1.TopRow
    Adata = PLST(j, Pdim0 + 1)
    For i = j To 1 Step -1
        If Adata <> PLST(i, Pdim0 + 1) Then
            j = i
            Exit For
'
        End If
    Next i
'
    Adata = PLST(j, Pdim0 + 1)
    For i = j To 1 Step -1
        If Adata <> PLST(i, Pdim0 + 1) Then
            j = i + 1
            Exit For
'
        End If
    Next i
'
    If i = 0 Then j = 1     '*** 行の先頭 **
'
    If Ptotal < G_row Then
        MSFlexGrid1.TopRow = 1
    ElseIf (Ptotal + 1) - j >= G_row Then
        MSFlexGrid1.TopRow = j
    Else
        MSFlexGrid1.TopRow = (Ptotal + 1) - G_row + 1
    End If
End Sub

Private Sub mnuJumpE_Click()
    If Ptotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Ptotal - G_row + 1
    End If
End Sub

Private Sub mnuJumpN_Click()
    Dim Adata As String
    Dim i As Integer, j As Integer
'
    j = MSFlexGrid1.TopRow
    Adata = PLST(j, Pdim0 + 1)
    For i = j To Ptotal
        If Adata <> PLST(i, Pdim0 + 1) Then
            j = i
            Exit For
'
        End If
    Next i
'
    If Ptotal < G_row Then
        MSFlexGrid1.TopRow = 1
    ElseIf (Ptotal + 1) - j >= G_row Then
        MSFlexGrid1.TopRow = j
    Else
        MSFlexGrid1.TopRow = (Ptotal + 1) - G_row + 1
    End If
End Sub

Private Sub mnuJumpT_Click()
    If Ptotal <= G_row Then
        Beep
    End If
'
    MSFlexGrid1.TopRow = 1
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 2          '*** 部品表フラグ ***
    Call mnuSousaSetumei
End Sub

Private Sub mnuVersion_Click()
    Call mnuVersionGamen
End Sub

Private Sub MSFlexGrid1_Click()
    txtKiji.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
    Dim Ginput As String
    Dim GinData As String, GmsgData As String, TMPcode As String
    Dim i As Integer
'
    MpointC = MSFlexGrid1.Col     '*** 列 ***
    MpointR = MSFlexGrid1.Row     '*** 行 ***
'
    If MpointR > Ptotal Then    '*** ptotalより大きい時はエラー ***
        Beep
        Exit Sub
    End If
'
    If MpointC = 0 Or MpointC = 1 Or MpointC = 2 Then '*** 削除 ***
        Beep
        MSFlexGrid1.Col = 2
        Gdata2(0) = MSFlexGrid1.Text
        MSFlexGrid1.Col = 4
        Gdata4(0) = MSFlexGrid1.Text
        GmsgData = "[" & str(MpointR) & " 行] < " & Gdata2(0) & " : " & Gdata4(0) & " > を _削除_ します。"
        i = MsgBox(GmsgData, vbExclamation Or vbOKCancel, STATUS)
        If i = vbOK Then
            Call DEL1gyou(MpointR)
            FLGchange = 1
        End If
    ElseIf MpointC = 3 Or MpointC = 4 Then      '*** 変更 ***
        Beep
        MSFlexGrid1.Col = 1
        Gdata1(0) = MSFlexGrid1.Text
        MSFlexGrid1.Col = 2
        Gdata2(0) = MSFlexGrid1.Text
        MSFlexGrid1.Col = 3
        Gdata3(0) = Trim(MSFlexGrid1.Text)
        MSFlexGrid1.Col = 4
        Gdata4(0) = MSFlexGrid1.Text
        Gdata5(0) = PLST(MpointR, 3)   '*** 部品指定 ***
        Gdata6(0) = PLST(MpointR, 4)   '*** 特記事項 ***
        Gdata7(0) = PLST(MpointR, 2)   '*** 備考欄 ***
'
        TMPcode = Gdata3(0)        '*** 元のコード番号 ***
        Plst_main_c.Show 1
'
        If FLGesc = 0 Then
            FLGchange = 1
'
            PLST(MpointR, 1) = Gdata3(0)   '*** 部品コード ***
            PLST(MpointR, 3) = Gdata5(0)   '*** 部品指定 ***
            PLST(MpointR, 4) = Gdata6(0)   '*** 特記事項 ***
            PLST(MpointR, 2) = Gdata7(0)   '*** 備考欄 ***
'
            Call set_G_line(MpointR)
'
            If FLGsubete = 1 Then   '*** すべて一緒に変更の時 ***
                For i = 1 To Ptotal
                    If PLST(i, 1) = TMPcode Then
                        PLST(i, 1) = Gdata3(0)
                        PLST(i, 3) = Gdata5(0)
                        PLST(i, 4) = Gdata6(0)
                        PLST(i, 2) = Gdata7(0)
'
                        Call set_G_line(i)
                    End If
                Next i
            End If
        End If
    ElseIf MpointC = 6 Then         '*** 備考欄変更 ***
        Beep
        MSFlexGrid1.Col = 2
        Gdata1(0) = MSFlexGrid1.Text
        MSFlexGrid1.Col = 4
        Gdata2(0) = MSFlexGrid1.Text
        MSFlexGrid1.Col = 6
        Gdata3(0) = Trim(MSFlexGrid1.Text)
        GinData = "< " & Gdata1(0) & " : " & Gdata2(0) & " > の備考欄を /変更/ します。" _
            & vbCrLf & vbCrLf & " <<< データーを入力してください。 >>>"
        Ginput = InputBox(GinData, STATUS, Gdata3(0))
        If Ginput <> "" Then
            PLST(MpointR, 2) = Ginput
            MSFlexGrid1.Text = " " & Ginput
            FLGchange = 1
        End If
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub optOndo_Click()
    If FlagBusy = True Then Exit Sub
'
    FlagBusy = True
    optTanka.Value = False
    optOndo.Value = True
    FlagBusy = False
'
    Call setMSFlexGrid1     '*** グリッドの諸元設定 ***
    Call set_Pname          '*** 部品名称の表示 ***
End Sub

Private Sub optTanka_Click()
    If FlagBusy = True Then Exit Sub
'
    FlagBusy = True
    optTanka.Value = True
    optOndo.Value = False
    FlagBusy = False
'
    Call setMSFlexGrid1     '*** グリッドの諸元設定 ***
    Call set_Pname          '*** 部品名称の表示 ***
End Sub

Private Sub txtDate_Click()
    txtDate.MousePointer = vbIbeam
End Sub

Private Sub txtDate_LostFocus()
    Dim Ddata As String
'
    Ddata = Trim(txtDate.Text)
    If Ddata <> Trim(Plistdate) Then
        Plistdate = Ddata
        FLGchange = 1
    End If
'
    txtDate.MousePointer = vbArrow
End Sub

Private Sub txtKiji_Click()
    txtKiji.MousePointer = vbIbeam
End Sub

Private Sub txtKiji_LostFocus()
    Dim Ddata As String
'
    Ddata = Trim(txtKiji.Text)
    If Ddata <> Trim(Remarks) Then
        Remarks = Ddata
        FLGchange = 1
    End If
'
    txtKiji.MousePointer = vbArrow
End Sub

Private Sub txtKmeisyou_Click()
    txtKmeisyou.MousePointer = vbIbeam
End Sub

Private Sub txtKmeisyou_LostFocus()
    Dim Ddata As String
'
    Ddata = Trim(txtKmeisyou.Text)
    If Ddata <> Trim(Plistname) Then
        Plistname = Ddata
        FLGchange = 1
    End If
'
    txtKmeisyou.MousePointer = vbArrow
End Sub

Private Sub MENU_settei()       '*** メニュー状態設定 ***
'
    If FLGconst = 1 Then       '*** 構成表画面存在 ***
        Me.mnuKousei.Checked = True
        Me.mnuKouseiP.Checked = True
    Else
        Me.mnuKousei.Checked = False
        Me.mnuKouseiP.Checked = False
    End If
'
    If FLGplst = 1 Then         '*** 部品表画面存在 ***
        Me.mnuBuhin.Checked = True
        Me.mnuBuhinP.Checked = True
    Else
        Me.mnuBuhin.Checked = False
        Me.mnuBuhinP.Checked = False
    End If
'
    If FLGplst2 = 1 Then        '*** 部品表画面２存在 ***
        Me.mnuBuhin2.Checked = True
        Me.mnuBuhin2P.Checked = True
    Else
        Me.mnuBuhin2.Checked = False
        Me.mnuBuhin2P.Checked = False
    End If
'
    If FLGplst = 1 And FLGplst2 = 1 Then    '*** 部品表２画面とも既に開いている ***
        Me.mnuORCAD.Enabled = False
        Me.mnuORCADP.Enabled = False
    Else
        Me.mnuORCAD.Enabled = True
        Me.mnuORCADP.Enabled = True
    End If
'
    If FLGplstWork = 1 Then        '*** OrCAD変換作業ファイル 編集画面存在 ***
        Me.mnuConvFile.Checked = True
        Me.mnuConvFileP.Checked = True
    Else
        Me.mnuConvFile.Checked = False
        Me.mnuConvFileP.Checked = False
    End If
'
    If FLGmaker = 1 Then        '*** メーカー画面存在 ***
        Me.mnuMakerment.Checked = True
        Me.mnuMakermentP.Checked = True
    Else
        Me.mnuMakerment.Checked = False
        Me.mnuMakermentP.Checked = False
    End If
'
    If FLGtrader = 1 Then       '*** 商社画面存在 ***
        Me.mnuTraderment.Checked = True
        Me.mnuTradermentP.Checked = True
    Else
        Me.mnuTraderment.Checked = False
        Me.mnuTradermentP.Checked = False
    End If
'
    If FLGitem = 1 Then         '*** 部品コード項目画面存在 ***
        Me.mnuCode.Checked = True
        Me.mnuHinsyu.Enabled = True
'
        Me.mnuCodeP.Checked = True
        Me.mnuHinsyuP.Enabled = True
'
        If FLGindex = 1 Then
            Me.mnuHinsyu.Checked = True
            Me.mnuPmain.Enabled = True
'
            Me.mnuHinsyuP.Checked = True
            Me.mnuPmainP.Enabled = True
'
            If FLGmain = 1 Then
                Me.mnuPmain.Checked = True
                Me.mnuPmainP.Checked = True
            Else
                Me.mnuPmain.Checked = False
                Me.mnuPmainP.Checked = False
            End If
        Else
            Me.mnuHinsyu.Checked = False
            Me.mnuPmain.Enabled = False
            Me.mnuPmain.Checked = False
'
            Me.mnuHinsyuP.Checked = False
            Me.mnuPmainP.Enabled = False
            Me.mnuPmainP.Checked = False
        End If
    Else
        Me.mnuCode.Checked = False
        Me.mnuCode.Enabled = True
        Me.mnuHinsyu.Checked = False
        Me.mnuHinsyu.Enabled = False
        Me.mnuPmain.Checked = False
        Me.mnuPmain.Enabled = False
'
        Me.mnuCodeP.Checked = False
        Me.mnuCodeP.Enabled = True
        Me.mnuHinsyuP.Checked = False
        Me.mnuHinsyuP.Enabled = False
        Me.mnuPmainP.Checked = False
        Me.mnuPmainP.Enabled = False
    End If
'
    Call MENU_shinki
End Sub

Private Sub MENU_shinki()
    If FLGshinki = 1 Then
        Me.mnufileSV.Enabled = False
        Me.mnufileRD.Enabled = False
'
        Me.mnufileSVP.Enabled = False
        Me.mnufileRDP.Enabled = False
    Else
        Me.mnufileSV.Enabled = True
        Me.mnufileRD.Enabled = True
'
        Me.mnufileSVP.Enabled = True
        Me.mnufileRDP.Enabled = True
    End If
End Sub
