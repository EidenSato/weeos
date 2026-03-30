VERSION 5.00
Begin VB.MDIForm Eeos2_mainMDI 
   BackColor       =   &H8000000C&
   Caption         =   "ＥEＯＳ２"
   ClientHeight    =   3060
   ClientLeft      =   165
   ClientTop       =   300
   ClientWidth     =   7500
   Icon            =   "Eeos2_mainMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   360
   End
   Begin VB.Menu mnuFileA 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuAllQuit 
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
      Begin VB.Menu mnu区切り線21 
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
      Begin VB.Menu mnuCodement 
         Caption         =   "部品ｺｰﾄﾞ表(&C)"
      End
      Begin VB.Menu mnuMakerment 
         Caption         =   "ﾒｰｶｰｺｰﾄﾞ表(&M)"
      End
      Begin VB.Menu mnuTraderment 
         Caption         =   "商社ｺｰﾄﾞ表(&T)"
      End
   End
   Begin VB.Menu mnuKankyou 
      Caption         =   "環境(&O)"
      Begin VB.Menu mnuSettei 
         Caption         =   "環境設定(&K)"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "ｵﾌﾟｼｮﾝ(&O)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuSetumei 
         Caption         =   "操作説明(&S)"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "改版履歴(&H)"
      End
      Begin VB.Menu mnu区切り線81 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ情報(&V)"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ﾎﾟｯﾌﾟｱｯﾌﾟﾒﾆｭｰ"
      Visible         =   0   'False
      Begin VB.Menu mnuKouseiP 
         Caption         =   "電気 構成表..."
      End
      Begin VB.Menu mnu区切り線90 
         Caption         =   "-"
      End
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
      Begin VB.Menu mnu区切り線91 
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
      Begin VB.Menu mnu区切り線92 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodementP 
         Caption         =   "部品ｺｰﾄﾞ表"
      End
      Begin VB.Menu mnuMakermentP 
         Caption         =   "ﾒｰｶｰｺｰﾄﾞ表"
      End
      Begin VB.Menu mnuTradermentP 
         Caption         =   "商社ｺｰﾄﾞ表"
      End
      Begin VB.Menu mnu区切り線93 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAQuitP 
         Caption         =   "EEOS２の終了"
      End
   End
End
Attribute VB_Name = "Eeos2_mainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*****************************
'*** ＥＥＯＳ２ メイン画面 ***
'*****************************
'
Option Explicit
'

Private Sub MDIForm_Initialize()
    Eeos2_mainMDI.Caption = EEOS_STATUS
'
    FLGconst = 0    '*** 構成表画面存在フラグ初期化 ***
    FLGplst = 0     '*** 部品表画面存在フラグ初期化 ***
    FLGlevel = 0    '*** 部品表 作業フラグ初期化 ***
    FLGplstWork = 0 '*** 部品表作成データ編集画面存在フラグ ***
    FLGitem = 0     '*** 部品コード項目画面存在フラグ初期化 ***
    FLGindex = 0    '*** 部品コード品種画面存在フラグ初期化 ***
    FLGmain = 0     '*** 部品コード品目画面存在フラグ初期化 ***
    FLGmaker = 0    '*** メーカー画面存在フラグ初期化 ***
    FLGtrader = 0   '*** 商社画面存在フラグ初期化 ***
'
    FLG_job_error_end = 0 '*** フラグの初期化 ***
End Sub

Private Sub MDIForm_Load()
    Dim i  As Integer
    Dim FILE_number As Integer
'
    Width = 12060           '*** 画面サイズ、表示位置設定 ***
    Height = 8700
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - 380 - Height) \ 2
'
    Call RDcont          '*** 環境設定/オプション読み込み ***
'
    If Xcont0(15) = "2" Then
        HyoujiBairitu = cUHBairitu
    ElseIf Xcont0(15) = "1" Then
        HyoujiBairitu = cHBairitu
    Else
        HyoujiBairitu = 1#
    End If
'
    If FLG_job_error_end = 1 Then
        Kankyow_Itiran.Show 1
        FLG_job_error_end = 0
    End If
'
    Timer1.Interval = 1     '*** 1ms ***
    Timer1.Enabled = True
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
'
    Select Case Xcont0(1)
    Case "1":
        Call mnuKousei_Click        '*** 構成表一覧表 ***
    Case "2":
        Call mnuBuhin_Click         '*** 部品表一覧表 ***
    Case "3":
        Call mnuCodement_Click      '*** 部品コード項目一覧表 ***
    Case "4":
        Call mnuMakerment_Click     '*** メーカーコード表 ***
    Case "5":
        Call mnuTraderment_Click    '*** 商社コード表 ***
    Case "6":
        Call mnuSettei_Click        '*** 環境設定 ***
    End Select
End Sub

Private Sub mnuAllQuit_Click()
    Unload Me
    End
End Sub

Private Sub mnuAQuitP_Click()
    Call mnuAllQuit_Click
End Sub

Private Sub mnuBuhin_Click()
    Call mnuDenkiBuhinhyou
End Sub

Private Sub mnuBuhinP_Click()
    Call mnuBuhin_Click
End Sub
'
Private Sub mnuBuhin2_Click()
    Call mnuDenkiBuhinhyou2
End Sub
'
Private Sub mnuBuhin2P_Click()
    Call mnuBuhin2_Click
End Sub

Private Sub mnuBuhinPRN_Click()
    Call mnuStandardBuhinhyouPrint
End Sub

Private Sub mnuBuhinPRNP_Click()
    Call mnuBuhinPRN_Click
End Sub

Private Sub mnuCodement_Click()
    Call mnuCodeBuhinMaintenance
End Sub

Private Sub mnuCodementP_Click()
    Call mnuCodement_Click
End Sub

Private Sub mnuFilePrnA_Click()
    Call mnuBuhinItiranhyouPrint
End Sub

Private Sub mnuFilePrnAP_Click()
    Call mnuFilePrnA_Click
End Sub

Private Sub mnuHistory_Click()
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
    Call mnuConvFile_Click
End Sub

Private Sub mnuSettei_Click()
    Call mnuKankyouSettei
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 0          '*** 作業フラグ、一般事項 ***
    Call mnuSousaSetumei
End Sub

Private Sub mnuSuuryo_Click()
    Call mnuBuhinSuuryohyouPrint
End Sub

Private Sub mnuSuuryoP_Click()
    Call mnuSuuryo_Click
End Sub

Private Sub mnuTraderment_Click()
    Call mnuCodeTraderMaintenance
End Sub

Private Sub mnuTradermentP_Click()
    Call mnuTraderment_Click
End Sub

Private Sub mnuVersion_Click()
    Call mnuVersionGamen
End Sub
