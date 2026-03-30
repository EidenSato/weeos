VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Trader_main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004000&
   Caption         =   "ＥＥＯＳ 商社コード  <一覧>"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   10095
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
   Icon            =   "Trader_main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6750
   ScaleWidth      =   10095
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5100
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8996
      _Version        =   393216
      Rows            =   21
      Cols            =   6
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
      SelectionMode   =   1
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
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Text            =   "商社名"
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   8400
      TabIndex        =   4
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "帳票形式(&L)"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.csv"
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "着目商社名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "ﾌｧｲﾙ(&X)"
      Begin VB.Menu mnuFilePrn 
         Caption         =   "商社一覧印刷(&P)"
      End
      Begin VB.Menu mnuFileWR 
         Caption         =   "一覧ﾌｧｲﾙ出力(&W)"
      End
      Begin VB.Menu mnu区切り線1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "閉じる(&Q)"
      End
      Begin VB.Menu mnu区切り線2 
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
      Begin VB.Menu mnuJumpC 
         Caption         =   "中心へｼﾞｬﾝﾌﾟ(&C)"
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
      Begin VB.Menu mnu区切り線4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ(&V)"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ﾎﾟｯﾌﾟｱｯﾌﾟﾒﾆｭｰ"
      Visible         =   0   'False
      Begin VB.Menu mnuJumpTP 
         Caption         =   "先頭へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpCP 
         Caption         =   "中心へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnuJumpEP 
         Caption         =   "最後部へｼﾞｬﾝﾌﾟ"
      End
      Begin VB.Menu mnu区切り線91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKouseihyouP 
         Caption         =   "構成表"
         Begin VB.Menu mnuKouseiP 
            Caption         =   "電気 構成表..."
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
         Begin VB.Menu mnu区切り線92 
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
      Begin VB.Menu mnuCodehyouP 
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
      End
      Begin VB.Menu mnu区切り線93 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackP 
         Caption         =   "閉じる"
      End
      Begin VB.Menu mnu区切り線94 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAQuitP 
         Caption         =   "EEOS２の終了"
      End
   End
End
Attribute VB_Name = "Trader_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'***********************************
'*** ＥＥＯＳ 商社コード 一覧 ***
'***   2000.12.15  by S.Fukazawa ***
'***********************************
'
Option Explicit
'
    Dim HeadTitle As String
    Dim FLGgyou As Integer
    Dim FLGpage As Integer
'
    Dim FLGoffsetX As Integer
    Dim FLGoffsetY As Integer
'                                   567twip=10mm,1440twip=1inch
    Private Const OrgWidth = 10215  '*** フォーム寸法初期値 ***
    Private Const OrgHeight = 6945
    Dim tempWidth As Integer
    Dim tempHeight As Integer
'
    Private Const c10mm = 567
    Private Const kijunX = c10mm * 1.8
    Private Const kijunY = c10mm * 0.9
    Private Const gyoukan = 1440 / 4
    Private Const moji_zureX = c10mm / 5
    Private Const moji_zureY = gyoukan / 4
    Private Const Gyoumax = 40            '*** 印刷行最大値 ***
    Private Const haba1X = c10mm * (20.5 - 2.5)
    Private Const haba1Y = gyoukan * (Gyoumax + 2)
    Private Const pichi1 = c10mm * 1.1
    Private Const pichi2 = c10mm * 2.6
    Private Const pichi3 = c10mm * 8#
    Private Const pichi4 = c10mm * 12#
    Private Const pichi5 = c10mm * 14#

Private Sub Form_Activate()
    FLGtrader = 1    '*** 商社画面存在フラグ設定 ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    Text1.Text = Trader(Trdps, 2)
    Text1.SetFocus
End Sub

Private Sub Form_Initialize()
    Trdp = 1         '*** 初期値設定 ***
    Trdps = 1
'
    HeadTitle = "商社ｺｰﾄﾞ <一覧>"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        Hup             '*** 上へ ***
'
    Case vbKeyUp        '*** ↑ ***
        H1up            '*** 一つ上へ ***
'
    Case vbKeyPageUp    '*** Roll Down
        Hdown           '*** 下へ ***
'
    Case vbKeyDown      '*** ↓ ***
        H1down          '*** 一つ下へ ***
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                       フォームのサイズの設定
    tempWidth = 360 + (OrgWidth - 720) * HyoujiBairitu + 360
    tempHeight = 360 + (OrgHeight - 720) * HyoujiBairitu + 360
                        '*** これで「Form_Resize」割り込みが発生する。 ***
    Width = tempWidth
    Height = tempHeight
'
    Call setFormArea    '*** フォームの表示位置の設定 ***
'
    Trader_main.Caption = HeadTitle
'
    Call RDcont         '*** 環境設定読み込み ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    DRVtrader = Xcont0(2) & "\TRADER.COD"
    Call RDtrader       '*** 商社コード 読み込み ***
'
    If FLG_job_error_end = 1 Then
        Kankyow_Itiran.Show 1
'
        FLG_job_error_end = 0
        Unload Me
    End If
'
    Call DSPlevel1      '*** 項目名表示 ***
    FLGtrd_data_change = 0   '*** 変更フラグ初期化 ***
'
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub

Private Sub Form_Resize()
'                   フォーム構成部品の表示位置の設定
    If Width > tempWidth Then
        FLGoffsetX = Width - tempWidth
    Else
        FLGoffsetX = 0
    End If
'
    If Height > tempHeight Then
        FLGoffsetY = (Height - tempHeight) / 2
    Else
        FLGoffsetY = 0
    End If
'                   グリッド幅の設定
    MSFlexGrid1.Width = 9375 * HyoujiBairitu + FLGoffsetX
    MSFlexGrid1.Font.Size = 10 * HyoujiBairitu
    MSFlexGrid1.RowHeightMin = 245 * HyoujiBairitu  '245
    MSFlexGrid1.Height = MSFlexGrid1.RowHeightMin * 21  '5145?
    MSFlexGrid1.Cols = 6
    MSFlexGrid1.ColWidth(0) = 545 * HyoujiBairitu
    MSFlexGrid1.ColWidth(1) = 900 * HyoujiBairitu  '840
    MSFlexGrid1.ColWidth(2) = 2400 * HyoujiBairitu + FLGoffsetX / 2
    MSFlexGrid1.ColWidth(3) = 2390 * HyoujiBairitu + (300 * HyoujiBairitu - 300)
    MSFlexGrid1.ColWidth(4) = 1160 * HyoujiBairitu
    MSFlexGrid1.ColWidth(5) = 1650 * HyoujiBairitu + FLGoffsetX / 2
'
    Call Buhin_Haichi    '*** 表示部品配置 ***
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FLGtrader = 0   '*** 商社画面存在フラグ初期化 ***
    Trdp = 1        '*** フラグ初期化 ***
    Trdps = 1       '***<再ロード時 "Form_Initialize"前に "MSFlexGrid1_Scroll" が発生してしまう予防>***
End Sub

Private Sub DSPlevel1()     '*** 一覧表示 ***
    Dim temp As String
'
    If Trdnum0 > 20 Then
        MSFlexGrid1.Rows = Trdnum0 + 2
    Else
        MSFlexGrid1.Rows = 22
    End If
'
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Text = "ｺｰﾄﾞ"
        .Col = 1
        .Text = "略 称"
        .Col = 2
        .Text = "      正式名称"
        .Col = 3
        .Text = "      電話/FAX番号"
        .Col = 4
        .Text = "担当者"
'       .Col = 5
'       .Text = "郵便番号"
'       .Col = 6
'       .Text = "事務所所在地"
        .Col = 5
        .Text = "取り扱いメーカー"
'       .Col = 8
'       .Text = "発注書最終番号"
'       .Col = 6
'       .Text = "更新日"
    End With
'
    With MSFlexGrid1
        .ColAlignment(0) = flexAlignCenterCenter    '*** 中央揃え ***
        .ColAlignment(1) = flexAlignCenterCenter    '*** 中央揃え ***
        .ColAlignment(2) = flexAlignLeftCenter      '*** 左揃え ***
        .ColAlignment(3) = flexAlignLeftCenter      '*** 左揃え ***
        .ColAlignment(4) = flexAlignCenterCenter    '*** 中央揃え ***
'       .ColAlignment(5) = flexAlignLeftCenter      '*** 左揃え ***
'       .ColAlignment(6) = flexAlignLeftCenter      '*** 左揃え ***
        .ColAlignment(5) = flexAlignLeftCenter      '*** 左揃え ***
'       .ColAlignment(8) = flexAlignCenterCenter    '*** 中央揃え ***
'       .ColAlignment(6) = flexAlignCenterCenter    '*** 中央揃え ***
    End With
'
    Call DATA_settei
'
End Sub
Private Sub DATA_settei()
    Dim i As Integer
'
    For i = 1 To Trdnum0
        MSFlexGrid1.Row = i
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = Trader(i, 0)
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = Trader(i, 2)
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = Trader(i, 1)
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = Trader(i, 3)
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Text = Trader(i, 4)
'       MSFlexGrid1.Col = 5
'       MSFlexGrid1.Text = Trader(i, 5)
'       MSFlexGrid1.Col = 6
'       MSFlexGrid1.Text = Trader(i, 6)
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = Trader(i, 7)
'       MSFlexGrid1.Col = 8
'       MSFlexGrid1.Text = Trader(i, 8)
'       MSFlexGrid1.Col = 9
'       MSFlexGrid1.Text = Trader(i, 9)
    Next i
'
    MSFlexGrid1.TopRow = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub H1down()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Trdnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Trdnum0 > j + 19 Then
        MSFlexGrid1.TopRow = j + 1
    Else
        Beep
        MSFlexGrid1.TopRow = Trdnum0 - 18
    End If
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub H1up()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Trdnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j = 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 1
    End If
'
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub Hdown()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Trdnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Trdnum0 > j + 29 Then
        MSFlexGrid1.TopRow = j + 10
    Else
        Beep
        MSFlexGrid1.TopRow = Trdnum0 - 18
    End If
'
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub Hup()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Trdnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j - 10 < 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 10
    End If
'
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub cmdUp_Click()
    Hup
End Sub

Private Sub cmdDown_Click()
    Hdown
End Sub

Private Sub cmdPage_Click()
    FLGtrd_data_change = 0   '*** 変更フラグ初期化 ***
'
    Trader_ment.Show 1
'                       '*** モーダルフォーム実行後、続いて実行される。***
    If FLGtrd_data_change = 1 Then   '*** 変更有り ***
        'Call DATA_settei
        Call DSPlevel1
        FLGtrd_data_change = 0   '*** 変更フラグ初期化 ***
    End If
'
    Text1.Text = Trader(Trdps, 2)
    Text1.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub mnuAQuitP_Click()
    Unload Me
    End
End Sub

Private Sub mnuBackP_Click()
    Unload Me
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
    If FLGitem = 1 Then
        Pcod_item.SetFocus
    Else
        Pcod_item.Show
    End If
End Sub

Private Sub mnuCodeP_Click()
    Call mnuCode_Click
End Sub

Private Sub mnuFilePrnA_Click()
    Call mnuBuhinItiranhyouPrint
End Sub

Private Sub mnuFilePrnAP_Click()
    Call mnuFilePrnA_Click
End Sub

Private Sub mnuHinsyu_Click()
    If FLGindex = 1 Then
        Pcod_index.SetFocus
    Else
        Pcod_index.Show
    End If
End Sub

Private Sub mnuHinsyuP_Click()
    Call mnuHinsyu_Click
End Sub

Private Sub mnuJumpCP_Click()
    Call mnuJumpC_Click
End Sub

Private Sub mnuJumpEP_Click()
    Call mnuJumpE_Click
End Sub

Private Sub mnuJumpTP_Click()
    Call mnuJumpT_Click
End Sub

Private Sub mnuKousei_Click()
    Call mnuDenkiKouseihyou
End Sub

Private Sub mnuKouseiP_Click()
    Call mnuKousei_Click
End Sub

Private Sub mnuMakerment_Click()
    If FLGmaker = 1 Then
        Maker_main.SetFocus
    Else
        Maker_main.Show
    End If
End Sub

Private Sub mnuMakermentP_Click()
    Call mnuMakerment_Click
End Sub

Private Sub mnuOption_Click()
    Option_Sel.Show 1   '*** オプション設定 ***
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
    If FLGmain = 1 Then
        Pcod_main.SetFocus
    Else
        Pcod_main.Show
    End If
End Sub

Private Sub mnuPmainP_Click()
    Call mnuPmain_Click
End Sub

Private Sub mnuReform_Click()
'                       フォームのサイズの設定
    Width = tempWidth
    Height = tempHeight '*** これで「Form_Resize」割り込みが発生する。 ***
'
    Call setFormArea    '*** フォームの表示位置の設定 ***
End Sub

Private Sub mnuSuuryo_Click()
    Call mnuBuhinSuuryohyouPrint
End Sub

Private Sub mnuSuuryoP_Click()
    Call mnuSuuryo_Click
End Sub

Private Sub MSFlexGrid1_Click()
    If MSFlexGrid1.Row <= Trdnum0 Then
        Trdps = MSFlexGrid1.Row
        Text1.Text = Trader(Trdps, 2)
    End If
'
    Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row <= Trdnum0 Then
        Trdps = MSFlexGrid1.Row
        Text1.Text = Trader(Trdps, 2)
'
        Call cmdPage_Click
    End If
'
    Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub MSFlexGrid1_Scroll()
    Trdp = MSFlexGrid1.TopRow
    Trdps = MSFlexGrid1.TopRow
    Text1.Text = Trader(Trdps, 2)
'
    Text1.SetFocus
End Sub

Private Sub mnuAQuit_Click()
    Unload Me
    End
End Sub

Private Sub mnuBack_Click()
    FLGjob = 0
    Unload Me
End Sub

Private Sub mnuFilePrn_Click()  '*** 商社 一覧印刷 ***
    Dim objPrinter As Printer
'
    flag_cancel = False
    Printer_Window.Show 1
        If flag_cancel = True Then Exit Sub
'
    For Each objPrinter In Printers
        If objPrinter.DeviceName = strMyPrinter Then
            Set Printer = objPrinter    'オブジェクトに代入
        End If
    Next
'
    Eeos2_mainMDI.MousePointer = vbHourglass    '*** 砂時計 ***
    MSFlexGrid1.MousePointer = flexHourglass
    Me.MousePointer = vbHourglass
'
    Trader_main.Caption = HeadTitle & "  =>  プリントバッファーに転送中 !!!"
    FLGpage = 1             '*** ページ初期化
    DoEvents                '*** 画面書き直し ***
'
    PRNheader  '*** ヘッダー印刷 (FLGgyou = 0)***
    DoEvents
'
    For Trdp = 1 To Trdnum0
        If FLGgyou >= Gyoumax Then
            PRNfooter
            Printer.EndDoc   '*** 改ページ ***
'
            PRNheader
        End If
'
        PRNkoumoku        '*** 項目印刷 ***
        DoEvents
'
    Next Trdp
'
    PRNfooter           '*** 項末印刷 ***
    Printer.EndDoc      '*** プリンター書き込み ***
'
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault             '*** 砂時計解除 ***
    MSFlexGrid1.MousePointer = flexArrow
    Eeos2_mainMDI.MousePointer = vbDefault
End Sub

Private Sub mnuFileWR_Click()   '*** 商社 一覧ファイル出力 ***
    Dim CSVfile As String, PDa As String, PD As String, Tradername As String
    Dim FILE_number As Integer
'
    On Error GoTo error_F
    With CommonDialog1
        .CancelError = True     ' CancelError プロパティを真 (True) に設定します。
        .Flags = cdlOFNHideReadOnly    ' Flags プロパティを設定します。
                                    ' リスト ボックスに表示されるフィルタを設定します。
        .Filter = "CSV形式 ファイル (*.csv)|*.csv|" & _
                    "テキスト ファイル (*.txt)|*.txt|" & _
                    "すべてのファイル (*.*)|*.*|"
        .FilterIndex = 0        ' "CSV形式 ファイル" を既定のフィルタとして指定します。
        .DialogTitle = HeadTitle & " => ファイル出力"
        .ShowOpen               ' [ファイルを開く] ダイアログ ボックスを表示します。
    End With
'
    Trader_main.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
    CSVfile = CommonDialog1.FileName
    On Error GoTo 0
'
    DoEvents                '*** 画面書き直し ***
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open CSVfile For Output As #FILE_number
    For Trdp = 1 To Trdnum0
        DoEvents            '*** 砂時計用 ***
'
        PDa = Chr(34) & Trader(Trdp, 0) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 2) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 1) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 3) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 4) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 5) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 6) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 7) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 8) & Chr(34) & "," _
            & Chr(34) & Trader(Trdp, 9) & Chr(34)
'
        Print #FILE_number, PDa  '*** データ書き込み ***
    Next Trdp
    Close #FILE_number
    Trader_main.MousePointer = vbDefault
    MSFlexGrid1.MousePointer = flexArrow
'
error_F:
    Trader_main.Caption = HeadTitle
'
End Sub

Private Sub mnuJumpC_Click()
    Dim Tyuusin As Integer
'
    Tyuusin = Trdnum0 \ 2
    If Tyuusin <= 10 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Tyuusin - 9
    End If
'
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub mnuJumpE_Click()
    If Trdnum0 < 21 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Trdnum0 - 18
    End If
'
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub mnuJumpT_Click()
    MSFlexGrid1.TopRow = 1
'
    Trdp = MSFlexGrid1.TopRow
    Trdps = Trdp
    Text1.Text = Trader(Trdps, 2)
End Sub

Private Sub mnuKaihan_Click()
    FLG_Setumei = 1     '*** 1: 改版履歴 ***
    Setumei_gamen.Show 1
End Sub

Private Sub mnuSettei_Click()
    Kankyow_Itiran.Show 1
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 5          '*** 商社コード表フラグ ***
    FLG_Setumei = 0     '*** 0: 操作説明 ***
    Setumei_gamen.Show 1
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
    If FLGtrader = 1 Then
        Trader_main.SetFocus
    Else
'       Trader_main.Show
    End If
End Sub

Private Sub mnuVersion_Click()
    Version_gamen.Show 1
End Sub

Private Sub PRNfooter()
'                   *** フッター印刷 ***
    Printer.CurrentX = kijunX + c10mm
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    SETfont_size 9, 1
    Printer.Print "M5103-05";
'
    Printer.CurrentX = kijunX + haba1X - c10mm * 4
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    SETfont_size 9, 0
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷";
End Sub

Private Sub PRNkoumoku()
'                   *** 項目印刷 ***
    Dim makername As String, PD As String
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Trader(Trdp, 0)
'
    Printer.CurrentX = kijunX + pichi1 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Trader(Trdp, 2)
'
    Printer.CurrentX = kijunX + pichi2 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Trader(Trdp, 1)
'
    Printer.CurrentX = kijunX + pichi3 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Trader(Trdp, 3)
'
    Printer.CurrentX = kijunX + pichi4 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Trader(Trdp, 4)
'
    Printer.CurrentX = kijunX + pichi5 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Trader(Trdp, 7)
'
'    Printer.CurrentX = kijunX + pichi6 + moji_zureX
'    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
'    Printer.Print Trader(Trdp, 6)
'
    FLGgyou = FLGgyou + 1
End Sub

Private Sub PRNheader()
'                   *** 一覧表ヘッダー印刷 ***
    Dim i As Integer
'
    FLGgyou = 0
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    Printer.Orientation = vbPRORPortrait    '*** ポートレート   210x290 ***
'   Printer.Orientation = vbPRORLandscape   '*** ランドスケープ 290x210 ***
'
    Printer.CurrentX = kijunX + haba1X - c10mm * 2
    Printer.CurrentY = kijunY - gyoukan + moji_zureY
    SETfont_size 9, 1      '*** フォント,サイズ設定 ***
    Printer.Print "ﾍﾟｰｼﾞ： " & FLGpage
    FLGpage = FLGpage + 1
'
    Printer.CurrentX = kijunX + c10mm / 2
    Printer.CurrentY = kijunY + moji_zureY
    SETfont_size 10, 1      '*** フォント,サイズ設定 ***
    Printer.Print "===  [  営電標準  ]                 電気部品コード・単価表                  ＜商社 一覧表＞  ==="
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan + moji_zureY
    SETfont_size 9, 1      '*** フォント,サイズ設定 ***
    Printer.Print "ｺｰﾄﾞ    略称               商社名                      電話/FAX番号         担当者       取り扱いメーカー"
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
'
    Printer.Line (kijunX, kijunY)-Step(haba1X, haba1Y), 0, B
'
    Printer.Line (kijunX, kijunY + gyoukan)-Step(haba1X, gyoukan), 0, B
End Sub

Private Sub Buhin_Haichi()       '*** 表示部品配置 ***
    MSFlexGrid1.Left = 360
    MSFlexGrid1.Top = 360 + FLGoffsetY
'
    Text1.Left = 360 + (3855 - 360) * HyoujiBairitu + FLGoffsetX
    Text1.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    Text1.FontSize = 10 * HyoujiBairitu
    Text1.Width = 975 * HyoujiBairitu
    Text1.Height = 285 * HyoujiBairitu
'
    Label1.Left = 360 + (2640 - 360) * HyoujiBairitu + FLGoffsetX
    Label1.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    Label1.FontSize = 10 * HyoujiBairitu
    Label1.Width = 1215 * HyoujiBairitu
    Label1.Height = Text1.Height
'
    cmdClose.Left = 360 + (8400 - 360) * HyoujiBairitu + FLGoffsetX
    cmdClose.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdClose.FontSize = 9 * HyoujiBairitu
    cmdClose.Width = 855 * HyoujiBairitu
    cmdClose.Height = 495 * HyoujiBairitu
'
    cmdDown.Left = 360 + (7440 - 360) * HyoujiBairitu + FLGoffsetX
    cmdDown.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Width = 735 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdUp.Left = 360 + (6600 - 360) * HyoujiBairitu + FLGoffsetX
    cmdUp.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Width = 735 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdPage.Left = 360 + (5040 - 360) * HyoujiBairitu + FLGoffsetX
    cmdPage.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdPage.FontSize = 9 * HyoujiBairitu
    cmdPage.Width = 1095 * HyoujiBairitu
    cmdPage.Height = 495 * HyoujiBairitu
End Sub

Private Sub MENU_settei()       '*** メニュー状態設定 ***
'
    If FLGconst = 1 Then        '*** 構成表画面存在 ***
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
    Else
        Me.mnuTraderment.Checked = False
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
            End If
        Else
            Me.mnuHinsyu.Checked = False
            Me.mnuPmain.Checked = False
            Me.mnuPmain.Enabled = False
'
            Me.mnuHinsyuP.Checked = False
            Me.mnuPmainP.Checked = False
            Me.mnuPmainP.Enabled = False
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
End Sub

Private Sub setFormArea()   '*** フォームの表示位置の設定 ***
    If Eeos2_mainMDI.ScaleHeight > Height Then
        Top = (Eeos2_mainMDI.ScaleHeight - Height) \ 2
    Else
        Top = 0
    End If
'
    If Eeos2_mainMDI.ScaleWidth > Width Then
        Left = (Eeos2_mainMDI.ScaleWidth - Width) * 2 \ 3
    Else
        Left = 0
    End If
End Sub
