VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Pcod_index 
   BackColor       =   &H00004000&
   Caption         =   "部品ｺｰﾄﾞ <品種一覧>"
   ClientHeight    =   6255
   ClientLeft      =   75
   ClientTop       =   645
   ClientWidth     =   10095
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Pcod_index.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6255
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdHelp 
      Caption         =   "？"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdKensaku 
      Caption         =   "部品検索(&S)"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   600
   End
   Begin VB.TextBox txtPoint 
      Alignment       =   2  '中央揃え
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
      Height          =   280
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "品目一覧(&N)"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "帳票形式(&L)"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   5640
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "コード表ファイル出力"
      Filter          =   "*.csv"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
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
   Begin VB.Label lblPoint 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "着目品種"
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
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "項目"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "ﾌｧｲﾙ(&X)"
      Begin VB.Menu mnuFilePrn 
         Caption         =   "品種一覧印刷(&P)"
      End
      Begin VB.Menu mnuFileWR 
         Caption         =   "一覧ﾌｧｲﾙ出力(&W)"
      End
      Begin VB.Menu mnu区切り線11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "閉じる(&Q)"
      End
      Begin VB.Menu mnu区切り線12 
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
Attribute VB_Name = "Pcod_index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'******************************************
'*** ＥＥＯＳ２ 電気部品コード 品種一覧   ***
'***          2001.01.30  by S.Fukazawa ***
'******************************************
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
    Private Const OrgWidth = 10215  '9840 *** フォーム寸法初期値 ***
    Private Const OrgHeight = 6945  '7155
    Dim tempWidth As Integer
    Dim tempHeight As Integer
'
    Private Const c10mm = 567
    Private Const kijunX = c10mm * 2.5
    Private Const kijunY = c10mm * 1
    Private Const gyoukan = 1440 / 6
    Private Const moji_zureX = c10mm / 5
    Private Const moji_zureY = gyoukan / 6
    Private Const Gyoumax = 58            '*** 印刷行最大値 ***
    Private Const haba1X = c10mm * (21 - 4)
    Private Const haba1Y = gyoukan * (Gyoumax + 4)
    Private Const pichi1 = c10mm * 1.3
    Private Const pichi2 = c10mm * 4#
    Private Const pichi3 = c10mm * 14#
    Private Const pichi4 = c10mm * 15.6
    Private Const pichi5 = c10mm * 4#
    Private Const pichi6 = c10mm * 9#
'
    Dim Time_up As Boolean

Private Sub Form_Activate()
    FLGindex = 1      '*** 部品コード項目画面存在フラグ設定 ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    lblItem.Caption = Aitem0(ips0, 0)
'
    If icps0 <> ips0 Then   '*** 現在の内容と異なるときは再読込 ***
        icps0 = ips0
        DRVindex0 = Xcont0(2) & "\" & Aitem0(icps0, 0) & "\" & Aitem0(icps0, 0) & "INDEX.COD"
        Call RDindex(DRVindex0, Bindex0(), Bnum0, Bdim0)    '*** INDEX.COD 読み込み ***
'
        jp0 = 1             '*** 初期値設定 ***
        jps0 = jp0
        jcps0 = 0
        Call DATA_settei
        FLGindex_data_change = 0   '*** 変更フラグ初期化 ***
    End If
'
    Call setPointtxt(jps0)
    txtPoint.SetFocus
End Sub

Private Sub Form_Initialize()
    jp0 = 1          '*** 初期値設定 ***
    jps0 = jp0
    jcps0 = 0
'
    HeadTitle = "部品ｺｰﾄﾞ <品種一覧>"
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
'                       フォームのサイズの設定
    tempWidth = 360 + (OrgWidth - 720) * HyoujiBairitu + 360
    tempHeight = 360 + (OrgHeight - 720) * HyoujiBairitu + 360
                        '*** これで「Form_Resize」割り込みが発生する。 ***
    Width = tempWidth
    Height = tempHeight
'
    Call setFormArea    '*** フォームの表示位置の設定
'
    Me.Caption = HeadTitle
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    icps0 = ips0
    DRVindex0 = Xcont0(2) & "\" & Aitem0(icps0, 0) & "\" & Aitem0(icps0, 0) & "INDEX.COD"
    Call RDindex(DRVindex0, Bindex0(), Bnum0, Bdim0)    '*** INDEX.COD 読み込み ***
'
    Call DSPlevel2              '*** 品種名表示 ***
    FLGindex_data_change = 0    '*** 変更フラグ初期化 ***
'
    Time_up = False
    Timer1.Enabled = False
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
        FLGoffsetY = (Height - tempHeight) \ 2
    Else
        FLGoffsetY = 0
    End If
'                   グリッド幅の設定
    MSFlexGrid1.Width = 9375 * HyoujiBairitu + FLGoffsetX
    MSFlexGrid1.Font.Size = 10 * HyoujiBairitu
    MSFlexGrid1.RowHeightMin = 245 * HyoujiBairitu  '245
    MSFlexGrid1.Height = MSFlexGrid1.RowHeightMin * 21  '5145?
    MSFlexGrid1.Cols = 6
    MSFlexGrid1.ColWidth(0) = 440 * HyoujiBairitu
    MSFlexGrid1.ColWidth(1) = 700 * HyoujiBairitu
    MSFlexGrid1.ColWidth(2) = 2270 * HyoujiBairitu + FLGoffsetX \ 2
    MSFlexGrid1.ColWidth(3) = 1080 * HyoujiBairitu
    MSFlexGrid1.ColWidth(4) = 840 * HyoujiBairitu
    MSFlexGrid1.ColWidth(5) = 3700 * HyoujiBairitu + (300 * HyoujiBairitu - 300) + FLGoffsetX \ 2
'
    Call Buhin_Haichi       '*** 表示部品配置 ***
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FLGindex = 0    '*** 部品コード品種画面存在フラグ初期化 ***
    jp0 = 1         '*** フラグ初期化 ***
    jps0 = jp0      '***<再ロード時 "Form_Initialize"前に "MSFlexGrid1_Scroll" が発生してしまう予防>***
'
    If FLGmain = 1 Then
        Unload Pcod_main    '*** 子画面も抹消 ***
    End If
End Sub

Private Sub H1down()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Bnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Bnum0 > j + 19 Then
        MSFlexGrid1.TopRow = j + 1
    Else
        Beep
        MSFlexGrid1.TopRow = Bnum0 - 18
    End If
'
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub H1up()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Bnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j = 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 1
    End If
'
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub Hdown()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Bnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Bnum0 > j + 29 Then
        MSFlexGrid1.TopRow = j + 10
    Else
        Beep
        MSFlexGrid1.TopRow = Bnum0 - 18
    End If
'
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub Hup()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Bnum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j - 10 < 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 10
    End If
'
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub DSPlevel2()
                    '*** 品種一覧表示 ***
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Text = "行"
        .ColAlignment(0) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 1
        .Text = "親ｺｰﾄﾞ"
        .ColAlignment(1) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 2
        .Text = "        品 種 代 表 型 名"
        .ColAlignment(2) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 3
        .Text = "図面表記"
        .ColAlignment(3) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 4
        .Text = "ﾒｰｶｰ"
        .ColAlignment(4) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 5
        .Text = "                    概    要"
        .ColAlignment(5) = flexAlignLeftCenter      '*** 左詰め ***
    End With
'
    Call DATA_settei
End Sub

Private Sub DATA_settei()
    Dim i As Integer
    Dim makername As String
    Dim hyouki As String
'
    If Bnum0 > 20 Then  '*** 行数設定 ***
        MSFlexGrid1.Rows = Bnum0 + 2
    Else
        MSFlexGrid1.Rows = 22
    End If
'
    For i = 1 To Bnum0
        MSFlexGrid1.Row = i
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = i                    '*** 中央揃え ***
        MSFlexGrid1.Col = 1
        If Bindex0(i, 15) = "1" Then
            MSFlexGrid1.CellForeColor = QBColor(10)
        Else
            MSFlexGrid1.CellForeColor = QBColor(15)
        End If
        MSFlexGrid1.Text = "L" & Bindex0(i, 0)  '*** 中央揃え ***
        MSFlexGrid1.Col = 2
        If Bindex0(i, 15) = "1" Then
            MSFlexGrid1.CellForeColor = QBColor(10)
        Else
            MSFlexGrid1.CellForeColor = QBColor(15)
        End If
        MSFlexGrid1.Text = " " & Bindex0(i, 3) & "xxx" & Bindex0(i, 4)
        MSFlexGrid1.Col = 3
            MSFlexGrid1.CellForeColor = QBColor(15)
        hyouki = Bindex0(i, 6)
            Call setFullName(hyouki)
        MSFlexGrid1.Text = hyouki               '*** 中央揃え ***
        MSFlexGrid1.Col = 4
        makername = Bindex0(i, 5)
            Call Makerget2(makername)
        MSFlexGrid1.Text = makername            '*** 中央揃え ***
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = " " & Bindex0(i, 1)
    Next i
'
        MSFlexGrid1.Row = i '*** Bnum0+1 を消去 ***
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = ""
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = ""
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = ""
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = ""
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Text = ""
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = ""
'
    If Bnum0 < 21 Then
        For i = Bnum0 + 1 To 21
            MSFlexGrid1.Row = i
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = ""
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = ""
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = ""
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = ""
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = ""
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Text = ""
        Next i
    End If
'
    MSFlexGrid1.TopRow = jp0
    Call setPointtxt(jp0)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Call Hdown
End Sub

Private Sub cmdHelp_Click()
    Dim temp As Integer
'
    temp = FLG_Setumei
    FLG_Setumei = 10
    Setumei_gamen.Show 1
'
    FLG_Setumei = temp
End Sub

Private Sub cmdKensaku_Click()
    Dim temp As Integer
'
    Pcod_retrieve.Show 1
'    jps0 = Val(Kouho(Kwork, 2))
'    kps0 = Val(Kouho(Kwork, 3))
'
    If flag_cancel = True Then Exit Sub
'
    temp = jps0
    If Bnum0 <= 20 Then
        MSFlexGrid1.TopRow = 1
    ElseIf jps0 <= 10 Then
        MSFlexGrid1.TopRow = 1
    ElseIf Bnum0 - jps0 <= 18 Then
        MSFlexGrid1.TopRow = Bnum0 - 18
    Else
        MSFlexGrid1.TopRow = jps0
    End If
'
    jps0 = temp
    Call setPointtxt(jps0)
'
    Call cmdMain_Click
End Sub

Private Sub cmdMain_Click()
    If FLGmain = 1 Then
        Pcod_main.SetFocus
    Else
        Pcod_main.Show
    End If
End Sub

Private Sub cmdPage_Click()
    FLGindex_data_change = 0     '*** 変更フラグ初期化 ***
'
    Pcod_index_c.Show 1
'                       '*** モーダルフォーム実行後、続いて実行される。***
    If FLGindex_data_change = 1 Then '*** 変更有り ***
        Call DATA_settei
        FLGindex_data_change = 0   '*** 変更フラグ初期化 ***
    End If
'
    Call setPointtxt(jps0)
    txtPoint.SetFocus
End Sub

Private Sub cmdUp_Click()
    Call Hup
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
'        Pcod_item.Show
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
'        Pcod_index.Show
    End If
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
    Call cmdMain_Click
End Sub

Private Sub mnuPmainP_Click()
    Call cmdMain_Click
End Sub

Private Sub mnuReform_Click()
    Width = tempWidth
    Height = tempHeight     '*** これで「Form_Resize」割り込みが発生する。 ***
'
    Call setFormArea        '*** フォームの表示位置の設定 ***
End Sub

Private Sub mnuSettei_Click()
    Call mnuKankyouSettei
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

Private Sub MSFlexGrid1_Click()
    If MSFlexGrid1.Row <= Bnum0 Then
        jps0 = MSFlexGrid1.Row
        Call setPointtxt(jps0)
    End If
'
    txtPoint.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row <= Bnum0 Then
        jps0 = MSFlexGrid1.Row
        Call setPointtxt(jps0)
'
        Call cmdMain_Click
'
    Else
        txtPoint.SetFocus
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub MSFlexGrid1_Scroll()
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
'
    txtPoint.SetFocus
End Sub

Private Sub mnuAQuit_Click()
    Unload Me
    End
End Sub

Private Sub mnuBack_Click()
    FLGjob = 0
    Unload Me
End Sub

Private Sub mnuFilePrn_Click()  '*** 品種一覧印刷 ***
    Dim j As Integer
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
    Me.Caption = HeadTitle & "  =>  プリントバッファーに転送中 !!!"
    FLGpage = 1             '*** ページ初期化
    DoEvents                '*** 画面書き直し ***
'
    Call PRNheader  '*** ヘッダー印刷 (FLGgyou = 0)***
    DoEvents
'
    For j = 1 To Bnum0
        If FLGgyou >= Gyoumax Then
            Call PRNfooter
            Printer.EndDoc  '*** 改ページ ***
'
            Call PRNheader
        End If
'
        Call PRNkoumoku(j)  '*** 項目印刷 ***
        DoEvents
'
    Next j
'
    Call PRNfooter          '*** 項末印刷 ***
    Printer.EndDoc          '*** プリンター書き込み ***
'
    Call timer_waite(1000)  '*** 表示認識待ち ***
'
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault             '*** 砂時計解除 ***
    MSFlexGrid1.MousePointer = flexArrow
    Eeos2_mainMDI.MousePointer = vbDefault
End Sub

Private Sub mnuFileWR_Click()   '*** 品種一覧ファイル出力 ***
    Dim j As Integer
    Dim CSVfile As String
    Dim PDa As String
    Dim PD As String
    Dim makername As String
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
    Eeos2_mainMDI.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
    Me.MousePointer = vbHourglass
    CSVfile = CommonDialog1.FileName
    On Error GoTo 0
    DoEvents
'
    Me.Caption = HeadTitle & "  =>  ファイルに出力中！"
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open CSVfile For Output As #FILE_number
    For j = 1 To Bnum0
        PD = Chr(34) & "L" & Bindex0(j, 0) & Chr(34) & "," _
            & Chr(34) & Bindex0(j, 1) & Chr(34) & "," _
            & Chr(34) & Bindex0(j, 2) & Chr(34) & "," _
            & Chr(34) & Bindex0(j, 3) & "xxx" & Bindex0(j, 4) & Chr(34) & ","
        makername = Bindex0(j, 5)
            Call Makerget2(makername)
        PDa = PD & Chr(34) & makername & Chr(34) & "," _
            & Chr(34) & Bindex0(j, 6) & Chr(34) & "," _
            & Chr(34) & Bindex0(j, 7) & Chr(34)
'
        Print #FILE_number, PDa  '*** データ書き込み ***
        DoEvents
    Next j
    Close #FILE_number
'
    Call timer_waite(1000)  '*** 表示認識待ち ***
'
error_F:
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
    MSFlexGrid1.MousePointer = flexArrow
    Eeos2_mainMDI.MousePointer = vbDefault
End Sub

Private Sub mnuJumpC_Click()
    Dim Tyuusin As Integer
'
    Tyuusin = Bnum0 \ 2
    If Tyuusin <= 10 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Tyuusin - 9
    End If
'
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub mnuJumpE_Click()
    If Bnum0 < 20 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Bnum0 - 18
    End If
'
    jp0 = MSFlexGrid1.TopRow
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub mnuJumpT_Click()
    MSFlexGrid1.TopRow = 1
'
    jp0 = MSFlexGrid1.TopRow + 1
    jps0 = jp0
    Call setPointtxt(jps0)
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 3          '*** 部品コード表フラグ ***
    Call mnuSousaSetumei
End Sub

Private Sub mnuVersion_Click()
    Call mnuVersionGamen
End Sub

Private Sub PRNfooter() '*** フッター印刷 ***
    Printer.CurrentX = kijunX + c10mm
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    Call SETfont_size(9, 1)
    Printer.Print "M5103-02";
'
    Printer.CurrentX = kijunX + haba1X - c10mm * 4
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    Call SETfont_size(9, 0)
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷";
End Sub

Private Sub PRNkoumoku(j As Integer)
'                   *** 項目印刷 ***
    Dim makername As String
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 4) + moji_zureY
    Printer.Print "L" & Bindex0(j, 0)
'
    Printer.CurrentX = kijunX + pichi1 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 4) + moji_zureY
    Printer.Print Bindex0(j, 3) & "xxx" & Bindex0(j, 4)
'
    Printer.CurrentX = kijunX + pichi2 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 4) + moji_zureY
    Printer.Print Bindex0(j, 1)
'
    Printer.CurrentX = kijunX + pichi3 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 4) + moji_zureY
    Printer.Print Bindex0(j, 6)
'
    Printer.CurrentX = kijunX + pichi4 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 4) + moji_zureY
        makername = Bindex0(j, 5)
            Call Makerget2(makername)
    Printer.Print makername
'
    Printer.CurrentX = kijunX + pichi5 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 5) + moji_zureY
    Printer.Print Bindex0(j, 2)
'
    Printer.CurrentX = kijunX + pichi6 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 5) + moji_zureY
    Printer.Print Bindex0(j, 7)
'
    FLGgyou = FLGgyou + 2
End Sub

Private Sub PRNheader() '*** 項目一覧表ヘッダー印刷 ***
    FLGgyou = 0
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    Printer.Orientation = vbPRORPortrait    '*** ポートレート 210x290 ***
'
    Printer.CurrentX = kijunX + haba1X - c10mm * 2
    Printer.CurrentY = kijunY - gyoukan + moji_zureY
    Call SETfont_size(9, 1)         '*** フォント,サイズ設定 ***
    Printer.Print "ﾍﾟｰｼﾞ： " & FLGpage
    FLGpage = FLGpage + 1
'
    Printer.CurrentX = kijunX + c10mm / 2
    Printer.CurrentY = kijunY + moji_zureY * 3.5
    Call SETfont_size(10, 1)        '*** フォント,サイズ設定 ***
    Printer.Print "===  [  営電標準  ]     電気部品コード・単価表                      ＜品種名一覧表＞  ==="
'
    Printer.CurrentX = kijunX + c10mm * 9.5
    Printer.CurrentY = kijunY + moji_zureY * 3.5
    Printer.Print Aitem0(icps0, 1)
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * 2 + moji_zureY
    SETfont_size 9, 1      '*** フォント,サイズ設定 ***
    Printer.Print "親コード    型 名             品種和名                                                図面表記    ﾒｰｶｰ"
'
    Printer.CurrentX = kijunX + c10mm * 4.8 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * 3 + moji_zureY
    Printer.Print "品種英名                            備 考"
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
'
    Printer.Line (kijunX, kijunY)-Step(haba1X, haba1Y), 0, B
'
    Printer.Line (kijunX, kijunY + gyoukan * 2)-Step(haba1X, gyoukan * 2), 0, B
End Sub

Private Sub Buhin_Haichi()       '*** 表示部品配置 ***
    MSFlexGrid1.Left = 360
    MSFlexGrid1.Top = 360 + FLGoffsetY
'
    lblItem.FontSize = 10 * HyoujiBairitu
    lblItem.Height = 240 * HyoujiBairitu
    lblItem.Left = 360
    lblItem.Top = 360 - lblItem.Height + FLGoffsetY
    lblItem.Width = 480 * HyoujiBairitu
'
    cmdHelp.FontSize = 10 * HyoujiBairitu
    cmdHelp.Height = 220 * HyoujiBairitu
    cmdHelp.Left = 360 + (5880 - 360) * HyoujiBairitu
    cmdHelp.Top = 360 - cmdHelp.Height + FLGoffsetY
    cmdHelp.Width = 255 * HyoujiBairitu
'
    txtPoint.Left = 360 + (1330 - 360) * HyoujiBairitu
    txtPoint.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    txtPoint.FontSize = 10 * HyoujiBairitu
    txtPoint.Width = 855 * HyoujiBairitu
    txtPoint.Height = 280 * HyoujiBairitu
'
    lblPoint.Left = 360
    lblPoint.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    lblPoint.FontSize = 10 * HyoujiBairitu
    lblPoint.Width = 975 * HyoujiBairitu
    lblPoint.Height = txtPoint.Height
'
    cmdKensaku.Left = 360 + (2400 - 360) * HyoujiBairitu
    cmdKensaku.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdKensaku.FontSize = 9 * HyoujiBairitu
    cmdKensaku.Width = 1095 * HyoujiBairitu
    cmdKensaku.Height = 495 * HyoujiBairitu
'
    cmdClose.Left = 360 + (8520 - 360) * HyoujiBairitu
    cmdClose.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdClose.FontSize = 9 * HyoujiBairitu
    cmdClose.Width = 855 * HyoujiBairitu
    cmdClose.Height = 495 * HyoujiBairitu
'
    cmdDown.Left = 360 + (7560 - 360) * HyoujiBairitu
    cmdDown.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Width = 735 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdUp.Left = 360 + (6720 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Width = 735 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdMain.Left = 360 + (5280 - 360) * HyoujiBairitu
    cmdMain.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdMain.FontSize = 9 * HyoujiBairitu
    cmdMain.Width = 1095 * HyoujiBairitu
    cmdMain.Height = 495 * HyoujiBairitu
'
    cmdPage.Left = 360 + (3840 - 360) * HyoujiBairitu
    cmdPage.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdPage.FontSize = 9 * HyoujiBairitu
    cmdPage.Width = 1095 * HyoujiBairitu
    cmdPage.Height = 495 * HyoujiBairitu
End Sub

Private Sub MENU_settei()   '*** メニュー状態設定 ***
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
    If FLGmaker = 1 Then       '*** メーカー画面存在 ***
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
'       Me.mnuHinsyuP.Enabled = True        '*** 自分なので存在しない ***
'
        If FLGindex = 1 Then
            Me.mnuHinsyu.Checked = True
            Me.mnuPmain.Enabled = True
'
'           Me.mnuHinsyuP.Checked = True    '*** 自分なので存在しない ***
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
        Me.mnuPmainP.Checked = False
        Me.mnuPmainP.Enabled = False
    End If
End Sub

Private Sub setFormArea()   '*** フォームの表示位置の設定 ***
    If Eeos2_mainMDI.ScaleHeight > Height Then
        Top = Eeos2_mainMDI.ScaleHeight - Height
    Else
        Top = 0
    End If
'
    If Eeos2_mainMDI.ScaleWidth > Width Then
        Left = (Eeos2_mainMDI.ScaleWidth - Width) \ 2
    Else
        Left = 0
    End If
End Sub

Private Sub setPointtxt(j As Integer)   '*** txtPoint.text を設定する ***
    txtPoint.Text = " L" & Bindex0(j, 0)
End Sub

Private Sub Timer1_Timer()
    Time_up = True
End Sub

Private Sub timer_waite(tm As Integer)
    Time_up = False
    Timer1.Interval = tm
    Timer1.Enabled = True
'
    Do
        DoEvents
    Loop While Time_up = False
'
    Timer1.Enabled = False
End Sub

