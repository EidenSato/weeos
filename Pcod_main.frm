VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Pcod_main 
   BackColor       =   &H00004000&
   Caption         =   "部品ｺｰﾄﾞ <品目一覧>"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   10125
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
   Icon            =   "Pcod_main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6510
   ScaleWidth      =   10125
   Begin VB.CommandButton cmdPDF 
      Caption         =   "PDF"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   21
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "？"
      Height          =   255
      Left            =   9240
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   0
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "Text8"
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Text7"
      Top             =   5760
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      MousePointer    =   1  '矢印
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Text6"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5640
      MousePointer    =   1  '矢印
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      MousePointer    =   1  '矢印
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   960
      Width           =   640
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      MousePointer    =   1  '矢印
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   375
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
      Left            =   8520
      TabIndex        =   15
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
      Left            =   5400
      TabIndex        =   14
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   7560
      TabIndex        =   13
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   6720
      TabIndex        =   12
      Top             =   5760
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "コード表ファイル出力"
      Filter          =   "*.csv"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4215
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   17
      Cols            =   7
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
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   18
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "着目部品"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "代表型名"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5040
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "親ｺｰﾄﾞ"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "項目"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "ﾌｧｲﾙ(&X)"
      Begin VB.Menu mnuFilePrn 
         Caption         =   "品目一覧印刷<縦>(&P)"
      End
      Begin VB.Menu mnuFilePrnL 
         Caption         =   "品目一覧印刷<横>(&L)"
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
         Begin VB.Menu mnuHinsyuP 
            Caption         =   "品種一覧"
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
Attribute VB_Name = "Pcod_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'******************************************
'*** ＥＥＯＳ２ 電気部品コード 品目一覧   ***
'***          2001.03.13  by S.Fukazawa ***
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
    Dim Gyoumax As Integer          '*** 印刷行最大値 ***
    Dim haba1X As Integer
    Dim haba1Y As Integer
    Dim kijunX As Integer
    Dim kijunY As Integer
    Dim Port_Land As Integer
'                                   567twip=10mm,1440twip=1inch
    Private Const OrgWidth = 10215  '*** フォーム寸法初期値 ***
    Private Const OrgHeight = 6945
    Dim tempWidth As Integer
    Dim tempHeight As Integer
'
    Private Const c10mm = 567
    Private Const gyoukan = 1440 / 4
    Private Const moji_zureX = c10mm / 5
    Private Const moji_zureY = gyoukan / 4
    Private Const pichi1 = c10mm * 1.5
    Private Const pichi2 = c10mm * 5.1
    Private Const pichi3 = c10mm * 9.95
    Private Const pichi4 = c10mm * 11.35
    Private Const pichi5 = c10mm * 12.95
    Private Const pichi6 = c10mm * 14.05
    Private Const pichi7 = c10mm * 15.4
    Private Const pichi8 = c10mm * 16.5
    Private Const pichi9 = c10mm * 17#
'
    Private Const ctuika = c10mm * 8#
'
    Dim Time_up As Boolean
'
'       // PDFﾌｧｲﾙを起動する
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Const SW_HIDE = 0
    Private Const SW_SHOWNORMAL = 1
    Private Const SW_SHOWMINIMIZED = 2
    Private Const SW_SHOWMAXIMIZED = 3
    Private Const SW_SHOWNOACTIVATE = 4
    Private Const SW_SHOW = 5
    Private Const SW_SHOWMINNOACTIVE = 7
    Private Const SW_SHOWNA = 8
    Dim DRVmainPDF  As String

Private Sub cmdHelp_Click()
    Dim temp As Integer
'
    temp = FLG_Setumei
    FLG_Setumei = 10
    Setumei_gamen.Show 1
'
    FLG_Setumei = temp
End Sub

Private Sub Form_Activate()
    Dim X As Integer
    Dim makername As String
'
    FLGmain = 1      '*** 部品コード品目画面存在フラグ設定 ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    If jcps0 <> jps0 Then   '*** 現在の内容と異なるときは再読込 ***
        jcps0 = jps0
'
        If Aitem0(ips0, 0) = "IC" Then
            DRVmain0 = Xcont0(2) & "\IC\IC" & Left$(Bindex0(jps0, 0), 1) _
                & "\L" & Bindex0(jps0, 0) & ".COD"
        Else
            DRVmain0 = Xcont0(2) & "\" & Aitem0(ips0, 0) _
                & "\L" & Bindex0(jps0, 0) & ".COD"
        End If
        Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)   '*** メインコード読み込み ***
'
        kp0 = 1             '*** 初期値設定 ***
        kps0 = kp0
        kcps0 = 0
        Call DATA_settei
        FLGmain_data_change = 0   '*** 変更フラグ初期化 ***
    End If
'
    Text1.Text = Aitem0(ips0, 0)
    Text2.Text = Aitem0(ips0, 1)
    Text3.Text = Bindex0(jps0, 1)
    Text4.Text = "L" & Bindex0(jps0, 0)
    Text6.Text = Bindex0(jps0, 3) & "xxx" & Bindex0(jps0, 4)
'
    If Bindex0(jps0, 5) = "998" Then X = 8 Else X = 5
    makername = Bindex0(jps0, X)
    Call Makerget1(makername)     '***ﾒｰｶｰ名取得 ***
    Text5.Text = makername
'
    Text1.SetFocus
End Sub

Private Sub Form_Initialize()
    kp0 = 1         '*** 初期値設定 ***
    kps0 = kp0
    kcps0 = 0
'
    HeadTitle = "部品ｺｰﾄﾞ <品目一覧>"
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
                        '*** フォームのサイズの設定
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
    If Aitem0(ips0, 0) = "IC" Then
        DRVmain0 = Xcont0(2) & "\IC\IC" & Left$(Bindex0(jps0, 0), 1) _
            & "\L" & Bindex0(jps0, 0) & ".COD"
    Else
        DRVmain0 = Xcont0(2) & "\" & Aitem0(ips0, 0) _
            & "\L" & Bindex0(jps0, 0) & ".COD"
    End If
'
    jcps0 = jps0
    Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)   '*** メインコード読み込み ***
'
    kp0 = 1             '*** 初期値設定 ***
    kps0 = kp0
    kcps0 = 0
'
    Call DSPlevel3              '*** 品目名表示 ***
    FLGmain_data_change = 0     '*** 変更フラグ初期化 ***
    Port_Land = 0               '*** ポートレートに設定 ***
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
    MSFlexGrid1.Height = MSFlexGrid1.RowHeightMin * 17
    MSFlexGrid1.Cols = 7
    MSFlexGrid1.ColWidth(0) = 445 * HyoujiBairitu
    MSFlexGrid1.ColWidth(1) = 2345 * HyoujiBairitu + FLGoffsetX \ 4
    MSFlexGrid1.ColWidth(2) = 3230 * HyoujiBairitu + (FLGoffsetX * 3) \ 4
    MSFlexGrid1.ColWidth(3) = 470 * HyoujiBairitu
    MSFlexGrid1.ColWidth(4) = 920 * HyoujiBairitu
    MSFlexGrid1.ColWidth(5) = 690 * HyoujiBairitu
    MSFlexGrid1.ColWidth(6) = 930 * HyoujiBairitu + (300 * HyoujiBairitu - 300)
'
    Call Buhin_Haichi       '*** 表示部品配置 ***
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FLGmain = 0     '*** 部品コード品種画面存在フラグ初期化 ***
    kp0 = 1         '*** フラグ初期化 ***
    kps0 = kp0      '***<再ロード時 "Form_Initialize"前に "MSFlexGrid1_Scroll" が発生してしまう予防>***
End Sub

Private Sub DSPlevel3()
                    '*** 品目一覧表示 ***
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Text = "ｺｰﾄﾞ"
        .ColAlignment(0) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 1
        .Text = "          個別型名"
        .ColAlignment(1) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 2
        .Text = "                    規格・備考"
        .ColAlignment(2) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 3
        .Text = "指定"
        .ColAlignment(3) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 4
        .Text = "平均単価"
        .ColAlignment(4) = flexAlignRightCenter     '*** 右詰め ***
        .Col = 5
        .Text = "在庫数"
        .ColAlignment(5) = flexAlignRightCenter     '*** 右詰め ***
        .Col = 6
        .Text = "更新日"
        .ColAlignment(6) = flexAlignCenterCenter    '*** 中央揃え ***
    End With
'
    Call DATA_settei
End Sub
Private Sub DATA_settei()
    Dim i As Integer
'
    If Cnum0 > 16 Then
        MSFlexGrid1.Rows = Cnum0 + 2
    Else
        MSFlexGrid1.Rows = 18
    End If
'
    For i = 1 To Cnum0
        MSFlexGrid1.Row = i
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = " -" & Cmain0(i, 0)
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = " " & Cmain0(i, 1)
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = " " & Cmain0(i, 2)
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = Cmain0(i, 3)
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Text = Format(Val(Cmain0(i, 5)), "####,###.0")
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = Format(Val(Cmain0(i, 8)), "###,###")
        MSFlexGrid1.Col = 6
        MSFlexGrid1.Text = Cmain0(i, 10)
    Next i
'
        MSFlexGrid1.Row = i '*** Cnum0+1 を消去 ***
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
        MSFlexGrid1.Col = 6
        MSFlexGrid1.Text = ""
'
    If Cnum0 < 17 Then
        For i = Cnum0 + 1 To 17
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
            MSFlexGrid1.Col = 6
            MSFlexGrid1.Text = ""
        Next i
    End If
'
    MSFlexGrid1.TopRow = kp0
    Call Realname    '*** 着目部品正式名称表示 ***
End Sub

Private Sub H1down()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Cnum0 <= 15 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Cnum0 > j + 14 Then
        MSFlexGrid1.TopRow = j + 1
    Else
        Beep
        MSFlexGrid1.TopRow = Cnum0 - 13
    End If
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub H1up()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Bnum0 <= 15 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j = 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 1
    End If
'
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub Hdown()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Cnum0 <= 15 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Cnum0 > j + 24 Then
        MSFlexGrid1.TopRow = j + 10
    Else
        Beep
        MSFlexGrid1.TopRow = Cnum0 - 13
    End If
'
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub Hup()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Cnum0 <= 15 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j - 10 < 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 10
    End If
'
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub cmdUp_Click()
    Call Hup
End Sub

Private Sub cmdDown_Click()
    Call Hdown
End Sub

Private Sub cmdPage_Click()
    FLGmain_data_change = 0     '*** 変更フラグ初期化 ***
'
    Pcod_main_c.Show 1
'                       '*** モーダルフォーム実行後、続いて実行される。***
    If FLGmain_data_change = 1 Then '*** 変更有り ***
        Call DATA_settei
        FLGmain_data_change = 0   '*** 変更フラグ初期化 ***
    End If
'
    Call Realname
    Text1.SetFocus
End Sub

Private Sub cmdPDF_Click()
    Dim nRC As Long
    
    nRC = ShellExecute(Me.hwnd, vbNullString, DRVmainPDF, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Buhin_Haichi()
                            '*** 表示部品配置 ***
    Text1.FontSize = 10 * HyoujiBairitu
    Text1.Height = 285 * HyoujiBairitu
    Text1.Left = 360 + (855 - 360) * HyoujiBairitu
    Text1.Top = 360 + FLGoffsetY
    Text1.Width = 375 * HyoujiBairitu
'
    Label1.FontSize = 10 * HyoujiBairitu
    Label1.Height = Text1.Height
    Label1.Left = 360
    Label1.Top = 360 + FLGoffsetY
    Label1.Width = 495 * HyoujiBairitu
'
    Text2.FontSize = 10 * HyoujiBairitu
    Text2.Height = 285 * HyoujiBairitu
    Text2.Left = 360 + (1320 - 360) * HyoujiBairitu
    Text2.Top = 360 + FLGoffsetY
    Text2.Width = 2295 * HyoujiBairitu
'
    Label3.FontSize = 10 * HyoujiBairitu
    Label3.Height = 270 * HyoujiBairitu
    Label3.Left = 360
    Label3.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    Label3.Width = 645 * HyoujiBairitu
'
    Text4.FontSize = 10 * HyoujiBairitu
    Text4.Height = 285 * HyoujiBairitu
    Text4.Left = 360
    Text4.Top = 360 + (960 - 360) * HyoujiBairitu + FLGoffsetY
    Text4.Width = 645 * HyoujiBairitu
'
    Label12.FontSize = 10 * HyoujiBairitu
    Label12.Height = 270 * HyoujiBairitu
    Label12.Left = 360 + (1080 - 360) * HyoujiBairitu
    Label12.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    Label12.Width = 975 * HyoujiBairitu
'
    Text6.FontSize = 10 * HyoujiBairitu
    Text6.Height = 285 * HyoujiBairitu
    Text6.Left = 360 + (1080 - 360) * HyoujiBairitu
    Text6.Top = 360 + (960 - 360) * HyoujiBairitu + FLGoffsetY
    Text6.Width = 2775 * HyoujiBairitu
'
    Label2.FontSize = 10 * HyoujiBairitu
    Label2.Height = 270 * HyoujiBairitu
    Label2.Left = 360 + (3960 - 360) * HyoujiBairitu
    Label2.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    Label2.Width = 495 * HyoujiBairitu
'
    Text3.FontSize = 10 * HyoujiBairitu
    Text3.Height = 285 * HyoujiBairitu
    Text3.Left = 360 + (3960 - 360) * HyoujiBairitu
    Text3.Top = 360 + (960 - 360) * HyoujiBairitu + FLGoffsetY
    Text3.Width = 5175 * HyoujiBairitu
'
    Text5.FontSize = 10 * HyoujiBairitu
    Text5.Height = 285 * HyoujiBairitu
    Text5.Left = 360 + (5655 - 360) * HyoujiBairitu
    Text5.Top = 360 + FLGoffsetY
    Text5.Width = 3495 * HyoujiBairitu
'
    Label10.FontSize = 10 * HyoujiBairitu
    Label10.Height = Text5.Height
    Label10.Left = 360 + (5040 - 360) * HyoujiBairitu
    Label10.Top = 360 + FLGoffsetY
    Label10.Width = 615 * HyoujiBairitu
'
    cmdHelp.FontSize = 10 * HyoujiBairitu
    cmdHelp.Height = 220 * HyoujiBairitu
    cmdHelp.Left = 360 + (9180 - 360) * HyoujiBairitu
    cmdHelp.Top = 360 + (1050 - 360) * HyoujiBairitu + FLGoffsetY
    cmdHelp.Width = 255 * HyoujiBairitu
'
    MSFlexGrid1.Left = 360
    MSFlexGrid1.Top = 360 + (1320 - 360) * HyoujiBairitu + FLGoffsetY
'
    Text7.FontSize = 10 * HyoujiBairitu
    Text7.Height = 285 * HyoujiBairitu
    Text7.Left = 360 + (1335 - 360) * HyoujiBairitu
    Text7.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    Text7.Width = 3615 * HyoujiBairitu
'
    Label4.FontSize = 10 * HyoujiBairitu
    Label4.Height = Text7.Height
    Label4.Left = 360
    Label4.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    Label4.Width = 975 * HyoujiBairitu
'
    Text8.FontSize = 10 * HyoujiBairitu
    Text8.Height = 285 * HyoujiBairitu
    Text8.Left = 360 + (1335 - 360) * HyoujiBairitu
    Text8.Top = 360 + (6120 - 360) * HyoujiBairitu + FLGoffsetY
    Text8.Width = 3255 * HyoujiBairitu
'
    Label5.FontSize = 10 * HyoujiBairitu
    Label5.Height = Text8.Height
    Label5.Left = 360 + (720 - 360) * HyoujiBairitu
    Label5.Top = 360 + (6120 - 360) * HyoujiBairitu + FLGoffsetY
    Label5.Width = 615 * HyoujiBairitu
'
    cmdPDF.FontSize = 9 * HyoujiBairitu
    cmdPDF.Height = 615 * HyoujiBairitu
    cmdPDF.Left = 360 + (4920 - 360) * HyoujiBairitu
    cmdPDF.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdPDF.Width = 255 * HyoujiBairitu
'
    cmdPage.FontSize = 9 * HyoujiBairitu
    cmdPage.Height = 495 * HyoujiBairitu
    cmdPage.Left = 360 + (5400 - 360) * HyoujiBairitu
    cmdPage.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdPage.Width = 1095 * HyoujiBairitu
'
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
    cmdUp.Left = 360 + (6720 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdUp.Width = 735 * HyoujiBairitu
'
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
    cmdDown.Left = 360 + (7560 - 360) * HyoujiBairitu
    cmdDown.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdDown.Width = 735 * HyoujiBairitu
'
    cmdClose.FontSize = 9 * HyoujiBairitu
    cmdClose.Height = 495 * HyoujiBairitu
    cmdClose.Left = 360 + (8520 - 360) * HyoujiBairitu
    cmdClose.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdClose.Width = 855 * HyoujiBairitu
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

Private Sub mnuFilePrn_Click()  '*** 品目一覧印刷 ***
    Port_Land = 0               '*** ポートレートに設定 ***
    Gyoumax = 40                '*** 印刷行最大値 ***
    haba1X = c10mm * (20.5 - 2.5)
    haba1Y = gyoukan * (Gyoumax + 3)
    kijunX = c10mm * 1.8
    kijunY = c10mm * 0.6
'
    Call Itiran_PRN
End Sub

Private Sub mnuFilePrnL_Click() '*** 品目一覧印刷 ***
    Port_Land = 1               '*** ランドスケープに設定 ***
    Gyoumax = 24                '*** 印刷行最大値 ***
    haba1X = c10mm * (28.5 - 2.5)
    haba1Y = gyoukan * (Gyoumax + 3)
    kijunX = c10mm * 1.5
    kijunY = c10mm * 2#
'
    Call Itiran_PRN
End Sub

Private Sub Itiran_PRN()
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
    Call PRNheader '*** ヘッダー印刷 (FLGgyou = 0)***
    DoEvents
'
    For kp0 = 1 To Cnum0
        If FLGgyou >= Gyoumax Then
            Call PRNfooter
            Printer.NewPage  '*** 改ページ ***
'
            Call PRNheader
        End If
'
        Call PRNkoumoku       '*** 項目印刷 ***
        DoEvents
'
    Next kp0
'
    Call PRNfooter          '*** 項末印刷 ***
    Printer.EndDoc          '*** プリンター書き込み ***
'
    Call timer_waite(1000)  '*** 表示認識待ち ***
'
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault        '*** 砂時計解除 ***
    MSFlexGrid1.MousePointer = flexArrow
    Eeos2_mainMDI.MousePointer = vbDefault
End Sub

Private Sub mnuHinsyu_Click()
    If FLGindex = 1 Then
        Pcod_index.SetFocus
    Else
'        Pcod_index.Show
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
    If FLGmain = 1 Then
        Pcod_main.SetFocus
    Else
'        Pcod_main.Show
    End If
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

Private Sub mnuAQuit_Click()
    Unload Me
    End
End Sub

Private Sub mnuBack_Click()
    FLGjob = 0
    Unload Me
End Sub

Private Sub mnuFileWR_Click()
    Dim i As Integer            '*** 品目一覧ファイル出力 ***
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
    For kp0 = 1 To Cnum0
        PDa = Chr(34) & Aitem0(ips0, 0) & Chr(34) & "," _
                & Chr(34) & "L" & Bindex0(jps0, 0) & "-" & Cmain0(kp0, 0) & Chr(34) & ","
'
        If Bindex0(jps0, 4) = "*" Then
            PD = Bindex0(jps0, 3) & Cmain0(kp0, 1)
        Else
            PD = Bindex0(jps0, 3) & Cmain0(kp0, 1) & Bindex0(jps0, 4)
        End If
'
        PDa = PDa & Chr(34) & PD & Chr(34) & "," _
                & Chr(34) & Bindex0(jps0, 1) & Chr(34) & "," _
                & Chr(34) & Cmain0(kp0, 2) & Chr(34) & ","
        PD = Cmain0(kp0, 3)
            Call TRSsitei2(PD)
        PDa = PDa & Chr(34) & PD & Chr(34) & "," _
                & Chr(34) & str(Val(Cmain0(kp0, 5))) & Chr(34) & ","
'
        If Bindex0(jps0, 5) <> "000" Then
            PD = Bindex0(jps0, 5)
            Call Makerget2(PD)
        Else
            PD = Cmain0(kp0, 13)
            Call Makerget2(PD)
        End If
'
        PDa = PDa & Chr(34) & PD & Chr(34) & "," _
                & Chr(34) & str(Val(Cmain0(kp0, 8))) & Chr(34) & "," _
                & Chr(34) & Cmain0(kp0, 10) & Chr(34) & ","
'
        PD = Cmain0(kp0, 12)
            Call TRSsyukko2(PD)
        PDa = PDa & Chr(34) & PD & Chr(34) & ","
'
        PD = Cmain0(kp0, 17)
            Call TRStouroku2(PD)
        PDa = PDa & Chr(34) & PD & Chr(34) & ","
'
        PD = Cmain0(kp0, 18)
            Call TRSkeijou2(PD)
        PDa = PDa & Chr(34) & PD & Chr(34)
'
        Print #FILE_number, PDa  '*** データ書き込み ***
        DoEvents
    Next kp0
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
    Tyuusin = Cnum0 \ 2
    If Tyuusin <= 8 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Tyuusin - 7
    End If
'
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub mnuJumpE_Click()
    If Cnum0 < 16 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Cnum0 - 14
    End If
'
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub mnuJumpT_Click()
    MSFlexGrid1.TopRow = 1
'
    kp0 = MSFlexGrid1.TopRow + 1
    kps0 = kp0
'
    Call Realname   '*** 着目部品正式名称表示 ***
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 3          '*** 部品コード表フラグ ***
    Call mnuSousaSetumei
End Sub

Private Sub mnuVersion_Click()
    Call mnuVersionGamen
End Sub

Private Sub PRNfooter()
'                   *** フッター印刷 ***
'
    Printer.CurrentX = kijunX + c10mm
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    SETfont_size 9, 1
    Printer.Print "M5103-03";
'
    Printer.CurrentX = kijunX + haba1X - c10mm * 4
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    SETfont_size 9, 0
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷";
End Sub

Private Sub PRNkoumoku()
'                   *** 項目印刷 ***
    Dim makername As String, PD As String
    Dim tuika As Integer, tyousetu As Integer, bityou As Integer
'
    If Port_Land = 0 Then
        tuika = 0
        tyousetu = 0
        bityou = 0
    Else
        tuika = ctuika
        tyousetu = c10mm
        bityou = moji_zureX / 2
    End If
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    Printer.Print "L" & Bindex0(jps0, 0) & "-" & Cmain0(kp0, 0)
'
    Printer.CurrentX = kijunX + pichi1 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    If Bindex0(jps0, 4) = "*" Then
        PD = Bindex0(jps0, 3) & Cmain0(kp0, 1)
    Else
        PD = Bindex0(jps0, 3) & Cmain0(kp0, 1) & Bindex0(jps0, 4)
    End If
    Printer.Print PD
'
    Printer.CurrentX = kijunX + pichi2 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    Printer.Print Cmain0(kp0, 2)
'
    Printer.CurrentX = kijunX + pichi3 + tuika - tyousetu - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    PD = Cmain0(kp0, 3)
    If Port_Land = 0 Then
        Call TRSsitei3(PD)  '*** <部品>無し ***
    Else
        Call TRSsitei2(PD)
    End If
    Printer.Print PD
'
    Printer.CurrentX = kijunX + pichi4 + tuika - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
        Call SET_migiyose(Format(Cmain0(kp0, 5), "###,##0.0"), PD, 9)
    Printer.Print PD
'
    Printer.CurrentX = kijunX + pichi5 + tuika - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    If Bindex0(jps0, 5) <> "000" Then
        PD = Bindex0(jps0, 5)
        Call Makerget2(PD)
    Else
        PD = Cmain0(kp0, 13)
        Call Makerget2(PD)
    End If
    Printer.Print PD
'
    Printer.CurrentX = kijunX + pichi6 + tuika - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    Printer.Print Cmain0(kp0, 10)
'
    Printer.CurrentX = kijunX + pichi7 + tuika - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    PD = Cmain0(kp0, 12)
        Call TRSsyukko2(PD)
    If Len(PD) = 2 Then
        PD = " " & PD
    End If
    Printer.Print PD
'
    Printer.CurrentX = kijunX + pichi8 + tuika - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    PD = Cmain0(kp0, 17)
        Call TRStouroku2(PD)
    Printer.Print PD
'
    Printer.CurrentX = kijunX + pichi9 + tuika - bityou + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 3) + moji_zureY
    PD = Cmain0(kp0, 18)
        Call TRSkeijou2(PD)
    Printer.Print PD
'
    FLGgyou = FLGgyou + 1
End Sub

Private Sub PRNheader()
'                   *** 項目一覧表ヘッダー印刷 ***
    Dim i As Integer
    Dim tuika As Integer
'
    If Port_Land = 0 Then
        tuika = 0
    Else
        tuika = ctuika
    End If
'
    FLGgyou = 0
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    If Port_Land = 0 Then
        Printer.Orientation = vbPRORPortrait    '*** ポートレート   210x290 ***
    Else
        Printer.Orientation = vbPRORLandscape   '*** ランドスケープ 290x210 ***
    End If
'
    Printer.CurrentX = kijunX + haba1X + tuika - c10mm * 2
    Printer.CurrentY = kijunY - gyoukan + moji_zureY
    SETfont_size 9, 1      '*** フォント,サイズ設定 ***
    Printer.Print "ﾍﾟｰｼﾞ： " & FLGpage
    FLGpage = FLGpage + 1
'
    Printer.CurrentX = kijunX + tuika / 2 + c10mm / 2
    Printer.CurrentY = kijunY + moji_zureY
    SETfont_size 10, 1      '*** フォント,サイズ設定 ***
    Printer.Print "===  [  営電標準  ]       電気部品コード・単価表                        ＜コード一覧表＞  ==="
'
    Printer.CurrentX = kijunX + tuika / 2 + c10mm * 10
    Printer.CurrentY = kijunY + moji_zureY
    SETfont_size 9, 1      '*** フォント,サイズ設定 ***
    Printer.Print Aitem0(ips0, 1)
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan + moji_zureY
    Printer.Print "  " & Bindex0(jps0, 1)
'
    Printer.CurrentX = kijunX + c10mm * 8 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan + moji_zureY
    Printer.Print Bindex0(jps0, 7)
'
    Printer.CurrentX = kijunX + tuika + c10mm * 15 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan + moji_zureY
    Printer.Print "表記： " & Bindex0(jps0, 6)
'
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * 2 + moji_zureY
'
    If Port_Land = 0 Then
        Printer.Print "部品ｺｰﾄﾞ      型名                    規格・備考                指定     単価     ﾒｰｶｰ   更新日  出庫 CAD 形状"
    Else
        Printer.Print "部品ｺｰﾄﾞ      型名                    規格・備考                                                            指定          単価     ﾒｰｶｰ   更新日  出庫 CAD 形状"
    End If
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
'
    Printer.Line (kijunX, kijunY)-Step(haba1X, haba1Y), 0, B
'
    Printer.Line (kijunX, kijunY + gyoukan)-Step(haba1X, gyoukan), 0, B
'
    Printer.Line (kijunX, kijunY + gyoukan * 2)-Step(haba1X, gyoukan), 0, B
End Sub

Private Sub Realname()
    Dim makername As String
'
    If Bindex0(jps0, 4) = "*" Then
        Text7.Text = "L" & Bindex0(jps0, 0) & "-" & Cmain0(kps0, 0) & ": " & Bindex0(jps0, 3) & Cmain0(kps0, 1)
    Else
        Text7.Text = "L" & Bindex0(jps0, 0) & "-" & Cmain0(kps0, 0) & ": " & Bindex0(jps0, 3) & Cmain0(kps0, 1) & Bindex0(jps0, 4)
    End If
'
    If Bindex0(jps0, 5) = "000" Then
        makername = Cmain0(kps0, 13)
        Call Makerget1(makername)     '***ﾒｰｶｰ名取得 ***
        Text8.Text = makername
'
        Label5.Visible = True
        Text8.Visible = True
    Else
        Label5.Visible = False
        Text8.Visible = False
    End If
'
    Call setPDFcommand
End Sub

Private Sub setPDFcommand()
    Dim PDFari As Boolean
    Dim FILE_number As Integer
'
    DRVmainPDF = Xcont0(2) & "\0_spec\L" & Bindex0(jps0, 0) & "-" & Cmain0(kps0, 0) & ".PDF"
'
    On Error GoTo errh_NoA
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVmainPDF For Input As #FILE_number
'
    Close #FILE_number
    PDFari = True
    GoTo next9
'
errh_NoA:
    Resume next1
next1:
    DRVmainPDF = Xcont0(2) & "\0_spec\L" & Bindex0(jps0, 0) & "-XX.PDF"
'
    On Error GoTo errh_NoB
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVmainPDF For Input As #FILE_number
'
    Close #FILE_number
    PDFari = True
    GoTo next9
'
errh_NoB:
    Resume next2
next2:
    PDFari = False
'
next9:
    On Error GoTo 0
    If PDFari = True Then
        cmdPDF.Visible = True
    Else
        cmdPDF.Visible = False
    End If
End Sub

Private Sub setFormArea()   '*** フォームの表示位置の設定 ***
    If Eeos2_mainMDI.ScaleHeight > (Height + 420) Then
        Top = Eeos2_mainMDI.ScaleHeight - Height - 420
    Else
        Top = 0
    End If
'
    If Eeos2_mainMDI.ScaleWidth > Width Then
        Left = Eeos2_mainMDI.ScaleWidth - Width
    Else
        Left = 0
    End If
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
    If FLGplst2 = 1 Then         '*** 部品表画面２存在 ***
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
'
            If FLGmain = 1 Then
                Me.mnuPmain.Checked = True
'               Me.mnuPmainP.Checked = True '*** 自分なので存在しない ***
            Else
                Me.mnuPmain.Checked = False
            End If
        Else
            Me.mnuHinsyu.Checked = False
            Me.mnuPmain.Enabled = False
            Me.mnuPmain.Checked = False
'
            Me.mnuHinsyuP.Checked = False
        End If
    Else
'       Me.mnuCode.Checked = False
'       Me.mnuCode.Enabled = True
'       Me.mnuHinsyu.Checked = False
'       Me.mnuHinsyu.Enabled = False
        Me.mnuPmain.Checked = False
        Me.mnuPmain.Enabled = False
'
'       Me.mnuHinsyuP.Checked = False
'       Me.mnuHinsyuP.Enabled = False
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    If MSFlexGrid1.Row <= Cnum0 Then
        kps0 = MSFlexGrid1.Row
'
        Call Realname   '*** 着目部品正式名称表示 ***
    End If
    Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row <= Cnum0 Then
        kps0 = MSFlexGrid1.Row
        Call Realname   '*** 着目部品正式名称表示 ***
'
        Call cmdPage_Click
    End If
    Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub MSFlexGrid1_Scroll()
    kp0 = MSFlexGrid1.TopRow
    kps0 = kp0
    Call Realname   '*** 着目部品正式名称表示 ***
'
    Text1.SetFocus
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

