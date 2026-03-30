VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Const_main 
   BackColor       =   &H00004000&
   Caption         =   "EＥOS 電気 構成表"
   ClientHeight    =   8055
   ClientLeft      =   645
   ClientTop       =   1110
   ClientWidth     =   10095
   FillColor       =   &H0000C000&
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
   Icon            =   "Const_main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   8055
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdHelp 
      Caption         =   "？"
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "<電気部品表ﾁｪｯｸ>"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtPrinting 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0000C000&
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Const_main.frx":030A
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtkbikou 
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
      Height          =   495
      IMEMode         =   4  '全角ひらがな
      Left            =   1320
      MousePointer    =   1  '矢印
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   6480
      Width           =   8415
   End
   Begin VB.CommandButton cmdsyuuryou 
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
      Left            =   8640
      TabIndex        =   17
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txttantou 
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
      Height          =   285
      IMEMode         =   4  '全角ひらがな
      Left            =   6840
      MousePointer    =   1  '矢印
      TabIndex        =   11
      Text            =   "深澤"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtdaisuu 
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
      Height          =   285
      Left            =   6840
      MousePointer    =   1  '矢印
      TabIndex        =   9
      Text            =   "1"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtkouban 
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
      Height          =   285
      Left            =   6840
      MousePointer    =   1  '矢印
      TabIndex        =   7
      Text            =   "M12-1234-12/13"
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtbangou 
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
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   5
      Text            =   "1234"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtmeisyou 
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
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   3
      Text            =   "TV SOUND MULTIPLEX MODULATOR"
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtKeisiki 
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
      Height          =   285
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   360
      TabIndex        =   18
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   21
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
   Begin VB.Label lbl_curr_file 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "現在のﾌｧｲﾙ名 ："
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
      Height          =   255
      Left            =   2880
      MousePointer    =   1  '矢印
      TabIndex        =   16
      Top             =   7200
      Width           =   3255
   End
   Begin VB.Label lblbikou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "コメント (&B)"
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
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lbltantou 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "技術担当者(&6)"
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
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbldaisuu 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "台  数  (&5)"
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
      Height          =   285
      Left            =   5280
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblkouban 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "工  番  (&4)"
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
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblbangou 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "図面番号(&3)"
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
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblmeisyou 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "名  称 (&2)"
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
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblkeisiki 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "形  式 (&1)"
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
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "ﾌｧｲﾙ(&X)"
      Begin VB.Menu mnuSinki 
         Caption         =   "新規作成(&N)"
      End
      Begin VB.Menu mnu区切り線11 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileSV 
         Caption         =   "上書き保存(&S)"
      End
      Begin VB.Menu mnufileNW 
         Caption         =   "名前を付けて保存(&R)..."
      End
      Begin VB.Menu mnu区切り線12 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileRD 
         Caption         =   "同一ﾌｧｲﾙ再読込(&C)"
      End
      Begin VB.Menu mnufileNR 
         Caption         =   "別ﾌｧｲﾙ読込(&F)..."
      End
      Begin VB.Menu mnu区切り線13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrn 
         Caption         =   "構成表印刷(&P)..."
      End
      Begin VB.Menu mnuFileWR 
         Caption         =   "構成表ﾌｧｲﾙ出力(&W)"
      End
      Begin VB.Menu mnuPurFileWR 
         Caption         =   "構成表資材ﾌｧｲﾙ(&P)"
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
      Enabled         =   0   'False
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
      Begin VB.Menu mnuPuKouseihyou 
         Caption         =   "構成表"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSinkiP 
         Caption         =   "新規作成"
      End
      Begin VB.Menu mnu区切り線921 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileSVP 
         Caption         =   "上書き保存"
      End
      Begin VB.Menu mnufileNWP 
         Caption         =   "名前を付けて保存..."
      End
      Begin VB.Menu mnu区切り線922 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileRDP 
         Caption         =   "同一ﾌｧｲﾙ再読込"
      End
      Begin VB.Menu mnufileNRP 
         Caption         =   "別ﾌｧｲﾙ読込..."
      End
      Begin VB.Menu mnu区切り線923 
         Caption         =   "-"
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
         Begin VB.Menu mnu区切り線931 
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
Attribute VB_Name = "Const_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'********************
'*** 構成表メイン ***
'********************
'
Option Explicit
'
    Dim HeadTitle As String
'
    Dim FLGoffsetX As Integer
    Dim FLGoffsetY As Integer
'                                   567twip=10mm,1440twip=1inch
    Private Const OrgWidth = 10215  '*** フォーム寸法初期値 ***
    Private Const OrgHeight = 8415
    Dim tempWidth As Integer
    Dim tempHeight As Integer
'
    Dim CHaiki As String
    Dim kijunX As Integer, kijunY As Integer
    Dim habaX As Integer, habaY As Integer
    Dim haba1X As Integer, haba2X As Integer, haba3X As Integer
    Dim haba4X As Integer, gyoukan As Integer
    Dim moji_zureX As Integer, moji_zureY As Integer
    Dim Tmpdata As String
    Dim G_row   As Integer
'
    Dim DRVconst As String       '*** 構成表ディレクトリ ***
    Dim CATno As String          '*** 型名 ***
    Dim CATname As String        '*** 品名 ***
    Dim Zuban As String          '*** 図番 ***
    Dim Person As String         '*** 担当者 ***
    Dim Orgdate As String        '*** 作成日 ***
    Dim Revdate As String        '*** 修正日 ***
    Dim Checkdate As String      '*** 数量表計算日 ***
    Dim Outdate As String        '*** ラベル印刷日 ***
    Dim Ktotal As Integer        '*** 構成表配列数 ***
    Dim Kdim As Integer          '*** 構成表次元数 ***
    Dim KLST() As String         '*** 構成表データ配列 ***
    Dim Kouban As String         '*** 工番 ***
    Dim Daisuu As String         '*** 台数 ***
    Dim Kbikou As String         '*** 備考欄 ***
    Dim KyobiA As String         '*** 予備１ ***
    Dim KyobiB As String         '*** 予備２ ***

Private Sub Form_Activate()
    FLGconst = 1        '*** 構成表画面存在フラグ設定 ***
    FLGjob = 1
    STATUS = HeadTitle  '*** 選択ウインドウのタイトル名称 ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    txtKeisiki.SetFocus
End Sub

Private Sub Form_Initialize()
    HeadTitle = STATUS
    FLGconst = 1
End Sub

Private Sub Form_Load()
    tempWidth = 360 + (OrgWidth - 720) * HyoujiBairitu + 360
    tempHeight = 240 + (OrgHeight - 480) * HyoujiBairitu + 240
                                '*** これで「Form_Resize」割り込みが発生する。 ***
    Me.Width = tempWidth
    Me.Height = tempHeight
'
'   FLGoffsetX = 0      '*** 初期化 ***
'   FLGoffsetY = 0
'
    Call setFormArea    '*** フォームの表示位置の設定
'
    Me.Caption = HeadTitle
'
    CHaiki = "構成表は変更されています。「廃棄終了」をキャンセルしますか？"
'
    If FLGshinki = 1 Then
        DSPconst0N      '*** 新規の初期化 ***
        DSPconst1N
    Else
        Call Copy_const_Temp2work
    End If
'
    DSPconst0
    DSPconst1
'
    FLGchange = 0       '*** 変更フラグクリアー ***
    FLGowari = 0        '*** 終了フラグりセット ***
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
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FLGconst = 0    '*** 構成表画面存在フラグ初期化 ***
'    kp0 = 1         '*** フラグ初期化 ***
'    kps0 = kp0      '***<再ロード時 "Form_Initialize"前に "MSFlexGrid1_Scroll" が発生してしまう予防>***
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
'
    If UnloadMode = vbFormControlMenu Then
        Call mnuBack_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub Copy_const_Temp2work()
    DRVconst = DRVconstT
    CATno = CATnoT
    CATname = CATnameT
    Zuban = ZubanT
    Person = PersonT
    Orgdate = OrgdateT
    Revdate = RevdateT
    Checkdate = CheckdateT
    Outdate = OutdateT
'
    Ktotal = KtotalT
    Kdim = KdimT
    ReDim KLST(Ktotal, Kdim) As String
'
    KLST() = KLSTT()
    Kouban = KoubanT
    Daisuu = DaisuuT
    Kbikou = KbikouT
    KyobiA = KyobiAT
    KyobiB = KyobiBT
End Sub

Private Sub Copy_const_work2temp()
    DRVconstT = DRVconst
    CATnoT = CATno
    CATnameT = CATname
    ZubanT = Zuban
    PersonT = Person
    OrgdateT = Orgdate
    RevdateT = Revdate
    CheckdateT = Checkdate
    OutdateT = Outdate
'
    KtotalT = Ktotal
    KdimT = Kdim
    ReDim KLSTT(KtotalT, KdimT) As String
'
    KLSTT() = KLST()
    KoubanT = Kouban
    DaisuuT = Daisuu
    KbikouT = Kbikou
    KyobiAT = KyobiA
    KyobiBT = KyobiB
End Sub

Private Sub DSPconst0()
    txtKeisiki.Text = " " & CATno
    txtmeisyou.Text = " " & CATname
    txtbangou.Text = " " & Zuban
    txtkouban.Text = " " & Kouban
    txtdaisuu.Text = " " & Daisuu
    txttantou.Text = " " & Person
    If Trim(Kbikou) <> "*" Then
        txtkbikou.Text = Kbikou
    End If
    lbl_curr_file.Caption = "現在のﾌｧｲﾙ名： " & CURR_file
End Sub

Private Sub DSPconst0N()
'               *** 構成表 要素 初期化 ***
    CATno = "*"
    CATname = "*"
    Zuban = "*"
    Person = "*"
    Orgdate = Format(Date, "yy/mm/dd")
    Revdate = "*"
    Checkdate = "*"
    Outdate = "*"
    Kouban = "*"
    Daisuu = 0
    Kbikou = "*"
    KyobiA = "*"
    KyobiB = "*"
End Sub

Private Sub DSPconst1()
'               *** 定義 ***
    Dim i As Integer, iend As Integer
'
    If Ktotal < G_row Then
        MSFlexGrid1.Rows = G_row + 2    '*** タイトル行＋予備１行 ***
        iend = G_row + 1
    Else
        MSFlexGrid1.Rows = Ktotal + 2
        iend = Ktotal + 1
    End If
'               *** 表題 ***
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Text = "行"
        .ColAlignment(0) = flexAlignCenterCenter    '*** 中央 ***
        .Col = 1
        .Text = "            名  称 "
        .ColAlignment(1) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 2
        .Text = "      回路図"
        .ColAlignment(2) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 3
        .Text = "   電気部品表"
        .ColAlignment(3) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 4
        .Text = "    ﾊﾟﾀｰﾝ番号"
        .ColAlignment(4) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 5
        .Text = "数"
        .ColAlignment(5) = flexAlignCenterCenter    '*** 中央 ***
        .Col = 6
        .Text = "         備  考"
        .ColAlignment(6) = flexAlignLeftCenter      '*** 左詰め ***
    End With
'               *** 内容表示 ***
    For i = 1 To Ktotal
        With MSFlexGrid1
            .Row = i
            .Col = 0
            .Text = i
'
            .Col = 1
                If Trim(KLST(i, 0)) = "〃" Then
            .Text = "     " & Trim(KLST(i, 0))
                ElseIf Trim(KLST(i, 0)) = "*" Or Trim(KLST(i, 0)) = "-" Then
            .Text = "  --------"
                Else
            .Text = " " & KLST(i, 0)
                End If
'
            .Col = 2
                If Trim(KLST(i, 1)) = "〃" Then
            .Text = "     " & Trim(KLST(i, 1))
                ElseIf Trim(KLST(i, 1)) = "*" Or Trim(KLST(i, 1)) = "-" Then
            .Text = "  --------"
                Else
            .Text = " " & KLST(i, 1)
                End If
'
            .Col = 3
                If Trim(KLST(i, 2)) = "〃" Then
            .Text = "     " & Trim(KLST(i, 2))
                ElseIf Trim(KLST(i, 2)) = "*" Or Trim(KLST(i, 2)) = "-" Then
            .Text = "  --------"
                Else
            .Text = " " & KLST(i, 2)
                End If
'
            .Col = 4
                If Trim(KLST(i, 3)) = "〃" Then
            .Text = "     " & Trim(KLST(i, 3))
                ElseIf Trim(KLST(i, 3)) = "*" Or Trim(KLST(i, 3)) = "-" Then
            .Text = "  --------"
                Else
            .Text = " " & KLST(i, 3)
                End If
'
            .Col = 5
            .Text = Trim(KLST(i, 4))
'
            .Col = 6
            .Text = " " & Trim(KLST(i, 5))
        End With
    Next i
'
    For i = Ktotal + 1 To iend
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
        End With
    Next i
End Sub

Private Sub DSPconst1N()
    Dim i As Integer
'
    Ktotal = 1
    Kdim = Ckdim
'
    ReDim KLST(Ktotal, Kdim) As String
'
    For i = 0 To Kdim
        KLST(1, i) = "*"
    Next i
End Sub

Private Sub Kurisage()
    Dim i As Integer, j As Integer
'
    For i = Ktotal - 1 To TMProw Step -1
        For j = 0 To Kdim
            KLST(i + 1, j) = KLST(i, j)
        Next j
    Next i
End Sub

Private Sub H1up()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If j = 1 Then
        Beep
    ElseIf Ktotal <= G_row Then
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
    If Ktotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Ktotal - (j + 10) >= G_row Then
        MSFlexGrid1.TopRow = j + 10
    Else
        Beep
        MSFlexGrid1.TopRow = Ktotal - G_row + 1
    End If
End Sub

Private Sub Hup()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Ktotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j - 10 <= 0 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 10
    End If
End Sub

Private Sub H1down()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Ktotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Ktotal - j >= G_row Then
        MSFlexGrid1.TopRow = j + 1
    Else
        Beep
    End If
End Sub

Private Sub PRNheader0()
'                   567twip=10mm,1440twip=1inch
    kijunX = 720
    kijunY = 1134
    habaX = 10500
    habaY = 1531
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    Printer.Orientation = vbPRORPortrait    '*** ﾎﾟｰﾄﾚｰﾄ ***
'
    Printer.CurrentX = kijunX + 3118
    Printer.CurrentY = 800
    SETfont_size 17, 1    '*** フォント,サイズ設定 ***
    Printer.Print "《　電  気  構　成　表　》"
'
    Printer.CurrentX = kijunX + 9300
    Printer.CurrentY = 920
    SETfont_size 10.8, 1    '*** フォント,サイズ設定 ***
    If Val(Left(Revdate, 2)) > 85 Then
        Printer.Print "19" & Revdate
    Else
        Printer.Print "20" & Revdate
    End If
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
    Printer.Line (kijunX, kijunY)-Step(habaX, habaY), 0, B
'
    Printer.CurrentX = kijunX + 567
    Printer.CurrentY = kijunY + 180
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Printer.Print "型    式    " & CATno
'
    Printer.CurrentX = kijunX + 567
    Printer.CurrentY = kijunY + 180 + 482
    Printer.Print "名    称    " & CATname
'
    Printer.CurrentX = kijunX + 567
    Printer.CurrentY = kijunY + 180 + 482 + 482
    Printer.Print "図面番号      " & Zuban
'
    Printer.CurrentX = kijunX + 6600
    Printer.CurrentY = kijunY + 180
    Printer.Print "工  番   " & Kouban
'
    Printer.CurrentX = kijunX + 6600
    Printer.CurrentY = kijunY + 180 + 482
    Printer.Print "台  数      " & Daisuu
'
    Printer.CurrentX = kijunX + 6600
    Printer.CurrentY = kijunY + 180 + 482 + 340
    Printer.Print " 技術"
'
    Printer.CurrentX = kijunX + 6600
    Printer.CurrentY = kijunY + 180 + 482 + 340 + 220
    Printer.Print "担当者"
'
    Printer.CurrentX = kijunX + 7626
    Printer.CurrentY = kijunY + 1144
    Printer.Print Person
'
    Printer.CurrentX = kijunX + 9623
    Printer.CurrentY = kijunY + 1300
    Printer.Print "検査"
End Sub

Private Sub PRNheader1(startX As Integer, startY As Integer)
'                   567twip=10mm,1440twip=1inch
    Dim X As Integer, Y As Integer
'
    haba1X = 2672
    haba2X = 1814
    haba3X = 454
    haba4X = 1932
    gyoukan = 480
    moji_zureX = 113
    moji_zureY = 128
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
    Printer.Line (startX, startY)-Step(haba1X, gyoukan), 0, B
    Printer.Line (startX + haba1X, startY)-Step(haba2X, gyoukan), 0, B
    Printer.Line (startX + haba1X + haba2X, startY)-Step(haba2X, gyoukan), 0, B
    Printer.Line (startX + haba1X + haba2X * 2, startY)-Step(haba2X, gyoukan), 0, B
    Printer.Line (startX + haba1X + haba2X * 3, startY)-Step(haba3X, gyoukan), 0, B
    Printer.Line (startX + haba1X + haba2X * 3 + haba3X, startY)-Step(haba4X, gyoukan), 0, B
'
    Printer.CurrentX = startX + 808
    Printer.CurrentY = startY + moji_zureY
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Printer.Print "名      称"
'
    Printer.CurrentX = startX + haba1X + 354
    Printer.CurrentY = startY + moji_zureY
    Printer.Print "回  路  図"
'
    Printer.CurrentX = startX + haba1X + haba2X + 368
    Printer.CurrentY = startY + moji_zureY
    Printer.Print "電気部品表"
'
    Printer.CurrentX = startX + haba1X + haba2X * 2 + 255
    Printer.CurrentY = startY + moji_zureY
    Printer.Print "パターン番号"
'
    Printer.CurrentX = startX + haba1X + haba2X * 3 + 118
    Printer.CurrentY = startY + moji_zureY
    Printer.Print "数"
'
    Printer.CurrentX = startX + haba1X + haba2X * 3 + haba3X + 595
    Printer.CurrentY = startY + moji_zureY
    Printer.Print "備   考"
End Sub

Private Sub PRNkoumoku(startX As Integer, startY As Integer)
    Dim i As Integer, j As Integer
    Dim s_itiY As Integer
    Dim Pdata As String
'
    s_itiY = startY
    j = 1               '*** ページフラグ ***
    For i = 1 To Ktotal
        Printer.DrawWidth = 6
        Printer.FillStyle = vbFSTransparent
        Printer.Line (startX, s_itiY)-Step(haba1X, gyoukan), 0, B
        Printer.Line (startX + haba1X, s_itiY)-Step(haba2X, gyoukan), 0, B
        Printer.Line (startX + haba1X + haba2X, s_itiY)-Step(haba2X, gyoukan), 0, B
        Printer.Line (startX + haba1X + haba2X * 2, s_itiY)-Step(haba2X, gyoukan), 0, B
        Printer.Line (startX + haba1X + haba2X * 3, s_itiY)-Step(haba3X, gyoukan), 0, B
        Printer.Line (startX + haba1X + haba2X * 3 + haba3X, s_itiY)-Step(haba4X, gyoukan), 0, B
'                       *** 名称 ***
        Printer.CurrentX = startX + moji_zureX
        Printer.CurrentY = s_itiY + moji_zureY
        SETfont_size 10.8, 1
        If Trim(KLST(i, 0)) = "〃" Then
            Printer.Print "     " & Trim(KLST(i, 0))
        ElseIf Trim(KLST(i, 0)) = "*" Or Trim(KLST(i, 0)) = "-" Then
            Printer.Print "   --------"
        Else
            Printer.Print KLST(i, 0)
        End If
'                       *** 回路図 ***
        Printer.CurrentX = startX + haba1X + moji_zureX
        Printer.CurrentY = s_itiY + moji_zureY
        If Trim(KLST(i, 1)) = "〃" Then
            Printer.Print "     " & Trim(KLST(i, 1))
        ElseIf Trim(KLST(i, 1)) = "*" Or Trim(KLST(i, 1)) = "-" Then
            Printer.Print "   --------"
        Else
            Printer.Print KLST(i, 1)
        End If
'                       *** 電気部品表 ***
        Printer.CurrentX = startX + haba1X + haba2X + moji_zureX
        Printer.CurrentY = s_itiY + moji_zureY
        If Trim(KLST(i, 2)) = "〃" Then
            Printer.Print "     " & Trim(KLST(i, 2))
        ElseIf Trim(KLST(i, 2)) = "*" Or Trim(KLST(i, 2)) = "-" Then
            Printer.Print "   --------"
        Else
            Printer.Print KLST(i, 2)
        End If
'                       *** パターン番号 ***
        Printer.CurrentX = startX + haba1X + haba2X * 2 + moji_zureX
        Printer.CurrentY = s_itiY + moji_zureY
        If Trim(KLST(i, 3)) = "〃" Then
            Printer.Print "     " & Trim(KLST(i, 3))
        ElseIf Trim(KLST(i, 3)) = "*" Or Trim(KLST(i, 3)) = "-" Then
            Printer.Print "   --------"
        Else
            Printer.Print KLST(i, 3)
        End If
'                       *** 数 ***
        Printer.CurrentY = s_itiY + moji_zureY
        If Val(KLST(i, 4)) > 99 Then
            Printer.CurrentX = startX + haba1X + haba2X * 3 + 68
            Printer.Print Trim(KLST(i, 4))
        ElseIf Val(KLST(i, 4)) > 9 Then
            Printer.CurrentX = startX + haba1X + haba2X * 3 + 119
            Printer.Print Trim(KLST(i, 4))
        Else
            Printer.CurrentX = startX + haba1X + haba2X * 3 + 176
            Printer.Print Trim(KLST(i, 4))
        End If
'                       *** 備考 ***
        Printer.CurrentX = startX + haba1X + haba2X * 3 + haba3X + moji_zureX
        Printer.CurrentY = s_itiY + moji_zureY
        If Trim(KLST(i, 5)) = "*" Then
            Printer.Print " "
        Else
            Printer.Print KLST(i, 5)
        End If
'
        s_itiY = s_itiY + gyoukan
'
        If (j = 1 And i Mod 25 = 0 And Ktotal <> 25) Or (j > 1 And (i - 25) Mod 28 = 0) Then
            Printer.CurrentX = startX + 300
            Printer.CurrentY = s_itiY + 57
            SETfont_size 8, 1
            Printer.Print "M5903-04"
'
            Printer.CurrentX = startX + 4422
            Printer.CurrentY = s_itiY + 57
            SETfont_size 10.8, 0
            Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷"
'
            Printer.CurrentX = startX + 9639
            Printer.CurrentY = s_itiY + 57
            Printer.Print "つづく"
'
            Printer.NewPage
'
            j = j + 1
            Printer.CurrentX = kijunX + 1418
            Printer.CurrentY = kijunY - Int(gyoukan / 2)
            SETfont_size 10.8, 0
            Printer.Print CATno     '*** 型番印刷 ***
'
            Printer.CurrentX = kijunX + 3969
            Printer.CurrentY = kijunY - Int(gyoukan / 2)
            Printer.Print CATname   '*** 名称印刷 ***
'
            Printer.CurrentX = kijunX + 9356
            Printer.CurrentY = kijunY - Int(gyoukan / 2)
            Printer.Print j; "ページ"
'
            Call PRNheader1(kijunX, kijunY) '*** 構成表項目見出し印刷 ***
'
            s_itiY = kijunY + gyoukan
        End If
    Next i
'
    If Ktotal < 11 Then
        For i = Ktotal + 1 To 11
            Printer.DrawWidth = 6
            Printer.FillStyle = vbFSTransparent
            Printer.Line (startX, s_itiY)-Step(haba1X, gyoukan), 0, B
            Printer.Line (startX + haba1X, s_itiY)-Step(haba2X, gyoukan), 0, B
            Printer.Line (startX + haba1X + haba2X, s_itiY)-Step(haba2X, gyoukan), 0, B
            Printer.Line (startX + haba1X + haba2X * 2, s_itiY)-Step(haba2X, gyoukan), 0, B
            Printer.Line (startX + haba1X + haba2X * 3, s_itiY)-Step(haba3X, gyoukan), 0, B
            Printer.Line (startX + haba1X + haba2X * 3 + haba3X, s_itiY)-Step(haba4X, gyoukan), 0, B
            s_itiY = s_itiY + gyoukan
        Next i
    End If
'
    Printer.CurrentX = startX + 300
    Printer.CurrentY = s_itiY + 57
    SETfont_size 8, 1
    Printer.Print "M5903-04"
'
    Printer.CurrentX = startX + 4422
    Printer.CurrentY = s_itiY + 57
    SETfont_size 10.8, 0
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷"
'
    If Kbikou <> "" Then
        s_itiY = s_itiY + 567
        Printer.CurrentX = startX
        Printer.CurrentY = s_itiY
        SETfont_size 10.8, 0
'
        i = InStr(Kbikou, " _")
        If i = 0 Then
            Printer.Print " 備考： " & Kbikou
        Else
            Printer.Print " 備考： " & Left(Kbikou, i - 1)
            Pdata = Mid(Kbikou, i + 2)
            i = InStr(Pdata, " _")
'
            Do While i > 0
                s_itiY = s_itiY + 340
                Printer.CurrentX = startX
                Printer.CurrentY = s_itiY
                Printer.Print " 　　　 " & Left(Pdata, i - 1)
                Pdata = Mid(Pdata, i + 2)
                i = InStr(Pdata, " _")
            Loop
'
            s_itiY = s_itiY + 340
            Printer.CurrentX = startX
            Printer.CurrentY = s_itiY
            Printer.Print " 　　　 " & Pdata
        End If
    End If
End Sub

Private Sub Sakujo()
    Dim klstx() As String
    Dim i As Integer, j As Integer
'
    For i = TMProw To Ktotal
        For j = 0 To Kdim
            KLST(i, j) = KLST(i + 1, j)
        Next j
    Next i
'
    ReDim klstx(Ktotal, Kdim) As String
'
    For i = 1 To Ktotal
        For j = 0 To Kdim
            klstx(i, j) = KLST(i, j)
        Next j
    Next i
'
    ReDim KLST(Ktotal, Kdim) As String
'
    For i = 1 To Ktotal
        For j = 0 To Kdim
            KLST(i, j) = klstx(i, j)
        Next j
    Next i
End Sub

Private Sub Zoukakdim()
    Dim klstx() As String
    Dim i As Integer, j As Integer
'
    ReDim klstx(Ktotal - 1, Kdim) As String
'
    For i = 1 To Ktotal - 1
        For j = 0 To Kdim
            klstx(i, j) = KLST(i, j)
        Next j
    Next i
'
    ReDim KLST(Ktotal, Kdim) As String
'
    For i = 1 To Ktotal - 1
        For j = 0 To Kdim
            KLST(i, j) = klstx(i, j)
        Next j
    Next i
End Sub

Private Sub checkButtonHyouji(sw As Integer)
    If sw = 1 Then
        cmdCheck.Caption = "部品表ﾁｪｯｸ ＯＫ"
    Else
        cmdCheck.Caption = "<電気部品表ﾁｪｯｸ>"
    End If
End Sub

Private Sub cmdCheck_Click()
    Dim i As Integer
    Dim Tmp1 As String, PFLnameP As String, DRVpartlistP As String
'
    If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
'        TMPdir1 = Xcont0(3)
        TMPdir2 = Xcont0(4)
    Else
'        FLGjob = 1
'        Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
'
        If FLGplst = 0 And FLGplst2 = 0 Then  '*** 部品表２画面とも開いていない ***
            FLGjob = 2
            Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir2, FLGesc ***
        End If
    End If
'
'    Folder_sel.Show 1               '*** "TMPdir1/2 => TMPplst","FLGesc => 0/1" ***
    If FLGesc = 1 Then Exit Sub     '*** ESC ***
'
    If TMPplst <> TMPdir2 Then
        TMPplst = TMPdir2
    End If
'
    For i = 1 To Ktotal
        If Left(KLST(i, 2), 1) = "B" Then
            PFLnameP = KLST(i, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP            '*** .PLT を検索 ***
            Tmp1 = Dir(DRVpartlistP)
'
            If Tmp1 = "" Then
                PFLnameP = KLST(i, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
'
                DRVpartlistP = TMPplst & "\" & PFLnameP        '*** xxxxxxxx.yyy を検索 ***
                Tmp1 = Dir(DRVpartlistP)
'
                If Tmp1 = "" Then
                    i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                    Exit Sub
'
                End If
            End If
        End If
'
    Next i
'
    Call checkButtonHyouji(1)
End Sub

Private Sub cmdHelp_Click()
    Call mnuSetumei_Click
End Sub

Private Sub cmdDown_Click()
    Hdown           '*** 下へ ***
End Sub

Private Sub cmdsyuuryou_Click()
    Dim i As Integer
'
    i = vbYes
    If FLGchange = 1 Then
        Beep
        i = MsgBox("変更内容を放棄しますか？", vbQuestion Or vbYesNo, HeadTitle)
    End If
'
    If i = vbYes Then
        FLGowari = 0
        FLGjob = 0
        FLGlevel = 0
        Unload Me           '*** 終了 ***
'
    Else
        Call Copy_const_work2temp
'
        Const_DirW.Show 1   '*** 書き込みダイアログ ***
'
        If FLGesc = 0 Then  '*** 終了 ***
            Unload Me
'
        End If
    End If
End Sub

Private Sub cmdUp_Click()
    Hup             '*** 上へ ***
End Sub

Private Sub mnuAQuit_Click()
    Dim i As Integer
'
    i = vbYes
    If FLGchange = 1 Then
        Beep
        i = MsgBox("変更内容を放棄しますか？", vbQuestion Or vbYesNo, HeadTitle)
    End If
'
    If i = vbYes Then
        FLGowari = 0
        Unload Me
        End
'
    Else
        Call Copy_const_work2temp
'
        Const_DirW.Show 1   '*** 書き込みダイアログ ***
'
        If FLGesc = 0 Then  '*** 終了 ***
            Unload Me
            End
'
        End If
    End If
End Sub

Private Sub mnuAQuitP_Click()
    Call mnuAQuit_Click
End Sub

Private Sub mnuBack_Click()
    Call cmdsyuuryou_Click
End Sub

Private Sub mnuBackP_Click()
    Call mnuBack_Click
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
    STATUS = "電気 構成表"
    FLGshinki = 0       '*** 新規フラグクリアー ***
'
    If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
        TMPdir1 = Xcont0(3)
        FLGesc = 0
    Else
        Sele_Dir.Show 1         '*** フォルダー選択 => TMPdir1 ***
    End If
'
    If FLGesc = 0 Then
        Const_DirR.Show 1       '*** 読み込みダイアログ ***
'
        Call Copy_const_Temp2work
'
        If FLGesc = 0 Then
            DSPconst0
            DSPconst1
'
            FLGchange = 0       '*** 変更フラグクリアー ***
            FLGowari = 0        '*** 終了フラグりセット ***
'
            Call MENU_shinki
        End If
    End If
End Sub

Private Sub mnufileNRP_Click()
    Call mnufileNR_Click
End Sub

Private Sub mnufileNW_Click()
    STATUS = "電気 構成表"
    Revdate = Format(Date, "yy/mm/dd")
'
    Call Copy_const_work2temp
'
    Const_DirW.Show 1   '*** 書き込みダイアログ ***
'
    If FLGesc = 0 Then
        Call DSPconst0
        Call DSPconst1
'
        FLGshinki = 0   '*** 新規フラグクリアー ***
        FLGchange = 0   '*** 変更フラグクリアー ***
    End If
End Sub

Private Sub mnufileNWP_Click()
    Call mnufileNW_Click
End Sub

Private Sub mnuFilePrn_Click()
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
    txtPrinting.Text = vbCrLf & "  !!! 構成表 " & CURR_file & " を プリントバッファーへ転送中 !!!"
    txtPrinting.Visible = True
    DoEvents                '*** 画面書き直し ***
'
    Call PRNheader0          '*** 構成表ヘッダー印刷 ***
    DoEvents
'
    Call PRNheader1(kijunX, kijunY + habaY)  '*** 構成表項目見出し印刷 ***
    DoEvents
'
    Call PRNkoumoku(kijunX, kijunY + habaY + gyoukan)  '*** 構成表項目印刷 ***
    DoEvents
'
    Printer.EndDoc      '*** プリンター書き込み ***
'
    txtPrinting.Visible = False
    Me.MousePointer = vbDefault             '*** 砂時計解除 ***
    Eeos2_mainMDI.MousePointer = vbDefault
    MSFlexGrid1.MousePointer = flexArrow
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
    i = vbNo
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbQuestion Or vbYesNo, HeadTitle)
    End If
'
    If FLGshinki = 1 Then
        i = vbYes
        Beep
    End If
'
    If i = vbNo Then
        Me.Caption = "  同じ構成表をもう一度読み込んでいます。ちょっと待ってね!"
        Me.MousePointer = vbHourglass
        DoEvents
        Call RDconst_lst(DRVconst, CATno, CATname, Zuban, Person, Orgdate, Revdate, Checkdate, Outdate, _
                KLST(), Ktotal, Kdim, Kouban, Daisuu, Kbikou, KyobiA, KyobiB)
        DoEvents
        Call DSPconst0
        DoEvents
        Call DSPconst1
        DoEvents
        FLGchange = 0    '*** 変更フラグクリアー ***
        Me.Caption = HeadTitle
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub mnufileRDP_Click()
    Call mnufileRD_Click
End Sub

Private Sub mnufileSV_Click()
    Me.Caption = "  構成表を上書き保存しています。ちょっと待ってね！"
    Me.MousePointer = vbHourglass
    DoEvents
'
    Revdate = Format(Date, "yy/mm/dd")    '*** 日付更新 ***
'
    Call WRconst_lst(DRVconst, CATno, CATname, Zuban, Person, Orgdate, Revdate, Checkdate, Outdate, _
                KLST(), Ktotal, Kdim, Kouban, Daisuu, Kbikou, KyobiA, KyobiB)   '*** 構成表 上書き保存 ***
    FLGshinki = 0   '*** 新規フラグクリアー ***
    FLGchange = 0   '*** 変更フラグクリアー ***
    DoEvents
'
    Call DSPconst0
    DoEvents
    Call DSPconst1
    DoEvents
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
End Sub

Private Sub mnufileSVP_Click()
    Call mnufileSV_Click
End Sub

Private Sub mnuFileWR_Click()   '*** CSV形式で出力 ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
    Dim CSVfile As String
    Dim PDa As String
    Dim Pdata As String
'
    Me.Caption = "  構成表を「CSV形式」で " & "I" & Trim(Zuban) & ".CSV として保存しています。ちょっと待ってね！"
    Me.MousePointer = vbHourglass
    DoEvents
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    CSVfile = TMPdir1 & "\I" & Trim(Zuban) & ".CSV"
    Open CSVfile For Output As #FILE_number
        PDa = "《 電 気 構 成 表 》"
        If Val(Left(Revdate, 2)) > 85 Then
            PDa = PDa & ",19" & Revdate
        Else
            PDa = PDa & ",20" & Revdate
        End If
        Print #FILE_number, PDa & ",型 式： " & CATno & ",名 称： " & CATname & ",*,*"
'
        PDa = "図面番号： " & Zuban & ",工 番： " & Kouban & ",台数： " & Daisuu & ",技術担当者： " & Person & ",*,*"
        Print #FILE_number, PDa
    
        PDa = "名 称" & ",回 路 図" & ",電気部品表" & ",パターン番号" & ",数" & ",備 考"
        Print #FILE_number, PDa
    
        For i = 1 To Ktotal
            If Trim(KLST(i, 0)) = "〃" Then
                PDa = "     " & Trim(KLST(i, 0))
            ElseIf Trim(KLST(i, 0)) = "*" Or Trim(KLST(i, 0)) = "-" Then
                PDa = "   --------"
            Else
                PDa = KLST(i, 0)
            End If
'                       *** 回路図 ***
            If Trim(KLST(i, 1)) = "〃" Then
                PDa = PDa & "," & Trim(KLST(i, 1))
            ElseIf Trim(KLST(i, 1)) = "*" Or Trim(KLST(i, 1)) = "-" Then
                PDa = PDa & ",--------"
            Else
                PDa = PDa & "," & KLST(i, 1)
            End If
'                       *** 電気部品表 ***
            If Trim(KLST(i, 2)) = "〃" Then
                PDa = PDa & "," & Trim(KLST(i, 2))
            ElseIf Trim(KLST(i, 2)) = "*" Or Trim(KLST(i, 2)) = "-" Then
                PDa = PDa & ",--------"
            Else
                PDa = PDa & "," & KLST(i, 2)
            End If
'                       *** パターン番号 ***
            If Trim(KLST(i, 3)) = "〃" Then
                PDa = PDa & "," & Trim(KLST(i, 3))
            ElseIf Trim(KLST(i, 3)) = "*" Or Trim(KLST(i, 3)) = "-" Then
                PDa = PDa & ",--------"
            Else
                PDa = PDa & "," & KLST(i, 3)
            End If
'                       *** 数 ***
                PDa = PDa & "," & Trim(KLST(i, 4))
'                       *** 備考 ***
            If Trim(KLST(i, 5)) = "*" Then
                PDa = PDa & ",*"
            Else
                PDa = PDa & "," & KLST(i, 5)
            End If
            Print #FILE_number, PDa
            DoEvents
        Next i
'
        If Kbikou <> "" Then
            j = InStr(Kbikou, " _")
            If j = 0 Then
                Print #FILE_number, "備考： " & Kbikou & ",*,*,*,*,*"
            Else
                Print #FILE_number, "備考： " & Left(Kbikou, j - 1) & ",*,*,*,*,*"
                Pdata = Mid(Kbikou, j + 2)
                j = InStr(Pdata, " _")
'
                Do While j > 0
                    Print #FILE_number, Left(Pdata, j - 1) & ",*,*,*,*,*"
                    Pdata = Mid(Pdata, j + 2)
                    j = InStr(Pdata, " _")
                Loop
'
                Print #FILE_number, Pdata & ",*,*,*,*,*"
            End If
        End If
    Close #FILE_number
    DoEvents
'
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuPurFileWR_Click()    '***資材用特殊ﾌｧｲﾙ出力
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
    Dim CSVfile As String
    Dim PDa As String, PDb As String, PDc As String, PDd As String, PDe As String, PDf As String
    Dim Pdata As String
'
    Me.Caption = "  構成表を「資材部特殊形式」で kosei.CSV として保存しています。ちょっと待ってね！"
    Me.MousePointer = vbHourglass
    DoEvents
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    CSVfile = TMPdir1 & "\kosei.CSV"
    Open CSVfile For Output As #FILE_number
        For i = 1 To Ktotal
            If Trim(KLST(i, 0)) = "〃" Then
                PDa = "     " & Trim(KLST(i, 0))
            ElseIf Trim(KLST(i, 0)) = "*" Or Trim(KLST(i, 0)) = "-" Then
                PDa = "   --------"
            Else
                PDa = KLST(i, 0)
            End If
'                       *** 回路図 ***
            If Trim(KLST(i, 1)) = "〃" Then
                PDb = Trim(KLST(i, 1))
            ElseIf Trim(KLST(i, 1)) = "*" Or Trim(KLST(i, 1)) = "-" Then
                PDb = "--------"
            Else
                PDb = KLST(i, 1)
            End If
'                       *** 電気部品表 ***
            If Trim(KLST(i, 2)) = "〃" Then
                PDc = Trim(KLST(i, 2))
            ElseIf Trim(KLST(i, 2)) = "*" Or Trim(KLST(i, 2)) = "-" Then
                PDc = "--------"
            Else
                PDc = KLST(i, 2)
            End If
'                       *** パターン番号 ***
            If Trim(KLST(i, 3)) = "〃" Then
                PDd = Trim(KLST(i, 3))
            ElseIf Trim(KLST(i, 3)) = "*" Or Trim(KLST(i, 3)) = "-" Then
                PDd = "--------"
            Else
                PDd = KLST(i, 3)
            End If
'                       *** 数 ***
                PDe = Trim(KLST(i, 4))
'                       *** 備考 ***
            If Trim(KLST(i, 5)) = "*" Then
                PDf = "*"
            Else
                PDf = KLST(i, 5)
            End If
'
            Write #FILE_number, PDa, PDb, PDc, PDd, PDe, PDf,
            Print #FILE_number,
            DoEvents
        Next i
    Close #FILE_number
    DoEvents
'
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuHinsyu_Click()
    Call mnuCodeHinsyuMaintenance
End Sub

Private Sub mnuHinsyuP_Click()
    Call mnuHinsyu_Click
End Sub

Private Sub mnuJumpC_Click()
    Dim Tyuusin As Integer
'
    Tyuusin = Ktotal \ 2
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

Private Sub mnuJumpE_Click()
    If Ktotal <= G_row Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Ktotal - G_row + 1
    End If
End Sub

Private Sub mnuJumpEP_Click()
    Call mnuJumpE_Click
End Sub

Private Sub mnuJumpT_Click()
    If Ktotal <= G_row Then
        Beep
    End If
'
    MSFlexGrid1.TopRow = 1
End Sub

Private Sub mnuJumpTP_Click()
    Call mnuJumpT_Click
End Sub

Private Sub mnuKaihan_Click()
    Call mnuKaihanRireki
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

Private Sub mnuReform_Click()
    Me.Width = 360 + (OrgWidth - 720) * HyoujiBairitu + 360
                                            '*** これで「Form_Resize」割り込みが発生する。 ***
    Me.Height = 240 + (OrgHeight - 480) * HyoujiBairitu + 240
'
    Call setFormArea    '*** フォームの表示位置の設定 ***
End Sub

Private Sub mnuSettei_Click()
    Call mnuKankyouSettei
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

Private Sub MSFlexGrid1_Click()
    txtkbikou.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
    Dim bunsyou As String, i As Integer
'
    TMProw = MSFlexGrid1.Row    '*** マウスボタンクリック位置記憶 ***
    TMPcol = MSFlexGrid1.Col
'
    If TMPcol = 0 Then
        If TMProw <= Ktotal Then
            Beep
            bunsyou = "<削除！ " & "行番号" & str(TMProw) & ">  この行を削除するなら「はい(Y)」ボタン、" & vbCrLf & " ここに行を挿入するなら「いいえ(N)」ボタンを押してください。"
            i = MsgBox(bunsyou, vbQuestion Or vbYesNoCancel, HeadTitle)
        Else
            Beep
            bunsyou = "最終行番号" & str(Ktotal) & " の次に追加します。"
            i = MsgBox(bunsyou, vbInformation Or vbOKCancel, HeadTitle)
            If i = vbOK Then
                i = vbNo
            End If
        End If
'
        If i = vbYes Then
            Ktotal = Ktotal - 1
            Call Sakujo
            FLGchange = 1   '*** 変更フラグセット ***
            Call DSPconst1
'
        ElseIf i = vbNo Then
            Ktotal = Ktotal + 1
'
            If TMProw <= Ktotal Then
                Call Zoukakdim      '*** ﾃﾞｨﾒﾝｼﾞｮﾝ増加 ***
                Call Kurisage       '*** 挿入処理 ***
            Else
                Call Zoukakdim      '*** ﾃﾞｨﾒﾝｼﾞｮﾝ増加 ***
            End If
'
            FLGchange = 1   '*** 変更フラグセット ***
            Call DSPconst1
'
            FLGtuika = 1    '*** 追加フラグセット ***
            KtotalT = Ktotal
            KdimT = Kdim
            ReDim KLSTT(KtotalT, KdimT) As String
            KLSTT() = KLST()
'
            Const_main_c.Show 1
'
            KLST() = KLSTT()
        Else
            FLGtuika = 0    '*** 追加フラグクリアー ***
        End If
    Else
        If TMProw > Ktotal Then
            Beep
'
        Else
            FLGtuika = 0        '*** 追加フラグクリアー ***
            KtotalT = Ktotal
            KdimT = Kdim
            ReDim KLSTT(KtotalT, KdimT) As String
            KLSTT() = KLST()
'
            Const_main_c.Show 1
'
            KLST() = KLSTT()
'
            If TMPcol = 3 And FLGchange = 1 Then
                Call checkButtonHyouji(0)
            End If
        End If
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 1          '*** 部品構成表フラグ ***
    Call mnuSousaSetumei
End Sub

Private Sub mnuSinki_Click()
    Dim i As Integer
'
    i = vbNo
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbQuestion Or vbYesNo, HeadTitle)
    End If
'
    If i = vbNo Then
        CURR_file = "shinki01.cod"
        Call DSPconst0N
        Call DSPconst0
        Call DSPconst1N
        Call DSPconst1
        FLGchange = 0   '*** 変更フラグクリアー ***
        FLGshinki = 1   '*** 新規フラグセット ***
'
        Call MENU_shinki
    End If
End Sub

Private Sub mnuVersion_Click()
    Call mnuVersionGamen
End Sub

Private Sub MENU_settei()       '*** メニュー状態設定 ***
'
'   If FLGconst = 1 Then       '*** 構成表画面存在 ***
'       Me.mnuKousei.Checked = True
'       Me.mnuKouseiP.Checked = True
'   Else
'       Me.mnuKousei.Checked = False
'       Me.mnuKouseiP.Checked = False
'   End If
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

Private Sub setFormArea()   '*** フォームの表示位置の設定 ***
    Top = 0
    Left = 0
End Sub

Private Sub txtbangou_Click()
    txtbangou.MousePointer = vbIbeam
End Sub

Private Sub txtbangou_LostFocus()
    Tmpdata = Trim(txtbangou.Text)
'
    If Tmpdata <> Zuban Then
        Zuban = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txtbangou.MousePointer = vbArrow
End Sub

Private Sub txtdaisuu_Click()
    txtdaisuu.MousePointer = vbIbeam
End Sub

Private Sub txtdaisuu_LostFocus()
    Tmpdata = Trim(txtdaisuu.Text)
'
    If Tmpdata <> Daisuu Then
        Daisuu = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txtdaisuu.MousePointer = vbArrow
End Sub

Private Sub txtkbikou_Click()
    txtkbikou.MousePointer = vbIbeam
End Sub

Private Sub txtkbikou_LostFocus()
    Tmpdata = Trim(txtkbikou.Text)
'
    If Tmpdata <> Kbikou Then
        Kbikou = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txtkbikou.MousePointer = vbArrow
End Sub

Private Sub txtKeisiki_Click()
    txtKeisiki.MousePointer = vbIbeam
End Sub

Private Sub txtKeisiki_LostFocus()
    Tmpdata = Trim(txtKeisiki.Text)
'
    If Tmpdata <> CATno Then
        CATno = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txtKeisiki.MousePointer = vbArrow
End Sub

Private Sub txtkouban_Click()
    txtkouban.MousePointer = vbIbeam
End Sub

Private Sub txtkouban_LostFocus()
    Tmpdata = Trim(txtkouban.Text)
'
    If Tmpdata <> Kouban Then
        Kouban = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txtkouban.MousePointer = vbArrow
End Sub

Private Sub txtmeisyou_Click()
    txtmeisyou.MousePointer = vbIbeam
End Sub

Private Sub txtmeisyou_LostFocus()
    Tmpdata = Trim(txtmeisyou.Text)
'
    If Tmpdata <> CATname Then
        CATname = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txtmeisyou.MousePointer = vbArrow
End Sub

Private Sub txttantou_Click()
    txttantou.MousePointer = vbIbeam
End Sub

Private Sub txttantou_LostFocus()
    Tmpdata = Trim(txttantou.Text)
'
    If Tmpdata <> Person Then
        Person = Tmpdata
        FLGchange = 1    '*** 変更フラグセット ***
    End If
'
    txttantou.MousePointer = vbArrow
End Sub

Private Sub DSPgamenBuhin()
    txtKeisiki.Left = 360 + (1690 - 360) * HyoujiBairitu
    txtKeisiki.Top = 360 + FLGoffsetY
    txtKeisiki.FontSize = 10 * HyoujiBairitu
    txtKeisiki.Width = 1455 * HyoujiBairitu
    txtKeisiki.Height = 285 * HyoujiBairitu
'
    lblkeisiki.Left = 360
    lblkeisiki.Top = 360 + FLGoffsetY
    lblkeisiki.FontSize = 10 * HyoujiBairitu
    lblkeisiki.Width = 1335 * HyoujiBairitu
    lblkeisiki.Height = txtKeisiki.Height
'
    txtmeisyou.Left = 360 + (1690 - 360) * HyoujiBairitu
    txtmeisyou.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    txtmeisyou.FontSize = 10 * HyoujiBairitu
    txtmeisyou.Width = 3375 * HyoujiBairitu
    txtmeisyou.Height = 285 * HyoujiBairitu
'
    lblmeisyou.Left = 360
    lblmeisyou.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    lblmeisyou.FontSize = 10 * HyoujiBairitu
    lblmeisyou.Width = 1335 * HyoujiBairitu
    lblmeisyou.Height = txtmeisyou.Height
'
    txtbangou.Left = 360 + (1690 - 360) * HyoujiBairitu
    txtbangou.Top = 360 + (1080 - 360) * HyoujiBairitu + FLGoffsetY
    txtbangou.FontSize = 10 * HyoujiBairitu
    txtbangou.Width = 1440 * HyoujiBairitu
    txtbangou.Height = 285 * HyoujiBairitu
'
    lblbangou.Left = 360
    lblbangou.Top = 360 + (1080 - 360) * HyoujiBairitu + FLGoffsetY
    lblbangou.FontSize = 10 * HyoujiBairitu
    lblbangou.Width = 1335 * HyoujiBairitu
    lblbangou.Height = txtbangou.Height
'
    Call checkButtonHyouji(0)
'
    cmdCheck.Left = 360 + (3360 - 360) * HyoujiBairitu
    cmdCheck.Top = 360 + (1170 - 360) * HyoujiBairitu + FLGoffsetY
    cmdCheck.FontSize = 9 * HyoujiBairitu
    cmdCheck.Width = 1695 * HyoujiBairitu
    cmdCheck.Height = 240 * HyoujiBairitu
'
    txtkouban.Left = 360 + (6850 - 360) * HyoujiBairitu
    txtkouban.Top = 360 + FLGoffsetY
    txtkouban.FontSize = 10 * HyoujiBairitu
    txtkouban.Width = 2655 * HyoujiBairitu
    txtkouban.Height = 285 * HyoujiBairitu
'
    lblkouban.Left = 360 + (5280 - 360) * HyoujiBairitu
    lblkouban.Top = 360 + FLGoffsetY
    lblkouban.FontSize = 10 * HyoujiBairitu
    lblkouban.Width = 1575 * HyoujiBairitu
    lblkouban.Height = txtkouban.Height
'
    txtdaisuu.Left = 360 + (6850 - 360) * HyoujiBairitu
    txtdaisuu.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    txtdaisuu.FontSize = 10 * HyoujiBairitu
    txtdaisuu.Width = 735 * HyoujiBairitu
    txtdaisuu.Height = 285 * HyoujiBairitu
'
    lbldaisuu.Left = 360 + (5280 - 360) * HyoujiBairitu
    lbldaisuu.Top = 360 + (720 - 360) * HyoujiBairitu + FLGoffsetY
    lbldaisuu.FontSize = 10 * HyoujiBairitu
    lbldaisuu.Width = 1575 * HyoujiBairitu
    lbldaisuu.Height = txtdaisuu.Height
'
    txttantou.Left = 360 + (6850 - 360) * HyoujiBairitu
    txttantou.Top = 360 + (1080 - 360) * HyoujiBairitu + FLGoffsetY
    txttantou.FontSize = 10 * HyoujiBairitu
    txttantou.Width = 1575 * HyoujiBairitu
    txttantou.Height = 285 * HyoujiBairitu
'
    lbltantou.Left = 360 + (5280 - 360) * HyoujiBairitu
    lbltantou.Top = 360 + (1080 - 360) * HyoujiBairitu + FLGoffsetY
    lbltantou.FontSize = 10 * HyoujiBairitu
    lbltantou.Width = 1575 * HyoujiBairitu
    lbltantou.Height = txttantou.Height
'
    cmdHelp.Left = 360 + (9120 - 360) * HyoujiBairitu
    cmdHelp.Top = 360 + (1140 - 360) * HyoujiBairitu + FLGoffsetY
    cmdHelp.FontSize = 10 * HyoujiBairitu
    cmdHelp.Width = 255 * HyoujiBairitu
    cmdHelp.Height = 255 * HyoujiBairitu
'
    MSFlexGrid1.Left = 360
    MSFlexGrid1.Top = 360 + (1440 - 360) * HyoujiBairitu + FLGoffsetY
'
    With MSFlexGrid1          '*** グリッドの幅の設定 ***
        .Width = 9375 * HyoujiBairitu + FLGoffsetX
        .RowHeightMin = 245 * HyoujiBairitu
        .Height = .RowHeightMin * 21  '5125 * HyoujiBairitu  '5175?
        .Font.Size = 10 * HyoujiBairitu
        .Cols = 7
        .ColWidth(0) = 300 * HyoujiBairitu
        .ColWidth(1) = 2184 * HyoujiBairitu + FLGoffsetX * 0.1
        .ColWidth(2) = 1488 * HyoujiBairitu + FLGoffsetX * 0.1
        .ColWidth(3) = 1488 * HyoujiBairitu + FLGoffsetX * 0.1
        .ColWidth(4) = 1512 * HyoujiBairitu + FLGoffsetX * 0.1
        .ColWidth(5) = 360 * HyoujiBairitu + (240 * HyoujiBairitu - 240)
        .ColWidth(6) = 1730 * HyoujiBairitu + FLGoffsetX * 0.6
    End With
    G_row = 20      '*** グリッドのデータ表示行数 ***
'
    txtkbikou.Left = 360 + (1330 - 360) * HyoujiBairitu
    txtkbikou.Top = 360 + (6600 - 360) * HyoujiBairitu + FLGoffsetY
    txtkbikou.FontSize = 10 * HyoujiBairitu
    txtkbikou.Width = 8415 * HyoujiBairitu
    txtkbikou.Height = 495 * HyoujiBairitu
'
    lblbikou.Left = 360
    lblbikou.Top = 360 + (6600 - 360) * HyoujiBairitu + FLGoffsetY
    lblbikou.FontSize = 10 * HyoujiBairitu
    lblbikou.Width = 975 * HyoujiBairitu
    lblbikou.Height = txtkbikou.Height
'
    lbl_curr_file.Left = 360 + (2880 - 360) * HyoujiBairitu
    lbl_curr_file.Top = 360 + (7320 - 360) * HyoujiBairitu + FLGoffsetY
    lbl_curr_file.FontSize = 10 * HyoujiBairitu
    lbl_curr_file.Width = 3255 * HyoujiBairitu
    lbl_curr_file.Height = 255 * HyoujiBairitu
'
    cmdUp.Left = 360 + (6720 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (7320 - 360) * HyoujiBairitu + FLGoffsetY
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Width = 735 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdDown.Left = 360 + (7680 - 360) * HyoujiBairitu
    cmdDown.Top = 360 + (7320 - 360) * HyoujiBairitu + FLGoffsetY
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Width = 735 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdsyuuryou.Left = 360 + (8640 - 360) * HyoujiBairitu
    cmdsyuuryou.Top = 360 + (7320 - 360) * HyoujiBairitu + FLGoffsetY
    cmdsyuuryou.FontSize = 10 * HyoujiBairitu
    cmdsyuuryou.Width = 975 * HyoujiBairitu
    cmdsyuuryou.Height = 495 * HyoujiBairitu
'
    txtPrinting.Width = 6600 * HyoujiBairitu
    txtPrinting.Height = 720 * HyoujiBairitu
    txtPrinting.FontSize = 10 * HyoujiBairitu
    txtPrinting.Left = (Me.Width - txtPrinting.Width) / 2
    txtPrinting.Top = (Me.Height - txtPrinting.Height) / 3
    txtPrinting.Visible = False
    txtPrinting.BackColor = &HC000&
    txtPrinting.ForeColor = &HFFFFFF
End Sub


