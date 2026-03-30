VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Plst_Suuryo 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "部品数量表印刷"
   ClientHeight    =   6255
   ClientLeft      =   2475
   ClientTop       =   1200
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
   Icon            =   "Plst_Suuryo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6255
   ScaleWidth      =   7215
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   4320
   End
   Begin VB.Frame frmTanka 
      BackColor       =   &H00008000&
      Caption         =   "備考欄 選択"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4920
      TabIndex        =   32
      Top             =   3960
      Width           =   1815
      Begin VB.OptionButton optTainetu 
         BackColor       =   &H00008000&
         Caption         =   "耐熱/ﾒｯｷ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optTanka 
         BackColor       =   &H00008000&
         Caption         =   "部品単価"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optBikou 
         BackColor       =   &H00008000&
         Caption         =   "  備 考"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame frmSoutou 
      BackColor       =   &H00008000&
      Caption         =   "「相当」記入位置"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2760
      TabIndex        =   28
      Top             =   3960
      Width           =   1815
      Begin VB.OptionButton optBikourann 
         BackColor       =   &H00008000&
         Caption         =   "備考欄先頭"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optBuhinmei 
         BackColor       =   &H00008000&
         Caption         =   "部品名末尾"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label lblTyu1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00008000&
         Caption         =   "(ﾌｧｲﾙ出力時有効)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmSort 
      BackColor       =   &H00008000&
      Caption         =   "ｿｰﾃｨﾝｸﾞ条件"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4920
      TabIndex        =   24
      Top             =   2640
      Width           =   1815
      Begin VB.OptionButton optMaker 
         BackColor       =   &H00008000&
         Caption         =   "ﾒｰｶｰ出現順"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optName 
         BackColor       =   &H00008000&
         Caption         =   "部品名称順"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optCode 
         BackColor       =   &H00008000&
         Caption         =   "ｺｰﾄﾞ出現順"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame frmSougou 
      BackColor       =   &H00008000&
      Caption         =   "全部品表印刷"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4920
      TabIndex        =   21
      Top             =   1560
      Width           =   1815
      Begin VB.OptionButton optSougou 
         BackColor       =   &H00008000&
         Caption         =   "総合集計"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optKobetu 
         BackColor       =   &H00008000&
         Caption         =   "個別集計"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "ﾌｧｲﾙ出力(&F)"
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtbaisuu 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   17
      Text            =   "1"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtdaisuu 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   12
      Text            =   "1000"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txttantou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   8
      Text            =   "名前を記入だ"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtmeisyou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   6
      Text            =   "TV SOUND MULTI MODULATOR"
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtkeisiki 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   4
      Text            =   "1000A-001"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame frmkinyuu 
      BackColor       =   &H00008000&
      Caption         =   "項目記入名称"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   1815
      Begin VB.OptionButton opthyou 
         BackColor       =   &H00008000&
         Caption         =   "構成表による"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optdirect 
         BackColor       =   &H00008000&
         Caption         =   "直接入力"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox txtkouban 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   10
      Text            =   "A11-1234-12/56,A12-1235-13/15"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "印刷(&P)"
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   4920
      TabIndex        =   20
      Top             =   5400
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   10
   End
   Begin VB.Label lblbaisuu 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "倍数"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lbldate 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "日付 ： 1997/09/19"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label lblkomei 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "小名称 ：ABCDEFGHIJKLMNOPQRSTUVWXY"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label lbldaisuu 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "台数"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbltantou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "担当者"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblmeisyou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "名 称"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblkeisiki 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "型式"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblkouban 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "工番"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblnamae 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "印刷 ﾌｧｲﾙ名 ：A1234-00.2AB"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   2520
      Width           =   3495
   End
End
Attribute VB_Name = "Plst_Suuryo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 部品数量表印刷 ***
'**********************
'
Option Explicit
'
Dim HeadTitle As String
Dim FLGkouseihyou As Integer, FLGgyou As Integer, FLGpage As Integer
Dim FLGtuzuki As Integer, FLGtuzukip As Integer
Dim FLGend As Integer, FLGheader As Integer
Dim FLGitemp As Integer
Dim kijunX As Integer, kijunY As Integer
Dim habaX As Integer, habaY As Integer
Dim haba1X As Integer, haba2X As Integer, haba3X As Integer
Dim haba4X As Integer, haba5X As Integer, haba6X As Integer
Dim haba7X As Integer, gyoukan As Integer
Dim moji_zureX As Integer, moji_zureY As Integer
Dim NowpointX As Integer, NowpointY As Integer
Dim Tmpdata As String, Baisuu As String
'
Dim PFLnameP As String
Dim DRVpartlistP As String
Dim PlistnameP As String
Dim PlistdateP As String
Dim RemarksP As String
Dim PtotalP As Integer
Dim PdimP As Integer
Dim PLSTP() As String
'
Dim Pdata_total() As String     '*** Pdata_total(Ckmax, Ckdim) ***
                                '*** 0: 項目、1: 部品ｺｰﾄﾞ、2: 備考、3: メーカ指定 ***
                                '*** 4: 特記事項、5: 個数 ***
Dim CKmax As Integer, Ckdim As Integer, CKsuu As Integer
Dim FNAME_suuryou As String
Dim TempFname As String
Dim Gyoumax As Integer
'
Dim FLGsort As Integer
Dim FLGsoutouhin As Integer
Dim Pdata_lines() As String     '*** Pdata_lines(CKsuu,line_dim) ***
                '*** 0:部品ｺｰﾄﾞ、1:規格・定格、2:個数、3:ﾒｰｶｰ、4:備考/単価 ***
Dim ipT As Integer
Dim Goukei As Double

Private Sub Form_Initialize()
    HeadTitle = "部品表 <数量表印刷>"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (7305 - 960) * HyoujiBairitu + 480
    Height = 480 + (6750 - 960) * HyoujiBairitu + 480
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
    Timer1.Interval = 10
    Timer1.Enabled = False
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    Me.Caption = HeadTitle
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
    Call S_Hyouji            '*** 画面初期化＆表示 ***
End Sub

Private Sub PRNhamidashi(p As Integer)
    Printer.Print "はみ出し部品  " & Pdata_total(p, 0) & "  " & Pdata_total(p, 1) & "  " & Pdata_total(p, 2)
End Sub

Private Sub PRNIndex(ip As Integer)
'                   *** 項目名印刷 ***
    Dim itiji As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print " " & Aitem0(ip, 1)
'
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
End Sub

Private Sub PRNfooter()
'                   *** フッター印刷 ***
    Dim Ptempd As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.CurrentX = NowpointX + 284
    Printer.CurrentY = NowpointY + moji_zureY
    SETfont_size 9, 1
    Printer.Print "M5904-04";
'
    Ptempd = FormatDateTime(Date, vbLongDate) & " 印刷"
    SETfont_size 10.8, 0
    Printer.CurrentX = NowpointX + 5250 - Printer.TextWidth(Ptempd) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd;
'
    Printer.CurrentX = NowpointX + 9500
    Printer.CurrentY = NowpointY + moji_zureY
    If FLGend = 0 Then
        Printer.Print "つづく"
    Else
        If optTanka.Value = True Then
            Printer.CurrentX = NowpointX + 7250
            Call SET_Yen0_Format(Goukei, Ptempd, 12)
            Printer.Print "おわり    合計："; Ptempd; "*"
        Else
            Printer.Print "おわり"
        End If
    End If
End Sub

Private Sub PRNkoumoku(n As Integer)
'                   *** 項目印刷 ***
    Dim Pdata As String
    Dim Pdata1 As String
'
    If FLGgyou >= Gyoumax Then
        FLGend = 0
        PRNfooter
        Printer.NewPage   '*** 改ページ ***
'
        PRNheader0
        PRNheader1
        FLGheader = 1   '*** ヘッダー印刷済みにｾｯﾄ ***
    End If
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "    " & Pdata_lines(n, 0)
'
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Pdata_lines(n, 1)
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
        SET_migiyose Pdata_lines(n, 2), Pdata, 6 '*** 6文字にする
    Printer.Print Pdata
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
        SET_migiyose str(Val(Pdata_lines(n, 2)) * Val(DaisuuT)), Pdata, 6  '*** 6文字にする
    Printer.Print Pdata
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Pdata_lines(n, 3)
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
'
    If optBikou.Value = True Then
        Pdata = Pdata_lines(n, 4)
        If Pdata = "*" Then
            Pdata = " "
        End If
        Printer.Print Pdata
'
    ElseIf optTanka.Value = True Then
        If 0 < Val(Pdata_lines(n, 4)) Then
            Call SET_Yen1_Format(Val(Pdata_lines(n, 4)), Pdata, 9)
            Call SET_Yen0_Format(Val(Pdata_lines(n, 4)) * Val(Pdata_lines(n, 2)) * Val(DaisuuT), Pdata1, 10)
            Printer.Print Pdata; Pdata1
            Goukei = Goukei + Val(Pdata_lines(n, 4)) * Val(Pdata_lines(n, 2)) * Val(DaisuuT)
        End If
'
    ElseIf optTainetu.Value = True Then
        Printer.CurrentX = NowpointX + moji_zureX - 40
'
        Pdata = Pdata_lines(n, 4)
        Printer.Print Pdata
    End If
End Sub

Private Sub PRNheader1()
'                   *** 部品表項目見出し印刷 ***
    Dim Ptempd As String
'                          567twip=10mm,1440twip=1inch,<16114>
    FLGgyou = 1
    haba1X = 1786
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Ptempd = "品名／ｺｰﾄﾞ番号"
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + (haba1X - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd
'
    haba2X = 3760
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Ptempd = "規 格・定 格"
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + (haba2X - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd
'
    haba3X = 900
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Ptempd = "個 数"
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + (haba3X - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd
'
    haba4X = 900
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Ptempd = "総 数"
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + (haba4X - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd
'
    haba5X = 900
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Ptempd = "ﾒｰｶｰ"
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + (haba5X - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd
'
    haba6X = 2254
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    If optTainetu.Value = True Then
        Ptempd = "MSL/耐熱/ ﾒｯｷ /RoHS"
    ElseIf optTanka.Value = True Then
        Ptempd = " 平均単価    ｘ総数"
    Else
        Ptempd = "備 考"
    End If
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + (haba6X - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Ptempd
End Sub

Private Sub PRNheader0()
'                   *** 部品数量表ヘッダー印刷 ***
    Dim i As Integer
    Dim Ptempd As String
'                   567twip=10mm,1440twip=1inch
    kijunX = 720
    kijunY = 1134
    habaX = 10500
    habaY = 340     'gyoukan
    gyoukan = 340
    moji_zureX = 113
    moji_zureY = 67
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    Printer.Orientation = vbPRORPortrait    '*** ﾎﾟｰﾄﾚｲﾄ ***
'
    Ptempd = "《　電  気  部  品  数  量  表　》"
    Printer.CurrentX = kijunX
    Printer.CurrentY = 750
    SETfont_size 17, 1    '*** フォント,サイズ設定 ***
    Printer.CurrentX = kijunX + (habaX - Printer.TextWidth(Ptempd)) / 2
    Printer.CurrentY = 750
    Printer.Print Ptempd
'
    Printer.CurrentX = kijunX + 9300
    Printer.CurrentY = 870
    SETfont_size 10.8, 1    '*** フォント,サイズ設定 ***
    Printer.Print FLGpage & " ページ"
    FLGpage = FLGpage + 1
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
'
    haba1X = 2320
    NowpointX = kijunX
    NowpointY = kijunY
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "型式： " & CATnoT
'
    haba2X = 4530
    NowpointX = kijunX + haba1X
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    If FLGall = 1 And optSougou.Value = True Then
        Printer.Print "名称： " & CATnameT
    Else
        Printer.Print "小名称：" & TempFname & "; " & PlistnameP
    End If
'
    haba3X = 2800
    NowpointX = kijunX + haba1X + haba2X
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "工番： " & KoubanT
'
    haba4X = 850
    NowpointX = kijunX + haba1X + haba2X + haba3X
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
        SET_migiyose str(Val(DaisuuT)), Ptempd, 4 '*** 4文字にする ***
    Printer.Print Ptempd & "台"
End Sub

Private Sub Klst_yomu()
    DRVconstT = TMPdir1 & "\constlst.cod"
    Call RDconst_lst(DRVconstT, CATnoT, CATnameT, ZubanT, PersonT, OrgdateT, RevdateT, CheckdateT, OutdateT, _
                KLSTT(), KtotalT, KdimT, KoubanT, DaisuuT, KbikouT, KyobiAT, KyobiBT)       '*** 構成表読み込み ***
'
    txtkeisiki.Text = CATnoT
    txtmeisyou.Text = CATnameT
    txttantou.Text = PersonT
    txtkouban.Text = KoubanT
    txtdaisuu.Text = DaisuuT
End Sub

Private Sub Plst_yomu()
'                   *** 部品表を読む ***
    Dim i As Integer
    Dim tmpdata0 As String, tmpdata1 As String, tmpdata2 As String
'
    Call RDpartlist(DRVpartlistP, PlistnameP, PlistdateP, RemarksP, PLSTP(), PtotalP, PdimP) '*** 部品表読み込み ***
'
    If FLGall = 0 Then
        lblnamae.Caption = "読み込みﾌｧｲﾙ名： " & PFLnameP
        lblkomei.Caption = "小名称 ： " & PlistnameP
        lblDate.Caption = "日付 ： " & PlistdateP
        txtbaisuu.Text = Baisuu
'
        Call remove_PLT(PFLnameP, TempFname)
    End If
'
    For i = 1 To PtotalP
        tmpdata1 = PLSTP(i, 0)
'
        Call GETsymbol(tmpdata0, tmpdata1, tmpdata2)
'
        PLSTP(i, PdimP + 1) = tmpdata1
        PLSTP(i, PdimP + 2) = tmpdata2
        PLSTP(i, PdimP + 3) = tmpdata0
    Next i
End Sub

Private Sub PRNmitouroku()
'                   *** 未登録部品印刷 ***
    Dim Pdata As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    Printer.Print "   未登録部品"
'
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Mid(Pdata_total(pp, 1), 2)
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
        SET_migiyose Pdata_total(pp, 5), Pdata, 6 '*** 6文字にする
    Printer.Print Pdata
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
        SET_migiyose str(Val(Pdata_total(pp, 5)) * Val(DaisuuT)), Pdata, 6 '*** 6文字にする
    Printer.Print Pdata
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
'
    If optBikou.Value = True Then
        Pdata = Pdata_total(pp, 2)
        If Pdata = "*" Then
            Pdata = " "
        End If
        Printer.Print Pdata
    End If
'
    If optTainetu.Value = True Then
        Pdata = " ??? "
        Printer.Print Pdata
    End If
End Sub
        
Private Sub Print_sort(Tlines As Integer)            '*** 条件によるソーティング ***
    Dim i As Integer, j As Integer, k As Integer
    Dim Ptemp(4) As String
'
    Select Case FLGsort
    Case 0  '*** ｺｰﾄﾞ出現順 ***
        '--- 変更なし ---
'
    Case 1  '*** 部品名称順 ***
        For i = 1 To Tlines - 1
            For j = Tlines - i To Tlines - 1
                If Pdata_lines(j + 1, 1) < Pdata_lines(j, 1) Then
                    For k = 0 To 4
                        Ptemp(k) = Pdata_lines(j, k)                '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 4
                        Pdata_lines(j, k) = Pdata_lines(j + 1, k)   '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 4
                        Pdata_lines(j + 1, k) = Ptemp(k)            '*** 入れ替え ***
                    Next k
                End If
            Next j
        Next i
'
    Case 2  '*** ﾒｰｶｰ出現順 ***
        For i = 1 To Tlines - 1
            For j = Tlines - i To Tlines - 1
                If Pdata_lines(j + 1, 3) < Pdata_lines(j, 3) Then
                    For k = 0 To 4
                        Ptemp(k) = Pdata_lines(j, k)                '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 4
                        Pdata_lines(j, k) = Pdata_lines(j + 1, k)   '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 4
                        Pdata_lines(j + 1, k) = Ptemp(k)            '*** 入れ替え ***
                    Next k
                End If
            Next j
        Next i
'
    End Select
End Sub

Private Sub Insatsu_MAIN()  '*** 印刷ルーチン ***
    Dim i As Integer, j As Integer, np As Integer, q As Integer
    Dim ip As Integer
    Dim ips As Integer
    Dim jp As Integer
    Dim jps As Integer
    Dim kp As Integer
    Dim MKsitei As String
    Dim Tmp1 As String, Tmp0 As String
    Dim FLGitti As Integer, FLGmojime As Integer
    Dim Pdata_codep As String
    Dim Pdata_makerp  As String
    Dim Pdata_kikakup As String
    Dim lines As Integer
    Dim Line_dim As Integer
'
    FLGpage = 1             '*** ページ初期化
    FLGheader = 0           '*** ヘッダー印刷初期化
    Gyoumax = 41            '*** 印刷行最大値 ***
    Goukei = 0              '*** 合計初期化 ***
'
    Line_dim = 4
    ReDim Pdata_lines(CKsuu, Line_dim)
'
    For ip = 1 To Anum0
        DRVindexT = Xcont0(2) & "\" & Aitem0(ip, 0) & "\" & Aitem0(ip, 0) & "INDEX.COD"
'
        Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)    '*** INDEX.COD 読み込み ***
        FLGitemp = 0    '*** 項目印刷ﾌﾗｸﾞｸﾘｱｰ
        pps = 1                 '*** Pdata_totalのスタート位置フラグ ***
        jps = pps
        lines = 0
'
        For jp = 1 To BnumT
            For pp = jps To CKsuu
                If Mid(Pdata_total(pp, 1), 2, 4) = BindexT(jp, 0) Then   '*** コード番号 ***
                    pps = pp
'
                    If FLGheader = 0 Then
                        Call PRNheader0     '*** 部品表ヘッダー印刷 ***
                        Call PRNheader1     '*** 部品表項目見出し印刷 ***
                        FLGheader = 1       '*** ヘッダー印刷済みにｾｯﾄ ***
                    End If
'
                    If FLGitemp = 0 Then
                        If FLGgyou >= Gyoumax - 1 Then
                            FLGend = 0
                            Call PRNfooter
                            Printer.NewPage   '*** 改ページ ***
'
                            Call PRNheader0
                            Call PRNheader1
                            FLGheader = 1   '*** ヘッダー印刷済みにｾｯﾄ ***
                        End If
'
                        Call PRNIndex(ip)   '*** 項目名印刷 ***
                        FLGitemp = 1        '*** 項目印刷済みにｾｯﾄ ***
'
                    End If
'
                    Call SET_DRVmain(DRVmainT, Aitem0(), ip, BindexT(), jp)
'
                    Call RDmain(DRVmainT, CmainT(), CnumT, CdimT)    '*** MAIN.COD 読み込み ***
'
                    For kp = 1 To CnumT
                        Pdata_codep = "L" & BindexT(jp, 0) & "-" & CmainT(kp, 0)
'
                        For np = pps To CKsuu
                            If Pdata_total(np, 1) = Pdata_codep Then   '*** 個別にﾃﾞｰﾀｰ作成
                                If BindexT(jp, 5) = "000" Then
                                    Pdata_makerp = CmainT(kp, 13)
'
                                ElseIf BindexT(jp, 5) = "998" Then
                                    If Pdata_total(np, 3) = "0" Then
                                        MKsitei = "0"
                                    Else
                                        MKsitei = Pdata_total(np, 3)
                                    End If
'
                                    Call GET998maker(Pdata_makerp, MKsitei, BindexT(), jp)
'
                                Else
                                    Pdata_makerp = BindexT(jp, 5)
                                End If
'
                                Call Makerget2(Pdata_makerp)    '*** メーカー略称取得 ***
'
                                Call GETkikaku(Pdata_kikakup, MKsitei, BindexT(), jp, CmainT(), kp) '*** 部品名取得 ***
'
                                If CmainT(kp, 16) = "1" Then         '*** 特記事項記入
                                    If Pdata_total(np, 4) = "" Or Pdata_total(np, 4) = "*" Then
                                        '
                                    Else
                                        Pdata_kikakup = Pdata_kikakup & Pdata_total(np, 4)
                                    End If
                                End If
'
                                Call GET_shitei2(CmainT(kp, 3), Pdata_kikakup, q)   '*** ! ? の記入 ***
'
                                Pdata_lines(lines + 1, 0) = Pdata_codep   '*** データ設定 ***
                                Pdata_lines(lines + 1, 1) = Pdata_kikakup
                                Pdata_lines(lines + 1, 2) = Pdata_total(np, 5)
                                Pdata_lines(lines + 1, 3) = Pdata_makerp
'
                                If optBikou.Value = True Then
                                    Pdata_lines(lines + 1, 4) = Pdata_total(np, 2)
'
                                ElseIf optTanka.Value = True Then
                                    Pdata_lines(lines + 1, 4) = CmainT(kp, 5)
'
                                ElseIf optTainetu.Value = True Then
                                    Tmp0 = CmainT(kp, 6)    '*** MSL記入
                                    Call TRS_Mlevel2(Tmp0)
                                    Tmp1 = Tmp0 & "/"
'
                                    Tmp0 = CmainT(kp, 11)       '*** 耐熱記入
                                    If Len(Tmp0) < 3 Then Tmp0 = "   "  '***3文字未満は無い
                                    Tmp1 = Tmp1 & Tmp0 & "/"
'
                                    If 5 < Len(CmainT(kp, 19)) Then    '*** ﾒｯｷ記入
                                        Tmp1 = Tmp1 & CmainT(kp, 19) & "/"
                                    ElseIf 5 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & " " & CmainT(kp, 19) & "/"
                                    ElseIf 4 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & " " & CmainT(kp, 19) & " /"
                                    ElseIf 3 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & "  " & CmainT(kp, 19) & " /"
                                    ElseIf 2 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & "  " & CmainT(kp, 19) & "  /"
                                    Else
                                        Tmp1 = Tmp1 & "      /"     '*** 1文字の金属メッキは無い ***
                                    End If
'
                                    If InStr(CmainT(kp, 19), "SnPb") <> 0 Then
                                        Tmp1 = Tmp1 & "----"
                                    ElseIf InStr(CmainT(kp, 2), "#<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(CmainT(kp, 2), "<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "RoHS"
                                    ElseIf InStr(CmainT(kp, 2), "#<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(CmainT(kp, 2), "<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "Pbﾌﾘ"
                                    ElseIf InStr(CmainT(kp, 2), "<Green>") <> 0 Then
                                        Tmp1 = Tmp1 & "Gree"
                                    ElseIf InStr(CmainT(kp, 2), "<Ro2>") <> 0 Then      '*** Ver2.1ﾆﾃ追加
                                        Tmp1 = Tmp1 & "Ro2"
                                    ElseIf InStr(CmainT(kp, 2), "<R863>") <> 0 Then     '*** Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "R863"
                                    ElseIf InStr(BindexT(jp, 1), "#<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(BindexT(jp, 1), "<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "RoHS"
                                    ElseIf InStr(BindexT(jp, 1), "#<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(BindexT(jp, 1), "<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "Pbﾌﾘ"
                                    ElseIf InStr(BindexT(jp, 1), "<Green>") <> 0 Then
                                        Tmp1 = Tmp1 & "Gree"
                                    ElseIf InStr(BindexT(jp, 1), "<Ro2>") <> 0 Then     '*** Ver2.1ﾆﾃ追加
                                        Tmp1 = Tmp1 & "Ro2"
                                    ElseIf InStr(BindexT(jp, 1), "<R863>") <> 0 Then    '*** Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "R863"
                                    End If
'
                                    Pdata_lines(lines + 1, 4) = Tmp1
                                End If
'
                                lines = lines + 1
'
'Debug.Print CKsuu, lines, Pdata_lines(lines, 0), Pdata_lines(lines, 1)
                            End If
                        Next np
                    Next kp
                    Exit For    '*** 次の親コードへ ***
'
                End If
            Next pp
        Next jp
'
        Call Print_sort(lines)      '*** 条件によるソーティング ***
'
        For np = 1 To lines
            Call PRNkoumoku(np)     '*** 部品表項目印刷 ***
        Next np
'
        For pp = 1 To CKsuu    '*** 未登録部品の選別 ***
            If Left(Pdata_total(pp, 1), 1) = "*" Then
                Tmp1 = Pdata_total(pp, 0)
                Call GET_koumoku(Tmp1, Aitem0(), ip)
'
                If Tmp1 = Aitem0(ip, 0) Then
                    If FLGheader = 0 Then
                        Call PRNheader0     '*** 部品表ヘッダー印刷 ***
                        Call PRNheader1     '*** 部品表項目見出し印刷 ***
                        FLGheader = 1       '*** ヘッダー印刷済みにｾｯﾄ ***
                    End If
'
                    If FLGitemp = 0 Then
                        Call PRNIndex(ip)   '*** 項目名印刷 ***
                        FLGitemp = 1        '*** 項目印刷済みにｾｯﾄ ***
                    End If
'
                    If FLGgyou >= Gyoumax Then
                        FLGend = 0
                        Call PRNfooter
                        Printer.NewPage   '*** 改ページ ***
'
                        FLGheader = 0
                        Call PRNheader0
                        Call PRNheader1
                        FLGheader = 1   '*** ヘッダー印刷済みにｾｯﾄ ***
                    End If
'
                    Call PRNmitouroku    '*** 未登録部品印刷 ***
                End If
            End If
        Next pp
    Next ip
'
    FLGend = 1
    Call PRNfooter          '*** 部品表項末印刷 ***
'
    For pp = 1 To CKsuu
        Tmp1 = Pdata_total(pp, 0)
        Call GET_koumoku(Tmp1, Aitem0(), Anum0)
'
        If Tmp1 = "**" Then
            Call PRNhamidashi(pp)   '*** はみ出し部品印刷 ***
        End If
    Next pp
'
    Printer.EndDoc      '*** プリンター書き込み ***
'
End Sub

Private Sub Sougou_insatu()     '*** 総合印刷 ***
    Dim FLGbanme As Integer
    Dim Tmp1 As String
    Dim i As Integer
'
    Me.Caption = "!!! 同一部品を集計中 !!!"
    Me.MousePointer = vbHourglass
    DoEvents
'
    ReDim Pdata_total(CKmax, Ckdim)
    CKsuu = 0
    DoEvents
'
    For FLGbanme = 1 To KtotalT
        If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
            PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP             '*** .PLT を検索 ***
            Tmp1 = Dir(DRVpartlistP)
            If Tmp1 = "" Then
                PFLnameP = KLSTT(FLGbanme, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
'
                DRVpartlistP = TMPplst & "\" & PFLnameP         '*** xxxxxxxx.yyy を検索 ***
                Tmp1 = Dir(DRVpartlistP)
'                If Tmp1 = "" Then
'                    PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
'                    DRVpartlistP = TMPplst & "\" & PFLnameP    '*** 共有フォルダーを検索 ***
'                    Tmp1 = Dir(DRVpartlistP)
                    If Tmp1 = "" Then
                        i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                        GoTo skip
'
                    End If
'                End If
            End If
            Baisuu = Val(KLSTT(FLGbanme, 4))
'
            Call Plst_yomu           '*** 部品表を読む ***
            DoEvents
'
            Call Plst_total          '*** 同一品目の集計 ***
            DoEvents
        End If
    Next FLGbanme
'
    Me.Caption = " !!! 部品数量表をプリントバッファーに転送中 !!!"
    DoEvents
'
    Call Insatsu_MAIN    '*** 印刷ルーチン ***
'
skip:
    Me.MousePointer = vbDefault
    DoEvents
    Timer1.Enabled = True
'    Unload Me
End Sub

Private Sub Kobetu_insatu()     '*** 個別印刷 ***
    Dim FLGbanme As Integer
    Dim Tmp1 As String
    Dim i As Integer
'
    For FLGbanme = 1 To KtotalT
        If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
            PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP             '*** .PLT を検索 ***
            Tmp1 = Dir(DRVpartlistP)
            If Tmp1 = "" Then
                PFLnameP = KLSTT(FLGbanme, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
'
                DRVpartlistP = TMPplst & "\" & PFLnameP         '*** xxxxxxxx.yyy を検索 ***
                Tmp1 = Dir(DRVpartlistP)
'                If Tmp1 = "" Then
'                    PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
'                    DRVpartlistP = TMPplst & "\" & PFLnameP    '*** 共有フォルダーを検索 ***
'                    Tmp1 = Dir(DRVpartlistP)
                    If Tmp1 = "" Then
                        i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                        GoTo skip
'
                    End If
'                End If
            End If
            Baisuu = Val(KLSTT(FLGbanme, 4))
'
            TempFname = KLSTT(FLGbanme, 2)
            Me.Caption = " !!! " & TempFname & " の同一部品を集計中 !!!"
            Me.MousePointer = vbHourglass
            DoEvents
'
            ReDim Pdata_total(CKmax, Ckdim) '*** 配列初期化 ***
            CKsuu = 0
            DoEvents
'
            Call Plst_yomu           '*** 部品表を読む ***
            DoEvents
'
            Call Plst_total          '*** 同一品目の集計 ***
            DoEvents
'
            Me.Caption = " !!! " & TempFname & " の部品数量表をプリントバッファーに転送中 !!!"
            DoEvents
'
            Call Insatsu_MAIN    '*** 印刷ルーチン ***
            DoEvents
        End If
    Next FLGbanme
'
skip:
    Me.MousePointer = vbDefault
    DoEvents
    Timer1.Enabled = True
'    Unload Me
End Sub

Private Sub Hitotu_insatu()     '*** 単独印刷 ***
    Me.Caption = "!!! 同一部品を集計中 !!!"
    Me.MousePointer = vbHourglass
    DoEvents
'
    ReDim Pdata_total(CKmax, Ckdim)
    CKsuu = 0
    DoEvents
'
    Call Plst_total          '*** 同一品目の集計 ***
'
    Me.Caption = " !!! 部品数量表をプリントバッファーに転送中 !!!"
    DoEvents
'
    Call Insatsu_MAIN    '*** 印刷ルーチン ***
'
    Me.MousePointer = vbDefault
    DoEvents
    Timer1.Enabled = True
'    Unload Me
End Sub

Private Sub file_MAIN()     '*** ファイル出力ルーチン ***
    Dim i As Integer, ip As Integer, j As Integer, jp As Integer, np As Integer, q As Integer
    Dim jps As Integer, pps As Integer
    Dim kp As Integer
    Dim Gyoumax As Integer
    Dim MKsitei As String
    Dim Tmp1 As String, Tmp0 As String
    Dim FLGitti As Integer, FLGmojime As Integer
    Dim Pdata_codep As String
    Dim Pdata_makerp  As String
    Dim Pdata_kikakup As String
    Dim lines As Integer
    Dim Line_dim As Integer
'
    Line_dim = 7
    ReDim Pdata_lines(CKsuu, Line_dim)
    Goukei = 0              '*** 合計初期化 ***
'
    Call WR_suuryou(0, 0, "", "", 0)    '*** 数量表項目ファイル出力 ***
'
    For ip = 1 To Anum0
        DRVindexT = Xcont0(2) & "\" & Aitem0(ip, 0) & "\" & Aitem0(ip, 0) & "INDEX.COD"
        Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)    '*** INDEX.COD 読み込み ***
'
        FLGitemp = 0    '*** 項目印刷ﾌﾗｸﾞｸﾘｱｰ
        pps = 1                 '*** Pdata_totalのスタート位置フラグ ***
        jps = pps
'
        lines = 0
'
        For jp = 1 To BnumT
            For pp = jps To CKsuu
                If Mid(Pdata_total(pp, 1), 2, 4) = BindexT(jp, 0) Then   '*** コード番号 ***
                    pps = pp
'
                    Call SET_DRVmain(DRVmainT, Aitem0(), ip, BindexT(), jp)
'
                    Call RDmain(DRVmainT, CmainT(), CnumT, CdimT)   '*** MAIN.COD 読み込み ***
'
                    For kp = 1 To CnumT
                        Pdata_codep = "L" & BindexT(jp, 0) & "-" & CmainT(kp, 0)
'
                        For np = pps To CKsuu
                            If Pdata_total(np, 1) = Pdata_codep Then   '*** 個別にﾃﾞｰﾀｰ作成
                                If BindexT(jp, 5) = "000" Then
                                    Pdata_makerp = CmainT(kp, 13)
'
                                ElseIf BindexT(jp, 5) = "998" Then
                                    If Pdata_total(np, 3) = "0" Then
                                        MKsitei = "0"
                                    Else
                                        MKsitei = Pdata_total(np, 3)
                                    End If
'
                                    Call GET998maker(Pdata_makerp, MKsitei, BindexT(), jp)
'
                                Else
                                    Pdata_makerp = BindexT(jp, 5)
                                End If
'
                                Call Makerget2(Pdata_makerp)    '*** メーカー略称取得 ***
'
                                Call GETkikaku(Pdata_kikakup, MKsitei, BindexT(), jp, CmainT(), kp) '*** 部品名取得 ***
'
                                If CmainT(kp, 16) = "1" Then         '*** 特記事項記入
                                    If Pdata_total(np, 4) = "" Or Pdata_total(np, 4) = "*" Then
                                        '
                                    Else
                                        Pdata_kikakup = Pdata_kikakup & Pdata_total(np, 4)
                                    End If
                                End If
'
                                Call GET_shitei2(CmainT(kp, 3), Pdata_kikakup, q)   '*** ! ? の記入 ***
'
                                Call WR_suuryou(1, np, Pdata_kikakup, Pdata_makerp, lines)  '*** 数量表項目バッファ出力 ***
'
                                If optBikou.Value = True Then
                                    Pdata_lines(lines + 1, 6) = Pdata_total(np, 2)
'
                                    Tmp0 = CmainT(kp, 6)    '*** MSL記入
                                    Call TRS_Mlevel2(Tmp0)
                                    Pdata_lines(lines + 1, 7) = Tmp0
                                End If
'
                                If optTanka.Value = True Then
                                    Pdata_lines(lines + 1, 6) = CmainT(kp, 5)
                                End If
'
                                If optTainetu.Value = True Then
                                    Tmp0 = CmainT(kp, 6)    '*** MSL記入
                                    Call TRS_Mlevel2(Tmp0)
                                    Tmp1 = Tmp0 & "/"
'
                                    Tmp0 = CmainT(kp, 11)       '*** 耐熱記入
                                    If Len(Tmp0) < 3 Then Tmp0 = "   "  '***3文字未満は無い
                                    Tmp1 = Tmp1 & Tmp0 & "/"
'
                                    If 5 < Len(CmainT(kp, 19)) Then    '*** ﾒｯｷ記入
                                        Tmp1 = Tmp1 & CmainT(kp, 19) & "/"
                                    ElseIf 5 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & " " & CmainT(kp, 19) & "/"
                                    ElseIf 4 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & " " & CmainT(kp, 19) & " /"
                                    ElseIf 3 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & "  " & CmainT(kp, 19) & " /"
                                    ElseIf 2 = Len(CmainT(kp, 19)) Then
                                        Tmp1 = Tmp1 & "  " & CmainT(kp, 19) & "  /"
                                    Else
                                        Tmp1 = Tmp1 & "      /"     '*** 1文字の金属メッキは無い ***
                                    End If
'
                                    If InStr(CmainT(kp, 19), "SnPb") <> 0 Then
                                        Tmp1 = Tmp1 & "----"
                                    ElseIf InStr(CmainT(kp, 2), "#<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(CmainT(kp, 2), "<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "RoHS"
                                    ElseIf InStr(CmainT(kp, 2), "#<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(CmainT(kp, 2), "<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "Pbﾌﾘ"
                                    ElseIf InStr(CmainT(kp, 2), "<Green>") <> 0 Then
                                        Tmp1 = Tmp1 & "Gree"
                                    ElseIf InStr(CmainT(kp, 2), "<Ro2>") <> 0 Then      '*** Ver2.1ﾆﾃ追加忘れ,Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "Ro2"
                                    ElseIf InStr(CmainT(kp, 2), "<R863>") <> 0 Then     '*** Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "R863"
                                    ElseIf InStr(BindexT(jp, 1), "#<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(BindexT(jp, 1), "<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "RoHS"
                                    ElseIf InStr(BindexT(jp, 1), "#<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(BindexT(jp, 1), "<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "Pbﾌﾘ"
                                    ElseIf InStr(BindexT(jp, 1), "<Green>") <> 0 Then
                                        Tmp1 = Tmp1 & "Gree"
                                    ElseIf InStr(BindexT(jp, 1), "<Ro2>") <> 0 Then     '*** Ver2.1ﾆﾃ追加忘れ,Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "Ro2"
                                    ElseIf InStr(BindexT(jp, 1), "<R863>") <> 0 Then    '*** Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "R863"
                                    End If
'
                                    Pdata_lines(lines + 1, 6) = Tmp1
                                End If
'
                                lines = lines + 1
'
                            End If
                        Next np
                    Next kp
                    Exit For    '*** 次の親コードへ ***
'
                End If
            Next pp
        Next jp
'
        Call file_SORT(FLGsort, lines) '*** 条件によるソーティング ***
'
        For np = 1 To lines
            Call WR_suuryou(2, 0, "", "", np)    '*** 数量表項目ファイル出力 ***
        Next np
'
        For pp = 1 To CKsuu    '*** 未登録部品の選別 ***
            If Left(Pdata_total(pp, 1), 1) = "*" Then
                Tmp1 = Pdata_total(pp, 0)
                Call GET_koumoku(Tmp1, Aitem0(), Anum0)
'
                If Tmp1 = Aitem0(ip, 0) Then
                    Pdata_kikakup = Mid(Pdata_total(pp, 1), 2)
                    Pdata_makerp = "*"
                    Call WR_suuryou(3, pp, Pdata_kikakup, Pdata_makerp, 0)  '*** 数量表項目ファイル出力 ***
'
                End If
            End If
        Next pp
    Next ip
'
    Call WR_suuryou(9, 0, " ", " ", 0)  '*** 数量表項目ファイル出力 ***
'
End Sub
        
Private Sub file_SORT(flg As Integer, Tlines As Integer)    '*** 条件によるソーティング ***
    Dim i As Integer, j As Integer, k As Integer
    Dim Ptemp(6) As String
'
    Select Case flg
    Case 0  '*** ｺｰﾄﾞ出現順 ***
        '--- 変更なし ---
'
    Case 1  '*** 部品名称順 ***
        For i = 1 To Tlines - 1
            For j = Tlines - i To Tlines - 1
                If Pdata_lines(j + 1, 2) < Pdata_lines(j, 2) Then
                    For k = 0 To 7
                        Ptemp(k) = Pdata_lines(j, k)                '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 7
                        Pdata_lines(j, k) = Pdata_lines(j + 1, k)   '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 7
                        Pdata_lines(j + 1, k) = Ptemp(k)            '*** 入れ替え ***
                    Next k
                End If
            Next j
        Next i
'
    Case 2  '*** ﾒｰｶｰ出現順 ***
        For i = 1 To Tlines - 1
            For j = Tlines - i To Tlines - 1
                If Pdata_lines(j + 1, 5) < Pdata_lines(j, 5) Then
                    For k = 0 To 7
                        Ptemp(k) = Pdata_lines(j, k)                '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 7
                        Pdata_lines(j, k) = Pdata_lines(j + 1, k)   '*** 入れ替え ***
                    Next k
'
                    For k = 0 To 7
                        Pdata_lines(j + 1, k) = Ptemp(k)            '*** 入れ替え ***
                    Next k
                End If
            Next j
        Next i
'
    End Select
End Sub

Private Sub Sougou_file()       '*** 総合ファイル出力 ***
    Dim i As Integer
    Dim FLGbanme As Integer
    Dim MSGcomment As String
    Dim Tmp1 As String
    Dim zubanTemp As String
'
    If Right(UCase(ZubanT), 4) = ".PLT" Then
        i = Len(ZubanT)
        zubanTemp = Mid(ZubanT, 1, i - 4)
    Else
        zubanTemp = ZubanT
    End If
'
    MSGcomment = "それでは " & CATnameT & " の部品数量表をファイル出力します。" & vbCrLf _
            & "ファイル名は K" & zubanTemp & ".CSV になります。" & vbCrLf _
            & "ファイルはちょっと大きくなるのでディスクの残容量に注意してください。"
'
    i = MsgBox(MSGcomment, vbExclamation Or vbOKCancel)
    If i = vbCancel Then GoTo errh_P    '*** キャンセルで処理中止 ***
'
    Me.Caption = "!!! 同一部品を集計中 !!!"
    Me.MousePointer = vbHourglass
    DoEvents
'
    ReDim Pdata_total(CKmax, Ckdim)
    CKsuu = 0
    DoEvents
'
    For FLGbanme = 1 To KtotalT
        If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
            PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP             '*** .PLT を検索 ***
            Tmp1 = Dir(DRVpartlistP)
            If Tmp1 = "" Then
                PFLnameP = KLSTT(FLGbanme, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
'
                DRVpartlistP = TMPplst & "\" & PFLnameP         '*** xxxxxxxx.yyy を検索 ***
                Tmp1 = Dir(DRVpartlistP)
'                If Tmp1 = "" Then
'                    PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
'                    DRVpartlistP = TMPplst & "\" & PFLnameP    '*** 共有フォルダーを検索 ***
'                    Tmp1 = Dir(DRVpartlistP)
                    If Tmp1 = "" Then
                        i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                        GoTo skip
'
                    End If
'                End If
            End If
'
            Baisuu = Val(KLSTT(FLGbanme, 4))
            Call Plst_yomu           '*** 部品表を読む ***
            DoEvents
'
            Call Plst_total          '*** 同一品目の集計 ***
            DoEvents
        End If
    Next FLGbanme
'
    Me.Caption = " !!! 部品数量表をファイルに出力中 !!!"
    DoEvents
'
    FNAME_suuryou = TMPdir1 & "\K" & zubanTemp & ".CSV"
    Call file_MAIN   '*** ファイル出力ルーチン ***
'
skip:
    Me.MousePointer = vbDefault
    DoEvents
    Timer1.Enabled = True
    Exit Sub
'    Unload Me
'
errh_P:
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
End Sub

Private Sub Kobetu_file()       '*** 個別ファイル出力 ***
    Dim i As Integer, j As Integer
    Dim nagasa As Integer
    Dim FLGbanme As Integer
    Dim MSGcomment As String
    Dim WRname() As String
    Dim Tmp1 As String
'
    ReDim WRname(KtotalT)
'
    MSGcomment = "それでは " & CATnameT & " の部品数量表を個別にファイル出力します。" & vbCrLf _
        & "ファイル名は ･･･Bxxxx-xx を[K]に変更した Kxxxx-xx.CSV になります。" & vbCrLf _
        & "ファイルはちょっと大きくなるのでディスクの残容量に注意してください。"
'
    i = MsgBox(MSGcomment, vbExclamation Or vbOKCancel)
    If i = vbCancel Then GoTo errh_P    '*** キャンセルで処理中止 ***
'
    Me.MousePointer = vbHourglass
    DoEvents
'
    For FLGbanme = 1 To KtotalT
        If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
            PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP             '*** .PLT を検索 ***
            Tmp1 = Dir(DRVpartlistP)
            If Tmp1 = "" Then
                PFLnameP = KLSTT(FLGbanme, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
'
                DRVpartlistP = TMPplst & "\" & PFLnameP         '*** xxxxxxxx.yyy を検索 ***
                Tmp1 = Dir(DRVpartlistP)
'                If Tmp1 = "" Then
'                    PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
'                    DRVpartlistP = TMPplst & "\" & PFLnameP    '*** 共有フォルダーを検索 ***
'                    Tmp1 = Dir(DRVpartlistP)
                    If Tmp1 = "" Then
                        i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                        GoTo skip
'
                    End If
'                End If
            End If
            Baisuu = Val(KLSTT(FLGbanme, 4))
'
            TempFname = KLSTT(FLGbanme, 2)
            Me.Caption = "!!! " & TempFname & " の同一部品を集計中 !!!"
'
            ReDim Pdata_total(CKmax, Ckdim)
            CKsuu = 0
            DoEvents
'
            Call Plst_yomu           '*** 部品表を読む ***
            DoEvents
'
            Call Plst_total          '*** 同一品目の集計 ***
'
            Me.Caption = " !!! 部品数量表 " & TempFname & " をファイルに出力中 !!!"
            DoEvents
'
            If Right(UCase(PFLnameP), 4) = ".PLT" Then
                i = Len(PFLnameP)
                FNAME_suuryou = TMPdir1 & "\K" & Mid(PFLnameP, 1 + 1, i - 4 - 1) & ".CSV"
            Else
                If InStr(1, PFLnameP, ".") = 9 Then
                    PFLnameP = Left(PFLnameP, 8) & Mid(PFLnameP, 10)
                End If
                FNAME_suuryou = TMPdir1 & "\K" & Mid(PFLnameP, 2) & ".CSV"
            End If
'
            For j = 1 To FLGbanme - 1   '*** ファイル名の重なりをチェックし名前を変える。 ***
                If (FNAME_suuryou = WRname(j)) Then
                    nagasa = Len(FNAME_suuryou)
                    FNAME_suuryou = Left(FNAME_suuryou, nagasa - 4) & "I.CSV"
                End If
            Next
            WRname(FLGbanme) = FNAME_suuryou
'
            Call file_MAIN   '*** ファイル出力ルーチン ***
        End If
    Next FLGbanme
'
skip:
    Me.MousePointer = vbDefault
    DoEvents
    Timer1.Enabled = True
    Exit Sub
'    Unload Me
'
errh_P:
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
End Sub

Private Sub Hitotu_file()       '*** 単独ファイル出力 ***
    Dim i As Integer
    Dim MSGcomment As String
    Dim zubanTemp As String
'
    If Right(UCase(PFLnameP), 4) = ".PLT" Then
        i = Len(PFLnameP)
        zubanTemp = Mid(PFLnameP, 1, i - 4)
    Else
        zubanTemp = PFLnameP
    End If
'
    MSGcomment = "それでは " & zubanTemp & " の部品数量表をファイル出力します。" & vbCrLf _
        & "ファイル名は ･･･Bxxxx-xx を[K]に変更した Kxxxx-xx.CSV になります。" & vbCrLf _
        & "ファイルはちょっと大きくなるのでディスクの残容量に注意してください。"
'
    i = MsgBox(MSGcomment, vbExclamation Or vbOKCancel)
    If i = vbCancel Then GoTo errh_P    '*** キャンセルで処理中止 ***
'
    Me.Caption = "!!! 同一部品を集計中 !!!"
    Me.MousePointer = vbHourglass
    DoEvents
'
    ReDim Pdata_total(CKmax, Ckdim)
    CKsuu = 0
    DoEvents
'
    Call Plst_total          '*** 同一品目の集計 ***
'
    Me.Caption = " !!! 部品数量表をファイルに出力中 !!!"
    DoEvents
'
    If Right(UCase(PFLnameP), 4) = ".PLT" Then
        i = Len(PFLnameP)
        FNAME_suuryou = TMPdir1 & "\K" & Mid(PFLnameP, 1 + 1, i - 1 - 4) & ".CSV"
    Else
        If InStr(1, PFLnameP, ".") = 9 Then
            PFLnameP = Left(PFLnameP, 8) & Mid(PFLnameP, 10)
        End If
        FNAME_suuryou = TMPdir1 & "\K" & Mid(PFLnameP, 2) & ".CSV"
    End If
'
    Call file_MAIN   '*** ファイル出力ルーチン ***
'
    Me.MousePointer = vbDefault
    DoEvents
    Timer1.Enabled = True
    Exit Sub
'    Unload Me
'
errh_P:
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdFile_Click()     '*** 数量表ファイル出力 ***
    CKmax = 200
    Ckdim = 5
'
    If FLGall = 1 Then
        If optSougou.Value = True Then
            Call Sougou_file '*** 総合 ***
        Else
            Call Kobetu_file '*** 個別 ***
        End If
    Else
        Call Hitotu_file     '*** 単独 ***
    End If
End Sub

Private Sub cmdGo_Click()       '*** 数量表印刷 ***
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
    CKmax = 200
    Ckdim = 5
'
    If FLGall = 1 Then
        If optSougou.Value = True Then
            Call Sougou_insatu   '*** 総合 ***
        Else
            Call Kobetu_insatu   '*** 個別 ***
        End If
    Else
        Call Hitotu_insatu       '*** 単独 ***
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub S_Hyouji()
    optCode.Value = True
    FLGsort = 0
'
    optBuhinmei.Value = True
    optBikou.Value = True
    FLGsoutouhin = 0
'
'    cmdGo.ToolTipText = "プリンタの印刷設定を変更するには" _
'                    & "コントロールパネルの「プリンタとFAX」で行ってください。"
'
    If FLGall = 0 Then      '*** １つの部品表印刷 ***
        optdirect.Value = True
        FLGkouseihyou = 0
'
        optSougou.Enabled = False
        optKobetu.Value = True
'
        txtkeisiki.Text = ""
        txtmeisyou.Text = ""
        txttantou.Text = ""
        txtkouban.Text = ""
        txtdaisuu.Text = ""
        Baisuu = "1"        '*** 仮数 ***
'
        PFLnameP = PFLnameT
        DRVpartlistP = DRVpartlistT
'
        Call Plst_yomu      '*** 部品表を読みﾌｧｲﾙ内容表示 ***
'
    Else                    '*** 構成表より連続印刷 ***
        opthyou.Value = True
        optdirect.Enabled = False
        FLGkouseihyou = 1
'
        optSougou.Value = True
'
        Call Klst_yomu       '*** 構成表を読み込み内容を表示 ***
'
        lblnamae.Caption = "印刷 ﾌｧｲﾙ名 ： <<構成表による>>"
        lblkomei.Caption = "小名称 ： "
        lblDate.Caption = "日付 ： "
        txtbaisuu.Text = ""
        lblkomei.Enabled = False
        lblDate.Enabled = False
        txtbaisuu.Enabled = False
        lblbaisuu.Enabled = False
    End If
End Sub

Private Sub Plst_total()    '*** 同一項目の集計 ***
    Dim i As Integer, j As Integer
    Dim FLGumu As Integer
'
    For i = 1 To PtotalP
        FLGumu = 0
        For j = 1 To CKsuu                          '*** 部品記号は一致しなくても良い ***
            If PLSTP(i, 1) = Pdata_total(j, 1) _
                And PLSTP(i, 2) = Pdata_total(j, 2) _
                    And PLSTP(i, 3) = Pdata_total(j, 3) _
                        And PLSTP(i, 4) = Pdata_total(j, 4) Then
'
                Pdata_total(j, 5) = str(Val(Pdata_total(j, 5)) + Val(Baisuu))
                FLGumu = 1
                Exit For
'
            End If
        Next j
'
        If CKsuu = 0 Or FLGumu = 0 Then
            UPckmax     '*** Pdata_total()の配列を必要なら増やす ***
'
            Pdata_total(CKsuu + 1, 0) = PLSTP(i, PdimP + 1) ' 0: 項目
            Pdata_total(CKsuu + 1, 1) = PLSTP(i, 1)         ' 1: 部品ｺｰﾄﾞ
            Pdata_total(CKsuu + 1, 2) = PLSTP(i, 2)         ' 2: 備考
            Pdata_total(CKsuu + 1, 3) = PLSTP(i, 3)         ' 3: メーカ指定
            Pdata_total(CKsuu + 1, 4) = PLSTP(i, 4)         ' 4: 特記事項
            Pdata_total(CKsuu + 1, 5) = Baisuu              ' 5: 個数 ***
            CKsuu = CKsuu + 1
        End If
    Next i
'
End Sub

Private Sub UPckmax()
'                       *** Pdata_total()の配列を必要なら増やす ***
    Dim i As Integer, j As Integer
    Dim Ptemp() As String
'
    If CKsuu = CKmax Then
        ReDim Ptemp(CKmax, Ckdim)
'
        For i = 1 To CKmax
            For j = 0 To Ckdim
                Ptemp(i, j) = Pdata_total(i, j)
            Next j
        Next i
'
        ReDim Pdata_total(CKmax + 50, Ckdim)
'
        For i = 1 To CKmax
            For j = 0 To Ckdim
                Pdata_total(i, j) = Ptemp(i, j)
            Next j
        Next i
        CKmax = CKmax + 50
    End If
End Sub
    
Private Sub WR_suuryou(FLG_WR_file As Integer, np As Integer, Pdata_kikakup As String, Pdata_makerp As String, line As Integer)
                                        '*** 数量表項目ファイル出力 ***
    Dim Ptemp As String, Ptemp2 As String
    Dim i As Integer
'
    Select Case FLG_WR_file
    Case 0      '*** ファイルオープン,ヘッダー記入 ***
        Open FNAME_suuryou For Output As #1
            If optTainetu.Value = True Then
                Write #1, "項目", "コード番号", "規格・定格", "個数", "総数", "メーカー", "MSL/耐熱/ ﾒｯｷ /RoHS"
            ElseIf optTanka.Value = True Then
                Write #1, "項目", "コード番号", "規格・定格", "個数", "総数", "メーカー", "単価"
            Else
                If Xcont0(8) = "1" Then
                    Write #1, "項目", "コード番号", "規格・定格", "個数", "総数", "メーカー", "備考", "MSL"
                Else
                    Write #1, "項目", "コード番号", "規格・定格", "個数", "総数", "メーカー", "備考"
                End If
            End If
'
            If FLGall = 1 And optSougou.Value = True Then
                If Xcont0(8) = "1" Then
                    Write #1, "*", "型式：" & CATnoT, "名称：" & CATnameT, "電 気 部 品 数 量 表", "工番：" & KoubanT, "台数：" & DaisuuT, "*", "*"
                Else
                    Write #1, "*", "型式：" & CATnoT, "名称：" & CATnameT, "電 気 部 品 数 量 表", "工番：" & KoubanT, "台数：" & DaisuuT, "*"
                End If
            Else
                If Xcont0(8) = "1" Then
                    Write #1, "*", "型式：" & CATnoT, "小名称：" & TempFname & "; " & PlistnameP, "電 気 部 品 数 量 表", "工番：" & KoubanT, "台数：" & DaisuuT, "*", "*"
                Else
                    Write #1, "*", "型式：" & CATnoT, "小名称：" & TempFname & "; " & PlistnameP, "電 気 部 品 数 量 表", "工番：" & KoubanT, "台数：" & DaisuuT, "*"
                End If
            End If
'
    Case 1      '*** 内容記入 ***
            Ptemp = Pdata_total(np, 0)
            Call GET_koumoku(Ptemp, Aitem0(), Anum0)
'
            Pdata_lines(line + 1, 0) = Ptemp
            Pdata_lines(line + 1, 1) = Pdata_total(np, 1)
            Pdata_lines(line + 1, 2) = Pdata_kikakup
            Pdata_lines(line + 1, 3) = Pdata_total(np, 5)
            Pdata_lines(line + 1, 4) = str(Val(Pdata_total(np, 5)) * Val(DaisuuT))
            Pdata_lines(line + 1, 5) = Pdata_makerp
            Pdata_lines(line + 1, 6) = Pdata_total(np, 2)
'
    Case 2
            If FLGsoutouhin = 1 And optBikou.Value = True Then
                i = InStr(1, Pdata_lines(line, 2), "相当")
                If i <> 0 Then
                    Pdata_lines(line, 2) = Left(Pdata_lines(line, 2), i - 1)
                    If Pdata_lines(line, 6) = "*" Then
                        Pdata_lines(line, 6) = "相当"
                    Else
                        Pdata_lines(line, 6) = "相当," & Pdata_lines(line, 6)
                    End If
                End If
            End If
            Write #1, Pdata_lines(line, 0), Pdata_lines(line, 1), Pdata_lines(line, 2), Pdata_lines(line, 3), Pdata_lines(line, 4), Pdata_lines(line, 5), Pdata_lines(line, 6), Pdata_lines(line, 7)
'
    Case 3      '*** 内容記入 ***
            Ptemp = Pdata_total(np, 0)
            Call GET_koumoku(Ptemp, Aitem0(), Anum0)
'
            If optTainetu.Value = True Then
                Write #1, Ptemp, "未登録部品", Pdata_kikakup, Pdata_total(np, 5), str(Val(Pdata_total(np, 5)) * Val(DaisuuT)), Pdata_makerp, "---"
            ElseIf optTanka.Value = True Then
                Write #1, Ptemp, "未登録部品", Pdata_kikakup, Pdata_total(np, 5), str(Val(Pdata_total(np, 5)) * Val(DaisuuT)), Pdata_makerp, "0"
            Else
                If Xcont0(8) = "1" Then
                    Write #1, Ptemp, "未登録部品", Pdata_kikakup, Pdata_total(np, 5), str(Val(Pdata_total(np, 5)) * Val(DaisuuT)), Pdata_makerp, Pdata_total(np, 2), "*"
                Else
                    Write #1, Ptemp, "未登録部品", Pdata_kikakup, Pdata_total(np, 5), str(Val(Pdata_total(np, 5)) * Val(DaisuuT)), Pdata_makerp, Pdata_total(np, 2)
                End If
            End If
'
    Case 9      '*** ファイルクローズ ***
        Close #1
'
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub

Private Sub Timer1_Timer()
    Call cmdQuit_Click
End Sub

Private Sub optBikourann_Click()
    FLGsoutouhin = 1
End Sub

Private Sub optBuhinmei_Click()
    FLGsoutouhin = 0
End Sub

Private Sub optCode_Click()
    FLGsort = 0
End Sub

Private Sub opthyou_Click()
    Klst_yomu       '*** 構成表を読み込み内容表示 ***
End Sub

Private Sub optMaker_Click()
    FLGsort = 2
End Sub

Private Sub optName_Click()
    FLGsort = 1
End Sub

Private Sub txtbaisuu_Click()
    txtbaisuu.MousePointer = vbIbeam
End Sub

Private Sub txtbaisuu_LostFocus()
    txtbaisuu.MousePointer = vbArrow
    Baisuu = Trim(txtbaisuu.Text)
End Sub

Private Sub txtdaisuu_Click()
    txtdaisuu.MousePointer = vbIbeam
End Sub

Private Sub txtdaisuu_LostFocus()
    txtdaisuu.MousePointer = vbArrow
    DaisuuT = Trim(txtdaisuu.Text)
End Sub

Private Sub txtKeisiki_Click()
    txtkeisiki.MousePointer = vbIbeam
End Sub

Private Sub txtKeisiki_LostFocus()
    txtkeisiki.MousePointer = vbArrow
    CATnoT = Trim(txtkeisiki.Text)
End Sub

Private Sub txtkouban_Click()
    txtkouban.MousePointer = vbIbeam
End Sub

Private Sub txtkouban_LostFocus()
    txtkouban.MousePointer = vbArrow
    KoubanT = Trim(txtkouban.Text)
End Sub

Private Sub txtmeisyou_Click()
    txtmeisyou.MousePointer = vbIbeam
End Sub

Private Sub txtmeisyou_LostFocus()
    txtmeisyou.MousePointer = vbArrow
    CATnameT = Trim(txtmeisyou.Text)
End Sub

Private Sub txttantou_Click()
    txttantou.MousePointer = vbIbeam
End Sub

Private Sub txttantou_LostFocus()
    txttantou.MousePointer = vbArrow
    PersonT = Trim(txttantou.Text)
End Sub

Private Sub DSPgamenBuhin()
    txtkeisiki.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtkeisiki.Top = 480
    txtkeisiki.FontSize = 10 * HyoujiBairitu
    txtkeisiki.Width = 1215 * HyoujiBairitu
    txtkeisiki.Height = 285 * HyoujiBairitu
'
    lblkeisiki.Left = 480
    lblkeisiki.Top = 480
    lblkeisiki.FontSize = 10 * HyoujiBairitu
    lblkeisiki.Width = 855 * HyoujiBairitu
    lblkeisiki.Height = txtkeisiki.Height
'
    txtmeisyou.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtmeisyou.Top = 480 + (840 - 480) * HyoujiBairitu
    txtmeisyou.FontSize = 10 * HyoujiBairitu
    txtmeisyou.Width = 3135 * HyoujiBairitu
    txtmeisyou.Height = 285 * HyoujiBairitu
'
    lblmeisyou.Left = 480
    lblmeisyou.Top = 480 + (840 - 480) * HyoujiBairitu
    lblmeisyou.FontSize = 10 * HyoujiBairitu
    lblmeisyou.Width = 855 * HyoujiBairitu
    lblmeisyou.Height = txtmeisyou.Height
'
    txttantou.Left = 480 + (1335 - 480) * HyoujiBairitu
    txttantou.Top = 480 + (1200 - 480) * HyoujiBairitu
    txttantou.FontSize = 10 * HyoujiBairitu
    txttantou.Width = 1215 * HyoujiBairitu
    txttantou.Height = 285 * HyoujiBairitu
'
    lbltantou.Left = 480
    lbltantou.Top = 480 + (1200 - 480) * HyoujiBairitu
    lbltantou.FontSize = 10 * HyoujiBairitu
    lbltantou.Width = 855 * HyoujiBairitu
    lbltantou.Height = txttantou.Height
'
    txtkouban.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtkouban.Top = 480 + (1560 - 480) * HyoujiBairitu
    txtkouban.FontSize = 10 * HyoujiBairitu
    txtkouban.Width = 3135 * HyoujiBairitu
    txtkouban.Height = 285 * HyoujiBairitu
'
    lblkouban.Left = 480
    lblkouban.Top = 480 + (1560 - 480) * HyoujiBairitu
    lblkouban.FontSize = 10 * HyoujiBairitu
    lblkouban.Width = 855 * HyoujiBairitu
    lblkouban.Height = txtkouban.Height
'
    txtdaisuu.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtdaisuu.Top = 480 + (1920 - 480) * HyoujiBairitu
    txtdaisuu.FontSize = 10 * HyoujiBairitu
    txtdaisuu.Width = 615 * HyoujiBairitu
    txtdaisuu.Height = 285 * HyoujiBairitu
'
    lbldaisuu.Left = 480
    lbldaisuu.Top = 480 + (1920 - 480) * HyoujiBairitu
    lbldaisuu.FontSize = 10 * HyoujiBairitu
    lbldaisuu.Width = 855 * HyoujiBairitu
    lbldaisuu.Height = txtdaisuu.Height
'
    lblnamae.Left = 480
    lblnamae.Top = 480 + (2520 - 480) * HyoujiBairitu
    lblnamae.FontSize = 10 * HyoujiBairitu
    lblnamae.Width = 3495 * HyoujiBairitu
    lblnamae.Height = 285 * HyoujiBairitu
'
    lblkomei.Left = 480
    lblkomei.Top = 480 + (2880 - 480) * HyoujiBairitu
    lblkomei.FontSize = 10 * HyoujiBairitu
    lblkomei.Width = 3975 * HyoujiBairitu
    lblkomei.Height = 285 * HyoujiBairitu
'
    lblDate.Left = 480
    lblDate.Top = 480 + (3240 - 480) * HyoujiBairitu
    lblDate.FontSize = 10 * HyoujiBairitu
    lblDate.Width = 2535 * HyoujiBairitu
    lblDate.Height = 285 * HyoujiBairitu
'
    txtbaisuu.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtbaisuu.Top = 480 + (3600 - 480) * HyoujiBairitu
    txtbaisuu.FontSize = 10 * HyoujiBairitu
    txtbaisuu.Width = 615 * HyoujiBairitu
    txtbaisuu.Height = 285 * HyoujiBairitu
'
    lblbaisuu.Left = 480
    lblbaisuu.Top = 480 + (3600 - 480) * HyoujiBairitu
    lblbaisuu.FontSize = 10 * HyoujiBairitu
    lblbaisuu.Width = 855 * HyoujiBairitu
    lblbaisuu.Height = txtbaisuu.Height
'
    frmkinyuu.Left = 480 + (4920 - 480) * HyoujiBairitu
    frmkinyuu.Top = 480
    frmkinyuu.FontSize = 10 * HyoujiBairitu
    frmkinyuu.Width = 1815 * HyoujiBairitu
    frmkinyuu.Height = 855 * HyoujiBairitu
'
    optdirect.Left = 120 * HyoujiBairitu
    optdirect.Top = 240 * HyoujiBairitu
    optdirect.FontSize = 10 * HyoujiBairitu
    optdirect.Width = 1455 * HyoujiBairitu
    optdirect.Height = 255 * HyoujiBairitu
'
    opthyou.Left = 120 * HyoujiBairitu
    opthyou.Top = 480 * HyoujiBairitu
    opthyou.FontSize = 10 * HyoujiBairitu
    opthyou.Width = 1455 * HyoujiBairitu
    opthyou.Height = 255 * HyoujiBairitu
'
    frmSougou.Left = 480 + (4920 - 480) * HyoujiBairitu
    frmSougou.Top = 480 + (1560 - 480) * HyoujiBairitu
    frmSougou.FontSize = 10 * HyoujiBairitu
    frmSougou.Width = 1815 * HyoujiBairitu
    frmSougou.Height = 855 * HyoujiBairitu
'
    optSougou.Left = 120 * HyoujiBairitu
    optSougou.Top = 240 * HyoujiBairitu
    optSougou.FontSize = 10 * HyoujiBairitu
    optSougou.Width = 1455 * HyoujiBairitu
    optSougou.Height = 255 * HyoujiBairitu
'
    optKobetu.Left = 120 * HyoujiBairitu
    optKobetu.Top = 480 * HyoujiBairitu
    optKobetu.FontSize = 10 * HyoujiBairitu
    optKobetu.Width = 1455 * HyoujiBairitu
    optKobetu.Height = 255 * HyoujiBairitu
'
    frmSort.Left = 480 + (4920 - 480) * HyoujiBairitu
    frmSort.Top = 480 + (2640 - 480) * HyoujiBairitu
    frmSort.FontSize = 10 * HyoujiBairitu
    frmSort.Width = 1815 * HyoujiBairitu
    frmSort.Height = 1095 * HyoujiBairitu
'
    optCode.Left = 120 * HyoujiBairitu
    optCode.Top = 240 * HyoujiBairitu
    optCode.FontSize = 10 * HyoujiBairitu
    optCode.Width = 1455 * HyoujiBairitu
    optCode.Height = 255 * HyoujiBairitu
'
    optName.Left = 120 * HyoujiBairitu
    optName.Top = 480 * HyoujiBairitu
    optName.FontSize = 10 * HyoujiBairitu
    optName.Width = 1455 * HyoujiBairitu
    optName.Height = 255 * HyoujiBairitu
'
    optMaker.Left = 120 * HyoujiBairitu
    optMaker.Top = 720 * HyoujiBairitu
    optMaker.FontSize = 10 * HyoujiBairitu
    optMaker.Width = 1455 * HyoujiBairitu
    optMaker.Height = 255 * HyoujiBairitu
'
    frmSoutou.Left = 480 + (2760 - 480) * HyoujiBairitu
    frmSoutou.Top = 480 + (3960 - 480) * HyoujiBairitu
    frmSoutou.FontSize = 10 * HyoujiBairitu
    frmSoutou.Width = 1815 * HyoujiBairitu
    frmSoutou.Height = 1095 * HyoujiBairitu
'
    lblTyu1.Left = 100 * HyoujiBairitu
    lblTyu1.Top = 240 * HyoujiBairitu
    lblTyu1.FontSize = 10 * HyoujiBairitu
    lblTyu1.Width = 1615 * HyoujiBairitu
    lblTyu1.Height = 255 * HyoujiBairitu
'
    optBuhinmei.Left = 120 * HyoujiBairitu
    optBuhinmei.Top = 480 * HyoujiBairitu
    optBuhinmei.FontSize = 10 * HyoujiBairitu
    optBuhinmei.Width = 1455 * HyoujiBairitu
    optBuhinmei.Height = 255 * HyoujiBairitu
'
    optBikourann.Left = 120 * HyoujiBairitu
    optBikourann.Top = 720 * HyoujiBairitu
    optBikourann.FontSize = 10 * HyoujiBairitu
    optBikourann.Width = 1455 * HyoujiBairitu
    optBikourann.Height = 255 * HyoujiBairitu
'
    frmTanka.Left = 480 + (4920 - 480) * HyoujiBairitu
    frmTanka.Top = 480 + (3960 - 480) * HyoujiBairitu
    frmTanka.FontSize = 10 * HyoujiBairitu
    frmTanka.Width = 1815 * HyoujiBairitu
    frmTanka.Height = 1095 * HyoujiBairitu
'
    optBikou.Left = 120 * HyoujiBairitu
    optBikou.Top = 240 * HyoujiBairitu
    optBikou.FontSize = 10 * HyoujiBairitu
    optBikou.Width = 1215 * HyoujiBairitu
    optBikou.Height = 255 * HyoujiBairitu
'
    optTanka.Left = 120 * HyoujiBairitu
    optTanka.Top = 480 * HyoujiBairitu
    optTanka.FontSize = 10 * HyoujiBairitu
    optTanka.Width = 1215 * HyoujiBairitu
    optTanka.Height = 255 * HyoujiBairitu
'
    optTainetu.Left = 120 * HyoujiBairitu
    optTainetu.Top = 720 * HyoujiBairitu
    optTainetu.FontSize = 10 * HyoujiBairitu
    optTainetu.Width = 1215 * HyoujiBairitu
    optTainetu.Height = 255 * HyoujiBairitu
'
    cmdGo.Left = 480 + (1200 - 480) * HyoujiBairitu
    cmdGo.Top = 480 + (5400 - 480) * HyoujiBairitu
    cmdGo.FontSize = 10 * HyoujiBairitu
    cmdGo.Width = 1335 * HyoujiBairitu
    cmdGo.Height = 495 * HyoujiBairitu
'
    cmdFile.Left = 480 + (3000 - 480) * HyoujiBairitu
    cmdFile.Top = 480 + (5400 - 480) * HyoujiBairitu
    cmdFile.FontSize = 10 * HyoujiBairitu
    cmdFile.Width = 1335 * HyoujiBairitu
    cmdFile.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 480 + (4920 - 480) * HyoujiBairitu
    cmdQuit.Top = 480 + (5400 - 480) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1335 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
End Sub


