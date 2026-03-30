VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Plst_PRNstd 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "標準部品表印刷"
   ClientHeight    =   4320
   ClientLeft      =   2475
   ClientTop       =   1200
   ClientWidth     =   6750
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
   Icon            =   "Plst_PRNstd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4320
   ScaleWidth      =   6750
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   600
   End
   Begin VB.TextBox txtPrinting 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0000C000&
      Height          =   285
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "Plst_PRNstd.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frmTanka 
      BackColor       =   &H00008000&
      Caption         =   "備考欄 選択"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4800
      TabIndex        =   20
      Top             =   1440
      Width           =   1455
      Begin VB.OptionButton optBikou 
         BackColor       =   &H00008000&
         Caption         =   "   備考"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTanka 
         BackColor       =   &H00008000&
         Caption         =   "部品単価"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txtbaisuu 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      MousePointer    =   1  '矢印
      TabIndex        =   17
      Text            =   "1"
      Top             =   3360
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
      Caption         =   "ﾀｲﾄﾙ記入"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton opthyou 
         BackColor       =   &H00008000&
         Caption         =   "構 成 表"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optdirect 
         BackColor       =   &H00008000&
         Caption         =   "直接入力"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
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
      Left            =   4800
      TabIndex        =   18
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   4800
      TabIndex        =   19
      Top             =   3480
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   10
      Orientation     =   2
   End
   Begin VB.Label lblbaisuu 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "倍数"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lbldate 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "日付 ： 1997/09/19"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Top             =   3360
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
      Top             =   3000
      Width           =   3855
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
      Top             =   2640
      Width           =   3375
   End
End
Attribute VB_Name = "Plst_PRNstd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 標準部品表印刷 ***
'**********************
'
Option Explicit
'
Dim FLGkouseihyou As Integer, FLGgyou As Integer, FLGpage As Integer
Dim FLGtuzuki As Integer, FLGtuzukip As Integer
Dim FLGend As Integer, FLGheader As Integer
Dim FLGitemp As Integer
Dim kijunX As Integer, kijunY As Integer
Dim haba1X As Integer, haba2X As Integer, haba3X As Integer, haba31X As Integer
Dim haba4X As Integer, haba5Xa As Integer, haba5Xb As Integer, haba5X As Integer
Dim haba6X As Integer, haba7X As Integer, gyoukan As Integer
Dim moji_zureX As Integer, moji_zureY As Integer
Dim NowpointX As Integer, NowpointY As Integer
Dim Tmpdata As String
Dim Tanka_total As Double
Dim BuhinBangouMojisuu As Integer
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
Dim Pdata_code() As String      '*** Pdata_code(FLGtuzuki) ***
Dim Pdata_kikaku() As String    '*** Pdata_kikaku(FLGtuzuki) ***
Dim Pdata_ghyouji() As String   '*** Pdata_ghyouji(FLGtuzuki) ***
Dim Pdata_bangou() As String    '*** Pdata_bangou(FLGtuzuki) ***
Dim Pdata_kosuu() As String     '*** Pdata_kosuu(FLGtuzuki) ***
Dim Pdata_sousuu() As String    '*** Pdata_sousuu(FLGtuzuki) ***
Dim Pdata_maker() As String     '*** Pdata_maker(FLGtuzuki) ***
Dim Pdata_bikou() As String     '*** Pdata_bikou(FLGtuzuki) ***
Dim Pdata_tanka() As Double     '*** Pdata_tanka(FLGtuzuki) ***
Dim Pdata_RoHS() As String      '*** Pdata_RoHS(FLGtuzuki) ***
'
Dim Pdata_codep As String       '*** 部品ｺｰﾄﾞ ***
Dim Pdata_kikakup As String     '*** 部品規格 ***
Dim Pdata_ghyoujip As String    '*** 現品表示 ***
Dim Pdata_bangoup As String     '*** 部品番号 ***
Dim Pdata_kosuup As String      '*** 個数 ***
Dim Pdata_sousuup As String     '*** 総数 ***
Dim Pdata_makerp As String      '*** ﾒｰｶｰ ***
Dim Pdata_bikoup As String      '*** 備考欄 ***
Dim Pdata_RoHSp As String       '*** ﾒｯｷ/RoHS ***

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (6840 - 960) * HyoujiBairitu + 480
    Height = 480 + (4800 - 960) * HyoujiBairitu + 480
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
    Plst_PRNstd.Caption = STATUS
'
    Call S_Hyouji            '*** 画面初期化＆表示 ***
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = vbFormControlMenu Then
'
'    End If
End Sub

Private Sub PRNhamidashi()
    Printer.Print "はみ出し部品  " & PLSTP(pp, 0) & "  " & PLSTP(pp, 1) & "  " & PLSTP(pp, 2)
End Sub

Private Sub PRNIndex()
'                   *** 項目名印刷 ***
    Dim itiji As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)   '*** フォントサイズ設定 ***
    itiji = str(FLGgyou - 3)
    If Len(itiji) = 3 Then
        itiji = Mid(itiji, 2)
    End If
    Printer.Print itiji
'
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print " " & Aitem0(ipT, 1)
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba31X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xa, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xb, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X + haba6X
    NowpointY = kijunY + gyoukan * FLGgyou
'    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
'
'    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5Xa * 2 + haba5X + haba6X * 2
'    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba7X, gyoukan), 0, B
End Sub

Private Sub PRNfooter()
'                   *** フッター印刷 ***
    Dim Pdata As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.CurrentX = NowpointX + 284
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(9, 1)
    Printer.Print "M5904-02";
'
    Printer.CurrentX = NowpointX + 7088
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 0)
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷";
'
    Printer.CurrentX = NowpointX + 15026
    Printer.CurrentY = NowpointY + moji_zureY
    If FLGend = 0 Then
        Printer.Print "つづく"
    Else
        If optTanka.Value = True Then
            Printer.CurrentX = NowpointX + 13000
            Call SET_Yen0_Format(Tanka_total, Pdata, 11)
            Printer.Print "おわり    合計："; Pdata; "*"
        Else
            Printer.Print "おわり"
        End If
    End If
End Sub

Function LenMbcs(ByVal str As String)           '*** 漢字混じり文字列を正確に数えるおまじない ***
   LenMbcs = LenB(StrConv(str, vbFromUnicode))
End Function

Private Sub PRNkoumoku(FLGtuzukip As Integer)
'                   *** 項目印刷 ***
    Dim itiji As String
    Dim Pdata As String
    Dim Pdata1 As String
    Dim mojityou As Integer
    Dim tempStrA As String, tempStrB As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)   '*** フォントサイズ設定 ***
    itiji = str(FLGgyou - 3)
    If Len(itiji) = 3 Then
        itiji = Mid(itiji, 2)
    End If
    Printer.Print itiji
'
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    If Pdata_code(FLGtuzukip) <> "" Then
        Printer.CurrentX = NowpointX + moji_zureX
        Printer.CurrentY = NowpointY + moji_zureY
        Printer.Print "  " & Pdata_code(FLGtuzukip)
    End If
'
    If Pdata_kikaku(FLGtuzukip) <> "" Then
        NowpointX = kijunX + haba1X + haba2X
        NowpointY = kijunY + gyoukan * FLGgyou
        Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
        Printer.CurrentX = NowpointX + moji_zureX
        Printer.CurrentY = NowpointY + moji_zureY
'
        If 32 < LenMbcs(Pdata_kikaku(FLGtuzukip)) Then
            Call SETfont_size(8, 1)    '*** フォントサイズ設定 ***
        ElseIf 26 < LenMbcs(Pdata_kikaku(FLGtuzukip)) Then
            Call SETfont_size(9, 1)    '*** フォントサイズ設定 ***
        End If
        Printer.Print Pdata_kikaku(FLGtuzukip)
'
        Call SETfont_size(10.8, 1)   '*** フォントサイズ設定 ***
    End If
'
    If Pdata_ghyouji(FLGtuzukip) <> "" Then
        NowpointX = kijunX + haba1X + haba2X + haba3X
        NowpointY = kijunY + gyoukan * FLGgyou
        Printer.Line (NowpointX, NowpointY)-Step(haba31X, gyoukan), 0, B
'
        If 23 < LenMbcs(Pdata_ghyouji(FLGtuzukip)) Then
            tempStrA = Left(Pdata_ghyouji(FLGtuzukip), 25)
            tempStrB = Mid(Pdata_ghyouji(FLGtuzukip), 26)
'
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY / 5
'
            Call SETfont_size(8, 1)    '*** フォントサイズ設定 ***
            Printer.Print tempStrA
'
            Printer.CurrentX = NowpointX + moji_zureX * 2
            Printer.CurrentY = NowpointY + moji_zureY * 2 + moji_zureY * 4 / 5
            Printer.Print tempStrB
        ElseIf 18 < LenMbcs(Pdata_ghyouji(FLGtuzukip)) Then
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY
'
            Call SETfont_size(9, 1)    '*** フォントサイズ設定 ***
            Printer.Print Pdata_ghyouji(FLGtuzukip)
        Else
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY
'
            Printer.Print Pdata_ghyouji(FLGtuzukip)
        End If
'
        Call SETfont_size(10.8, 1)   '*** フォントサイズ設定 ***
'
        NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X
        NowpointY = kijunY + gyoukan * FLGgyou
        Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
'        If Pdata_bangou(FLGtuzukip) <> "" Then
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY
            Printer.Print Pdata_bangou(FLGtuzukip)
'        End If
    Else
        NowpointX = kijunX + haba1X + haba2X
        NowpointY = kijunY + gyoukan * FLGgyou
        Printer.Line (NowpointX, NowpointY)-Step(haba3X / 2 - 56, gyoukan), 0, B
'
        NowpointX = kijunX + haba1X + haba2X + haba3X / 2 - 56
        NowpointY = kijunY + gyoukan * FLGgyou
        Printer.Line (NowpointX, NowpointY)-Step(haba3X / 2 + 56 + haba31X + haba4X, gyoukan), 0, B
        If Pdata_bangou(FLGtuzukip) <> "" Then
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY
'
            mojityou = Len(Pdata_bangou(FLGtuzukip))
            If mojityou < BuhinBangouMojisuu Then
                Printer.Print Spc(BuhinBangouMojisuu - mojityou); Pdata_bangou(FLGtuzukip)
            Else
                Printer.Print Pdata_bangou(FLGtuzukip)
            End If
        End If
    End If
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xa, gyoukan), 0, B
    If Pdata_kosuu(FLGtuzukip) <> "" Then
        Printer.CurrentX = NowpointX + moji_zureX - 80
        Printer.CurrentY = NowpointY + moji_zureY
            SET_migiyose Trim(Pdata_kosuu(FLGtuzukip)), Pdata, 5  '*** 5文字にする
        Printer.Print Pdata
    End If
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xb, gyoukan), 0, B
    If Pdata_sousuu(FLGtuzukip) <> "" Then
        Printer.CurrentX = NowpointX + moji_zureX - 50
        Printer.CurrentY = NowpointY + moji_zureY
            SET_migiyose Trim(Pdata_sousuu(FLGtuzukip)), Pdata, 5 '*** 5文字にする
        Printer.Print Pdata
    End If
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    If Pdata_maker(FLGtuzukip) <> "" Then
        Printer.CurrentX = NowpointX + moji_zureX - 30
        Printer.CurrentY = NowpointY + moji_zureY
            SET_tyuuou Pdata_maker(FLGtuzukip), Pdata, 6  '*** 6文字にする
        Printer.Print Pdata
    End If
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    If Pdata_RoHS(FLGtuzukip) <> "" Then
        Printer.CurrentX = NowpointX + moji_zureX - 80
        Printer.CurrentY = NowpointY + moji_zureY
        Printer.Print Pdata_RoHS(FLGtuzukip)
    End If
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X + haba6X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba7X, gyoukan), 0, B
'
    If optBikou.Value = True Then
        If Pdata_bikou(FLGtuzukip) <> "*" Then
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY
            Printer.Print Pdata_bikou(FLGtuzukip)
        End If
    Else    'If optTanka.Value = True Then
        If 0 < Pdata_tanka(FLGtuzukip) Then
            Printer.CurrentX = NowpointX + moji_zureX
            Printer.CurrentY = NowpointY + moji_zureY
            Call SET_Yen1_Format(Pdata_tanka(FLGtuzukip), Pdata, 9)
            Call SET_Yen0_Format(Pdata_tanka(FLGtuzukip) * Val(Pdata_sousuu(FLGtuzukip)), Pdata1, 11)
            Printer.Print Pdata; Pdata1
            Tanka_total = Tanka_total + Pdata_tanka(FLGtuzukip) * Val(Pdata_sousuu(FLGtuzukip))
        End If
    End If
End Sub

Private Sub PRNheader1()
'                   *** 部品表項目見出し印刷 ***
'                   567twip=10mm,1440twip=1inch,<16114>
    FLGgyou = 3
    haba1X = 425
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)   '*** フォントサイズ設定 ***
    Printer.Print "行"
'
    haba2X = 1786
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "品名／ｺｰﾄﾞ番号"
'
    haba3X = 3033
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 425 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "品番・形名・製品名"
'
    haba31X = 2208      '2448 -240
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba31X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 558 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "現品表示"
'
    haba4X = 1844       '2776
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 206 + moji_zureX     '672
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "部 品 番 号"
'
    haba5Xa = 640
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xa, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 10 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "個数"
'
    haba5Xb = 700
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xb, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 23 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "総数"
'
    haba5X = 970
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 140 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "ﾒｰｶｰ"
'
    haba6X = 2144   '1804 +240 +100
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
'    Printer.CurrentX = NowpointX + 154 + moji_zureX
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "MSL/耐熱/ﾒｯｷ/RoHS"
'
    haba7X = 2366
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X + haba6X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba7X, gyoukan), 0, B
    
    If optBikou.Value = True Then
        Printer.CurrentX = NowpointX + moji_zureX
        Printer.CurrentY = NowpointY + moji_zureY
        Printer.Print "        備考"
    Else    'If optTanka.Value = True Then
        Printer.CurrentX = NowpointX + moji_zureX
        Printer.CurrentY = NowpointY + moji_zureY
        Printer.Print " 平均単価    ｘ総数"
    End If
End Sub

Private Sub PRNheader0()
'                   *** 部品表ヘッダー印刷 ***
    Dim i As Integer
'                   567twip=10mm,1440twip=1inch
    kijunX = 56     '113
    kijunY = 1020
    gyoukan = 340
    moji_zureX = 113
    moji_zureY = 60
    FLGgyou = 0
    FLGheader = 1   '*** ヘッダー印刷済みにセット
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    Printer.Orientation = vbPRORLandscape   '*** ﾗﾝﾄﾞｽｹｰﾌﾟ <16114>***
'
    Printer.CurrentX = kijunX + 6010
    Printer.CurrentY = 955
    Call SETfont_size(17, 1)       '*** フォント,サイズ設定 ***
    Printer.Print "《　電  気  部　品　表　》"
'
    Printer.CurrentX = kijunX + 13600
    Printer.CurrentY = 1090
    Call SETfont_size(10.8, 1)     '*** フォント,サイズ設定 ***
'
    Printer.Print PlistdateP '*** 日付印刷 ***
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
'
    If FLGpage = 1 Then
        Printer.Line (kijunX, kijunY)-Step(0, gyoukan)
        Call SETfont_size(10.8, 1)     '*** フォントサイズ設定 ***
        Printer.CurrentX = kijunX + moji_zureX
        Printer.CurrentY = kijunY + moji_zureY
        Printer.Print "検査："
    End If
'
    FLGgyou = 1
    haba1X = 2437
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Call SETfont_size(10.8, 1)   '*** フォントサイズ設定 ***
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "型式： " & CATnoT
'
    haba2X = 4366
    NowpointX = kijunX + haba1X
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "名称： " & CATnameT
'
    haba3X = 2325
    NowpointX = kijunX + haba1X + haba2X
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "担当者： " & PersonT
'
    haba4X = 4097
    NowpointX = kijunX + haba1X + haba2X + haba3X
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "工番： " & KoubanT
'
    haba5X = 2889
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "台数： " & Trim(DaisuuT)
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + 1360
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "x倍数：" & Trim(BaisuuT)
'
    haba1X = 2437
    FLGgyou = 2
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
'
    Call remove_PLT(PFLnameP, Tmpdata)
'
    Printer.Print "図番： " & Tmpdata
'
    haba2X = 3646
    NowpointX = kijunX + haba1X
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "小名称： " & PlistnameP
'
    haba3X = 8616
    NowpointX = kijunX + haba1X + haba2X
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    If RemarksP = "*" Then
        RemarksP = ""
    End If
    Printer.Print "記事： ";
    Printer.CurrentY = NowpointY + moji_zureY * 1.3
    Call SETfont_size(9, 2)     '*** フォント,サイズ設定 ***
    Printer.Print RemarksP
'
    haba4X = 1415
    NowpointX = kijunX + haba1X + haba2X + haba3X
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)     '*** フォント,サイズ設定 ***
    Printer.Print "ﾍﾟｰｼﾞ： " & FLGpage
    FLGpage = FLGpage + 1
End Sub

Private Sub Klst_yomu()
    DRVconstT = TMPdir1 & "\constlst.cod"
    Call RDconst_lst(DRVconstT, CATnoT, CATnameT, ZubanT, PersonT, OrgdateT, RevdateT, CheckdateT, OutdateT, _
                KLSTT(), KtotalT, KdimT, KoubanT, DaisuuT, KbikouT, KyobiAT, KyobiBT)   '*** 構成表読み込み ***
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
    lblnamae.Caption = "印刷 ﾌｧｲﾙ名 ： " & PFLnameP
'
    Call RDpartlist(DRVpartlistP, PlistnameP, PlistdateP, RemarksP, PLSTP(), PtotalP, PdimP)    '*** 部品表読み込み ***
'
    If FLGall = 0 Then
        lblkomei.Caption = "小名称 ： " & PlistnameP
        lblDate.Caption = "日付 ： " & PlistdateP
        txtbaisuu.Text = BaisuuT
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
    Dim itiji As String
    Dim Pdata As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    SETfont_size 10.8, 1  '*** フォントサイズ設定 ***
    itiji = str(FLGgyou - 3)
    If Len(itiji) = 3 Then
        itiji = Mid(itiji, 2)
    End If
    Printer.Print itiji
'
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "  未登録部品"
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Mid(PLSTP(pp, 1), 2)
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba31X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print PLSTP(pp, 0)
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xa, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX - 50
    Printer.CurrentY = NowpointY + moji_zureY
        Call SET_migiyose("1", Pdata, 5)   '*** 5文字にする
    Printer.Print Pdata
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5Xb, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX - 90
    Printer.CurrentY = NowpointY + moji_zureY
        Call SET_migiyose(Trim(str(Val(BaisuuT) * Val(DaisuuT))), Pdata, 5)  '*** 5文字にする
    Printer.Print Pdata
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba31X + haba4X + haba5Xa + haba5Xb + haba5X + haba6X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba7X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
'
    If optBikou.Value = True Then
        Printer.Print PLSTP(pp, 2)
    End If
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer, j As Integer, np As Integer, q As Integer
    Dim Ckdim As Integer, Gyoumax As Integer
    Dim Bmoji As Integer, Cmoji As Integer, Xmoji As Integer, Ymoji As Integer
    Dim TMPmoji As String
    Dim MKsitei As String
    Dim Tmp1 As String
    Dim FLGitti As Integer, FLGmojime As Integer
    Dim FLGbanme As Integer
    Dim objPrinter As Printer
'
    If opthyou.Value = True Then
        Call Klst_yomu              '*** 最新の構成表を読み込み内容表示 ***
    End If
'
    CATnoT = Trim(txtkeisiki.Text)
    CATnameT = Trim(txtmeisyou.Text)
    PersonT = Trim(txttantou.Text)
    KoubanT = Trim(txtkouban.Text)
    DaisuuT = Trim(txtdaisuu.Text)
    BaisuuT = Trim(txtbaisuu.Text)
'
    If FLGall = 1 Then      '*** ファイルの有無チェック ***
        For FLGbanme = 1 To KtotalT
            If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
                PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
                DRVpartlistP = TMPplst & "\" & PFLnameP            '*** .PLT を検索 ***
                Tmp1 = Dir(DRVpartlistP)
                If Tmp1 = "" Then
                    PFLnameP = KLSTT(FLGbanme, 2)
                    If Len(PFLnameP) > 8 Then
                        PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                    End If
'
                    DRVpartlistP = TMPplst & "\" & PFLnameP        '*** xxxxxxxx.yyy を検索 ***
                    Tmp1 = Dir(DRVpartlistP)
'                    If Tmp1 = "" Then
'                        PFLnamep = KLSTT(FLGbanme, 2) & ".PLT"
'                        DRVpartlistp = TMPdir2 & "\" & PFLnamep    '*** 共有フォルダーを検索 ***
'                        Tmp1 = Dir(DRVpartlistp)
                        If Tmp1 = "" Then
                            i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                            Exit Sub
'
                        End If
'                    End If
                End If
            End If
'
        Next FLGbanme
    End If
'
    flag_cancel = False
    Printer_Window.Show 1
        If flag_cancel = True Then GoTo OWARI
'
    For Each objPrinter In Printers
        If objPrinter.DeviceName = strMyPrinter Then
            Set Printer = objPrinter    'オブジェクトに代入
        End If
    Next
'
    Plst_PRNstd.MousePointer = vbHourglass   '*** 砂時計 ***
    DoEvents
'
    FLGbanme = 1
    If FLGall = 1 Then
KURIKAESHI:
        If FLGbanme > KtotalT Then GoTo OWARI    '*** 終わりにｼﾞｬﾝﾌﾟ ***
'
        If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
            PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP            '*** .PLT を検索 ***
'
            Tmp1 = Dir(DRVpartlistP)
            If Tmp1 = "" Then
                PFLnameP = KLSTT(FLGbanme, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
                DRVpartlistP = TMPplst & "\" & PFLnameP        '*** xxxxxxxx.yyy を検索 ***
'
'                Tmp1 = Dir(DRVpartlistp)
'                If Tmp1 = "" Then
'                    PFLnamep = KLSTT(FLGbanme, 2) & ".PLT"
'                    DRVpartlistp = TMPdir2 & "\" & PFLnamep    '*** 共有フォルダーを検索 ***
'                End If
'
            End If
            BaisuuT = KLSTT(FLGbanme, 4) '*** 倍数設定 ***
            FLGbanme = FLGbanme + 1
        Else
            FLGbanme = FLGbanme + 1
            GoTo KURIKAESHI     '*** 次項目へ ***
        End If
'
        Call Plst_yomu           '*** 部品表を読む ***
'
    End If
'
    txtPrinting.Text = vbCrLf & " !!! 部品表 " & PFLnameP & " をプリントバッファーに転送中 !!!"
    txtPrinting.Visible = True
    DoEvents                '*** 画面書き直し ***
'
    FLGpage = 1             '*** ページ初期化
    FLGheader = 0           '*** ヘッダー印刷初期化
    Gyoumax = 28            '*** 印刷行最大値 ***
    Cmoji = 20              '*** 備考欄印刷文字数 ***
    Tanka_total = 0         '*** 単価合計 ***
'
    Ckdim = 100  '*** ﾃﾞｨﾒﾝｼﾞｮﾝ数
    ReDim Pdata_code(Ckdim)
    ReDim Pdata_kikaku(Ckdim)
    ReDim Pdata_ghyouji(Ckdim)
    ReDim Pdata_bangou(Ckdim)
    ReDim Pdata_kosuu(Ckdim)
    ReDim Pdata_sousuu(Ckdim)
    ReDim Pdata_maker(Ckdim)
    ReDim Pdata_bikou(Ckdim)
    ReDim Pdata_tanka(Ckdim)
    ReDim Pdata_RoHS(Ckdim)
    pps = 1
'
    For ipT = 1 To Anum0
        DRVindexT = Xcont0(2) & "\" & Aitem0(ipT, 0) & "\" & Aitem0(ipT, 0) & "INDEX.COD"
'
        Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)      '*** INDEX.COD 読み込み ***
        FLGitemp = 0    '*** 項目印刷ﾌﾗｸﾞｸﾘｱｰ
        jpsT = pps
'
        For jpT = 1 To BnumT
            For pp = jpsT To PtotalP
                If ((Left(PLSTP(pp, 1), 1) = "L") And (Mid(PLSTP(pp, 1), 2, 4) = BindexT(jpT, 0))) Then
                    pps = pp
'
                    If FLGheader = 0 Then
                        Call PRNheader0  '*** 部品表ヘッダー印刷 ***
                        Call PRNheader1  '*** 部品表項目見出し印刷 ***
                    End If
'
                    If FLGitemp = 0 Then
                        If FLGgyou >= Gyoumax - 1 Then
                            FLGend = 0
                            Call PRNfooter
                            Printer.NewPage   '*** 改ページ ***
'
                            FLGheader = 0
                            Call PRNheader0
                            Call PRNheader1
                        End If
'
                        Call PRNIndex        '*** 項目名印刷 ***
                        FLGitemp = 1    '*** 項目印刷済みにｾｯﾄ ***
                    End If
'
                    Call SET_DRVmain(DRVmainT, Aitem0(), ipT, BindexT(), jpT)
'
                    Call RDmain(DRVmainT, CmainT(), CnumT, CdimT) '*** MAIN.COD 読み込み ***
'
                    FLGtuzuki = 1       '*** ﾃﾞｰﾀｰｸﾘｱｰのため
                    For kpT = 1 To CnumT
                        If FLGtuzuki > 0 Then
                            For i = 0 To Ckdim
                                Pdata_code(i) = ""
                                Pdata_kikaku(i) = ""
                                Pdata_bangou(i) = ""
                                Pdata_ghyouji(i) = ""
                                Pdata_kosuu(i) = ""
                                Pdata_sousuu(i) = ""
                                Pdata_maker(i) = ""
                                Pdata_bikou(i) = ""
                                Pdata_tanka(i) = 0
                                Pdata_RoHS(i) = ""
                            Next i
                        End If
'
                        Pdata_codep = "L" & BindexT(jpT, 0) & "-" & CmainT(kpT, 0)
                        FLGtuzuki = 0       '*** 項目があれば１～
'
                        For np = pps To PtotalP
                            If PLSTP(np, 1) = Pdata_codep Then   '*** 個別にﾃﾞｰﾀｰ作成
                                If BindexT(jpT, 5) = "000" Then
                                    Pdata_makerp = CmainT(kpT, 13)
'
                                ElseIf BindexT(jpT, 5) = "998" Then
                                    If PLSTP(np, 3) = "0" Then
                                        MKsitei = "0"
                                    Else
                                        MKsitei = PLSTP(np, 3)
                                    End If
'
                                    Call GET998maker(Pdata_makerp, MKsitei, BindexT(), jpT)
'
                                Else
                                    Pdata_makerp = BindexT(jpT, 5)
                                End If
'
                                Call Makerget2(Pdata_makerp)   '*** メーカー略称取得 ***
'
                                Call GETkikaku(Pdata_kikakup, MKsitei, BindexT(), jpT, CmainT(), kpT) '*** 部品名取得 ***
'
                                Call GET_shitei2(CmainT(kpT, 3), Pdata_kikakup, q)   '*** ! ? の記入 ***
'
                                If CmainT(kpT, 16) = "1" Then         '*** 特記事項記入
                                    If PLSTP(np, 4) = "" Or PLSTP(np, 4) = "*" Then
                                        '
                                    Else
                                        Pdata_kikakup = Pdata_kikakup & PLSTP(np, 4)
                                    End If
                                End If
'
                                Pdata_ghyoujip = CmainT(kpT, 9)
                                If Pdata_ghyoujip = "" Or Pdata_ghyoujip = "0" Then
                                    Pdata_ghyoujip = "-"
                                End If
'
                                Pdata_bikoup = PLSTP(np, 2)
'
                                FLGitti = 0     '*** 一致ﾌﾗｸﾞｸﾘｱｰ
                                For i = 0 To FLGtuzuki              '*** 一致する項目を集計する
                                    If Pdata_kikaku(i) = Pdata_kikakup And _
                                                            Pdata_bikou(i) = Pdata_bikoup Then
'                                                                   '*** 部品番号集計 ***
                                        Pdata_bangou(i) = Pdata_bangou(i) & "," & PLSTP(np, 0)
'                                                                   '*** 個数合計 ***
                                        Pdata_kosuu(i) = str(Val(Pdata_kosuu(i)) + 1)
                                        FLGitti = 1     '*** 一致ﾌﾗｸﾞｾｯﾄ
                                        Exit For
'
                                    End If
                                Next i
'
                                If FLGitti = 0 Then
                                    Pdata_code(FLGtuzuki) = Pdata_codep
                                    Tmp1 = CmainT(kpT, 18)
                                        If Tmp1 = "6" Or Tmp1 = "7" Or Tmp1 = "8" Or Tmp1 = "9" _
                                            Or Tmp1 = "A" Or Tmp1 = "B" Or Tmp1 = "C" Or Tmp1 = "D" _
                                            Or Tmp1 = "E" Or Tmp1 = "F" Then            '*** SOP表示 ***
                                            Pdata_code(FLGtuzuki) = Pdata_code(FLGtuzuki) + " sop"
                                        End If
                                    Pdata_kikaku(FLGtuzuki) = Pdata_kikakup
                                    Pdata_bangou(FLGtuzuki) = PLSTP(np, 0)
                                    Pdata_ghyouji(FLGtuzuki) = Pdata_ghyoujip
                                    Pdata_kosuu(FLGtuzuki) = "1"
                                    Pdata_maker(FLGtuzuki) = Pdata_makerp
'
                                    TMPmoji = CmainT(kpT, 6)    '*** MSL記入
                                    Call TRS_Mlevel2(TMPmoji)
                                    Tmp1 = TMPmoji & "/"
'
                                    TMPmoji = CmainT(kpT, 11)       '*** 耐熱記入
                                    If Len(TMPmoji) < 3 Then TMPmoji = "   "  '***3文字未満は無い
                                    Tmp1 = Tmp1 & TMPmoji & "/"
'
                                    If 5 < Len(CmainT(kpT, 19)) Then    '*** ﾒｯｷ記入
                                        Tmp1 = Tmp1 & CmainT(kpT, 19) & "/"
                                    ElseIf 5 = Len(CmainT(kpT, 19)) Then
                                        Tmp1 = Tmp1 & " " & CmainT(kpT, 19) & "/"
                                    ElseIf 4 = Len(CmainT(kpT, 19)) Then
                                        Tmp1 = Tmp1 & " " & CmainT(kpT, 19) & " /"
                                    ElseIf 3 = Len(CmainT(kpT, 19)) Then
                                        Tmp1 = Tmp1 & "  " & CmainT(kpT, 19) & " /"
                                    ElseIf 2 = Len(CmainT(kpT, 19)) Then
                                        Tmp1 = Tmp1 & "  " & CmainT(kpT, 19) & "  /"
                                    Else
                                        Tmp1 = Tmp1 & "      /"     '*** 1文字の金属メッキは無い ***
                                    End If
'
                                    If InStr(CmainT(kpT, 19), "SnPb") <> 0 Then '*** RoHS記入
                                        Tmp1 = Tmp1 & "----"
                                    ElseIf InStr(CmainT(kpT, 2), "#<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(CmainT(kpT, 2), "<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "RoHS"
                                    ElseIf InStr(CmainT(kpT, 2), "#<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(CmainT(kpT, 2), "<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "Pbﾌﾘ"
                                    ElseIf InStr(CmainT(kpT, 2), "<Green>") <> 0 Then
                                        Tmp1 = Tmp1 & "Gree"
                                    ElseIf InStr(CmainT(kpT, 2), "<Ro2>") <> 0 Then     '*** Ver2.1ﾆﾃ追加
                                        Tmp1 = Tmp1 & "Ro2"
                                    ElseIf InStr(CmainT(kpT, 2), "<R863>") <> 0 Then    '*** Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "R863"
                                    ElseIf InStr(BindexT(jpT, 1), "#<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(BindexT(jpT, 1), "<RoHS>") <> 0 Then
                                        Tmp1 = Tmp1 & "RoHS"
                                    ElseIf InStr(BindexT(jpT, 1), "#<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "=>>="
                                    ElseIf InStr(BindexT(jpT, 1), "<Pbﾌﾘｰ>") <> 0 Then
                                        Tmp1 = Tmp1 & "Pbﾌﾘ"
                                    ElseIf InStr(BindexT(jpT, 1), "<Green>") <> 0 Then
                                        Tmp1 = Tmp1 & "Gree"
                                    ElseIf InStr(BindexT(jpT, 1), "<Ro2>") <> 0 Then    '*** Ver2.1ﾆﾃ追加
                                        Tmp1 = Tmp1 & "Ro2"
                                    ElseIf InStr(BindexT(jpT, 1), "<R863>") <> 0 Then   '*** Ver2.3ﾆﾃ追加
                                        Tmp1 = Tmp1 & "R863"
                                    End If
'
                                    Pdata_RoHS(FLGtuzuki) = Tmp1
'
'                                    If optBikou.Value = True Then   '*** 常に使用 ***
                                        Pdata_bikou(FLGtuzuki) = Pdata_bikoup
'                                    End If
'
                                    If optTanka.Value = True Then
                                        Pdata_tanka(FLGtuzuki) = Val(CmainT(kpT, 5))
                                    End If
'
                                    FLGtuzuki = FLGtuzuki + 1
                                End If
                            End If
                        Next np
'
                        If 0 < FLGtuzuki Then
                            For i = 0 To FLGtuzuki - 1      '*** 総数計算 ***
                                Pdata_sousuu(i) = str(Val(Pdata_kosuu(i)) * Val(BaisuuT) * Val(DaisuuT))
                            Next i
'
                            Bmoji = 15      '*** 部品番号印刷文字数, 初期値, +37 *** 24,28
                            BuhinBangouMojisuu = 15 + 35 - 1    '***印刷ルーチンで使用
'
                            For i = 0 To Ckdim - 1      '*** 部品番号行分割 ***
                                Do While Bmoji < Len(Pdata_bangou(i)) And i <= Ckdim - 1
                                    FLGmojime = Bmoji
                                    Do While Mid(Pdata_bangou(i), FLGmojime, 1) <> ","
                                        FLGmojime = FLGmojime - 1
                                    Loop
                                    Tmp1 = Mid(Pdata_bangou(i), FLGmojime + 1)
                                    Pdata_bangou(i) = Left(Pdata_bangou(i), FLGmojime)
'
                                    For j = Ckdim - 1 To i + 1 Step -1      '*** 全体をずらす
                                        Pdata_code(j + 1) = Pdata_code(j)
                                        Pdata_kikaku(j + 1) = Pdata_kikaku(j)
                                        Pdata_ghyouji(j + 1) = Pdata_ghyouji(j)
                                        Pdata_bangou(j + 1) = Pdata_bangou(j)
                                        Pdata_kosuu(j + 1) = Pdata_kosuu(j)
                                        Pdata_sousuu(j + 1) = Pdata_sousuu(j)
                                        Pdata_maker(j + 1) = Pdata_maker(j)
                                        Pdata_bikou(j + 1) = Pdata_bikou(j)
                                        Pdata_tanka(j + 1) = Pdata_tanka(j)
                                        Pdata_RoHS(j + 1) = Pdata_RoHS(j)
                                    Next j
 '
                                    Pdata_code(i + 1) = ""
                                    Pdata_kikaku(i + 1) = ""
                                    Pdata_ghyouji(i + 1) = ""
                                    Pdata_bangou(i + 1) = Tmp1
                                    Pdata_kosuu(i + 1) = ""
                                    Pdata_sousuu(i + 1) = ""
                                    Pdata_maker(i + 1) = ""
                                    Pdata_bikou(i + 1) = ""
                                    Pdata_tanka(i + 1) = 0
                                    Pdata_RoHS(i + 1) = ""
                                    FLGtuzuki = FLGtuzuki + 1
                                    i = i + 1
'
                                    If Bmoji = 15 Then
                                        Bmoji = Bmoji + 35
                                    End If
                                Loop
'
                                If Len(Pdata_bangou(i)) <= Bmoji Then
                                    Bmoji = 15
                                End If
                            Next i
'
                            If optBikou.Value = True Then
                                For i = 0 To Ckdim - 1      '*** 備考欄行分割 *** =>実質 not use
                                    Do While Cmoji < LenB(StrConv(Pdata_bikou(i), _
                                                        vbFromUnicode)) And i <= Ckdim - 1
                                        For j = 1 To Cmoji
                                            TMPmoji = Left(Pdata_bikou(i), j)
                                            Xmoji = Len(TMPmoji)
                                            Ymoji = LenB(StrConv(TMPmoji, vbFromUnicode))   '***<注意>***
                                            If Cmoji <= Ymoji Then
                                                Exit For
                                            End If
                                        Next j
'
                                        Tmp1 = Mid(Pdata_bikou(i), Xmoji + 1)
                                        Pdata_bikou(i) = Left(Pdata_bikou(i), Xmoji)
'
                                        If Pdata_code(i + 1) = "" Then
                                            Pdata_bikou(i + 1) = Tmp1
                                        Else
                                            For j = Ckdim - 1 To i + 1 Step -1
                                                Pdata_code(j + 1) = Pdata_code(j)
                                                Pdata_kikaku(j + 1) = Pdata_kikaku(j)
                                                Pdata_ghyouji(j + 1) = Pdata_ghyouji(j)
                                                Pdata_bangou(j + 1) = Pdata_bangou(j)
                                                Pdata_kosuu(j + 1) = Pdata_kosuu(j)
                                                Pdata_sousuu(j + 1) = Pdata_sousuu(j)
                                                Pdata_maker(j + 1) = Pdata_maker(j)
                                                Pdata_bikou(j + 1) = Pdata_bikou(j)
                                                Pdata_RoHS(j + 1) = Pdata_RoHS(j)
                                            Next j
'
                                            Pdata_code(i + 1) = ""
                                            Pdata_kikaku(i + 1) = ""
                                            Pdata_ghyouji(i + 1) = ""
                                            Pdata_bangou(i + 1) = ""
                                            Pdata_kosuu(i + 1) = ""
                                            Pdata_sousuu(i + 1) = ""
                                            Pdata_maker(i + 1) = ""
                                            Pdata_bikou(i + 1) = Tmp1
                                            Pdata_RoHS(i + 1) = ""
                                        End If
'
                                        FLGtuzuki = FLGtuzuki + 1
                                        i = i + 1
                                    Loop
                                Next i
                            End If
'
                            For i = 0 To FLGtuzuki - 1
                                If FLGgyou >= Gyoumax Then
                                        FLGend = 0
                                        Call PRNfooter
                                        Printer.NewPage '*** 改ページ ***
'
                                        FLGheader = 0
                                        Call PRNheader0
                                        Call PRNheader1
                                End If
'
                                Call PRNkoumoku(i)     '*** 部品表項目印刷 ***
                            Next i
                        End If
                    Next kpT
                    Exit For    '*** ppの検索終了
'
                End If
            Next pp
        Next jpT
'
        For pp = 1 To PtotalP    '*** 未登録部品の選別 ***
            If Left(PLSTP(pp, 1), 1) = "*" Then
                If PLSTP(pp, PdimP + 1) = Aitem0(ipT, 3) Or _
                   PLSTP(pp, PdimP + 1) = Aitem0(ipT, 4) Or _
                   PLSTP(pp, PdimP + 1) = Aitem0(ipT, 5) Then
'
                    If FLGheader = 0 Then
                        Call PRNheader0  '*** 部品表ヘッダー印刷 ***
                        Call PRNheader1  '*** 部品表項目見出し印刷 ***
                    End If
'
                    If FLGitemp = 0 Then
                        If FLGgyou >= Gyoumax - 1 Then
                            FLGend = 0
                            Call PRNfooter
                            Printer.NewPage   '*** 改ページ ***
'
                            FLGheader = 0
                            Call PRNheader0
                            Call PRNheader1
                        End If
'
                        Call PRNIndex        '*** 項目名印刷 ***
                        FLGitemp = 1    '*** 項目印刷済みにｾｯﾄ ***
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
                    End If
'
                    Call PRNmitouroku    '*** 未登録部品印刷 ***
                End If
            End If
        Next pp
    Next ipT
'
    FLGend = 1
    Call PRNfooter           '*** 部品表項末印刷 ***
'
    For pp = 1 To PtotalP
        Tmp1 = PLSTP(pp, PdimP + 1)
        Call GET_koumoku(Tmp1, Aitem0(), Anum0)
'
        If Tmp1 = "**" Then
            Call PRNhamidashi       '*** はみ出し部品印刷 ***
        End If
    Next pp
'
    Printer.EndDoc      '*** プリンター書き込み ***
'
    If FLGall = 1 Then GoTo KURIKAESHI      '*** 構成表による連続印刷 ***
'
OWARI:
    Plst_PRNstd.MousePointer = vbDefault        '*** 砂時計解除 ***
    DoEvents
    Timer1.Enabled = True
    Exit Sub
'    Unload Me
'
errh_P:
    txtPrinting.Visible = False
    Plst_PRNstd.MousePointer = vbDefault        '*** 砂時計解除 ***
'
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub S_Hyouji()
    optBikou.Value = True
'
'    cmdGo.ToolTipText = "プリンタの印刷設定を変更するには" _
'                    & "コントロールパネルの「プリンタとFAX」で行ってください。"
'
    If FLGall = 0 Then      '*** １つの部品表印刷 ***
        optdirect.Value = True
        FLGkouseihyou = 0
'
        txtkeisiki.Text = ""
        txtmeisyou.Text = ""
        txttantou.Text = ""
        txtkouban.Text = ""
        DaisuuT = "1"           '*** 仮数 ***
        txtdaisuu.Text = DaisuuT
        BaisuuT = "1"           '*** 仮数 ***
        txtbaisuu.Text = BaisuuT
'
        PFLnameP = PFLnameT
        DRVpartlistP = DRVpartlistT
'
        Call Plst_yomu           '*** 部品表を読みﾌｧｲﾙ内容表示 ***
'
    Else                    '*** 構成表より連続印刷 ***
        opthyou.Value = True
        optdirect.Enabled = False
        FLGkouseihyou = 1
'
        Call Klst_yomu       '*** 構成表を読み込み内容を表示 ***
'
        lblnamae.Caption = "印刷 ﾌｧｲﾙ名 ： <<構成表による>>"
        lblkomei.Caption = "小名称 ： "
        lblDate.Caption = "日付 ： "
        txtbaisuu.Text = "-"
        lblkomei.Enabled = False
        lblDate.Enabled = False
        txtbaisuu.Enabled = False
        lblbaisuu.Enabled = False
    End If
End Sub

Private Sub opthyou_Click()
    Call Klst_yomu       '*** 構成表を読み込み内容表示 ***
End Sub

Private Sub Timer1_Timer()
    Call cmdQuit_Click
End Sub

Private Sub txtbaisuu_Click()
    txtbaisuu.MousePointer = vbIbeam
End Sub

Private Sub txtbaisuu_LostFocus()
    txtbaisuu.MousePointer = vbArrow
End Sub

Private Sub txtdaisuu_Click()
    txtdaisuu.MousePointer = vbIbeam
End Sub

Private Sub txtdaisuu_LostFocus()
    txtdaisuu.MousePointer = vbArrow
End Sub

Private Sub txtKeisiki_Click()
    txtkeisiki.MousePointer = vbIbeam
End Sub

Private Sub txtKeisiki_LostFocus()
    txtkeisiki.MousePointer = vbArrow
End Sub

Private Sub txtkouban_Click()
    txtkouban.MousePointer = vbIbeam
End Sub

Private Sub txtkouban_LostFocus()
    txtkouban.MousePointer = vbArrow
End Sub

Private Sub txtmeisyou_Click()
    txtmeisyou.MousePointer = vbIbeam
End Sub

Private Sub txtmeisyou_LostFocus()
    txtmeisyou.MousePointer = vbArrow
End Sub

Private Sub txttantou_Click()
    txttantou.MousePointer = vbIbeam
End Sub

Private Sub txttantou_LostFocus()
    txttantou.MousePointer = vbArrow
End Sub

Private Sub DSPgamenBuhin()
    txtkeisiki.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtkeisiki.Top = 360 + (480 - 360) * HyoujiBairitu
    txtkeisiki.FontSize = 10 * HyoujiBairitu
    txtkeisiki.Width = 1215 * HyoujiBairitu
    txtkeisiki.Height = 285 * HyoujiBairitu
'
    lblkeisiki.Left = 480
    lblkeisiki.Top = 360 + (480 - 360) * HyoujiBairitu
    lblkeisiki.FontSize = 10 * HyoujiBairitu
    lblkeisiki.Width = 855 * HyoujiBairitu
    lblkeisiki.Height = txtkeisiki.Height
'
    txtmeisyou.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtmeisyou.Top = 360 + (840 - 360) * HyoujiBairitu
    txtmeisyou.FontSize = 10 * HyoujiBairitu
    txtmeisyou.Width = 3135 * HyoujiBairitu
    txtmeisyou.Height = 285 * HyoujiBairitu
'
    lblmeisyou.Left = 480
    lblmeisyou.Top = 360 + (840 - 360) * HyoujiBairitu
    lblmeisyou.FontSize = 10 * HyoujiBairitu
    lblmeisyou.Width = 855 * HyoujiBairitu
    lblmeisyou.Height = txtmeisyou.Height
'
    txttantou.Left = 480 + (1335 - 480) * HyoujiBairitu
    txttantou.Top = 360 + (1200 - 360) * HyoujiBairitu
    txttantou.FontSize = 10 * HyoujiBairitu
    txttantou.Width = 1215 * HyoujiBairitu
    txttantou.Height = 285 * HyoujiBairitu
'
    lbltantou.Left = 480
    lbltantou.Top = 360 + (1200 - 360) * HyoujiBairitu
    lbltantou.FontSize = 10 * HyoujiBairitu
    lbltantou.Width = 855 * HyoujiBairitu
    lbltantou.Height = txttantou.Height
'
    txtkouban.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtkouban.Top = 360 + (1560 - 360) * HyoujiBairitu
    txtkouban.FontSize = 10 * HyoujiBairitu
    txtkouban.Width = 3135 * HyoujiBairitu
    txtkouban.Height = 285 * HyoujiBairitu
'
    lblkouban.Left = 480
    lblkouban.Top = 360 + (1560 - 360) * HyoujiBairitu
    lblkouban.FontSize = 10 * HyoujiBairitu
    lblkouban.Width = 855 * HyoujiBairitu
    lblkouban.Height = txtkouban.Height
'
    txtdaisuu.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtdaisuu.Top = 360 + (1920 - 360) * HyoujiBairitu
    txtdaisuu.FontSize = 10 * HyoujiBairitu
    txtdaisuu.Width = 615 * HyoujiBairitu
    txtdaisuu.Height = 285 * HyoujiBairitu
'
    lbldaisuu.Left = 480
    lbldaisuu.Top = 360 + (1920 - 360) * HyoujiBairitu
    lbldaisuu.FontSize = 10 * HyoujiBairitu
    lbldaisuu.Width = 855 * HyoujiBairitu
    lbldaisuu.Height = txtdaisuu.Height
'
    lblnamae.Left = 480
    lblnamae.Top = 360 + (2640 - 360) * HyoujiBairitu
    lblnamae.FontSize = 10 * HyoujiBairitu
    lblnamae.Width = 3375 * HyoujiBairitu
    lblnamae.Height = 285 * HyoujiBairitu
'
    lblkomei.Left = 480
    lblkomei.Top = 360 + (3000 - 360) * HyoujiBairitu
    lblkomei.FontSize = 10 * HyoujiBairitu
    lblkomei.Width = 3855 * HyoujiBairitu
    lblkomei.Height = 285 * HyoujiBairitu
'
    lblDate.Left = 480
    lblDate.Top = 360 + (3360 - 360) * HyoujiBairitu
    lblDate.FontSize = 10 * HyoujiBairitu
    lblDate.Width = 2535 * HyoujiBairitu
    lblDate.Height = 285 * HyoujiBairitu
'
    txtbaisuu.Left = 480 + (3720 - 480) * HyoujiBairitu
    txtbaisuu.Top = 360 + (3360 - 360) * HyoujiBairitu
    txtbaisuu.FontSize = 10 * HyoujiBairitu
    txtbaisuu.Width = 615 * HyoujiBairitu
    txtbaisuu.Height = 285 * HyoujiBairitu
'
    lblbaisuu.Left = 480 + (3240 - 480) * HyoujiBairitu
    lblbaisuu.Top = 360 + (3360 - 360) * HyoujiBairitu
    lblbaisuu.FontSize = 10 * HyoujiBairitu
    lblbaisuu.Width = 495 * HyoujiBairitu
    lblbaisuu.Height = txtbaisuu.Height
'
    frmkinyuu.Left = 480 + (4800 - 480) * HyoujiBairitu
    frmkinyuu.Top = 360
    frmkinyuu.FontSize = 10 * HyoujiBairitu
    frmkinyuu.Width = 1455 * HyoujiBairitu
    frmkinyuu.Height = 855 * HyoujiBairitu
'
    optdirect.Left = 120 * HyoujiBairitu
    optdirect.Top = 240 * HyoujiBairitu
    optdirect.FontSize = 10 * HyoujiBairitu
    optdirect.Width = 1095 * HyoujiBairitu
    optdirect.Height = 255 * HyoujiBairitu
'
    opthyou.Left = 120 * HyoujiBairitu
    opthyou.Top = 480 * HyoujiBairitu
    opthyou.FontSize = 10 * HyoujiBairitu
    opthyou.Width = 1095 * HyoujiBairitu
    opthyou.Height = 255 * HyoujiBairitu
'
    frmTanka.Left = 480 + (4800 - 480) * HyoujiBairitu
    frmTanka.Top = 360 + (1440 - 360) * HyoujiBairitu
    frmTanka.FontSize = 10 * HyoujiBairitu
    frmTanka.Width = 1455 * HyoujiBairitu
    frmTanka.Height = 855 * HyoujiBairitu
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
    cmdGo.Left = 480 + (4800 - 480) * HyoujiBairitu
    cmdGo.Top = 360 + (2760 - 360) * HyoujiBairitu
    cmdGo.FontSize = 10 * HyoujiBairitu
    cmdGo.Width = 1455 * HyoujiBairitu
    cmdGo.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 480 + (4800 - 480) * HyoujiBairitu
    cmdQuit.Top = 360 + (3480 - 360) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1455 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
'
    txtPrinting.Width = 5880 * HyoujiBairitu
    txtPrinting.Height = 720 * HyoujiBairitu
    txtPrinting.FontSize = 10 * HyoujiBairitu
    txtPrinting.Left = (Me.Width - txtPrinting.Width) / 2
    txtPrinting.Top = (Me.Height - txtPrinting.Height) / 3
    txtPrinting.Visible = False
    txtPrinting.BackColor = &HC000&
    txtPrinting.ForeColor = &HFFFFFF
End Sub

