VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Plst_PRNlst 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "部品一覧表印刷"
   ClientHeight    =   4095
   ClientLeft      =   2475
   ClientTop       =   1200
   ClientWidth     =   6630
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
   Icon            =   "Plst_PRNlst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4095
   ScaleWidth      =   6630
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3480
      Top             =   360
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "ﾌｧｲﾙ出力(&F)"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtmeisyou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  '矢印
      TabIndex        =   6
      Text            =   "TV SOUND MULTI MODULATOR"
      Top             =   1200
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
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame frmkinyuu 
      BackColor       =   &H00008000&
      Caption         =   "項目記入"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      Begin VB.OptionButton opthyou 
         BackColor       =   &H00008000&
         Caption         =   "構成表"
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
   Begin VB.CommandButton cmdGo 
      Caption         =   "印刷(&P)"
      Default         =   -1  'True
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   10
   End
   Begin VB.Label lbldate 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "日付 ： 1997/09/19"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblkomei 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "小名称 ：ABCDEFGHIJKLMNOPQRSTUVWXY"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   3855
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
      Top             =   1200
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
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblnamae 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "印刷 ﾌｧｲﾙ名 ：A1234-00.2AB"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
End
Attribute VB_Name = "Plst_PRNlst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 部品一覧表印刷 ***
'**********************
'
Option Explicit
'
Dim FLGgyou As Integer, FLGpage As Integer
Dim FLGend As Integer, FLGheader As Integer
Dim FLGitemp As Integer, Filenum As Integer
Dim kijunX As Integer, kijunY As Integer
Dim haba1X As Integer, haba2X As Integer, haba3X As Integer
Dim haba4X As Integer, haba5X As Integer, haba6X As Integer
Dim gyoukan As Integer
Dim moji_zureX As Integer, moji_zureY As Integer
Dim NowpointX As Integer, NowpointY As Integer
Dim Tmpdata As String
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
Dim Pdata_codep As String       '*** 部品ｺｰﾄﾞ ***
Dim Pdata_kikakup As String     '*** 部品規格 ***
Dim Pdata_bangoup As String     '*** 部品番号 ***
Dim Pdata_makerp As String      '*** ﾒｰｶｰ ***
Dim Pdata_bikoup As String      '*** 備考欄 ***
Dim Pdata_teikaku As String     '*** 定格など ***

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 480 + (6750 - 960) * HyoujiBairitu + 480
    Height = 480 + (4470 - 960) * HyoujiBairitu + 480
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
    Me.Caption = STATUS
    CATnoT = ""
    CATnameT = ""
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
    If UnloadMode = vbFormControlMenu Then
        '
    End If
End Sub

Private Sub PRNhamidashi()
    Printer.Print "はみ出し部品  " & PLSTP(pp, 0) & "  " & PLSTP(pp, 1) & "  " & PLSTP(pp, 2)
End Sub

Private Sub FLEindex()
'                   *** 項目名ファイル出力 ***
    Write #Filenum, Aitem0(ipT, 1), "", "", "", "", "", "", "", "", ""
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
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Aitem0(ipT, 1)
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
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.CurrentX = NowpointX + 567
    Printer.CurrentY = NowpointY + moji_zureY
    SETfont_size 9, 1
    Printer.Print "M5904-03";
'
    Printer.CurrentX = NowpointX + 4252
    Printer.CurrentY = NowpointY + moji_zureY
    SETfont_size 10.8, 0
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷";
'
    Printer.CurrentX = NowpointX + 9072
    Printer.CurrentY = NowpointY + moji_zureY
    If FLGend = 0 Then
        Printer.Print "つづく"
    Else
        Printer.Print "おわり"
    End If
End Sub

Private Sub FLEkoumoku()
'                   *** 項目ファイル出力 ***
    Dim maxV As String, maxA As String, maxW As String, maxT As String, maxR As String
    Dim ist As Integer
'
    ist = InStr(1, Pdata_teikaku, "!V:")
    If ist = 0 Then
        maxV = ""
    Else
        maxV = Mid(Pdata_teikaku, ist + 3)
        ist = InStr(1, maxV, ",")
        If ist = 0 Then
            maxV = ""
        Else
            maxV = Mid(maxV, 1, ist - 1)
        End If
    End If
'
    ist = InStr(1, Pdata_teikaku, "!A:")
    If ist = 0 Then
        maxA = ""
    Else
        maxA = Mid(Pdata_teikaku, ist + 3)
        ist = InStr(1, maxA, ",")
        If ist = 0 Then
            maxA = ""
        Else
            maxA = Mid(maxA, 1, ist - 1)
        End If
    End If
'
    ist = InStr(1, Pdata_teikaku, "!W:")
    If ist = 0 Then
        maxW = ""
    Else
        maxW = Mid(Pdata_teikaku, ist + 3)
        ist = InStr(1, maxW, ",")
        If ist = 0 Then
            maxW = ""
        Else
            maxW = Mid(maxW, 1, ist - 1)
        End If
    End If
'
    ist = InStr(1, Pdata_teikaku, "!T:")
    If ist = 0 Then
        maxT = ""
    Else
        maxT = Mid(Pdata_teikaku, ist + 3)
        ist = InStr(1, maxT, ",")
        If ist = 0 Then
            maxT = ""
        Else
            maxT = Mid(maxT, 1, ist - 1)
        End If
    End If
'
    ist = InStr(1, Pdata_teikaku, "!R:")
    If ist = 0 Then
        maxR = ""
    Else
        maxR = Mid(Pdata_teikaku, ist + 3)
        ist = InStr(1, maxR, ",")
        If ist = 0 Then
            maxR = ""
        Else
            maxR = Mid(maxR, 1, ist - 1)
        End If
    End If
'
    Write #Filenum, Pdata_bangoup, Pdata_codep, Pdata_kikakup, Pdata_makerp, Pdata_bikoup, maxV, maxA, maxW, maxT, maxR
End Sub

Private Sub PRNkoumoku()
'                   *** 項目印刷 ***
    Dim itiji As String
    Dim Pdata As String
'
    FLGgyou = FLGgyou + 1
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)  '*** フォントサイズ設定 ***
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
    Printer.Print Pdata_bangoup
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "  " & Pdata_codep
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Pdata_kikakup
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Pdata_makerp
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Pdata_bikoup
End Sub

Private Sub FLEheader1()
    Write #Filenum, "番号", "ｺｰﾄﾞ番号", "品 名", "ﾒｰｶｰ", "備 考", "最大電圧", "最大電流", "最大損失", "最大接合温度", "熱抵抗"
End Sub

Private Sub PRNheader1()
'                   *** 部品一覧表項目見出し印刷 ***
'                   567twip=10mm,1440twip=1inch,<11320/10490>
    FLGgyou = 3
    haba1X = 425
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)  '*** フォントサイズ設定 ***
    Printer.Print "行"
'
    haba2X = 850
    NowpointX = kijunX + haba1X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 113 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Call SETfont_size(10.8, 1)  '*** フォントサイズ設定 ***
    Printer.Print "番号"
'
    haba3X = 1786
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX + 340
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "ｺｰﾄﾞ番号"
'
    haba4X = 3883
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 1588 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "品 名"
'
    haba5X = 964
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 170 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "ﾒｰｶｰ"
'
    haba6X = 2582
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + 907 + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "備 考"
End Sub

Private Sub FLEheader0()
'                   *** 部品表ヘッダーファイル出力 ***
    Dim i As Integer
'
    Call remove_PLT(PFLnameP, Tmpdata)
'
    Write #Filenum, "形式： " & Trim(CATnoT), "名称： " & Trim(CATnameT), _
            "電気部品一覧表", "図番： " & Trim(Tmpdata), "小名称： " & Trim(PlistnameP), "", "", "", "", ""
'
    FLGheader = 1       '*** 出力済み ***
End Sub

Private Sub PRNheader0()
'                   *** 部品表ヘッダー印刷 ***
    Dim i As Integer
'                   567twip=10mm,1440twip=1inch,<11320/10490>
    kijunX = 830
    kijunY = 737
    gyoukan = 340
    moji_zureX = 113
    moji_zureY = 60
    FLGgyou = 0
    FLGheader = 1   '*** ヘッダー印刷済みにセット
'
    Printer.PaperSize = vbPRPSA4            '*** A4 ***
    Printer.Orientation = vbPRORPortrait    '*** ﾎﾟｰﾄﾚｰﾄ ***
'
    Printer.CurrentX = kijunX + 2268
    Printer.CurrentY = 737
    Call SETfont_size(17, 1)    '*** フォント,サイズ設定 ***
    Printer.Print "《　電  気  部　品　一  覧  表　》"
'
    Printer.CurrentX = kijunX + 9072
    Printer.CurrentY = 850
    Call SETfont_size(10.8, 1)  '*** フォント,サイズ設定 ***
'
    Printer.Print FLGpage & " ﾍﾟｰｼﾞ"    '*** ページ印刷 ***
    FLGpage = FLGpage + 1
'
    Printer.DrawWidth = 6
    Printer.FillStyle = vbFSTransparent
'
    FLGgyou = 1
    haba1X = 2437
    NowpointX = kijunX
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba1X, gyoukan), 0, B
    Call SETfont_size(10.8, 1)  '*** フォント,サイズ設定 ***
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "型式： " & CATnoT
'
    haba2X = 5471   '6183
    NowpointX = kijunX + haba1X
    Printer.Line (NowpointX, NowpointY)-Step(haba2X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "名称： " & CATnameT
'
    haba3X = 2582   '1870
    NowpointX = kijunX + haba1X + haba2X
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print PlistdateP     '*** 日付印刷 ***
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
    haba3X = 4407
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
End Sub

Private Sub Klst_yomu()
    DRVconstT = TMPdir1 & "\constlst.cod"
    Call RDconst_lst(DRVconstT, CATnoT, CATnameT, ZubanT, PersonT, OrgdateT, RevdateT, CheckdateT, OutdateT, _
                KLSTT(), KtotalT, KdimT, KoubanT, DaisuuT, KbikouT, KyobiAT, KyobiBT)   '*** 構成表読み込み ***
'
    txtkeisiki.Text = CATnoT
    txtMeisyou.Text = CATnameT
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

Private Sub FLEmitouroku()
'                   *** 未登録部品ファイル出力 ***
    Write #Filenum, PLSTP(pp, 0), "未登録部品", Mid(PLSTP(pp, 1), 2), "******", PLSTP(pp, 2), "", "", "", "", ""
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
    Call SETfont_size(10.8, 1)  '*** フォントサイズ設定 ***
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
    Printer.Print PLSTP(pp, 0)
'
    NowpointX = kijunX + haba1X + haba2X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba3X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "  未登録部品"
'
    NowpointX = kijunX + haba1X + haba2X + haba3X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba4X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print Mid(PLSTP(pp, 1), 2)
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba5X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print "******"
'
    NowpointX = kijunX + haba1X + haba2X + haba3X + haba4X + haba5X
    NowpointY = kijunY + gyoukan * FLGgyou
    Printer.Line (NowpointX, NowpointY)-Step(haba6X, gyoukan), 0, B
    Printer.CurrentX = NowpointX + moji_zureX
    Printer.CurrentY = NowpointY + moji_zureY
    Printer.Print PLSTP(pp, 2)
End Sub

Private Sub cmdFile_Click()
    FLGfile = 1             '*** 一覧表ファイル出力 ***
    cmdGo_Click
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer, j As Integer, np As Integer, q As Integer
    Dim nlen As Integer
    Dim Ckdim As Integer, Gyoumax As Integer
    Dim Bmoji As Integer
    Dim MKsitei As String
    Dim Tmp1 As String
    Dim FLGmojime As Integer
    Dim FLGbanme As Integer
    Dim FFLname As String
    Dim objPrinter As Printer
'
    If FLGfile = 2 Then             '*** 未定 → 一覧表印刷 ***
        FLGfile = 0
    End If
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
                    DRVpartlistP = TMPplst & "\" & PFLnameP        '*** xxxxxxxx.yyy を検索 ***
                    Tmp1 = Dir(DRVpartlistP)
'                    If Tmp1 = "" Then
'                        PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
'                        DRVpartlistP = TMPplst & "\" & PFLnameP    '*** 共有フォルダーを検索 ***
'                        Tmp1 = Dir(DRVpartlistP)
                        If Tmp1 = "" Then
                            i = MsgBox("ファイル " & DRVpartlistP & "が見つかりません。", vbCritical)
                            Exit Sub
'
                        End If
'                    End If
                End If
            End If
        Next FLGbanme
    End If
'
    If FLGfile = 0 Then
        flag_cancel = False
        Printer_Window.Show 1
            If flag_cancel = True Then GoTo OWARI
'
        For Each objPrinter In Printers
            If objPrinter.DeviceName = strMyPrinter Then
                Set Printer = objPrinter    'オブジェクトに代入
            End If
        Next
    End If
'
    Me.MousePointer = vbHourglass      '*** 砂時計 ***
    FLGbanme = 1
    If FLGall = 1 Then
KURIKAESHI:
        If FLGbanme > KtotalT Then GoTo OWARI    '*** 終わりにｼﾞｬﾝﾌﾟ ***
'
        If Left(KLSTT(FLGbanme, 2), 1) = "B" And Val(KLSTT(FLGbanme, 4)) <> 0 Then
            PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
            DRVpartlistP = TMPplst & "\" & PFLnameP            '*** .PLT を検索 ***
            Tmp1 = Dir(DRVpartlistP)
            If Tmp1 = "" Then
                PFLnameP = KLSTT(FLGbanme, 2)
                If Len(PFLnameP) > 8 Then
                    PFLnameP = Left(PFLnameP, 8) & "." & Mid(PFLnameP, 9)
                End If
                DRVpartlistP = TMPplst & "\" & PFLnameP        '*** xxxxxxxx.yyy を検索 ***
'
'                Tmp1 = Dir(DRVpartlistP)
'                If Tmp1 = "" Then
'                    PFLnameP = KLSTT(FLGbanme, 2) & ".PLT"
'                    DRVpartlistP = TMPplst & "\" & PFLnameP    '*** 共有フォルダーを検索 ***
'                End If
'
            End If
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
    If FLGfile = 1 Then
        nlen = Len(PFLnameP)
        If InStr(1, UCase(PFLnameP), ".PLT") <> 0 Then
            nlen = nlen - 4
        End If
'
        FFLname = "J" & Mid(PFLnameP, 2, nlen - 1) & ".CSV"
        Tmp1 = " !!! 部品一覧表 " & FFLname & " をファイルに出力中 !!!"
'
        Filenum = FreeFile     '*** 空いているファイル番号を得る ***
        Open TMPdir1 & "\" & FFLname For Output As #Filenum   '*** ファイルオープン ***
    Else
        Tmp1 = " !!! 部品一覧表 " & PFLnameP & " をプリントバッファーに転送中 !!!"
    End If
'
    Me.Caption = Tmp1
    FLGpage = 1             '*** ページ初期化
    FLGheader = 0           '*** ヘッダー印刷初期化
    Gyoumax = 43            '*** 印刷行最大値 ***
    pps = 1
'
    For ipT = 1 To Anum0
        DRVindexT = Xcont0(2) & "\" & Aitem0(ipT, 0) & "\" & Aitem0(ipT, 0) & "INDEX.COD"
'
        Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)      '*** INDEX.COD 読み込み ***
        FLGitemp = 0    '*** 項目印刷ﾌﾗｸﾞｸﾘｱｰ
        jpsT = pps
'
        For pp = jpsT To PtotalP
'
            For jpT = 1 To BnumT
                If ((Left(PLSTP(pp, 1), 1) = "L") And (Mid(PLSTP(pp, 1), 2, 4) = BindexT(jpT, 0))) Then
                    pps = pp
'
                    If FLGheader = 0 Then
                        If FLGfile = 1 Then
                            Call FLEheader0  '*** 部品表ヘッダー出力 ***
                            Call FLEheader1  '*** 部品表項目見出し出力 ***
                        Else
                            Call PRNheader0  '*** 部品表ヘッダー印刷 ***
                            Call PRNheader1  '*** 部品表項目見出し印刷 ***
                        End If
                    End If
'
                    If FLGitemp = 0 Then
                        If FLGfile = 1 Then
                            Call FLEindex        '*** 項目名ファイル出力 ***
                        Else
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
                        End If
'
                        FLGitemp = 1    '*** 項目印刷済みにｾｯﾄ ***
                    End If
'
                    Call SET_DRVmain(DRVmainT, Aitem0(), ipT, BindexT(), jpT)
'
                    Call RDmain(DRVmainT, CmainT(), CnumT, CdimT)   '*** MAIN.COD 読み込み ***
'
                    For kpT = 1 To CnumT
                        Pdata_codep = "L" & BindexT(jpT, 0) & "-" & CmainT(kpT, 0)
'
                            If PLSTP(pp, 1) = Pdata_codep Then   '*** 個別にﾃﾞｰﾀｰ作成
                                Pdata_bangoup = PLSTP(pp, 0)
'
                                If BindexT(jpT, 5) = "000" Then
                                    Pdata_makerp = CmainT(kpT, 13)
'
                                ElseIf BindexT(jpT, 5) = "998" Then
                                    If PLSTP(pp, 3) = "0" Then
                                        MKsitei = "0"
                                    Else
                                        MKsitei = PLSTP(pp, 3)
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
                                Pdata_teikaku = CmainT(kpT, 2)        '*** 定格など ***
'
                                Call GET_shitei2(CmainT(kpT, 3), Pdata_kikakup, q)   '*** ! ? の記入 ***
'
                                If CmainT(kpT, 16) = "1" Then         '*** 特記事項記入
                                    If PLSTP(pp, 4) = "" Or PLSTP(pp, 4) = "*" Then
                                        '
                                    Else
                                        Pdata_kikakup = Pdata_kikakup & PLSTP(pp, 4)
                                    End If
                                End If
                                Pdata_bikoup = PLSTP(pp, 2)
'
                                If FLGfile = 1 Then
                                    Call FLEkoumoku      '*** 部品表項目ファイル出力 ***
                                Else
                                    Call PRNkoumoku      '*** 部品表項目印刷 ***
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
                                End If
'
                                Exit For    '*** ppの検索終了
                            End If
                    Next kpT
'
                End If
            Next jpT
        Next pp
'
        For pp = 1 To PtotalP    '*** 未登録部品の選別 ***
            If Left(PLSTP(pp, 1), 1) = "*" Then
                If PLSTP(pp, PdimP + 1) = Aitem0(ipT, 3) Or _
                   PLSTP(pp, PdimP + 1) = Aitem0(ipT, 4) Or _
                   PLSTP(pp, PdimP + 1) = Aitem0(ipT, 5) Then
'
                    If FLGheader = 0 Then
                        If FLGfile = 1 Then
                            Call FLEheader0
                            Call FLEheader1
                        Else
                            Call PRNheader0  '*** 部品表ヘッダー印刷 ***
                            Call PRNheader1  '*** 部品表項目見出し印刷 ***
                        End If
                    End If
'
                    If FLGitemp = 0 Then
                        If FLGfile = 1 Then
                            Call FLEindex        '*** 項目名ファイル出力 ***
                        Else
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
                        End If
'
                        FLGitemp = 1    '*** 項目印刷済みにｾｯﾄ ***
                    End If
'
                    If FLGfile = 1 Then
                        Call FLEmitouroku    '*** 未登録部品ファイル出力 ***
                    Else
                        Call PRNmitouroku    '*** 未登録部品印刷 ***
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
                    End If
                End If
            End If
        Next pp
    Next ipT
'
    FLGend = 1
'
    If FLGfile = 1 Then
        Close #Filenum
    Else
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
    End If
'
    If FLGall = 1 Then GoTo KURIKAESHI      '*** 構成表による連続印刷 ***
'
OWARI:
    Me.MousePointer = vbDefault    '*** 砂時計解除 ***
    Timer1.Enabled = True
    Exit Sub
'    Unload Me
'
errh_P:
    Me.Caption = STATUS
    Me.MousePointer = vbDefault    '*** 砂時計解除 ***
'
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Call cmdQuit_Click
End Sub

Private Sub S_Hyouji()
'    cmdGo.ToolTipText = "プリンタの印刷設定を変更するには" _
'                    & "コントロールパネルの「プリンタとFAX」で行ってください。"
'
    If FLGall = 0 Then      '*** １つの部品表印刷 ***
        optdirect.Value = True
'
        txtkeisiki.Text = CATnoT
        txtMeisyou.Text = CATnameT
'
        PFLnameP = PFLnameT
        DRVpartlistP = DRVpartlistT
'
        Call Plst_yomu           '*** 部品表を読みﾌｧｲﾙ内容表示 ***
'
    Else                    '*** 構成表より連続印刷 ***
        opthyou.Value = True
'
        Call Klst_yomu       '*** 構成表を読み込み内容を表示 ***
'
        lblnamae.Caption = "印刷 ﾌｧｲﾙ名 ： <<構成表による>>"
        lblkomei.Caption = "小名称 ： "
        lblDate.Caption = "日付 ： "
        lblkomei.Enabled = False
        lblDate.Enabled = False
    End If
'
    Select Case FLGfile
    Case 0
        cmdGo.Enabled = True
        cmdFile.Enabled = False
    Case 1
        cmdGo.Enabled = False
        cmdFile.Enabled = True
    Case 2
        cmdGo.Enabled = True
        cmdFile.Enabled = True
    End Select
End Sub

Private Sub optDirect_Click()
'
End Sub

Private Sub opthyou_Click()
    Call Klst_yomu       '*** 構成表を読み込み内容表示 ***
End Sub

Private Sub txtKeisiki_Click()
    txtkeisiki.MousePointer = vbIbeam
End Sub

Private Sub txtKeisiki_LostFocus()
    txtkeisiki.MousePointer = vbArrow
    CATnoT = Trim(txtkeisiki.Text)
End Sub

Private Sub txtmeisyou_Click()
    txtMeisyou.MousePointer = vbIbeam
End Sub

Private Sub txtmeisyou_LostFocus()
    txtMeisyou.MousePointer = vbArrow
    CATnameT = Trim(txtMeisyou.Text)
End Sub

Private Sub DSPgamenBuhin()
    txtkeisiki.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtkeisiki.Top = 480 + (720 - 480) * HyoujiBairitu
    txtkeisiki.FontSize = 10 * HyoujiBairitu
    txtkeisiki.Width = 1215 * HyoujiBairitu
    txtkeisiki.Height = 285 * HyoujiBairitu
'
    lblkeisiki.Left = 480
    lblkeisiki.Top = 480 + (720 - 480) * HyoujiBairitu
    lblkeisiki.FontSize = 10 * HyoujiBairitu
    lblkeisiki.Width = 855 * HyoujiBairitu
    lblkeisiki.Height = txtkeisiki.Height
'
    txtMeisyou.Left = 480 + (1335 - 480) * HyoujiBairitu
    txtMeisyou.Top = 480 + (1200 - 480) * HyoujiBairitu
    txtMeisyou.FontSize = 10 * HyoujiBairitu
    txtMeisyou.Width = 3135 * HyoujiBairitu
    txtMeisyou.Height = 285 * HyoujiBairitu
'
    lblMeisyou.Left = 480
    lblMeisyou.Top = 480 + (1200 - 480) * HyoujiBairitu
    lblMeisyou.FontSize = 10 * HyoujiBairitu
    lblMeisyou.Width = 855 * HyoujiBairitu
    lblMeisyou.Height = txtMeisyou.Height
'
    lblnamae.Left = 480
    lblnamae.Top = 480 + (1920 - 480) * HyoujiBairitu
    lblnamae.FontSize = 10 * HyoujiBairitu
    lblnamae.Width = 3375 * HyoujiBairitu
    lblnamae.Height = 285 * HyoujiBairitu
'
    lblkomei.Left = 480
    lblkomei.Top = 480 + (2400 - 480) * HyoujiBairitu
    lblkomei.FontSize = 10 * HyoujiBairitu
    lblkomei.Width = 3855 * HyoujiBairitu
    lblkomei.Height = 285 * HyoujiBairitu
'
    lblDate.Left = 480
    lblDate.Top = 480 + (2880 - 480) * HyoujiBairitu
    lblDate.FontSize = 10 * HyoujiBairitu
    lblDate.Width = 2535 * HyoujiBairitu
    lblDate.Height = 285 * HyoujiBairitu
'
    frmkinyuu.Left = 480 + (4800 - 480) * HyoujiBairitu
    frmkinyuu.Top = 480
    frmkinyuu.FontSize = 10 * HyoujiBairitu
    frmkinyuu.Width = 1335 * HyoujiBairitu
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
    cmdGo.Left = 480 + (4800 - 480) * HyoujiBairitu
    cmdGo.Top = 480 + (1680 - 480) * HyoujiBairitu
    cmdGo.FontSize = 10 * HyoujiBairitu
    cmdGo.Width = 1335 * HyoujiBairitu
    cmdGo.Height = 495 * HyoujiBairitu
'
    cmdFile.Left = 480 + (4800 - 480) * HyoujiBairitu
    cmdFile.Top = 480 + (2400 - 480) * HyoujiBairitu
    cmdFile.FontSize = 10 * HyoujiBairitu
    cmdFile.Width = 1335 * HyoujiBairitu
    cmdFile.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 480 + (4800 - 480) * HyoujiBairitu
    cmdQuit.Top = 480 + (3120 - 480) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1335 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
End Sub

