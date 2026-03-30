VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Plst_PRNlbl 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "部品ラベル印刷"
   ClientHeight    =   4215
   ClientLeft      =   2475
   ClientTop       =   1485
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
   Icon            =   "Plst_PRNlbl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4215
   ScaleWidth      =   6630
   Begin VB.TextBox txtbaisuu 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      MousePointer    =   1  '矢印
      TabIndex        =   17
      Text            =   "1"
      Top             =   3120
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
      Enabled         =   0   'False
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
      Enabled         =   0   'False
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
      Height          =   615
      Left            =   4800
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "終了(&Q)"
      Height          =   615
      Left            =   4800
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      FontSize        =   10
   End
   Begin VB.Label lblComment 
      BackColor       =   &H00808000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾌﾟﾘﾝﾀ設定 ： 用紙幅 => 3302、用紙長 =>5334 (mm/10)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   360
      TabIndex        =   20
      Top             =   3600
      Width           =   4335
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
      Top             =   3120
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
      Top             =   3120
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
      Top             =   2760
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
      Enabled         =   0   'False
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
      Enabled         =   0   'False
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
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "終了(&X)"
      Begin VB.Menu mnuBack 
         Caption         =   "ﾒﾆｭｰへ戻る(&B)"
      End
      Begin VB.Menu mnuAQuit 
         Caption         =   "EＥOSの終了(&X)"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnusetumei 
         Caption         =   "説明(&H)"
      End
      Begin VB.Menu mnuversion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ(&V)"
      End
   End
End
Attribute VB_Name = "Plst_PRNlbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 部品ラベル印刷 ***
'**********************
'
Option Explicit
'
    Dim FLGkouseihyou As Integer
    Dim FLGinit As Integer, FLGcolumn As Integer
    Dim kijunX As Integer, kijunY As Integer
    Dim TacckHabaX As Integer
    Dim haba1X As Integer, haba2X As Integer, haba3X As Integer, haba4X As Integer
    Dim NowpointX As Integer, NowpointY As Long
    Dim gyoukan As Integer
    Dim Tmpdata As String, Baisuu As String
'
    Dim Pdata_List() As String
    Dim Pdata_bangou() As String
    Dim Pdata_code() As String      '*** Pdata_code(FLGtuzuki) ***
    Dim Pdata_kikaku() As String    '*** Pdata_kikaku(FLGtuzuki) ***
    Dim Pdata_kosuu() As String     '*** Pdata_kosuu(FLGtuzuki) ***
    Dim Pdata_sousuu() As String    '*** Pdata_sousuu(FLGtuzuki) ***
    Dim Pdata_maker() As String     '*** Pdata_maker(FLGtuzuki) ***
    Dim Pdata_bikou() As String     '*** Pdata_bikou(FLGtuzuki) ***
'
    Dim Itti_List() As String
    Dim Itti_Bangou() As String
    Dim Itti_code() As String
    Dim Itti_kikaku() As String
    Dim Itti_kosuu() As String
    Dim Itti_sousuu() As String
    Dim Itti_maker() As String
    Dim Itti_bikou() As String
'
    Dim Pdata_Listp As String       '*** 部品表名称 ***
    Dim Pdata_bangoup As String     '*** 通し番号 ***
    Dim Pdata_codep As String       '*** 部品ｺｰﾄﾞ ***
    Dim Pdata_kikakup As String     '*** 部品規格 ***
    Dim Pdata_kosuup As String      '*** 個数 ***
    Dim Pdata_sousuup As String     '*** 総数 ***
    Dim Pdata_makerp As String      '*** ﾒｰｶｰ ***
    Dim Pdata_bikoup As String      '*** 備考欄 ***

Private Sub PRNinit0()
'       567twip = 10mm, 1440twip = 1inch,<16114>
'       1/6inch = 240twip, 7/6inch = 1680twip
    kijunX = 0
    kijunY = 720
    NowpointX = kijunX
    NowpointY = kijunY
    Printer.CurrentX = NowpointX
    Printer.CurrentY = NowpointY
    SETfont_size 10.8, 0    '*** フォントサイズ設定 ***
'    Printer.Height = 15120  '*** 10.5inch ***
'    Printer.Height = 30240  '*** 21.0inch ***
'    Printer.Width = 18720   '*** 13.0inch ***
'
    gyoukan = 240           '*** 行間 1/6inch ***
'
    TacckHabaX = 3458
    haba1X = 1700
    haba2X = 2041
    haba3X = 1758
    haba4X = 1814
'    kijunY = 0
'
End Sub

Private Sub PRNkoumoku(Frow As Integer)
'                   *** 項目印刷 ***
    Dim i As Integer
'
    Frow = Frow - 1
'
        Printer.CurrentY = NowpointY
    For i = 0 To Frow
        Printer.CurrentX = NowpointX + TacckHabaX * i
        Printer.Print Kouban;
'
        Printer.CurrentX = NowpointX + haba1X + TacckHabaX * i
        Printer.Print CATno;
    Next i
    Printer.Print " "
'
        NowpointY = NowpointY + gyoukan
        Printer.CurrentY = NowpointY
    For i = 0 To Frow
        Printer.CurrentX = NowpointX + TacckHabaX * i
        Printer.Print Pdata_List(i);
'
        Printer.CurrentX = NowpointX + haba2X + TacckHabaX * i
        Printer.Print "<" & Pdata_bangou(i) & " >";
    Next i
    Printer.Print " "
'
        NowpointY = NowpointY + gyoukan
        Printer.CurrentY = NowpointY
    For i = 0 To Frow
        Printer.CurrentX = NowpointX + TacckHabaX * i
        Printer.Print Pdata_code(i);
'
        Printer.CurrentX = NowpointX + haba3X + TacckHabaX * i
        Printer.Print Pdata_maker(i);
    Next i
    Printer.Print " "
'
        NowpointY = NowpointY + gyoukan
        Printer.CurrentY = NowpointY
    For i = 0 To Frow
        Printer.CurrentX = NowpointX + TacckHabaX * i
        Printer.Print Pdata_kikaku(i);
    Next i
    Printer.Print " "
'
    Tmpdata = " ="
        NowpointY = NowpointY + gyoukan
        Printer.CurrentY = NowpointY
    For i = 0 To Frow
        Printer.CurrentX = NowpointX + TacckHabaX * i
        Printer.Print Pdata_bikou(i);
'
        Printer.CurrentX = NowpointX + TacckHabaX * i + haba4X
        Printer.Print Tmpdata & Pdata_sousuu(i) & Tmpdata;
    Next i
    Printer.Print " "
'
    FLGcolumn = FLGcolumn + 1
'
    If FLGcolumn = 17 Then
        Printer.EndDoc  '*** 印刷 ***
'
        FLGcolumn = 0
        NowpointX = kijunX
        NowpointY = kijunY
    Else
        NowpointY = NowpointY + gyoukan * 3
    End If
End Sub

Private Sub Klst_yomu()
    DRVconst = XCONT0(4) & "\constlst.cod"
    RDconst_lst     '*** 構成表読み込み ***
'
    txtkeisiki.Text = CATno
    txtmeisyou.Text = CATname
    txttantou.Text = Person
    txtkouban.Text = Kouban
    txtdaisuu.Text = Daisuu
End Sub

Private Sub Plst_yomu()
'                   *** 部品表を読む ***
    Dim i As Integer
    Dim tmpdata1 As String, tmpdata2 As String
'
    lblnamae.Caption = "印刷 ﾌｧｲﾙ名 ： " & UCase(PFLname)
'
    RDpartlist      '*** 部品表読み込み ***
'
    If FLGall = 0 Then
        lblkomei.Caption = "小名称 ： " & Plistname
        lbldate.Caption = "日付 ： " & Plistdate
        txtbaisuu.Text = Baisuu
    End If
'
    For i = 1 To Ptotal
        tmpdata1 = PLST(i, 0)
'
        GETsymbol tmpdata1, tmpdata2
'
        PLST(i, Pdim0 + 1) = tmpdata1
        PLST(i, Pdim0 + 2) = tmpdata2
    Next i
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer, j As Integer, np As Integer, q As Integer
    Dim RowDim As Integer, FLGrow As Integer
    Dim MKsitei As String
    Dim Tmp1 As String
    Dim Hizuke As String
    Dim FLGitti As Integer, FLGtuzuki As Integer, TuzukiDim As Integer
    Dim FLGbanme As Integer, FLGmojime As Integer
'
    FLGinit = 0             '*** プリンタ初期化フラグリセット ***
    TuzukiDim = 50          '*** 同一品名異指定/未登録 メモリー初期化 ***
    ReDim Itti_List(TuzukiDim)
    ReDim Itti_Bangou(TuzukiDim)
    ReDim Itti_code(TuzukiDim)
    ReDim Itti_kikaku(TuzukiDim)
    ReDim Itti_kosuu(TuzukiDim)
    ReDim Itti_sousuu(TuzukiDim)
    ReDim Itti_maker(TuzukiDim)
    ReDim Itti_bikou(TuzukiDim)
'
    RowDim = 5              '*** 印刷メモリー初期化 ***
    ReDim Pdata_List(RowDim)
    ReDim Pdata_bangou(RowDim)
    ReDim Pdata_code(RowDim)
    ReDim Pdata_kikaku(RowDim)
    ReDim Pdata_kosuu(RowDim)
    ReDim Pdata_sousuu(RowDim)
    ReDim Pdata_maker(RowDim)
    ReDim Pdata_bikou(RowDim)
    FLGrow = 0
    Reset_Pdata RowDim
    FLGcolumn = 0       '*** 17片 ***
'
    Hizuke = Date
'
    i = MsgBox("プリンタに５列のタック用紙をセットしてください。", vbInformation Or vbOKCancel)
    If i = vbCancel Then
        Exit Sub
    End If
'
    If FLGall = 1 Then      '*** ファイルの有無チェック ***
        For FLGbanme = 1 To Ktotal
            If Left(KLST(FLGbanme, 2), 1) = "B" And Val(KLST(FLGbanme, 4)) <> 0 Then
                PFLname = KLST(FLGbanme, 2)
                If Len(PFLname) > 8 Then
                    PFLname = Left(PFLname, 8) & "." & Mid(PFLname, 9)
                End If
                DRVpartlist = XCONT0(4) & "\" & PFLname
                Tmp1 = Dir(DRVpartlist)
                If Tmp1 = "" Then
                    i = MsgBox("ファイル " & DRVpartlist & "が見つかりません。", vbCritical)
                    Exit Sub
'
                End If
            End If
        Next FLGbanme
    End If
'
    On Error GoTo errh_P
    With CommonDialog1
        .CancelError = True
        .Flags = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDUseDevModeCopies
        .ShowPrinter   '*** Printer Default = True
    End With
'
    Plst_PRNlbl.MousePointer = ccHourglass  '*** 砂時計 ***
    FLGbanme = 1
    If FLGall = 1 Then
KURIKAESHI:
        If FLGbanme > Ktotal Then GoTo OWARI    '*** 終わりにｼﾞｬﾝﾌﾟ ***
'
        If Left(KLST(FLGbanme, 2), 1) = "B" And Val(KLST(FLGbanme, 4)) <> 0 Then
            PFLname = KLST(FLGbanme, 2)
            Pdata_Listp = UCase(PFLname)
            If Len(PFLname) > 8 Then
                PFLname = Left(PFLname, 8) & "." & Mid(PFLname, 9)
            End If
'
            DRVpartlist = XCONT0(4) & "\" & PFLname
            Baisuu = Val(KLST(FLGbanme, 4))
            FLGbanme = FLGbanme + 1
        Else
            FLGbanme = FLGbanme + 1
            GoTo KURIKAESHI     '*** 次項目へ ***
        End If
'
        Plst_yomu           '*** 部品表を読む ***
'
    Else
        If Len(PFLname) < 9 Then
            Pdata_Listp = UCase(PFLname)
        Else
            Pdata_Listp = UCase(Left(PFLname, 8) & Mid(PFLname, 10))
        End If
    End If
'
    Plst_PRNlbl.Caption = " !!! 部品表 " & UCase(PFLname) & " のデータをプリントバッファーに転送中 !!!"
    pps = 1
    Pdata_bangoup = " 1"    '*** 通し番号初期化 ***
'
    For ip = 1 To Anum0
        DRVindex = XCONT0(2) & "\" & AITEM(ip, 0) & "\" & AITEM(ip, 0) & "INDEX.COD"
'
        RDindex         '*** INDEX.COD 読み込み ***
        jps = pps
'
        For jp = 1 To Bnum0
            For pp = jps To Ptotal
                If Mid(PLST(pp, 1), 2, 4) = BINDEX(jp, 0) Then
                    pps = pp
'
                    If FLGinit = 0 Then
                        PRNinit0    '*** プリンタ初期化 ***
                        FLGinit = 1
                    End If
'
                    SET_DRVmain ip, jp
'
                    RDmain  '*** MAIN.COD 読み込み ***
'
                    For kp = 1 To Cnum0
                        Pdata_codep = "L" & BINDEX(jp, 0) & "-" & CMAIN(kp, 0)
                        FLGitti = 0
                        FLGtuzuki = 0
                        Reset_Itti TuzukiDim    '*** 一致メモリー初期化 ***
'
                        For np = pps To Ptotal
                            If PLST(np, 1) = Pdata_codep Then   '*** 個別にﾃﾞｰﾀｰ作成
                                If BINDEX(jp, 5) = "000" Then
                                    Pdata_makerp = CMAIN(kp, 13)
'
                                ElseIf BINDEX(jp, 5) = "998" Then
                                    If PLST(np, 3) = "0" Then
                                        MKsitei = "0"
                                    Else
                                        MKsitei = PLST(np, 3)
                                    End If
'
                                    GET998maker Pdata_makerp, MKsitei
'
                                Else
                                    Pdata_makerp = BINDEX(jp, 5)
                                End If
'
                                Makerget2 Pdata_makerp  '*** メーカー略称取得 ***
'
                                GETkikaku Pdata_kikakup, MKsitei    '*** 部品名取得 ***
'
                                If CMAIN(kp, 16) = "1" Then         '*** 特記事項記入
                                    If PLST(np, 4) = "" Or PLST(np, 4) = "*" Then
                                        '
                                    Else
                                        Pdata_kikakup = Pdata_kikakup & "[" & PLST(np, 4) & "]"
                                    End If
                                End If
                                Pdata_bikoup = PLST(np, 2)
'
                                FLGitti = 0     '*** 一致ﾌﾗｸﾞｸﾘｱｰ
                                For i = 0 To FLGtuzuki              '*** 一致する項目を集計する
                                    If Itti_kikaku(i) = Pdata_kikakup And _
                                                            Itti_bikou(i) = Pdata_bikoup Then
'                                                                   '*** 個数合計 ***
                                        Itti_kosuu(i) = Str(Val(Itti_kosuu(i)) + 1)
                                        FLGitti = 1     '*** 一致ﾌﾗｸﾞｾｯﾄ
                                        Exit For
                                    End If
                                Next i
'
                                If FLGitti = 0 Then
                                    Set_Itti FLGtuzuki
                                    FLGtuzuki = FLGtuzuki + 1
                                End If
                                CMAIN(kp, 10) = Hizuke
                            End If
                        Next np
'
                        If CMAIN(kp, 12) = "0" Then
                            '   非出庫の時は印刷しない。
                        Else
                            If FLGtuzuki > 0 Then
                                For i = 0 To FLGtuzuki - 1      '*** 総数計算 ***
                                    Itti_sousuu(i) = Str(Val(Itti_kosuu(i)) * Val(Baisuu) * Val(Daisuu))
                                Next i
'
                                For i = 0 To FLGtuzuki - 1
                                    Set_Pdata FLGrow, i
                                    FLGrow = FLGrow + 1
'
                                    If FLGrow = 5 Then
                                        PRNkoumoku FLGrow   '*** 部品表項目印刷 ***
                                        FLGrow = 0
                                        Reset_Pdata RowDim
'
                                    End If
'
                                Next i
                            End If
                        End If
                    Next kp
'
                    WRmain      '*** 日付データセーブ ***
                    Exit For    '*** ppの検索終了
'
                End If
            Next pp
        Next jp
'
        FLGtuzuki = 0
        Reset_Itti TuzukiDim    '*** 一致メモリー初期化 ***
'
        For pp = 1 To Ptotal    '*** 未登録部品の選別 ***
            If Left(PLST(pp, 1), 1) = "*" Then
                If PLST(pp, Pdim0 + 1) = AITEM(ip, 3) Or _
                   PLST(pp, Pdim0 + 1) = AITEM(ip, 4) Or _
                   PLST(pp, Pdim0 + 1) = AITEM(ip, 5) Then
'
                    Pdata_codep = "未登録"
                    Pdata_kikakup = Mid(PLST(pp, 1), 2)
                    Pdata_makerp = " ****"
                    Pdata_bikoup = PLST(pp, 2)
'
                    FLGitti = 0     '*** 一致ﾌﾗｸﾞｸﾘｱｰ
                    For i = 0 To FLGtuzuki              '*** 一致する項目を集計する
                        If Itti_kikaku(i) = Pdata_kikakup And _
                                                Itti_bikou(i) = Pdata_bikoup Then
'                                                                   '*** 個数合計 ***
                            Itti_kosuu(i) = Str(Val(Itti_kosuu(i)) + 1)
                            FLGitti = 1     '*** 一致ﾌﾗｸﾞｾｯﾄ
                            Exit For
'
                        End If
                    Next i
'
                    If FLGitti = 0 Then
                        Set_Itti FLGtuzuki
                        FLGtuzuki = FLGtuzuki + 1
                    End If
                End If
            End If
        Next pp
'
        If FLGtuzuki > 0 Then
            For i = 0 To FLGtuzuki - 1      '*** 総数計算 ***
                Itti_sousuu(i) = Str(Val(Itti_kosuu(i)) * Val(Baisuu) * Val(Daisuu))
            Next i
'
            For i = 0 To FLGtuzuki - 1
                Set_Pdata FLGrow, i
                FLGrow = FLGrow + 1
'
                If FLGrow = 5 Then
                    PRNkoumoku FLGrow
                    FLGrow = 0
                    Reset_Pdata RowDim
                End If
            Next i
        End If
'
    Next ip
'
    If FLGall = 1 Then GoTo KURIKAESHI      '*** 構成表による連続印刷 ***
'
OWARI:
    If FLGrow > 0 Then
        PRNkoumoku FLGrow
        FLGrow = 0
        Reset_Pdata RowDim
    End If
'
    Printer.EndDoc      '*** プリンター書き込み ***
'
    Plst_PRNttl.MousePointer = ccDefault    '*** 砂時計解除 ***
    Parts_lst1.Show
    Unload Me
'
errh_P:
    Plst_PRNstd.Caption = STATUS
    Plst_PRNttl.MousePointer = ccDefault    '*** 砂時計解除 ***
'
End Sub

Private Sub cmdQuit_Click()
    Parts_lst1.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    S_Hyouji            '*** 画面初期化＆表示 ***
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 6750
    Height = 4875
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
'
    Plst_PRNstd.Caption = STATUS2
'
    S_Hyouji            '*** 画面初期化＆表示 ***
End Sub

Private Sub Reset_Pdata(Kazu As Integer)
    Dim i As Integer
'
    For i = 0 To Kazu
        Pdata_List(i) = ""
        Pdata_bangou(i) = ""
        Pdata_code(i) = ""
        Pdata_kikaku(i) = ""
        Pdata_kosuu(i) = ""
        Pdata_sousuu(i) = ""
        Pdata_maker(i) = ""
        Pdata_bikou(i) = ""
    Next i
End Sub

Private Sub Reset_Itti(Kazu As Integer)
    Dim i As Integer
'
    For i = 0 To Kazu
        Itti_List(i) = ""
        Itti_Bangou(i) = ""
        Itti_code(i) = ""
        Itti_kikaku(i) = ""
        Itti_kosuu(i) = ""
        Itti_sousuu(i) = ""
        Itti_maker(i) = ""
        Itti_bikou(i) = ""
    Next i
End Sub

Private Sub S_Hyouji()
    If FLGall = 0 Then      '*** １つの部品表印刷 ***
        optdirect.Value = True
        FLGkouseihyou = 0
'
        txtkeisiki.Text = ""
        txtmeisyou.Text = ""
        txttantou.Text = ""
        txtkouban.Text = ""
        txtdaisuu.Text = ""
        Daisuu = "1"
        Baisuu = "1"        '*** 仮数 ***
'
        Plst_yomu           '*** 部品表を読みﾌｧｲﾙ内容表示 ***
'
    Else                    '*** 構成表より連続印刷 ***
        opthyou.Value = True
        optdirect.Enabled = False
        FLGkouseihyou = 1
'
        Klst_yomu       '*** 構成表を読み込み内容を表示 ***
'
        lblnamae.Caption = "印刷 ﾌｧｲﾙ名 ： <<構成表による>>"
        lblkomei.Caption = "小名称 ： "
        lbldate.Caption = "日付 ： "
        txtbaisuu.Text = ""
        lblkomei.Enabled = False
        lbldate.Enabled = False
        txtbaisuu.Enabled = False
        lblbaisuu.Enabled = False
    End If
End Sub

Private Sub Set_Itti(FLGtuzuki As Integer)
    Itti_List(FLGtuzuki) = Pdata_Listp
    Itti_Bangou(FLGtuzuki) = Pdata_bangoup
    If CMAIN(kp, 12) = "1" Or Pdata_codep = "未登録" Then
        Pdata_bangoup = Str(Val(Pdata_bangoup) + 1)
    End If
    Itti_code(FLGtuzuki) = Pdata_codep
    Itti_kikaku(FLGtuzuki) = Pdata_kikakup
    Itti_kosuu(FLGtuzuki) = "1"
    Itti_maker(FLGtuzuki) = Pdata_makerp
    Itti_bikou(FLGtuzuki) = Pdata_bikoup
End Sub

Private Sub Set_Pdata(FLGrow As Integer, i As Integer)
    Pdata_List(FLGrow) = Itti_List(i)
    Pdata_bangou(FLGrow) = Itti_Bangou(i)
    Pdata_code(FLGrow) = Itti_code(i)
    Pdata_kikaku(FLGrow) = Itti_kikaku(i)
    Pdata_kosuu(FLGrow) = Itti_kosuu(i)
    Pdata_maker(FLGrow) = Itti_maker(i)
    Pdata_bikou(FLGrow) = Itti_bikou(i)
    Pdata_sousuu(FLGrow) = Itti_sousuu(i)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Parts_lst1.Show
    End If
End Sub

Private Sub mnuAQuit_Click()
    Unload Me
    End
End Sub

Private Sub mnuBack_Click()
    Eeos_00.Show
    Unload Me
End Sub

Private Sub mnusetumei_Click()
    EeosSetumei.Show 1
End Sub

Private Sub mnuversion_Click()
    EeosVersion.Show 1
End Sub

Private Sub opthyou_Click()
    Klst_yomu       '*** 構成表を読み込み内容表示 ***
End Sub

Private Sub txtbaisuu_Click()
    txtbaisuu.MousePointer = ccIBeam
End Sub

Private Sub txtbaisuu_LostFocus()
    txtbaisuu.MousePointer = ccArrow
    Baisuu = Trim(txtbaisuu.Text)
End Sub

Private Sub txtdaisuu_Click()
    txtdaisuu.MousePointer = ccIBeam
End Sub

Private Sub txtdaisuu_LostFocus()
    txtdaisuu.MousePointer = ccArrow
    Daisuu = Trim(txtdaisuu.Text)
End Sub

Private Sub txtKeisiki_Click()
    txtkeisiki.MousePointer = ccIBeam
End Sub

Private Sub txtKeisiki_LostFocus()
    txtkeisiki.MousePointer = ccArrow
    CATno = Trim(txtkeisiki.Text)
End Sub

Private Sub txtkouban_Click()
    txtkouban.MousePointer = ccIBeam
End Sub

Private Sub txtkouban_LostFocus()
    txtkouban.MousePointer = ccArrow
    Kouban = Trim(txtkouban.Text)
End Sub

Private Sub txtmeisyou_Click()
    txtmeisyou.MousePointer = ccIBeam
End Sub

Private Sub txtmeisyou_LostFocus()
    txtmeisyou.MousePointer = ccArrow
    CATname = Trim(txtmeisyou.Text)
End Sub

Private Sub txttantou_Click()
    txttantou.MousePointer = ccIBeam
End Sub

Private Sub txttantou_LostFocus()
    txttantou.MousePointer = ccArrow
    Person = Trim(txttantou.Text)
End Sub

