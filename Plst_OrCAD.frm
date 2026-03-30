VERSION 5.00
Begin VB.Form Plst_OrCAD 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "OrCADからの変換"
   ClientHeight    =   3975
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9615
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Plst_OrCAD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMeisyou0 
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
      Height          =   270
      Left            =   3000
      MousePointer    =   1  '矢印
      TabIndex        =   17
      Text            =   "123456789012345678901234567890"
      Top             =   2280
      Width           =   6135
   End
   Begin VB.TextBox txtCodeno 
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
      Height          =   270
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtpno 
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
      Height          =   270
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   13
      Text            =   " U12"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   3000
   End
   Begin VB.TextBox txtKiji 
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
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "123456789012345678901234567890"
      Top             =   840
      Width           =   7935
   End
   Begin VB.TextBox txtDate 
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
      Height          =   285
      Left            =   7200
      TabIndex        =   7
      Text            =   "1997/03/31"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtKmeisyou 
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
      Height          =   285
      Left            =   3720
      MousePointer    =   1  '矢印
      TabIndex        =   5
      Text            =   "1234567890123456789012345"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtZuban 
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
      Height          =   285
      Left            =   1200
      MousePointer    =   1  '矢印
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "A1234-001AR12"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "決定(&F)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "中止(&C)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      ToolTipText     =   "作業を中断します。作成データは記憶されます。"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblMeisyou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部 品 名 称"
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
      Height          =   255
      Left            =   3000
      MousePointer    =   1  '矢印
      TabIndex        =   16
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Label lblCodeno 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ｺｰﾄﾞ番号"
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
      Height          =   255
      Left            =   1680
      MousePointer    =   1  '矢印
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblPno 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部品番号"
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
      Height          =   255
      Left            =   480
      MousePointer    =   1  '矢印
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblSindbar 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1320
      Width           =   7935
   End
   Begin VB.Label lblSindo 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "進度 ："
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
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblKiji 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "記事 ："
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
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblDate 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "日付 ："
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
      Left            =   6600
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblKmeisyou 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "小名称 ："
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
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblZuban 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "図番 ："
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
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Plst_OrCAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'***********************
'*** OrCADからの変換 ***
'***********************
'
Option Explicit
'
Dim HeadTitle As String
'
Private FLGsindo As Integer
Private FLGworknew As Integer
Private FLGwork_read As Integer
Private CHaiki As String
Private Pmodu As String, Pitem As String, Pno As String
Private Fpdata As Integer, Fpmax As Integer
'
Private Pdata() As String       '*** 一時エリア ***

Private Sub Form_Activate()
    FLGjob = 2
    FLGlevel = 2    '*** 部品表 OrCADﾃﾞｰﾀの変換 ***
    STATUS = "OrCADﾃﾞｰﾀの変換"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'    If FLGplst = 1 And FLGplst2 = 1 Then    '*** 部品表２画面とも既に開いている ***
'        Unload Me
'    End If
'                   フォームの表示位置の設定
    Width = 480 + (9705 - 960) * HyoujiBairitu + 480
    Height = 480 + (4350 - 960) * HyoujiBairitu + 480
'
    Left = (Eeos2_mainMDI.ScaleWidth - Width) \ 2
    Top = (Eeos2_mainMDI.ScaleHeight - Height) \ 2 - 480 - (1800 - 480) * HyoujiBairitu
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    Me.Caption = STATUS & "  X" & Mid(PFLnameT, 2) & " を読み込み中 !!!"
    Me.MousePointer = vbHourglass
    DoEvents
'
    txtZuban.Text = PFLnameT
    txtKmeisyou.Text = ""
    txtDate.Text = ""
    txtKiji.Text = ""
    lblSindbar.Caption = ""
    txtpno.Text = ""
    txtCodeNo.Text = ""
    txtMeisyou0.Text = ""
    FLGesc = 0
    FLGworknew = 0
    FLGwork_read = 0
    FLGsindo = 0
    FLGoption = 0
    icpsT = 0
    PdimT = cPdim0
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
    Timer1.Enabled = True   '*** タイマー動作開始 ***
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
'
    Me.MousePointer = vbDefault
    If FLGworknew = 1 Then
        i = MsgBox("今回の作成データを上書き保存しますか？", vbQuestion Or vbYesNo, STATUS)
        If i = vbYes Then
            Me.MousePointer = vbHourglass
            DoEvents
            Call WRplstWork '*** 部品表作成データセーブ ***
        End If
    ElseIf FLGwork_read = 1 Then
        Me.MousePointer = vbHourglass
        DoEvents
        Call WRplstWork     '*** 部品表作成データセーブ ***
    End If
'
    Me.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdFix_Click()
    Dim i As Integer
'
    If FLGworknew = 1 Then
        i = MsgBox("今回の作成データを上書き保存しますか？", vbQuestion Or vbYesNo, STATUS)
        If i = vbYes Then
            Me.MousePointer = vbHourglass
            Call WRplstWork  '*** 部品表作成データセーブ ***
        End If
    Else
        Me.MousePointer = vbHourglass
        Call WRplstWork      '*** 部品表作成データセーブ ***
    End If
'
    Kill DRVcadplst '*** OrCADからのデータを削除 ***
'
    PlistnameT = Trim(txtKmeisyou.Text)
    PlistdateT = Trim(txtDate.Text)
    RemarksT = Trim(txtKiji.Text)
        If RemarksT = "" Then RemarksT = "*"
'
    Me.MousePointer = vbHourglass
    DoEvents
'
    Call WRpartlist(DRVpartlistT, PlistnameT, PlistdateT, RemarksT, PLSTT(), PtotalT, PdimT)   '*** 部品表データセーブ ***
'
'    PFLname = PFLnameT
'    DRVpartlist = DRVpartlistT
'
    If FLGplst = 0 Then
        Plst_main.Show      '*** 標準部品表へ渡す ***
    ElseIf FLGplst2 = 0 Then
        Plst_main2.Show
    End If
'
    Me.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False  '*** タイマー動作停止 ***
'
    If FLGesc = 1 Then
        Call cmdCancel_Click     '*** 作業中断、終了 ***
    End If
'
    Select Case FLGsindo
    Case 0
        Call RD_CadPlst      '*** OrCAD部品表の読み込み ***
        Timer1.Enabled = True   '*** タイマー動作開始 ***
        Exit Sub
'
    Case 1
        Call Bunseki         '*** 内容分析 ***
            If FLGesc = 1 Then
                FLGsindo = 99   '*** 作業中断、終了 ***
            End If
'
        Timer1.Enabled = True   '*** タイマー動作開始 ***
        Exit Sub
'
    Case 2
        Call Kensaku         '*** 部品割り当て ***
            If FLGesc = 1 Then
                FLGsindo = 99   '*** 作業中断、終了 ***
            End If
'
        Timer1.Enabled = True   '*** タイマー動作開始 ***
        Exit Sub
'
    Case Else
'
    End Select
End Sub

Private Sub DSPtxtmeisyou0()
    Dim Pdata As String, Bsitei As String
'
    Bsitei = Gdata5(0)
    jpT = jps0
    kpT = kps0
    Call GETkikaku(Pdata, Bsitei, BindexT(), jpT, CmainT(), kpT)
    txtMeisyou0.Text = Pdata
End Sub

Private Sub RD_CadPlst()
                        '*** OrCAD部品表を読み込む ***
    Dim i As Integer, j As Integer, k As Integer
    Dim Ptemp As String, Pdata3 As String
'
    Fpdata = 1      '*** 初期値 ***
    Fpmax = 200     '*** 最大値 ***
    ReDim Pdata(Fpmax)
'
    Open DRVcadplst For Input As #1
        Line Input #1, Ptemp
'
        Call CONVtab(Ptemp)        '*** TAB => " " ***
'
        i = InStr(1, Ptemp, "Revised")
        If i <> 0 Then
            PlistnameT = Trim(Left(Ptemp, i - 1))    '*** 小名称 ***
'
            PlistdateT = Trim(Mid(Ptemp, i + 8))     '*** 日付 ***
            i = InStr(1, PlistdateT, ",")
            j = InStr(i + 1, PlistdateT, ",")
            If j <> 0 Then                          '*** Windows版は ","が２つある ***
                PlistdateT = Trim(Mid(PlistdateT, i + 1))
            End If
        Else
            PlistnameT = "不明"                     '*** 小名称 ***
            PlistdateT = Format(Date, "yy/mm/dd")   '*** 日付 ***
        End If
'
        Line Input #1, Ptemp                    '*** データー空読み <４行>
        Line Input #1, Ptemp                    '*** データー空読み
        Line Input #1, Ptemp                    '*** データー空読み
        Line Input #1, Ptemp                    '*** データー空読み
'
        Do While EOF(1) = False
            Line Input #1, Ptemp                '*** データー読み込み
'
            Call CONVtab(Ptemp)        '*** TAB => " ",前後の空白削除 ***
'
            If Ptemp = "" Then                  '*** 改行だけ ***
                '
            ElseIf InStr(1, Ptemp, "Revised") <> 0 Then '*** 部品データ行では無い ***
                '
            ElseIf InStr(1, Ptemp, "Revision") <> 0 Then '*** 部品データ行では無い ***
                '
            ElseIf InStr(1, Ptemp, "Bill") <> 0 Then '*** 部品データ行では無い ***
                '
            ElseIf InStr(1, Ptemp, "Item") <> 0 Then '*** 部品データ行では無い ***
                '
            ElseIf InStr(1, Ptemp, "---") <> 0 Then '*** 部品データ行では無い ***
                '
            ElseIf InStr(1, Ptemp, "___") <> 0 Then '*** 部品データ行では無い ***
                '
            Else
                If Fpdata = Fpmax Then
                    Call UP_Pdata        '*** Pdataディメンジョン増加 ***
                End If
'
                Pdata(Fpdata) = Ptemp
                Fpdata = Fpdata + 1
            End If
        Loop
    Close #1        '*** ファイルの読み終わり ***
'
    If Pdata(Fpdata) = "" Then Fpdata = Fpdata - 1
'
    PtotalT = 0
    For i = 1 To Fpdata     '*** 部品数合計 ***
        j = InStr(1, Pdata(i), " ")
        If j <> 0 Then      '*** 「行番号 数量 部品番号 名称」or「部品番号」or「部品番号 名称」 ***
            Pdata3 = Trim(Mid(Pdata(i), j + 1))
            k = InStr(1, Pdata3, " ")     '*** ２つ目の空白は数量と部品番号の境界 ***
            If k <> 0 Then
                PtotalT = PtotalT + Val(Trim(Mid(Pdata3, 1, k - 1)))
            End If
        End If
    Next i
'
    txtKmeisyou.Text = PlistnameT        '*** 途中経過表示 ***
    txtDate.Text = PlistdateT
    Me.Caption = STATUS & " 只今 " & PFLnameT & "を分析中 !!!"
    FLGesc = 0
    FLGsindo = 1
End Sub

Private Sub CONVtab(Ptemp As String)    '*** TAB => " " ***
    Dim i As Integer, j As Integer
'
    j = 1
    i = InStr(j, Ptemp, vbTab)
    Do While i <> 0
        Ptemp = Left(Ptemp, i - 1) & " " & Mid(Ptemp, i + 1)
        j = i + 1
        i = InStr(j, Ptemp, vbTab)
    Loop
    Ptemp = Trim(Ptemp)             '*** 前後の空白を削除 ***
End Sub

Private Sub Bunseki()
                        '*** ここから内容分析 ***
    Dim i As Integer, j As Integer, k As Integer, k0 As Integer, L As Integer, m As Integer
    Dim p As Integer, q As Integer, r As Integer
    Dim fplst As Integer, pf1 As Integer, Nflag As Integer
    Dim Pdata3 As String, Pdata2 As String, Pdata1 As String
    Dim Pname As String, plstno0 As String, plstno1 As String, Tempno As String
'
    ReDim PLSTT(PtotalT, PdimT + 3)
    fplst = 1       '*** 部品表カウントフラグ ***
'
    For i = 1 To Fpdata
        p = InStr(1, Pdata(i), " ")
        If p <> 0 Then      '*** 「行番号 数量 部品番号 名称」or「部品番号 名称」のブロック ***
            Pdata3 = Trim(Mid(Pdata(i), p + 1))
            q = InStr(1, Pdata3, " ")           '*** ２つ目の空白は数量と部品番号の境界 ***
            If q <> 0 Then
                Pdata2 = Trim(Mid(Pdata3, q + 1))
                r = InStr(1, Pdata2, " ")       '*** ３つ目の空白は部品番号と部品名称の境界 ***
'
                plstno0 = Trim(Mid(Pdata2, 1, r - 1))   '*** この部品番号群を抽出する ***
                Pname = Trim(Mid(Pdata2, r + 1))        '*** 部品名称(Pname)抽出 ***
            Else            '*** 「部品番号 名称」のブロック ***
                plstno0 = Trim(Mid(Pdata(i), 1, r - 1))   '*** この部品番号(群)を抽出する ***
                'Pname =<* 部品名称は前のデータを使う/先頭に来ることは無い *>
            End If
        Else                '*** 「部品番号」だけのブロック ***
                plstno0 = Pdata(i)
                'Pname =<* 部品名称は前のデータを使う/先頭に来ることは無い *>
        End If
'
        L = 0       '*** 部品番号群を分解する ***
        pf1 = Len(plstno0)
        For k = 1 To pf1         '*** 左から見て","の位置で部品番号の境を見つける ***
            plstno1 = Mid(plstno0, k, 1)
            If plstno1 = "," Then
                Call SetPdata(plstno0, fplst, k, L, Pname)
                L = 0
            Else
                If k = pf1 Then   '*** カンマのない行末 ***
                    Call SetPdata(plstno0, fplst, k + 1, L + 1, Pname)
                Else
                    L = L + 1       '*** まだ区切りにならない ***
                End If
            End If
        Next k
    Next i
'
    For i = 1 To PtotalT         '*** 部品表データ整備 ***
        Pitem = PLSTT(i, 0)
        Call GETsymbol(Pmodu, Pitem, Pno)
        PLSTT(i, PdimT + 1) = Pitem         '*** シンボル U,Q etc
        PLSTT(i, PdimT + 2) = Pno           '*** 番号 1,2,1A etc
        PLSTT(i, PdimT + 3) = Pmodu         '*** モジュール番号 A1,A2 etc
    Next i
'
    k = 1           '*** 1 から Ptotalまで並べ換える ***
    k0 = 1
    Nflag = 0
    For i = 1 To Anum0
        For j = 3 To 5
            Pitem = Aitem0(i, j)
            If Pitem <> "*" Then
                Nflag = 0
                For L = k To PtotalT
                    If PLSTT(L, PdimT + 1) = Pitem Then
                        Call NarabeKae1(k, L, Nflag)
                    End If
                Next L
                If Nflag = 1 Then
                    If k <> k0 + 1 Then
                        For L = k0 To k - 2
                            For m = L + 1 To k - 1
                                If PLSTT(L, PdimT + 3) = "" And PLSTT(m, PdimT + 3) = "" Then
                                    If Val(PLSTT(L, PdimT + 2)) > Val(PLSTT(m, PdimT + 2)) Then
                                        Call NarabeKae2(L, m)
                                    Else
                                        If Val(PLSTT(L, PdimT + 2)) < Val(PLSTT(m, PdimT + 2)) Then
                                            '
                                        Else
                                            If PLSTT(L, PdimT + 2) > PLSTT(m, PdimT + 2) Then
                                                Call NarabeKae2(L, m)
                                            End If
                                        End If
                                    End If
                                ElseIf PLSTT(L, PdimT + 3) = "" And PLSTT(m, PdimT + 3) <> "" Then
                                    '
                                Else
                                    If Val(Mid(PLSTT(L, PdimT + 3), 2)) > Val(Mid(PLSTT(m, PdimT + 3), 2)) Then
                                        Call NarabeKae2(L, m)
'
                                    ElseIf Val(Mid(PLSTT(L, PdimT + 3), 2)) = Val(Mid(PLSTT(m, PdimT + 3), 2)) Then
                                        If Val(PLSTT(L, PdimT + 2)) < Val(PLSTT(m, PdimT + 2)) Then
                                            '
                                        Else
                                            If PLSTT(L, PdimT + 2) > PLSTT(m, PdimT + 2) Then
                                                Call NarabeKae2(L, m)
                                            End If
                                        End If
                                    End If
                                End If
                            Next m
                        Next L
                    End If
                    k0 = k
                End If
            End If
        Next j
    Next i
'
    Me.MousePointer = vbDefault
    Tempno = ""         '*** 部品番号の重複検査 ***
    For i = 1 To PtotalT
        If Tempno = PLSTT(i, 0) Then
            j = MsgBox("部品番号 " & Tempno & " が重複しています。" & vbCrLf & "先に進めますか？", vbCritical Or vbYesNo, STATUS)
            If j = vbNo Then
                FLGesc = 1
                Exit Sub
'
            End If
        Else
            Tempno = PLSTT(i, 0)
        End If
'
        If Right(PLSTT(i, 0), 1) = "?" Then
            j = MsgBox("部品番号に " & PLSTT(i, 0) & " があります。" & vbCrLf & "先に進めますか？", vbCritical Or vbYesNo, STATUS)
            If j = vbNo Then
                FLGesc = 1
                Exit Sub
'
            End If
        End If
    Next i
'
    Me.Caption = STATUS
    FLGsindo = 2
    FLGesc = 0
End Sub

Private Sub Kensaku()
    Dim i As Integer, j As Integer, Ks As Integer, Ke As Integer, Ku As Integer, L As Integer
    Dim m As Integer, n As Integer
'
    i = MsgBox("以前の部品表作成データを使用しますか？", vbQuestion Or vbYesNo, STATUS)
    If i = vbYes Then
        Me.MousePointer = vbHourglass
        Me.Caption = STATUS & " 部品表作成データ読み込み中 !!!"
'
        Call RDplstWork          '*** 部品表作成データ読み込み ***
'
        FLGworknew = 0
        FLGwork_read = 1
        Me.Caption = STATUS
    Else
        ReDim PlstWork(cPLSTWORKmax, cPLSTWORKdim)  '*** 部品表作成データ初期化 ***
        FLGworknew = 1
        FLGwork_read = 0
    End If
    Me.MousePointer = vbDefault
'
    Ks = 1              '*** 部品コードの検索 ***
    Ke = 0
    For i = 1 To Anum0
        If Ke = PtotalT Then Exit For
'
        Ku = 0      '*** 該当フラグ ***
        For j = 3 To 5
            Pitem = Aitem0(i, j)
            If Pitem = "*" Then Exit For
'
            For L = Ks To PtotalT
                If PLSTT(L, PdimT + 1) = Pitem Then
                    Ke = L
                    Ku = 1
                End If
            Next L
        Next j
'
        If Ku = 1 Then
            ipsT = i
            ipT = i
            Call SET_DRVindex(DRVindexT, Aitem0(), ipT)
            Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)
'
            jpsT = 1
            jcpsT = 0
            For L = Ks To Ke
                For m = 0 To cPLSTWORKmax
                    If PLSTT(L, 1) = PlstWork(m, 1) And PLSTT(L, PdimT + 1) = PlstWork(m, 4) Then
                        Call OnajiMono(L, m)
                        GoTo tugino_L
'
                    End If
                Next m
'
                For m = cPLSTWORKmax To 1 Step -1   '*** 新名称を先頭に入れるので後ろにずらす ***
                    For n = 0 To cPLSTWORKdim
                        PlstWork(m, n) = PlstWork(m - 1, n)
                    Next n
                Next m
'
                PlstWork(0, 1) = PLSTT(L, 1)         '*** 回路図上の部品名称 ***
'
                mps = L
                Plst_Select.Show 1          '*** 品種・品目選択 ***
'
                If FLGesc = 1 Then        '*** 作業中止 ***
                    For n = 0 To cPLSTWORKdim
                        PlstWork(0, n) = ""
                    Next n
                    Exit Sub        '*** キャンセル終了 ***
'
                Else
                    PlstWork(0, 0) = PLSTT(L, 1)         '*** Lxxxx-xx ***
                   'PlstWork(0, 1) = PLSTt(L, 1)         '*** 部品名称(代入済み) ***
                    PlstWork(0, 2) = PLSTT(L, 2)         '*** 備考 ***
                    PlstWork(0, 3) = PLSTT(L, 3)         '*** メーカー指定 ***
                    PlstWork(0, 4) = PLSTT(L, PdimT + 1) '*** 部品項目記号 ***
'
                    txtpno.Text = PLSTT(L, 0)
                    txtCodeNo.Text = PLSTT(L, 1)
                    txtMeisyou0.Text = PlstWork(0, 1)
                End If
tugino_L:
                Call Hyouji_sindo(L)       '*** 進度バー設定 ***
            Next L
            Ks = Ke + 1
        End If
    Next i
'
    txtpno.Text = ""
    txtCodeNo.Text = ""
    txtMeisyou0.Text = ""
    Call Hyouji_sindo(PtotalT + 1)          '*** 進度バー終了設定 ***
'
    FLGsindo = 3
    FLGesc = 0
End Sub

Private Sub UP_Pdata()
                    '*** 配列を増やす ***
    Dim i As Integer, j As Integer
    Dim CADtemp() As String
'
    ReDim CADtemp(Fpmax)
    For i = 1 To Fpmax
        CADtemp(i) = Pdata(i)
    Next i
'
    ReDim Pdata(Fpmax + 100)
    For i = 1 To Fpmax
        Pdata(i) = CADtemp(i)
    Next i
'
    Fpmax = Fpmax + 100
End Sub

Private Sub SetPdata(plstno0 As String, j As Integer, k As Integer, L As Integer, Pname As String)
    PLSTT(j, 0) = Mid(plstno0, k - L, L)     '*** 部品番号
    PLSTT(j, 1) = Pname                      '*** 部品名
    PLSTT(j, 2) = "*"                        '*** 備考
    PLSTT(j, 3) = "0"                        '*** 部品指定しない
    PLSTT(j, 4) = "*"                        '*** 特記事項
    j = j + 1
End Sub

Private Sub NarabeKae1(ksub As Integer, Lsub As Integer, Nflag As Integer)
    Dim n As Integer, temp As String
'
    If ksub <> Lsub Then
        For n = 0 To PdimT + 3
            temp = PLSTT(ksub, n)
            PLSTT(ksub, n) = PLSTT(Lsub, n)
            PLSTT(Lsub, n) = temp
        Next n
    End If
'
    ksub = ksub + 1
    Nflag = 1
End Sub

Private Sub NarabeKae2(Lsub As Integer, Msub As Integer)
    Dim n As Integer, temp As String
'
    For n = 0 To PdimT + 3
        temp = PLSTT(Lsub, n)
        PLSTT(Lsub, n) = PLSTT(Msub, n)
        PLSTT(Msub, n) = temp
    Next n
End Sub

Private Sub OnajiMono(Lsub As Integer, Msub As Integer)
    Dim Indata As String
    Dim Tmptop As Integer, Tmpleft As Integer
'
    Tmptop = Me.Top + txtMeisyou0.Top + 960
    Tmpleft = Me.Left + txtMeisyou0.Left
'
    PLSTT(Lsub, 1) = PlstWork(Msub, 0)   '*** コード番号 ***
'
    txtpno.Text = PLSTT(Lsub, 0)
    txtCodeNo.Text = PLSTT(Lsub, 1)
    txtMeisyou0.Text = PlstWork(Msub, 1)
'
    If PlstWork(Msub, 2) <> "*" Then
        Beep
'        Indata = InputBox("備考欄に記入する文字を入力してください。", STATUS, "*", Tmpleft, Tmptop)
        Indata = InputBox("備考欄に記入する文字を入力してください。", STATUS, PlstWork(Msub, 2), Tmpleft, Tmptop)
        PlstWork(Msub, 2) = Trim(Indata)
    End If
    PLSTT(Lsub, 2) = PlstWork(Msub, 2)   '*** 備考 ***
'
    PLSTT(Lsub, 3) = PlstWork(Msub, 3)
'
    If PLSTT(Lsub, 4) <> "*" Then
        Beep
'        Indata = InputBox("特記事項に記入する文字を入力してください。", STATUS, "*", Tmpleft, Tmptop)
        Indata = InputBox("特記事項に記入する文字を入力してください。", STATUS, PLSTT(Lsub, 4), Tmpleft, Tmptop)
        PLSTT(Lsub, 4) = Indata
    End If
'
    Call CHGplstWork(Msub)
'
End Sub

Private Sub Hyouji_sindo(Lsub As Integer)      '*** 進度バー設定 ***
    Dim i As Integer, j As Integer
    Dim Xdata As String
'
    If Lsub = PtotalT + 1 Then
        lblSindbar.BackColor = &HC000&
        Xdata = "  >>> 部品割り当ては終了しました。記事欄を記入したら「決定(F)」を押してください。 <<<"
    Else
        i = (CDbl(Lsub) * 40) / CDbl(PtotalT)
        Xdata = ""
        For j = 1 To i
            Xdata = Xdata & "■"
        Next j
    End If
    lblSindbar.Caption = Xdata
End Sub

Private Sub DSPgamenBuhin()
    txtZuban.Left = 480 + (1215 - 480) * HyoujiBairitu
    txtZuban.Top = 480
    txtZuban.FontSize = 10 * HyoujiBairitu
    txtZuban.Width = 1575 * HyoujiBairitu
    txtZuban.Height = 285 * HyoujiBairitu
'
    lblZuban.Left = 480
    lblZuban.Top = 480
    lblZuban.FontSize = 10 * HyoujiBairitu
    lblZuban.Width = 735 * HyoujiBairitu
    lblZuban.Height = txtZuban.Height
'
    txtKmeisyou.Left = 480 + (3735 - 480) * HyoujiBairitu
    txtKmeisyou.Top = 480
    txtKmeisyou.FontSize = 10 * HyoujiBairitu
    txtKmeisyou.Width = 2775 * HyoujiBairitu
    txtKmeisyou.Height = 285 * HyoujiBairitu
'
    lblKmeisyou.Left = 480 + (2880 - 480) * HyoujiBairitu
    lblKmeisyou.Top = 480
    lblKmeisyou.FontSize = 10 * HyoujiBairitu
    lblKmeisyou.Width = 855 * HyoujiBairitu
    lblKmeisyou.Height = txtKmeisyou.Height
'
    txtDate.Left = 480 + (7215 - 480) * HyoujiBairitu
    txtDate.Top = 480
    txtDate.FontSize = 10 * HyoujiBairitu
    txtDate.Width = 1935 * HyoujiBairitu
    txtDate.Height = 285 * HyoujiBairitu
'
    lblDate.Left = 480 + (6600 - 480) * HyoujiBairitu
    lblDate.Top = 480
    lblDate.FontSize = 10 * HyoujiBairitu
    lblDate.Width = 615 * HyoujiBairitu
    lblDate.Height = txtDate.Height
'
    txtKiji.Left = 480 + (1215 - 480) * HyoujiBairitu
    txtKiji.Top = 480 + (840 - 480) * HyoujiBairitu
    txtKiji.FontSize = 10 * HyoujiBairitu
    txtKiji.Width = 7935 * HyoujiBairitu
    txtKiji.Height = 285 * HyoujiBairitu
'
    lblKiji.Left = 480
    lblKiji.Top = 480 + (840 - 480) * HyoujiBairitu
    lblKiji.FontSize = 10 * HyoujiBairitu
    lblKiji.Width = 735 * HyoujiBairitu
    lblKiji.Height = txtKiji.Height
'
    lblSindbar.Left = 480 + (1215 - 480) * HyoujiBairitu
    lblSindbar.Top = 480 + (1320 - 480) * HyoujiBairitu
    lblSindbar.FontSize = 10 * HyoujiBairitu
    lblSindbar.Width = 7935 * HyoujiBairitu
    lblSindbar.Height = 285 * HyoujiBairitu
'
    lblSindo.Left = 480
    lblSindo.Top = 480 + (1320 - 480) * HyoujiBairitu
    lblSindo.FontSize = 10 * HyoujiBairitu
    lblSindo.Width = 735 * HyoujiBairitu
    lblSindo.Height = 285 * HyoujiBairitu
'
    lblPno.Left = 480
    lblPno.Top = 480 + (2040 - 480) * HyoujiBairitu
    lblPno.FontSize = 10 * HyoujiBairitu
    lblPno.Width = 975 * HyoujiBairitu
    lblPno.Height = 255 * HyoujiBairitu
'
    txtpno.Left = 480
    txtpno.Top = 480 + (2280 - 480) * HyoujiBairitu
    txtpno.FontSize = 10 * HyoujiBairitu
    txtpno.Width = 975 * HyoujiBairitu
    txtpno.Height = 270 * HyoujiBairitu
'
    lblCodeno.Left = 480 + (1680 - 480) * HyoujiBairitu
    lblCodeno.Top = 480 + (2040 - 480) * HyoujiBairitu
    lblCodeno.FontSize = 10 * HyoujiBairitu
    lblCodeno.Width = 1095 * HyoujiBairitu
    lblCodeno.Height = 255 * HyoujiBairitu
'
    txtCodeNo.Left = 480 + (1680 - 480) * HyoujiBairitu
    txtCodeNo.Top = 480 + (2280 - 480) * HyoujiBairitu
    txtCodeNo.FontSize = 10 * HyoujiBairitu
    txtCodeNo.Width = 1095 * HyoujiBairitu
    txtCodeNo.Height = 270 * HyoujiBairitu
'
    lblMeisyou.Left = 480 + (3000 - 480) * HyoujiBairitu
    lblMeisyou.Top = 480 + (2040 - 480) * HyoujiBairitu
    lblMeisyou.FontSize = 10 * HyoujiBairitu
    lblMeisyou.Width = 6135 * HyoujiBairitu
    lblMeisyou.Height = 255 * HyoujiBairitu
'
    txtMeisyou0.Left = 480 + (3000 - 480) * HyoujiBairitu
    txtMeisyou0.Top = 480 + (2280 - 480) * HyoujiBairitu
    txtMeisyou0.FontSize = 10 * HyoujiBairitu
    txtMeisyou0.Width = 6135 * HyoujiBairitu
    txtMeisyou0.Height = 270 * HyoujiBairitu
'
    cmdCancel.Left = 480 + (6240 - 480) * HyoujiBairitu
    cmdCancel.Top = 480 + (2880 - 480) * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1095 * HyoujiBairitu
    cmdCancel.Height = 615 * HyoujiBairitu
'
    cmdFix.Left = 480 + (7920 - 480) * HyoujiBairitu
    cmdFix.Top = 480 + (2880 - 480) * HyoujiBairitu
    cmdFix.FontSize = 10 * HyoujiBairitu
    cmdFix.Width = 1095 * HyoujiBairitu
    cmdFix.Height = 615 * HyoujiBairitu
End Sub


