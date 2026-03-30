VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Pcod_item 
   BackColor       =   &H00004000&
   Caption         =   "部品ｺｰﾄﾞ <項目一覧>"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   1650
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
   Icon            =   "Pcod_item.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6495
   ScaleWidth      =   10095
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
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
      SelectionMode   =   1
      AllowUserResizing=   1
      MousePointer    =   1
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
   Begin VB.CommandButton cmdTanka_new 
      Caption         =   "単価ﾃﾞｰﾀ更新"
      Enabled         =   0   'False
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
      Left            =   2760
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   0
      Text            =   "U,D"
      Top             =   5760
      Width           =   1305
   End
   Begin VB.CommandButton cmdIndex 
      Caption         =   "品種一覧(&N)"
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
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
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
      TabIndex        =   4
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "着目項目"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "ﾌｧｲﾙ(&X)"
      Begin VB.Menu mnuFilePrn 
         Caption         =   "項目一覧印刷(&P)..."
      End
      Begin VB.Menu mnuFileWR 
         Caption         =   "一覧ﾌｧｲﾙ出力(&W)..."
      End
      Begin VB.Menu mnuCodeWR 
         Caption         =   "ｺｰﾄﾞ表ﾌｧｲﾙ出力(&F)..."
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
Attribute VB_Name = "Pcod_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'****************************************
'*** ＥＥＯＳ 電気部品コード 項目一覧 ***
'***        2001.01.24  by S.Fukazawa ***
'****************************************
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
    Private Const kijunX = c10mm * 3
    Private Const kijunY = c10mm * 1
    Private Const gyoukan = 1440 / 4
    Private Const moji_zureX = c10mm / 5
    Private Const moji_zureY = gyoukan / 4
    Private Const Gyoumax = 40            '*** 印刷行最大値 ***
    Private Const haba1X = c10mm * (21 - 4.5)
    Private Const haba1Y = gyoukan * (Gyoumax + 2)
    Private Const pichi1 = c10mm * 0.7
    Private Const pichi2 = c10mm * 3.4
    Private Const pichi3 = c10mm * 4.3
    Private Const pichi4 = c10mm * 5.1
    Private Const pichi5 = c10mm * 5.9
    Private Const pichi6 = c10mm * 7#
'
    Dim Time_up As Boolean

Private Sub Form_Activate()
    FLGitem = 1      '*** 部品コード項目画面存在フラグ設定 ***
'
    Call MENU_settei    '*** メニュー状態設定 ***
'
    Text1.Text = Aitem0(ips0, 1)
    Text1.SetFocus
End Sub

Private Sub Form_Initialize()
    ip0 = 1         '*** 初期値設定 ***
    ips0 = ip0
    icps0 = 0
'
    HeadTitle = "部品ｺｰﾄﾞ <項目一覧>"
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
    Call DSPlevel1      '*** 項目名表示 ***
    FLGitem_data_change = 0     '*** 変更フラグ初期化 ***
'
    If Xcont0(8) = "_" Then
        cmdTanka_new.Enabled = True
        cmdTanka_new.Visible = True
    Else
        cmdTanka_new.Enabled = False
        cmdTanka_new.Visible = False
    End If
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
'                    グリッドの幅の設定
    MSFlexGrid1.Width = 9375 * HyoujiBairitu + FLGoffsetX
    MSFlexGrid1.Font.Size = 10 * HyoujiBairitu
    MSFlexGrid1.RowHeightMin = 245 * HyoujiBairitu  '245
    MSFlexGrid1.Height = MSFlexGrid1.RowHeightMin * 21  '5145?
    MSFlexGrid1.Cols = 7
    MSFlexGrid1.ColWidth(0) = 430 * HyoujiBairitu
    MSFlexGrid1.ColWidth(1) = 540 * HyoujiBairitu
    MSFlexGrid1.ColWidth(2) = 1430 * HyoujiBairitu
    MSFlexGrid1.ColWidth(3) = 540 * HyoujiBairitu
    MSFlexGrid1.ColWidth(4) = 540 * HyoujiBairitu
    MSFlexGrid1.ColWidth(5) = 540 * HyoujiBairitu
    MSFlexGrid1.ColWidth(6) = 5070 * HyoujiBairitu + (300 * HyoujiBairitu - 300) + FLGoffsetX
'
    Call Buhin_Haichi       '*** 表示部品配置 ***
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FLGitem = 0     '*** 部品コード項目画面存在フラグ初期化 ***
    ip0 = 1         '*** フラグ初期化 ***
    ips0 = ip0      '***<再ロード時 "Form_Initialize"前に "MSFlexGrid1_Scroll" が発生してしまう予防>***
'
    If FLGindex = 1 Then
        Unload Pcod_index   '*** 子画面も抹消 ***
    End If
'
    If FLGmain = 1 Then
'        Unload Pcod_main    '*** 子画面も抹消 ***
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdTanka_new_Click()
    Dim Tanka() As String       '*** 単価データ ***
    Dim it As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim temp As String
    Dim Hajimete As Integer     '*** 初めてのときはセーブしない ***
    Dim FILE_number As Integer
    Dim Kiroku As String        '*** 更新記録 ***
'
    Beep
    i = MsgBox("コード表の単価データを「K_Tanka.csv」に従って書き換えます。", vbExclamation Or vbYesNo, HeadTitle)
    If i = vbNo Then Exit Sub
'
    Eeos2_mainMDI.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    MSFlexGrid1.MousePointer = flexHourglass
    DoEvents
'
    Hajimete = 1
    icps0 = 0
    jcps0 = 0
    Kiroku = "L0000-00"
'
    it = 4      '*** データ項目数 ***
                '*** 0: 項目、1:コード番号、2:品名、3:単価、4:入荷年月 ***
    ReDim Tanka(it)
'
    FILE_number = FreeFile      '*** 空いているファイル番号を得る ***
    Open Xcont0(2) & "\2_Tanka\K_Tanka.CSV" For Input As #FILE_number
        Do While EOF(FILE_number) = False
            For i = 0 To it
                Input #FILE_number, temp
                Tanka(i) = Trim(temp)
            Next i
'
            If (Kiroku = Tanka(1)) Then
                GoTo loopin     '*** 同じだから跳ばす ***
            End If
            Kiroku = Tanka(1)
'
            If Left(Tanka(1), 1) = "L" Then
                DoEvents
'
                For i = 1 To Anum0
                    If Tanka(0) = Aitem0(i, 0) Then
                        If icps0 <> i Then
                            Call SET_DRVindex(DRVindex0, Aitem0(), i)
                            Call RDindex(DRVindex0, Bindex0(), Bnum0, Bdim0)
                            icps0 = i
                        End If
                        Exit For
'
                    End If
                Next i

                If i <= Anum0 Then      '*** 該当項目がないときは終わり ***
                    For j = 1 To Bnum0
                        If Mid(Tanka(1), 2, 4) = Bindex0(j, 0) Then
                            If icps0 <> i Then  '*** 初期化 ***
                                icps0 = i
                                jcps0 = 0
                            End If
'
                            If jcps0 <> j Then
                                If Hajimete <> 1 Then
                                    Call WRmain(DRVmain0, Cmain0(), Cnum0, Cdim0)
                                End If
                                Hajimete = 0
'
                                Call SET_DRVmain(DRVmain0, Aitem0(), icps0, Bindex0(), j)
                                Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)
                                jcps0 = j
                            End If
                            Exit For
'
                        End If
                    Next j
'
                    If j <= Bnum0 Then      '*** 該当親コードがないときは終わり ***
                        For k = 1 To Cnum0
                            If Mid(Tanka(1), 7, 2) = Cmain0(k, 0) Then
                                If Left(Tanka(3), 1) = "\" Then
                                    Tanka(3) = Mid(Tanka(3), 2)
                                End If
                                Cmain0(k, 5) = Tanka(3)
                            '   Cmain0(k, 8) = "0"      '*** 在庫数 ***
                                Cmain0(k, 10) = Tanka(4)
                                Me.Caption = HeadTitle & "  " & DRVmain0
                                Exit For
'
                            End If
                        Next k
                    End If
'
                End If
            End If
loopin: Loop
'
        If Hajimete <> 1 Then
            Call WRmain(DRVmain0, Cmain0(), Cnum0, Cdim0)
        End If
    Close #FILE_number
'
    MSFlexGrid1.MousePointer = flexArrow
    Me.MousePointer = vbDefault
    Eeos2_mainMDI.MousePointer = vbDefault
    Me.Caption = HeadTitle
End Sub

Private Sub cmdTanka_newX_Click()
    Dim Tanka() As String       '*** 単価データ ***
    Dim it As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim temp As String
    Dim Hajimete As Integer     '*** 初めてのときはセーブしない ***
    Dim FILE_number As Integer
'
    Beep
    i = MsgBox("コード表の単価データを「K_Tanka.csv」に従って書き換えます。", vbExclamation Or vbYesNo, HeadTitle)
    If i = vbNo Then Exit Sub
'
    Hajimete = 1
    icps0 = 0
    jcps0 = 0
'
    it = 7      '*** データ項目数 ***
                '*** 0:項目、1:コード番号、2:品名、3:項目説明、4:品種説明、5:指定、6:単価、7:メーカー ***
    ReDim Tanka(it)
'
    FILE_number = FreeFile      '*** 空いているファイル番号を得る ***
    Open Xcont0(2) & "\K_Tanka.csv" For Input As #FILE_number
        Do While EOF(FILE_number) = False
            For i = 0 To it
                Input #FILE_number, temp
                Tanka(i) = Trim(temp)
            Next i
'
            For i = 1 To Anum0
                If Tanka(0) = Aitem0(i, 0) Then
                    If icps0 <> i Then
                        Call SET_DRVindex(DRVindex0, Aitem0(), i)
                        Call RDindex(DRVindex0, Bindex0(), Bnum0, Bdim0)
                        icps0 = i
                    End If
                    Exit For
'
                End If
            Next i
'
            If i <= Anum0 Then          '*** 該当項目がないときは終わり ***
                For j = 1 To Bnum0
                    If Mid(Tanka(1), 2, 4) = Bindex0(j, 0) Then
                        If jcps0 <> j Then
                            If Hajimete <> 1 Then
                                Call WRmain(DRVmain0, Cmain0(), Cnum0, Cdim0)
                            End If
                            Hajimete = 0
'
                            Call SET_DRVmain(DRVmain0, Aitem0(), icps0, Bindex0(), j)
                            Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)
                            jcps0 = j
                        End If
                        Exit For
'
                    End If
                Next j
'
                If j <= Bnum0 Then      '*** 該当親コードがないときは終わり ***
                    For k = 1 To Cnum0
                        If Mid(Tanka(1), 7, 2) = Cmain0(k, 0) Then
                            Cmain0(k, 5) = Tanka(6)
                            Cmain0(k, 8) = "0"
                            Me.Caption = HeadTitle & "  " & DRVmain0
                            Exit For
'
                        End If
                    Next k
                End If
'
            End If
        Loop
'
        If Hajimete <> 1 Then
            Call WRmain(DRVmain0, Cmain0(), Cnum0, Cdim0)
        End If
    Close #FILE_number
'
    Me.Caption = HeadTitle
End Sub

Private Sub cmdUp_Click()
    Call Hup
End Sub

Private Sub cmdDown_Click()
    Call Hdown
End Sub

Private Sub DSPlevel1()
                    '*** 項目一覧表示 ***
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Text = "行"
        .ColAlignment(0) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 1
        .Text = "略"
        .ColAlignment(1) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 2
        .Text = "   項 目 名"
        .ColAlignment(2) = flexAlignLeftCenter      '*** 左詰め ***
        .Col = 3
        .Text = "  図"
        .ColAlignment(3) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 4
        .Text = "面 記"
        .ColAlignment(4) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 5
        .Text = "号   "
        .ColAlignment(5) = flexAlignCenterCenter    '*** 中央揃え ***
        .Col = 6
        .Text = "       備   考"
        .ColAlignment(6) = flexAlignLeftCenter      '*** 左詰め ***
    End With
'
    Call DATA_settei
End Sub

Private Sub DATA_settei()
    Dim i As Integer, j As Integer
'
    If Anum0 > 20 Then      '*** 行数設定 ***
        MSFlexGrid1.Rows = Anum0 + 2
    Else
        MSFlexGrid1.Rows = 22
    End If
'
    For i = 1 To Anum0
        MSFlexGrid1.Row = i
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = i
        For j = 0 To 1
            MSFlexGrid1.Col = j + 1
            MSFlexGrid1.Text = Aitem0(i, j)
        Next j
        For j = 3 To 6
            MSFlexGrid1.Col = j
            If j = 6 Then
                MSFlexGrid1.Text = " " & Aitem0(i, j)
            Else
                MSFlexGrid1.Text = Aitem0(i, j)
            End If
        Next j
    Next i
'
    MSFlexGrid1.Row = Anum0 + 1     '*** 予備行クリアー ***
    For j = 0 To 6
        MSFlexGrid1.Col = j
        MSFlexGrid1.Text = ""
    Next j
'
    MSFlexGrid1.TopRow = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub Hup()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Anum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j - 10 < 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 10
    End If
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub Hdown()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Anum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Anum0 > j + 29 Then
        MSFlexGrid1.TopRow = j + 10
    Else
        Beep
        MSFlexGrid1.TopRow = Anum0 - 18
    End If
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub H1up()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Anum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf j = 1 Then
        Beep
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = j - 1
    End If
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub H1down()
    Dim j As Integer
'
    j = MSFlexGrid1.TopRow
'
    If Anum0 <= 20 Then
        Beep
        MSFlexGrid1.TopRow = 1
    ElseIf Anum0 > j + 19 Then
        MSFlexGrid1.TopRow = j + 1
    Else
        Beep
        MSFlexGrid1.TopRow = Anum0 - 18
    End If
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub cmdIndex_Click()
    If FLGindex = 1 Then
        Pcod_index.SetFocus
    Else
        Pcod_index.Show
    End If
End Sub

Private Sub cmdPage_Click()
    FLGitem_data_change = 0     '*** 変更フラグ初期化 ***
'
    Pcod_item_c.Show 1
'                       '*** モーダルフォーム実行後、続いて実行される。***
    If FLGitem_data_change = 1 Then '*** 変更有り ***
        Call DATA_settei
        FLGitem_data_change = 0   '*** 変更フラグ初期化 ***
    End If
'
    Text1.Text = Aitem0(ips0, 1)
    Text1.SetFocus
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

Private Sub mnuFilePrnA_Click()
    Call mnuBuhinItiranhyouPrint
End Sub

Private Sub mnuFilePrnAP_Click()
    Call mnuFilePrnA_Click
End Sub

Private Sub mnuHinsyu_Click()
    Call cmdIndex_Click
End Sub

Private Sub mnuHinsyuP_Click()
    Call cmdIndex_Click
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
    Call mnuConvFile_Click
End Sub

Private Sub mnuPmain_Click()
    If FLGmain = 1 Then
        If icps0 = ips0 And jcps0 = jps0 Then
            Pcod_main.SetFocus
        Else
            Pcod_index.SetFocus
        End If
    Else
'       If FLGindex = 1 Then
'           Pcod_index.SetFocus
'       Else
'           Pcod_index.Show
'       End If
    End If
End Sub

Private Sub mnuPmainP_Click()
    Call mnuPmain_Click
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
    If MSFlexGrid1.Row <= Anum0 Then
        ips0 = MSFlexGrid1.Row
        Text1.Text = Aitem0(ips0, 1)
    End If
'
    Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row <= Anum0 Then
        ips0 = MSFlexGrid1.Row
        Text1.Text = Aitem0(ips0, 1)
'
        Call cmdIndex_Click
'
    Else
        Text1.SetFocus
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** 右ボタン処理 ***
    End If
End Sub

Private Sub MSFlexGrid1_Scroll()
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
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

Private Sub mnuCodeWR_Click()   '*** コード表ファイル出力 ***
    Dim ans As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim CSVfile As String, PDa As String, PD As String
    Dim FILE_number As Integer
'
    ans = MsgBox("全項目・全品種のコード表をファイルに出力します。" & vbCrLf & _
        "フルデータ出力の時は「はい」、従来と同じ時は「いいえ」を押してください。", vbYesNoCancel Or vbInformation)
    If ans = vbYes Or ans = vbNo Then
        On Error GoTo error_F
'
        With CommonDialog1
            .CancelError = True     ' CancelError プロパティを真 (True) に設定します。
            .Flags = cdlOFNHideReadOnly    ' Flags プロパティを設定します。
                                    ' リスト ボックスに表示されるフィルタを設定します。
            .Filter = "CSV形式 ファイル (*.csv)|*.csv|" & _
                    "テキスト ファイル (*.txt)|*.txt|" & _
                    "すべてのファイル (*.*)|*.*|"
            .FilterIndex = 0        ' "CSV形式 ファイル" を既定のフィルタとして指定します。
            .DialogTitle = "部品ｺｰﾄﾞ表 => ファイル出力"
            .ShowOpen               ' [ファイルを開く] ダイアログ ボックスを表示します。
        End With
'
        Eeos2_mainMDI.MousePointer = vbHourglass
        MSFlexGrid1.MousePointer = flexHourglass
        CSVfile = CommonDialog1.FileName
        On Error GoTo 0
'
        DoEvents
'
        FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
        Open CSVfile For Output As #FILE_number
        For i = 1 To Anum0
            Me.Caption = "電気部品コード表 =>  「 " & Aitem0(i, 1) & " 」をファイルに出力中！ あと " & Anum0 - i & "項目です。"
'
            Call SET_DRVindex(DRVindexT, Aitem0(), i)
            Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)
'
            For j = 1 To BnumT
                Call SET_DRVmain(DRVmainT, Aitem0(), i, BindexT(), j)
                Call RDmain(DRVmainT, CmainT(), CnumT, CdimT)
'
                DoEvents    '*** 砂時計をちゃんと表示させるために ***
'
                For k = 1 To CnumT
                    PDa = Chr(34) & Aitem0(i, 0) & Chr(34) & "," _
                        & Chr(34) & "L" & BindexT(j, 0) & "-" & CmainT(k, 0) & Chr(34) & ","
                    If BindexT(j, 4) = "*" Then
                        PD = BindexT(j, 3) & CmainT(k, 1)
                    Else
                        PD = BindexT(j, 3) & CmainT(k, 1) & BindexT(j, 4)
                    End If
                    PDa = PDa & Chr(34) & PD & Chr(34) & "," _
                        & Chr(34) & BindexT(j, 1) & Chr(34) & "," _
                        & Chr(34) & CmainT(k, 2) & Chr(34) & ","
                    PD = CmainT(k, 3)
                        Call TRSsitei2(PD)
                    PDa = PDa & Chr(34) & PD & Chr(34) & "," _
                        & Val(CmainT(k, 5)) & ","
                    If BindexT(j, 5) <> "000" Then
                        PD = BindexT(j, 5)
                        Call Makerget2(PD)
                    Else
                        PD = CmainT(k, 13)
                        Call Makerget2(PD)
                    End If
                    PDa = PDa & Chr(34) & PD & Chr(34) & "," _
                        & Chr(34) & str(Val(CmainT(k, 8))) & Chr(34)
'
                    If ans = vbYes Then
                        PDa = PDa & "," & Chr(34) & CmainT(k, 10) & Chr(34) & ","
'
                        PD = CmainT(k, 12)
                            Call TRSsyukko2(PD)
                        PDa = PDa & Chr(34) & PD & Chr(34) & ","
'
                        PD = CmainT(k, 17)
                            Call TRStouroku2(PD)
                        PDa = PDa & Chr(34) & PD & Chr(34) & ","
'
                        PD = CmainT(k, 18)
                            Call TRSkeijou2(PD)
                        PDa = PDa & Chr(34) & PD & Chr(34)
'
                        PD = CmainT(k, 6)
                            Call TRS_Mlevel(PD)
                        PDa = PDa & Chr(34) & PD & Chr(34)
                    End If
'
                    Print #FILE_number, PDa  '*** データ書き込み ***
                Next k
            Next j
        Next i
        Close #FILE_number
        Eeos2_mainMDI.MousePointer = vbDefault
        MSFlexGrid1.MousePointer = flexArrow
    End If
error_F:
    Me.Caption = HeadTitle
End Sub

Private Sub spc_mnuCodeWR_Click()   '*** コード表ファイル出力 ***
    Dim ans As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
'
    ans = MsgBox("全項目・全品種のコード表の指定部分をクリアー(→ * )します。", vbYesNo Or vbInformation)
    If ans = vbYes Then
        On Error GoTo error_F
'
        Eeos2_mainMDI.MousePointer = vbHourglass
        MSFlexGrid1.MousePointer = flexHourglass
        On Error GoTo 0
'
        DoEvents
'
        For i = 1 To Anum0
            Me.Caption = "電気部品コード表 =>  「 " & Aitem0(i, 1) & " 」を処理中！ あと " & Anum0 - i & "項目です。"
'
            Call SET_DRVindex(DRVindexT, Aitem0(), i)
            Call RDindex(DRVindexT, BindexT(), BnumT, BdimT)
'
            For j = 1 To BnumT
                Call SET_DRVmain(DRVmainT, Aitem0(), i, BindexT(), j)
                Call RDmain(DRVmainT, CmainT(), CnumT, CdimT)
'
                DoEvents    '*** 砂時計をちゃんと表示させるために ***
'
                For k = 1 To CnumT
                    CmainT(k, 6) = "*"      '*** 指定部分をクリアーする ***
                Next k
'
                Call WRmain(DRVmainT, CmainT(), CnumT, CdimT)
            Next j
        Next i
        Eeos2_mainMDI.MousePointer = vbDefault
        MSFlexGrid1.MousePointer = flexArrow
    End If
error_F:
    Me.Caption = HeadTitle
End Sub

Private Sub mnuFilePrn_Click()  '*** 項目一覧印刷 ***
    Dim i As Integer
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
    FLGpage = 1             '*** ページフラグ初期化
    DoEvents                '*** 画面書き直し ***
'
    Call PRNheader  '*** ヘッダー印刷 (FLGgyou = 0)***
    DoEvents
'
    For i = 1 To Anum0
        If FLGgyou >= Gyoumax Then
            Call PRNfooter
            Printer.EndDoc  '*** 改ページ ***
'
            Call PRNheader
        End If
'
        Call PRNkoumoku(i)  '*** 項目印刷 ***
        DoEvents
'
    Next i
'
    Call PRNfooter          '*** 部品表項末印刷 ***
    Printer.EndDoc      '*** プリンター書き込み ***
'
    Call timer_waite(1000)  '*** 表示認識待ち ***
'
    Me.Caption = HeadTitle
    Me.MousePointer = vbDefault             '*** 砂時計解除 ***
    MSFlexGrid1.MousePointer = flexArrow
    Eeos2_mainMDI.MousePointer = vbDefault
End Sub

Private Sub mnuFileWR_Click()   '*** 項目一覧ファイル出力 ***
    Dim i As Integer
    Dim CSVfile As String
    Dim PDa As String
    Dim PD As String
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
    CSVfile = CommonDialog1.FileName
    On Error GoTo 0
    DoEvents
'
    Me.Caption = HeadTitle & "  =>  ファイルに出力中！"
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open CSVfile For Output As #FILE_number
    For i = 1 To Anum0
        DoEvents
'
        PDa = Chr(34) & Aitem0(i, 0) & Chr(34) & "," _
            & Chr(34) & Aitem0(i, 1) & Chr(34) & "," _
            & Chr(34) & Aitem0(i, 2) & Chr(34) & "," _
            & Chr(34) & Aitem0(i, 3) & Chr(34) & "," _
            & Chr(34) & Aitem0(i, 4) & Chr(34) & "," _
            & Chr(34) & Aitem0(i, 5) & Chr(34) & "," _
            & Chr(34) & Aitem0(i, 6) & Chr(34)
'
        Print #FILE_number, PDa  '*** データ書き込み ***
    Next i
    Close #FILE_number
'
    Call timer_waite(1000)  '*** 表示認識待ち ***
'
error_F:
    Me.Caption = HeadTitle
    Eeos2_mainMDI.MousePointer = vbDefault  '*** 砂時計解除 ***
    MSFlexGrid1.MousePointer = flexArrow
End Sub

Private Sub mnuJumpC_Click()
    Dim Tyuusin As Integer
'
    Tyuusin = Anum0 \ 2
    If Tyuusin <= 10 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Tyuusin - 9
    End If
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub mnuJumpE_Click()
    If Anum0 < 20 Then
        MSFlexGrid1.TopRow = 1
    Else
        MSFlexGrid1.TopRow = Anum0 - 18
    End If
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
End Sub

Private Sub mnuJumpT_Click()
    MSFlexGrid1.TopRow = 1
'
    ip0 = MSFlexGrid1.TopRow
    ips0 = ip0
    Text1.Text = Aitem0(ips0, 1)
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
    Printer.CurrentX = kijunX + c10mm
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    Call SETfont_size(9, 1)
    Printer.Print "M5103-01";
'
    Printer.CurrentX = kijunX + haba1X - c10mm * 4
    Printer.CurrentY = kijunY + haba1Y + moji_zureY
    Call SETfont_size(9, 0)
    Printer.Print FormatDateTime(Date, vbLongDate) & " 印刷";
End Sub

Private Sub PRNkoumoku(i As Integer)
'                   *** 項目印刷 ***
    Printer.CurrentX = kijunX + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 0)
'
    Printer.CurrentX = kijunX + pichi1 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 1)
'
    Printer.CurrentX = kijunX + pichi2 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 2)
'
    Printer.CurrentX = kijunX + pichi3 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 3)
'
    Printer.CurrentX = kijunX + pichi4 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 4)
'
    Printer.CurrentX = kijunX + pichi5 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 5)
'
    Printer.CurrentX = kijunX + pichi6 + moji_zureX
    Printer.CurrentY = kijunY + gyoukan * (FLGgyou + 2) + moji_zureY
    Printer.Print Aitem0(i, 6)
'
    FLGgyou = FLGgyou + 1
End Sub

Private Sub PRNheader()
'                   *** 項目一覧表ヘッダー印刷 ***
    Dim i As Integer
'
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
    Printer.CurrentY = kijunY + moji_zureY
    Call SETfont_size(10, 1)        '*** フォント,サイズ設定 ***
    Printer.Print "===  [  営電標準  ]           電気部品コード・単価表             ＜項目名一覧表＞  ==="
'
    Printer.CurrentX = kijunX + c10mm + moji_zureX
    Printer.CurrentY = kijunY + gyoukan + moji_zureY
    SETfont_size 9, 1      '*** フォント,サイズ設定 ***
    Printer.Print "項目名        DSK    図面記号                             備 考"
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
    Text1.FontSize = 10 * HyoujiBairitu
    Text1.Height = 285 * HyoujiBairitu
    Text1.Left = 360 + (1330 - 360) * HyoujiBairitu
    Text1.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    Text1.Width = 1305 * HyoujiBairitu
'
    Label1.FontSize = 10 * HyoujiBairitu
    Label1.Height = Text1.Height
    Label1.Left = 360
    Label1.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    Label1.Width = 975 * HyoujiBairitu
'
    cmdTanka_new.FontSize = 9 * HyoujiBairitu
    cmdTanka_new.Left = 360 + (2670 - 360) * HyoujiBairitu
    cmdTanka_new.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdTanka_new.Width = 1215 * HyoujiBairitu
    cmdTanka_new.Height = 495 * HyoujiBairitu
'
    cmdPage.FontSize = 9 * HyoujiBairitu
    cmdPage.Left = 360 + (4080 - 360) * HyoujiBairitu
    cmdPage.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdPage.Width = 1095 * HyoujiBairitu
    cmdPage.Height = 495 * HyoujiBairitu
'
    cmdIndex.FontSize = 9 * HyoujiBairitu
    cmdIndex.Left = 360 + (5400 - 360) * HyoujiBairitu
    cmdIndex.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdIndex.Width = 1095 * HyoujiBairitu
    cmdIndex.Height = 495 * HyoujiBairitu
'
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Left = 360 + (6720 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdUp.Width = 735 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Left = 360 + (7560 - 360) * HyoujiBairitu
    cmdDown.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdDown.Width = 735 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdClose.FontSize = 9 * HyoujiBairitu
    cmdClose.Left = 360 + (8520 - 360) * HyoujiBairitu
    cmdClose.Top = 360 + (5760 - 360) * HyoujiBairitu + FLGoffsetY
    cmdClose.Width = 855 * HyoujiBairitu
    cmdClose.Height = 495 * HyoujiBairitu
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
'       Me.mnuCodeP.Checked = True  '*** 自分なので存在しない ***
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
'
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
        Me.mnuHinsyuP.Checked = False
        Me.mnuHinsyuP.Enabled = False
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
        Left = 0    'Eeos2_mainMDI.ScaleWidth - Width
    Else
        Left = 0
    End If
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


