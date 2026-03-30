VERSION 5.00
Begin VB.Form Plst_append 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "部品追加"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Plst_append.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox txtTokki 
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
      Height          =   315
      Left            =   6360
      MousePointer    =   1  '矢印
      TabIndex        =   13
      Text            =   "*"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox cbomaker 
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
      Height          =   315
      Left            =   3600
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox cboCode1 
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
      Height          =   315
      Left            =   1440
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   9
      Top             =   1440
      Width           =   7335
   End
   Begin VB.ComboBox cboCode0 
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
      Height          =   315
      Left            =   1440
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1080
      Width           =   7335
   End
   Begin VB.TextBox txtMeisyou0 
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
      Height          =   270
      Left            =   2640
      MousePointer    =   1  '矢印
      TabIndex        =   3
      Text            =   "123456789012345678901234567890"
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox txtCodeno 
      Alignment       =   2  '中央揃え
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
      Height          =   270
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   720
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
      Left            =   4800
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "決定(&Q)"
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
      TabIndex        =   14
      Top             =   2640
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
      Left            =   360
      MousePointer    =   1  '矢印
      TabIndex        =   1
      Text            =   " U12"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblTokki 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "特記事項"
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
      Height          =   315
      Left            =   5400
      MousePointer    =   1  '矢印
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblmaker 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ指定"
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
      Height          =   315
      Left            =   2640
      MousePointer    =   1  '矢印
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblcmain 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品目"
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
      Height          =   315
      Left            =   960
      MousePointer    =   1  '矢印
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblbindex 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種"
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
      Height          =   315
      Left            =   960
      MousePointer    =   1  '矢印
      TabIndex        =   6
      Top             =   1080
      Width           =   495
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
      Left            =   2640
      MousePointer    =   1  '矢印
      TabIndex        =   2
      Top             =   480
      Width           =   5895
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
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   4
      Top             =   480
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
      Left            =   360
      MousePointer    =   1  '矢印
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Plst_append"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 部品表追加画面 ***
'**********************
'
Option Explicit
'
    Private FLGsindo As Integer
    Private FLGtyoufuku As Integer
    Private Pmodu As String, Pitem As String, Pno As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 360 + (9225 - 720) * HyoujiBairitu + 360
    Height = 480 + (4020 - 840) * HyoujiBairitu + 360
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
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    Me.Caption = STATUS & "＜部品追加＞"
'
    txtpno.Text = ""
    txtCodeno.Text = ""
    txtMeisyou0.Text = ""
    FLGesc = 1              '*** ESCフラグ セット ***
    FLGtuika = 0
    FLGsindo = 0
    icpsT = 0
'
    Call Checksindo     '*** 進度チェック ***
End Sub

Private Sub cboCode0_Click()
    Dim Moji As String, temp As String, Dtemp As String
'
    jpsT = cboCode0.ListIndex
    If jpsT = 0 Then
        Gdata3(0) = "未登録部品"
        txtCodeno.Text = Gdata3(0)
        FLGsindo = 4
        Call Checksindo
'
    Else
        FLGsindo = 2
        Call Checksindo
'
        temp = BindexT(jpsT, 0)
        Call GET_jps(temp, Aitem0(), ipsT, BindexT(), jpsT, jcpsT, DRVmainT, CmainT(), CnumT, CdimT, kcpsT)
        cboCode1.Clear
        For kpT = 1 To CnumT
            Moji = "       " & CmainT(kpT, 0) & "   xxx" & CmainT(kpT, 1) & "xxx"
            Call ADD_kuuhaku(Moji, 18)
            Dtemp = CmainT(kpT, 3)
            Call TRSsitei3(Dtemp)
            Moji = Moji & " " & Dtemp
            Call ADD_kuuhaku(Moji, 28)
            cboCode1.AddItem Moji & "  " & CmainT(kpT, 2)
        Next kpT
        kpsT = 1
        cboCode1.ListIndex = kpsT - 1
        Gdata3(0) = "L" & BindexT(jpsT, 0) & "-" & CmainT(kpsT, 0)
        Call SET_cbomaker
    End If
End Sub

Private Sub cboCode1_Click()
'
    kpsT = cboCode1.ListIndex + 1
'
    If CmainT(kpsT, 16) = "1" Then    '*** 特記事項 ***
        lblTokki.Enabled = True
        txtTokki.Enabled = True
        txtTokki.Text = Gdata6(0)
    Else
        lblTokki.Enabled = False
        txtTokki.Enabled = False
    End If
'
    FLGsindo = 3
    Call Checksindo
'
    Gdata3(0) = "L" & BindexT(jpsT, 0) & "-" & CmainT(kpsT, 0)
    txtCodeno.Text = " " & Gdata3(0)
    Call DSPtxtmeisyou0     '*** 名称表示 ***
    Gdata5(0) = "0"
    Gdata6(0) = ""
    cmdQuit.SetFocus
End Sub

Private Sub cbomaker_Change()
    Select Case cbomaker.ListIndex
    Case 0
        Gdata5(0) = "0"
    Case 1
        Gdata5(0) = "8"
    Case 2
        Gdata5(0) = "9"
    Case 3
        Gdata5(0) = "10"
    End Select
End Sub

Private Sub cmdCancel_Click()
    FLGesc = 1      '*** ESCフラグ セット ***
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Dim i As Integer
'
    If FLGsindo < 3 Then
        Beep
        i = MsgBox("まだ部品名称が確定していません。", vbCritical, STATUS)
    Else
'       Gdata3(0) = "L" & BINDEX(jps, 0) & "-" & CMAIN(kps, 0) '*** コード番号 ***
'       Gdata4(0)          '*** 未登録部品は "*" & "部品名"、未登録の時だけデーター入力 ***
'
        If Left(Gdata3(0), 1) <> "L" Then       '*** 未登録部品 ***
            Gdata4(0) = "*" & Trim(txtMeisyou0.Text)
            Gdata3(0) = Gdata4(0)
'
            MEM_ips = ipsT   '*** 今のデータを記憶 ***
            MEM_jps = 0
            MEM_kps = 0
'
        Else                                    '*** 登録部品 ***
            MEM_ips = ipsT   '*** 今のデータを記憶 ***
            MEM_jps = jpsT
            MEM_kps = kpsT
        End If
'
'       Gdata5(0)          '*** 部品指定 ***
'       Gdata6(0)          '*** 特記事項 ***
'
        FLGesc = 0      '*** ESCフラグ リセット ***
        FLGtuika = 1    '*** 追加フラグ セット ***
'
        Unload Me
    End If
End Sub

Private Sub Checksindo()
                            '*** 進度に合わせてテキストボックスを許可する。 ***
    Select Case FLGsindo
    Case 0                  '*** 未入力 ***
        lblCodeno.Enabled = False
        txtCodeno.Enabled = False
        lblMeisyou.Enabled = False
        txtMeisyou0.Enabled = False
        lblbindex.Enabled = False
        cboCode0.Enabled = False
        lblcmain.Enabled = False
        cboCode1.Enabled = False
        lblmaker.Enabled = False
        cbomaker.Enabled = False
        lblTokki.Enabled = False
        txtTokki.Enabled = False
        cmdQuit.Enabled = False
    Case 1                  '*** 部品番号入力完了 ***
        lblCodeno.Enabled = False
        txtCodeno.Enabled = False
        lblMeisyou.Enabled = True
        txtMeisyou0.Enabled = True
        lblbindex.Enabled = True
        cboCode0.Enabled = True
        lblcmain.Enabled = False
        cboCode1.Enabled = False
    Case 2                  '*** 品目名確定 ***
        lblCodeno.Enabled = False
        txtCodeno.Enabled = False
        lblMeisyou.Enabled = True
        txtMeisyou0.Enabled = True
        lblbindex.Enabled = True
        cboCode0.Enabled = True
        lblcmain.Enabled = False
        cboCode1.Enabled = False
    Case 3                  '*** 部品名確定 ***
        lblCodeno.Enabled = True
        txtCodeno.Enabled = True
        lblMeisyou.Enabled = True
        txtMeisyou0.Enabled = True
        lblbindex.Enabled = True
        cboCode0.Enabled = True
        lblcmain.Enabled = True
        cboCode1.Enabled = True
        cmdQuit.Enabled = True
    Case 4                  '*** 未登録部品 ***
        lblCodeno.Enabled = False
        txtCodeno.Enabled = False
        lblMeisyou.Enabled = True
        txtMeisyou0.Enabled = True
        lblbindex.Enabled = True
        cboCode0.Enabled = True
        lblcmain.Enabled = False
        cboCode1.Enabled = False
        cmdQuit.Enabled = True
    End Select
End Sub

Private Sub DSPtxtmeisyou0()
    Dim Pdata As String, Bsitei As String
'
    Bsitei = Gdata5(0)
    jpT = jpsT
    kpT = kpsT
    Call GETkikaku(Pdata, Bsitei, BindexT(), jpT, CmainT(), kpT)
    txtMeisyou0.Text = Pdata
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1      '*** エスケープフラグセット ***
    End If
End Sub

Private Sub SET_cbomaker()
    Dim i As Integer
    Dim Mdata As String
'
    If BindexT(jpsT, 5) = "998" Then
        cbomaker.Enabled = True
        lblmaker.Enabled = True
'
            cbomaker.Clear
            cbomaker.AddItem "複数指"
        For i = 8 To 10
            Mdata = BindexT(jpsT, i)
            Call Makerget2(Mdata)
            cbomaker.AddItem Mdata
        Next i
'
        Select Case Val(Gdata5(0))
        Case 0
            cbomaker.ListIndex = 0
        Case 8
            cbomaker.ListIndex = 1
        Case 9
            cbomaker.ListIndex = 2
        Case 10
            cbomaker.ListIndex = 3
        End Select
    Else
        cbomaker.Enabled = False
        lblmaker.Enabled = False
    End If
End Sub

Private Sub txtMeisyou0_GotFocus()
    If Left(Trim(Gdata3(0)), 1) = "L" Then
        txtMeisyou0.Locked = True
    End If
End Sub

Private Sub txtMeisyou0_LostFocus()
    txtMeisyou0.Locked = False
End Sub

Private Sub txtpno_LostFocus()
    Dim i As Integer
'
    If txtpno.Text = "" Then
        Beep
'       txtpno.SetFocus     '*** 進度操作により他に移行しない ***
    Else
        Call CHKbangou       '*** 部品番号重複チェック
        If FLGtyoufuku = 1 Then
            Beep
            i = MsgBox("部品番号が重複しています。", vbCritical, STATUS)
            txtpno.SetFocus
        Else
            FLGsindo = 1
            Call Checksindo     '*** 進度チェック ***
            Call DSPoya_code
            If Gdata1(0) = "**" Then
                Beep
                i = MsgBox("入力した部品番号はリストに無いので使用できません。", vbCritical, STATUS)
                    FLGsindo = 0
                    Call Checksindo     '*** 進度を戻す ***
                txtpno.SetFocus
                Exit Sub
'
            End If
            cboCode0.SetFocus
        End If
    End If
End Sub

Private Sub DSPoya_code()
                            '*** 親コード表示 ***
    Dim FLGmaker As String
    Dim Moji As String
'
    Gdata1(0) = Pitem
    Call GET_koumoku(Gdata1(0), Aitem0(), Anum0)    '*** 項目名の検索 ***
    If Gdata1(0) = "**" Then
        Exit Sub
    End If
'
                                '*** ip の決定 ***
    Call GET_ips(Gdata1(0), Aitem0(), Anum0, Adim0, ipsT, icpsT, DRVindexT, BindexT(), BnumT, BdimT, _
                    jcpsT, kcpsT)
'
        cboCode0.Clear
        cboCode0.AddItem "未登録部品       (上の欄に直接記入してね！)"
    For jpT = 1 To BnumT                 '*** データー裏表示 ***
        FLGmaker = BindexT(jpT, 5)
        Call Makerget2(FLGmaker)
        Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
        Call ADD_kuuhaku(Moji, 24)
        cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
    Next jpT
'
    If MEM_ips = ipsT Then
        cboCode0.ListIndex = MEM_jps
        cboCode1.ListIndex = MEM_kps - 1
    Else
        cboCode0.ListIndex = 0
    End If
End Sub

Private Sub txtTokki_LostFocus()
    Gdata6(0) = txtTokki.Text
End Sub

Private Sub CHKbangou()
                        '*** 部品番号が重複していないかを調べる ***
    Dim i As Integer
'
    Pitem = Trim(txtpno.Text)
    Call GETsymbol(Pmodu, Pitem, Pno)   '*** 部品番号ヘッダーの部分だけを取り出す。 ***
'
    Gdata2(0) = Pmodu & Pitem & Pno
    txtpno.Text = Gdata2(0)
'
    FLGtyoufuku = 0
    For i = 1 To PtotalT
        If PLSTT(i, 0) = Gdata2(0) Then
            FLGtyoufuku = 1
        End If
    Next i
End Sub

Private Sub DSPgamenBuhin()
    lblPno.Left = 360
    lblPno.Top = 480
    lblPno.FontSize = 10 * HyoujiBairitu
    lblPno.Width = 975 * HyoujiBairitu
    lblPno.Height = 255 * HyoujiBairitu
'
    txtpno.Left = 360
    txtpno.Top = 480 + (720 - 480) * HyoujiBairitu
    txtpno.FontSize = 10 * HyoujiBairitu
    txtpno.Width = 975 * HyoujiBairitu
    txtpno.Height = 270 * HyoujiBairitu
'
    lblCodeno.Left = 360 + (1440 - 360) * HyoujiBairitu
    lblCodeno.Top = 480
    lblCodeno.FontSize = 10 * HyoujiBairitu
    lblCodeno.Width = 1095 * HyoujiBairitu
    lblCodeno.Height = 255 * HyoujiBairitu
'
    txtCodeno.Left = 360 + (1440 - 360) * HyoujiBairitu
    txtCodeno.Top = 480 + (720 - 480) * HyoujiBairitu
    txtCodeno.FontSize = 10 * HyoujiBairitu
    txtCodeno.Width = 1095 * HyoujiBairitu
    txtCodeno.Height = 270 * HyoujiBairitu
'
    lblMeisyou.Left = 360 + (2640 - 360) * HyoujiBairitu
    lblMeisyou.Top = 480
    lblMeisyou.FontSize = 10 * HyoujiBairitu
    lblMeisyou.Width = 5895 * HyoujiBairitu
    lblMeisyou.Height = 255 * HyoujiBairitu
'
    txtMeisyou0.Left = 360 + (2640 - 360) * HyoujiBairitu
    txtMeisyou0.Top = 480 + (720 - 480) * HyoujiBairitu
    txtMeisyou0.FontSize = 10 * HyoujiBairitu
    txtMeisyou0.Width = 5895 * HyoujiBairitu
    txtMeisyou0.Height = 270 * HyoujiBairitu
'
    cboCode0.Left = 360 + (1455 - 360) * HyoujiBairitu
    cboCode0.Top = 480 + (1080 - 480) * HyoujiBairitu
    cboCode0.FontSize = 10 * HyoujiBairitu
    cboCode0.Width = 7335 * HyoujiBairitu
'   cboCode0.Height = 315 * HyoujiBairitu
'
    lblbindex.Left = 360 + (960 - 360) * HyoujiBairitu
    lblbindex.Top = 480 + (1080 - 480) * HyoujiBairitu
    lblbindex.FontSize = 10 * HyoujiBairitu
    lblbindex.Width = 495 * HyoujiBairitu
    lblbindex.Height = cboCode0.Height
'
    cboCode1.Left = 360 + (1455 - 360) * HyoujiBairitu
    cboCode1.Top = 480 + (1440 - 480) * HyoujiBairitu
    cboCode1.FontSize = 10 * HyoujiBairitu
    cboCode1.Width = 7335 * HyoujiBairitu
'   cboCode1.Height = 315 * HyoujiBairitu
'
    lblcmain.Left = 360 + (960 - 360) * HyoujiBairitu
    lblcmain.Top = 480 + (1440 - 480) * HyoujiBairitu
    lblcmain.FontSize = 10 * HyoujiBairitu
    lblcmain.Width = 495 * HyoujiBairitu
    lblcmain.Height = cboCode1.Height
'
    cbomaker.Left = 360 + (3615 - 360) * HyoujiBairitu
    cbomaker.Top = 480 + (1920 - 480) * HyoujiBairitu
    cbomaker.FontSize = 10 * HyoujiBairitu
    cbomaker.Width = 1455 * HyoujiBairitu
'   cbomaker.Height = 315 * HyoujiBairitu
'
    lblmaker.Left = 360 + (2640 - 360) * HyoujiBairitu
    lblmaker.Top = 480 + (1920 - 480) * HyoujiBairitu
    lblmaker.FontSize = 10 * HyoujiBairitu
    lblmaker.Width = 975 * HyoujiBairitu
    lblmaker.Height = cbomaker.Height
'
    txtTokki.Left = 360 + (6375 - 360) * HyoujiBairitu
    txtTokki.Top = 480 + (1920 - 480) * HyoujiBairitu
    txtTokki.FontSize = 10 * HyoujiBairitu
    txtTokki.Width = 2175 * HyoujiBairitu
    txtTokki.Height = 315 * HyoujiBairitu
'
    lblTokki.Left = 360 + (5400 - 360) * HyoujiBairitu
    lblTokki.Top = 480 + (1920 - 480) * HyoujiBairitu
    lblTokki.FontSize = 10 * HyoujiBairitu
    lblTokki.Width = 975 * HyoujiBairitu
    lblTokki.Height = txtTokki.Height
'
    cmdCancel.Left = 360 + (4800 - 360) * HyoujiBairitu
    cmdCancel.Top = 480 + (2640 - 480) * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1095 * HyoujiBairitu
    cmdCancel.Height = 615 * HyoujiBairitu
'
    cmdQuit.Left = 360 + (6360 - 360) * HyoujiBairitu
    cmdQuit.Top = 480 + (2640 - 480) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1095 * HyoujiBairitu
    cmdQuit.Height = 615 * HyoujiBairitu
End Sub
