VERSION 5.00
Begin VB.Form Plst_main_c 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "部品変更"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Plst_main_c.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Parts_lst5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.ComboBox cboCode2 
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
      ItemData        =   "Plst_main_c.frx":030A
      Left            =   7680
      List            =   "Plst_main_c.frx":030C
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdOnaji 
      Caption         =   "一つ前と同じ(&F)"
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
      Left            =   5160
      TabIndex        =   19
      ToolTipText     =   "一つ前に行った変更と同じ部品に変更します。"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSubete 
      Caption         =   "すべて変更(&A)"
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
      Left            =   3240
      TabIndex        =   18
      ToolTipText     =   "他の番号の部品も一度に変更します。"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtSuisyou 
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
      Height          =   270
      Left            =   360
      MousePointer    =   1  '矢印
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   2160
      Width           =   1095
   End
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
      Width           =   8055
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
      ItemData        =   "Plst_main_c.frx":030E
      Left            =   1440
      List            =   "Plst_main_c.frx":0310
      MousePointer    =   1  '矢印
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1080
      Width           =   8055
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
      TabIndex        =   5
      Text            =   "123456789012345678901234567890"
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox txtCodeno 
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
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "L1234-56"
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
      Left            =   1440
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "変更実行(&Q)"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
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
      Locked          =   -1  'True
      MousePointer    =   1  '矢印
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   " U12"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblSuisyou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "推奨品ｺｰﾄﾞ"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      MousePointer    =   1  '矢印
      TabIndex        =   16
      ToolTipText     =   "このコード番号の部品に変更した方が良いですよ。"
      Top             =   1920
      Width           =   1095
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
      TabIndex        =   4
      Top             =   480
      Width           =   4935
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
      TabIndex        =   2
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
Attribute VB_Name = "Plst_main_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**********************
'*** 部品表変更画面 ***
'**********************
'
Option Explicit
'
    Private Fcbo0 As Integer
    Private PRTindex() As String
'    Private PRTcmain() As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
'                   フォームの表示位置の設定
    Width = 360 + (9960 - 720) * HyoujiBairitu + 360
    Height = 480 + (4095 - 840) * HyoujiBairitu + 360
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
    Me.Caption = STATUS & "＜部品変更＞"
'
    txtpno.Text = Gdata2(0)
    If Left(Gdata3(0), 1) = "L" Then
        txtCodeno.Text = " " & Gdata3(0)
    Else
        txtCodeno.Text = Gdata3(0)
    End If
'
    txtMeisyou0.Text = Gdata4(0)
    txtSuisyou.Text = ""
'
    If Gdata1(1) = Gdata1(0) Then
        cmdOnaji.Enabled = True
    Else
        cmdOnaji.Enabled = False
    End If
'
    Fcbo0 = 0   '*** コンボボックス誤動作防止 ***
    FLGsubete = 0
'
    Call DSP_gamen0
End Sub

Private Sub cboCode2_Click()
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim temp As String
'
    If Fcbo0 = 0 Then Exit Sub  '*** コンボボックスの誤動作防止 ***
'
    If cboCode2.ListIndex = 1 Then
        For j = 1 To BnumT - 1
            For i = 1 To BnumT - 1
                k = Val(PRTindex(i, 0))
                m = Val(PRTindex(i + 1, 0))
                If Mid(PRTindex(m, 1), 12, 5) < Mid(PRTindex(k, 1), 12, 5) Then
                    temp = PRTindex(i, 0)
                    PRTindex(i, 0) = PRTindex(i + 1, 0)
                    PRTindex(i + 1, 0) = temp
                End If
            Next i
        Next j
    Else
        For i = 1 To BnumT
            PRTindex(i, 0) = str(i)
        Next i
    End If
'
        cboCode0.Clear
        cboCode0.AddItem "            未登録部品"
    For i = 1 To BnumT
        j = Val(PRTindex(i, 0))
        cboCode0.AddItem PRTindex(j, 1)
    Next i
'
    For i = 1 To BnumT
        If jpsT = Val(PRTindex(i, 0)) Then
            j = i
            Exit For
        End If
    Next i
'
    If i = BnumT + 1 Then j = 0
    cboCode0.ListIndex = j
'
'    cboCode2.Visible = False
End Sub

Private Sub cboCode0_Click()
    Dim Moji As String, Dtemp As String
    Dim FLGbcod As String
'
    If Fcbo0 = 0 Then Exit Sub  '*** コンボボックスの誤動作防止 ***
'
    jpsT = cboCode0.ListIndex
    If jpsT = 0 Then
        cboCode1.Enabled = False
        lblcmain.Enabled = False
        Gdata3(0) = "未登録部品"
    Else
        cboCode1.Enabled = True
        lblcmain.Enabled = True
        jpsT = Val(PRTindex(jpsT, 0))
'
        FLGbcod = BindexT(jpsT, 0)
        Call GET_jps(FLGbcod, Aitem0(), ipsT, BindexT(), jpsT, jcpsT, DRVmainT, CmainT(), CnumT, CdimT, kcpsT)
        cboCode1.Clear
        For kpT = 1 To CnumT
            Moji = "       " & CmainT(kpT, 0) & "   xxx" & CmainT(kpT, 1) & "xxx"
            Call ADD_kuuhaku(Moji, 18)
            Dtemp = CmainT(kpT, 3)
            Call TRSsitei3(Dtemp)
            Moji = Moji & " " & Dtemp
            Call ADD_kuuhaku(Moji, 28)
            cboCode1.AddItem Moji & " " & CmainT(kpT, 2)
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
        Gdata6(0) = "*"
    End If
'
    Gdata3(0) = "L" & BindexT(jpsT, 0) & "-" & CmainT(kpsT, 0)
End Sub

Private Sub cbomaker_Click()
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

Private Sub cmdOnaji_Click()
    Gdata1(0) = Gdata1(1)   '*** 前のデータを復活 ***
    Gdata2(0) = Gdata2(1)
    Gdata3(0) = Gdata3(1)
    Gdata4(0) = Gdata4(1)
    Gdata5(0) = Gdata5(1)
    Gdata6(0) = Gdata6(1)
    Gdata7(0) = Gdata7(1)
'
    FLGesc = 0      '*** ESCフラグ リセット ***
    FLGsubete = 0
    Unload Me
End Sub

Private Sub cmdQuit_Click()
'   Gdata3(0) = "L" & BINDEX(jps, 0) & "-" & CMAIN(kps, 0) '*** コード番号 ***
'   Gdata4(0)          '*** 未登録部品は "*" & "部品名"
    If Left(Gdata3(0), 1) <> "L" Then
        Gdata4(0) = "*" & Trim(txtMeisyou0.Text)
        Gdata3(0) = Gdata4(0)
    End If
'   Gdata5(0)          '*** 部品指定 ***
'   Gdata6(0)          '*** 特記事項 ***
'
    Gdata1(1) = Gdata1(0)   '*** 前のデータとして保存 ***
    Gdata2(1) = Gdata2(0)
    Gdata3(1) = Gdata3(0)
    Gdata4(1) = Gdata4(0)
    Gdata5(1) = Gdata5(0)
    Gdata6(1) = Gdata6(0)
    Gdata7(1) = Gdata7(0)
'
    FLGesc = 0      '*** ESCフラグ リセット ***
    FLGsubete = 0
    Unload Me
End Sub

Private Sub cmdSubete_Click()
    Dim Indata As String
    Dim i As Integer
'
'   Gdata3(0) = "L" & BINDEX(jps, 0) & "-" & CMAIN(kps, 0) '*** コード番号 ***
'   Gdata4(0)          '*** 未登録部品は "*" & "部品名"
    If Left(Gdata3(0), 1) <> "L" Then
        Gdata4(0) = "*" & Trim(txtMeisyou0.Text)
        Gdata3(0) = Gdata4(0)
    End If
'   Gdata5(0)          '*** 部品指定 ***
'   Gdata6(0)          '*** 特記事項 ***
'
    FLGesc = 0      '*** ESCフラグ リセット ***
'
    Indata = InputBox("備考欄を入力してください。" & vbCrLf & vbCrLf & "何も書くことがないときも「*」を入れておいて下さい。", STATUS & " <部品変更>", Gdata7(0), (Screen.Width - 6000) \ 2, (Screen.Height - 1000) \ 2)
    If Indata = "" Then
        Exit Sub
'
    Else
        Gdata7(0) = Indata
    End If
'
    i = MsgBox("同じコード番号すべてを一括で変更しますがよろしいですね！", vbExclamation Or vbOKCancel, STATUS & " <部品変更>")
    If i = vbOK Then
        FLGsubete = 1
    Else
        FLGsubete = 0
    End If
'
    Unload Me
End Sub

Private Sub DSP_gamen0()
    Dim Tmpcode0 As String, Tmpcode1 As String
    Dim FLGmaker As String
    Dim Moji As String, Dtemp As String
'
    If Left(Trim(Gdata3(0)), 1) = "L" Then     '*** 登録部品 ***
        Tmpcode0 = Mid(Trim(Gdata3(0)), 2, 4)  '*** コード番号分離　***
        Tmpcode1 = Right(Trim(Gdata3(0)), 2)
'
        Dtemp = Trim(Gdata1(0))                 '*** 部品名称特定 ***
        Call GET_ips(Dtemp, Aitem0(), Anum0, Adim0, ipsT, icpsT, DRVindexT, BindexT(), BnumT, BdimT, jcpsT, kcpsT)
'
'       Dtemp = Tmpcode0
        Call GET_jps(Tmpcode0, Aitem0(), ipsT, BindexT(), jpsT, jcpsT, DRVmainT, CmainT(), CnumT, CdimT, kcpsT)
'
        Call GET_kps(Tmpcode1, CmainT(), CnumT, kpsT, kcpsT)
'
'                        ************ データー表示
        ReDim PRTindex(BnumT, 1)
'
            cboCode0.AddItem "            未登録部品"
        For jpT = 1 To BnumT
            FLGmaker = BindexT(jpT, 5)
            Call Makerget2(FLGmaker)
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
'
            PRTindex(jpT, 0) = str(jpT)
            PRTindex(jpT, 1) = Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
        Next jpT
        cboCode0.ListIndex = jpsT
'
'        ReDim PRTcmain(CnumT, 2)
        For kpT = 1 To CnumT
            Moji = "       " & CmainT(kpT, 0) & "   xxx" & CmainT(kpT, 1) & "xxx"
            Call ADD_kuuhaku(Moji, 18)
            Dtemp = CmainT(kpT, 3)
            Call TRSsitei3(Dtemp)
            Moji = Moji & " " & Dtemp
            Call ADD_kuuhaku(Moji, 28)
            cboCode1.AddItem Moji & " " & CmainT(kpT, 2)
'
'            PRTcmain(kpT, 0) = str(kpT)
'            PRTcmain(kpT, 1) = str(kpT)
'            PRTcmain(kpT, 2) = Moji & " " & CmainT(kpT, 2)
        Next kpT
        cboCode1.ListIndex = kpsT - 1
'
        Fcbo0 = 1
        SET_cbomaker
'
        If Val(CmainT(kpsT, 16)) = 0 Then
            lblTokki.Enabled = False
            txtTokki.Enabled = False
        Else
            lblTokki.Enabled = True
            txtTokki.Enabled = True
            txtTokki.Text = Gdata6(0)      '*** 特記事項 ***
        End If
'
        Select Case CmainT(kpsT, 3)
        Case "*", "1", "2"
            lblSuisyou.Enabled = False
            txtSuisyou.Enabled = False
            lblSuisyou.Caption = "推奨品ｺｰﾄﾞ"
            txtSuisyou.Text = ""
        Case "0"
            lblSuisyou.Enabled = True
            txtSuisyou.Enabled = True
            lblSuisyou.Caption = "代品コード"
            txtSuisyou.Text = CmainT(kpsT, 4)
        Case "3", "4"
            lblSuisyou.Enabled = True
            txtSuisyou.Enabled = True
            lblSuisyou.Caption = "推奨品ｺｰﾄﾞ"
            txtSuisyou.Text = CmainT(kpsT, 4)
        End Select
'
    Else            '*** 未登録部品 ***
        Dtemp = Trim(Gdata1(0))                 '*** 部品項目取得 ***
        Call GET_ips(Dtemp, Aitem0(), Anum0, Adim0, ipsT, icpsT, DRVindexT, BindexT(), BnumT, BdimT, jcpsT, kcpsT)
'
'                                               **** データー表示
        ReDim PRTindex(BnumT, 1)
'
            cboCode0.AddItem "未登録部品       (修正をするときは上の欄を直接直してね！)"
        For jpT = 1 To BnumT
            Moji = " L" & BindexT(jpT, 0) & "-     " & BindexT(jpT, 3) & "xxxxx" & BindexT(jpT, 4)
            Call ADD_kuuhaku(Moji, 24)
            cboCode0.AddItem Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
'
            PRTindex(jpT, 0) = str(jpT)
            PRTindex(jpT, 1) = Moji & " " & FLGmaker & "  " & BindexT(jpT, 1)
        Next jpT
        cboCode0.ListIndex = 0
'
        cboCode1.Enabled = False
        lblcmain.Enabled = False
        cbomaker.Enabled = False
        lblmaker.Enabled = False
        lblTokki.Enabled = False
        txtTokki.Enabled = False
        lblSuisyou.Enabled = False
        txtSuisyou.Enabled = False
        Fcbo0 = 1
    End If
'
'    cboCode2.Visible = True
    cboCode2.Clear
    cboCode2.AddItem " 品種コード順"    '*** 0 ***
    cboCode2.AddItem " 品種部品名順"    '*** 1 ***
    Fcbo0 = 0
    cboCode2.ListIndex = 0
    Fcbo0 = 1
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
    If Left(Trim(Gdata3(0)), 1) = "L" Then  '*** 未登録の時だけデーター入力 ***
        txtMeisyou0.Locked = True
    End If
End Sub

Private Sub txtMeisyou0_LostFocus()
    txtMeisyou0.Locked = False
End Sub

Private Sub txtTokki_LostFocus()
    Gdata6(0) = txtTokki.Text
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
    lblMeisyou.Width = 4935 * HyoujiBairitu
    lblMeisyou.Height = 255 * HyoujiBairitu
'
    txtMeisyou0.Left = 360 + (2640 - 360) * HyoujiBairitu
    txtMeisyou0.Top = 480 + (720 - 480) * HyoujiBairitu
    txtMeisyou0.FontSize = 10 * HyoujiBairitu
    txtMeisyou0.Width = 4935 * HyoujiBairitu
    txtMeisyou0.Height = 270 * HyoujiBairitu
'
    cboCode0.Left = 360 + (1455 - 360) * HyoujiBairitu
    cboCode0.Top = 480 + (1080 - 480) * HyoujiBairitu
    cboCode0.FontSize = 10 * HyoujiBairitu
    cboCode0.Width = 8055 * HyoujiBairitu
'   cboCode0.Height = 315 * HyoujiBairitu
'
    cboCode2.Top = 480 + (720 - 480) * HyoujiBairitu
    cboCode2.FontSize = 10 * HyoujiBairitu
    cboCode2.Width = 1815 * HyoujiBairitu
    cboCode2.Left = cboCode0.Left + cboCode0.Width - cboCode2.Width
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
    cboCode1.Width = 8055 * HyoujiBairitu
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
    lblSuisyou.Left = 360
    lblSuisyou.Top = 480 + (1920 - 480) * HyoujiBairitu
    lblSuisyou.FontSize = 10 * HyoujiBairitu
    lblSuisyou.Width = 1095 * HyoujiBairitu
    lblSuisyou.Height = 240 * HyoujiBairitu
'
    txtSuisyou.Left = 360
    txtSuisyou.Top = 480 + (2160 - 480) * HyoujiBairitu
    txtSuisyou.FontSize = 10 * HyoujiBairitu
    txtSuisyou.Width = 1095 * HyoujiBairitu
    txtSuisyou.Height = 270 * HyoujiBairitu
'
    cmdCancel.Left = 360 + (1440 - 360) * HyoujiBairitu
    cmdCancel.Top = 480 + (2640 - 480) * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1335 * HyoujiBairitu
    cmdCancel.Height = 615 * HyoujiBairitu
'
    cmdSubete.Left = 360 + (3240 - 360) * HyoujiBairitu
    cmdSubete.Top = 480 + (2640 - 480) * HyoujiBairitu
    cmdSubete.FontSize = 10 * HyoujiBairitu
    cmdSubete.Width = 1455 * HyoujiBairitu
    cmdSubete.Height = 615 * HyoujiBairitu
'
    cmdOnaji.Left = 360 + (5160 - 360) * HyoujiBairitu
    cmdOnaji.Top = 480 + (2640 - 480) * HyoujiBairitu
    cmdOnaji.FontSize = 10 * HyoujiBairitu
    cmdOnaji.Width = 1455 * HyoujiBairitu
    cmdOnaji.Height = 615 * HyoujiBairitu
'
    cmdQuit.Left = 360 + (7080 - 360) * HyoujiBairitu
    cmdQuit.Top = 480 + (2640 - 480) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1455 * HyoujiBairitu
    cmdQuit.Height = 615 * HyoujiBairitu
End Sub

