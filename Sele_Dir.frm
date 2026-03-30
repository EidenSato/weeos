VERSION 5.00
Begin VB.Form Sele_Dir 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ｺｰﾄﾞ表 #0"
   ClientHeight    =   5085
   ClientLeft      =   3045
   ClientTop       =   2550
   ClientWidth     =   7590
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Sele_Dir.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5085
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFolder 
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
      Height          =   525
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   6615
   End
   Begin VB.CommandButton cmdStd 
      Caption         =   "標準(&S)"
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
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "環境設定で指定したﾌｫﾙﾀﾞ"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "中止(&Q)"
      Default         =   -1  'True
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
      Left            =   5880
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "決定(&G)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2190
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   5175
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label lblHyoudai 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00004000&
      Caption         =   "***"
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
      Height          =   435
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label lbl_Folder 
      BackColor       =   &H00004000&
      Caption         =   "ﾌｧｲﾙの場所 ："
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
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Sele_Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'******************
'*  Dir 設定 画面 *
'******************
'
Option Explicit
'
    Dim FLGstartP As Integer
    Dim tempDIR As String

Private Sub Form_Load()
    FLGstartP = 1       '*** スタートイニシャライズ中セット ***
'
    Width = 480 + (7680 - 960) * HyoujiBairitu + 480
    Height = 480 + (5440 - 840) * HyoujiBairitu + 360
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
    Call GamenSettei    '*** 画面部品諸元指定 ***
'
    Me.Caption = STATUS
'
    Select Case FLGjob
    Case 1:
        lblHyoudai.Caption = "<< 構成表/部品表が収容されている個別ﾌｫﾙﾀﾞｰ(Partlist)を指定して下さい。>>"
'
        txtFolder.Text = Xcont0(3)
        TMPdir1 = Xcont0(3)
        tempDIR = TMPdir1
'
    Case 2:
        If FLGlevel = 2 Then    '*** 部品表 OrCADﾃﾞｰﾀの変換 ***
            lblHyoudai.Caption = "<< 変換した部品表を収容するﾌｫﾙﾀﾞｰ(Partlist)を指定して下さい。>>" & vbCrLf _
                                & "   OrCADで作成した部品表は構成表と同じ親ﾌｫﾙﾀﾞｰの(Cad_plst)に入れてください。"
'
            txtFolder.Text = Xcont0(3)
            TMPdir2 = Xcont0(3)
            tempDIR = TMPdir2
        Else
            lblHyoudai.Caption = "<< 部品表が収容されている共有ﾌｫﾙﾀﾞｰ(Partlist)を指定して下さい。>>"
'
            txtFolder.Text = Xcont0(4)
            TMPdir2 = Xcont0(4)
            tempDIR = TMPdir2
        End If
    End Select
'
    On Error GoTo NoDir
'                       '*** データセットしただけで「Drive1_Change」が発生する。 ***
    Drive1.Drive = tempDIR
'                       '*** データセットしただけで「Dir1_Change」が発生する。 ***
    Dir1.Path = tempDIR
'
    FLGstartP = 0       '*** スタートイニシャライズ終了セット ***
    FLGesc = 0
    Exit Sub
'
NoDir:
    Resume Next
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
'
    FLGstartP = 0       '*** スタートイニシャライズ終了セット ***
    FLGesc = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        FLGesc = 1
    End If
End Sub

Private Sub cmdQuit_Click()
    FLGesc = 1
    Unload Me
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer
'
    tempDIR = txtFolder.Text
'
    i = InStr(1, UCase(txtFolder.Text), "CAD_PLST")
    If (i <> 0) Then
        Beep
        i = MsgBox("ﾌｫﾙﾀﾞｰの指定が間違っています。", vbCritical, STATUS)
        Exit Sub
    End If
'
    i = InStr(1, UCase(txtFolder.Text), "PARTLIST")
    Select Case FLGjob
    Case 1: '構成表/CAD_PLST
        If (i <> 0) Then
            TMPdir1 = Left$(tempDIR, i - 1) & "PARTLIST"
            TMPdir3 = Left$(tempDIR, i - 1) & "CAD_PLST"
        Else
            TMPdir1 = tempDIR & "\PARTLIST"
            TMPdir3 = tempDIR & "\CAD_PLST"
        End If
'
    Case 2: '部品表
        If (i <> 0) Then
            TMPdir2 = Left$(tempDIR, i - 1) & "PARTLIST"
        Else
            TMPdir2 = tempDIR & "\PARTLIST"
        End If
    End Select
'
    FLGesc = 0
    Unload Me
End Sub

Private Sub cmdStd_Click()
    Select Case FLGjob
    Case 1:
        txtFolder.Text = Xcont0(3)
        tempDIR = Xcont0(3)
    Case 2:
        txtFolder.Text = Xcont0(4)
        tempDIR = Xcont0(4)
    End Select
End Sub

Private Sub Dir1_Change()
    If FLGstartP = 1 Then Exit Sub  '*** 初めての「Form_Load」の時は無視する ***
'
    Call Dir1_Click
End Sub

Private Sub Dir1_Click()
    Dim Moji As String
    Dim n As Integer
'               *** カレントパスの取得 ***
    Moji = Dir1.Path
    n = Len(Moji)
    If Right$(Moji, 1) = "\" Then
        tempDIR = Left$(Moji, n - 1)
    Else
        tempDIR = Moji
    End If
'
    txtFolder.Text = tempDIR
End Sub

Private Sub Drive1_Change()
    If FLGstartP = 1 Then Exit Sub  '*** 初めての「Form_Load」の時は無視する ***
'               *** カレントドライブの取得 ***
    Dir1.Path = Drive1.Drive
End Sub

Private Sub GamenSettei()
    lblHyoudai.Left = 360
    lblHyoudai.Top = 360
    lblHyoudai.FontSize = 10 * HyoujiBairitu
    lblHyoudai.Width = 6855 * HyoujiBairitu
    lblHyoudai.Height = 435 * HyoujiBairitu
'
    lbl_folder.Left = 480
    lbl_folder.Top = 360 + (960 - 360) * HyoujiBairitu
    lbl_folder.FontSize = 9 * HyoujiBairitu
    lbl_folder.Width = 1455 * HyoujiBairitu
    lbl_folder.Height = 195 * HyoujiBairitu
'
    txtFolder.Left = 480
    txtFolder.Top = 360 + (1200 - 360) * HyoujiBairitu
    txtFolder.FontSize = 10 * HyoujiBairitu
    txtFolder.Width = 6615 * HyoujiBairitu
    txtFolder.Height = 525 * HyoujiBairitu
'
    Drive1.Left = 480
    Drive1.Top = 360 + (1920 - 360) * HyoujiBairitu
    Drive1.FontSize = 10 * HyoujiBairitu
    Drive1.Width = 5175 * HyoujiBairitu
'   Drive1.Height = 300 * HyoujiBairitu
'
    Dir1.Left = 480
    Dir1.Top = 360 + (2280 - 360) * HyoujiBairitu
    Dir1.FontSize = 10 * HyoujiBairitu
    Dir1.Width = 5175 * HyoujiBairitu
    Dir1.Height = 2190 * HyoujiBairitu
'
    cmdGo.Left = 480 + (5880 - 480) * HyoujiBairitu
    cmdGo.Top = 360 + (1920 - 360) * HyoujiBairitu
    cmdGo.FontSize = 10 * HyoujiBairitu
    cmdGo.Width = 1215 * HyoujiBairitu
    cmdGo.Height = 735 * HyoujiBairitu
'
    cmdStd.Left = 480 + (5880 - 480) * HyoujiBairitu
    cmdStd.Top = 360 + (2880 - 360) * HyoujiBairitu
    cmdStd.FontSize = 10 * HyoujiBairitu
    cmdStd.Width = 1215 * HyoujiBairitu
    cmdStd.Height = 615 * HyoujiBairitu
'
    cmdQuit.Left = 480 + (5880 - 480) * HyoujiBairitu
    cmdQuit.Top = 360 + (3840 - 360) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1215 * HyoujiBairitu
    cmdQuit.Height = 615 * HyoujiBairitu
End Sub


