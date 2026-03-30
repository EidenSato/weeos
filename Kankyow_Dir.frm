VERSION 5.00
Begin VB.Form Kankyow_Dir 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ｺｰﾄﾞ表 #0"
   ClientHeight    =   3465
   ClientLeft      =   3045
   ClientTop       =   2550
   ClientWidth     =   3870
   FillColor       =   &H00FFFFFF&
   Icon            =   "Kankyow_Dir.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   3465
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdKettei 
      Caption         =   "決定"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2190
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Kankyow_Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*********************
'*  環境設定 Dir画面  *
'*********************
'
Option Explicit
'
    Dim FLGstartP As Integer

Private Sub cmdKettei_Click()
'               *** カレントパスの設定 ***
    Select Case TMPiti
    Case 1
        Kankyow_Itiran.Text2.Text = TMPdir1
    Case 2
        Kankyow_Itiran.Text3.Text = TMPdir1
    Case 3
        Kankyow_Itiran.Text4.Text = TMPdir1
    Case 4
'        Kankyow_Itiran.Text5.Text = TMPdir1
'    Case 5
        Kankyow_Itiran.Text13.Text = TMPdir1
    End Select
'
    With Kankyow_Itiran
        .Command2.Enabled = True
        .Command3.Enabled = True
    End With
'
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
        TMPdir1 = Left$(Moji, n - 1)
    Else
        TMPdir1 = Moji
    End If
End Sub

Private Sub Drive1_Change()
    If FLGstartP = 1 Then Exit Sub  '*** 初めての「Form_Load」の時は無視する ***
'               *** カレントドライブの取得 ***
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    FLGstartP = 1       '*** スタートイニシャライズ中セット ***
'
    Left = Kankyow_Itiran.Left + Tmpleft + 30
    Top = Kankyow_Itiran.Top + Tmptop + 30
    Me.Width = 3975 * HyoujiBairitu + 15
    Me.Height = (3960 - 480) * HyoujiBairitu + 480
'
    Call DSPgamenBuhin  '*** 画面部品設定 ***
'
    Select Case TMPiti
    Case 1
        Kankyow_Dir.Caption = "ｺｰﾄﾞ表 ﾌｫﾙﾀﾞｰ"
    Case 2
        Kankyow_Dir.Caption = "構成/部品表 ﾌｫﾙﾀﾞｰ"
    Case 3
        Kankyow_Dir.Caption = "部品表 ﾌｫﾙﾀﾞｰ１"
    Case 4
'        Kankyow_Dir.Caption = "部品表 ﾌｫﾙﾀﾞｰ２"
'    Case 5
        Kankyow_Dir.Caption = "ﾜｰｸﾌｫﾙﾀﾞｰ"
    End Select
'
    On Error GoTo NoDir
'
'                       '*** データセットしただけで「Drive1_Change」が発生する。 ***
    Drive1.Drive = TMPdir1
'                       '*** データセットしただけで「Dir1_Change」が発生する。 ***
    Dir1.Path = TMPdir1
'
    FLGstartP = 0       '*** スタートイニシャライズ終了セット ***
    Exit Sub
'
NoDir:
    Resume Next
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
'
    FLGstartP = 0       '*** スタートイニシャライズ終了セット ***
End Sub

Private Sub DSPgamenBuhin()
    Drive1.Left = 0
    Drive1.Top = 0
    Drive1.FontSize = 10 * HyoujiBairitu
    Drive1.Width = 3855 * HyoujiBairitu
'
    Dir1.Left = 0
    Dir1.Top = 360 * HyoujiBairitu
    Dir1.FontSize = 10 * HyoujiBairitu
    Dir1.Width = 3855 * HyoujiBairitu
    Dir1.Height = 2190 * HyoujiBairitu
'
    cmdKettei.Left = 360 * HyoujiBairitu
    cmdKettei.Top = 2760 * HyoujiBairitu
    cmdKettei.FontSize = 10 * HyoujiBairitu
    cmdKettei.Width = 1335 * HyoujiBairitu
    cmdKettei.Height = 495 * HyoujiBairitu
'
    cmdCancel.Left = 2160 * HyoujiBairitu
    cmdCancel.Top = 2760 * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1335 * HyoujiBairitu
    cmdCancel.Height = 495 * HyoujiBairitu
End Sub

