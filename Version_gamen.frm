VERSION 5.00
Begin VB.Form Version_gamen 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "ÇdÇdOSáU  ﬁ∞ºﬁÆ›"
   ClientHeight    =   1665
   ClientLeft      =   1875
   ClientTop       =   1530
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Version_gamen.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   1665
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_ok 
      Caption         =   "ï¬Ç∂ÇÈ(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "Version_gamen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*********************
'*  EEOSáU ÉoÅ[ÉWÉáÉì *
'*********************
'
Option Explicit

Private Sub cmd_OK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'                   ÉtÉHÅ[ÉÄÇÃï\é¶à íuÇÃê›íË
    Width = 480 + (6825 - 960) * HyoujiBairitu + 480
    Height = 600 + (2070 - 1200) * HyoujiBairitu + 600
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
    lblVersion.Left = 480
    lblVersion.Top = 600
    lblVersion.FontSize = 12 * HyoujiBairitu
    lblVersion.Width = 5775 * HyoujiBairitu
    lblVersion.Height = 255 * HyoujiBairitu
'
    cmd_OK.Left = 480 + (5040 - 480) * HyoujiBairitu
    cmd_OK.Top = 600 + (960 - 600) * HyoujiBairitu
    cmd_OK.FontSize = 10 * HyoujiBairitu
    cmd_OK.Width = 1215 * HyoujiBairitu
    cmd_OK.Height = 495 * HyoujiBairitu
'
    Version_gamen.Caption = "ÇdÇdÇnÇrÇQ ÉoÅ[ÉWÉáÉì"
    lblVersion.Caption = EEOS_Version
End Sub
