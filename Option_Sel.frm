VERSION 5.00
Begin VB.Form Option_Sel 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "ÉIÉvÉVÉáÉìê›íË"
   ClientHeight    =   4455
   ClientLeft      =   3045
   ClientTop       =   2550
   ClientWidth     =   6375
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Option_Sel.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   4455
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSize 
      BackColor       =   &H00008000&
      Caption         =   "ï∂éöÇÃëÂÇ´Ç≥"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   5175
      Begin VB.OptionButton optTokudai 
         BackColor       =   &H00008000&
         Caption         =   "í¥ëÂÇ´Ç¢ª≤Ωﬁ(14p)"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   2040
      End
      Begin VB.OptionButton optOokiku 
         BackColor       =   &H00008000&
         Caption         =   "ëÂÇ´Ç¢ª≤Ωﬁ(12p)"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1920
      End
      Begin VB.OptionButton optHyoujyun 
         BackColor       =   &H00008000&
         Caption         =   "ïWèÄª≤Ωﬁ(10p)"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "∑¨›æŸ(&E)"
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
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame fraFolder 
      BackColor       =   &H00008000&
      Caption         =   "ç\ê¨ï\/ïîïiï\ ì«Ç›çûÇ›ÉtÉHÉãÉ_Å["
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   5175
      Begin VB.OptionButton optSelected 
         BackColor       =   &H00008000&
         Caption         =   "å¬ï Ã´Ÿ¿ﬁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   3015
      End
      Begin VB.OptionButton optDirect 
         BackColor       =   &H00008000&
         Caption         =   "ÉtÉ@ÉCÉãÇäJÇ≠éûÇÕñàâÒëIëÇ∑ÇÈÅB"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optSTD 
         BackColor       =   &H00008000&
         Caption         =   "ã§óLÃ´Ÿ¿ﬁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblOptin_STD 
         BackStyle       =   0  'ìßñæ
         Caption         =   " Åö ä¬ã´ê›íËÇ…è]Ç¢å≈íËÇ∑ÇÈÅB"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "ï¬Ç∂ÇÈ(&Q)"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "çXêV(&U)"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Option_Sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'************************
'*  ÉIÉvÉVÉáÉìê›íË âÊñ   *
'************************
'
Option Explicit
'

Private Sub Form_Load()
    Call GamenSettei    '*** âÊñ ïîïièîå≥éwíË ***
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
    Option_Sel.Caption = "ÉIÉvÉVÉáÉìê›íË"
'
    Call FolderSettei
    Call MojiSettei
'
    cmdGo.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    Call WRcont
    cmdGo.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub cmdQuit_Click()
    If cmdGo.Enabled = True Then
        Call WRcont
    End If
    Unload Me
End Sub

Private Sub optDirect_Click()
    Xcont0(16) = "0"    '*** ñàâÒëIë ***
    Xcont0(17) = "0"    '*** ã§óL ***
    cmdGo.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub optSTD_Click()
    Xcont0(16) = "1"    '*** å≈íË ***
    Xcont0(17) = "0"    '*** ã§óL ***
    cmdGo.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub optSelected_Click()
    Xcont0(16) = "1"    '*** å≈íË ***
    Xcont0(17) = "1"    '*** å¬ï  ***
    cmdGo.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub optHyoujyun_Click()
    Xcont0(15) = "*"    '*** ïWèÄ ***
    cmdGo.Enabled = True
    cmdCancel.Enabled = True
'
    Call GamenSettei
End Sub

Private Sub optOokiku_Click()
    Xcont0(15) = "1"    '*** ëÂÇ´Ç¢ ***
    cmdGo.Enabled = True
    cmdCancel.Enabled = True
'
    Call GamenSettei
End Sub

Private Sub optTokudai_Click()
    Xcont0(15) = "2"    '*** í¥ëÂÇ´Ç¢ ***
    cmdGo.Enabled = True
    cmdCancel.Enabled = True
'
    Call GamenSettei
End Sub

Private Sub FolderSettei()
    If Xcont0(16) = "1" Then
        If Xcont0(17) = "1" Then
            optSelected.Value = True
        Else
            optSTD.Value = True
        End If
    Else
        optdirect.Value = True
    End If
End Sub

Private Sub MojiSettei()
    If Xcont0(15) = "2" Then
        optTokudai.Value = True
    ElseIf Xcont0(15) = "1" Then
        optOokiku.Value = True
    Else
        optHyoujyun.Value = True
    End If
End Sub

Private Sub GamenSettei()
    If Xcont0(15) = "2" Then
        HyoujiBairitu = cUHBairitu
    ElseIf Xcont0(15) = "1" Then
        HyoujiBairitu = cHBairitu
    Else
        HyoujiBairitu = 1#
    End If
'
    fraFolder.Top = 360
    fraFolder.Left = 600
    fraFolder.Width = 5175 * HyoujiBairitu
    fraFolder.Height = 1575 * HyoujiBairitu
    fraFolder.FontSize = 10 * HyoujiBairitu
    optdirect.Top = 300 * HyoujiBairitu
    optdirect.Left = 840 * HyoujiBairitu
    optdirect.Width = 3375 * HyoujiBairitu
    optdirect.FontSize = 10 * HyoujiBairitu
    lblOptin_STD.Top = 600 * HyoujiBairitu
    lblOptin_STD.Left = 840 * HyoujiBairitu
    lblOptin_STD.Width = 3015 * HyoujiBairitu
    lblOptin_STD.FontSize = 10 * HyoujiBairitu
    optSTD.Top = 900 * HyoujiBairitu
    optSTD.Left = 1080 * HyoujiBairitu
    optSTD.Width = 3015 * HyoujiBairitu
    optSTD.FontSize = 10 * HyoujiBairitu
    optSelected.Top = 1200 * HyoujiBairitu
    optSelected.Left = 1080 * HyoujiBairitu
    optSelected.Width = 3015 * HyoujiBairitu
    optSelected.FontSize = 10 * HyoujiBairitu
'
    fraSize.Top = fraFolder.Top + fraFolder.Height + 360
    fraSize.Width = 5175 * HyoujiBairitu
    fraSize.Height = 1335 * HyoujiBairitu
    fraSize.FontSize = 10 * HyoujiBairitu
    optHyoujyun.Top = 240 * HyoujiBairitu
    optHyoujyun.Left = 1560 * HyoujiBairitu
    optHyoujyun.Width = 1815 * HyoujiBairitu
    optHyoujyun.FontSize = 10 * HyoujiBairitu
    optOokiku.Top = 600 * HyoujiBairitu
    optOokiku.Left = 1560 * HyoujiBairitu
    optOokiku.Width = 1920 * HyoujiBairitu
    optOokiku.FontSize = 10 * HyoujiBairitu
    optTokudai.Top = 960 * HyoujiBairitu
    optTokudai.Left = 1560 * HyoujiBairitu
    optTokudai.Width = 2040 * HyoujiBairitu
    optTokudai.FontSize = 10 * HyoujiBairitu
'
    cmdCancel.Top = fraSize.Top + fraSize.Height + 360 * HyoujiBairitu
    cmdCancel.Left = 600
    cmdCancel.Width = 1575 * HyoujiBairitu
    cmdCancel.Height = 495 * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdGo.Top = cmdCancel.Top
    cmdGo.Left = cmdCancel.Left + cmdCancel.Width + 240 * HyoujiBairitu
    cmdGo.Width = 1575 * HyoujiBairitu
    cmdGo.Height = 495 * HyoujiBairitu
    cmdGo.FontSize = 10 * HyoujiBairitu
    cmdQuit.Top = cmdCancel.Top
    cmdQuit.Left = cmdGo.Left + cmdGo.Width + 240 * HyoujiBairitu
    cmdQuit.Width = 1575 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
'
    Me.Width = 600 + fraFolder.Width + 600
    Me.Height = 420 + cmdCancel.Top + cmdCancel.Height + 360
End Sub
