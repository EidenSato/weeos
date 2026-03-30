VERSION 5.00
Begin VB.Form Pcod_retrieve 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "ïîïiÉRÅ[Éhåüçı"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Pcod_retrieve.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDecision 
      BackColor       =   &H00FFFF80&
      Caption         =   "åàíË(&D)"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   5280
      MaskColor       =   &H8000000F&
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   5
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtPartName 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
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
      Height          =   280
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox txtCodeNo 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
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
      Height          =   280
      Index           =   2
      Left            =   480
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "L1234-56"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtPartName 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   280
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txtCodeNo 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   280
      Index           =   1
      Left            =   480
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "L1234-56"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "åüçı(&S)"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPartName 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
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
      Height          =   280
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtCodeNo 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
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
      Height          =   280
      Index           =   0
      Left            =   480
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "L1234-56"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtPoint 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
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
      Height          =   280
      Left            =   3000
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "ï¬Ç∂ÇÈ(&Q)"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblPoint 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "åüçıï∂éöóÒ (ï∂éöóÒÇä‹Çﬁ)"
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
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "Pcod_retrieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'**************************************
'*** ÇdÇdÇnÇrÇQ ìdãCïîïiÉRÅ[Éh åüçı ***
'***      2006.10.13  by S.Fukazawa ***
'**************************************
'
Option Explicit
'                                   567twip=10mm,1440twip=1inch
Private Const OrgWidth = 6700   '*** ÉtÉHÅ[ÉÄê°ñ@èâä˙íl ***
Private Const OrgHeight = 3990
Dim HikakuMoji As String
Dim Kouho(32, 3) As String
Dim Kpoint As Integer           '*** max 15 ***
Dim Kwork As Integer
Dim Koumoku As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageDown  '*** Roll Up
        Call cmdUp_Click    '*** è„Ç÷ ***
        Call cmdUp_Click    '*** è„Ç÷ ***
        Call cmdUp_Click    '*** è„Ç÷ ***
'
    Case vbKeyUp        '*** Å™ ***
        Call cmdUp_Click    '*** è„Ç÷ ***
'
    Case vbKeyPageUp    '*** Roll Down
        Call cmdDown_Click  '*** â∫Ç÷ ***
        Call cmdDown_Click  '*** â∫Ç÷ ***
        Call cmdDown_Click  '*** â∫Ç÷ ***
'
    Case vbKeyDown      '*** Å´ ***
        Call cmdDown_Click  '*** â∫Ç÷ ***
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
'
    Me.Width = 480 + (OrgWidth - 960) * HyoujiBairitu + 480
    Me.Height = 480 + (OrgHeight - 960) * HyoujiBairitu + 480
'
    If Eeos2_mainMDI.ScaleHeight > Me.Height Then
        Me.Top = (Eeos2_mainMDI.ScaleHeight - Me.Height) / 6 * 5
    Else
        Me.Top = 0
    End If
'
    If Eeos2_mainMDI.ScaleWidth > Me.Width Then
        Me.Left = (Eeos2_mainMDI.ScaleWidth - Me.Width) / 2
    Else
        Me.Left = 0
    End If
'
    Call Buhin_Haichi       '*** ï\é¶ïîïiîzíu ***
'
    For i = 0 To 2
        txtCodeno(i).Text = "---"
        txtPartName(i).Text = "---"
    Next i
'
End Sub

Private Sub cmdQuit_Click()
    flag_cancel = True
    Unload Me
End Sub

Private Sub cmdRetrieve_Click()
    Dim i As Integer, j As Integer, k As Integer
'
    HikakuMoji = txtPoint.Text
'
    Koumoku = Aitem0(ips0, 0)
    ipT = ips0
'
    Me.MousePointer = vbHourglass
'
    For i = 0 To 2
        txtCodeno(i).Text = "---"
        txtPartName(i).Text = "---"
    Next i
        Kouho(1, 1) = "åüçıíÜÅI"
    DoEvents
'
    For i = 0 To 32
        Kouho(i, 0) = "---"
        Kouho(i, 1) = "---"
    Next i
    Kpoint = 0
'
    For jpT = 1 To Bnum0
        If Aitem0(ipT, 0) = "IC" Then
            DRVmainT = Xcont0(2) & "\IC\IC" & Left$(Bindex0(jpT, 0), 1) _
                & "\L" & Bindex0(jpT, 0) & ".COD"
        Else
            DRVmainT = Xcont0(2) & "\" & Aitem0(ipT, 0) _
                & "\L" & Bindex0(jpT, 0) & ".COD"
        End If
        Call RDmain(DRVmainT, CmainT(), CnumT, CdimT)   '*** ÉÅÉCÉìÉRÅ[Éhì«Ç›çûÇ› ***
'
        For kpT = 1 To CnumT
            k = InStr(1, Bindex0(jpT, 3) + CmainT(kpT, 1) + Bindex0(jpT, 4), HikakuMoji)
            If k <> 0 Then
                Kpoint = Kpoint + 1
                Kouho(Kpoint, 0) = "L" & Bindex0(jpT, 0) & "-" & CmainT(kpT, 0)
                Kouho(Kpoint, 1) = Bindex0(jpT, 3) + CmainT(kpT, 1) + Bindex0(jpT, 4)
                Kouho(Kpoint, 2) = str(jpT)
                Kouho(Kpoint, 3) = str(kpT)
                If 31 <= Kpoint Then
                    Kouho(32, 1) = "- à»â∫è»ó™ -"
                    GoTo nukeru
                End If
'
            End If
        Next kpT
    Next jpT
'
nukeru:
    If Kpoint = 0 Then
        Kouho(1, 1) = "äYìññ≥ÇµÅI"
    End If
'
    Me.MousePointer = vbDefault
'
    txtCodeno(0).Text = Kouho(0, 0)
    txtPartName(0).Text = Kouho(0, 1)
'
    txtCodeno(1).Text = Kouho(1, 0)
    txtPartName(1).Text = Kouho(1, 1)
'
    txtCodeno(2).Text = Kouho(2, 0)
    txtPartName(2).Text = Kouho(2, 1)
    Kwork = 1
End Sub

Private Sub cmdUp_Click()
    Kwork = Kwork - 1
    If 0 < Kwork Then
        txtCodeno(0).Text = Kouho(Kwork - 1, 0)
        txtPartName(0).Text = Kouho(Kwork - 1, 1)
'
        txtCodeno(1).Text = Kouho(Kwork, 0)
        txtPartName(1).Text = Kouho(Kwork, 1)
'
        txtCodeno(2).Text = Kouho(Kwork + 1, 0)
        txtPartName(2).Text = Kouho(Kwork + 1, 1)
    Else
        Kwork = Kwork + 1
    End If
End Sub

Private Sub cmdDecision_Click()
    If Kouho(Kwork, 0) = "---" Then
        flag_cancel = True
    Else
        jps0 = Val(Kouho(Kwork, 2))
        kps0 = Val(Kouho(Kwork, 3))
'
        flag_cancel = False
    End If
'
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Kwork = Kwork + 1
    If Kwork <= Kpoint Then
        txtCodeno(0).Text = Kouho(Kwork - 1, 0)
        txtPartName(0).Text = Kouho(Kwork - 1, 1)
'
        txtCodeno(1).Text = Kouho(Kwork, 0)
        txtPartName(1).Text = Kouho(Kwork, 1)
'
        txtCodeno(2).Text = Kouho(Kwork + 1, 0)
        txtPartName(2).Text = Kouho(Kwork + 1, 1)
    Else
        Kwork = Kwork - 1
    End If
End Sub

Private Sub Buhin_Haichi()
    lblPoint.FontSize = 10 * HyoujiBairitu
    lblPoint.Left = 480
    lblPoint.Top = 480
    lblPoint.Width = 2535 * HyoujiBairitu
'
    txtPoint.FontSize = 10 * HyoujiBairitu
    txtPoint.Height = 280 * HyoujiBairitu
    txtPoint.Left = lblPoint.Left + lblPoint.Width + 10
    txtPoint.Top = lblPoint.Top
    txtPoint.Width = (3135 - 10) * HyoujiBairitu
'
    lblPoint.Height = txtPoint.Height
'
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Height = 375 * HyoujiBairitu
    cmdQuit.Left = 480 + (3840 - 480) * HyoujiBairitu
    cmdQuit.Top = 480 + (2880 - 480) * HyoujiBairitu
    cmdQuit.Width = 1215 * HyoujiBairitu
'
    cmdRetrieve.FontSize = 10 * HyoujiBairitu
    cmdRetrieve.Height = 495 * HyoujiBairitu
    cmdRetrieve.Left = 480 + (3840 - 480) * HyoujiBairitu
    cmdRetrieve.Top = 480 + (960 - 480) * HyoujiBairitu
    cmdRetrieve.Width = 1215 * HyoujiBairitu
'
    txtCodeno(0).FontSize = 10 * HyoujiBairitu
    txtCodeno(0).Height = 280 * HyoujiBairitu
    txtCodeno(0).Left = 480
    txtCodeno(0).Top = 480 + (1680 - 480) * HyoujiBairitu
    txtCodeno(0).Width = (1095 - 10) * HyoujiBairitu
'
    txtPartName(0).FontSize = 10 * HyoujiBairitu
    txtPartName(0).Height = 280 * HyoujiBairitu
    txtPartName(0).Left = txtCodeno(0).Left + txtCodeno(0).Width + 10
    txtPartName(0).Top = txtCodeno(0).Top
    txtPartName(0).Width = 3495 * HyoujiBairitu
'
    txtCodeno(1).FontSize = 10 * HyoujiBairitu
    txtCodeno(1).Height = 280 * HyoujiBairitu
    txtCodeno(1).Left = 480
    txtCodeno(1).Top = 480 + (2040 - 480) * HyoujiBairitu
    txtCodeno(1).Width = (1095 - 10) * HyoujiBairitu
'
    txtPartName(1).FontSize = 10 * HyoujiBairitu
    txtPartName(1).Height = 280 * HyoujiBairitu
    txtPartName(1).Left = txtCodeno(1).Left + txtCodeno(1).Width + 10
    txtPartName(1).Top = txtCodeno(1).Top
    txtPartName(1).Width = 3495 * HyoujiBairitu
'
    txtCodeno(2).FontSize = 10 * HyoujiBairitu
    txtCodeno(2).Height = 280 * HyoujiBairitu
    txtCodeno(2).Left = 480
    txtCodeno(2).Top = 480 + (2400 - 480) * HyoujiBairitu
    txtCodeno(2).Width = (1095 - 10) * HyoujiBairitu
'
    txtPartName(2).FontSize = 10 * HyoujiBairitu
    txtPartName(2).Height = 280 * HyoujiBairitu
    txtPartName(2).Left = txtCodeno(2).Left + txtCodeno(2).Width + 10
    txtPartName(2).Top = txtCodeno(2).Top
    txtPartName(2).Width = 3495 * HyoujiBairitu
'
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Height = 375 * HyoujiBairitu
    cmdUp.Left = 480 + (5280 - 480) * HyoujiBairitu
    cmdUp.Top = txtPartName(0).Top + txtPartName(0).Height / 2 - cmdUp.Height
    cmdUp.Width = 855 * HyoujiBairitu
'
    cmdDecision.FontSize = 10 * HyoujiBairitu
    cmdDecision.Height = 495 * HyoujiBairitu
    cmdDecision.Left = 480 + (5280 - 480) * HyoujiBairitu
    cmdDecision.Top = txtPartName(1).Top + txtPartName(1).Height / 2 - cmdDecision.Height / 2
    cmdDecision.Width = 855 * HyoujiBairitu
'
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Height = 375 * HyoujiBairitu
    cmdDown.Left = 480 + (5280 - 480) * HyoujiBairitu
    cmdDown.Top = txtPartName(2).Top + txtPartName(2).Height / 2
    cmdDown.Width = 855 * HyoujiBairitu
End Sub
