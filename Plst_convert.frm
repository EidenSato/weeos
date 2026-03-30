VERSION 5.00
Begin VB.Form Plst_convert 
   BackColor       =   &H00004000&
   Caption         =   "OrCADïœä∑ çÏã∆Ãß≤Ÿ ï“èW"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11055
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Plst_convert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   4110
   ScaleWidth      =   11055
   Begin VB.CommandButton cmdBottom 
      Caption         =   "ÅÅ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   36
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdTop 
      Caption         =   "ÅÅ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   35
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "àÍçsÅ®çÌèú"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   30
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmd5DOWN 
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   34
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmd5UP 
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   33
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdDOWN 
      Caption         =   "Å´"
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
      Left            =   9960
      TabIndex        =   32
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdUP 
      Caption         =   "Å™"
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
      Left            =   9960
      TabIndex        =   31
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtBikou 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   4
      Left            =   7800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "*"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtShitei 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   4
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   4
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtKigou 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   " U"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtCodeno 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   4
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtMeisyou 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "123456789012345"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtBikou 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   3
      Left            =   7800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "*"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtShitei 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   3
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   3
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtKigou 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   " U"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtCodeno 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   3
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtMeisyou 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   3
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "123456789012345"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtBikou 
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
      Height          =   270
      Index           =   2
      Left            =   7800
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   17
      Text            =   "*"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtShitei 
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
      Height          =   270
      Index           =   2
      Left            =   6600
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   16
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
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
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Index           =   2
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtKigou 
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
      Height          =   270
      Index           =   2
      Left            =   1920
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   13
      Text            =   " U"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtCodeno 
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
      Height          =   270
      Index           =   2
      Left            =   4800
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   15
      Text            =   " L1234-56"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtMeisyou 
      BackColor       =   &H00008000&
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
      Height          =   270
      Index           =   2
      Left            =   3000
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   14
      Text            =   "123456789012345"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtBikou 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   1
      Left            =   7800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "*"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtShitei 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   1
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   1
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "-"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtKigou 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " U"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtCodeno 
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
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtMeisyou 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   270
      Index           =   1
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "123456789012345"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtBikou 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   7800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "*"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtShitei 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "-"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "çXêV(&U)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      TabIndex        =   38
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "∑¨›æŸ(&E)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   37
      Top             =   3240
      Width           =   1455
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
      Left            =   8280
      TabIndex        =   40
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtMeisyou 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "123456789012345"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtCodeno 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   " L1234-56"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtKigou 
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
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   " U"
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblComment 
      Alignment       =   2  'íÜâõëµÇ¶
      Appearance      =   0  'Ã◊Øƒ
      BackColor       =   &H00004000&
      Caption         =   "0ÅF≈º, 8-10ÅFéwíË“∞∂∞"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   6480
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   46
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblBikou 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "îı  çl"
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
      Left            =   7800
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   45
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblShitei 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "“∞∂∞éwíË"
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
      Left            =   6600
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   44
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "çÄñ⁄î‘çÜ"
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
      Left            =   960
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   43
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblMeisyou 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ïîïiï\é¶ ñºèÃ"
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
      Left            =   3000
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   42
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblCodeno 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ëŒâû ∫∞ƒﬁî‘çÜ"
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
      Left            =   4800
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   41
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblKigou 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00008000&
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ïîïiãLçÜ"
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
      Left            =   1920
      MousePointer    =   1  'ñÓàÛ
      TabIndex        =   39
      Top             =   480
      Width           =   975
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Ãß≤Ÿ(&X)"
      Begin VB.Menu mnuCancel 
         Caption         =   "∑¨›æŸ(&E)"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "çXêV(&U)"
      End
      Begin VB.Menu mnuSQuit 
         Caption         =   "ï¬Ç∂ÇÈ(&Q)"
      End
      Begin VB.Menu mnuãÊêÿÇËê¸12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAQuit 
         Caption         =   "EEOSÇQÇÃèIóπ(&X)"
      End
   End
   Begin VB.Menu mnuKouseihyou 
      Caption         =   "ç\ê¨ï\(&K)"
      Begin VB.Menu mnuKousei 
         Caption         =   "ìdãC ç\ê¨ï\(&C)..."
      End
   End
   Begin VB.Menu mnuBuhinhyou 
      Caption         =   "ïîïiï\(&P)"
      Begin VB.Menu mnuBuhin 
         Caption         =   "ìdãC ïîïiï\(&C)..."
      End
      Begin VB.Menu mnuBuhin2 
         Caption         =   "ìdãC ïîïiï\ÇQ(&D)..."
      End
      Begin VB.Menu mnuORCAD 
         Caption         =   "OrCADïœä∑(&O)..."
      End
      Begin VB.Menu mnuConvFile 
         Caption         =   "ïœä∑çÏã∆Ãß≤Ÿ(&W)"
      End
      Begin VB.Menu mnuãÊêÿÇËê¸31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuhinPRN 
         Caption         =   "ïîïiï\àÛç¸(&P)..."
      End
      Begin VB.Menu mnuFilePrnA 
         Caption         =   "àÍóóï\àÛç¸(&L)..."
      End
      Begin VB.Menu mnuSuuryo 
         Caption         =   "êîó ï\àÛç¸(&T)..."
      End
   End
   Begin VB.Menu mnuCodehyou 
      Caption         =   "∫∞ƒﬁï\(&C)"
      Begin VB.Menu mnuCode 
         Caption         =   "çÄñ⁄àÍóó(&M)"
      End
      Begin VB.Menu mnuHinsyu 
         Caption         =   "ïiéÌàÍóó(&I)"
      End
      Begin VB.Menu mnuPmain 
         Caption         =   "ïiñ⁄àÍóó(&P)"
      End
      Begin VB.Menu mnuMakerment 
         Caption         =   "“∞∂∞∫∞ƒﬁï\(&M)"
      End
      Begin VB.Menu mnuTraderment 
         Caption         =   "è§é–∫∞ƒﬁï\(&T)"
      End
   End
   Begin VB.Menu mnuJump 
      Caption         =   "ºﬁ¨›Ãﬂ(&J)"
      Begin VB.Menu mnuJumpT 
         Caption         =   "êÊì™Ç÷ºﬁ¨›Ãﬂ(&T)"
      End
      Begin VB.Menu mnuJumpC 
         Caption         =   "íÜêSÇ÷ºﬁ¨›Ãﬂ(&C)"
      End
      Begin VB.Menu mnuJumpE 
         Caption         =   "ç≈å„ïîÇ÷ºﬁ¨›Ãﬂ(&E)"
      End
   End
   Begin VB.Menu mnuWindou 
      Caption         =   "≥≤›ƒﬁ≥(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileH 
         Caption         =   "è„â∫Ç…ï¿Ç◊Çƒï\é¶(&H)"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "ç∂âEÇ…ï¿Ç◊Çƒï\é¶(&V)"
      End
      Begin VB.Menu mnuTileC 
         Caption         =   "èdÇÀÇƒï\é¶(&C)"
      End
      Begin VB.Menu mnuReform 
         Caption         =   "èâä˙à íuÇ…ñﬂÇ∑(&S)"
      End
   End
   Begin VB.Menu mnuKnakyou 
      Caption         =   "ä¬ã´(&O)"
      Begin VB.Menu mnuSettei 
         Caption         =   "ä¬ã´ê›íË(&K)"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "µÃﬂºÆ›(&O)"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "ÕŸÃﬂ(&H)"
      Begin VB.Menu mnuSetumei 
         Caption         =   "ëÄçÏê‡ñæ(&S)"
      End
      Begin VB.Menu mnuKaihan 
         Caption         =   "â¸î≈óöó(&H)"
      End
      Begin VB.Menu mnuãÊêÿÇËê¸81 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   " ﬁ∞ºﬁÆ›(&V)"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ŒﬂØÃﬂ±ØÃﬂ“∆≠∞"
      Visible         =   0   'False
      Begin VB.Menu mnuJumpTP 
         Caption         =   "êÊì™Ç÷ºﬁ¨›Ãﬂ"
      End
      Begin VB.Menu mnuJumpCP 
         Caption         =   "íÜêSÇ÷ºﬁ¨›Ãﬂ"
      End
      Begin VB.Menu mnuJumpEP 
         Caption         =   "ç≈å„ïîÇ÷ºﬁ¨›Ãﬂ"
      End
      Begin VB.Menu mnuãÊêÿÇËê¸91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKouseihyouP 
         Caption         =   "ç\ê¨ï\"
         Begin VB.Menu mnuKouseiP 
            Caption         =   "ìdãC ç\ê¨ï\..."
         End
      End
      Begin VB.Menu mnuPuBuhinhyou 
         Caption         =   "ïîïiï\"
         Begin VB.Menu mnuBuhinP 
            Caption         =   "ìdãC ïîïiï\..."
         End
         Begin VB.Menu mnuBuhin2P 
            Caption         =   "ìdãC ïîïiï\ÇQ..."
         End
         Begin VB.Menu mnuORCADP 
            Caption         =   "OrCADïœä∑..."
         End
         Begin VB.Menu mnuãÊêÿÇËê¸951 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBuhinPRNP 
            Caption         =   "ïîïiï\àÛç¸..."
         End
         Begin VB.Menu mnuFilePrnAP 
            Caption         =   "àÍóóï\àÛç¸..."
         End
         Begin VB.Menu mnuSuuryoP 
            Caption         =   "êîó ï\àÛç¸..."
         End
      End
      Begin VB.Menu mnuPuCodehyou 
         Caption         =   "∫∞ƒﬁï\"
         Begin VB.Menu mnuCodeP 
            Caption         =   "çÄñ⁄àÍóó"
         End
         Begin VB.Menu mnuHinsyuP 
            Caption         =   "ïiéÌàÍóó"
         End
         Begin VB.Menu mnuPmainP 
            Caption         =   "ïiñ⁄àÍóó"
         End
         Begin VB.Menu mnuMakermentP 
            Caption         =   "“∞∂∞∫∞ƒﬁï\"
         End
         Begin VB.Menu mnuTradermentP 
            Caption         =   "è§é–∫∞ƒﬁï\"
         End
      End
      Begin VB.Menu mnuãÊêÿÇËê¸95 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackP 
         Caption         =   "ï¬Ç∂ÇÈ"
      End
      Begin VB.Menu mnuãÊêÿÇËê¸96 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAQuitP 
         Caption         =   "EEOSÇQÇÃèIóπ"
      End
   End
End
Attribute VB_Name = "Plst_convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*****************************
'* OrCADïœä∑ çÏã∆Ãß≤Ÿ ï“èW ***
'*****************************
'
Option Explicit
'
Dim HeadTitle As String
Dim CHaiki As String
'
Dim FLGoffsetX As Integer
Dim FLGoffsetY As Integer
'                                   567twip=10mm,1440twip=1inch
Private Const OrgWidth = 11175  '*** ÉtÉHÅ[ÉÄê°ñ@èâä˙íl ***
Private Const OrgHeight = 4590
Dim tempWidth As Integer
Dim tempHeight As Integer
'
Dim FLGchange As Boolean
Dim Disp_Pointer As Integer
Dim TEMP_code As String
Dim TEMP_maker As String
Dim TEMP_bikou As String
Dim TEMP_Meisyou As String
Dim TEMP_Kigou As String

Private Sub Form_Activate()
    FLGplstWork = 1
    FLGjob = 2
    FLGlevel = 4    '*** OrCADïœä∑ çÏã∆Ãß≤Ÿ ï“èW ***
    STATUS = HeadTitle  '*** ëIëÉEÉCÉìÉhÉEÇÃÉ^ÉCÉgÉãñºèÃ ***
'
    Call MENU_settei    '*** ÉÅÉjÉÖÅ[èÛë‘ê›íË ***
'
    txtKigou(2).SetFocus
End Sub

Private Sub Form_Initialize()
    HeadTitle = STATUS
    FLGplstWork = 1
End Sub

Private Sub Form_Load()
                            '*** ÉtÉHÅ[ÉÄÇÃÉTÉCÉYÇÃê›íË
    tempWidth = 360 + (OrgWidth - 720) * HyoujiBairitu + 360
    tempHeight = 360 + (OrgHeight - 720) * HyoujiBairitu + 360
'
    Me.Width = tempWidth    '*** Ç±ÇÍÇ≈ÅuForm_ResizeÅväÑÇËçûÇ›Ç™î≠ê∂Ç∑ÇÈÅB ***
    Me.Height = tempHeight
'
    Call setFormArea        '*** ÉtÉHÅ[ÉÄÇÃï\é¶à íuÇÃê›íË
'
    FLGoffsetX = 0          '*** èâä˙âª ***
    FLGoffsetY = 0
'
    Me.Caption = HeadTitle
'
    CHaiki = "çÏã∆ÉtÉ@ÉCÉãÇÕïœçXÇ≥ÇÍÇƒÇ¢Ç‹Ç∑ÅBÅuîpä¸èIóπÅvÇÉLÉÉÉìÉZÉãÇµÇ‹Ç∑Ç©ÅH"
'
    FLGchange = False
    Call SET_Command_Button
'
    Call DSPgamenBuhin  '*** âÊñ ïîïiê›íË ***
'
    Call RDplstWork
'
    Disp_Pointer = 0
    Call Data_Display(Disp_Pointer)
'
    TEMP_code = ""
    TEMP_maker = ""
    TEMP_bikou = ""
    TEMP_Meisyou = ""
    TEMP_Kigou = ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup  '*** âEÉ{É^Éìèàóù ***
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        '
    End If
'
    FLGplstWork = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageDown      '*** Roll Up
        Call cmd5UP_Click   '*** è„Ç÷ ***
'
    Case vbKeyUp            '*** Å™ ***
        Call cmdUp_Click    '*** àÍÇ¬è„Ç÷ ***
'
    Case vbKeyPageUp        '*** Roll Down
        Call cmd5DOWN_Click '*** â∫Ç÷ ***
'
    Case vbKeyDown          '*** Å´ ***
        Call cmdDown_Click  '*** àÍÇ¬â∫Ç÷ ***
    End Select
End Sub

Private Sub Form_Resize()
'                   ÉtÉHÅ[ÉÄç\ê¨ïîïiÇÃï\é¶à íuÇÃê›íË
    If Me.Width > tempWidth Then
        FLGoffsetX = (Me.Width - tempWidth) \ 2
    Else
        FLGoffsetX = 0
    End If
'
    If Me.Height > tempHeight Then
        FLGoffsetY = (Me.Height - tempHeight) \ 2
    Else
        FLGoffsetY = 0
    End If
'
    Call DSPgamenBuhin  '*** âÊñ ïîïiê›íË ***
End Sub

Private Sub cmdQuit_Click()
    Call cmdUpdate_Click
'
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Call WRplstWork
'
    FLGchange = False
    Call SET_Command_Button
End Sub

Private Sub cmdTop_Click()
    Disp_Pointer = 0
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub cmd5UP_Click()
    Disp_Pointer = Disp_Pointer - 5
    If Disp_Pointer < 0 Then Disp_Pointer = 0
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub cmdUp_Click()
    Disp_Pointer = Disp_Pointer - 1
    If Disp_Pointer < 0 Then Disp_Pointer = 0
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub cmdDown_Click()
    Disp_Pointer = Disp_Pointer + 1
    If cPLSTWORKmax < Disp_Pointer Then Disp_Pointer = cPLSTWORKmax
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub cmd5DOWN_Click()
    Disp_Pointer = Disp_Pointer + 5
    If cPLSTWORKmax < Disp_Pointer Then Disp_Pointer = cPLSTWORKmax
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub cmdBottom_Click()
    Disp_Pointer = cPLSTWORKmax
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    Dim j As Integer
'
    For i = Disp_Pointer To cPLSTWORKmax - 1
        For j = 0 To cPLSTWORKdim
            PlstWork(i, j) = PlstWork(i + 1, j)
        Next j
    Next i
'
        For j = 0 To cPLSTWORKdim
            PlstWork(i, j) = ""
        Next j
'
    Call Data_Display(Disp_Pointer)
'
    FLGchange = True
    Call SET_Command_Button
End Sub

Private Sub mnuCancel_Click()
    Call cmdCancel_Click
End Sub

Private Sub mnuUpdate_Click()
    Call cmdUpdate_Click
End Sub

Private Sub mnuSQuit_Click()
    Call cmdQuit_Click
End Sub

Private Sub mnuAQuit_Click()
    Dim i As Integer
    If FLGchange = 1 Then
        Beep
        i = MsgBox(CHaiki, vbQuestion Or vbYesNo, STATUS)
        If i = vbNo Then
            Unload Me
            End
'
        End If
'
    Else
        Unload Me
        End
'
    End If
End Sub

Private Sub mnuKousei_Click()
    Call mnuDenkiKouseihyou
End Sub

Private Sub mnuBuhin_Click()
    Call mnuDenkiBuhinhyou
End Sub

Private Sub mnuBuhin2_Click()
    Call mnuDenkiBuhinhyou2
End Sub

Private Sub mnuORCAD_Click()
    Call mnuOrCAD_Henkan
End Sub

Private Sub mnuConvFile_Click()
'
End Sub

Private Sub mnuBuhinPRN_Click()
    Call mnuStandardBuhinhyouPrint
End Sub

Private Sub mnuFilePrnA_Click()
    Call mnuBuhinItiranhyouPrint
End Sub

Private Sub mnuSuuryo_Click()
    Call mnuBuhinSuuryohyouPrint
End Sub

Private Sub mnuCode_Click()
    Call mnuCodeBuhinMaintenance
End Sub

Private Sub mnuHinsyu_Click()
    Call mnuCodeHinsyuMaintenance
End Sub

Private Sub mnuPmain_Click()
    Call mnuCodePmainMaintenance
End Sub

Private Sub mnuMakerment_Click()
    Call mnuCodeMakerMaintenance
End Sub

Private Sub mnuTraderment_Click()
    Call mnuCodeTraderMaintenance
End Sub

Private Sub mnuJumpT_Click()
    Call cmdTop_Click
End Sub

Private Sub mnuJumpC_Click()
    Disp_Pointer = cPLSTWORKmax / 2
'
    Call Data_Display(Disp_Pointer)
End Sub

Private Sub mnuJumpE_Click()
    Call cmdBottom_Click
End Sub

Private Sub mnuTileH_Click()
    Eeos2_mainMDI.Arrange vbTileHorizontal  '*** ï¿Ç◊Çƒï\é¶ ***
End Sub

Private Sub mnuTileV_Click()
    Eeos2_mainMDI.Arrange vbTileVertical    '*** ï¿Ç◊Çƒï\é¶ ***
End Sub

Private Sub mnuTileC_Click()
    Eeos2_mainMDI.Arrange vbCascade         '*** èdÇÀÇƒï\é¶ ***
End Sub

Private Sub mnuReform_Click()
    Me.Width = tempWidth
    Me.Height = tempHeight '*** Ç±ÇÍÇ≈ÅuForm_ResizeÅväÑÇËçûÇ›Ç™î≠ê∂Ç∑ÇÈÅB ***
'
    Call setFormArea    '*** ÉtÉHÅ[ÉÄÇÃï\é¶à íuÇÃê›íË ***
End Sub

Private Sub mnuSettei_Click()
    Call mnuKankyouSettei
End Sub

Private Sub mnuOption_Click()
    Call mnuOptionSettei
End Sub

Private Sub mnuSetumei_Click()
    FLGjob = 2          '*** ïîïiï\ÉtÉâÉO ***
    Call mnuSousaSetumei
End Sub

Private Sub mnuKaihan_Click()
    Call mnuKaihanRireki
End Sub

Private Sub mnuVersion_Click()
    Call mnuVersionGamen
End Sub

Private Sub mnuJumpTP_Click()
    Call mnuJumpT_Click
End Sub

Private Sub mnuJumpCP_Click()
    Call mnuJumpC_Click
End Sub

Private Sub mnuJumpEP_Click()
    Call mnuJumpE_Click
End Sub

Private Sub mnuKouseiP_Click()
    Call mnuKousei_Click
End Sub

Private Sub mnuBuhinP_Click()
    Call mnuBuhin_Click
End Sub

Private Sub mnuBuhin2P_Click()
    Call mnuBuhin2_Click
End Sub

Private Sub mnuORCADP_Click()
    Call mnuORCAD_Click
End Sub

Private Sub mnuBuhinPRNP_Click()
    Call mnuBuhinPRN_Click
End Sub

Private Sub mnuFilePrnAP_Click()
    Call mnuFilePrnA_Click
End Sub

Private Sub mnuSuuryoP_Click()
    Call mnuSuuryo_Click
End Sub

Private Sub mnuCodeP_Click()
    Call mnuCode_Click
End Sub

Private Sub mnuHinsyuP_Click()
    Call mnuHinsyu_Click
End Sub

Private Sub mnuPmainP_Click()
    Call mnuPmain_Click
End Sub

Private Sub mnuMakermentP_Click()
    Call mnuMakerment_Click
End Sub

Private Sub mnuTradermentP_Click()
    Call mnuTraderment_Click
End Sub

Private Sub mnuBackP_Click()
    Call cmdQuit_Click
End Sub

Private Sub mnuAQuitP_Click()
    Call mnuAQuit_Click
End Sub

Private Sub txtKigou_GotFocus(Index As Integer)
    If Index = 2 Then
        txtKigou(Index).MousePointer = vbIbeam
        TEMP_Kigou = txtKigou(Index).Text
    End If
End Sub

Private Sub txtKigou_LostFocus(Index As Integer)
        txtKigou(Index).MousePointer = vbArrow
'
    If TEMP_Kigou <> txtKigou(Index).Text And Index = 2 Then
        PlstWork(Index - 2 + Disp_Pointer, 4) = txtKigou(Index).Text
        FLGchange = True
        Call SET_Command_Button
    End If
End Sub

Private Sub txtMeisyou_GotFocus(Index As Integer)
    If Index = 2 Then
        txtMeisyou(Index).MousePointer = vbIbeam
        TEMP_Meisyou = txtMeisyou(Index).Text
    End If
End Sub

Private Sub txtmeisyou_LostFocus(Index As Integer)
        txtMeisyou(Index).MousePointer = vbArrow
'
    If TEMP_Meisyou <> txtMeisyou(Index).Text And Index = 2 Then
        PlstWork(Index - 2 + Disp_Pointer, 1) = txtMeisyou(Index).Text
        FLGchange = True
        Call SET_Command_Button
    End If
End Sub

Private Sub txtCodeno_GotFocus(Index As Integer)
    If Index = 2 Then
        txtCodeno(Index).MousePointer = vbIbeam
        TEMP_code = txtCodeno(Index).Text
    End If
End Sub

Private Sub txtCodeno_LostFocus(Index As Integer)
        txtCodeno(Index).MousePointer = vbArrow
'
    If TEMP_code <> txtCodeno(Index).Text And Index = 2 Then
        PlstWork(Index - 2 + Disp_Pointer, 0) = txtCodeno(Index).Text
        FLGchange = True
        Call SET_Command_Button
    End If
End Sub

Private Sub txtShitei_GotFocus(Index As Integer)
    If Index = 2 Then
        txtShitei(Index).MousePointer = vbIbeam
        TEMP_maker = txtShitei(Index).Text
    End If
End Sub

Private Sub txtShitei_LostFocus(Index As Integer)
        txtShitei(Index).MousePointer = vbArrow
'
    If TEMP_maker <> txtShitei(Index).Text And Index = 2 Then
        PlstWork(Index - 2 + Disp_Pointer, 3) = txtShitei(Index).Text
        FLGchange = True
        Call SET_Command_Button
    End If
End Sub

Private Sub txtBikou_GotFocus(Index As Integer)
    If Index = 2 Then
        txtBikou(Index).MousePointer = vbIbeam
        TEMP_bikou = txtBikou(Index).Text
    End If
End Sub

Private Sub txtBikou_LostFocus(Index As Integer)
        txtBikou(Index).MousePointer = vbArrow
'
    If TEMP_bikou <> txtBikou(Index).Text And Index = 2 Then
        PlstWork(Index - 2 + Disp_Pointer, 2) = txtBikou(Index).Text
        FLGchange = True
        Call SET_Command_Button
    End If
End Sub

Private Sub SET_Command_Button()
    If FLGchange = True Then
        cmdCancel.Enabled = True
        cmdUpdate.Enabled = True
    Else
        cmdCancel.Enabled = False
        cmdUpdate.Enabled = False
    End If
End Sub

Private Sub setFormArea()   '*** ÉtÉHÅ[ÉÄÇÃï\é¶à íuÇÃê›íË ***
        Me.Top = 0
'
    If Eeos2_mainMDI.ScaleWidth > Me.Width Then
        Me.Left = (Eeos2_mainMDI.ScaleWidth - Me.Width) * 2 \ 3
    Else
        Me.Left = 0
    End If
End Sub

Private Sub Data_Display(top_no As Integer)
    Dim i As Integer
'
    For i = 0 To 4
        If (i + top_no - 2) < 0 Then
            txtNumber(i).Text = "-"
            txtKigou(i).Text = "-"
            txtMeisyou(i).Text = "-"
            txtCodeno(i).Text = "-"
            txtShitei(i).Text = "-"
            txtBikou(i).Text = "-"
        ElseIf cPLSTWORKmax < (i + top_no - 2) Then
            txtNumber(i).Text = "-"
            txtKigou(i).Text = "-"
            txtMeisyou(i).Text = "-"
            txtCodeno(i).Text = "-"
            txtShitei(i).Text = "-"
            txtBikou(i).Text = "-"
        Else
            txtNumber(i).Text = str(i + top_no - 2)
            txtKigou(i).Text = PlstWork(i + top_no - 2, 4)
            txtMeisyou(i).Text = PlstWork(i + top_no - 2, 1)
            txtCodeno(i).Text = PlstWork(i + top_no - 2, 0)
            txtShitei(i).Text = PlstWork(i + top_no - 2, 3)
            txtBikou(i).Text = PlstWork(i + top_no - 2, 2)
        End If
    Next i
End Sub

Private Sub DSPgamenBuhin()
    Dim i As Integer
'
    lblNumber.Left = 360 + (960 - 360) * HyoujiBairitu + FLGoffsetX
    lblNumber.Top = 480 + FLGoffsetY
    lblNumber.FontSize = 10 * HyoujiBairitu
    lblNumber.Width = 855 * HyoujiBairitu
    lblNumber.Height = 255 * HyoujiBairitu
'
    lblKigou.Left = 360 + (1920 - 360) * HyoujiBairitu + FLGoffsetX
    lblKigou.Top = 480 + FLGoffsetY
    lblKigou.FontSize = 10 * HyoujiBairitu
    lblKigou.Width = 975 * HyoujiBairitu
    lblKigou.Height = 255 * HyoujiBairitu
'
    lblMeisyou.Left = 360 + (3000 - 360) * HyoujiBairitu + FLGoffsetX
    lblMeisyou.Top = 480 + FLGoffsetY
    lblMeisyou.FontSize = 10 * HyoujiBairitu
    lblMeisyou.Width = 1695 * HyoujiBairitu
    lblMeisyou.Height = 255 * HyoujiBairitu
'
    lblCodeno.Left = 360 + (4800 - 360) * HyoujiBairitu + FLGoffsetX
    lblCodeno.Top = 480 + FLGoffsetY
    lblCodeno.FontSize = 10 * HyoujiBairitu
    lblCodeno.Width = 1695 * HyoujiBairitu
    lblCodeno.Height = 255 * HyoujiBairitu
'
    lblShitei.Left = 360 + (6600 - 360) * HyoujiBairitu + FLGoffsetX
    lblShitei.Top = 480 + FLGoffsetY
    lblShitei.FontSize = 10 * HyoujiBairitu
    lblShitei.Width = 1095 * HyoujiBairitu
    lblShitei.Height = 255 * HyoujiBairitu
'
    lblBikou.Left = 360 + (7800 - 360) * HyoujiBairitu + FLGoffsetX
    lblBikou.Top = 480 + FLGoffsetY
    lblBikou.FontSize = 10 * HyoujiBairitu
    lblBikou.Width = 1935 * HyoujiBairitu
    lblBikou.Height = 255 * HyoujiBairitu
'
    For i = 0 To 4
        txtNumber(i).Left = 360 + (960 - 360) * HyoujiBairitu + FLGoffsetX
        txtNumber(i).Top = 480 + (960 - 480) * HyoujiBairitu + 360 * HyoujiBairitu * i + FLGoffsetY
        txtNumber(i).FontSize = 10 * HyoujiBairitu
        txtNumber(i).Width = 855 * HyoujiBairitu
        txtNumber(i).Height = 270 * HyoujiBairitu
'
        txtKigou(i).Left = 360 + (1920 - 360) * HyoujiBairitu + FLGoffsetX
        txtKigou(i).Top = 480 + (960 - 480) * HyoujiBairitu + 360 * HyoujiBairitu * i + FLGoffsetY
        txtKigou(i).FontSize = 10 * HyoujiBairitu
        txtKigou(i).Width = 975 * HyoujiBairitu
        txtKigou(i).Height = 270 * HyoujiBairitu
'
        txtMeisyou(i).Left = 360 + (3000 - 360) * HyoujiBairitu + FLGoffsetX
        txtMeisyou(i).Top = 480 + (960 - 480) * HyoujiBairitu + 360 * HyoujiBairitu * i + FLGoffsetY
        txtMeisyou(i).FontSize = 10 * HyoujiBairitu
        txtMeisyou(i).Width = 1695 * HyoujiBairitu
        txtMeisyou(i).Height = 270 * HyoujiBairitu
'
        txtCodeno(i).Left = 360 + (4800 - 360) * HyoujiBairitu + FLGoffsetX
        txtCodeno(i).Top = 480 + (960 - 480) * HyoujiBairitu + 360 * HyoujiBairitu * i + FLGoffsetY
        txtCodeno(i).FontSize = 10 * HyoujiBairitu
        txtCodeno(i).Width = 1695 * HyoujiBairitu
        txtCodeno(i).Height = 270 * HyoujiBairitu
'
        txtShitei(i).Left = 360 + (6600 - 360) * HyoujiBairitu + FLGoffsetX
        txtShitei(i).Top = 480 + (960 - 480) * HyoujiBairitu + 360 * HyoujiBairitu * i + FLGoffsetY
        txtShitei(i).FontSize = 10 * HyoujiBairitu
        txtShitei(i).Width = 1095 * HyoujiBairitu
        txtShitei(i).Height = 270 * HyoujiBairitu
'
        txtBikou(i).Left = 360 + (7800 - 360) * HyoujiBairitu + FLGoffsetX
        txtBikou(i).Top = 480 + (960 - 480) * HyoujiBairitu + 360 * HyoujiBairitu * i + FLGoffsetY
        txtBikou(i).FontSize = 10 * HyoujiBairitu
        txtBikou(i).Width = 1935 * HyoujiBairitu
        txtBikou(i).Height = 270 * HyoujiBairitu
    Next i
'
    cmdDelete.Left = 360 + FLGoffsetX
    cmdDelete.Top = 480 + (1200 - 480) * HyoujiBairitu + FLGoffsetY
    cmdDelete.FontSize = 10 * HyoujiBairitu
    cmdDelete.Width = 495 * HyoujiBairitu
    cmdDelete.Height = 1215 * HyoujiBairitu
'
    cmdTop.Left = 360 + (9960 - 360) * HyoujiBairitu + FLGoffsetX
    cmdTop.Top = 480 + (600 - 480) * HyoujiBairitu + FLGoffsetY
    cmdTop.FontSize = 10 * HyoujiBairitu
    cmdTop.Width = 735 * HyoujiBairitu
    cmdTop.Height = 255 * HyoujiBairitu
'
    cmd5UP.Left = 360 + (9960 - 360) * HyoujiBairitu + FLGoffsetX
    cmd5UP.Top = 480 + (840 - 480) * HyoujiBairitu + FLGoffsetY
    cmd5UP.FontSize = 10 * HyoujiBairitu
    cmd5UP.Width = 735 * HyoujiBairitu
    cmd5UP.Height = 375 * HyoujiBairitu
'
    cmdUP.Left = 360 + (9960 - 360) * HyoujiBairitu + FLGoffsetX
    cmdUP.Top = 480 + (1320 - 480) * HyoujiBairitu + FLGoffsetY
    cmdUP.FontSize = 10 * HyoujiBairitu
    cmdUP.Width = 735 * HyoujiBairitu
    cmdUP.Height = 495 * HyoujiBairitu
'
    cmdDOWN.Left = 360 + (9960 - 360) * HyoujiBairitu + FLGoffsetX
    cmdDOWN.Top = 480 + (1920 - 480) * HyoujiBairitu + FLGoffsetY
    cmdDOWN.FontSize = 10 * HyoujiBairitu
    cmdDOWN.Width = 735 * HyoujiBairitu
    cmdDOWN.Height = 495 * HyoujiBairitu
'
    cmd5DOWN.Left = 360 + (9960 - 360) * HyoujiBairitu + FLGoffsetX
    cmd5DOWN.Top = 480 + (2520 - 480) * HyoujiBairitu + FLGoffsetY
    cmd5DOWN.FontSize = 10 * HyoujiBairitu
    cmd5DOWN.Width = 735 * HyoujiBairitu
    cmd5DOWN.Height = 375 * HyoujiBairitu
'
    cmdBottom.Left = 360 + (9960 - 360) * HyoujiBairitu + FLGoffsetX
    cmdBottom.Top = 480 + (2895 - 480) * HyoujiBairitu + FLGoffsetY
    cmdBottom.FontSize = 10 * HyoujiBairitu
    cmdBottom.Width = 735 * HyoujiBairitu
    cmdBottom.Height = 255 * HyoujiBairitu
'
    lblComment.Left = 360 + (6480 - 360) * HyoujiBairitu + FLGoffsetX
    lblComment.Top = 480 + (2760 - 480) * HyoujiBairitu + FLGoffsetY
    lblComment.FontSize = 10 * HyoujiBairitu
    lblComment.Width = 2055 * HyoujiBairitu
    lblComment.Height = 255 * HyoujiBairitu
'
    cmdCancel.Left = 360 + (4680 - 360) * HyoujiBairitu + FLGoffsetX
    cmdCancel.Top = 480 + (3240 - 480) * HyoujiBairitu + FLGoffsetY
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1455 * HyoujiBairitu
    cmdCancel.Height = 495 * HyoujiBairitu
'
    cmdUpdate.Left = 360 + (6480 - 360) * HyoujiBairitu + FLGoffsetX
    cmdUpdate.Top = 480 + (3240 - 480) * HyoujiBairitu + FLGoffsetY
    cmdUpdate.FontSize = 10 * HyoujiBairitu
    cmdUpdate.Width = 1455 * HyoujiBairitu
    cmdUpdate.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 360 + (8280 - 360) * HyoujiBairitu + FLGoffsetX
    cmdQuit.Top = 480 + (3240 - 480) * HyoujiBairitu + FLGoffsetY
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1455 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
End Sub

Private Sub MENU_settei()       '*** ÉÅÉjÉÖÅ[èÛë‘ê›íË ***
'
    If FLGconst = 1 Then        '*** ç\ê¨ï\âÊñ ë∂ç› ***
        Me.mnuKousei.Checked = True
        Me.mnuKouseiP.Checked = True
    Else
        Me.mnuKousei.Checked = False
        Me.mnuKouseiP.Checked = False
    End If
'
    If FLGplst = 1 Then         '*** ïîïiï\âÊñ ë∂ç› ***
        Me.mnuBuhin.Checked = True
        Me.mnuBuhinP.Checked = True
    Else
        Me.mnuBuhin.Checked = False
        Me.mnuBuhinP.Checked = False
    End If
'
    If FLGplst2 = 1 Then        '*** ïîïiï\âÊñ ÇQë∂ç› ***
        Me.mnuBuhin2.Checked = True
        Me.mnuBuhin2P.Checked = True
    Else
        Me.mnuBuhin2.Checked = False
        Me.mnuBuhin2P.Checked = False
    End If
'
    If FLGplst = 1 And FLGplst2 = 1 Then    '*** ïîïiï\ÇQâÊñ Ç∆Ç‡ä˘Ç…äJÇ¢ÇƒÇ¢ÇÈ ***
        Me.mnuORCAD.Enabled = False
        Me.mnuORCADP.Enabled = False
    Else
        Me.mnuORCAD.Enabled = True
        Me.mnuORCADP.Enabled = True
    End If
'
    If FLGplstWork = 1 Then        '*** OrCADïœä∑çÏã∆ÉtÉ@ÉCÉã ï“èWâÊñ ë∂ç› ***
        Me.mnuConvFile.Checked = True
    Else
        Me.mnuConvFile.Checked = False
    End If
'
    If FLGmaker = 1 Then       '*** ÉÅÅ[ÉJÅ[âÊñ ë∂ç› ***
        Me.mnuMakerment.Checked = True
        Me.mnuMakermentP.Checked = True
    Else
        Me.mnuMakerment.Checked = False
        Me.mnuMakermentP.Checked = False
    End If
'
    If FLGtrader = 1 Then       '*** è§é–âÊñ ë∂ç› ***
        Me.mnuTraderment.Checked = True
        Me.mnuTradermentP.Checked = True
    Else
        Me.mnuTraderment.Checked = False
        Me.mnuTradermentP.Checked = False
    End If
'
    If FLGitem = 1 Then         '*** ïîïiÉRÅ[ÉhçÄñ⁄âÊñ ë∂ç› ***
        Me.mnuCode.Checked = True
        Me.mnuHinsyu.Enabled = True
'
        Me.mnuCodeP.Checked = True
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


