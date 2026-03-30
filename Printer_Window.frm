VERSION 5.00
Begin VB.Form Printer_Window 
   BackColor       =   &H00004000&
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "プリンタ 選択"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル(&Q)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "印刷実行(&P)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ListBox lstPrinters 
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
      Height          =   1425
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Shape Shape_waku 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblAvailable 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "利用可能"
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
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblComment 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblPrinterSelect 
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "出力先："
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
      TabIndex        =   3
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "Printer_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************
'*** プリンタ選択 ***
'********************
'
Option Explicit
'
Private Sub Form_Load()
'   プリンターを ListBox に列挙する
    Dim objPrinter As Printer
    Dim i As Integer
'
    Me.Width = 480 + (6255 - 960) * HyoujiBairitu + 480
    Me.Height = 480 + (4200 - 960) * HyoujiBairitu + 480
    Printer_Window.BackColor = &H206000
'
    Call DSPgamenBuhin
'
    strMyPrinter = Printer.DeviceName '現在のプリンター
    lblPrinterSelect.Caption = "出力先： " & strMyPrinter
    lblAvailable.Caption = vbCrLf & "利用" & vbCrLf & "可能" & vbCrLf & vbCrLf & "プリンタ"
    lstPrinters.Clear
'
    For Each objPrinter In Printers
        lstPrinters.AddItem objPrinter.DeviceName
        If strMyPrinter = objPrinter.DeviceName Then
            lstPrinters.Selected(i) = True '現在のプリンタを選択状態に
        End If
        i = i + 1
    Next
'
    lblComment.Caption = "選択したプリンタの設定はコントロールパネルの" & vbCrLf & "「プリンタとＦＡＸ」の設定で行います。"
    flag_cancel = True
End Sub

Private Sub lstPrinters_DblClick()
    strMyPrinter = lstPrinters.List(lstPrinters.ListIndex)  '指定のプリンタ
    lblPrinterSelect.Caption = "出力先： " & strMyPrinter
End Sub

Private Sub cmdCancel_Click()
    flag_cancel = True
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    flag_cancel = False
    Unload Me
End Sub
 
Private Sub DSPgamenBuhin()
    Shape_waku.BorderWidth = 5
    Shape_waku.Top = 64
    Shape_waku.Left = 65
    Shape_waku.Width = Me.Width - 80 - 130
    Shape_waku.Height = Me.Height - 360 - 128
'
    lblPrinterSelect.Left = 480
    lblPrinterSelect.Top = 240 + (360 - 240) * HyoujiBairitu
    lblPrinterSelect.FontSize = 10 * HyoujiBairitu
    lblPrinterSelect.Width = 5175 * HyoujiBairitu
    lblPrinterSelect.Height = 265 * HyoujiBairitu
'
    lstPrinters.Left = 480 + (1200 - 480) * HyoujiBairitu
    lstPrinters.Top = 240 + (720 - 240) * HyoujiBairitu
    lstPrinters.FontSize = 10 * HyoujiBairitu
    lstPrinters.Width = 4455 * HyoujiBairitu
    lstPrinters.Height = 1490 * HyoujiBairitu
'
    lblAvailable.Left = 480
    lblAvailable.Top = 240 + (720 - 240) * HyoujiBairitu
    lblAvailable.FontSize = 10 * HyoujiBairitu
    lblAvailable.Width = 735 * HyoujiBairitu
    lblAvailable.Height = lstPrinters.Height
'
    lblComment.Left = 480
    lblComment.Top = 240 + (2280 - 240) * HyoujiBairitu
    lblComment.FontSize = 10 * HyoujiBairitu
    lblComment.Width = 5175 * HyoujiBairitu
    lblComment.Height = 480 * HyoujiBairitu
'
    cmdPrint.Left = 480 + (1200 - 480) * HyoujiBairitu
    cmdPrint.Top = 240 + (3000 - 240) * HyoujiBairitu
    cmdPrint.FontSize = 10 * HyoujiBairitu
    cmdPrint.Width = 1575 * HyoujiBairitu
    cmdPrint.Height = 495 * HyoujiBairitu
'
    cmdCancel.Left = 480 + (3360 - 480) * HyoujiBairitu
    cmdCancel.Top = 240 + (3000 - 240) * HyoujiBairitu
    cmdCancel.FontSize = 10 * HyoujiBairitu
    cmdCancel.Width = 1575 * HyoujiBairitu
    cmdCancel.Height = 495 * HyoujiBairitu
End Sub

