VERSION 5.00
Begin VB.Form Pcod_main_c 
   BackColor       =   &H00004000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "部品コード ＜品目 帳票形式＞"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Pcod_main_c.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6375
   ScaleWidth      =   10215
   Begin VB.ComboBox cboMSLevel 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6600
      TabIndex        =   62
      Text            =   "cboBsitei"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSakujyo 
      Caption         =   "項目削除(&D)"
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
      Left            =   1080
      TabIndex        =   55
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox cboKeijou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Text            =   "cboKeijou"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cboCad 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3840
      TabIndex        =   52
      Text            =   "cboCad"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cboBikouran 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Pcod_main_c.frx":030A
      Left            =   5160
      List            =   "Pcod_main_c.frx":030C
      TabIndex        =   49
      Text            =   "cboBikouran"
      Top             =   3480
      Width           =   3855
   End
   Begin VB.ComboBox cboSyukko 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   46
      Text            =   "cboSyukko"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.ComboBox cboBsitei 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5160
      TabIndex        =   18
      Text            =   "cboBsitei"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox cboMaker0 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      Text            =   "cboMaker0"
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdKettei 
      Caption         =   "決定保存(&G)"
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
      TabIndex        =   58
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdTuika 
      Caption         =   "品目追加(&A)"
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
      Left            =   1080
      TabIndex        =   56
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdKousin 
      Caption         =   "内容変更(&C)"
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
      TabIndex        =   57
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "閉じる(&Q)"
      Height          =   495
      Left            =   7800
      TabIndex        =   61
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   6360
      TabIndex        =   60
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   495
      Left            =   5160
      TabIndex        =   59
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      MousePointer    =   1  '矢印
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "Text26"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5040
      MousePointer    =   1  '矢印
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text25"
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6240
      MousePointer    =   1  '矢印
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "Text24"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      MousePointer    =   1  '矢印
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "Text23"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9000
      MousePointer    =   1  '矢印
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "Text22"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "Text21"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      MousePointer    =   1  '矢印
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "Text20"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      MousePointer    =   1  '矢印
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Text19"
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "Text18"
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "Text17"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6240
      MousePointer    =   1  '矢印
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "Text16"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      MousePointer    =   1  '矢印
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Text15"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Text14"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      MousePointer    =   1  '矢印
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "Text13"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "Text12"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      MousePointer    =   1  '矢印
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "Text11"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '右揃え
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      MousePointer    =   1  '矢印
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "Text10"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      MousePointer    =   1  '矢印
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "Text9"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      MousePointer    =   1  '矢印
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Text8"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      MousePointer    =   1  '矢印
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text7"
      Top             =   1440
      Width           =   5775
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MaxLength       =   256
      MousePointer    =   1  '矢印
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Text6"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      MousePointer    =   1  '矢印
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      MousePointer    =   1  '矢印
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MousePointer    =   1  '矢印
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "更新年月日"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      TabIndex        =   53
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "足ピン数"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   28
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label21 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "MTBF FIT表"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   30
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "形  状"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   21
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "CAD登録"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   50
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "特記事項"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      TabIndex        =   47
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "出庫非出庫"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   44
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "半田 耐熱温度 ℃"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      TabIndex        =   26
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "発注単位"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   32
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "標準在庫数"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      TabIndex        =   42
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "端子ﾒｯｷ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   24
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "現品表示"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "在庫数"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   40
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "予約数"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   38
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "MSL(吸湿ﾚﾍﾞﾙ)"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      TabIndex        =   36
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "平均単価"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   34
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "代品ｺｰﾄﾞ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "部品指定"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ﾒｰｶｰ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品目名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "品種名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ｺｰﾄﾞ番号"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "項目名"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Pcod_main_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'********************
'* 品目 細目一覧表示 *
'********************
'
Option Explicit
'
    Dim HeadTitle As String
    Dim FLG_CorA As Integer
    Dim FLG_Motonoiro As Long

Private Sub Form_Initialize()
    HeadTitle = "部品ｺｰﾄﾞ <品目 帳票形式>"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If FLG_CorA = 1 Or FLG_CorA = 2 Then Exit Sub   '*** 変更追加画面の時はキーが効かない ***
'
    Select Case KeyCode
    Case vbKeyPageDown      '*** Roll Up
        Call Hup
        Call DSPlevel31     '*** 品目名表示 ***
'
    Case vbKeyUp            '*** ↑ ***
        Call H1up
        Call DSPlevel31     '*** 品目名表示 ***
'
    Case vbKeyPageUp        '*** Roll Down
        Call Hdown
        Call DSPlevel31     '*** 品目名表示 ***
'
    Case vbKeyDown          '*** ↓ ***
        Call H1down
        Call DSPlevel31     '*** 品目名表示 ***
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim makername As String
'                   フォームの表示位置の設定
    Width = 360 + (10305 - 720) * HyoujiBairitu + 360
    Height = 360 + (6915 - 720) * HyoujiBairitu + 360
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
    FLG_Motonoiro = Me.BackColor
'
    Me.Caption = HeadTitle
'
    If Aitem0(ips0, 0) = "IC" Then
        DRVmain0 = Xcont0(2) & "\IC\IC" & Left$(Bindex0(jps0, 0), 1) _
        & "\L" & Bindex0(jps0, 0) & ".COD"
    Else
        DRVmain0 = Xcont0(2) & "\" & Aitem0(ips0, 0) _
        & "\L" & Bindex0(jps0, 0) & ".COD"
    End If
'
    Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)   '*** メインコード読み込み ***
'
    Text1.Text = Aitem0(ips0, 0)
    Text2.Text = Aitem0(ips0, 1)
    Text4.Text = Bindex0(jps0, 3) & "xxx" & Bindex0(jps0, 4)
    Text5.Text = Bindex0(jps0, 1)
'
    DSPlevel31      '*** 品目一覧表示 ***
'
    Call DSPpointer(1)  '*** マウスポインターOFF ***
'
    If Xcont0(8) = "_" Then
        cmdKettei.Visible = True
        cmdKettei.Enabled = False
        cmdSakujyo.Visible = True
        cmdSakujyo.Enabled = True
        cmdTuika.Enabled = True
        cmdkousin.Enabled = True
    Else
        cmdKettei.Visible = False
        cmdSakujyo.Visible = False
        cmdTuika.Enabled = False
        cmdkousin.Enabled = False
    End If
'
    cboMaker0.Visible = False
    cboBsitei.Visible = False
    cboSyukko.Visible = False
    cboBikouran.Visible = False
    cboCad.Visible = False
    cboKeijou.Visible = False
    cboMSLevel.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
'
    End If
End Sub

Private Sub DSPlevel31()
                    '*** 品目一覧表示 ***
    Dim makername As String
    Dim Dtemp As String
'
    Text3.Text = "L" & Bindex0(jps0, 0) & "-" & Cmain0(kps0, 0)
'
    If Bindex0(jps0, 5) = "000" Then
        makername = Cmain0(kps0, 13)
    ElseIf Bindex0(jps0, 5) = "998" Then
        makername = Bindex0(jps0, 8)
    Else
        makername = Bindex0(jps0, 5)
    End If
'
    Call Makerget1(makername)      '***ﾒｰｶｰ名取得 ***
    Text25.Text = makername
'
    Text6.Text = Cmain0(kps0, 1)
    Text7.Text = Cmain0(kps0, 2)
        Dtemp = Cmain0(kps0, 3)
        Call TRSsitei(Dtemp)
    Text8.Text = Dtemp
    Text9.Text = Cmain0(kps0, 4)
    Text10.Text = Format(Cmain0(kps0, 5), "#,##0.00")
        Dtemp = Cmain0(kps0, 6)     '***Ver2.00ｶﾗ MSLﾄｼﾃ 使用
        Call TRS_Mlevel(Dtemp)
    Text11.Text = Dtemp
    Text12.Text = Format(Cmain0(kps0, 7), "#,##0")
    Text13.Text = Format(Cmain0(kps0, 8), "#,##0")
    Text14.Text = Cmain0(kps0, 9)
        If Cmain0(kps0, 19) = "0" Then
            Cmain0(kps0, 19) = "*"
        End If
    Text15.Text = Cmain0(kps0, 19)
    Text16.Text = Format(Cmain0(kps0, 14), "#,##0")
    Text17.Text = Format(Cmain0(kps0, 15), "#,##0")
        Dtemp = Cmain0(kps0, 12)
        Call TRSsyukko(Dtemp)
    Text18.Text = Dtemp
        Dtemp = Cmain0(kps0, 16)
        Call TRSbikouran(Dtemp)
    Text19.Text = Dtemp
        Dtemp = Cmain0(kps0, 17)
        Call TRStouroku(Dtemp)
    Text20.Text = Dtemp
        Dtemp = Cmain0(kps0, 18)
        Call TRSkeijou(Dtemp)
    Text21.Text = Dtemp
    Text22.Text = Cmain0(kps0, 20)
    Text23.Text = Format(Cmain0(kps0, 21), "#,##0")
    Text24.Text = Cmain0(kps0, 10)
    Text26.Text = Cmain0(kps0, 11)
End Sub

Private Sub DSPpointer(X As Integer)
    If X = 1 Then
        Text1.MousePointer = 1
        Text2.MousePointer = 1
        Text3.MousePointer = 1
        Text4.MousePointer = 1
        Text5.MousePointer = 1
        Text6.MousePointer = 1
        Text7.MousePointer = 1
        Text8.MousePointer = 1
        Text9.MousePointer = 1
        Text10.MousePointer = 1
        Text11.MousePointer = 1
        Text12.MousePointer = 1
        Text13.MousePointer = 1
        Text14.MousePointer = 1
        Text15.MousePointer = 1
        Text16.MousePointer = 1
        Text17.MousePointer = 1
        Text18.MousePointer = 1
        Text19.MousePointer = 1
        Text20.MousePointer = 1
        Text21.MousePointer = 1
        Text22.MousePointer = 1
        Text23.MousePointer = 1
        Text24.MousePointer = 1
        Text25.MousePointer = 1
        Text26.MousePointer = 1
    Else
        Text3.MousePointer = 0
        Text6.MousePointer = 0
        Text7.MousePointer = 0
        Text9.MousePointer = 0
        Text10.MousePointer = 0
        Text11.MousePointer = 0
        Text12.MousePointer = 0
        Text13.MousePointer = 0
        Text15.MousePointer = 0
        Text16.MousePointer = 0
        Text17.MousePointer = 0
        Text22.MousePointer = 0
        Text23.MousePointer = 0
        Text26.MousePointer = 0
    End If
End Sub

Private Sub cmdDown_Click()
    Call H1down         '*** 一つ下へ ***
    Call DSPlevel31     '*** 品目名表示 ***
'
    Text3.SetFocus
End Sub

Private Sub cmdKettei_Click()
    Dim Tempcode As String
    Dim i As Integer, j As Integer
'
    Select Case FLG_CorA
    Case 1
        GET_Chg_Data    '*** 変更データ取り込み ***
'
    Case 2
        Tempcode = Mid(Trim(Text3.Text), 7)
        i = Len(Tempcode)
        If i <> 2 Or Val(Tempcode) < 0 Or Val(Tempcode) > 100 Then
            j = MsgBox("正しいコード番号を記入してください。", vbCritical)
            Exit Sub
'
        End If
'
        For i = 1 To Cnum0
            If Cmain0(i, 0) = Tempcode Then
                j = MsgBox("同じコード番号がすでにあります。", vbExclamation)
                Exit Sub
'
            End If
        Next i
'
        If Cnum0 = 0 Then
            kps0 = 1     '*** 初めてのデータ ***
            Cnum0 = 1
'
        Else
            For i = 1 To Cnum0
                If Cmain0(i, 0) > Tempcode Then
                Exit For
'
                End If
            Next i
'
            kps0 = i
            Call Ins_cmain      '*** データを１つずらす ***
        End If
'
        Cmain0(kps0, 0) = Tempcode
        Call GET_Chg_Data       '*** 変更データ取り込み ***
'
    End Select
'
    Me.MousePointer = vbHourglass
    DoEvents
'                               '*** 表示を元に戻す ***
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.TabStop = False
    Text7.TabStop = False
    Text9.TabStop = False
    Text10.TabStop = False
    Text11.TabStop = False
    Text12.TabStop = False
    Text13.TabStop = False
    Text14.TabStop = False
    Text15.TabStop = False
    Text16.TabStop = False
    Text17.TabStop = False
    Text22.TabStop = False
    Text23.TabStop = False
    Text26.TabStop = False
    Me.BackColor = FLG_Motonoiro
'
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    cmdkousin.Enabled = True
    cmdTuika.Enabled = True
    cmdSakujyo.Enabled = True
    cmdKettei.Enabled = False
'
    Text24.Enabled = True
'
    Text25.Visible = True
    Text25.Enabled = True
    cboMaker0.Visible = False
'
    Text8.Visible = True
    cboBsitei.Visible = False
'
    Text18.Visible = True
    cboSyukko.Visible = False
'
    Text19.Visible = True
    cboBikouran.Visible = False
'
    Text20.Visible = True
    cboCad.Visible = False
'
    Text21.Visible = True
    cboKeijou.Visible = False
'
    Text11.Visible = True
    cboMSLevel.Visible = False
'
    Call WRmain(DRVmain0, Cmain0(), Cnum0, Cdim0)   '*** MAINデータセーブ ***
    DoEvents
    Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)   '*** MAINデータ再読み込み ***
    DoEvents
    Call DSPlevel31     '*** 品目一覧表示 ***
    DoEvents
'
    Call DSPpointer(1)      '*** マウスポインタセット ***
'
    Me.MousePointer = vbDefault
    FLG_CorA = 0
    FLGmain_data_change = 1 '*** 変更フラグ設定(メイン画面の表示を変更する)***
'
    Text3.SetFocus
End Sub

Private Sub cmdSakujyo_Click()
    Dim i As Integer, j As Integer
'
    If Cnum0 = 1 Then
        i = MsgBox("最後の１品目は削除できません。", vbCritical)
    Else
        i = MsgBox("品目 < L" & Bindex0(jps0, 0) & "-" & Cmain0(kps0, 0) & " > を削除して良ろしいですか？", vbYesNo)
        If i = vbYes Then
            Me.MousePointer = vbHourglass
            DoEvents
'
            For i = kps0 To Cnum0 - 1
                For j = 0 To Cdim0
                    Cmain0(i, j) = Cmain0(i + 1, j)
                Next j
            Next i
            If kps0 = Cnum0 Then
                kps0 = Cnum0 - 1
            End If
            Cnum0 = Cnum0 - 1
            DoEvents
'
            Call WRmain(DRVmain0, Cmain0(), Cnum0, Cdim0)    '*** 品種データセーブ ***
            DoEvents
            Call RDmain(DRVmain0, Cmain0(), Cnum0, Cdim0)    '*** 品種データ再読み込み ***
            DoEvents
            Call DSPlevel31
            FLGmain_data_change = 1 '*** 変更フラグ設定(メイン画面の表示を変更する)***
'
            Me.MousePointer = vbDefault
        End If
    End If
'
    Text6.SetFocus
End Sub

Private Sub cmdTuika_Click()
    FLG_CorA = 2
    Set_Gamen_CorA      '*** 追加画面設定 ***
'
    Call DSPpointer(0)  '*** マウスポインタセット ***
'
    Text3.SetFocus
End Sub

Private Sub cmdUp_Click()
    Call H1up           '*** 一つ上へ ***
    Call DSPlevel31     '*** 品目名表示 ***
'
    Text3.SetFocus
End Sub

Private Sub cmdkousin_Click()
    FLG_CorA = 1
    Set_Gamen_CorA      '*** 変更画面設定 ***
'
    Call DSPpointer(0)  '*** マウスポインタセット ***
'
    Text6.SetFocus
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub GET_Chg_Data()      '*** 変更データ取り込み ***
    Dim i As Integer
    Dim TempAno As Integer, TempBno As Integer, TempCno As Integer
    Dim Dtemp As String
'
    If Bindex0(jps0, 5) = "000" Then
        Cmain0(kps0, 13) = Maker(cboMaker0.ListIndex, 0)
    Else
        Cmain0(kps0, 13) = "*"
    End If
'
    Cmain0(kps0, 1) = Trim(Text6.Text)
    Cmain0(kps0, 2) = Trim(Text7.Text)
    TempAno = cboBsitei.ListIndex
    If TempAno = 5 Then
        Cmain0(kps0, 3) = "*"
    Else
        Cmain0(kps0, 3) = Trim(str(TempAno))
    End If
            Dtemp = Cmain0(kps0, 3)
        Call TRSsitei(Dtemp)
    Text8.Text = Dtemp
    Cmain0(kps0, 4) = Trim(Text9.Text)
'
    Cmain0(kps0, 11) = Trim(Text26.Text)
'
            Dtemp = Text10.Text
        i = InStr(1, Dtemp, ",", vbTextCompare)     '1,000/1,000,000
        If i <> 0 Then
            Dtemp = Left(Dtemp, i - 1) & Mid(Dtemp, i + 1)
        End If
'
        i = InStr(1, Dtemp, ",", vbTextCompare)     '1,000
        If i <> 0 Then
            Dtemp = Left(Dtemp, i - 1) & Mid(Dtemp, i + 1)
        End If
    Cmain0(kps0, 5) = Trim(str(Val(Dtemp)))
'
            TempAno = cboMSLevel.ListIndex
            Dtemp = Trim(str(TempAno))
    Cmain0(kps0, 6) = Dtemp
        Call TRS_Mlevel(Dtemp)
    Text11.Text = Dtemp
    Cmain0(kps0, 7) = Trim(Text12.Text)
'
            Dtemp = Text13.Text
        i = InStr(1, Dtemp, ",", vbTextCompare)
        If i <> 0 Then
            Dtemp = Left(Dtemp, i - 1) & Mid(Dtemp, i + 1)
        End If
    Cmain0(kps0, 8) = Trim(str(Val(Dtemp)))
    Cmain0(kps0, 9) = Trim(Text14.Text)
    Cmain0(kps0, 14) = Trim(Text16.Text)
'
            Dtemp = Text17.Text
        i = InStr(1, Dtemp, ",", vbTextCompare)
        If i <> 0 Then
            Dtemp = Left(Dtemp, i - 1) & Mid(Dtemp, i + 1)
        End If
    Cmain0(kps0, 15) = Trim(str(Val(Dtemp)))
    TempAno = cboBsitei.ListIndex
    If TempAno = 5 Then
        Cmain0(kps0, 3) = "*"
    Else
        Cmain0(kps0, 3) = Trim(str(TempAno))
    End If
            Dtemp = Cmain0(kps0, 3)
        Call TRSsitei(Dtemp)
    Text8.Text = Dtemp
'
    TempAno = cboSyukko.ListIndex
    If TempAno = 0 Then
        Cmain0(kps0, 12) = "1"
    Else
        Cmain0(kps0, 12) = "0"
    End If
            Dtemp = Cmain0(kps0, 12)
        Call TRSsyukko(Dtemp)
    Text18.Text = Dtemp
'
    TempAno = cboBikouran.ListIndex
    If TempAno = 0 Then
        Cmain0(kps0, 16) = "1"
    Else
        Cmain0(kps0, 16) = "0"
    End If
            Dtemp = Cmain0(kps0, 16)
        Call TRSbikouran(Dtemp)
    Text19.Text = Dtemp
'
    TempAno = cboCad.ListIndex
    If TempAno = 0 Then
        Cmain0(kps0, 17) = "*"
    Else
        Cmain0(kps0, 17) = "1"
    End If
            Dtemp = Cmain0(kps0, 17)
        Call TRStouroku(Dtemp)
    Text20.Text = Dtemp
'
            TempAno = cboKeijou.ListIndex
        Call TRSkeijouKigou(TempAno, Dtemp)
    Cmain0(kps0, 18) = Dtemp
            
        Call TRSkeijou(Dtemp)
    Text21.Text = Dtemp
'
    Cmain0(kps0, 20) = Trim(Text22.Text)
'
            Dtemp = Text23.Text
        i = InStr(1, Dtemp, ",", vbTextCompare)     '1,000
        If i <> 0 Then
            Dtemp = Left(Dtemp, i - 1) & Mid(Dtemp, i + 1)
        End If
    Cmain0(kps0, 21) = Trim(str(Val(Dtemp)))
    Cmain0(kps0, 19) = Trim(Text15.Text)
    Cmain0(kps0, 10) = Format(Date, "yy/mm/dd")
End Sub

Private Sub Set_Gamen_CorA()    '*** 追加変更画面設定 ***
    Dim i As Integer
    Dim Mdata As String, Dtemp As String
'
    Text1.Enabled = False
    Text2.Enabled = False
    If FLG_CorA = 1 Then
        Text3.Enabled = False
    Else
        Text3.Enabled = True
        Text3.TabStop = True
    End If
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.TabStop = True
    Text7.TabStop = True
    Text9.TabStop = True
    Text10.TabStop = True
    Text11.TabStop = True
    Text12.TabStop = True
    Text13.TabStop = True
    Text14.TabStop = True
    Text15.TabStop = True
    Text16.TabStop = True
    Text17.TabStop = True
    Text22.TabStop = True
    Text23.TabStop = True
    Text24.Enabled = False
    Text26.TabStop = True
    Me.BackColor = &H808080
'
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If FLG_CorA = 1 Then
        cmdkousin.Enabled = True
        cmdTuika.Enabled = False
    Else
        cmdkousin.Enabled = False
        cmdTuika.Enabled = True
    End If
    cmdSakujyo.Enabled = False
    cmdKettei.Enabled = True
'
    If Bindex0(jps0, 5) = "000" Then  '*** メーカー個別指定 ***
        Text25.Visible = False
        cboMaker0.Visible = True
        cboMaker0.Clear
        cboMaker0.AddItem "*"
        For i = 1 To Maknum0
            cboMaker0.AddItem Maker(i, 2)
        Next i
'
        For i = 1 To Maknum0
            If Cmain0(kps0, 13) = Maker(i, 0) Then
                cboMaker0.ListIndex = i
                Exit For
            End If
            cboMaker0.ListIndex = 0
        Next i
    Else
        Text25.Enabled = False
    End If
'
    Text8.Visible = False       '*** 部品指定 ***
    cboBsitei.Visible = True
    cboBsitei.Clear
'
    Dtemp = "0"
        Call TRSsitei(Dtemp)
            cboBsitei.AddItem Dtemp
    Dtemp = "1"
        Call TRSsitei(Dtemp)
            cboBsitei.AddItem Dtemp
    Dtemp = "2"
        Call TRSsitei(Dtemp)
            cboBsitei.AddItem Dtemp
    Dtemp = "3"
        Call TRSsitei(Dtemp)
            cboBsitei.AddItem Dtemp
    Dtemp = "4"
        Call TRSsitei(Dtemp)
            cboBsitei.AddItem Dtemp
    Dtemp = "*"
        Call TRSsitei(Dtemp)
            cboBsitei.AddItem Dtemp
    If Cmain0(kps0, 3) = "*" Then
        cboBsitei.ListIndex = 5
    Else
        cboBsitei.ListIndex = Val(Cmain0(kps0, 3))
    End If
'
    Text18.Visible = False      '*** 出庫/非出庫 ***
    cboSyukko.Visible = True
    cboSyukko.Clear
'
    Dtemp = "1"
        Call TRSsyukko(Dtemp)
            cboSyukko.AddItem Dtemp
    Dtemp = "0"
        Call TRSsyukko(Dtemp)
            cboSyukko.AddItem Dtemp
    If Cmain0(kps0, 12) = "1" Then
        cboSyukko.ListIndex = 0
    Else
        cboSyukko.ListIndex = 1
    End If
'
    Text19.Visible = False      '*** 特記事項/備考欄印刷 ***
    cboBikouran.Visible = True
    cboBikouran.Clear
'
    Dtemp = "1"
        Call TRSbikouran(Dtemp)
            cboBikouran.AddItem Dtemp
    Dtemp = "0"
        Call TRSbikouran(Dtemp)
            cboBikouran.AddItem Dtemp
    If Cmain0(kps0, 16) = "1" Then
        cboBikouran.ListIndex = 0
    Else
        cboBikouran.ListIndex = 1
    End If
'
    Text20.Visible = False      '*** CAD登録 ***
    cboCad.Visible = True
    cboCad.Clear
'
    Dtemp = "*"
        Call TRStouroku(Dtemp)
            cboCad.AddItem Dtemp
    Dtemp = "1"
        Call TRStouroku(Dtemp)
            cboCad.AddItem Dtemp
'
    If Cmain0(kps0, 17) = "1" Then
        cboCad.ListIndex = 1
    Else
        cboCad.ListIndex = 0
    End If
'
    Text21.Visible = False      '*** 形状 ***
    cboKeijou.Visible = True
    cboKeijou.Clear
'
    Dtemp = "0"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "1"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "2"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "3"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "4"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "5"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "6"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "7"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "8"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "9"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "A"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "B"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "C"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "D"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "E"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "F"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
    Dtemp = "*"
        Call TRSkeijou(Dtemp)
            cboKeijou.AddItem Dtemp
'
    Dtemp = Cmain0(kps0, 18)
        Call TRSkeijouNo(Dtemp)
    cboKeijou.ListIndex = Val(Dtemp)

'
    Text11.Visible = False      '*** MSL ***
    cboMSLevel.Visible = True
    cboMSLevel.Clear
'
    Dtemp = "0"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "1"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "2"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "3"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "4"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "5"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "6"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "7"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "8"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "9"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "10"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "11"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "12"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "13"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "14"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
    Dtemp = "*"
        Call TRS_Mlevel(Dtemp)
            cboMSLevel.AddItem Dtemp
'
    Dtemp = Cmain0(kps0, 6)
    If Dtemp = "*" Then
        Dtemp = "15"
    End If
    cboMSLevel.ListIndex = Val(Dtemp)
End Sub

Private Sub Ins_cmain()
    Dim i As Integer, j As Integer
'
    For i = Cnum0 To kps0 Step -1
        For j = 0 To Cdim0
            Cmain0(i + 1, j) = Cmain0(i, j)
        Next j
    Next i
    Cnum0 = Cnum0 + 1
End Sub

Private Sub Hup()       '*** 10上へ ***
    If kps0 > 10 Then
        kps0 = kps0 - 10
    Else
        Beep
        kps0 = 1
    End If
End Sub

Private Sub Hdown()     '*** 10下へ ***
    If kps0 + 10 <= Cnum0 Then
        kps0 = kps0 + 10
    Else
        Beep
        kps0 = Cnum0
    End If
End Sub

Private Sub H1down()    '*** 1下へ ***
    If kps0 < Cnum0 Then
        kps0 = kps0 + 1
    Else
        Beep
    End If
End Sub

Private Sub H1up()      '*** 1上へ ***
    If kps0 > 1 Then
        kps0 = kps0 - 1
    Else
        Beep
    End If
End Sub

Private Sub DSPgamenBuhin()
    Text1.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text1.Top = 360
    Text1.FontSize = 10 * HyoujiBairitu
    Text1.Width = 375 * HyoujiBairitu
    Text1.Height = 285 * HyoujiBairitu
'
    Label1.Left = 360
    Label1.Top = 360
    Label1.FontSize = 10 * HyoujiBairitu
    Label1.Width = 1095 * HyoujiBairitu
    Label1.Height = Text1.Height
'
    Text2.Left = 360 + (1920 - 360) * HyoujiBairitu
    Text2.Top = 360
    Text2.FontSize = 10 * HyoujiBairitu
    Text2.Width = 2055 * HyoujiBairitu
    Text2.Height = Text1.Height
'
    Text3.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text3.Top = 360 + (720 - 360) * HyoujiBairitu
    Text3.FontSize = 10 * HyoujiBairitu
    Text3.Width = 1095 * HyoujiBairitu
    Text3.Height = 285 * HyoujiBairitu
'
    Label2.Left = 360
    Label2.Top = 360 + (720 - 360) * HyoujiBairitu
    Label2.FontSize = 10 * HyoujiBairitu
    Label2.Width = 1095 * HyoujiBairitu
    Label2.Height = Text3.Height
'
    Text25.Left = 360 + (5055 - 360) * HyoujiBairitu
    Text25.Top = 360 + (720 - 360) * HyoujiBairitu
    Text25.FontSize = 10 * HyoujiBairitu
    Text25.Width = 3975 * HyoujiBairitu
    Text25.Height = 285 * HyoujiBairitu
'
    Label5.Left = 360 + (4080 - 360) * HyoujiBairitu
    Label5.Top = 360 + (720 - 360) * HyoujiBairitu
    Label5.FontSize = 10 * HyoujiBairitu
    Label5.Width = 975 * HyoujiBairitu
    Label5.Height = Text25.Height
'
    cboMaker0.Left = 360 + (5055 - 360) * HyoujiBairitu
    cboMaker0.Top = 360 + (720 - 360) * HyoujiBairitu
    cboMaker0.FontSize = 10 * HyoujiBairitu
    cboMaker0.Width = 3975 * HyoujiBairitu
'   cboMaker0.Height = 285 * HyoujiBairitu
'
    Text4.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text4.Top = 360 + (1080 - 360) * HyoujiBairitu
    Text4.FontSize = 10 * HyoujiBairitu
    Text4.Width = 2535 * HyoujiBairitu
    Text4.Height = 285 * HyoujiBairitu
'
    Label3.Left = 360
    Label3.Top = 360 + (1080 - 360) * HyoujiBairitu
    Label3.FontSize = 10 * HyoujiBairitu
    Label3.Width = 1095 * HyoujiBairitu
    Label3.Height = Text4.Height
'
    Text5.Left = 360 + (4080 - 360) * HyoujiBairitu
    Text5.Top = 360 + (1080 - 360) * HyoujiBairitu
    Text5.FontSize = 10 * HyoujiBairitu
    Text5.Width = 4935 * HyoujiBairitu
    Text5.Height = Text4.Height
'
    Text6.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text6.Top = 360 + (1440 - 360) * HyoujiBairitu
    Text6.FontSize = 10 * HyoujiBairitu
    Text6.Width = 2535 * HyoujiBairitu
    Text6.Height = 285 * HyoujiBairitu
'
    Label4.Left = 360
    Label4.Top = 360 + (1440 - 360) * HyoujiBairitu
    Label4.FontSize = 10 * HyoujiBairitu
    Label4.Width = 1095 * HyoujiBairitu
    Label4.Height = Text6.Height
'
    Text7.Left = 360 + (4080 - 360) * HyoujiBairitu
    Text7.Top = 360 + (1440 - 360) * HyoujiBairitu
    Text7.FontSize = 10 * HyoujiBairitu
    Text7.Width = 5775 * HyoujiBairitu
    Text7.Height = Text6.Height
'
    Text14.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text14.Top = 360 + (1920 - 360) * HyoujiBairitu
    Text14.FontSize = 10 * HyoujiBairitu
    Text14.Width = 2535 * HyoujiBairitu
    Text14.Height = 285 * HyoujiBairitu
'
    Label12.Left = 360
    Label12.Top = 360 + (1920 - 360) * HyoujiBairitu
    Label12.FontSize = 10 * HyoujiBairitu
    Label12.Width = 1095 * HyoujiBairitu
    Label12.Height = Text14.Height
'
    Text8.Left = 360 + (5175 - 360) * HyoujiBairitu
    Text8.Top = 360 + (1920 - 360) * HyoujiBairitu
    Text8.FontSize = 10 * HyoujiBairitu
    Text8.Width = 2535 * HyoujiBairitu
    Text8.Height = 285 * HyoujiBairitu
'
    Label6.Left = 360 + (4080 - 360) * HyoujiBairitu
    Label6.Top = 360 + (1920 - 360) * HyoujiBairitu
    Label6.FontSize = 10 * HyoujiBairitu
    Label6.Width = 1095 * HyoujiBairitu
    Label6.Height = Text8.Height
'
    cboBsitei.Left = 360 + (5175 - 360) * HyoujiBairitu
    cboBsitei.Top = 360 + (1920 - 360) * HyoujiBairitu
    cboBsitei.FontSize = 10 * HyoujiBairitu
    cboBsitei.Width = 2535 * HyoujiBairitu
'   cboBsitei.Height = 285 * HyoujiBairitu
'
    Text9.Left = 360 + (8775 - 360) * HyoujiBairitu
    Text9.Top = 360 + (1920 - 360) * HyoujiBairitu
    Text9.FontSize = 10 * HyoujiBairitu
    Text9.Width = 1095 * HyoujiBairitu
    Text9.Height = 285 * HyoujiBairitu
'
    Label7.Left = 360 + (7800 - 360) * HyoujiBairitu
    Label7.Top = 360 + (1920 - 360) * HyoujiBairitu
    Label7.FontSize = 10 * HyoujiBairitu
    Label7.Width = 975 * HyoujiBairitu
    Label7.Height = Text9.Height
'
    Text21.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text21.Top = 360 + (2400 - 360) * HyoujiBairitu
    Text21.FontSize = 10 * HyoujiBairitu
    Text21.Width = 1215 * HyoujiBairitu
    Text21.Height = 285 * HyoujiBairitu
'
    Label20.Left = 360
    Label20.Top = 360 + (2400 - 360) * HyoujiBairitu
    Label20.FontSize = 10 * HyoujiBairitu
    Label20.Width = 1095 * HyoujiBairitu
    Label20.Height = Text21.Height
'
    cboKeijou.Left = 360 + (1455 - 360) * HyoujiBairitu
    cboKeijou.Top = 360 + (2400 - 360) * HyoujiBairitu
    cboKeijou.FontSize = 10 * HyoujiBairitu
    cboKeijou.Width = 1095 * HyoujiBairitu
'   cboKeijou.Height = 285 * HyoujiBairitu
'
    Text15.Left = 360 + (3855 - 360) * HyoujiBairitu
    Text15.Top = 360 + (2400 - 360) * HyoujiBairitu
    Text15.FontSize = 10 * HyoujiBairitu
    Text15.Width = 1215 * HyoujiBairitu
    Text15.Height = 285 * HyoujiBairitu
'
    Label13.Left = 360 + (2760 - 360) * HyoujiBairitu
    Label13.Top = 360 + (2400 - 360) * HyoujiBairitu
    Label13.FontSize = 10 * HyoujiBairitu
    Label13.Width = 1095 * HyoujiBairitu
    Label13.Height = Text15.Height
'
    Text26.Left = 360 + (6855 - 360) * HyoujiBairitu
    Text26.Top = 360 + (2400 - 360) * HyoujiBairitu
    Text26.FontSize = 10 * HyoujiBairitu
    Text26.Width = 855 * HyoujiBairitu
    Text26.Height = 285 * HyoujiBairitu
'
    Label16.Left = 360 + (5160 - 360) * HyoujiBairitu
    Label16.Top = 360 + (2400 - 360) * HyoujiBairitu
    Label16.FontSize = 10 * HyoujiBairitu
    Label16.Width = 1695 * HyoujiBairitu
    Label16.Height = Text26.Height
'
    Text23.Left = 360 + (8775 - 360) * HyoujiBairitu
    Text23.Top = 360 + (2400 - 360) * HyoujiBairitu
    Text23.FontSize = 10 * HyoujiBairitu
    Text23.Width = 1095 * HyoujiBairitu
    Text23.Height = 285 * HyoujiBairitu
'
    Label22.Left = 360 + (7800 - 360) * HyoujiBairitu
    Label22.Top = 360 + (2400 - 360) * HyoujiBairitu
    Label22.FontSize = 10 * HyoujiBairitu
    Label22.Width = 975 * HyoujiBairitu
    Label22.Height = Text23.Height
'
    Text17.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text17.Top = 360 + (2880 - 360) * HyoujiBairitu
    Text17.FontSize = 10 * HyoujiBairitu
    Text17.Width = 1215 * HyoujiBairitu
    Text17.Height = 285 * HyoujiBairitu
'
    Label15.Left = 360
    Label15.Top = 360 + (2880 - 360) * HyoujiBairitu
    Label15.FontSize = 10 * HyoujiBairitu
    Label15.Width = 1095 * HyoujiBairitu
    Label15.Height = Text17.Height
'
    Text10.Left = 360 + (3855 - 360) * HyoujiBairitu
    Text10.Top = 360 + (2880 - 360) * HyoujiBairitu
    Text10.FontSize = 10 * HyoujiBairitu
    Text10.Width = 1215 * HyoujiBairitu
    Text10.Height = 285 * HyoujiBairitu
'
    Label8.Left = 360 + (2760 - 360) * HyoujiBairitu
    Label8.Top = 360 + (2880 - 360) * HyoujiBairitu
    Label8.FontSize = 10 * HyoujiBairitu
    Label8.Width = 1095 * HyoujiBairitu
    Label8.Height = Text10.Height
'
    Text11.Left = 360 + (6600 - 360) * HyoujiBairitu
    Text11.Top = 360 + (2760 - 360) * HyoujiBairitu
    Text11.FontSize = 10 * HyoujiBairitu
    Text11.Width = 1095 * HyoujiBairitu
    Text11.Height = 285 * HyoujiBairitu
'
    cboMSLevel.Left = 360 + (6600 - 360) * HyoujiBairitu
    cboMSLevel.Top = 360 + (2760 - 360) * HyoujiBairitu
    cboMSLevel.FontSize = 10 * HyoujiBairitu
    cboMSLevel.Width = 1095 * HyoujiBairitu
'   cboMSLevel.Height = 285 * HyoujiBairitu
'
    Label9.Left = 360 + (5160 - 360) * HyoujiBairitu
    Label9.Top = 360 + (2760 - 360) * HyoujiBairitu
    Label9.FontSize = 10 * HyoujiBairitu
    Label9.Width = 1455 * HyoujiBairitu
    Label9.Height = Text11.Height
'
    Text22.Left = 360 + (9015 - 360) * HyoujiBairitu
    Text22.Top = 360 + (2760 - 360) * HyoujiBairitu
    Text22.FontSize = 10 * HyoujiBairitu
    Text22.Width = 855 * HyoujiBairitu
    Text22.Height = 285 * HyoujiBairitu
'
    Label21.Left = 360 + (7815 - 360) * HyoujiBairitu
    Label21.Top = 360 + (2760 - 360) * HyoujiBairitu
    Label21.FontSize = 10 * HyoujiBairitu
    Label21.Width = 1215 * HyoujiBairitu
    Label21.Height = Text22.Height
'
    Text12.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text12.Top = 360 + (3240 - 360) * HyoujiBairitu
    Text12.FontSize = 10 * HyoujiBairitu
    Text12.Width = 1215 * HyoujiBairitu
    Text12.Height = 285 * HyoujiBairitu
'
    Label10.Left = 360
    Label10.Top = 360 + (3240 - 360) * HyoujiBairitu
    Label10.FontSize = 10 * HyoujiBairitu
    Label10.Width = 1095 * HyoujiBairitu
    Label10.Height = Text12.Height
'
    Text13.Left = 360 + (3855 - 360) * HyoujiBairitu
    Text13.Top = 360 + (3240 - 360) * HyoujiBairitu
    Text13.FontSize = 10 * HyoujiBairitu
    Text13.Width = 1215 * HyoujiBairitu
    Text13.Height = 285 * HyoujiBairitu
'
    Label11.Left = 360 + (2760 - 360) * HyoujiBairitu
    Label11.Top = 360 + (3240 - 360) * HyoujiBairitu
    Label11.FontSize = 10 * HyoujiBairitu
    Label11.Width = 1095 * HyoujiBairitu
    Label11.Height = Text13.Height
'
    Text16.Left = 360 + (6255 - 360) * HyoujiBairitu
    Text16.Top = 360 + (3240 - 360) * HyoujiBairitu
    Text16.FontSize = 10 * HyoujiBairitu
    Text16.Width = 1215 * HyoujiBairitu
    Text16.Height = 285 * HyoujiBairitu
'
    Label14.Left = 360 + (5160 - 360) * HyoujiBairitu
    Label14.Top = 360 + (3240 - 360) * HyoujiBairitu
    Label14.FontSize = 10 * HyoujiBairitu
    Label14.Width = 1095 * HyoujiBairitu
    Label14.Height = Text16.Height
'
    Text18.Left = 360 + (1455 - 360) * HyoujiBairitu
    Text18.Top = 360 + (3720 - 360) * HyoujiBairitu
    Text18.FontSize = 10 * HyoujiBairitu
    Text18.Width = 2535 * HyoujiBairitu
    Text18.Height = 285 * HyoujiBairitu
'
    Label17.Left = 360
    Label17.Top = 360 + (3720 - 360) * HyoujiBairitu
    Label17.FontSize = 10 * HyoujiBairitu
    Label17.Width = 1095 * HyoujiBairitu
    Label17.Height = Text18.Height
'
    cboSyukko.Left = 360 + (1455 - 360) * HyoujiBairitu
    cboSyukko.Top = 360 + (3720 - 360) * HyoujiBairitu
    cboSyukko.FontSize = 10 * HyoujiBairitu
    cboSyukko.Width = 2535 * HyoujiBairitu
'   cboSyukko.Height = 285 * HyoujiBairitu
'
    Text19.Left = 360 + (5175 - 360) * HyoujiBairitu
    Text19.Top = 360 + (3720 - 360) * HyoujiBairitu
    Text19.FontSize = 10 * HyoujiBairitu
    Text19.Width = 3855 * HyoujiBairitu
    Text19.Height = 285 * HyoujiBairitu
'
    Label18.Left = 360 + (4080 - 360) * HyoujiBairitu
    Label18.Top = 360 + (3720 - 360) * HyoujiBairitu
    Label18.FontSize = 10 * HyoujiBairitu
    Label18.Width = 1095 * HyoujiBairitu
    Label18.Height = Text19.Height
'
    cboBikouran.Left = 360 + (5175 - 360) * HyoujiBairitu
    cboBikouran.Top = 360 + (3720 - 360) * HyoujiBairitu
    cboBikouran.FontSize = 10 * HyoujiBairitu
    cboBikouran.Width = 3855 * HyoujiBairitu
'   cboBikouran.Height = 285 * HyoujiBairitu
'
    Text20.Left = 360 + (3855 - 360) * HyoujiBairitu
    Text20.Top = 360 + (4200 - 360) * HyoujiBairitu
    Text20.FontSize = 10 * HyoujiBairitu
    Text20.Width = 1215 * HyoujiBairitu
    Text20.Height = 285 * HyoujiBairitu
'
    Label19.Left = 360 + (2760 - 360) * HyoujiBairitu
    Label19.Top = 360 + (4200 - 360) * HyoujiBairitu
    Label19.FontSize = 10 * HyoujiBairitu
    Label19.Width = 1095 * HyoujiBairitu
    Label19.Height = Text20.Height
'
    cboCad.Left = 360 + (3855 - 360) * HyoujiBairitu
    cboCad.Top = 360 + (4200 - 360) * HyoujiBairitu
    cboCad.FontSize = 10 * HyoujiBairitu
    cboCad.Width = 1215 * HyoujiBairitu
'   cboCad.Height = 285 * HyoujiBairitu
'
    Text24.Left = 360 + (6255 - 360) * HyoujiBairitu
    Text24.Top = 360 + (4200 - 360) * HyoujiBairitu
    Text24.FontSize = 10 * HyoujiBairitu
    Text24.Width = 1215 * HyoujiBairitu
    Text24.Height = 285 * HyoujiBairitu
'
    Label23.Left = 360 + (5160 - 360) * HyoujiBairitu
    Label23.Top = 360 + (4200 - 360) * HyoujiBairitu
    Label23.FontSize = 10 * HyoujiBairitu
    Label23.Width = 1095 * HyoujiBairitu
    Label23.Height = Text24.Height
'
    cmdSakujyo.Left = 360 + (1080 - 360) * HyoujiBairitu
    cmdSakujyo.Top = 360 + (4800 - 360) * HyoujiBairitu
    cmdSakujyo.FontSize = 9 * HyoujiBairitu
    cmdSakujyo.Width = 1335 * HyoujiBairitu
    cmdSakujyo.Height = 495 * HyoujiBairitu
'
    cmdKettei.Left = 360 + (2760 - 360) * HyoujiBairitu
    cmdKettei.Top = 360 + (4800 - 360) * HyoujiBairitu
    cmdKettei.FontSize = 9 * HyoujiBairitu
    cmdKettei.Width = 1335 * HyoujiBairitu
    cmdKettei.Height = 495 * HyoujiBairitu
'
    cmdTuika.Left = 360 + (1080 - 360) * HyoujiBairitu
    cmdTuika.Top = 360 + (5520 - 360) * HyoujiBairitu
    cmdTuika.FontSize = 9 * HyoujiBairitu
    cmdTuika.Width = 1335 * HyoujiBairitu
    cmdTuika.Height = 495 * HyoujiBairitu
'
    cmdkousin.Left = 360 + (2760 - 360) * HyoujiBairitu
    cmdkousin.Top = 360 + (5520 - 360) * HyoujiBairitu
    cmdkousin.FontSize = 9 * HyoujiBairitu
    cmdkousin.Width = 1335 * HyoujiBairitu
    cmdkousin.Height = 495 * HyoujiBairitu
'
    cmdUp.Left = 360 + (5160 - 360) * HyoujiBairitu
    cmdUp.Top = 360 + (5520 - 360) * HyoujiBairitu
    cmdUp.FontSize = 10 * HyoujiBairitu
    cmdUp.Width = 855 * HyoujiBairitu
    cmdUp.Height = 495 * HyoujiBairitu
'
    cmdDown.Left = 720 + (6360 - 720) * HyoujiBairitu
    cmdDown.Top = 360 + (5520 - 360) * HyoujiBairitu
    cmdDown.FontSize = 10 * HyoujiBairitu
    cmdDown.Width = 855 * HyoujiBairitu
    cmdDown.Height = 495 * HyoujiBairitu
'
    cmdQuit.Left = 360 + (7800 - 360) * HyoujiBairitu
    cmdQuit.Top = 360 + (5520 - 360) * HyoujiBairitu
    cmdQuit.FontSize = 10 * HyoujiBairitu
    cmdQuit.Width = 1215 * HyoujiBairitu
    cmdQuit.Height = 495 * HyoujiBairitu
End Sub
