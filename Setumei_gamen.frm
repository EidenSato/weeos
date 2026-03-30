VERSION 5.00
Begin VB.Form Setumei_gamen 
   BackColor       =   &H00004000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "EＥOS２ 説明"
   ClientHeight    =   5550
   ClientLeft      =   1845
   ClientTop       =   1545
   ClientWidth     =   9855
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
   Icon            =   "Setumei_gamen.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   5550
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtnaiyou 
      BackColor       =   &H00008000&
      ForeColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   480
      MousePointer    =   1  '矢印
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Setumei_gamen.frx":030A
      Top             =   480
      Width           =   8895
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "閉じる (&Q)"
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   4920
      Width           =   1815
   End
End
Attribute VB_Name = "Setumei_gamen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*********************
'*** EEOSⅡ の説明 ***
'*********************
'
Option Explicit
    
Private Sub cmd_OK_Click()
    Unload Me
End Sub

Private Sub GHyouji()
    Dim i As Integer
'
    EEOS_Setumei = vbCrLf _
        & Trim(str(1)) & ". GRPH ＋ ｷｰ でトピックを選択できますよ！" & vbCrLf _
        & vbCrLf _
        & Trim(str(2)) & ". 項目の移動は TAB ｷｰ ですよ！" & vbCrLf _
        & vbCrLf _
        & Trim(str(3)) & ". スタートした時に開くウインドウは「環境(O)」で指定できます。" & vbCrLf _
        & vbCrLf _
        & Trim(str(4)) & ". 実行したい作業はメニューから選んでね！" & vbCrLf _
        & vbCrLf _
        & Trim(str(5)) & ". マウス右ボタンクリックでもメニューが呼び出せます。" & vbCrLf
    i = 5
'
    Select Case FLGjob
    Case 0      '*** 無表示スタート画面 ***
        '
'
    Case 1      '*** 構成表画面 ***
        EEOS_Setumei = EEOS_Setumei _
            & vbCrLf _
            & Trim(str(i + 1)) & ". 「ROLL UP」「ROLL DOWN」キーで10行スクロールしますよ" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 2)) & ". 「↑」「↓」キーで1行スクロールしますよ。" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 3)) & ". 記入内容の変更はそのコラムを「ダブルクリック」してね！" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 4)) & ". 追加または削除は行番号を「ダブルクリック」してね！" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 5)) & ". 欄外備考欄で印刷の時改行したい所には「 _(空白+ｱﾝﾀﾞｰｽｺｱ)」を書いてね！" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 6)) & ". 行末備考欄に書ききれない時には「注１参照」などと書いて欄外備考欄に「注１：・・・」と書くとすっきりしますよ！" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 7)) & ". 部品表が２画面開いているときには「OrCAD変換」はできません。" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 8)) & ". 構成表フォルダは部品表個別フォルダと共用です。" & vbCrLf
'
    Case 2      '*** 部品表画面 ***
        EEOS_Setumei = EEOS_Setumei _
            & vbCrLf _
            & Trim(str(i + 1)) & ". 「ROLL UP」「ROLL DOWN」キーで10行スクロールしますよ" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 2)) & ". 「↑」「↓」キーで1行スクロールしますよ。" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 3)) & ". 部品表が２画面開いているときには「OrCAD変換」はできません。" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 4)) & ". 部品表個別フォルダは構成表フォルダと共用なので構成表を呼び出すことで指定できます。" & vbCrLf
'
        Select Case FLGlevel
        Case 1, 3
            EEOS_Setumei = EEOS_Setumei _
                & vbCrLf _
                & Trim(str(i + 5)) & ". 記入内容の変更はそのコラムを「ダブルクリック」してね！" & vbCrLf _
                & vbCrLf _
                & Trim(str(i + 6)) & ". 削除は行番号を「ダブルクリック」してね！" & vbCrLf _
                & vbCrLf _
                & Trim(str(i + 7)) & ". 部品表印刷＜全部＞の時、印刷する部品表は指定した「個別フォルダ」または「共有フォルダ」の中から探します。" & vbCrLf
'
        Case 4
            EEOS_Setumei = EEOS_Setumei _
                & vbCrLf _
                & Trim(str(i + 5)) & ". 記入内容の変更はそのコラムを「直接変更」してね！" & vbCrLf _
                & vbCrLf _
                & Trim(str(i + 6)) & ". 削除は「一行削除」を「クリック」してね！" & vbCrLf
        End Select
'
    Case 3, 4, 5    '*** 部品コード、メーカー、商社コード表画面 ***
        EEOS_Setumei = EEOS_Setumei _
            & vbCrLf _
            & Trim(str(i + 1)) & ". 「ROLL UP」「ROLL DOWN」キーで10行スクロールしますよ" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 2)) & ". 「↑」「↓」キーで1行スクロールしますよ。" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 3)) & ". 「コード一覧画面」では見たい行を「ダブルクリック」すると「下位画面」が開きます。" & vbCrLf _
            & vbCrLf _
            & Trim(str(i + 4)) & ". 「内容変更」「新規追加」は「ｽｰﾊﾟﾊﾞｲｻﾞｰ」でなければできません。" & vbCrLf
    End Select
End Sub

Private Sub HelpHyouji()
    Dim i As Integer
    Dim temp0 As String, temp1 As String, temp2 As String, temp3 As String
'
    Select Case FLG_Setumei
    Case 10      '***  ***
        temp0 = "                      ＥＥＯＳコード表 RoHS対応記号 規定" & vbCrLf _
        & vbCrLf _
        & "１．表記記号(<Pbﾌﾘｰ>、<RoHS>、<Green>、<Ro2>、<R863>)" & vbCrLf _
        & "  RoHS指令に準拠している部品､端子部に限って鉛フリーの部品またはRoHS指令指定元素" & vbCrLf _
        & "(鉛/水銀/ｶﾄﾞﾐｳﾑ/六価ｸﾛﾑ/PBB/PBDE)を含有していない部品には次の記号を品種概要または" & vbCrLf _
        & "規格備考欄に記入する｡" & vbCrLf _
        & "  (a)端子部限定鉛フリー                                     →<Pbﾌﾘ>" & vbCrLf _
        & "        ☆内部に鉛半田を使用/構成材料にも含有" & vbCrLf _
        & "  (b)RoHS指令に準拠                                         →<RoHS>" & vbCrLf _
        & "        ☆内部に鉛等を含有するも規定量未満/適用除外部位として使用" & vbCrLf _
        & "  (c)RoHS指令指定元素無含有                                 →<Green>" & vbCrLf _
        & "        ☆内部に鉛等を含有(自然材料に含有)するも極微量" & vbCrLf _
        & "        ☆環境負荷物質を含有(自然材料に含有)するも極微量" & vbCrLf _
        & "        ☆判定は「グリーン調達調査票、納入仕様書」にて行う" & vbCrLf _
        & "  (d)RoHS2指令に準拠                                        →<Ro2>" & vbCrLf _
        & "  (e)RoHS2指令 +(EU)2015/863に準拠                          →<R863>" & vbCrLf _
        & vbCrLf
        temp1 = "２．追加情報記号(#)" & vbCrLf _
        & "  メーカーからの新規供給品は上記対応が済んでいるが､倉庫在庫に非対応品が混ざって" & vbCrLf _
        & "いる可能性のある品種/品目には '#' 記号を前に付加する。" & vbCrLf _
        & "    ex. #、#<Pbﾌﾘｰ>、#<RoHS>、#<Green>" & vbCrLf _
        & "    ☆部品表印刷、数量表印刷時には「RoHS」記号に替えて「=>>=」を表示する。" & vbCrLf _
        & "    ★RoHS完全移行時には残非対応品は廃棄する必要がある｡" & vbCrLf _
        & vbCrLf
        temp2 = "３．端子メッキ記号(SnPb/Sn/NiPdAu/SnAgCu・・・)" & vbCrLf _
        & "  端子部に使われている金属材料のうち表面に出ている材料を表す｡" & vbCrLf _
        & "  (a)従来の共晶半田(ｽｽﾞ/鉛)メッキ       → SnPb" & vbCrLf _
        & "  (b)スズ(tin) メッキ                   → Sn" & vbCrLf _
        & "  (c)スズ/ビスマス メッキ               → SnBi" & vbCrLf _
        & "  (d)ニッケル/パラジウム/金 メッキ      → NiPdAu" & vbCrLf _
        & "  (e)金 メッキ                          → Au" & vbCrLf _
        & "  (f)ニッケル/金 メッキ                 → NiAu" & vbCrLf _
        & "  (g)銀 メッキ                          → Ag" & vbCrLf _
        & "  (h)ニッケル メッキ                    → Ni" & vbCrLf _
        & "  (i)スズ/銅 メッキ                     → SnCu" & vbCrLf _
        & "  (j)スズ/銀 メッキ                     → SnAg" & vbCrLf _
        & "  (k)スズ/銀/銅 メッキ                  → SnAgCu" & vbCrLf _
        & "    ★SnAgCu処理は従来半田で従来温度プロファイル(230℃)での実装の場合は濡れ性が" & vbCrLf _
        & "    劣るので使用できない｡" & vbCrLf _
        & vbCrLf
        temp3 = "４．半田耐熱温度(230/250/260)" & vbCrLf _
        & "  リフロー半田する時のピーク温度で、JEITAの規定により「250℃、260℃/10秒」の" & vbCrLf _
        & "２つのランクがある。従来品は「230℃/10秒」が標準的。個別温度プロファイルを一度は" & vbCrLf _
        & "確認することを推奨する｡" & vbCrLf _
        & "  フロー半田(主に挿入実装品)は「260℃」が標準になっている。" & vbCrLf _
        & "    ★パッケージの関係で最高温度がさらに低い部品があるので注意を要す｡" & vbCrLf
'
    Case 1      '***  ***
'
    End Select
'
    EEOS_Help = temp0 + temp1 + temp2 + temp3
End Sub

Private Sub Form_Load()     '*** 画面設定 ***
    Width = 480 + (9945 - 960) * HyoujiBairitu + 480
    Height = 480 + (5925 - 720) * HyoujiBairitu + 240
'
    Left = Eeos2_mainMDI.Left + (Eeos2_mainMDI.ScaleWidth - Width) \ 2
    Top = Eeos2_mainMDI.Top + 480 + (Eeos2_mainMDI.ScaleHeight - Height) \ 2
'
    txtnaiyou.Left = 480
    txtnaiyou.Top = 480
    txtnaiyou.FontSize = 10 * HyoujiBairitu
    txtnaiyou.Width = 8895 * HyoujiBairitu
    txtnaiyou.Height = 4335 * HyoujiBairitu
'
    cmd_ok.Left = 480 + (7440 - 480) * HyoujiBairitu
    cmd_ok.Top = 480 + (4920 - 480) * HyoujiBairitu
    cmd_ok.FontSize = 10 * HyoujiBairitu
    cmd_ok.Width = 1815 * HyoujiBairitu
    cmd_ok.Height = 495 * HyoujiBairitu
'
    Select Case FLG_Setumei
    Case 0      '*** 0: 操作説明 ***
        Setumei_gamen.Caption = "ＥＥＯＳ２ 操作説明"
        Call GHyouji
        txtnaiyou.Text = EEOS_Setumei
'
    Case 1      '*** 1: 改版履歴 ***
        Setumei_gamen.Caption = "ＥＥＯＳ２ 改版履歴"
        txtnaiyou.Text = EEOS_Kaihan
'
    Case Else     '*** 10～: Help<RoHS/Pbﾌﾘｰ 説明> ***
        Setumei_gamen.Caption = "ＥＥＯＳ２ Help<RoHS/Pbﾌﾘｰ記号 説明>"
        Call HelpHyouji
        txtnaiyou.Text = EEOS_Help
'
    End Select
End Sub
