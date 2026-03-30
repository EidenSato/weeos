VERSION 5.00
Begin VB.Form Eeos2_start 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'なし
   Caption         =   "Weeos"
   ClientHeight    =   4710
   ClientLeft      =   375
   ClientTop       =   1140
   ClientWidth     =   8175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "Eeos2_start.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4710
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   3720
   End
   Begin VB.Label lblSetumei 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H0000C000&
      Caption         =   "Eiden Engineering Office System 2  Ver0.9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   3960
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      Height          =   4695
      Index           =   8
      Left            =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   25
      Height          =   135
      Index           =   7
      Left            =   1200
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   25
      Height          =   135
      Index           =   6
      Left            =   1080
      Top             =   600
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   30
      Height          =   2655
      Index           =   5
      Left            =   7320
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   30
      Height          =   2655
      Index           =   4
      Left            =   720
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   20
      Height          =   2535
      Index           =   3
      Left            =   6720
      Shape           =   2  '楕円
      Top             =   960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   20
      Height          =   615
      Index           =   2
      Left            =   1200
      Shape           =   2  '楕円
      Top             =   3360
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   20
      Height          =   615
      Index           =   1
      Left            =   1200
      Shape           =   2  '楕円
      Top             =   600
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   20
      Height          =   2535
      Index           =   0
      Left            =   720
      Shape           =   2  '楕円
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLOGO 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H0000C000&
      Caption         =   "EEOS2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
   End
End
Attribute VB_Name = "Eeos2_start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*********************************
'*** ＥＥＯＳ２ スタートアップ ***
'*********************************
'
Option Explicit
'

Private Sub Form_Load()
    Width = 8175
    Height = 4710
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
'
    Timer1.Interval = 2000
    Timer1.Enabled = True
'
    DIRcont0 = App.Path   '*** カレントディレクトリーの記憶 ***
'
    Call EEOS_set   '*** このプログラム中で使用する文字列の設定 ***
'
    FLGjob = 0          '*** 作業フラグ初期化 ***
End Sub

Public Sub EEOS_set()
'                       *** このプログラム中で使用する文字列の設定 ***
    Dim Dtemp1 As String, Dtemp2 As String, Dtemp3 As String, Dtemp4 As String, Dtemp5 As String
    Dim Dtemp6 As String, Dtemp7 As String, Dtemp8 As String, Dtemp9 As String
    Dim VersionNumber As String
'
    VersionNumber = "Ver" & Trim(App.Major) & "." & Trim(App.Minor) & Trim(App.Revision)
    lblSetumei.Caption = "Eiden Engineering Office System 2 <" & VersionNumber & ">"
    EEOS_STATUS = "ＥＥＯＳ２(Eiden Engineering Office System 2)"
    EEOS_Version = "EEOS2 " & VersionNumber & " 2021/09/02"
    EEOS_Kaihan = vbCrLf _
        & "        ********** 改版履歴 **********" & vbCrLf _
        & " 2021/09/02 Ver2.30" & vbCrLf _
        & "    「概要」「規格」欄の「<>」区分に「R863」を追加。：" & vbCrLf _
        & "    各画面の表示を対応。" & vbCrLf _
        & " 2020/03/02 Ver2.20" & vbCrLf _
        & "    部品表一覧：子コードの「Ro2」表示を優先に変更。" & vbCrLf _
        & "    部品表印刷：子コードの「Ro2」表示を優先に変更。" & vbCrLf _
        & "    部品数量表印刷：子コードの「Ro2」表示を優先に変更。：" & vbCrLf _
        & " 2019/09/02 Ver2.10" & vbCrLf _
        & "    「概要」「規格」欄の「<>」区分に「Ro2」を追加。：" & vbCrLf _
        & "    部品表一覧：「Ro2」を追加。" & vbCrLf _
        & "    部品表印刷：「Ro2」を追加。" & vbCrLf _
        & "    部品数量表印刷「Ro2」を追加。：" & vbCrLf _
        & " 2016/12/19 Ver2.04" & vbCrLf _
        & "    MSL/吸湿ﾚﾍﾞﾙ/ﾊﾟﾗﾒｰﾀ追加(data6を目的変更)。：" & vbCrLf _
        & "    部品コード表：data6を「MSL」データに変更、表示ﾚｲｱｳﾄ一部変更。" & vbCrLf _
        & "    部品コード表ﾌｧｲﾙ出力：data6「MSL」データを追加出力。" & vbCrLf _
        & "    部品表一覧：耐熱性欄に「MSL」データを併記追加、表示幅を一部変更。" & vbCrLf _
        & "    部品表印刷：ﾒｯｷ/RoHS欄に「MSL」データ、耐熱温度を併記追加、現品表示欄/個数/総数の表示幅を変更。" & vbCrLf _
        & "    部品数量表印刷：耐熱性,端子ﾒｯｷ/RoHS欄に「MSL」データを併記追加。" & vbCrLf _
        & "    部品数量表印刷：ﾕｰｻﾞｰ：管理部の時には「備考欄」選択でも「MSL」データを行末に追加。" & vbCrLf _
        & "    オプション設定：構成表・部品表読み込み時のフォルダー設定選択を追加。"
    Dtemp1 = vbCrLf _
        & " 2011/01/18 Ver1.91" & vbCrLf _
        & "    構成表ﾌｧｲﾙ出力：資材用特殊形式出力を追加。" & vbCrLf _
        & " 2010/04/20 Ver1.81" & vbCrLf _
        & "    ｵﾌﾟｼｮﾝ設定：文字ｻｲｽﾞをもう２ﾎﾟｲﾝﾄ大きいｻｲｽﾞの画面を追加。" & vbCrLf _
        & "    部品コード表：たまに親コードが間違って印刷されるのを修正した積もり。" & vbCrLf _
        & " 2010/04/19 Ver1.80" & vbCrLf _
        & "    部品表：<Green>が表示されなかったので修正、表示部が狭いので<Gree>表示とする。" & vbCrLf _
        & "    部品表印刷：<Green>が表示されなかったので修正、表示部が狭いので<Gree>表示とする。" & vbCrLf _
        & "        現品表示の文字が長い時部品番号の上に重なってしまうので修正。" & vbCrLf _
        & "    部品コード表：納入仕様書が登録されている部品は「品目一覧画面」の着目部品表示の右側に" & vbCrLf _
        & "        「PDFｺﾏﾝﾄﾞﾎﾞﾀﾝ」が表示される。そして、そこをｸﾘｯｸすると対応PDFﾌｧｲﾙが開く様に修正。" & vbCrLf _
        & " 2008/10/10 Ver1.72" & vbCrLf _
        & "    部品表印刷：未登録部品の印刷位置が未修正だったので修正。" & vbCrLf _
        & "        ２行目の部品番号が１行目と離れすぎていて見にくいので右寄せにする。" & vbCrLf _
        & " 2008/09/18 Ver1.71" & vbCrLf _
        & "    部品コード表：内容修正し易いように枠を移動。" & vbCrLf _
        & "    部品表印刷：「現品表示」欄が狭すぎたので少し拡大。"
    Dtemp2 = vbCrLf _
        & " 2008/09/16 Ver1.70" & vbCrLf _
        & "    部品コード表：data6を「現品表示」データに変更。" & vbCrLf _
        & "    部品表印刷：「特記事項」欄の「[]」枠印刷を廃止し、自由位置に「[]」を表示できる様に修正。" & vbCrLf _
        & "      「現品表示」欄追加、「端子メッキ、RoHS表示」欄追加、このためオプション選択項目変更。" & vbCrLf _
        & "    部品数量表印刷：同一コードでも「特記事項」「備考」欄が違ったら行を分ける様に修正。" & vbCrLf _
        & " 2008/03/31 Ver1.64" & vbCrLf _
        & "    構成表：「環境設定」で個別・共有フォルダを変更した後「部品表ﾁｪｯｸ」を実行すると変更以前の" & vbCrLf _
        & "      フォルダを探しに行ってしまうバグを修正。" & vbCrLf _
        & " 2008/01/04 Ver1.63" & vbCrLf _
        & "    部品表：「足ﾋﾟﾝ数合計表示」一つの部品が「1000ﾋﾟﾝ以上」の時正しくカウントしないバグを修正。" & vbCrLf _
        & "    部品表印刷：一つの部品表を印刷した後、別の部品表を選択し印刷したつもりなのに、また同じ" & vbCrLf _
        & "      部品表が印刷されるバグを修正。" & vbCrLf _
        & "        備考欄に記入事項が有ったとき無用な空白行が追加される事があるバグを修正。" & vbCrLf _
        & "    構成表：画面を開いてすぐに「部品表チェック」を実行すると有るはずの部品表が「見つかりませ" & vbCrLf _
        & "      ん」とエラー表示するバグを修正。" & vbCrLf _
        & " 2007/10/26 Ver1.62" & vbCrLf _
        & "    部品表印刷：「数量表ファイル出力」の時「備考欄印刷：耐熱/ﾒｯｷ」の出力機能が正しく動作して" & vbCrLf _
        & "      いなかったのを修正。"
    Dtemp3 = vbCrLf _
        & " 2007/10/15 Ver1.61" & vbCrLf _
        & "    部品表印刷：「構成表による印刷」の時ディスクを入れ替えても新しい構成表の工番にならない" & vbCrLf _
        & "      バグを修正。備考欄の文字数が多い時の行分割印刷機能が正しく機能していなかったのを修正。" & vbCrLf _
        & " 2007/03/06 Ver1.60" & vbCrLf _
        & "    構成表：「チェック」ボタンを設け記載されている電気部品表の有無をチェックする機能を追加。" & vbCrLf _
        & "    部品数量表印刷：部品表と同じように備考欄への「耐熱温度と端子メッキ」オプションスイッチ" & vbCrLf _
        & "      を追加。" & vbCrLf _
        & "    部品コード表：「項目一覧」「品目一覧」画面に「？」ボタンを設けRoHS対応説明表示を追加。" & vbCrLf _
        & " 2007/01/24 Ver1.51" & vbCrLf _
        & "    主画面：なにもウインドウを開いていないときにも「部品表２」を開けるように変更。" & vbCrLf _
        & "    構成表/部品表：毎回フォルダー選択になっている時、指定ボックスが初めの１回しか開かない" & vbCrLf _
        & "      バグを修正。" & vbCrLf _
        & "    部品表更新：初めに出るフォルダ選択画面をキャンセルすると部品表が開けなくなるバグを修正。" & vbCrLf _
        & "    部品表更新/部品表印刷：オプション選択 端子メッキがSnPb時 /---- と表示するように変更。" & vbCrLf _
        & "    部品コード表：「品目一覧」の平均単価欄を10万円台も表示できるように少し広くする。" & vbCrLf _
        & "      「品種一覧」「項目一覧」画面の着目項目欄を重ねても見えるように左側に寄せる。"
    Dtemp4 = vbCrLf _
        & " 2007/01/15 Ver1.50" & vbCrLf _
        & "    部品表更新：未登録部品を修正する時、「インデックスエラー」が出るバグを修正。" & vbCrLf _
        & "        フォルダーを毎回指定にしている時、個別フォルダー指定をキャンセルしたら共通フォルダー" & vbCrLf _
        & "      指定もスキップするように変更。" & vbCrLf _
        & "        部品が「RoHS対応」または「鉛ﾌﾘｰ」の時には状況を画面表示するように変更。" & vbCrLf _
        & "    部品表印刷：部品表更新画面を開き次に印刷を実行した時、ファイル名を更新画面で開いた名前" & vbCrLf _
        & "      で印刷することがあるバグを修正。" & vbCrLf _
        & "        備考欄への「耐熱温度と端子メッキ」オプションスイッチをONにした時、部品が「RoHS対応」" & vbCrLf _
        & "      または「鉛ﾌﾘｰ」の時には状況を印刷するように変更。" & vbCrLf _
        & "    部品表OrCADからの変換：備考欄入力が数回目から要求されないバグを修正。" & vbCrLf _
        & "        作成データに備考欄記述があった場合選択画面に表示するように変更。" & vbCrLf _
        & "    OrCAD変換作業ファイル編集：OrCAD変換時に使用する変換作業データの編集画面を追加した。" & vbCrLf _
        & "    部品コード表：「部品検索」時「品種一覧」画面が最終ページになる品種の時、正しい「品目一覧」" & vbCrLf _
        & "      画面が開かないバグを修正。"
    Dtemp5 = vbCrLf _
        & " 2006/11/22 Ver1.43" & vbCrLf _
        & "    部品表更新：修正後、「終了」「廃棄終了中止→Yes」を選択したときに空の部品表が出来るバグ" & vbCrLf _
        & "      を修正。" & vbCrLf _
        & "    部品コード表：画面サイズを変更した時に「部品検索」ボタンが正しい位置に動かないバグを修正。" & vbCrLf _
        & " 2006/10/24 Ver1.42" & vbCrLf _
        & "    部品コード表：「部品検索」画面の「△/▽」ボタン表示を他画面と統一するため「UP/DOWN」に" & vbCrLf _
        & "      変更。  合わせて、「↑/↓」キー、「PageUp/PageDown」キーにも対応。" & vbCrLf _
        & " 2006/10/18 Ver1.41" & vbCrLf _
        & "    部品表更新：名前を付けて保存をした後、上書き保存するとその前のファイルに保存してしまう" & vbCrLf _
        & "      バグを修正。" & vbCrLf _
        & " 2006/10/17 Ver1.40" & vbCrLf _
        & "    部品表更新：２画面開けるように改修、メモリーエリアの扱いを個別画面に分割。" & vbCrLf _
        & "        ウインドウサイズを元の大きさに戻すコマンドを実行した時、正しく画面表示されないバグ" & vbCrLf _
        & "      を修正。" & vbCrLf _
        & "       「部品単価/形状」の欄をオプションスイッチで「半田耐熱/端子メッキ」表示できるように" & vbCrLf _
        & "      修正。" & vbCrLf _
        & "    部品コード表：品種一覧画面に「部品検索」ボタンを追加し、品名(部分)からのコード番号検索" & vbCrLf _
        & "      画面を追加。"
    Dtemp6 = vbCrLf _
        & " 2006/10/06 Ver1.32" & vbCrLf _
        & "    部品表更新：品種変更時の選択候補表示にコード順/品名順表示オプションを追加。" & vbCrLf _
        & "        部品番号AnXmを追加した時に正しい位置に挿入されないバグを修正。" & vbCrLf _
        & "    部品表印刷：Ver1.30のフォルダ指定の扱いが間違っていたので修正。" & vbCrLf _
        & "        文字数の多い部品名称が部品番号の欄に食い込むので文字サイズを少し小さくなるように修正。" & vbCrLf _
        & "    印刷時プリンタ選択画面： プリンタ一覧にコメントを追加。" & vbCrLf _
        & " 2006/07/07 Ver1.31" & vbCrLf _
        & "    構成表/部品表/部品一覧表/部品数量表印刷：印刷時プリンタ選択画面を追加。" & vbCrLf _
        & " 2006/07/04 Ver1.30" & vbCrLf _
        & "    環境設定：部品表のフォルダを個別フォルダーと共有フォルダーが指定できるようにパラメータを" & vbCrLf _
        & "      追加。" & vbCrLf _
        & "    構成表/部品表/部品一覧表/部品数量表印刷：印刷時のプリンタ設定画面を表示しないように" & vbCrLf _
        & "      変更。" & vbCrLf _
        & "    部品コード表：未使用の代品コード欄を半田付け時の最高耐熱温度欄に、出庫数欄を端子メッキ" & vbCrLf _
        & "      欄に変更。" & vbCrLf _
        & "    部品表印刷：備考欄に最高耐熱温度と端子メッキを印刷するように選択肢を追加。"
    Dtemp7 = vbCrLf _
        & " 2006/06/23 Ver1.23" & vbCrLf _
        & "    部品一覧表印刷：いきなり一覧表印刷/ファイル出力を行うとファイルが生成されないバグを修正。" & vbCrLf _
        & "    部品一覧表印刷：ファイル出力(CSV)の形式が間違っていた("","" になっていなかった)ので修正。" & vbCrLf _
        & " 2006/04/26 Ver1.22" & vbCrLf _
        & "    部品表印刷：左上のアイコンが旧アイコンだったのを修正。" & vbCrLf _
        & "    部品表/部品表印刷他：ファイル名など「小文字=>大文字変換」を中止する。" & vbCrLf _
        & " 2004/05/12 Ver1.21" & vbCrLf _
        & "    構成表印刷：行数が25行の時に改ﾍﾟｰｼﾞしてしまうバグを修正。" & vbCrLf _
        & "    環境設定：環境設定をするとオプション設定が無効になるバグを修正。" & vbCrLf _
        & "    部品コード表：データ修正のあと背景色が黒になるバグを修正。" & vbCrLf _
        & "    構成表/部品表：廃棄終了時の警告文(Yes/No)を同じ表現に変更。" & vbCrLf _
        & "    部品コード表：ファイル出力フォーマットに間違いがあったので修正。"
    Dtemp8 = vbCrLf _
        & " 2004/04/08 Ver1.20" & vbCrLf _
        & "    全般：各ファイルが環境設定通りのフォルダに無い時、異常終了しない様に変更。" & vbCrLf _
        & "    オプション設定：表示文字の大きさを10ﾎﾟｲﾝﾄと12ﾎﾟｲﾝﾄを選べる選択項目を追加。" & vbCrLf _
        & "    部品表：OrCADからの変換 図面タイトルが無くても動作できるように変更。" & vbCrLf _
        & "    部品表：     〃     絞り込み項目に「標準部品を含む」を追加。" & vbCrLf _
        & "    部品表：部品番号のモジュール番号正式対応。（「A1P1」表記を許容）" & vbCrLf _
        & "    部品表：ファイル名に８文字以上を許容し、拡張子を「.PLT」と規定した。" & vbCrLf _
        & "    部品表：変更/追加時にも選択欄内に部品指定(標準部品ﾅﾄﾞ)の表示を追加した。" & vbCrLf _
        & "    部品表：追加時に追加した部品を一覧表の上から２行目から中央10行目に変更した。" & vbCrLf _
        & "    メーカーコード表：コード番号のアルファベット表記対応。" & vbCrLf _
        & "    商社コード表：４桁コード番号表示対応。" & vbCrLf _
        & "    部品コード表：品種一覧画面で標準部品が含まれる品種のコード番号を強調。" & vbCrLf _
        & "    部品コード表：品種一覧画面に「図面表記」欄を追加して確認し易くした。"
    Dtemp9 = vbCrLf _
        & " 2002/06/04 Ver1.10" & vbCrLf _
        & "    部品表印刷：備考欄に代えて平均単価/総計を記入するオプションを追加。" & vbCrLf _
        & "    数量表印刷：備考欄に代えて平均単価/総計を記入するオプションを追加。" & vbCrLf _
        & "    数量表ﾌｧｲﾙ出力：備考欄に代えて平均単価を記入するオプションを追加。" & vbCrLf _
        & " 2002/05/09 Ver1.03" & vbCrLf _
        & "    数量表印刷：同じコード番号でも部品記号が異なっていると行が分かれるのを修正。" & vbCrLf _
        & " 2002/01/28 Ver1.02" & vbCrLf _
        & "    部品表：部品表フォルダ指定が曖昧だったのを改善。" & vbCrLf _
        & " 2002/01/15 Ver1.01" & vbCrLf _
        & "    部品コード表：画面でダブルクリックすると帳票画面だったのを下位画面の呼び出しに変更。" & vbCrLf _
        & " 2002/01/11 Ver1.00" & vbCrLf _
        & "    ＥＥＯＳのマルチウインドウ化版リリース→＜ＥＥＯＳ２＞"
'
    EEOS_Kaihan = EEOS_Kaihan & Dtemp1 & Dtemp2 & Dtemp3 & Dtemp4 & Dtemp5 & Dtemp6 & Dtemp7 & Dtemp8 & Dtemp9
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
'
    Eeos2_mainMDI.Show
    Unload Me
End Sub


