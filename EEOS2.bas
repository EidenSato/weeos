Attribute VB_Name = "EEOS_common"
'********************************
'***   ＥＥＯＳ２ 共通宣言    ***
'***                          ***
'*** 2000.11.24 by S.Fukazawa ***
'********************************
'
Option Compare Binary
'
'               *** 環境設定 ***
Public DIRcont0 As String       '*** アプリケーションルートディレクトリ ***
Public EEOS_STATUS As String    '*** ステータス情報 ***
Public EEOS_Version As String   '*** バージョン情報 ***
Public EEOS_Kaihan As String    '*** 改版履歴 ***
Public EEOS_Setumei As String   '*** 操作説明 ***
Public EEOS_Help As String      '*** Help説明 ***
Public FLG_Setumei As Integer   '*** 説明内容分類フラグ ***
                                '*** 0: 操作説明、 1: 改版履歴、 2: ﾊﾞｰｼﾞｮﾝ情報 ***
Public FLGjob As Integer        '*** 作業フラグ ***
                                '*** 1: 構成表、2: 部品表、3: 部品ｺｰﾄﾞ表、4: ﾒｰｶｰｺｰﾄﾞ表、
                                '*** 5: 商社ｺｰﾄﾞ表 ***
Public Xnum As Integer          '*** XCONT0()の最大配列数 ***
Public Xcont0() As String       '*** 環境コントロールデータ配列 ***
'
Public Const cXnum0 = 17        '*** 最大配列数基準値 ***
'
Public HyoujiBairitu As Double  '*** 画面表示の文字倍率 ***
Public Const cHBairitu = 1.24
Public Const cUHBairitu = 1.44
'
Public flag_cancel As Boolean
Public strMyPrinter As String
'
'               *** TEMPエリア ***
Public TMPdir1 As String        '*** 構成表(個別)DIR 一時メモリー ***
Public TMPdir2 As String        '*** 部品表(共有)DIR 一時メモリー ***
Public TMPdir3 As String        '*** CAD_PLST DIR 一時メモリー    ***
Public TMPplst As String        '*** 部品表(実行)DIR 一時メモリー ***
Public Tmpleft As Integer       '*** left位置 一時メモリー ***
Public Tmptop As Integer        '***  Top位置 一時メモリー ***
Public TMPiti As Integer        '*** スロット位置 一時メモリー ***
'
Public TMPcol As Integer        '*** 行 位置 一時メモリー ***
Public TMProw As Integer        '*** 列 位置 一時メモリー ***
'
'               *** 商社コード ***
Public FLGtrader As Integer     '*** 商社画面存在フラグ ***
Public DRVtrader As String      '*** 商社コードディレクトリ ***
Public Trdnum0 As Integer       '*** 商社コード配列数 ***
Public Trddim0 As Integer       '*** 商社コード次元数 ***
Public Trader() As String       '*** 商社コードデータ配列 ***
Public Trdp As Integer          '*** 商社コード着目点 ***
Public Trdps As Integer         '*** 商社コード仮着目点 ***
Public FLGtrd_data_change As Integer    '*** 商社コード内容変更フラグ ***
'
'               *** メーカーコード ***
Public FLGmaker As Integer      '*** メーカー画面存在フラグ ***
Public DRVmaker As String       '*** メーカーコードディレクトリ ***
Public Maknum0 As Integer       '*** メーカーコード配列数 ***
Public Makdim0 As Integer       '*** メーカーコード次元数 ***
Public Maker() As String        '*** メーカーコードデータ配列 ***
Public Makp As Integer          '*** メーカーコード着目点 ***
Public Makps As Integer         '*** メーカーコード仮着目点 ***
Public FLGmak_data_change As Integer    '*** メーカーコード内容変更フラグ ***
'
'               *** 一覧用 部品コード項目 ***
Public FLGitem As Integer       '*** 部品コード項目画面存在フラグ ***
Public DRVitem0 As String       '*** 部品コード項目ディレクトリ ***
Public Anum0 As Integer         '*** 部品コード配列数 ***
Public Adim0 As Integer         '*** 部品コード次元数 ***
Public Aitem0() As String       '*** 部品コードデータ配列 ***
Public ip0 As Integer           '*** 部品コード着目点 ***
Public ips0 As Integer          '*** 部品コード仮着目点 ***
Public icps0 As Integer         '*** 現在メモリー上にある項目を示す ***
Public FLGitem_data_change As Integer   '*** 部品コード項目内容変更フラグ ***
'
Public Const cAdim0 = 6         '*** 部品コード次元数基準値 ***
'
'               *** 一時使用 部品コード項目 ***
Public ipT As Integer           '*** 部品コード着目点 ***
Public ipsT As Integer          '*** 部品コード仮着目点 ***
Public icpsT As Integer         '*** 現在メモリー上にある項目を示す ***
'
'               *** 一覧用 品種コード ***
Public FLGindex As Integer      '*** 部品コード品種画面存在フラグ ***
Public DRVindex0 As String      '*** 部品コード品種ディレクトリ ***
Public Bnum0 As Integer         '*** 品種コード配列数 ***
Public Bdim0 As Integer         '*** 品種コード次元数 ***
Public Bindex0() As String      '*** 品種コードデータ配列 ***
Public jp0 As Integer           '*** 品種コード着目点 ***
Public jps0 As Integer          '*** 品種コード仮着目点 ***
Public jcps0 As Integer         '*** 現在メモリー上にある品種を示す ***
Public FLGindex_data_change As Integer   '*** 部品コード品種内容変更フラグ ***
'
Public Const cBdim0 = 15        '*** 品種コード次元数基準値 ***
'
'               *** 一時使用 品種コード ***
Public DRVindexT As String      '*** 部品コード品種ディレクトリ ***
Public BnumT As Integer         '*** 品種コード配列数 ***
Public BdimT As Integer         '*** 品種コード次元数 ***
Public BindexT() As String      '*** 品種コードデータ配列 ***
Public jpT As Integer           '*** 品種コード着目点 ***
Public jpsT As Integer          '*** 品種コード仮着目点 ***
Public jcpsT As Integer         '*** 現在メモリー上にある品種を示す ***
'
'
'               *** 一覧用 MAINコード ***
Public FLGmain As Integer       '*** 部品コード品目画面存在フラグ ***
Public DRVmain0 As String       '*** 部品コード品目ディレクトリ ***
Public Cnum0 As Integer         '*** 品目コード配列数 ***
Public Cdim0 As Integer         '*** 品目コード次元数 ***
Public Cmain0() As String       '*** 品目コードデータ配列 ***
Public kp0 As Integer           '*** 品目コード着目点 ***
Public kps0 As Integer          '*** 品目コード仮着目点 ***
Public kcps0 As Integer         '*** 現在メモリー上にある項目を示す ***
Public FLGmain_data_change As Integer   '*** 部品コード品種内容変更フラグ ***
'
Public Const cCdim0 = 21        '*** 品目コード次元数基準値 ***
'
'               *** 一時使用 MAINコード ***
Public DRVmainT As String       '*** 部品コード品目ディレクトリ ***
Public CnumT As Integer         '*** 品目コード配列数 ***
Public CdimT As Integer         '*** 品目コード次元数 ***
Public CmainT() As String       '*** 品目コードデータ配列 ***
Public kpT As Integer           '*** 品目コード着目点 ***
Public kpsT As Integer          '*** 品目コード仮着目点 ***
Public kcpsT As Integer         '*** 現在メモリー上にある項目を示す ***
'
'               *** 一覧用 構成表関連 ***
Public FLGconst As Integer      '*** 構成表画面存在フラグ ***
Public CURR_file As String      '*** 表示中のファイル名 ***
'Public DRVconst As String       '*** 構成表ディレクトリ ***
'Public CATno As String          '*** 型名 ***
'Public CATname As String        '*** 品名 ***
'Public Zuban As String          '*** 図番 ***
'Public Person As String         '*** 担当者 ***
'Public Orgdate As String        '*** 作成日 ***
'Public Revdate As String        '*** 修正日 ***
'Public Checkdate As String      '*** 数量表計算日 ***
'Public Outdate As String        '*** ラベル印刷日 ***
'Public Ktotal As Integer        '*** 構成表配列数 ***
'Public Kdim As Integer          '*** 構成表次元数 ***
'Public KLST() As String         '*** 構成表データ配列 ***
'Public Kouban As String         '*** 工番 ***
'Public Daisuu As String         '*** 台数 ***
'Public Kbikou As String         '*** 備考欄 ***
'Public KyobiA As String         '*** 予備１ ***
'Public KyobiB As String         '*** 予備２ ***
'
Public Const Ckdim = 5          '*** 構成表次元数基準値 ***
'
'               *** 一時使用 構成表関連 ***
Public DRVconstT As String      '*** 構成表ディレクトリ ***
Public CATnoT As String         '*** 型名 ***
Public CATnameT As String       '*** 品名 ***
Public ZubanT As String         '*** 図番 ***
Public PersonT As String        '*** 担当者 ***
Public OrgdateT As String       '*** 作成日 ***
Public RevdateT As String       '*** 修正日 ***
Public CheckdateT As String     '*** 数量表計算日 ***
Public OutdateT As String       '*** ラベル印刷日 ***
Public KtotalT As Integer       '*** 構成表配列数 ***
Public KdimT As Integer         '*** 構成表次元数 ***
Public KLSTT() As String        '*** 構成表データ配列 ***
Public KoubanT As String        '*** 工番 ***
Public DaisuuT As String        '*** 台数 ***
Public KbikouT As String        '*** 備考欄 ***
Public KyobiAT As String        '*** 予備１ ***
Public KyobiBT As String        '*** 予備２ ***
Public BaisuuT As String        '*** 倍数 ***
'
'               *** 一覧用 部品表関連 ***
Public FLGplst As Integer       '*** 部品表画面存在フラグ ***
Public FLGlevel As Integer      '*** 部品表作業フラグ ***
'
Public DRVcadplst As String     '*** 変換部品表ディレクトリ ***
Public mp As Integer            '*** 部品表着目点 ***
Public mps As Integer           '*** 部品表仮着目点 ***
'
'Public DRVpartlist As String    '*** 部品表ディレクトリ ***
'Public PFLname As String        '*** 表示ファイル名 ***
'Public Plistname As String      '*** 機種名 ***
'Public Plistdate As String      '*** 作成日 ***
'Public Remarks As String        '*** 備考欄 ***
'Public Ptotal As Integer        '*** 部品表配列数 ***
'Public Pdim0 As Integer         '*** 部品表次元数 ***
'Public PLST() As String         '*** 部品表データ配列 ***
'
Public Const cPdim0 = 4        '*** 部品表次元数基準値 ***
'
'               *** 一覧用 部品表２関連 ***
Public FLGplst2 As Integer      '*** 部品表画面存在フラグ ***
'
'               *** 一時使用 部品表関連 ***
Public DRVpartlistT As String   '*** 部品表ディレクトリ ***
Public PFLnameT As String       '*** 表示ファイル名 ***
Public PlistnameT As String     '*** 機種名 ***
Public PlistdateT As String     '*** 作成日 ***
Public RemarksT As String       '*** 備考欄 ***
Public PtotalT As Integer       '*** 部品表配列数 ***
Public PdimT As Integer         '*** 部品表次元数 ***
Public PLSTT() As String        '*** 部品表データ配列 ***
'
'               *** 部品表関連フラグ ***
Public pp As Integer
Public pps As Integer
Public MpointC As Integer
Public MpointR As Integer
Public Gdata1(1) As String
Public Gdata2(1) As String
Public Gdata3(1) As String
Public Gdata4(1) As String
Public Gdata5(1) As String
Public Gdata6(1) As String
Public Gdata7(1) As String
Public MEM_ips As Integer
Public MEM_jps As Integer
Public MEM_kps As Integer
Public FLGsubete As Integer
'
'               *** 部品表作成データ ***
Public FLGplstWork As Integer   '*** 部品表作成データ編集画面存在フラグ ***
Public DRVplstWork As String
Public PlstWork() As String
Public FLGoption As Integer
Public SearchKey As String
'
Public Const cPLSTWORKmax = 300
Public Const cPLSTWORKdim = 4
'
'               *** 一般フラグ ***
Public FLGshinki As Integer     '*** 新規作成フラグ ***
Public FLGchange As Integer
Public FLGtuika As Integer
Public FLGsakujo As Integer
Public STATUS As String
Public STATUS2 As String
Public FLGesc As Integer
Public FLGall As Integer    '*** 0:個別  1:すべて
Public FLGfile As Integer   '*** 0:印刷  1:ファイル出力  2:未定
Public FLGowari As Integer  '*** 0:無効  1:処理の終わり  9:プログラムの終了
Public FLG_job_error_end As Integer  '*** 0:処理の正常終了  1:エラーによる未実行終了

Public Sub GET_koumoku(rdata As String, aitem() As String, anum As Integer)
'               *** 項目記号の取得 ***
    Dim i As Integer, j As Integer
'
    For i = 1 To anum
        For j = 3 To 5
            If rdata = aitem(i, j) Then
                rdata = aitem(i, 0)
                Exit Sub
'
            End If
        Next j
    Next i
'
    rdata = "**"
End Sub

Public Sub GET_ips(Tdata As String, aitem() As String, anum As Integer, adim As Integer, ips As Integer, icps As Integer, _
                    drvindex As String, bindex() As String, bnum As Integer, bdim As Integer, jcps As Integer, kcps As Integer)
    Dim i As Integer
'
    i = 1
    Do While i <= anum
        If Trim(Tdata) = aitem(i, 0) Then Exit Do
        i = i + 1
    Loop
'
    ips = i
    If icps <> ips Then
        Call SET_DRVindex(drvindex, aitem(), ips)
        Call RDindex(drvindex, bindex(), bnum, bdim)
        icps = ips
        jcps = 0
        kcps = 0
    End If
End Sub

Public Sub GET_jps(Tdata As String, aitem() As String, ips As Integer, bindex() As String, jps As Integer, jcps As Integer, _
                    drvmain As String, cmain() As String, cnum As Integer, cdim As Integer, kcps As Integer)
    Dim j As Integer
'
    j = 1
    Do While jp <= bnum
        If Tdata = bindex(j, 0) Then Exit Do
        j = j + 1
    Loop
'
    jps = j
    If jcps <> jps Then
        Call SET_DRVmain(drvmain, aitem(), ips, bindex(), jps)
        Call RDmain(drvmain, cmain(), cnum, cdim)
        jcps = jps
        kcps = 0
    End If
End Sub

Public Sub GET_kps(Tdata As String, cmain() As String, cnum As Integer, kps As Integer, kcps As Integer)
    Dim k As Integer
'
    k = 1
    Do While k <= cnum
        If Tdata = cmain(k, 0) Then Exit Do
        k = k + 1
    Loop
'
    kps = k
    kcps = kps
End Sub

Public Sub SET_DRVmain(Drvtemp As String, aitem() As String, ipp As Integer, bindex() As String, jpp As Integer)
'                   *** ﾒｲﾝｺｰﾄﾞﾌｧｲﾙ名作成 ***
    If aitem(ipp, 0) = "IC" Then
        Drvtemp = Xcont0(2) & "\IC\IC" & Left(bindex(jpp, 0), 1) & "\L" & bindex(jpp, 0) & ".COD"
    Else
        Drvtemp = Xcont0(2) & "\" & aitem(ipp, 0) & "\L" & bindex(jpp, 0) & ".COD"
    End If
End Sub

Public Sub SET_DRVindex(Drvtemp As String, aitem() As String, ipp As Integer)
    Drvtemp = Xcont0(2) & "\" & aitem(ipp, 0) & "\" & aitem(ipp, 0) & "INDEX.COD"
End Sub

Public Sub SETdrvmain(Koumoku As String, bcod As String)
    If Koumoku = "IC" Then
        drvmain = Xcont0(2) & "\IC\IC" & Left(bcod, 1) & "\L" & bcod & ".COD"
    Else
        drvmain = Xcont0(2) & "\" & Koumoku & "\L" & bcod & ".COD"
    End If
End Sub

Public Sub SET_Yen0_Format(Indata As Double, Pdata As String, Mojisuu As Integer)
    Dim al As Integer       '*** 指定文字長にする
'
    Pdata = "\" & Trim(Format(Indata, "###,###,##0"))
    al = Len(Pdata)
    Do While al < Mojisuu
        Pdata = " " & Pdata
        al = Len(Pdata)
    Loop
End Sub

Public Sub SET_Yen1_Format(Indata As Double, Pdata As String, Mojisuu As Integer)
    Dim al As Integer       '*** 指定文字長にする
'
    Pdata = "\" & Trim(Format(Indata, "#,###,##0.0"))
    al = Len(Pdata)
    Do While al < Mojisuu
        Pdata = " " & Pdata
        al = Len(Pdata)
    Loop
End Sub

Function LenMbcs(ByVal str As String)           '*** 漢字混じり文字列を正確に数えるおまじない ***
   LenMbcs = LenB(StrConv(str, vbFromUnicode))
End Function

Public Sub SET_migiyose(Indata As String, Pdata As String, Mojisuu As Integer)
    Dim al As Integer       '*** 指定文字長にする
'
    Pdata = Indata
    al = Len(Pdata)
    Do While al < Mojisuu
        Pdata = " " & Pdata
        al = Len(Pdata)
    Loop
End Sub

Public Sub SET_tyuuou(Indata As String, Pdata As String, Mojisuu As Integer)
    Dim al As Integer       '*** 指定文字長にする
    Dim i As Integer
'
    Pdata = Indata
    al = (Mojisuu - LenMbcs(Pdata)) / 2
    For i = 1 To al
        Pdata = " " & Pdata
    Next i
End Sub

Public Sub GET_line(i As Integer, Line_no As String)
'                   *** ３桁の行番号文字列を求める ***
'*** グリッドの表示位置は ColAlignment(col n) を使用する。***
'*** ０：左、１：右、２：中央 ***
'
    If i < 10 Then
        Line_no = " " & str(i)
    ElseIf i < 100 Then
        Line_no = str(i)
    Else
        Line_no = Trim(str(i))
    End If
End Sub

Public Sub GET_shitei1(cmain3 As String, kikaku As String, j As Integer)
    If cmain3 = "0" Then            '*** 使用禁止部品 ***
        kikaku = "!" & kikaku
        j = 1
    ElseIf cmain3 = "3" Then        '*** 在庫限り部品 ***
        kikaku = "?" & kikaku
        j = 1
    ElseIf cmain3 = "4" Then        '*** 変更推奨部品 ***
        kikaku = "\" & kikaku
        j = 1
    Else
        j = 0
    End If
End Sub

Public Sub GET_shitei2(cmain3 As String, kikaku As String, j As Integer)
    If cmain3 = "0" Then            '*** 使用禁止部品 ***
        kikaku = "!" & kikaku
        j = 1
    ElseIf cmain3 = "3" Then        '*** 在庫限り部品 ***
        kikaku = "?" & kikaku
        j = 1
    Else
        j = 0
    End If
End Sub

Public Sub GETsymbol(Pmodu As String, Pitem As String, Pno As String)
'                                       *** 部品の項目と番号に分ける ***
    Dim i As Integer    '*** 番号の始まる文字位置 ***
    Dim j As Integer    '*** ２つ目の番号の始まり位置 ***
    Dim Pname As String, Sdata As String, Qname As String
    Dim kekka As Integer
'
    Pmodu = ""
    Pname = Trim(Pitem)
    i = 2
    Sdata = Mid(Pname, i, 1)
    kekka = Sdata Like "[A-Za-z]"
    Do While kekka = True
        i = i + 1
        Sdata = Mid(Pname, i, 1)
        kekka = Sdata Like "[A-Za-z]"
    Loop
'
    Pitem = Left(Pname, i - 1)
    Pitem = StrConv(Pitem, vbUpperCase)
    Pno = Mid(Pname, i)
'
    If Pitem = "A" Then     '*** "A"の時はモジュールそのものかを判定 ***
        Qname = Pno
        j = 2
        Sdata = Mid(Qname, j, 1)
        kekka = Sdata Like "[0-9]"
        Do While kekka = True
            j = j + 1
            Sdata = Mid(Qname, j, 1)
            kekka = Sdata Like "[0-9]"
        Loop
'
        If j - 1 <> Len(Pno) Then    '*** ２番目の項目番号がある ***
            Pmodu = "A" & Left(Qname, j - 1)
            Pname = Mid(Qname, j)
            i = 2
            Sdata = Mid(Pname, i, 1)
            kekka = Sdata Like "[A-Za-z]"
            Do While kekka = True
                i = i + 1
                Sdata = Mid(Pname, i, 1)
                kekka = Sdata Like "[A-Za-z]"
            Loop
'
            Pitem = Left(Pname, i - 1)
            Pitem = StrConv(Pitem, vbUpperCase)
            Pno = Mid(Pname, i)
        End If
    End If
End Sub

Public Sub RDplstWork()
'                                       *** 部品表作成データ読み込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    ReDim PlstWork(cPLSTWORKmax, cPLSTWORKdim)
'
    On Error GoTo NoFile
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVplstWork For Input As #FILE_number
        For i = 0 To cPLSTWORKmax
            For j = 0 To cPLSTWORKdim
                If EOF(1) = True Then GoTo nukeru
'
                Line Input #FILE_number, PlstWork(i, j)
'
            Next j
        Next i
nukeru:
    Close #FILE_number
    On Error GoTo 0
    Exit Sub
'
NoFile: Beep
    i = MsgBox("部品表作成データファイルが無いので新たに作成します。", vbExclamation, STATUS2)
    Resume YOMANAI
'
YOMANAI:
    On Error GoTo 0
End Sub

Public Sub RDpartlist(Drvtemp As String, Plistnamex As String, Plistdatex As String, Remarksx As String, _
                    Plstx() As String, Ptotalx As Integer, Pdimx As Integer)        '*** 部品表読み込み ***
    Dim i As Integer, j As Integer, PdimTx As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Input As #FILE_number
        Line Input #FILE_number, Plistnamex
        Line Input #FILE_number, Plistdatex
        Line Input #FILE_number, Remarksx
'
        Input #FILE_number, Ptotalx, PdimTx
'
        Pdimx = cPdim0
            If Pdimx < PdimTx Then
        ReDim Plstx(Ptotalx, PdimTx + 3)
            Else
        ReDim Plstx(Ptotalx, Pdimx + 3)
            End If
'
        For i = 1 To Ptotalx
            For j = 0 To PdimTx
                Line Input #FILE_number, Plstx(i, j)
            Next j
        Next i
    Close #FILE_number
'
    For i = 1 To Ptotalx '*** 内容整理 ***
        If Plstx(i, 2) = "" Then Plstx(i, 2) = "*"
        If Plstx(i, 4) = "" Then Plstx(i, 4) = "*"
    Next i
End Sub

Public Sub RDconst_lst(Drvtemp As String, CATnox As String, CATnamex As String, Zubanx As String, Personx As String, _
                Orgdatex As String, Revdatex As String, Checkdatex As String, Outdatex As String, _
                klstx() As String, Ktotalx As Integer, Kdimx As Integer, _
                Koubanx As String, Daisuux As String, Kbikoux As String, KyobiAx As String, KyobiBx As String)
'                   *** 構成表読み込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Input As #FILE_number
        Line Input #FILE_number, CATnox
        Line Input #FILE_number, CATnamex
        Line Input #FILE_number, Zubanx
        Line Input #FILE_number, Personx
        Line Input #FILE_number, Orgdatex
        Line Input #FILE_number, Revdatex
        Line Input #FILE_number, Checkdatex
        Line Input #FILE_number, Outdatex
'
        Input #FILE_number, Ktotalx, Kdimx
            ReDim klstx(Ktotalx, Kdimx)
        For i = 1 To Ktotalx
            For j = 0 To Kdimx
                Line Input #FILE_number, klstx(i, j)
            Next j
        Next i
'
        If EOF(1) = False Then
            Line Input #FILE_number, Koubanx
            Line Input #FILE_number, Daisuux
            Line Input #FILE_number, Kbikoux
            Line Input #FILE_number, KyobiAx
            Line Input #FILE_number, KyobiBx
        Else
            Koubanx = ""
            Daisuux = "1"
            Kbikoux = ""
            KyobiAx = ""
            KyobiBx = ""
        End If
    Close #FILE_number
End Sub

Public Sub GETkikaku(Pdata As String, Bsitei As String, bindex() As String, jp As Integer, cmain() As String, kp As Integer)
'                   *** 部品指定名称取得 ***
    Dim msub1 As Integer, msub2 As Integer
'
    msub1 = 3           '*** 初期値
    msub2 = 4
'
    If bindex(jp, 5) = "998" Then
        If Bsitei = "0" Then
            Pdata = bindex(jp, msub1) + cmain(kp, 1)
'
            If bindex(jp, msub2) <> "*" Then
                Pdata = Pdata + bindex(jp, msub2)
            End If
            Pdata = Pdata + "相当"
        Else
            If Bsitei = "8" Then
                msub1 = 3
                msub2 = 4
            ElseIf Bsitei = "9" Then
                msub1 = 11
                msub2 = 12
            Else
                msub1 = 13
                msub2 = 14
            End If
'
            Pdata = bindex(jp, msub1) + cmain(kp, 1)
'
            If bindex(jp, msub2) <> "*" Then
                Pdata = Pdata + bindex(jp, msub2)
            End If
        End If
'
    Else
        Pdata = bindex(jp, msub1) + cmain(kp, 1)
'
        If bindex(jp, msub2) <> "*" Then
            Pdata = Pdata + bindex(jp, msub2)
        End If
    End If
End Sub

Public Sub GET998maker(Pdata As String, Bsitei As String, bindex() As String, jp As Integer)
'                   *** メーカーｺｰﾄﾞ取得 ***
'
    If Bsitei = "8" Then
        Pdata = bindex(jp, 8)
    ElseIf Bsitei = "9" Then
        Pdata = bindex(jp, 9)
    ElseIf Bsitei = "10" Then
        Pdata = bindex(jp, 10)
    Else
        Pdata = "******"
    End If
End Sub

Public Sub SETfont_size(f_size As Integer, f_name As Integer)
'                   *** ﾌｫﾝﾄｻｲｽﾞの設定 ***
'
    Dim X As Integer, Y As Integer
'
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.Print
    Printer.CurrentX = X
    Printer.CurrentY = Y
'
    Select Case f_name
    Case 2
        Printer.Font.Name = "ＭＳ Ｐゴシック"
    Case 1
        Printer.Font.Name = "ＭＳ ゴシック"
    Case Else
        Printer.Font.Name = "ＭＳ 明朝"
    End Select
'
    Printer.Print
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Font.Size = f_size
End Sub

Public Sub TRSbikouran(Dtemp As String)
'                   *** 特記事項/備考欄印刷の有無 ***
    If Dtemp = "1" Then
        Dtemp = "1:特記事項有り、出庫ラベルに印刷"
    Else
        Dtemp = "0:特記事項無し"
    End If
End Sub

Public Sub TRSkeijou(Dtemp As String)
'                   *** 部品形状変換 ***
    Select Case Dtemp
    Case "0"
        Dtemp = "0:DIP"
    Case "1"
        Dtemp = "1:SK-DIP"
    Case "2"
        Dtemp = "2:SH-DIP"
    Case "3"
        Dtemp = "3:SIP"
    Case "4"
        Dtemp = "4:ZIP"
    Case "5"
        Dtemp = "5:PGA"
    Case "6"
        Dtemp = "6:SOP"
    Case "7"
        Dtemp = "7:QFP"
    Case "8"
        Dtemp = "8:SOJ"
    Case "9"
        Dtemp = "9:PLCC"
    Case "A"
        Dtemp = "A:BGA"
    Case "B"
        Dtemp = "B:FBGA"
    Case "C"
        Dtemp = "C:CSP"
    Case "D"
        Dtemp = "D:UFBGA"
    Case "E"
        Dtemp = "E:SSOP"
    Case "F"
        Dtemp = "F:TSSOP"
    Case Else
        Dtemp = "*:ｿﾉ他"
    End Select
End Sub

Public Sub TRSkeijou2(Dtemp As String)
'                   *** 部品形状変換（簡易表示） ***
    Select Case Dtemp
    Case "0"
        Dtemp = "DIP"
    Case "1"
        Dtemp = "SK-DIP"
    Case "2"
        Dtemp = "SH-DIP"
    Case "3"
        Dtemp = "SIP"
    Case "4"
        Dtemp = "ZIP"
    Case "5"
        Dtemp = "PGA"
    Case "6"
        Dtemp = "SOP"
    Case "7"
        Dtemp = "QFP"
    Case "8"
        Dtemp = "SOJ"
    Case "9"
        Dtemp = "PLCC"
    Case "A"
        Dtemp = "BGA"
    Case "B"
        Dtemp = "FBGA"
    Case "C"
        Dtemp = "CSP"
    Case "D"
        Dtemp = "UFBGA"
    Case "E"
        Dtemp = "SSOP"
    Case "F"
        Dtemp = "TSSOP"
    Case Else
        Dtemp = "ｿﾉ他"
    End Select
End Sub

Public Sub TRSkeijouNo(Dtemp As String)
'                   *** 部品形状変換 ***
    Select Case Dtemp
    Case "0"
        Dtemp = "0"
    Case "1"
        Dtemp = "1"
    Case "2"
        Dtemp = "2"
    Case "3"
        Dtemp = "3"
    Case "4"
        Dtemp = "4"
    Case "5"
        Dtemp = "5"
    Case "6"
        Dtemp = "6"
    Case "7"
        Dtemp = "7"
    Case "8"
        Dtemp = "8"
    Case "9"
        Dtemp = "9"
    Case "A"
        Dtemp = "10"
    Case "B"
        Dtemp = "11"
    Case "C"
        Dtemp = "12"
    Case "D"
        Dtemp = "13"
    Case "E"
        Dtemp = "14"
    Case "F"
        Dtemp = "15"
    Case Else
        Dtemp = "16"
    End Select
End Sub

Public Sub TRSkeijouKigou(TempAno As Integer, Dtemp As String)
'                   *** 部品形状変換 ***
    Select Case TempAno
    Case 0
        Dtemp = "0"
    Case 1
        Dtemp = "1"
    Case 2
        Dtemp = "2"
    Case 3
        Dtemp = "3"
    Case 4
        Dtemp = "4"
    Case 5
        Dtemp = "5"
    Case 6
        Dtemp = "6"
    Case 7
        Dtemp = "7"
    Case 8
        Dtemp = "8"
    Case 9
        Dtemp = "9"
    Case 10
        Dtemp = "A"
    Case 11
        Dtemp = "B"
    Case 12
        Dtemp = "C"
    Case 13
        Dtemp = "D"
    Case 14
        Dtemp = "E"
    Case 15
        Dtemp = "F"
    Case Else
        Dtemp = "*"
    End Select
End Sub

Public Sub TRSsitei(Dtemp As String)
'                   *** 部品指定変換 ***
    Select Case Dtemp
    Case "0"
        Dtemp = "0:使用禁止部品"
    Case "1"
        Dtemp = "1:標準部品"
    Case "2"
        Dtemp = "2:客先指定部品"
    Case "3"
        Dtemp = "3:在庫限り部品"
    Case "4"
        Dtemp = "4:変更推奨部品"
    Case Else
        Dtemp = "*:非標準部品"
    End Select
End Sub

Public Sub TRSsitei2(Dtemp As String)
'                   *** 部品指定変換簡易表示 ***
    Select Case Dtemp
    Case "0"
        Dtemp = "使用禁止部品"
    Case "1"
        Dtemp = "標準部品"
    Case "2"
        Dtemp = "客先指定部品"
    Case "3"
        Dtemp = "在庫限り部品"
    Case "4"
        Dtemp = "変更推奨部品"
    Case Else
        Dtemp = "非標準部品"
    End Select
End Sub

Public Sub TRSsitei3(Dtemp As String)
'                   *** 部品指定変換超簡易表示 ***
    Select Case Dtemp
    Case "0"
        Dtemp = "使用禁止"
    Case "1"
        Dtemp = "  標準  "
    Case "2"
        Dtemp = "客先指定"
    Case "3"
        Dtemp = "在庫限り"
    Case "4"
        Dtemp = "変更推奨"
    Case Else
        Dtemp = " 非標準 "
    End Select
End Sub

Sub Main()

End Sub

Public Sub Makerget1(makername As String)
'                   *** メーカー正式名 取得 ***
    Dim k As Integer
'
    If makername = "******" Then
        makername = "複数指定"
        Exit Sub
'
    Else
        For k = 1 To Maknum0
            If makername = Maker(k, 0) Then
                makername = Maker(k, 2)
                Exit Sub
'
            End If
        Next k
    End If
End Sub

Public Sub Makerget2(makername As String)
'                   *** メーカー略称 取得 ***
    Dim k As Integer
'
    If makername = "******" Then
        makername = "複数指"
        Exit Sub
'
    Else
        For k = 1 To Maknum0
            If makername = Maker(k, 0) Then
                makername = Maker(k, 1)
                Exit Sub
'
            End If
        Next k
    End If
End Sub

Public Sub RDcont()
'                    *** 環境設定読み込み ***
    Dim i As Integer
    Dim FILE_number As Integer
    Dim xUP As Integer
'
    On Error GoTo errh_K
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DIRcont0 & "\control0.dat" For Input As #FILE_number
        Input #FILE_number, Xnum
'
        ReDim Xcont0(cXnum0)
        For i = 1 To Xnum
            Input #FILE_number, Xcont0(i)
        Next i
    Close #FILE_number
'
    Xnum = cXnum0
    FLG_job_error_end = 0
'
    If Xcont0(3) = "*" Then     '*** Ver1.3 にて構成表と部品表のフォルダ指定が追加/変更になった。 ***
        Xcont0(3) = Xcont0(4)
    End If
'
    If Xcont0(17) = "0" Or Xcont0(17) = "1" Then
        '*** Ver2.03 にて オプション設定/環境設定/「個別//共有」フォルダー選択を追加 ***
    Else
        Xcont0(17) = "0"    '*** 共有選択 ***
    End If
'
    Exit Sub
'
errh_K: Beep
    DoEvents
    Resume Settei
'
Settei:
    i = MsgBox("「環境設定ファイル」が無いか、内容が正しくありません。" & vbCrLf & _
        "環境設定・オプション設定をもう一度行ってください。", vbInformation, "EEOS2")
'
    ReDim Xcont0(cXnum0)
    Call Kankyou_syokika    '*** 環境設定初期化 ***
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DIRcont0 & "\control0.dat" For Output As #FILE_number
        Write #FILE_number, cXnum0
'
        For i = 1 To cXnum0
            Print #FILE_number, Xcont0(i)
        Next i
    Close #FILE_number
'
    Xnum = cXnum0
    FLG_job_error_end = 1
    On Error GoTo 0
'
End Sub

Public Sub RDindex(Drvtemp As String, bindex() As String, bnum As Integer, bdim As Integer)
                    '*** INDEX.COD 読み込み ***
    Dim i As Integer, j As Integer, BdimT As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Input As #FILE_number
        Input #FILE_number, bnum, BdimT
'
        If BdimT < cBdim0 Then bdim = cBdim0 Else bdim = BdimT
        ReDim bindex(bnum + 1, bdim)
        For i = 1 To bnum
            For j = 0 To BdimT
                Line Input #FILE_number, bindex(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub RDitem(Drvtemp As String, aitem() As String, anum As Integer, adim As Integer)
                                    '*** ITEM.COD 読み込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    On Error GoTo NoFile
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Input As #FILE_number
        Input #FILE_number, anum, adim
'
        ReDim aitem(anum + 1, adim)
        For i = 1 To anum
            For j = 0 To adim
                Line Input #FILE_number, aitem(i, j)
            Next j
        Next i
    Close #FILE_number
    FLG_job_error_end = 0
    On Error GoTo 0
    Exit Sub
'
NoFile: Beep
    Resume error_skip
'
error_skip:
    i = MsgBox("「部品コード表」が環境設定で指定された場所に見当たりません。", vbExclamation Or vbOKOnly, HeadTitle)
    FLG_job_error_end = 1
    On Error GoTo 0
End Sub

Public Sub RDmain(Drvtemp As String, cmain() As String, cnum As Integer, cdim As Integer)
                    '*** MAIN.COD 読み込み ***
    Dim i As Integer, j As Integer, CdimT As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Input As #FILE_number
        Input #FILE_number, cnum, CdimT
'
        If CdimT < cCdim0 Then cdim = cCdim0 Else cdim = CdimT
        ReDim cmain(cnum + 1, cdim)
'
        For i = 1 To cnum
            For j = 0 To CdimT
                Line Input #FILE_number, cmain(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub RDmaker()
                    '*** メーカーコード 読み込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    On Error GoTo NoFile
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVmaker For Input As #FILE_number
        Input #FILE_number, Maknum0, Makdim0
'
        ReDim Maker(Maknum0 + 1, Makdim0)
        For i = 1 To Maknum0
            For j = 0 To Makdim0
                Line Input #FILE_number, Maker(i, j)
            Next j
        Next i
    Close #FILE_number
    FLG_job_error_end = 0
    On Error GoTo 0
    Exit Sub
'
NoFile: Beep
    Resume error_skip
'
error_skip:
    i = MsgBox("「メーカーコード」が環境設定で指定された場所に見当たりません。", vbExclamation Or vbOKOnly, HeadTitle)
    FLG_job_error_end = 1
    On Error GoTo 0
End Sub

Public Sub RDtrader()
                    '*** 商社コード 読み込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    On Error GoTo NoFile
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVtrader For Input As #FILE_number
        Input #FILE_number, Trdnum0, Trddim0
'
        ReDim Trader(Trdnum0 + 1, Trddim0)
        For i = 1 To Trdnum0
            For j = 0 To Trddim0
                Line Input #FILE_number, Trader(i, j)
            Next j
        Next i
    Close #FILE_number
    FLG_job_error_end = 0
    On Error GoTo 0
    Exit Sub
'
NoFile:
    Resume error_skip
'
error_skip:
    i = MsgBox("「商社コード」が環境設定で指定された場所に見当たりません。", vbExclamation Or vbOKOnly, HeadTitle)
    FLG_job_error_end = 1
    On Error GoTo 0
End Sub

Public Sub TRSsyukko(Dtemp As String)
'                   *** 出庫ラベル印刷変換 ***
    If Dtemp = "1" Then
        Dtemp = "1:ラベル印刷する"
    Else
        Dtemp = "0:払い出さない"
    End If
End Sub

Public Sub TRSsyukko2(Dtemp As String)
'                   *** 出庫ラベル印刷変換簡易表示 ***
    If Dtemp = "1" Then
        Dtemp = "出庫"
    Else
        Dtemp = "非出庫"
    End If
End Sub

Public Sub TRStouroku(Dtemp As String)
'                   *** CAD登録有無 ***
    If Dtemp = "1" Then
        Dtemp = "1:済"
    Else
        Dtemp = "*:未"
    End If
End Sub

Public Sub TRStouroku2(Dtemp As String)
'                   *** CAD登録有無簡易表示 ***
    If Dtemp = "1" Then
        Dtemp = "済"
    Else
        Dtemp = "--"
    End If
End Sub

Public Sub TRS_Mlevel(Dtemp As String)
'                   *** MSL(吸湿ﾚﾍﾞﾙ)表示 ***
    If Dtemp = "0" Then
        Dtemp = "0:対象外"
    ElseIf Dtemp = "1" Then
        Dtemp = "1"
    ElseIf Dtemp = "2" Then
        Dtemp = "2"
    ElseIf Dtemp = "3" Then
        Dtemp = "2a"
    ElseIf Dtemp = "4" Then
        Dtemp = "3"
    ElseIf Dtemp = "5" Then
        Dtemp = "4"
    ElseIf Dtemp = "6" Then
        Dtemp = "5"
    ElseIf Dtemp = "7" Then
        Dtemp = "5a"
    ElseIf Dtemp = "8" Then
        Dtemp = "6:JEDEC"
    ElseIf Dtemp = "9" Then
        Dtemp = "2S"
    ElseIf Dtemp = "10" Then
        Dtemp = "2aS"
    ElseIf Dtemp = "11" Then
        Dtemp = "3S"
    ElseIf Dtemp = "12" Then
        Dtemp = "4S"
    ElseIf Dtemp = "13" Then
        Dtemp = "5S"
    ElseIf Dtemp = "14" Then
        Dtemp = "5aS"
    Else    '"*" etc
        Dtemp = "---"
    End If
End Sub

Public Sub TRS_Mlevel2(Dtemp As String)
'                   *** MSL(吸湿ﾚﾍﾞﾙ)表示 ***
    If Dtemp = "0" Then
        Dtemp = " 0 "
    ElseIf Dtemp = "1" Then
        Dtemp = " 1 "
    ElseIf Dtemp = "2" Then
        Dtemp = " 2 "
    ElseIf Dtemp = "3" Then
        Dtemp = "2a "
    ElseIf Dtemp = "4" Then
        Dtemp = " 3 "
    ElseIf Dtemp = "5" Then
        Dtemp = " 4 "
    ElseIf Dtemp = "6" Then
        Dtemp = " 5 "
    ElseIf Dtemp = "7" Then
        Dtemp = "5a "
    ElseIf Dtemp = "8" Then
        Dtemp = " 6 "
    ElseIf Dtemp = "9" Then
        Dtemp = " 2S"
    ElseIf Dtemp = "10" Then
        Dtemp = "2aS"
    ElseIf Dtemp = "11" Then
        Dtemp = " 3S"
    ElseIf Dtemp = "12" Then
        Dtemp = " 4S"
    ElseIf Dtemp = "13" Then
        Dtemp = " 5S"
    ElseIf Dtemp = "14" Then
        Dtemp = "5aS"
    Else    '"*" etc
        Dtemp = "---"
    End If
End Sub

Public Sub WRconst_lst(Drvtemp As String, CATnox As String, CATnamex As String, Zubanx As String, Personx As String, _
                Orgdatex As String, Revdatex As String, Checkdatex As String, Outdatex As String, _
                klstx() As String, Ktotalx As Integer, Kdimx As Integer, _
                Koubanx As String, Daisuux As String, Kbikoux As String, KyobiAx As String, KyobiBx As String)
'                   *** 構成表書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
On Error GoTo Wr_inhibit
'
kaisi:
    Open Drvtemp For Output As #FILE_number
        Print #FILE_number, CATnox
        Print #FILE_number, CATnamex
        Print #FILE_number, Zubanx
        Print #FILE_number, Personx
        Print #FILE_number, Orgdatex
        Print #FILE_number, Revdatex
        Print #FILE_number, Checkdatex
        Print #FILE_number, Outdatex
'
        Write #FILE_number, Ktotalx, Kdimx
        For i = 1 To Ktotalx
            For j = 0 To Kdimx
                Print #FILE_number, klstx(i, j)
            Next j
        Next i
'
        Print #FILE_number, Koubanx
        Print #FILE_number, Daisuux
        Print #FILE_number, Kbikoux
        Print #FILE_number, KyobiAx
        Print #FILE_number, KyobiBx
    Close #FILE_number
Exit Sub
'
Wr_inhibit:
    Beep
    i = MsgBox("構成表を収容するﾌｫﾙﾀﾞｰが書き込み禁止です！", vbExclamation Or vbOKCancel, STATUS)
    If i = vbOK Then
        FLG_job_error_end = 0   '*** フラグりセット ***
        Resume kaisi
    End If
'
    Resume Utikiri
'
Utikiri:
    FLG_job_error_end = 1   '*** フラグセット ***
    On Error GoTo 0
End Sub

Public Sub WRcont()
'                    *** 環境設定 書き込み ***
    Dim i As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DIRcont0 & "\control0.dat" For Output As #FILE_number
        Write #FILE_number, Xnum
'
        For i = 1 To Xnum
            Print #FILE_number, Xcont0(i)
        Next i
    Close #FILE_number
End Sub

Public Sub WRindex(Drvtemp As String, bindex() As String, bnum As Integer, bdim As Integer)
                    '*** INDEX.COD 書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Output As #FILE_number
                Write #FILE_number, bnum, bdim
'
        For i = 1 To bnum
            For j = 0 To bdim
                Print #FILE_number, bindex(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub WRitem(Drvtemp As String, aitem() As String, anum As Integer, adim As Integer)
                                        '*** ITEM.COD 書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Output As #FILE_number
                Write #FILE_number, anum, adim
'
        For i = 1 To anum
            For j = 0 To adim
                Print #FILE_number, aitem(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub WRmain(Drvtemp As String, cmain() As String, cnum As Integer, cdim As Integer)
                    '*** MAIN.COD 書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open Drvtemp For Output As #FILE_number
                Write #FILE_number, cnum, cdim
'
        For i = 1 To cnum
            For j = 0 To cdim
                Print #FILE_number, cmain(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub WRmaker()
                    '*** メーカーコード 書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVmaker For Output As #FILE_number
        Write #FILE_number, Maknum0, Makdim0
'
        For i = 1 To Maknum0
            For j = 0 To Makdim0
                Print #FILE_number, Maker(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub WRtrader()
                    '*** 商社コード 書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVtrader For Output As #FILE_number
        Write #FILE_number, Trdnum0, Trddim0
'
        For i = 1 To Trdnum0
            For j = 0 To Trddim0
                Print #FILE_number, Trader(i, j)
            Next j
        Next i
    Close #FILE_number
End Sub

Public Sub WRpartlist(Drvtemp As String, Plistnamex As String, Plistdatex As String, Remarksx As String, _
                    Plstx() As String, Ptotalx As Integer, Pdimx As Integer)        '*** 部品表書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
'
On Error GoTo Wr_inhibit
'
kaisi:
    Open Drvtemp For Output As #FILE_number
        Print #FILE_number, Plistnamex
        Print #FILE_number, Plistdatex
        Print #FILE_number, Remarksx
        Write #FILE_number, Ptotalx, Pdimx
'
        For i = 1 To Ptotalx
            For j = 0 To Pdimx
                Print #FILE_number, Plstx(i, j)
            Next j
        Next i
    Close #FILE_number
Exit Sub
'
Wr_inhibit:
    Beep
    i = MsgBox("部品表を収容するﾌｫﾙﾀﾞｰが書き込み禁止です！", vbOKCancel, STATUS)
    If i = vbOK Then
        FLG_job_error_end = 0   '*** フラグりセット ***
        Resume kaisi
    End If
'
    Resume Utikiri
Utikiri:
    FLG_job_error_end = 1   '*** フラグセット ***
End Sub

Public Sub ADD_kuuhaku(Moji As String, Jisuu As Long)
    Do While Len(Moji) <= Jisuu
        Moji = Moji & " "
    Loop
End Sub

Public Sub WRplstWork()
'                   *** 部品表作成データ書き込み ***
    Dim i As Integer, j As Integer
    Dim FILE_number As Integer
'
On Error GoTo Wr_inhibit
'
kaisi:
    FILE_number = FreeFile     '*** 空いているファイル番号を得る ***
    Open DRVplstWork For Output As #FILE_number
        For i = 0 To cPLSTWORKmax
            For j = 0 To cPLSTWORKdim
                Print #FILE_number, PlstWork(i, j)
            Next j
        Next i
    Close #FILE_number
Exit Sub
'
Wr_inhibit:
    Beep
    i = MsgBox("部品表作成データを収容するﾌｫﾙﾀﾞｰが書き込み禁止です！", vbOKCancel, STATUS)
    If i = vbOK Then
        FLG_job_error_end = 0   '*** フラグりセット ***
        Resume kaisi
    End If
'
    Resume Utikiri
Utikiri:
    FLG_job_error_end = 1   '*** フラグセット ***
End Sub

Public Sub CHGplstWork(point As Integer)
'                   *** 部品表作成データ順番組み替え ***
    Dim i As Integer, j As Integer
    Dim Tmp(4) As String
'
    If point > 0 Then   '*** 0の時は何もしない。 ***
        For j = 0 To cPLSTWORKdim
            Tmp(j) = PlstWork(point, j)
        Next j
'
        For i = point To 1 Step -1
            For j = 0 To cPLSTWORKdim
                PlstWork(i, j) = PlstWork(i - 1, j)
            Next j
        Next i
'
        For j = 0 To cPLSTWORKdim
            PlstWork(0, j) = Tmp(j)
        Next j
    End If
End Sub

Public Sub GETtrader1(Tradername As String)
'                   *** 商社正式名 取得 ***
    Dim k As Integer
'
    For k = 1 To Trdnum0
        If Tradername = Trader(k, 0) Then
            Tradername = Trader(k, 1)
            Exit Sub
'
        End If
    Next k
End Sub

Public Sub GETtrader2(Tradername As String)
'                   *** 商社略称 取得 ***
    Dim k As Integer
'
    For k = 1 To Trdnum0
        If Tradername = Trader(k, 0) Then
            Tradername = Trader(k, 2)
            Exit Sub
'
        End If
    Next k
End Sub

Public Sub Kankyou_syokika()    '*** 環境設定初期化 ***
    Xcont0(1) = "0"         '*** 無選択 ***
    Xcont0(2) = DIRcont0
    Xcont0(3) = DIRcont0
    Xcont0(4) = DIRcont0
    Xcont0(5) = "*" 'DIRcont0
    Xcont0(6) = "*"
    Xcont0(7) = "*"
    Xcont0(8) = "*"
    Xcont0(9) = "*"
    Xcont0(10) = "*"
    Xcont0(11) = "*"
    Xcont0(12) = "*"
    Xcont0(13) = DIRcont0
    Xcont0(14) = "*"
    Xcont0(15) = "*"
    Xcont0(16) = "0"
    Xcont0(17) = "0"
End Sub

Public Sub setFullName(hyouki As String)
                                    '*** 図面表記データに「<ﾌﾙﾈｰﾑ>」があったら「ﾌﾙﾈｰﾑ」にする ***
    Dim i As Long
'
    i = InStr(1, hyouki, "<ﾌﾙﾈｰﾑ>")
    If i <> 0 Then
        hyouki = "ﾌﾙﾈｰﾑ"
    End If
End Sub

Public Sub removeFullName(hyouki As String)
                            '*** 図面表記データに「<ﾌﾙﾈｰﾑ>,<ﾅｼ>」があったら文字列から削除する ***
    Dim i As Long
'
    i = InStr(1, hyouki, "<ﾌﾙﾈｰﾑ>")
    If i <> 0 Then
        hyouki = Mid$(hyouki, 1, i - 1)
    Else
        i = InStr(1, hyouki, "<ﾅｼ>")
        If i <> 0 Then
            hyouki = Mid$(hyouki, 1, i - 1)
        End If
    End If
End Sub

Public Sub remove_PLT(Indata As String, Outdata As String)
    Dim i As Integer
'
    If Right(UCase(Indata), 4) = ".PLT" Then
        i = Len(Indata)
        Outdata = Left(Indata, i - 4)
    Else
        If Mid(Indata, 9, 1) = "." Then
            Outdata = Left(Indata, 8) & Mid(Indata, 10)
        Else
            Outdata = Indata
        End If
    End If
End Sub
