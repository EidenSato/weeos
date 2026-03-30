Attribute VB_Name = "EEOS_menu"
'********************************
'***   ＥＥＯＳ２ メニュー    ***
'***                          ***
'*** 2006.06.23 by S.Fukazawa ***
'********************************
'
Option Compare Binary
Option Explicit
'

Public Sub mnuDenkiKouseihyou()
    If FLGconst = 1 Then
        Const_main.SetFocus
    Else
        FLGjob = 1
        FLGshinki = 0       '*** 新規フラグクリアー ***
        STATUS = "電気 構成表"
'
        FLGesc = 0
        If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
            TMPdir1 = Xcont0(3)
        Else
            Sele_Dir.Show 1         '*** フォルダー選択 => TMPdir1 ***
        End If
'
        If FLGesc = 0 Then
            Const_DirR.Show 1       '*** 読み込みダイアログ ***
'
            If FLGesc = 0 Then
                Const_main.Show
            End If
        End If
    End If
End Sub

Public Sub mnuDenkiBuhinhyou()
    If FLGplst = 1 Then
        Plst_main.SetFocus
    Else
        FLGjob = 2
        FLGlevel = 1    '*** 部品表 更新・確認 ***
        STATUS = "電気 部品表"
'
        FLGesc = 0
        If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
            TMPdir1 = Xcont0(3)
            TMPdir2 = Xcont0(4)
        Else
            If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
                FLGjob = 1
                Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
            End If
'
            If FLGesc = 0 And FLGplst2 = 0 Then     '*** 部品表２画面が開いていない ***
                FLGjob = 2
                Sele_Dir.Show 1 '*** フォルダー選択 => TMPdir2, FLGesc ***
            End If
        End If
'
        If FLGesc = 0 Then
            Plst_DirR.Show 1
'
            If FLGesc = 0 Then
                Plst_main.Show
            End If
        End If
    End If
End Sub

Public Sub mnuDenkiBuhinhyou2()
    If FLGplst2 = 1 Then
        Plst_main2.SetFocus
    Else
        FLGjob = 2
        FLGlevel = 1    '*** 部品表 更新・確認 ***
        STATUS = "電気 部品表２"
'
        FLGesc = 0
        If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
            TMPdir1 = Xcont0(3)
            TMPdir2 = Xcont0(4)
        Else
            If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
                FLGjob = 1
                Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
            End If
'
                 
            If FLGesc = 0 And FLGplst = 0 Then      '*** 部品表１画面が開いていない ***
                FLGjob = 2
                Sele_Dir.Show 1 '*** フォルダー選択 => TMPdir2, FLGesc ***
            End If
        End If
'
        If FLGesc = 0 Then
            Plst_DirR.Show 1
'
            If FLGesc = 0 Then
                Plst_main2.Show
            End If
        End If
    End If
End Sub

Public Sub mnuOrCAD_Henkan()
    Dim i As Integer
'
    If FLGplst = 1 And FLGplst2 = 1 Then    '*** 部品表２画面とも既に開いている ***
        Exit Sub
    End If
'
    FLGjob = 2
    FLGlevel = 2    '*** 部品表 OrCADﾃﾞｰﾀの変換 ***
    STATUS = "OrCADﾃﾞｰﾀの変換"
'
    FLGesc = 0
    If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
        TMPdir1 = Xcont0(3)
        TMPdir2 = Xcont0(3)
    Else
        If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
            FLGjob = 1
            Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
        End If
'
        If FLGesc = 0 And FLGplst = 0 And FLGplst2 = 0 Then '*** 部品表２画面とも開いていない ***
            FLGjob = 2
            Sele_Dir.Show 1 '*** フォルダー選択 => TMPdir2, FLGesc ***
        End If
    End If
'
    i = InStr(1, UCase(TMPdir1), "PARTLIST")
    TMPdir3 = Left$(TMPdir1, i - 1) & "CAD_PLST"
'
    If FLGesc = 0 Then
        Plst_DirR.Show 1
'
        If FLGesc = 0 Then
            Plst_OrCAD.Show
        End If
    End If
End Sub

Public Sub mnuConvFile_Edit()
    If FLGplstWork = 1 Then
        Plst_convert.SetFocus
    Else
        FLGjob = 2
        FLGlevel = 4    '*** OrCAD変換 作業ﾌｧｲﾙ 編集 ***
        STATUS = "OrCAD変換 作業ﾌｧｲﾙ 編集"
'
        FLGesc = 0
        If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
            TMPdir1 = Xcont0(3)
        Else
            If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
                Sele_Dir.Show 1         '*** フォルダー選択 => TMPdir1 ***
            End If
        End If
'
        If FLGesc = 0 Then
            DRVplstWork = TMPdir1 & "\PLSTWORK.DAT"
'
            Plst_convert.Show
        End If
    End If
End Sub

Public Sub mnuStandardBuhinhyouPrint()
    FLGjob = 2
    FLGlevel = 5    '*** 部品表 標準部品表印刷 ***
    STATUS = "標準部品表印刷"
'
    FLGesc = 0
    If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
        TMPdir1 = Xcont0(3)         '*** 構成表(個別)フォルダ ***
        TMPdir2 = Xcont0(4)         '*** 部品表(共有)フォルダ ***
    Else
        If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
            FLGjob = 1
            Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
        End If
'
        If FLGesc = 0 And FLGplst = 0 And FLGplst2 = 0 Then '*** 部品表２画面とも開いていない ***
            FLGjob = 2
            Sele_Dir.Show 1 '*** フォルダー選択 => TMPdir2, FLGesc ***
        End If
    End If
'
    If FLGesc = 0 Then
        Plst_DirR.Show 1
'
        If FLGesc = 0 Then
            FLGfile = 2                 '*** 未定 ***
            Plst_PRNstd.Show 1  '*** 標準部品表印刷/ファイル出力 ***
        End If
    End If
End Sub

Public Sub mnuBuhinItiranhyouPrint()
    FLGjob = 2
    FLGlevel = 6    '*** 部品表 一覧表印刷 ***
    STATUS = "部品 一覧表"
'
    FLGesc = 0
    If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
        TMPdir1 = Xcont0(3)
        TMPdir2 = Xcont0(4)
    Else
        If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
            FLGjob = 1
            Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
        End If
'
        If FLGesc = 0 And FLGplst = 0 And FLGplst2 = 0 Then '*** 部品表２画面とも開いていない ***
            FLGjob = 2
            Sele_Dir.Show 1 '*** フォルダー選択 => TMPdir2, FLGesc ***
        End If
    End If
'
    If FLGesc = 0 Then
        Plst_DirR.Show 1
'
        If FLGesc = 0 Then
            FLGfile = 2         '*** 未定 ***
            Plst_PRNlst.Show 1  '*** 一覧表印刷/ファイル出力 ***
        End If
    End If
End Sub

Public Sub mnuBuhinSuuryohyouPrint()
    FLGjob = 2
    FLGlevel = 7    '*** 部品表 数量表印刷 ***
    STATUS = "部品 数量表"
'
    FLGesc = 0
    If Xcont0(16) = "1" Then    '*** オプション設定 <固定> ***
        TMPdir1 = Xcont0(3)
        TMPdir2 = Xcont0(4)
    Else
        If FLGconst = 0 Then    '*** 構成表画面が開いていない ***
            FLGjob = 1
            Sele_Dir.Show 1     '*** フォルダー選択 => TMPdir1, FLGesc ***
        End If
'
        If FLGesc = 0 And FLGplst = 0 And FLGplst2 = 0 Then '*** 部品表２画面とも開いていない ***
            FLGjob = 2
            Sele_Dir.Show 1 '*** フォルダー選択 => TMPdir2, FLGesc ***
        End If
    End If
'
    If FLGesc = 0 Then
        Plst_DirR.Show 1
'
        If FLGesc = 0 Then
            Plst_Suuryo.Show 1
        End If
    End If
End Sub

Public Sub mnuCodeBuhinMaintenance()
    If FLGitem = 1 Then
        Pcod_item.SetFocus
    Else
        Pcod_item.Show
    End If
End Sub

Public Sub mnuCodeHinsyuMaintenance()
    If FLGindex = 1 Then
        Pcod_index.SetFocus
    Else
        Pcod_index.Show
    End If
End Sub

Public Sub mnuCodePmainMaintenance()
    If FLGmain = 1 Then
        Pcod_main.SetFocus
    Else
        Pcod_main.Show
    End If
End Sub

Public Sub mnuCodeMakerMaintenance()
    If FLGmaker = 1 Then
        Maker_main.SetFocus
    Else
        Maker_main.Show
    End If
End Sub

Public Sub mnuCodeTraderMaintenance()
    If FLGtrader = 1 Then
        Trader_main.SetFocus
    Else
        Trader_main.Show
    End If
End Sub

Public Sub mnuKankyouSettei()
    Kankyow_Itiran.Show 1
End Sub

Public Sub mnuOptionSettei()
    Option_Sel.Show 1   '*** オプション設定 ***
End Sub

Public Sub mnuSousaSetumei()
    FLG_Setumei = 0     '*** 0: 操作説明 ***
    Setumei_gamen.Show 1
End Sub

Public Sub mnuKaihanRireki()
    FLG_Setumei = 1     '*** 1: 改版履歴 ***
    Setumei_gamen.Show 1
End Sub

Public Sub mnuVersionGamen()
    Version_gamen.Show 1
End Sub

