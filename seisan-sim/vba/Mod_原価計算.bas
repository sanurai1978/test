Attribute VB_Name = "Mod_原価計算"
Option Explicit
'==============================================================
' Mod_原価計算
' 【目的】品目ごとの標準原価と販売単価を計算する共通部品。
'         品目シートの計算列、出荷の販売単価、レポートの在庫金額で使う。
' 【業務背景】HTML 版と同じ式にする。
'   ・購入品の標準原価 ＝ 購入単価
'   ・製造品の標準原価 ＝ 材料費＋社内加工費＋外注加工費（全階層の合計）
'       社内加工費 ＝ 各工程の 加工時間（分/個）÷60 × 作業区の賃率（円/時）
'                    （段取時間は含めない）
'       外注加工費 ＝ 外注工程の外注単価
'       子品目の原価 × 員数 を親に積み上げる
'   ・販売単価 ＝ 品目マスタの販売単価。空欄なら標準原価×1.35 を
'                 10円単位に切り上げ
' 【前提】BOM に循環が無いこと（循環があると無限に積み上げるため、
'         20段を超えたら打ち切る）。
' 【流れ】品目・BOM・工程表・作業区を配列に読む → 品目ごとに再帰で積み上げ
'==============================================================

Private Const Lg_最大段数 As Long = 20      ' BOM の深さの上限（循環で止まらなくなるのを防ぐ）

' 事業部の全品目の標準原価（丸め前）を返す
'   引数：St_事業部　戻り値：Dictionary（品番 → 標準原価。丸めていない値）
'   表示や金額計算に使うときは Fn_丸め を通す（HTML 版の stdCost は r2 済み）。
Public Function Fn_標準原価集計(ByVal St_事業部 As String) As Object
    Dim Ob_結果 As Object
    Dim Va_品目 As Variant
    Dim Va_BOM As Variant
    Dim Va_工程 As Variant
    Dim Ob_賃率 As Object
    Dim Lg_行 As Long
    Set Ob_結果 = CreateObject("Scripting.Dictionary")
    Va_品目 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_品目), 列_品目.次列 - 1)
    Va_BOM = Fn_表読込(ThisWorkbook.Worksheets(St_シート_BOM), 列_BOM.次列 - 1)
    Va_工程 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_工程表), 列_工程表.次列 - 1)
    Set Ob_賃率 = Fn_賃率辞書(St_事業部)
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_事業部 Then
            Call Fn_品目原価(St_事業部, Fn_文字(Va_品目(Lg_行, 列_品目.品番)), Va_品目, Va_BOM, Va_工程, Ob_賃率, Ob_結果, 0)
        End If
    Next Lg_行
    Set Fn_標準原価集計 = Ob_結果
End Function

' 販売単価（品目マスタの値。空欄なら標準原価の35%増しを10円単位に切り上げ）
'   引数：Va_マスタ単価（品目シートの販売単価のセル値）、Db_標準原価（丸め済み）
'   戻り値：販売単価（円）
Public Function Fn_販売単価(ByVal Va_マスタ単価 As Variant, ByVal Db_標準原価 As Double) As Double
    If Fn_文字(Va_マスタ単価) <> "" Then
        Fn_販売単価 = Fn_数(Va_マスタ単価)
    Else
        ' 10円単位に切り上げ（HTML 版の Math.ceil(x*1.35/10)*10。誤差で1段上がらないよう微小値を引く）
        Fn_販売単価 = -Int(-(Db_標準原価 * 1.35 / 10 - Db_誤差)) * 10
    End If
End Function

' 作業区の賃率（作業区コード → 円/時）
Private Function Fn_賃率辞書(ByVal St_事業部 As String) As Object
    Dim Ob_結果 As Object
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Set Ob_結果 = CreateObject("Scripting.Dictionary")
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_作業区), 列_作業区.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_作業区.事業部)) = St_事業部 Then
            Ob_結果.Item(Fn_文字(Va_表(Lg_行, 列_作業区.作業区))) = Fn_数(Va_表(Lg_行, 列_作業区.賃率))
        End If
    Next Lg_行
    Set Fn_賃率辞書 = Ob_結果
End Function

' 1品目の標準原価を計算して Ob_結果 に入れ、その値を返す（子品目は再帰で先に計算）
'   引数：St_事業部、St_品番、各表の配列、Ob_賃率、Ob_結果（計算済みのメモ）、Lg_段数
'   戻り値：標準原価（丸め前）。品目マスタに無い品番は 0。
Private Function Fn_品目原価(ByVal St_事業部 As String, ByVal St_品番 As String, _
        ByVal Va_品目 As Variant, ByVal Va_BOM As Variant, ByVal Va_工程 As Variant, _
        ByVal Ob_賃率 As Object, ByVal Ob_結果 As Object, ByVal Lg_段数 As Long) As Double
    Dim Lg_行 As Long
    Dim Lg_品目行 As Long
    Dim Db_原価 As Double
    ' 計算済みならそれを返す
    If Ob_結果.Exists(St_品番) Then
        Fn_品目原価 = Ob_結果.Item(St_品番)
        Exit Function
    End If
    ' 循環などで深くなりすぎたら打ち切る
    If Lg_段数 > Lg_最大段数 Then
        Fn_品目原価 = 0
        Exit Function
    End If
    ' 品目マスタの行を探す
    Lg_品目行 = 0
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_事業部 And Fn_文字(Va_品目(Lg_行, 列_品目.品番)) = St_品番 Then
            Lg_品目行 = Lg_行
            Exit For
        End If
    Next Lg_行
    If Lg_品目行 = 0 Then
        Fn_品目原価 = 0
        Exit Function
    End If
    ' 購入品は購入単価
    If Not Fn_製造品か(Fn_文字(Va_品目(Lg_品目行, 列_品目.区分))) Then
        Db_原価 = Fn_数(Va_品目(Lg_品目行, 列_品目.購入単価))
    Else
        Db_原価 = 0
        ' 自分の工程の加工費（社内は 加工時間/60×賃率、外注は外注単価）
        For Lg_行 = 1 To Fn_行数(Va_工程)
            If Fn_文字(Va_工程(Lg_行, 列_工程表.事業部)) = St_事業部 And Fn_文字(Va_工程(Lg_行, 列_工程表.品番)) = St_品番 Then
                If Fn_文字(Va_工程(Lg_行, 列_工程表.外注)) = St_印 Then
                    Db_原価 = Db_原価 + Fn_数(Va_工程(Lg_行, 列_工程表.外注単価))
                Else
                    Db_原価 = Db_原価 + Fn_数(Va_工程(Lg_行, 列_工程表.加工時間)) / 60 * _
                        Fn_辞書値(Ob_賃率, Fn_文字(Va_工程(Lg_行, 列_工程表.作業区)))
                End If
            End If
        Next Lg_行
        ' 子品目の原価×員数を積み上げる
        For Lg_行 = 1 To Fn_行数(Va_BOM)
            If Fn_文字(Va_BOM(Lg_行, 列_BOM.事業部)) = St_事業部 And Fn_文字(Va_BOM(Lg_行, 列_BOM.親品番)) = St_品番 Then
                Db_原価 = Db_原価 + Fn_数(Va_BOM(Lg_行, 列_BOM.員数)) * _
                    Fn_品目原価(St_事業部, Fn_文字(Va_BOM(Lg_行, 列_BOM.子品番)), Va_品目, Va_BOM, Va_工程, Ob_賃率, Ob_結果, Lg_段数 + 1)
            End If
        Next Lg_行
    End If
    Ob_結果.Item(St_品番) = Db_原価
    Fn_品目原価 = Db_原価
End Function

' 品目シートの計算列（標準原価・適用販売単価）を書き直す（両事業部とも）
'   マスタを直接編集したあとでも、マクロを実行すれば最新の値になるようにするため。
Public Sub Pr_品目計算列更新()
    Dim Ws_品目 As Worksheet
    Dim Va_表 As Variant
    Dim Va_出力 As Variant
    Dim Ob_加工 As Object
    Dim Ob_組立 As Object
    Dim Lg_行 As Long
    Dim Db_原価 As Double
    Dim St_事業部 As String
    Dim St_品番 As String
    Set Ws_品目 = ThisWorkbook.Worksheets(St_シート_品目)
    Va_表 = Fn_表読込(Ws_品目, 列_品目.次列 - 1)
    If Fn_行数(Va_表) = 0 Then Exit Sub
    Set Ob_加工 = Fn_標準原価集計(St_加工)
    Set Ob_組立 = Fn_標準原価集計(St_組立)
    ReDim Va_出力(1 To Fn_行数(Va_表), 1 To 2)
    For Lg_行 = 1 To Fn_行数(Va_表)
        St_事業部 = Fn_文字(Va_表(Lg_行, 列_品目.事業部))
        St_品番 = Fn_文字(Va_表(Lg_行, 列_品目.品番))
        If St_事業部 = St_加工 Then
            Db_原価 = Fn_丸め(Fn_辞書値(Ob_加工, St_品番))
        Else
            Db_原価 = Fn_丸め(Fn_辞書値(Ob_組立, St_品番))
        End If
        Va_出力(Lg_行, 1) = Db_原価
        ' 販売単価は製造品だけ（購入品は売らない）
        If Fn_製造品か(Fn_文字(Va_表(Lg_行, 列_品目.区分))) Then
            Va_出力(Lg_行, 2) = Fn_販売単価(Va_表(Lg_行, 列_品目.販売単価), Db_原価)
        Else
            Va_出力(Lg_行, 2) = ""
        End If
    Next Lg_行
    Ws_品目.Range(Ws_品目.Cells(Lg_開始行, 列_品目.標準原価), _
                  Ws_品目.Cells(Lg_開始行 + Fn_行数(Va_表) - 1, 列_品目.適用販売単価)).Value = Va_出力
End Sub
