Attribute VB_Name = "Mod_在庫計算"
Option Explicit
'==============================================================
' Mod_在庫計算（在庫と日程の計算）
' 【目的】在庫の「引当」「有効在庫」「手配済」を計算する共通部品。
'         在庫一覧・MRP・各業務のチェックから呼ばれる。
' 【業務背景】HTML 版の次の式と同じにする。
'   ・受注引当   ＝ 状態が「引当済」の受注の残数（数量－出荷済）の合計
'   ・製造用確保 ＝ 完了していない製造用ピッキングの明細数量の合計
'   ・引当       ＝ 受注引当＋製造用確保
'   ・有効在庫   ＝ 在庫－引当（自由に使える数）
'   ・手配済     ＝ 計画・着手の製造指図の見込み出来高＋発注済の発注数
'   ・見込み出来高＝ 指示数。着手後は、完了した最後の工程の良品数
'                   （途中で不良が出たら、その分だけ出来高が減る）
' 【前提】戻り値の Dictionary は「品番 → 数量」。無い品番は 0 とみなす
'         （呼ぶ側は Fn_辞書値 で読む）。
'         Dictionary には数値・文字だけを入れる（オブジェクトは入れない）。
'==============================================================

' Dictionary の値を数で読む（キーが無ければ 0）
'   引数：Ob_辞書、St_キー　戻り値：数
Public Function Fn_辞書値(ByVal Ob_辞書 As Object, ByVal St_キー As String) As Double
    If Ob_辞書.Exists(St_キー) Then
        Fn_辞書値 = Ob_辞書.Item(St_キー)
    Else
        Fn_辞書値 = 0
    End If
End Function

' Dictionary の値に数を足す（キーが無ければ作る）
Public Sub Pr_辞書加算(ByVal Ob_辞書 As Object, ByVal St_キー As String, ByVal Db_数 As Double)
    If Ob_辞書.Exists(St_キー) Then
        Ob_辞書.Item(St_キー) = Fn_丸め(Ob_辞書.Item(St_キー) + Db_数)
    Else
        Call Ob_辞書.Add(St_キー, Fn_丸め(Db_数))
    End If
End Sub

' 受注引当：品番ごとの「引当済の受注の残数」の合計
'   引数：St_事業部　戻り値：Dictionary（品番 → 数量）
Public Function Fn_受注引当集計(ByVal St_事業部 As String) As Object
    Dim Ob_結果 As Object
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Set Ob_結果 = CreateObject("Scripting.Dictionary")
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_受注), 列_受注.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_受注.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_受注.状態)) = St_状態_引当済 Then
            Call Pr_辞書加算(Ob_結果, Fn_文字(Va_表(Lg_行, 列_受注.品番)), _
                Fn_数(Va_表(Lg_行, 列_受注.数量)) - Fn_数(Va_表(Lg_行, 列_受注.出荷済)))
        End If
    Next Lg_行
    Set Fn_受注引当集計 = Ob_結果
End Function

' 製造用確保：完了していない製造用ピッキングの明細を品番ごとに合計
'   引数：St_事業部　戻り値：Dictionary（品番 → 数量）
'   製造用ピッキングを出した時点で部材を確保する（ほかの指図に回さない）ため。
Public Function Fn_製造用確保集計(ByVal St_事業部 As String) As Object
    Dim Ob_結果 As Object
    Dim Ob_対象 As Object
    Dim Va_頭 As Variant
    Dim Va_明細 As Variant
    Dim Lg_行 As Long
    Set Ob_結果 = CreateObject("Scripting.Dictionary")
    Set Ob_対象 = CreateObject("Scripting.Dictionary")
    ' 完了していない製造用ピッキングの番号を集める
    Va_頭 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_ピッキング), 列_ピッキング.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_頭)
        If Fn_文字(Va_頭(Lg_行, 列_ピッキング.事業部)) = St_事業部 _
           And Fn_文字(Va_頭(Lg_行, 列_ピッキング.種別)) = "製造" _
           And Fn_文字(Va_頭(Lg_行, 列_ピッキング.状態)) <> St_状態_完了 Then
            Ob_対象.Item(Fn_文字(Va_頭(Lg_行, 列_ピッキング.ピッキングNo))) = 1
        End If
    Next Lg_行
    ' その明細を品番ごとに足す
    Va_明細 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_ピッキング明細), 列_ピッキング明細.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_明細)
        If Fn_文字(Va_明細(Lg_行, 列_ピッキング明細.事業部)) = St_事業部 Then
            If Ob_対象.Exists(Fn_文字(Va_明細(Lg_行, 列_ピッキング明細.ピッキングNo))) Then
                Call Pr_辞書加算(Ob_結果, Fn_文字(Va_明細(Lg_行, 列_ピッキング明細.品番)), _
                    Fn_数(Va_明細(Lg_行, 列_ピッキング明細.数量)))
            End If
        End If
    Next Lg_行
    Set Fn_製造用確保集計 = Ob_結果
End Function

' 製造指図の見込み出来高（指図No → 数量）
'   引数：St_事業部　戻り値：Dictionary（指図No → 見込み出来高）
'   計画中は指示数。着手後は、完了した工程のうち最後のものの良品数。
'   工程は順に完了していくので「完了した中で工程順が最大のもの」を使う。
Public Function Fn_指図出来高集計(ByVal St_事業部 As String) As Object
    Dim Ob_結果 As Object
    Dim Ob_最終工程 As Object
    Dim Va_指図 As Variant
    Dim Va_工程 As Variant
    Dim Lg_行 As Long
    Dim St_指図 As String
    Set Ob_結果 = CreateObject("Scripting.Dictionary")
    Set Ob_最終工程 = CreateObject("Scripting.Dictionary")
    ' まず指示数を入れる
    Va_指図 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_製造指図), 列_製造指図.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_指図)
        If Fn_文字(Va_指図(Lg_行, 列_製造指図.事業部)) = St_事業部 Then
            Ob_結果.Item(Fn_文字(Va_指図(Lg_行, 列_製造指図.指図No))) = Fn_数(Va_指図(Lg_行, 列_製造指図.数量))
        End If
    Next Lg_行
    ' 完了した工程があれば、工程順が最大のものの良品数で置き換える（計画中の指図は置き換えない）
    Va_工程 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_指図工程), 列_指図工程.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_工程)
        If Fn_文字(Va_工程(Lg_行, 列_指図工程.事業部)) = St_事業部 _
           And Fn_文字(Va_工程(Lg_行, 列_指図工程.状態)) = St_状態_完了 Then
            St_指図 = Fn_文字(Va_工程(Lg_行, 列_指図工程.指図No))
            If Fn_辞書値(Ob_最終工程, St_指図) < Fn_数(Va_工程(Lg_行, 列_指図工程.工程順)) Then
                Ob_最終工程.Item(St_指図) = Fn_数(Va_工程(Lg_行, 列_指図工程.工程順))
                Ob_結果.Item(St_指図) = Fn_数(Va_工程(Lg_行, 列_指図工程.良品))
            End If
        End If
    Next Lg_行
    ' 計画中の指図は指示数に戻す（HTML 版の moOut は計画中なら工程を見ない）
    For Lg_行 = 1 To Fn_行数(Va_指図)
        If Fn_文字(Va_指図(Lg_行, 列_製造指図.事業部)) = St_事業部 _
           And Fn_文字(Va_指図(Lg_行, 列_製造指図.状態)) = St_状態_計画 Then
            Ob_結果.Item(Fn_文字(Va_指図(Lg_行, 列_製造指図.指図No))) = Fn_数(Va_指図(Lg_行, 列_製造指図.数量))
        End If
    Next Lg_行
    Set Fn_指図出来高集計 = Ob_結果
End Function

' 手配済：品番ごとの「製造中の見込み出来高＋発注残」
'   引数：St_事業部　戻り値：Dictionary（品番 → 数量）
'   MRP で「もう手配してある分」として不足から差し引くため。
Public Function Fn_手配済集計(ByVal St_事業部 As String) As Object
    Dim Ob_結果 As Object
    Dim Ob_出来高 As Object
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim St_状態 As String
    Set Ob_結果 = CreateObject("Scripting.Dictionary")
    Set Ob_出来高 = Fn_指図出来高集計(St_事業部)
    ' 計画・着手の製造指図
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_製造指図), 列_製造指図.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        St_状態 = Fn_文字(Va_表(Lg_行, 列_製造指図.状態))
        If Fn_文字(Va_表(Lg_行, 列_製造指図.事業部)) = St_事業部 And (St_状態 = St_状態_計画 Or St_状態 = St_状態_着手) Then
            Call Pr_辞書加算(Ob_結果, Fn_文字(Va_表(Lg_行, 列_製造指図.品番)), _
                Fn_辞書値(Ob_出来高, Fn_文字(Va_表(Lg_行, 列_製造指図.指図No))))
        End If
    Next Lg_行
    ' 発注済の発注
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_発注), 列_発注.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_発注.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_発注.状態)) = St_状態_発注済 Then
            Call Pr_辞書加算(Ob_結果, Fn_文字(Va_表(Lg_行, 列_発注.品番)), Fn_数(Va_表(Lg_行, 列_発注.数量)))
        End If
    Next Lg_行
    Set Fn_手配済集計 = Ob_結果
End Function

' 1品目の有効在庫（在庫－受注引当－製造用確保）
'   引数：St_事業部、St_品番　戻り値：有効在庫
'   受注の引当や、ピッキング指示を出せるかの判定に使う。
Public Function Fn_有効在庫(ByVal St_事業部 As String, ByVal St_品番 As String) As Double
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_品番)
    If IsEmpty(Va_品目) Then
        Fn_有効在庫 = 0
        Exit Function
    End If
    Fn_有効在庫 = Fn_丸め(Fn_数(Va_品目(列_品目.在庫数)) _
        - Fn_辞書値(Fn_受注引当集計(St_事業部), St_品番) - Fn_辞書値(Fn_製造用確保集計(St_事業部), St_品番))
End Function

' 受注に紐付いた、まだ完成していない製造指図の見込み出来高の合計（HTML 版の pegOut）
'   引数：St_事業部、St_受注No、St_品番　戻り値：数量
'   受注生産で「この受注向けにもう作っている数」を数えるため。
Public Function Fn_紐付出来高(ByVal St_事業部 As String, ByVal St_受注No As String, ByVal St_品番 As String) As Double
    Dim Ob_出来高 As Object
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim St_状態 As String
    Dim Db_計 As Double
    Set Ob_出来高 = Fn_指図出来高集計(St_事業部)
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_製造指図), 列_製造指図.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        St_状態 = Fn_文字(Va_表(Lg_行, 列_製造指図.状態))
        If Fn_文字(Va_表(Lg_行, 列_製造指図.事業部)) = St_事業部 _
           And Fn_文字(Va_表(Lg_行, 列_製造指図.品番)) = St_品番 _
           And Fn_文字(Va_表(Lg_行, 列_製造指図.受注No)) = St_受注No _
           And (St_状態 = St_状態_計画 Or St_状態 = St_状態_着手) Then
            Db_計 = Db_計 + Fn_辞書値(Ob_出来高, Fn_文字(Va_表(Lg_行, 列_製造指図.指図No)))
        End If
    Next Lg_行
    Fn_紐付出来高 = Fn_丸め(Db_計)
End Function

' 受注に対して出荷用ピッキング中の数量（HTML 版の pkOpen）
'   引数：St_事業部、St_受注No　戻り値：数量
'   ピッキング中の受注は引当を解除できない、などの判定に使う。
Public Function Fn_出荷ピッキング中(ByVal St_事業部 As String, ByVal St_受注No As String) As Double
    Dim Ob_対象 As Object
    Dim Va_頭 As Variant
    Dim Va_明細 As Variant
    Dim Lg_行 As Long
    Dim Db_計 As Double
    Set Ob_対象 = CreateObject("Scripting.Dictionary")
    Va_頭 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_ピッキング), 列_ピッキング.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_頭)
        If Fn_文字(Va_頭(Lg_行, 列_ピッキング.事業部)) = St_事業部 _
           And Fn_文字(Va_頭(Lg_行, 列_ピッキング.種別)) = "出荷" _
           And Fn_文字(Va_頭(Lg_行, 列_ピッキング.状態)) <> St_状態_完了 _
           And Fn_文字(Va_頭(Lg_行, 列_ピッキング.対象伝票)) = St_受注No Then
            Ob_対象.Item(Fn_文字(Va_頭(Lg_行, 列_ピッキング.ピッキングNo))) = 1
        End If
    Next Lg_行
    ' 出荷用ピッキングは明細が1行（その受注の製品）なので、1行目の数量を足す
    Va_明細 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_ピッキング明細), 列_ピッキング明細.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_明細)
        If Fn_文字(Va_明細(Lg_行, 列_ピッキング明細.事業部)) = St_事業部 _
           And Fn_数(Va_明細(Lg_行, 列_ピッキング明細.行No)) = 1 Then
            If Ob_対象.Exists(Fn_文字(Va_明細(Lg_行, 列_ピッキング明細.ピッキングNo))) Then
                Db_計 = Db_計 + Fn_数(Va_明細(Lg_行, 列_ピッキング明細.数量))
            End If
        End If
    Next Lg_行
    Fn_出荷ピッキング中 = Fn_丸め(Db_計)
End Function

' 所要日数（HTML 版の prodDays）
'   引数：St_事業部、St_品番、Db_数量　戻り値：日数
'   社内工程の（段取＋加工×数量）を作業区の1日の能力で割って合計し、切り上げる（最低1日）。
'   外注工程は社内の時間を使わず、外注リードタイムの日数をそのまま足す。
Public Function Fn_所要日数(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal Db_数量 As Double) As Long
    Dim Va_工程 As Variant
    Dim Ob_能力 As Object
    Dim Va_作業区 As Variant
    Dim Lg_行 As Long
    Dim Db_日 As Double
    Dim Lg_外注日 As Long
    Dim Db_能力 As Double
    Dim Db_分 As Double
    ' 作業区の能力（分/日）
    Set Ob_能力 = CreateObject("Scripting.Dictionary")
    Va_作業区 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_作業区), 列_作業区.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_作業区)
        If Fn_文字(Va_作業区(Lg_行, 列_作業区.事業部)) = St_事業部 Then
            Ob_能力.Item(Fn_文字(Va_作業区(Lg_行, 列_作業区.作業区))) = Fn_数(Va_作業区(Lg_行, 列_作業区.能力))
        End If
    Next Lg_行
    Va_工程 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_工程表), 列_工程表.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_工程)
        If Fn_文字(Va_工程(Lg_行, 列_工程表.事業部)) = St_事業部 And Fn_文字(Va_工程(Lg_行, 列_工程表.品番)) = St_品番 Then
            If Fn_文字(Va_工程(Lg_行, 列_工程表.外注)) = St_印 Then
                Lg_外注日 = Lg_外注日 + CLng(Fn_数(Va_工程(Lg_行, 列_工程表.外注LT)))
            Else
                Db_能力 = Fn_辞書値(Ob_能力, Fn_文字(Va_工程(Lg_行, 列_工程表.作業区)))
                If Db_能力 <= 0 Then Db_能力 = 450      ' 能力が未設定なら HTML 版の既定（450分/日）
                Db_分 = Fn_丸め(Fn_数(Va_工程(Lg_行, 列_工程表.段取時間)) + Fn_数(Va_工程(Lg_行, 列_工程表.加工時間)) * Db_数量)
                Db_日 = Db_日 + Db_分 / Db_能力
            End If
        End If
    Next Lg_行
    Fn_所要日数 = CLng(Fn_切上げ(Db_日))
    If Fn_所要日数 < 1 Then Fn_所要日数 = 1
    Fn_所要日数 = Fn_所要日数 + Lg_外注日
End Function
