Attribute VB_Name = "Mod_伝票作成"
Option Explicit
'==============================================================
' Mod_伝票作成
' 【目的】複数の業務から使う「伝票を作る」共通部品を置く。
'         ・製造指図の作成（工程の複写と手配No.の採番を含む）
'         ・発注の作成
'         ・製造指図・ピッキングの削除（取消のとき）
'         ・在庫の移動と履歴、ピッキングの作成、工程完了・自動引当、出荷の計上、見込の消し込み
' 【業務背景】製造指図は、作った時点の工程表を複写して持つ（HTML 版と同じ）。
'         あとで工程表を変えても、作成済みの製造指図には反映しない。
'         複写するとき、工程1つごとに手配No.（全社連番）を振る。
'==============================================================

' 品目の工程表を、工程順に並べて指図工程シートへ複写する（工程ごとに手配No.を振る）
'   引数：St_事業部、St_指図No、St_品番
'   戻り値なし。複写した工程の状態は「未」、良品・不良・作業時間は空。
Public Sub Pr_指図工程複写(ByVal St_事業部 As String, ByVal St_指図No As String, ByVal St_品番 As String)
    Dim Ws_指図工程 As Worksheet
    Dim Va_工程 As Variant
    Dim Lg_対象行() As Long
    Dim Lg_行 As Long
    Dim Lg_件数 As Long
    Dim Lg_添字 As Long
    Dim Lg_比較 As Long
    Dim Lg_一時 As Long
    Set Ws_指図工程 = ThisWorkbook.Worksheets(St_シート_指図工程)
    Va_工程 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_工程表), 列_工程表.次列 - 1)
    ' この品目の工程の行番号を集める
    Lg_件数 = 0
    For Lg_行 = 1 To Fn_行数(Va_工程)
        If Fn_文字(Va_工程(Lg_行, 列_工程表.事業部)) = St_事業部 And Fn_文字(Va_工程(Lg_行, 列_工程表.品番)) = St_品番 Then
            Lg_件数 = Lg_件数 + 1
            ReDim Preserve Lg_対象行(1 To Lg_件数)
            Lg_対象行(Lg_件数) = Lg_行
        End If
    Next Lg_行
    If Lg_件数 = 0 Then Exit Sub
    ' 工程順に並べる（件数が少ないので単純な並べ替えでよい）
    For Lg_添字 = 1 To Lg_件数 - 1
        For Lg_比較 = Lg_添字 + 1 To Lg_件数
            If Fn_数(Va_工程(Lg_対象行(Lg_比較), 列_工程表.工程順)) < Fn_数(Va_工程(Lg_対象行(Lg_添字), 列_工程表.工程順)) Then
                Lg_一時 = Lg_対象行(Lg_添字)
                Lg_対象行(Lg_添字) = Lg_対象行(Lg_比較)
                Lg_対象行(Lg_比較) = Lg_一時
            End If
        Next Lg_比較
    Next Lg_添字
    ' 1工程ずつ、手配No.を振って追記する（並びは 列_指図工程 と同じ）
    For Lg_添字 = 1 To Lg_件数
        Lg_行 = Lg_対象行(Lg_添字)
        Call Fn_行追記(Ws_指図工程, Array(St_事業部, St_指図No, _
            Va_工程(Lg_行, 列_工程表.工程順), Fn_手配No採番(), _
            Va_工程(Lg_行, 列_工程表.工程名), Va_工程(Lg_行, 列_工程表.作業区), _
            Va_工程(Lg_行, 列_工程表.段取時間), Va_工程(Lg_行, 列_工程表.加工時間), _
            Va_工程(Lg_行, 列_工程表.外注), Va_工程(Lg_行, 列_工程表.外注先), _
            Va_工程(Lg_行, 列_工程表.外注単価), Va_工程(Lg_行, 列_工程表.外注LT), _
            St_状態_未, "", "", ""))
    Next Lg_添字
End Sub

' 製造指図を1件作る（HTML 版の newMO）
'   引数：St_事業部、St_品番、Db_数量、Lg_着手日（今日から何日後）、Lg_完成日（同）、
'         St_受注No（受注生産で紐付ける受注。無ければ空）、Dt_今日
'   戻り値：作った指図No
'   状態は「計画」。工程は作成時点の工程表を複写し、工程ごとに手配No.を振る。
Public Function Fn_製造指図作成(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal Db_数量 As Double, _
        ByVal Lg_着手日 As Long, ByVal Lg_完成日 As Long, ByVal St_受注No As String, ByVal Dt_今日 As Date) As String
    Dim St_指図No As String
    St_指図No = Fn_採番(St_事業部, "MO")
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_製造指図), Array(St_事業部, St_指図No, St_品番, Db_数量, _
        St_状態_計画, Dt_今日 + Lg_着手日, Dt_今日 + Lg_完成日, St_受注No, "", "", ""))
    Call Pr_指図工程複写(St_事業部, St_指図No, St_品番)
    Fn_製造指図作成 = St_指図No
End Function

' 発注を1件作る（HTML 版の newPO）
'   引数：St_事業部、St_品番、Db_数量、Dt_今日
'   戻り値：作った発注No
'   入荷予定は 今日＋品目のリードタイム。状態は「発注済」。
Public Function Fn_発注作成(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal Db_数量 As Double, _
        ByVal Dt_今日 As Date) As String
    Dim St_発注No As String
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_品番)
    St_発注No = Fn_採番(St_事業部, "PO")
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_発注), Array(St_事業部, St_発注No, St_品番, Db_数量, _
        Dt_今日 + Fn_数(Va_品目(列_品目.LT)), St_状態_発注済))
    Fn_発注作成 = St_発注No
End Function

' 製造指図を消す（ヘッダと工程の行。取消のとき）
'   引数：St_事業部、St_指図No
Public Sub Pr_製造指図削除(ByVal St_事業部 As String, ByVal St_指図No As String)
    Call Pr_キー一致行削除(St_シート_指図工程, 列_指図工程.次列 - 1, St_事業部, 列_指図工程.指図No, St_指図No)
    Call Pr_キー一致行削除(St_シート_製造指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_指図No)
End Sub

' ピッキング指示を消す（ヘッダと明細の行。取消のとき）
'   引数：St_事業部、St_ピッキングNo
Public Sub Pr_ピッキング削除(ByVal St_事業部 As String, ByVal St_ピッキングNo As String)
    If St_ピッキングNo = "" Then Exit Sub
    Call Pr_キー一致行削除(St_シート_ピッキング明細, 列_ピッキング明細.次列 - 1, St_事業部, 列_ピッキング明細.ピッキングNo, St_ピッキングNo)
    Call Pr_キー一致行削除(St_シート_ピッキング, 列_ピッキング.次列 - 1, St_事業部, 列_ピッキング.ピッキングNo, St_ピッキングNo)
End Sub

' 事業部とキーが一致する行をすべて消す
Private Sub Pr_キー一致行削除(ByVal St_シート名 As String, ByVal Lg_列数 As Long, ByVal St_事業部 As String, _
        ByVal Lg_キー列 As Long, ByVal St_キー As String)
    Dim Ws_表 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート名)
    Va_表 = Fn_表読込(Ws_表, Lg_列数)
    ' 下の行から消す（上から消すと行番号がずれるため）
    For Lg_行 = Fn_行数(Va_表) To 1 Step -1
        If Fn_文字(Va_表(Lg_行, 1)) = St_事業部 And Fn_文字(Va_表(Lg_行, Lg_キー列)) = St_キー Then
            Ws_表.Rows(Lg_行 + Lg_開始行 - 1).Delete
        End If
    Next Lg_行
End Sub

'==============================================================
' 段階3：在庫の移動・ピッキング・工程の完了・出荷
'==============================================================

' 在庫を動かし、履歴に1行残す（HTML 版の tx）
'   引数：St_事業部、St_品番、St_種別（払出・完成・入荷・出荷・出荷取消・調整）、Db_増減、St_伝票
'   在庫数は品目シートの在庫数の列。在庫を変えるのは必ずこの部品を通す（履歴を漏らさないため）。
Public Sub Pr_在庫移動(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal St_種別 As String, _
        ByVal Db_増減 As Double, ByVal St_伝票 As String)
    Dim Ws_品目 As Worksheet
    Dim Lg_行 As Long
    Dim Db_残 As Double
    Set Ws_品目 = ThisWorkbook.Worksheets(St_シート_品目)
    Lg_行 = Fn_行検索(Ws_品目, 列_品目.次列 - 1, St_事業部, 列_品目.品番, St_品番)
    If Lg_行 = 0 Then Err.Raise Lg_中断番号 + 1, , "品目 " & St_品番 & " が品目シートにありません"
    Db_残 = Fn_丸め(Fn_数(Ws_品目.Cells(Lg_行, 列_品目.在庫数).Value) + Db_増減)
    Ws_品目.Cells(Lg_行, 列_品目.在庫数).Value = Db_残
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_履歴), Array(St_事業部, Now, St_種別, St_品番, Fn_丸め(Db_増減), Db_残, St_伝票))
End Sub

' ピッキング指示を作る（HTML 版の newPK）
'   引数：St_事業部、St_種別（製造／出荷）、St_対象伝票、St_内容、Va_品番（配列）、Va_数量（配列）、St_配送方法
'   戻り値：ピッキングNo。明細の棚番は、いまの品目マスタの棚番。済は空（未）。
Public Function Fn_ピッキング作成(ByVal St_事業部 As String, ByVal St_種別 As String, ByVal St_対象伝票 As String, _
        ByVal St_内容 As String, ByVal Va_品番 As Variant, ByVal Va_数量 As Variant, ByVal St_配送方法 As String) As String
    Dim St_番号 As String
    Dim Lg_添字 As Long
    Dim Va_品目 As Variant
    St_番号 = Fn_採番(St_事業部, "PK")
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング), Array(St_事業部, St_番号, St_種別, St_対象伝票, St_内容, _
        St_状態_作業中, St_配送方法, ""))
    For Lg_添字 = LBound(Va_品番) To UBound(Va_品番)
        Va_品目 = Fn_品目情報(St_事業部, CStr(Va_品番(Lg_添字)))
        Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング明細), Array(St_事業部, St_番号, Lg_添字 - LBound(Va_品番) + 1, _
            Va_品番(Lg_添字), Va_数量(Lg_添字), Va_品目(列_品目.棚番), ""))
    Next Lg_添字
    Fn_ピッキング作成 = St_番号
End Function

' 工程が終わったときの共通処理（HTML 版の opFinish）
'   引数：St_事業部、St_指図No、Bl_最終工程、Db_良品
'   最終工程か、良品が0なら、製造指図を完成にして良品を製品の在庫に入れる。
'   受注生産で受注に紐付いていれば、その受注に自動で引き当てる。
Public Sub Pr_工程完了後処理(ByVal St_事業部 As String, ByVal St_指図No As String, ByVal Bl_最終工程 As Boolean, ByVal Db_良品 As Double)
    Dim Ws_指図 As Worksheet
    Dim Lg_行 As Long
    If Not Bl_最終工程 And Db_良品 <> 0 Then Exit Sub
    Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
    Lg_行 = Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_指図No)
    Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value = St_状態_完成
    Ws_指図.Cells(Lg_行, 列_製造指図.出来高).Value = Db_良品
    If Db_良品 > 0 Then
        Call Pr_在庫移動(St_事業部, Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.品番).Value), "完成", Db_良品, St_指図No)
        Call Pr_受注自動引当(St_事業部, Lg_行)
    End If
End Sub

' 受注に紐付いた製造指図が完成したら、その受注に自動で引き当てる（HTML 版の pegDone）
'   引数：St_事業部、Lg_指図行（製造指図シートの行）
'   受注が「受付」で、有効在庫が受注残以上なら「引当済」にする。足りなければ保留。結果は指図の備考に書く。
Public Sub Pr_受注自動引当(ByVal St_事業部 As String, ByVal Lg_指図行 As Long)
    Dim Ws_指図 As Worksheet
    Dim Ws_受注 As Worksheet
    Dim St_受注No As String
    Dim Lg_受注行 As Long
    Dim Db_残 As Double
    Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
    Set Ws_受注 = ThisWorkbook.Worksheets(St_シート_受注)
    St_受注No = Fn_文字(Ws_指図.Cells(Lg_指図行, 列_製造指図.受注No).Value)
    If St_受注No = "" Then Exit Sub
    Lg_受注行 = Fn_行検索(Ws_受注, 列_受注.次列 - 1, St_事業部, 列_受注.受注No, St_受注No)
    If Lg_受注行 = 0 Then Exit Sub
    If Fn_文字(Ws_受注.Cells(Lg_受注行, 列_受注.状態).Value) <> St_状態_受付 Then Exit Sub
    Db_残 = Fn_丸め(Fn_数(Ws_受注.Cells(Lg_受注行, 列_受注.数量).Value) - Fn_数(Ws_受注.Cells(Lg_受注行, 列_受注.出荷済).Value))
    If Fn_有効在庫(St_事業部, Fn_文字(Ws_指図.Cells(Lg_指図行, 列_製造指図.品番).Value)) >= Db_残 Then
        Ws_受注.Cells(Lg_受注行, 列_受注.状態).Value = St_状態_引当済
        Ws_指図.Cells(Lg_指図行, 列_製造指図.備考).Value = "受注 " & St_受注No & " に自動で引き当てました。"
    Else
        Ws_指図.Cells(Lg_指図行, 列_製造指図.備考).Value = "受注 " & St_受注No & " の必要数に足りないため、引当は保留です。"
    End If
End Sub

' 出荷を計上する（HTML 版の doShip）
'   引数：St_事業部、Lg_受注行（受注シートの行）、Db_数量、St_配送方法、Dt_今日
'   戻り値：出荷No
'   在庫を減らし、見込生産なら見込を消し込み、受注の出荷済を増やす（残0で出荷済）。
'   出荷実績には、その時点の販売単価を固定して残す（あとで単価を変えても請求額が変わらないように）。
Public Function Fn_出荷計上(ByVal St_事業部 As String, ByVal Lg_受注行 As Long, ByVal Db_数量 As Double, _
        ByVal St_配送方法 As String, ByVal Dt_今日 As Date) As String
    Dim Ws_受注 As Worksheet
    Dim St_出荷No As String
    Dim St_品番 As String
    Dim Va_品目 As Variant
    Dim Db_出荷済 As Double
    Dim Db_単価 As Double
    Set Ws_受注 = ThisWorkbook.Worksheets(St_シート_受注)
    St_品番 = Fn_文字(Ws_受注.Cells(Lg_受注行, 列_受注.品番).Value)
    Va_品目 = Fn_品目情報(St_事業部, St_品番)
    St_出荷No = Fn_採番(St_事業部, "SH")
    Call Pr_在庫移動(St_事業部, St_品番, "出荷", -Db_数量, St_出荷No)
    If Fn_文字(Va_品目(列_品目.生産方式)) <> St_受注生産 Then Call Pr_見込消込(St_事業部, St_品番, Db_数量)
    Db_出荷済 = Fn_丸め(Fn_数(Ws_受注.Cells(Lg_受注行, 列_受注.出荷済).Value) + Db_数量)
    Ws_受注.Cells(Lg_受注行, 列_受注.出荷済).Value = Db_出荷済
    If Fn_丸め(Fn_数(Ws_受注.Cells(Lg_受注行, 列_受注.数量).Value) - Db_出荷済) <= 0 Then
        Ws_受注.Cells(Lg_受注行, 列_受注.状態).Value = St_状態_出荷済
    End If
    Db_単価 = Fn_販売単価(Va_品目(列_品目.販売単価), Fn_丸め(Fn_辞書値(Fn_標準原価集計(St_事業部), St_品番)))
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_出荷実績), Array(St_事業部, St_出荷No, _
        Ws_受注.Cells(Lg_受注行, 列_受注.受注No).Value, St_品番, Db_数量, Ws_受注.Cells(Lg_受注行, 列_受注.得意先).Value, _
        Dt_今日, St_配送方法, Db_単価))
    Fn_出荷計上 = St_出荷No
End Function

' 見込の消し込み：出荷した分を、期日の早い見込から差し引く。0になった見込は消す（HTML 版の consumeFc）
'   引数：St_事業部、St_品番、Db_数量
Public Sub Pr_見込消込(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal Db_数量 As Double)
    Dim Ws_見込 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_順() As Long
    Dim Lg_件数 As Long
    Dim Lg_行 As Long
    Dim Lg_添字 As Long
    Dim Lg_一時 As Long
    Dim Db_使う As Double
    Dim Db_残 As Double
    Set Ws_見込 = ThisWorkbook.Worksheets(St_シート_見込)
    Va_表 = Fn_表読込(Ws_見込, 列_見込.次列 - 1)
    ' この品目の見込を集め、期日の早い順に（同じ期日は HTML 版の並び＝新しい行が先）
    For Lg_行 = Fn_行数(Va_表) To 1 Step -1
        If Fn_文字(Va_表(Lg_行, 列_見込.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_見込.品番)) = St_品番 Then
            Lg_件数 = Lg_件数 + 1
            ReDim Preserve Lg_順(1 To Lg_件数)
            Lg_順(Lg_件数) = Lg_行
        End If
    Next Lg_行
    If Lg_件数 = 0 Then Exit Sub
    For Lg_行 = 2 To Lg_件数
        Lg_一時 = Lg_順(Lg_行)
        Lg_添字 = Lg_行 - 1
        Do While Lg_添字 >= 1
            If CDate(Va_表(Lg_順(Lg_添字), 列_見込.期日)) <= CDate(Va_表(Lg_一時, 列_見込.期日)) Then Exit Do
            Lg_順(Lg_添字 + 1) = Lg_順(Lg_添字)
            Lg_添字 = Lg_添字 - 1
        Loop
        Lg_順(Lg_添字 + 1) = Lg_一時
    Next Lg_行
    ' 早い順に差し引く
    Db_残 = Db_数量
    For Lg_添字 = 1 To Lg_件数
        Lg_行 = Lg_順(Lg_添字)
        Db_使う = Fn_数(Va_表(Lg_行, 列_見込.数量))
        If Db_残 < Db_使う Then Db_使う = Db_残
        Ws_見込.Cells(Lg_行 + Lg_開始行 - 1, 列_見込.数量).Value = Fn_丸め(Fn_数(Va_表(Lg_行, 列_見込.数量)) - Db_使う)
        Db_残 = Fn_丸め(Db_残 - Db_使う)
    Next Lg_添字
    ' 0 になった見込を消す（下の行から）
    Va_表 = Fn_表読込(Ws_見込, 列_見込.次列 - 1)
    For Lg_行 = Fn_行数(Va_表) To 1 Step -1
        If Fn_文字(Va_表(Lg_行, 列_見込.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_見込.品番)) = St_品番 _
           And Fn_数(Va_表(Lg_行, 列_見込.数量)) <= 0 Then
            Ws_見込.Rows(Lg_行 + Lg_開始行 - 1).Delete
        End If
    Next Lg_行
End Sub

' 数を HTML 版と同じ見た目の文字にする（桁区切り、小数は2桁まで）
Public Function Fn_表示数(ByVal Db_値 As Double) As String
    Dim St_文 As String
    St_文 = Format$(Fn_丸め(Db_値), "#,##0.##")
    If Right$(St_文, 1) = "." Then St_文 = Left$(St_文, Len(St_文) - 1)
    Fn_表示数 = St_文
End Function
