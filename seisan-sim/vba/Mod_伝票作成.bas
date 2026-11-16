Attribute VB_Name = "Mod_伝票作成"
Option Explicit
'==============================================================
' Mod_伝票作成
' 【目的】複数の業務から使う「伝票を作る」共通部品を置く。
'         ・製造指図の作成（工程の複写と手配No.の採番を含む）
'         ・発注の作成
'         ・製造指図・ピッキングの削除（取消のとき）
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
