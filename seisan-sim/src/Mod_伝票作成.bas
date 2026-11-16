Attribute VB_Name = "Mod_伝票作成"
Option Explicit
'==============================================================
' Mod_伝票作成
' 【目的】複数の業務から使う「伝票を作る」共通部品を置く。
'         段階1では、製造指図の工程を工程表から複写する部品だけ。
'         （段階2以降で、製造指図・発注の作成などをここに足す）
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
