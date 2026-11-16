Attribute VB_Name = "Mod_マスタ点検"
Option Explicit
'==============================================================
' Mod_マスタ点検
' 【目的】直接編集されたマスタ（品目・BOM・作業区・工程表）に、計算を
'         狂わせる誤りが無いかを確かめる。誤りがあれば一覧を見せて中断する。
' 【業務背景】HTML 版は入力画面で誤りを弾くが、Excel 版はマスタを
'         シートで直接編集するので、MRP や製造指図の作成の前に点検する。
' 【点検すること】（現在の事業部の行だけ）
'   ・品目：品番の重複・空、区分、製造品の生産方式、数値の列
'   ・BOM ：親・子が品目にある、親が製造品、員数が正の数、循環が無い
'   ・作業区：コードの重複、賃率・能力が数
'   ・工程表：品目が製造品、工程順が数、社内工程は作業区があり時間が数、
'            外注工程は外注先・単価・LT がある
'   ・製造品なのに工程が1つも無い
' 【前提】呼ぶ側のメインから Call Pr_マスタ点検(事業部) で呼ぶ。
'==============================================================

Private Const Lg_表示上限 As Long = 15      ' メッセージに出す誤りの数の上限（長くなりすぎないように）

' マスタを点検し、誤りがあれば一覧を表示して中断する
'   引数：St_事業部
Public Sub Pr_マスタ点検(ByVal St_事業部 As String)
    Dim Co_誤り As Collection
    Dim Ob_品目 As Object
    Dim Ob_作業区 As Object
    Set Co_誤り = New Collection
    Set Ob_品目 = Fn_品目点検(St_事業部, Co_誤り)
    Set Ob_作業区 = Fn_作業区点検(St_事業部, Co_誤り)
    Call Pr_BOM点検(St_事業部, Ob_品目, Co_誤り)
    Call Pr_工程表点検(St_事業部, Ob_品目, Ob_作業区, Co_誤り)
    Call Pr_誤り表示(Co_誤り)
End Sub

' 品目の点検。戻り値：Dictionary（品番 → "製造" か "購入"）
Private Function Fn_品目点検(ByVal St_事業部 As String, ByVal Co_誤り As Collection) As Object
    Dim Ob_品目 As Object
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim St_品番 As String
    Dim St_区分 As String
    Dim St_行 As String
    Set Ob_品目 = CreateObject("Scripting.Dictionary")
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_品目), 列_品目.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_品目.事業部)) = St_事業部 Then
            St_行 = "品目 " & (Lg_行 + Lg_開始行 - 1) & "行目："
            St_品番 = Fn_文字(Va_表(Lg_行, 列_品目.品番))
            St_区分 = Fn_文字(Va_表(Lg_行, 列_品目.区分))
            If St_品番 = "" Then
                Co_誤り.Add St_行 & "品番が空です"
            ElseIf Ob_品目.Exists(St_品番) Then
                Co_誤り.Add St_行 & "品番 " & St_品番 & " が重複しています"
            ElseIf Fn_製造品か(St_区分) Then
                Ob_品目.Add St_品番, "製造"
            ElseIf St_区分 = "材料" Or St_区分 = "部品" Then
                Ob_品目.Add St_品番, "購入"
            Else
                Co_誤り.Add St_行 & "区分は 材料・部品・中間品・加工品・完成品 のどれかにしてください（いま：" & St_区分 & "）"
            End If
            If Fn_製造品か(St_区分) Then
                If Fn_文字(Va_表(Lg_行, 列_品目.生産方式)) <> St_見込生産 And Fn_文字(Va_表(Lg_行, 列_品目.生産方式)) <> St_受注生産 Then
                    Co_誤り.Add St_行 & St_品番 & " の生産方式は 見込生産 か 受注生産 にしてください"
                End If
            ElseIf St_区分 = "材料" Or St_区分 = "部品" Then
                If Not Fn_数か(Va_表(Lg_行, 列_品目.購入単価), 0) Then Co_誤り.Add St_行 & St_品番 & " の購入単価は0以上の数にしてください"
                If Not Fn_数か(Va_表(Lg_行, 列_品目.LT), 0) Then Co_誤り.Add St_行 & St_品番 & " のLTは0以上の数にしてください"
            End If
            If Not Fn_数か(Va_表(Lg_行, 列_品目.安全在庫), 0) Then Co_誤り.Add St_行 & St_品番 & " の安全在庫は0以上の数にしてください"
            If Not Fn_数か(Va_表(Lg_行, 列_品目.在庫数), -1E+15) Then Co_誤り.Add St_行 & St_品番 & " の在庫数が数ではありません"
            If Fn_文字(Va_表(Lg_行, 列_品目.販売単価)) <> "" And Not Fn_数か(Va_表(Lg_行, 列_品目.販売単価), 0) Then
                Co_誤り.Add St_行 & St_品番 & " の販売単価は空欄か0以上の数にしてください"
            End If
        End If
    Next Lg_行
    Set Fn_品目点検 = Ob_品目
End Function

' 作業区の点検。戻り値：Dictionary（作業区コード → 1）
Private Function Fn_作業区点検(ByVal St_事業部 As String, ByVal Co_誤り As Collection) As Object
    Dim Ob_作業区 As Object
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim St_コード As String
    Dim St_行 As String
    Set Ob_作業区 = CreateObject("Scripting.Dictionary")
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_作業区), 列_作業区.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_作業区.事業部)) = St_事業部 Then
            St_行 = "作業区 " & (Lg_行 + Lg_開始行 - 1) & "行目："
            St_コード = Fn_文字(Va_表(Lg_行, 列_作業区.作業区))
            If St_コード = "" Or Ob_作業区.Exists(St_コード) Then
                Co_誤り.Add St_行 & "作業区コードが空か重複しています（" & St_コード & "）"
            Else
                Ob_作業区.Add St_コード, 1
            End If
            If Not Fn_数か(Va_表(Lg_行, 列_作業区.賃率), 0) Then Co_誤り.Add St_行 & St_コード & " の賃率は0以上の数にしてください"
            If Not Fn_数か(Va_表(Lg_行, 列_作業区.能力), 1) Then Co_誤り.Add St_行 & St_コード & " の能力は1以上の数にしてください"
        End If
    Next Lg_行
    Set Fn_作業区点検 = Ob_作業区
End Function

' BOM の点検（品目の有無、員数、循環）
Private Sub Pr_BOM点検(ByVal St_事業部 As String, ByVal Ob_品目 As Object, ByVal Co_誤り As Collection)
    Dim Va_表 As Variant
    Dim Ob_段 As Object
    Dim Lg_行 As Long
    Dim Lg_回 As Long
    Dim Bl_変化 As Boolean
    Dim St_親 As String
    Dim St_子 As String
    Dim St_行 As String
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_BOM), 列_BOM.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_BOM.事業部)) = St_事業部 Then
            St_行 = "BOM " & (Lg_行 + Lg_開始行 - 1) & "行目："
            St_親 = Fn_文字(Va_表(Lg_行, 列_BOM.親品番))
            St_子 = Fn_文字(Va_表(Lg_行, 列_BOM.子品番))
            If Not Ob_品目.Exists(St_親) Then
                Co_誤り.Add St_行 & "親品番 " & St_親 & " が品目にありません"
            ElseIf Ob_品目.Item(St_親) <> "製造" Then
                Co_誤り.Add St_行 & "親品番 " & St_親 & " は購入品です（BOM の親は製造品だけ）"
            End If
            If Not Ob_品目.Exists(St_子) Then Co_誤り.Add St_行 & "子品番 " & St_子 & " が品目にありません"
            If St_親 = St_子 Then Co_誤り.Add St_行 & "親と子が同じ品番です（" & St_親 & "）"
            If Not Fn_数か(Va_表(Lg_行, 列_BOM.員数), 0.000001) Then Co_誤り.Add St_行 & "員数は0より大きい数にしてください"
        End If
    Next Lg_行
    ' 循環：ローレベルコードの計算が品目数より多く回っても落ち着かなければ循環している
    Set Ob_段 = CreateObject("Scripting.Dictionary")
    Bl_変化 = True
    Lg_回 = 0
    Do While Bl_変化 And Lg_回 <= Ob_品目.Count + 1
        Bl_変化 = False
        Lg_回 = Lg_回 + 1
        For Lg_行 = 1 To Fn_行数(Va_表)
            If Fn_文字(Va_表(Lg_行, 列_BOM.事業部)) = St_事業部 Then
                St_親 = Fn_文字(Va_表(Lg_行, 列_BOM.親品番))
                St_子 = Fn_文字(Va_表(Lg_行, 列_BOM.子品番))
                If Fn_辞書値(Ob_段, St_子) < Fn_辞書値(Ob_段, St_親) + 1 Then
                    Ob_段.Item(St_子) = Fn_辞書値(Ob_段, St_親) + 1
                    Bl_変化 = True
                End If
            End If
        Next Lg_行
    Loop
    If Bl_変化 Then Co_誤り.Add "BOM：親子の関係が循環しています（ある品目が自分自身の部品になっています）"
End Sub

' 工程表の点検と、工程の無い製造品の点検
Private Sub Pr_工程表点検(ByVal St_事業部 As String, ByVal Ob_品目 As Object, ByVal Ob_作業区 As Object, ByVal Co_誤り As Collection)
    Dim Va_表 As Variant
    Dim Ob_工程あり As Object
    Dim Va_品番 As Variant
    Dim Lg_行 As Long
    Dim St_品番 As String
    Dim St_行 As String
    Set Ob_工程あり = CreateObject("Scripting.Dictionary")
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_工程表), 列_工程表.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_工程表.事業部)) = St_事業部 Then
            St_行 = "工程表 " & (Lg_行 + Lg_開始行 - 1) & "行目："
            St_品番 = Fn_文字(Va_表(Lg_行, 列_工程表.品番))
            Ob_工程あり.Item(St_品番) = 1
            If Not Ob_品目.Exists(St_品番) Then
                Co_誤り.Add St_行 & "品番 " & St_品番 & " が品目にありません"
            ElseIf Ob_品目.Item(St_品番) <> "製造" Then
                Co_誤り.Add St_行 & "品番 " & St_品番 & " は購入品です（工程は製造品だけ）"
            End If
            If Not Fn_数か(Va_表(Lg_行, 列_工程表.工程順), 0) Then Co_誤り.Add St_行 & "工程順が数ではありません"
            If Fn_文字(Va_表(Lg_行, 列_工程表.外注)) = St_印 Then
                If Fn_文字(Va_表(Lg_行, 列_工程表.外注先)) = "" Then Co_誤り.Add St_行 & "外注工程の外注先が空です"
                If Not Fn_数か(Va_表(Lg_行, 列_工程表.外注単価), 0) Then Co_誤り.Add St_行 & "外注単価は0以上の数にしてください"
                If Not Fn_数か(Va_表(Lg_行, 列_工程表.外注LT), 0) Then Co_誤り.Add St_行 & "外注LTは0以上の数にしてください"
            Else
                If Not Ob_作業区.Exists(Fn_文字(Va_表(Lg_行, 列_工程表.作業区))) Then
                    Co_誤り.Add St_行 & "作業区 " & Fn_文字(Va_表(Lg_行, 列_工程表.作業区)) & " が作業区シートにありません"
                End If
                If Not Fn_数か(Va_表(Lg_行, 列_工程表.段取時間), 0) Or Not Fn_数か(Va_表(Lg_行, 列_工程表.加工時間), 0) Then
                    Co_誤り.Add St_行 & "段取時間・加工時間は0以上の数にしてください"
                End If
            End If
        End If
    Next Lg_行
    For Each Va_品番 In Ob_品目.Keys
        If Ob_品目.Item(Va_品番) = "製造" And Not Ob_工程あり.Exists(CStr(Va_品番)) Then
            Co_誤り.Add "工程表：製造品 " & Va_品番 & " の工程が1つもありません"
        End If
    Next Va_品番
End Sub

' 値が、下限以上の数か
Private Function Fn_数か(ByVal Va_値 As Variant, ByVal Db_下限 As Double) As Boolean
    Fn_数か = False
    If IsError(Va_値) Then Exit Function
    If Fn_文字(Va_値) = "" Or Not IsNumeric(Va_値) Then Exit Function
    Fn_数か = (CDbl(Va_値) >= Db_下限)
End Function

' 誤りがあれば一覧を表示して中断する
Private Sub Pr_誤り表示(ByVal Co_誤り As Collection)
    Dim St_文 As String
    Dim Lg_添字 As Long
    If Co_誤り.Count = 0 Then Exit Sub
    For Lg_添字 = 1 To Co_誤り.Count
        If Lg_添字 > Lg_表示上限 Then
            St_文 = St_文 & "…ほか " & (Co_誤り.Count - Lg_表示上限) & " 件" & vbCrLf
            Exit For
        End If
        St_文 = St_文 & "・" & Co_誤り(Lg_添字) & vbCrLf
    Next Lg_添字
    MsgBox "マスタに誤りがあるため処理を止めました。直してからもう一度実行してください。" & vbCrLf & vbCrLf & St_文, vbExclamation, "マスタ点検"
    Err.Raise Lg_中断番号
End Sub
