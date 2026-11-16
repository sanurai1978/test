Attribute VB_Name = "Mod_負荷"
Option Explicit
'==============================================================
' Mod_負荷
' 【目的】作業区ごとの負荷（今日から14日分）を数字の一覧で「負荷一覧」シートに出す。
' 【業務背景】HTML 版「生産計画」タブの「作業区の負荷」と同じ計算。ガント表示は作らない。
'   計画・着手の製造指図の、まだ完了していない工程を、着手予定日から順に
'   作業区の1日の能力（分）ずつ積んでいく（山積み。能力は無限として積む）。
'   ・完了した工程は飛ばし、次の工程の数量はその良品数にする
'   ・外注工程は社内の時間を使わず、外注LTの日数だけ先へ送る
'     （HTML 版と同じく、日の区切りは 450分/日 で数える）
'   ・表示は能力に対する％（四捨五入）。100 を超えた日は能力オーバー
' 【前提】休日は区別しない（HTML 版と同じ）。
' 【処理の流れ】0. 変数格納 → 1. チェック → 2. 山積み計算 → 3. 書込 → 4. 表示
'==============================================================

Private Const Db_日の分 As Double = 450     ' 外注工程で日を送るときの1日（HTML 版の CAP）
Private Const Lg_表見出し行 As Long = 5     ' ％の表の見出し行

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_負荷 As Worksheet            ' 負荷一覧シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Lg_日数 As Long                 ' 設定「負荷表示日数」（14）
Private Dt_今日 As Date                 ' 今日（表の1列目）

'==============================================================
' Pr_負荷一覧更新（メイン／目次）  ボタン：メニュー「負荷一覧（更新）」
'==============================================================
Public Sub Pr_負荷一覧更新()
    Dim Va_負荷 As Variant

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_負荷一覧, St_シート_作業区, St_シート_指図工程)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. シート・事業部・表示日数を格納（読み直さない）
    Call Pr_入力チェック                                    ' 1. 事業部と日数を確認（誤った条件で積まない）→NGなら中断
    Call Pr_画面更新停止                                    ' 1b. 画面更新を止める（速くする）
    Va_負荷 = Fn_山積み()                                   ' 2. 作業区×日の負荷（分）を積む（HTMLと同じ計算）
    Call Pr_負荷書込(Va_負荷)                               ' 3. ％と分の表を書く（能力オーバーの日を見つける）
    Call Pr_負荷表示                                        ' 4. 負荷一覧シートを開く（結果を見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_負荷 = ThisWorkbook.Worksheets(St_シート_負荷一覧)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Lg_日数 = CLng(Fn_数(Fn_設定値("負荷表示日数")))
    Dt_今日 = Date
End Sub

' 事業部と表示日数を確かめる
Private Sub Pr_入力チェック()
    If St_事業部 <> St_加工 And St_事業部 <> St_組立 Then Call Pr_入力エラー("設定シートの「現在の事業部」が正しくありません。")
    If Lg_日数 < 1 Or Lg_日数 > 60 Then Call Pr_入力エラー("設定シートの「負荷表示日数」は1〜60にしてください。")
End Sub

' 山積みする（HTML 版の loadMap）
'   戻り値：2要素の配列。(0)＝作業区の情報（行×3：コード・名前・能力）、(1)＝負荷（行×日、分）
Private Function Fn_山積み() As Variant
    Dim Va_作業区 As Variant, Va_指図 As Variant, Va_工程 As Variant
    Dim Va_情報() As Variant, Db_負荷() As Double
    Dim Ob_作業区 As Object
    Dim Lg_作業区数 As Long, Lg_行 As Long, Lg_工程 As Long, Lg_日 As Long, Lg_区 As Long
    Dim St_指図 As String, St_状態 As String
    Dim Lg_着手 As Long
    Dim Db_累計 As Double, Db_数量 As Double, Db_分 As Double, Db_能力 As Double, Db_余り As Double, Db_使う As Double
    ' 作業区（この事業部）
    Va_作業区 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_作業区), 列_作業区.次列 - 1)
    Set Ob_作業区 = CreateObject("Scripting.Dictionary")
    For Lg_行 = 1 To Fn_行数(Va_作業区)
        If Fn_文字(Va_作業区(Lg_行, 列_作業区.事業部)) = St_事業部 Then
            Lg_作業区数 = Lg_作業区数 + 1
            Ob_作業区.Item(Fn_文字(Va_作業区(Lg_行, 列_作業区.作業区))) = Lg_作業区数
        End If
    Next Lg_行
    If Lg_作業区数 = 0 Then Call Pr_入力エラー("この事業部の作業区がありません。")
    ReDim Va_情報(1 To Lg_作業区数, 1 To 3)
    ReDim Db_負荷(1 To Lg_作業区数, 1 To Lg_日数)
    For Lg_行 = 1 To Fn_行数(Va_作業区)
        If Ob_作業区.Exists(Fn_文字(Va_作業区(Lg_行, 列_作業区.作業区))) And Fn_文字(Va_作業区(Lg_行, 列_作業区.事業部)) = St_事業部 Then
            Lg_区 = Ob_作業区.Item(Fn_文字(Va_作業区(Lg_行, 列_作業区.作業区)))
            Va_情報(Lg_区, 1) = Va_作業区(Lg_行, 列_作業区.作業区)
            Va_情報(Lg_区, 2) = Va_作業区(Lg_行, 列_作業区.作業区名)
            Va_情報(Lg_区, 3) = Fn_数(Va_作業区(Lg_行, 列_作業区.能力))
            If Va_情報(Lg_区, 3) <= 0 Then Va_情報(Lg_区, 3) = Db_日の分
        End If
    Next Lg_行
    ' 計画・着手の製造指図ごとに、未完了の工程を着手日から積む
    Va_指図 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_製造指図), 列_製造指図.次列 - 1)
    Va_工程 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_指図工程), 列_指図工程.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_指図)
        St_状態 = Fn_文字(Va_指図(Lg_行, 列_製造指図.状態))
        If Fn_文字(Va_指図(Lg_行, 列_製造指図.事業部)) = St_事業部 And (St_状態 = St_状態_計画 Or St_状態 = St_状態_着手) Then
            St_指図 = Fn_文字(Va_指図(Lg_行, 列_製造指図.指図No))
            Lg_着手 = Fn_日数差(Va_指図(Lg_行, 列_製造指図.着手予定), Dt_今日)
            If Lg_着手 < 0 Then Lg_着手 = 0
            Db_累計 = 0
            Db_数量 = Fn_数(Va_指図(Lg_行, 列_製造指図.数量))
            ' 工程は工程順に並んでいる（作成時に工程順で複写している）
            For Lg_工程 = 1 To Fn_行数(Va_工程)
                If Fn_文字(Va_工程(Lg_工程, 列_指図工程.事業部)) = St_事業部 And Fn_文字(Va_工程(Lg_工程, 列_指図工程.指図No)) = St_指図 Then
                    If Fn_文字(Va_工程(Lg_工程, 列_指図工程.状態)) = St_状態_完了 Then
                        ' 完了した工程は積まない。次の工程の数量はこの工程の良品数
                        Db_数量 = Fn_数(Va_工程(Lg_工程, 列_指図工程.良品))
                    ElseIf Fn_文字(Va_工程(Lg_工程, 列_指図工程.外注)) = St_印 Then
                        ' 外注：今日の残りを切り上げて日を改め、外注LTの日数だけ先へ送る
                        Db_累計 = -Int(-(Db_累計 / Db_日の分 - Db_誤差)) * Db_日の分 + Fn_数(Va_工程(Lg_工程, 列_指図工程.外注LT)) * Db_日の分
                    Else
                        Db_分 = Fn_丸め(Fn_数(Va_工程(Lg_工程, 列_指図工程.段取時間)) + Fn_数(Va_工程(Lg_工程, 列_指図工程.加工時間)) * Db_数量)
                        Lg_区 = 0
                        If Ob_作業区.Exists(Fn_文字(Va_工程(Lg_工程, 列_指図工程.作業区))) Then Lg_区 = Ob_作業区.Item(Fn_文字(Va_工程(Lg_工程, 列_指図工程.作業区)))
                        If Lg_区 > 0 Then Db_能力 = Va_情報(Lg_区, 3) Else Db_能力 = Db_日の分
                        ' 1日の能力の残りに入るだけ入れ、あふれた分は翌日へ
                        Do While Db_分 > 0.001
                            Lg_日 = Lg_着手 + Int(Db_累計 / Db_能力 + Db_誤差)
                            Db_余り = Db_能力 - (Db_累計 - Db_能力 * Fix(Db_累計 / Db_能力))
                            Db_使う = Db_余り
                            If Db_分 < Db_使う Then Db_使う = Db_分
                            If Lg_日 < Lg_日数 And Lg_区 > 0 Then Db_負荷(Lg_区, Lg_日 + 1) = Db_負荷(Lg_区, Lg_日 + 1) + Db_使う
                            Db_累計 = Db_累計 + Db_使う
                            Db_分 = Db_分 - Db_使う
                        Loop
                    End If
                End If
            Next Lg_工程
        End If
    Next Lg_行
    Fn_山積み = Array(Va_情報, Db_負荷)
End Function

' ％の表と分の表を書く
Private Sub Pr_負荷書込(ByVal Va_負荷 As Variant)
    Dim Va_情報 As Variant
    Dim Va_分 As Variant
    Dim Va_率() As Variant
    Dim Va_分表() As Variant
    Dim Lg_区 As Long
    Dim Lg_日 As Long
    Dim Lg_超過 As Long
    Dim Lg_分表行 As Long
    Dim Lg_作業区数 As Long
    Va_情報 = Va_負荷(0)
    Va_分 = Va_負荷(1)
    Lg_作業区数 = UBound(Va_情報, 1)
    ReDim Va_率(1 To Lg_作業区数, 1 To Lg_日数 + 2)
    ReDim Va_分表(1 To Lg_作業区数, 1 To Lg_日数 + 2)
    For Lg_区 = 1 To Lg_作業区数
        Va_率(Lg_区, 1) = Va_情報(Lg_区, 1): Va_率(Lg_区, 2) = Va_情報(Lg_区, 2)
        Va_分表(Lg_区, 1) = Va_情報(Lg_区, 1): Va_分表(Lg_区, 2) = Va_情報(Lg_区, 2)
        For Lg_日 = 1 To Lg_日数
            Va_率(Lg_区, Lg_日 + 2) = Int(Va_分(Lg_区, Lg_日) / Va_情報(Lg_区, 3) * 100 + 0.5)
            Va_分表(Lg_区, Lg_日 + 2) = Fn_丸め(Va_分(Lg_区, Lg_日))
            If Va_分(Lg_区, Lg_日) > Va_情報(Lg_区, 3) + 0.01 Then Lg_超過 = Lg_超過 + 1
        Next Lg_日
    Next Lg_区
    Ws_負荷.Cells.Clear
    Ws_負荷.Range("B1").Value = "作業区の負荷（" & Fn_事業部名(St_事業部) & "。今日から " & Lg_日数 & " 日）"
    Ws_負荷.Range("B1").Font.Bold = True
    Ws_負荷.Range("B1").Font.Size = 14
    Ws_負荷.Range("B2").Value = "計画・着手の製造指図の未完了の工程を、着手予定日から1日の能力ずつ積んだ値です。休日は区別しません。100を超えた日は能力オーバー（着手日をずらして山をならす）。"
    Ws_負荷.Range("B3").Value = "能力超過：" & Lg_超過 & " 箇所　（更新：" & Format$(Now, "yyyy/mm/dd hh:nn") & "）"
    If Lg_超過 > 0 Then Ws_負荷.Range("B3").Font.Color = RGB(192, 0, 0)
    ' ％の表
    Call Pr_表書込(Lg_表見出し行, "能力に対する％", Va_率)
    Call Pr_色付け(Lg_表見出し行 + 1, Lg_作業区数)
    ' 分の表
    Lg_分表行 = Lg_表見出し行 + Lg_作業区数 + 3
    Call Pr_表書込(Lg_分表行, "負荷（分）", Va_分表)
    Ws_負荷.Columns("B:C").AutoFit
End Sub

' 1つの表（見出し＋作業区ごとの行）を書く
Private Sub Pr_表書込(ByVal Lg_見出し行 As Long, ByVal St_題名 As String, ByVal Va_表 As Variant)
    Dim Lg_日 As Long
    Ws_負荷.Cells(Lg_見出し行 - 1, 2).Value = "■ " & St_題名
    Ws_負荷.Cells(Lg_見出し行 - 1, 2).Font.Bold = True
    Ws_負荷.Cells(Lg_見出し行, 2).Value = "作業区"
    Ws_負荷.Cells(Lg_見出し行, 3).Value = "作業区名"
    For Lg_日 = 1 To Lg_日数
        Ws_負荷.Cells(Lg_見出し行, Lg_日 + 3).Value = Dt_今日 + Lg_日 - 1
    Next Lg_日
    Ws_負荷.Range(Ws_負荷.Cells(Lg_見出し行, 4), Ws_負荷.Cells(Lg_見出し行, Lg_日数 + 3)).NumberFormat = "m/d"
    With Ws_負荷.Range(Ws_負荷.Cells(Lg_見出し行, 2), Ws_負荷.Cells(Lg_見出し行, Lg_日数 + 3))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    Ws_負荷.Range(Ws_負荷.Cells(Lg_見出し行 + 1, 2), Ws_負荷.Cells(Lg_見出し行 + UBound(Va_表, 1), Lg_日数 + 3)).Value = Va_表
End Sub

' ％の表に色を付ける（80超＝黄、100超＝赤。HTML 版の l2・l3 と同じ区切り）
Private Sub Pr_色付け(ByVal Lg_開始 As Long, ByVal Lg_作業区数 As Long)
    Dim Lg_行 As Long
    Dim Lg_列 As Long
    Dim Db_値 As Double
    For Lg_行 = Lg_開始 To Lg_開始 + Lg_作業区数 - 1
        For Lg_列 = 4 To Lg_日数 + 3
            Db_値 = Fn_数(Ws_負荷.Cells(Lg_行, Lg_列).Value)
            If Db_値 > 100 Then
                Ws_負荷.Cells(Lg_行, Lg_列).Interior.Color = RGB(255, 199, 206)
            ElseIf Db_値 > 80 Then
                Ws_負荷.Cells(Lg_行, Lg_列).Interior.Color = RGB(255, 235, 156)
            End If
        Next Lg_列
    Next Lg_行
End Sub

' 負荷一覧シートを開く
Private Sub Pr_負荷表示()
    Ws_負荷.Activate
    Ws_負荷.Range("A1").Select
End Sub
