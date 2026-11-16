Attribute VB_Name = "Mod_在庫調整"
Option Explicit
'==============================================================
' Mod_在庫調整
' 【目的】入力_在庫調整 シートの操作（棚卸の結果で在庫を直す）と、履歴の表示を受け持つ。
' 【業務背景】HTML 版の在庫タブの「在庫調整（棚卸）」と同じ。
'   実際に数えた数量を入れると、差分を「調整」として履歴に残して在庫を直す。
'   引当済（受注引当＋製造用確保）より少なくはできない（先に引当を解除する）。
'   品目シートの在庫数を直接書き換えると履歴が残らないので、在庫はここで直す。
' 【前提】入力_在庫調整 シートは初期化で作る（Pr_入力在庫調整シート作成）。
'==============================================================

' ---- 入力_在庫調整 シートの配置 ----
Private Const St_品番セル As String = "C6"          ' 品番
Private Const St_実在庫セル As String = "C7"        ' 実際に数えた数量
Private Const Lg_履歴行 As Long = 10                ' 履歴を書く行

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_在庫調整 シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private St_品番 As String               ' C6
Private Va_実在庫 As Variant            ' C7

'==============================================================
' Pr_在庫調整（メイン）  ボタン：入力_在庫調整「在庫を更新」
'==============================================================
Public Sub Pr_在庫調整()
    Dim Db_差 As Double

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力在庫調整, St_シート_品目, St_シート_履歴)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Call Pr_調整チェック                                    ' 1. 品番・数量・引当済以上か確認（確保した在庫を消さない）→NGなら中断
    Call Pr_画面更新停止                                    ' 1b. 画面更新を止める（速くする）
    Db_差 = Fn_在庫を直す()                                 ' 2. 差分を調整として在庫を直す（履歴に残す）
    Call Pr_調整結果表示(Db_差)                             ' 3. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 4. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力在庫調整)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    St_品番 = Fn_文字(Ws_入力.Range(St_品番セル).Value)
    Va_実在庫 = Ws_入力.Range(St_実在庫セル).Value
End Sub

' 品番があるか、数量が0以上の数か、引当済以上か
Private Sub Pr_調整チェック()
    Dim Va_品目 As Variant
    Dim Db_引当 As Double
    Va_品目 = Fn_品目情報(St_事業部, St_品番)
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("品番「" & St_品番 & "」が " & Fn_事業部名(St_事業部) & " の品目にありません。")
    If Fn_文字(Va_実在庫) = "" Or Not IsNumeric(Va_実在庫) Then Call Pr_入力エラー("0以上の数を入力してください。")
    If CDbl(Va_実在庫) < 0 Then Call Pr_入力エラー("0以上の数を入力してください。")
    Db_引当 = Fn_丸め(Fn_数(Va_品目(列_品目.在庫数)) - Fn_有効在庫(St_事業部, St_品番))
    If CDbl(Va_実在庫) < Db_引当 Then
        Call Pr_入力エラー("引当済（" & Fn_表示数(Db_引当) & "）より少なくはできません。先に引当を解除してください。")
    End If
End Sub

' 差分があれば「調整」で在庫を動かす。戻り値：差分（0なら変化なし）
Private Function Fn_在庫を直す() As Double
    Dim Va_品目 As Variant
    Dim Db_差 As Double
    Va_品目 = Fn_品目情報(St_事業部, St_品番)
    Db_差 = Fn_丸め(Fn_丸め(CDbl(Va_実在庫)) - Fn_数(Va_品目(列_品目.在庫数)))
    If Db_差 <> 0 Then Call Pr_在庫移動(St_事業部, St_品番, "調整", Db_差, "棚卸")
    Fn_在庫を直す = Db_差
End Function

' 結果欄への表示
Private Sub Pr_調整結果表示(ByVal Db_差 As Double)
    If Db_差 = 0 Then
        Call Pr_結果表示(Ws_入力, St_品番 & "：在庫は変わりません")
    Else
        Call Pr_結果表示(Ws_入力, St_品番 & "：在庫を " & IIf(Db_差 > 0, "+", "") & Fn_表示数(Db_差) & " 調整しました")
    End If
End Sub

' 履歴の一覧を書き直す
Private Sub Pr_一覧更新()
    Call Pr_一覧転記(Ws_入力, Lg_履歴行, St_シート_履歴, 列_履歴.次列 - 1, St_事業部, "在庫の動き（履歴）")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力在庫調整再表示()
    If Not Fn_シートあり(St_シート_入力在庫調整) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_在庫調整 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力在庫調整シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力在庫調整)
    Call Pr_入力シート枠(Ws_表, "在庫調整（棚卸）", "実際に数えた数量を入力します。差分が「調整」として履歴に残ります。引当済より少なくはできません。")
    Call Pr_入力見出し(Ws_表, "B5", "在庫調整")
    Call Pr_入力欄(Ws_表, St_品番セル, "品番", "", "")
    Call Pr_入力欄(Ws_表, St_実在庫セル, "実在庫（数えた数）", "", "")
    Call Pr_ボタン作成(Ws_表, "E6", 2, "在庫を更新", "Pr_在庫調整", "調整_更新")
End Sub
