Attribute VB_Name = "Mod_発注"
Option Explicit
'==============================================================
' Mod_発注
' 【目的】入力_発注 シートの操作を受け持つ。
'         段階2：発注（手動）、発注の取消。段階3：入荷の計上（在庫が増える）。
' 【業務背景】HTML 版の「発注」タブと同じ。購入品（材料・部品）を発注すると
'         「発注済」になり、入荷予定は 今日＋品目のリードタイム。
'         発注済の数量は MRP の「手配済」に入る。
' 【前提】入力_発注 シートは初期化で作る（Pr_入力発注シート作成）。
'==============================================================

' ---- 入力_発注 シートの配置 ----
Private Const St_登録品番セル As String = "C6"      ' 発注：品目
Private Const St_登録数量セル As String = "C7"      ' 発注：数量
Private Const St_対象発注セル As String = "C10"     ' 発注の操作：発注No
Private Const Lg_一覧行 As Long = 14                ' 発注一覧を書く行

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_発注 シート
Private Ws_発注 As Worksheet            ' 発注シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Dt_今日 As Date                 ' 今日（入荷予定の基準）
Private St_登録品番 As String           ' C6
Private Va_登録数量 As Variant          ' C7
Private St_対象発注 As String           ' C10

'==============================================================
' Pr_発注登録（メイン）  ボタン：入力_発注「発注する」
'==============================================================
Public Sub Pr_発注登録()
    Dim St_発注No As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力発注, St_シート_発注)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Call Pr_登録チェック                                    ' 1. 購入品か・数量を確認（誤った発注を作らない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2. 画面更新を止める（速くする）
    St_発注No = Fn_発注追加()                               ' 3. 発注を作る（入荷予定＝今日＋LT）
    Call Pr_登録結果表示(St_発注No)                         ' 4. 発注Noを結果欄に出す（次の操作に使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_発注取消（メイン）  ボタン：入力_発注「発注を取り消す」
'==============================================================
Public Sub Pr_発注取消()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力発注, St_シート_発注)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象発注行()                                 ' 1. 発注No の行を探す（操作する発注を決める）→無ければ中断
    Call Pr_取消チェック(Lg_行)                             ' 2. 発注済か確認（入荷済の記録を消さない）→NGなら中断
    Call Pr_取消確認                                        ' 3. 削除してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_発注行削除(Lg_行)                               ' 4. 発注の行を消す（手配済から外す）
    Call Pr_操作結果表示("発注を取り消しました")            ' 5. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 6. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_入荷計上（メイン）  ボタン：入力_発注「入荷を計上」
'==============================================================
Public Sub Pr_入荷計上()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力発注, St_シート_発注, St_シート_履歴)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象発注行()                                 ' 1. 発注No の行を探す（操作する発注を決める）→無ければ中断
    Call Pr_取消チェック(Lg_行)                             ' 2. 発注済か確認（二重に入荷しない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_入荷書込(Lg_行)                                 ' 3. 在庫を増やし入荷済にする（使える在庫にする）
    Call Pr_操作結果表示("入荷を計上しました")              ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力発注)
    Set Ws_発注 = ThisWorkbook.Worksheets(St_シート_発注)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_今日 = Date
    St_登録品番 = Fn_文字(Ws_入力.Range(St_登録品番セル).Value)
    Va_登録数量 = Ws_入力.Range(St_登録数量セル).Value
    St_対象発注 = Fn_文字(Ws_入力.Range(St_対象発注セル).Value)
End Sub

' 発注の入力を確かめる（購入品、数量）
Private Sub Pr_登録チェック()
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("品目「" & St_登録品番 & "」が " & Fn_事業部名(St_事業部) & " の品目にありません。")
    If Fn_製造品か(Fn_文字(Va_品目(列_品目.区分))) Then Call Pr_入力エラー("発注できるのは購入品（材料・部品）だけです。")
    If Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)) = 0 Then Call Pr_入力エラー(St_数量エラー)
End Sub

' 発注を作る。戻り値：発注No
Private Function Fn_発注追加() As String
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    Fn_発注追加 = Fn_発注作成(St_事業部, St_登録品番, Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)), Dt_今日)
End Function

' 発注No の行を探す。戻り値：発注シートの行番号。空・見つからなければ中断
Private Function Fn_対象発注行() As Long
    Dim Lg_行 As Long
    If St_対象発注 = "" Then Call Pr_入力エラー("操作する発注No（例：PO-0001）を入力してください。")
    Lg_行 = Fn_行検索(Ws_発注, 列_発注.次列 - 1, St_事業部, 列_発注.発注No, St_対象発注)
    If Lg_行 = 0 Then Call Pr_入力エラー("発注 " & St_対象発注 & " は " & Fn_事業部名(St_事業部) & " にありません。")
    Fn_対象発注行 = Lg_行
End Function

' 発注済か（取消・入荷とも、発注済のときだけ）
Private Sub Pr_取消チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_発注.Cells(Lg_行, 列_発注.状態).Value) <> St_状態_発注済 Then
        Call Pr_入力エラー("発注 " & St_対象発注 & " は「発注済」ではありません（入荷済）。")
    End If
End Sub

' 入荷：在庫を増やし（履歴「入荷」）、状態を入荷済にする
Private Sub Pr_入荷書込(ByVal Lg_行 As Long)
    Call Pr_在庫移動(St_事業部, Fn_文字(Ws_発注.Cells(Lg_行, 列_発注.品番).Value), "入荷", _
        Fn_数(Ws_発注.Cells(Lg_行, 列_発注.数量).Value), St_対象発注)
    Ws_発注.Cells(Lg_行, 列_発注.状態).Value = St_状態_入荷済
End Sub

' 取り消してよいか確認する
Private Sub Pr_取消確認()
    Call Pr_削除確認("発注 " & St_対象発注 & " を取り消します（発注シートの行を削除します）。")
End Sub

' 発注の行を消す
Private Sub Pr_発注行削除(ByVal Lg_行 As Long)
    Call Pr_行削除(Ws_発注, Array(Lg_行))
End Sub

' 結果欄への表示
Private Sub Pr_登録結果表示(ByVal St_発注No As String)
    Call Pr_結果表示(Ws_入力, "発注 " & St_発注No & " を作りました（発注済）")
    Ws_入力.Range(St_対象発注セル).Value = St_発注No
End Sub

Private Sub Pr_操作結果表示(ByVal St_文 As String)
    Call Pr_結果表示(Ws_入力, "発注 " & St_対象発注 & "：" & St_文)
End Sub

' 発注一覧を書き直す
Private Sub Pr_一覧更新()
    Call Pr_一覧転記(Ws_入力, Lg_一覧行, St_シート_発注, 列_発注.次列 - 1, St_事業部, "発注一覧")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力発注再表示()
    If Not Fn_シートあり(St_シート_入力発注) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_発注 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力発注シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力発注)
    Call Pr_入力シート枠(Ws_表, "発注", "購入品（材料・部品）を発注します。発注済 → 入荷（在庫が増える）。入荷予定は今日＋リードタイムです。")
    Call Pr_入力見出し(Ws_表, "B5", "発注")
    Call Pr_入力欄(Ws_表, St_登録品番セル, "発注する品目（品番）", "", "選択肢_購入品")
    Call Pr_入力欄(Ws_表, St_登録数量セル, "数量", 50, "")
    Call Pr_ボタン作成(Ws_表, "E6", 2, "発注する", "Pr_発注登録", "発注_登録")
    Call Pr_入力見出し(Ws_表, "B9", "発注の操作（発注No を入れてボタンを押す）")
    Call Pr_入力欄(Ws_表, St_対象発注セル, "発注No", "", "")
    Call Pr_ボタン作成(Ws_表, "E10", 2, "発注を取り消す", "Pr_発注取消", "発注_取消")
    Call Pr_ボタン作成(Ws_表, "G10", 2, "入荷を計上", "Pr_入荷計上", "発注_入荷")
End Sub
