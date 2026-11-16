Attribute VB_Name = "Mod_受注"
Option Explicit
'==============================================================
' Mod_受注
' 【目的】入力_受注 シートの操作（受注登録・在庫引当・引当解除・受注取消・
'         受注生産の受注向け製造指図の作成）を受け持つ。
' 【業務背景】HTML 版の「受注」タブと同じ挙動にする。
'   ・受注は「受付」で登録する。在庫の引当は手で行い、有効在庫が受注残以上の
'     ときだけ「引当済」になる（足りなければ MRP で手配する）。
'   ・引当解除は、出荷用ピッキング中なら不可。
'   ・取消は「受付」で未出荷の受注だけ。紐付いていた製造指図は紐付きを外す。
'   ・受注生産の品目は、受注向けの製造指図を作れる。数量は
'     受注残 － すでに紐付いて作っている数。完成予定を納期に合わせて着手日を逆算する。
' 【前提】入力_受注 シートは初期化で作る（Pr_入力受注シート作成）。
'==============================================================

' ---- 入力_受注 シートの配置 ----
Private Const St_登録品番セル As String = "C6"      ' 受注登録：製品（品番）
Private Const St_登録数量セル As String = "C7"      ' 受注登録：数量
Private Const St_得意先セル As String = "C8"        ' 受注登録：得意先
Private Const St_納期セル As String = "C9"          ' 受注登録：納期（今日から何日後）
Private Const St_対象受注セル As String = "C12"     ' 受注の操作：受注No
Private Const Lg_一覧行 As Long = 16                ' 受注一覧を書く行

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_受注 シート
Private Ws_受注 As Worksheet            ' 受注シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Dt_今日 As Date                 ' 今日（納期の基準）
Private St_登録品番 As String           ' C6
Private Va_登録数量 As Variant          ' C7（確認前の値）
Private St_得意先 As String             ' C8
Private Va_納期日数 As Variant          ' C9（確認前の値）
Private St_対象受注 As String           ' C12

'==============================================================
' Pr_受注登録（メイン）  ボタン：入力_受注「受注を登録」
'==============================================================
Public Sub Pr_受注登録()
    Dim St_受注No As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力受注, St_シート_受注)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Call Pr_登録チェック                                    ' 1. 製品・数量・納期を確認（誤った受注を作らない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2. 画面更新を止める（速くする）
    St_受注No = Fn_受注追加()                               ' 3. 受注を「受付」で登録（需要として MRP に載せる）
    Call Pr_登録結果表示(St_受注No)                         ' 4. 登録した番号を結果欄に出す（次の操作に使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_受注引当（メイン）  ボタン：入力_受注「在庫を引き当てる」
'==============================================================
Public Sub Pr_受注引当()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力受注, St_シート_受注)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象受注行()                                 ' 1. 受注No の行を探す（操作する受注を決める）→無ければ中断
    Call Pr_引当チェック(Lg_行)                             ' 2. 受付か・有効在庫が足りるか確認（在庫の二重確保を防ぐ）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_受注状態変更(Lg_行, St_状態_引当済)              ' 3. 状態を引当済にする（在庫をこの受注に確保）
    Call Pr_操作結果表示("在庫を引き当てました")            ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_受注引当解除（メイン）  ボタン：入力_受注「引当を解除」
'==============================================================
Public Sub Pr_受注引当解除()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力受注, St_シート_受注)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象受注行()                                 ' 1. 受注No の行を探す（操作する受注を決める）→無ければ中断
    Call Pr_引当解除チェック(Lg_行)                         ' 2. 引当済か・ピッキング中でないか確認（作業中の品を外さない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_受注状態変更(Lg_行, St_状態_受付)                ' 3. 状態を受付に戻す（確保した在庫を放す）
    Call Pr_操作結果表示("引当を解除しました")              ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_受注取消（メイン）  ボタン：入力_受注「受注を取り消す」
'==============================================================
Public Sub Pr_受注取消()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力受注, St_シート_受注)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象受注行()                                 ' 1. 受注No の行を探す（操作する受注を決める）→無ければ中断
    Call Pr_取消チェック(Lg_行)                             ' 2. 受付で未出荷か確認（出荷済の記録を壊さない）→NGなら中断
    Call Pr_取消確認                                        ' 3. 削除してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_紐付解除                                        ' 4. 紐付いた製造指図の受注Noを外す（宙に浮いた紐付を残さない）
    Call Pr_受注行削除(Lg_行)                               ' 5. 受注の行を消す（需要から外す）
    Call Pr_操作結果表示("受注を取り消しました")            ' 6. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 7. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_受注向け製造指図作成（メイン）  ボタン：入力_受注「この受注向けの製造指図を作る」
'==============================================================
Public Sub Pr_受注向け製造指図作成()
    Dim Lg_行 As Long
    Dim Db_不足 As Double
    Dim St_指図No As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力受注, St_シート_受注, St_シート_製造指図)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象受注行()                                 ' 1. 受注No の行を探す（操作する受注を決める）→無ければ中断
    Call Pr_受注生産チェック(Lg_行)                         ' 2. 受付で受注生産の品目か確認（見込生産は MRP で作る）→NGなら中断
    Call Pr_マスタ点検(St_事業部)                           ' 3. マスタの誤りを点検（工程の複写を狂わせない）→NGなら中断
    Db_不足 = Fn_紐付不足数(Lg_行)                          ' 4. 受注残－紐付き製造中を出す（作りすぎない）→0以下なら中断
    Call Pr_画面更新停止                                    ' 4b. 画面更新を止める（速くする）
    St_指図No = Fn_受注向け指図追加(Lg_行, Db_不足)         ' 5. 納期から逆算して製造指図を作る（受注に紐付ける）
    Call Pr_指図作成結果表示(St_指図No, Db_不足)            ' 6. 作った指図を結果欄に出す（次の操作に使う）
    Call Pr_画面最新化                                      ' 7. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力受注)
    Set Ws_受注 = ThisWorkbook.Worksheets(St_シート_受注)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_今日 = Date
    St_登録品番 = Fn_文字(Ws_入力.Range(St_登録品番セル).Value)
    Va_登録数量 = Ws_入力.Range(St_登録数量セル).Value
    St_得意先 = Fn_文字(Ws_入力.Range(St_得意先セル).Value)
    Va_納期日数 = Ws_入力.Range(St_納期セル).Value
    St_対象受注 = Fn_文字(Ws_入力.Range(St_対象受注セル).Value)
End Sub

' 受注登録の入力を確かめる（製品は加工品・完成品、数量、納期 0〜365日）
Private Sub Pr_登録チェック()
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("製品「" & St_登録品番 & "」が " & Fn_事業部名(St_事業部) & " の品目にありません。")
    If Not Fn_製造品か(Fn_文字(Va_品目(列_品目.区分))) Or Fn_文字(Va_品目(列_品目.区分)) = "中間品" Then
        Call Pr_入力エラー("受注できるのは加工品・完成品だけです（" & St_登録品番 & " は " & Va_品目(列_品目.区分) & "）。")
    End If
    If Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)) = 0 Then Call Pr_入力エラー(St_数量エラー)
    If Fn_日数変換(Va_納期日数, 0, 365) < 0 Then Call Pr_入力エラー("納期は0〜365の整数（今日から何日後）で入力してください。")
End Sub

' 受注を1件追加する。戻り値：受注No
'   得意先が空なら「得意先未設定」（HTML 版と同じ）。
Private Function Fn_受注追加() As String
    Dim St_受注No As String
    Dim St_名前 As String
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    St_名前 = St_得意先
    If St_名前 = "" Then St_名前 = "得意先未設定"
    St_受注No = Fn_採番(St_事業部, "SO")
    Call Fn_行追記(Ws_受注, Array(St_事業部, St_受注No, St_登録品番, Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)), 0, _
        St_名前, Dt_今日 + Fn_日数変換(Va_納期日数, 0, 365), St_状態_受付))
    Fn_受注追加 = St_受注No
End Function

' 受注No の行を探す。戻り値：受注シートの行番号。空・見つからなければ中断
Private Function Fn_対象受注行() As Long
    Dim Lg_行 As Long
    If St_対象受注 = "" Then Call Pr_入力エラー("操作する受注No（例：SO-0001）を入力してください。")
    Lg_行 = Fn_行検索(Ws_受注, 列_受注.次列 - 1, St_事業部, 列_受注.受注No, St_対象受注)
    If Lg_行 = 0 Then Call Pr_入力エラー("受注 " & St_対象受注 & " は " & Fn_事業部名(St_事業部) & " にありません。")
    Fn_対象受注行 = Lg_行
End Function

' 引当できるか（受付で、有効在庫が受注残以上）
Private Sub Pr_引当チェック(ByVal Lg_行 As Long)
    Dim St_品番 As String
    Dim Db_残 As Double
    Dim Db_有効 As Double
    If Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.状態).Value) <> St_状態_受付 Then
        Call Pr_入力エラー("受注 " & St_対象受注 & " は「" & Ws_受注.Cells(Lg_行, 列_受注.状態).Value & "」なので引き当てられません（受付のときだけ）。")
    End If
    St_品番 = Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.品番).Value)
    Db_残 = Fn_丸め(Fn_数(Ws_受注.Cells(Lg_行, 列_受注.数量).Value) - Fn_数(Ws_受注.Cells(Lg_行, 列_受注.出荷済).Value))
    Db_有効 = Fn_有効在庫(St_事業部, St_品番)
    If Db_有効 < Db_残 Then
        Call Pr_入力エラー("有効在庫が足りません（有効 " & Db_有効 & " ／ 必要 " & Db_残 & "）。MRP で手配してください。")
    End If
End Sub

' 引当を解除できるか（引当済で、出荷用ピッキング中でない）
Private Sub Pr_引当解除チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.状態).Value) <> St_状態_引当済 Then
        Call Pr_入力エラー("受注 " & St_対象受注 & " は引当済ではありません。")
    End If
    If Fn_出荷ピッキング中(St_事業部, St_対象受注) > 0 Then
        Call Pr_入力エラー("ピッキング中のため解除できません。先にピッキング指示を取り消してください。")
    End If
End Sub

' 取り消せるか（受付で、まだ1つも出荷していない）
Private Sub Pr_取消チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.状態).Value) <> St_状態_受付 Then
        Call Pr_入力エラー("取り消せるのは「受付」の受注だけです（いま：" & Ws_受注.Cells(Lg_行, 列_受注.状態).Value & "）。引当済なら先に引当を解除してください。")
    End If
    If Fn_数(Ws_受注.Cells(Lg_行, 列_受注.出荷済).Value) > 0 Then
        Call Pr_入力エラー("一部出荷済みの受注は取り消せません。")
    End If
End Sub

' 取り消してよいか確認する。「いいえ」なら中断
Private Sub Pr_取消確認()
    Call Pr_削除確認("受注 " & St_対象受注 & " を取り消します（受注シートの行を削除します）。")
End Sub

' 受注の行を消す
Private Sub Pr_受注行削除(ByVal Lg_行 As Long)
    Call Pr_行削除(Ws_受注, Array(Lg_行))
End Sub

' 受注の状態を書き換える
Private Sub Pr_受注状態変更(ByVal Lg_行 As Long, ByVal St_状態 As String)
    Ws_受注.Cells(Lg_行, 列_受注.状態).Value = St_状態
End Sub

' この受注に紐付いた製造指図の受注Noを空にする（HTML 版：受注取消で m.so=null）
Private Sub Pr_紐付解除()
    Dim Ws_指図 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
    Va_表 = Fn_表読込(Ws_指図, 列_製造指図.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_製造指図.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_製造指図.受注No)) = St_対象受注 Then
            Ws_指図.Cells(Lg_行 + Lg_開始行 - 1, 列_製造指図.受注No).Value = ""
        End If
    Next Lg_行
End Sub

' 受注生産の受注向けに指図を作れるか（受付で、品目が受注生産）
Private Sub Pr_受注生産チェック(ByVal Lg_行 As Long)
    Dim Va_品目 As Variant
    If Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.状態).Value) <> St_状態_受付 Then
        Call Pr_入力エラー("受注 " & St_対象受注 & " は受付ではありません（引当済・出荷済の受注には作りません）。")
    End If
    Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.品番).Value))
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("受注の品目が品目マスタにありません。")
    If Fn_文字(Va_品目(列_品目.生産方式)) <> St_受注生産 Then
        Call Pr_入力エラー("この受注の品目は見込生産です。見込生産の品目は MRP で手配します。")
    End If
End Sub

' 受注向けに作る数（受注残 － すでに紐付いて作っている数）。0以下なら中断
Private Function Fn_紐付不足数(ByVal Lg_行 As Long) As Double
    Dim Db_残 As Double
    Dim Db_不足 As Double
    Db_残 = Fn_丸め(Fn_数(Ws_受注.Cells(Lg_行, 列_受注.数量).Value) - Fn_数(Ws_受注.Cells(Lg_行, 列_受注.出荷済).Value))
    Db_不足 = Fn_切上げ(Fn_丸め(Db_残 - Fn_紐付出来高(St_事業部, St_対象受注, Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.品番).Value))))
    If Db_不足 <= 0 Then Call Pr_入力エラー("紐付いた製造指図で必要数は足りています。")
    Fn_紐付不足数 = Db_不足
End Function

' 受注向けの製造指図を作る。完成予定＝納期になるよう着手日を逆算（今日より前なら今日）
'   戻り値：指図No
Private Function Fn_受注向け指図追加(ByVal Lg_行 As Long, ByVal Db_数量 As Double) As String
    Dim St_品番 As String
    Dim Lg_日数 As Long
    Dim Lg_着手 As Long
    St_品番 = Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.品番).Value)
    Lg_日数 = Fn_所要日数(St_事業部, St_品番, Db_数量)
    Lg_着手 = Fn_日数差(Ws_受注.Cells(Lg_行, 列_受注.納期).Value, Dt_今日) - Lg_日数
    If Lg_着手 < 0 Then Lg_着手 = 0
    Fn_受注向け指図追加 = Fn_製造指図作成(St_事業部, St_品番, Db_数量, Lg_着手, Lg_着手 + Lg_日数, St_対象受注, Dt_今日)
End Function

' 結果欄への表示
Private Sub Pr_登録結果表示(ByVal St_受注No As String)
    Call Pr_結果表示(Ws_入力, "受注 " & St_受注No & " を登録しました（受付）")
    Ws_入力.Range(St_対象受注セル).Value = St_受注No
End Sub

Private Sub Pr_操作結果表示(ByVal St_文 As String)
    Call Pr_結果表示(Ws_入力, "受注 " & St_対象受注 & "：" & St_文)
End Sub

Private Sub Pr_指図作成結果表示(ByVal St_指図No As String, ByVal Db_数量 As Double)
    Call Pr_結果表示(Ws_入力, "受注 " & St_対象受注 & " 向けに製造指図 " & St_指図No & "（" & Db_数量 & "）を作りました")
End Sub

' 受注一覧を書き直す
Private Sub Pr_一覧更新()
    Call Pr_一覧転記(Ws_入力, Lg_一覧行, St_シート_受注, 列_受注.次列 - 1, St_事業部, "受注一覧")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力受注再表示()
    If Not Fn_シートあり(St_シート_入力受注) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_受注 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力受注シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力受注)
    Call Pr_入力シート枠(Ws_表, "受注", _
        "受付 → 引当（在庫を確保）→ 出荷。見込生産の品目は在庫から引き当て、足りなければ MRP で補充。受注生産の品目は受注ごとに製造指図を紐付けます。")
    Call Pr_入力見出し(Ws_表, "B5", "受注登録")
    Call Pr_入力欄(Ws_表, St_登録品番セル, "製品（品番）", "", "選択肢_製品")
    Call Pr_入力欄(Ws_表, St_登録数量セル, "数量", 10, "")
    Call Pr_入力欄(Ws_表, St_得意先セル, "得意先（架空）", "東和機械", "")
    Call Pr_入力欄(Ws_表, St_納期セル, "納期（今日から何日後）", 14, "")
    Call Pr_ボタン作成(Ws_表, "E6", 2, "受注を登録", "Pr_受注登録", "受注_登録")
    Call Pr_入力見出し(Ws_表, "B11", "受注の操作（受注No を入れてボタンを押す）")
    Call Pr_入力欄(Ws_表, St_対象受注セル, "受注No", "", "")
    Call Pr_ボタン作成(Ws_表, "E12", 2, "在庫を引き当てる", "Pr_受注引当", "受注_引当")
    Call Pr_ボタン作成(Ws_表, "G12", 2, "引当を解除", "Pr_受注引当解除", "受注_引当解除")
    Call Pr_ボタン作成(Ws_表, "I12", 2, "受注を取り消す", "Pr_受注取消", "受注_取消")
    Call Pr_ボタン作成(Ws_表, "E14", 4, "この受注向けの製造指図を作る（受注生産）", "Pr_受注向け製造指図作成", "受注_向け指図")
End Sub
