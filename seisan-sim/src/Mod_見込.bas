Attribute VB_Name = "Mod_見込"
Option Explicit
'==============================================================
' Mod_見込
' 【目的】入力_見込 シートの操作（見込の登録・数量変更・削除）と、
'         品目ごとの見込の消し込み状況の表示を受け持つ。
' 【業務背景】HTML 版の「見込」タブと同じ。
'   見込生産の品目は、受注を待たずに「これだけ売れそう」という見込で作って
'   在庫を持つ。受注が入ると見込を消し込み（受注と見込の多い方を需要にする）、
'   出荷すると見込から差し引く（段階3）。受注生産の品目は見込を持たない。
' 【前提】入力_見込 シートは初期化で作る（Pr_入力見込シート作成）。
'==============================================================

' ---- 入力_見込 シートの配置 ----
Private Const St_登録品番セル As String = "C6"      ' 見込登録：品目
Private Const St_登録数量セル As String = "C7"      ' 見込登録：数量
Private Const St_期日セル As String = "C8"          ' 見込登録：いつまでに（今日から何日後）
Private Const St_対象IDセル As String = "C11"       ' 見込の操作：見込ID
Private Const St_変更数量セル As String = "C12"     ' 見込の操作：新しい数量
Private Const Lg_表行 As Long = 15                  ' 消し込み状況の表を書く行

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_見込 シート
Private Ws_見込 As Worksheet            ' 見込シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Dt_今日 As Date                 ' 今日
Private St_登録品番 As String           ' C6
Private Va_登録数量 As Variant          ' C7
Private Va_期日日数 As Variant          ' C8
Private St_対象ID As String             ' C11
Private Va_変更数量 As Variant          ' C12

'==============================================================
' Pr_見込登録（メイン）  ボタン：入力_見込「見込を登録」
'==============================================================
Public Sub Pr_見込登録()
    Dim Lg_見込ID As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力見込, St_シート_見込)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Call Pr_登録チェック                                    ' 1. 見込生産の品目か・数量・期日を確認（誤った見込を作らない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2. 画面更新を止める（速くする）
    Lg_見込ID = Fn_見込追加()                               ' 3. 見込を登録（MRP の需要に載せる）
    Call Pr_登録結果表示(Lg_見込ID)                         ' 4. 登録したIDを結果欄に出す（次の操作に使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_見込数量変更（メイン）  ボタン：入力_見込「数量を変更」
'==============================================================
Public Sub Pr_見込数量変更()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力見込, St_シート_見込)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象見込行()                                 ' 1. 見込ID の行を探す（操作する見込を決める）→無ければ中断
    Call Pr_変更チェック(Lg_行)                             ' 2. 新しい数量を確認（誤った数量にしない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_数量書換(Lg_行)                                 ' 3. 数量を書き換える（見込を見直す）
    Call Pr_操作結果表示("数量を変更しました")              ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_見込削除（メイン）  ボタン：入力_見込「見込を削除」
'==============================================================
Public Sub Pr_見込削除()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力見込, St_シート_見込)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象見込行()                                 ' 1. 見込ID の行を探す（操作する見込を決める）→無ければ中断
    Call Pr_削除前確認                                      ' 2. 削除してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_見込行削除(Lg_行)                               ' 3. 見込の行を消す（需要から外す）
    Call Pr_操作結果表示("削除しました")                    ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力見込)
    Set Ws_見込 = ThisWorkbook.Worksheets(St_シート_見込)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_今日 = Date
    St_登録品番 = Fn_文字(Ws_入力.Range(St_登録品番セル).Value)
    Va_登録数量 = Ws_入力.Range(St_登録数量セル).Value
    Va_期日日数 = Ws_入力.Range(St_期日セル).Value
    St_対象ID = Fn_文字(Ws_入力.Range(St_対象IDセル).Value)
    Va_変更数量 = Ws_入力.Range(St_変更数量セル).Value
End Sub

' 見込登録の入力を確かめる（見込生産の加工品・完成品、数量、期日 0〜365日）
Private Sub Pr_登録チェック()
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("品目「" & St_登録品番 & "」が " & Fn_事業部名(St_事業部) & " の品目にありません。")
    If Not Fn_製造品か(Fn_文字(Va_品目(列_品目.区分))) Or Fn_文字(Va_品目(列_品目.区分)) = "中間品" _
       Or Fn_文字(Va_品目(列_品目.生産方式)) <> St_見込生産 Then
        Call Pr_入力エラー("見込を登録できるのは、見込生産の加工品・完成品だけです。")
    End If
    If Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)) = 0 Then Call Pr_入力エラー(St_数量エラー)
    If Fn_日数変換(Va_期日日数, 0, 365) < 0 Then Call Pr_入力エラー("日数は0〜365の整数で入力してください。")
End Sub

' 見込を1件追加する。戻り値：見込ID
Private Function Fn_見込追加() As Long
    Dim Lg_見込ID As Long
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    Lg_見込ID = Fn_見込ID採番(St_事業部)
    Call Fn_行追記(Ws_見込, Array(St_事業部, Lg_見込ID, St_登録品番, Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)), _
        Dt_今日 + Fn_日数変換(Va_期日日数, 0, 365)))
    Fn_見込追加 = Lg_見込ID
End Function

' 見込ID の行を探す。戻り値：見込シートの行番号。空・見つからなければ中断
Private Function Fn_対象見込行() As Long
    Dim Lg_行 As Long
    If St_対象ID = "" Then Call Pr_入力エラー("操作する見込ID（数字）を入力してください。")
    Lg_行 = Fn_行検索(Ws_見込, 列_見込.次列 - 1, St_事業部, 列_見込.見込ID, St_対象ID)
    If Lg_行 = 0 Then Call Pr_入力エラー("見込ID " & St_対象ID & " は " & Fn_事業部名(St_事業部) & " にありません。")
    Fn_対象見込行 = Lg_行
End Function

' 新しい数量を確かめる
Private Sub Pr_変更チェック(ByVal Lg_行 As Long)
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Ws_見込.Cells(Lg_行, 列_見込.品番).Value))
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("見込の品目が品目マスタにありません。")
    If Fn_数量変換(Va_変更数量, Fn_小数可か(Va_品目)) = 0 Then Call Pr_入力エラー(St_数量エラー)
End Sub

' 数量を書き換える
Private Sub Pr_数量書換(ByVal Lg_行 As Long)
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Ws_見込.Cells(Lg_行, 列_見込.品番).Value))
    Ws_見込.Cells(Lg_行, 列_見込.数量).Value = Fn_数量変換(Va_変更数量, Fn_小数可か(Va_品目))
End Sub

' 削除してよいか確認する
Private Sub Pr_削除前確認()
    Call Pr_削除確認("見込ID " & St_対象ID & " を削除します（見込シートの行を削除します）。")
End Sub

' 見込の行を消す
Private Sub Pr_見込行削除(ByVal Lg_行 As Long)
    Call Pr_行削除(Ws_見込, Array(Lg_行))
End Sub

' 結果欄への表示
Private Sub Pr_登録結果表示(ByVal Lg_見込ID As Long)
    Call Pr_結果表示(Ws_入力, "見込ID " & Lg_見込ID & " を登録しました")
    Ws_入力.Range(St_対象IDセル).Value = Lg_見込ID
End Sub

Private Sub Pr_操作結果表示(ByVal St_文 As String)
    Call Pr_結果表示(Ws_入力, "見込ID " & St_対象ID & "：" & St_文)
End Sub

' 品目ごとの消し込み状況の表と、見込の一覧を書き直す
'   見込＝その品目の見込の合計、受注残＝出荷済でない受注の残数、未消込＝見込－受注残（0未満は0）
Private Sub Pr_一覧更新()
    Dim Va_品目 As Variant
    Dim Va_受注 As Variant
    Dim Va_見込 As Variant
    Dim Lg_行 As Long
    Dim Lg_子 As Long
    Dim Lg_出力行 As Long
    Dim St_品番 As String
    Dim Db_見込 As Double
    Dim Db_受注残 As Double
    Dim Bl_受注生産 As Boolean
    Va_品目 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_品目), 列_品目.次列 - 1)
    Va_受注 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_受注), 列_受注.次列 - 1)
    Va_見込 = Fn_表読込(Ws_見込, 列_見込.次列 - 1)
    Ws_入力.Range(Ws_入力.Cells(Lg_表行, 1), Ws_入力.Cells(Ws_入力.Rows.Count, 12)).Clear
    Ws_入力.Cells(Lg_表行, 2).Value = "■ 品目ごとの生産方式と見込の消し込み（" & Fn_事業部名(St_事業部) & "）"
    Ws_入力.Cells(Lg_表行, 2).Font.Bold = True
    Ws_入力.Range(Ws_入力.Cells(Lg_表行 + 1, 2), Ws_入力.Cells(Lg_表行 + 1, 7)).Value = Array("品番", "品名", "方式", "見込", "受注残", "未消込の見込")
    With Ws_入力.Range(Ws_入力.Cells(Lg_表行 + 1, 2), Ws_入力.Cells(Lg_表行 + 1, 7))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    Lg_出力行 = Lg_表行 + 2
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_事業部 And Fn_製造品か(Fn_文字(Va_品目(Lg_行, 列_品目.区分))) _
           And Fn_文字(Va_品目(Lg_行, 列_品目.区分)) <> "中間品" Then
            St_品番 = Fn_文字(Va_品目(Lg_行, 列_品目.品番))
            Bl_受注生産 = (Fn_文字(Va_品目(Lg_行, 列_品目.生産方式)) = St_受注生産)
            Db_見込 = 0: Db_受注残 = 0
            For Lg_子 = 1 To Fn_行数(Va_見込)
                If Fn_文字(Va_見込(Lg_子, 列_見込.事業部)) = St_事業部 And Fn_文字(Va_見込(Lg_子, 列_見込.品番)) = St_品番 Then
                    Db_見込 = Db_見込 + Fn_数(Va_見込(Lg_子, 列_見込.数量))
                End If
            Next Lg_子
            For Lg_子 = 1 To Fn_行数(Va_受注)
                If Fn_文字(Va_受注(Lg_子, 列_受注.事業部)) = St_事業部 And Fn_文字(Va_受注(Lg_子, 列_受注.品番)) = St_品番 _
                   And Fn_文字(Va_受注(Lg_子, 列_受注.状態)) <> St_状態_出荷済 Then
                    Db_受注残 = Db_受注残 + Fn_数(Va_受注(Lg_子, 列_受注.数量)) - Fn_数(Va_受注(Lg_子, 列_受注.出荷済))
                End If
            Next Lg_子
            Ws_入力.Cells(Lg_出力行, 2).Value = St_品番
            Ws_入力.Cells(Lg_出力行, 3).Value = Va_品目(Lg_行, 列_品目.品名)
            Ws_入力.Cells(Lg_出力行, 4).Value = Va_品目(Lg_行, 列_品目.生産方式)
            Ws_入力.Cells(Lg_出力行, 6).Value = Fn_丸め(Db_受注残)
            If Bl_受注生産 Then
                Ws_入力.Cells(Lg_出力行, 5).Value = "－"
                Ws_入力.Cells(Lg_出力行, 7).Value = "－"
            Else
                Ws_入力.Cells(Lg_出力行, 5).Value = Fn_丸め(Db_見込)
                Ws_入力.Cells(Lg_出力行, 7).Value = IIf(Db_見込 - Db_受注残 > 0, Fn_丸め(Db_見込 - Db_受注残), 0)
            End If
            Lg_出力行 = Lg_出力行 + 1
        End If
    Next Lg_行
    Call Pr_一覧転記(Ws_入力, Lg_出力行 + 1, St_シート_見込, 列_見込.次列 - 1, St_事業部, "見込一覧")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力見込再表示()
    If Not Fn_シートあり(St_シート_入力見込) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_見込 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力見込シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力見込)
    Call Pr_入力シート枠(Ws_表, "見込（フォーキャスト）", _
        "見込生産の品目について、いつまでにどれだけ売れそうかを入れます。受注が入ると見込を消し込み、出荷すると見込から差し引きます。")
    Call Pr_入力見出し(Ws_表, "B5", "見込登録")
    Call Pr_入力欄(Ws_表, St_登録品番セル, "品目（見込生産の品番）", "", "選択肢_見込品")
    Call Pr_入力欄(Ws_表, St_登録数量セル, "見込数量", 20, "")
    Call Pr_入力欄(Ws_表, St_期日セル, "いつまでに（今日から何日後）", 14, "")
    Call Pr_ボタン作成(Ws_表, "E6", 2, "見込を登録", "Pr_見込登録", "見込_登録")
    Call Pr_入力見出し(Ws_表, "B10", "見込の操作（見込ID を入れてボタンを押す）")
    Call Pr_入力欄(Ws_表, St_対象IDセル, "見込ID", "", "")
    Call Pr_入力欄(Ws_表, St_変更数量セル, "新しい数量（変更のとき）", "", "")
    Call Pr_ボタン作成(Ws_表, "E11", 2, "数量を変更", "Pr_見込数量変更", "見込_変更")
    Call Pr_ボタン作成(Ws_表, "G11", 2, "見込を削除", "Pr_見込削除", "見込_削除")
End Sub
