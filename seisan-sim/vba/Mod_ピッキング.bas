Attribute VB_Name = "Mod_ピッキング"
Option Explicit
'==============================================================
' Mod_ピッキング
' 【目的】入力_ピッキング シートの操作（読込・全部済にする・確定・取消）を受け持つ。
' 【業務背景】HTML 版の「ピッキング」タブと同じ。ピッキングは倉庫の棚から品物を集める作業。
'   ・明細は棚番の順に並べる（その順に倉庫を回れば集め終わる）。
'   ・集めた行の「済」に○を付ける。全部の行が済になったら確定できる。
'   ・製造用：確定すると部材を払い出し（在庫が減る）、製造指図が「着手」になる。
'   ・出荷用：確定すると出荷を計上する（在庫が減り、出荷実績に載る）。
'   ・取消：ピッキング指示を消す。製造用なら製造指図は計画のまま（部材の確保が外れる）。
' 【前提】入力_ピッキング シートは初期化で作る（Pr_入力ピッキングシート作成）。
'==============================================================

' ---- 入力_ピッキング シートの配置 ----
Private Const St_対象セル As String = "C6"          ' ピッキングNo
Private Const St_表示中セル As String = "C8"        ' 明細を表示しているピッキングNo（確定のときの照合に使う）
Private Const St_内容セル As String = "D8"          ' 表示中のピッキングの内容
Private Const Lg_明細見出し行 As Long = 10          ' 明細の見出し行
Private Const Lg_明細開始行 As Long = 11            ' 明細の1行目
Private Const Lg_明細最大 As Long = 18              ' 明細を出す最大行数
Private Const Lg_一覧行 As Long = 31                ' ピッキング一覧を書く行

' 明細の表示の列（B 列から）
Private Enum 列_明細表示
    済 = 2              ' B：集めたら○
    行No                ' C
    棚番                ' D
    品番                ' E
    品名                ' F
    数量                ' G
    単位                ' H
    次列
End Enum

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_ピッキング シート
Private Ws_頭 As Worksheet              ' ピッキングシート
Private Ws_明細 As Worksheet            ' ピッキング明細シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Dt_今日 As Date                 ' 今日（出荷日）
Private St_対象 As String               ' C6
Private St_表示中 As String             ' C8

'==============================================================
' Pr_ピッキング読込（メイン）  ボタン：入力_ピッキング「読込」
'==============================================================
Public Sub Pr_ピッキング読込()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力ピッキング, St_シート_ピッキング)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象行()                                     ' 1. ピッキングNo の行を探す（どの指示か決める）→無ければ中断
    Call Pr_画面更新停止                                    ' 1b. 画面更新を止める（速くする）
    Call Pr_明細表示(Lg_行)                                 ' 2. 明細を棚番順に出す（その順に棚を回る）
    Call Pr_読込結果表示                                    ' 3. 結果欄に出す（次の操作を案内）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_ピッキング全部済（メイン）  ボタン：入力_ピッキング「全部済にする」
'==============================================================
Public Sub Pr_ピッキング全部済()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力ピッキング, St_シート_ピッキング)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象行()                                     ' 1. ピッキングNo の行を探す（どの指示か決める）→無ければ中断
    Call Pr_作業中チェック(Lg_行)                           ' 2. 作業中か確認（完了した指示は変えない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_全行済(Lg_行)                                   ' 3. 明細をすべて済にする（残りをまとめて集めた扱い）
    Call Pr_明細表示(Lg_行)                                 ' 4. 明細を出し直す（結果を見せる）
    Call Pr_操作結果表示("すべての行を済にしました。「確定」を押してください") ' 5. 結果欄に出す（次の操作を案内）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_ピッキング確定（メイン）  ボタン：入力_ピッキング「確定」
'==============================================================
Public Sub Pr_ピッキング確定()
    Dim Lg_行 As Long
    Dim St_結果 As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力ピッキング, St_シート_ピッキング, St_シート_ピッキング明細)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象行()                                     ' 1. ピッキングNo の行を探す（どの指示か決める）→無ければ中断
    Call Pr_作業中チェック(Lg_行)                           ' 2. 作業中か・表示中の明細と同じか確認（別の指示の印で確定しない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_済の保存(Lg_行)                                 ' 3. 画面の済の印を明細に保存する（作業の進みを残す）
    Call Pr_全済チェック(Lg_行)                             ' 4. 全行が済か確認（集め残しを出さない）→NGなら中断
    Call Pr_出荷在庫チェック(Lg_行)                         ' 5. 出荷用なら在庫が足りるか確認（在庫をマイナスにしない）→NGなら中断
    St_結果 = Fn_確定処理(Lg_行)                            ' 6. 製造用は払出して着手、出荷用は出荷を計上（在庫を動かす）
    Call Pr_明細表示(Lg_行)                                 ' 7. 明細を出し直す（完了を見せる）
    Call Pr_操作結果表示(St_結果)                           ' 8. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 9. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_ピッキング取消（メイン）  ボタン：入力_ピッキング「取消」
'==============================================================
Public Sub Pr_ピッキング取消()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力ピッキング, St_シート_ピッキング, St_シート_ピッキング明細)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象行()                                     ' 1. ピッキングNo の行を探す（どの指示か決める）→無ければ中断
    Call Pr_取消チェック(Lg_行)                             ' 2. 作業中か確認（完了した記録を消さない）→NGなら中断
    Call Pr_取消確認                                        ' 3. 削除してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_指示削除(Lg_行)                                 ' 4. 指図の紐付けを外して指示を消す（確保を外す）
    Call Pr_明細消去                                        ' 5. 表示中の明細を消す（消した指示を残さない）
    Call Pr_操作結果表示("ピッキング指示を取り消しました")  ' 6. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 7. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力ピッキング)
    Set Ws_頭 = ThisWorkbook.Worksheets(St_シート_ピッキング)
    Set Ws_明細 = ThisWorkbook.Worksheets(St_シート_ピッキング明細)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_今日 = Date
    St_対象 = Fn_文字(Ws_入力.Range(St_対象セル).Value)
    St_表示中 = Fn_文字(Ws_入力.Range(St_表示中セル).Value)
End Sub

' ピッキングNo の行を探す。戻り値：ピッキングシートの行。空・見つからなければ中断
Private Function Fn_対象行() As Long
    Dim Lg_行 As Long
    If St_対象 = "" Then Call Pr_入力エラー("ピッキングNo（例：PK-0001）を入力してください。")
    Lg_行 = Fn_行検索(Ws_頭, 列_ピッキング.次列 - 1, St_事業部, 列_ピッキング.ピッキングNo, St_対象)
    If Lg_行 = 0 Then Call Pr_入力エラー("ピッキング " & St_対象 & " は " & Fn_事業部名(St_事業部) & " にありません。")
    Fn_対象行 = Lg_行
End Function

' 明細のシート行を棚番の順に返す（棚番が空のものは最後。HTML 版の (loc||'~') の並び）
'   戻り値：行番号の配列（1始まり）。明細が無ければ Empty
Private Function Fn_明細行() As Variant
    Dim Va_表 As Variant
    Dim Lg_結果() As Long
    Dim Lg_件数 As Long
    Dim Lg_行 As Long
    Dim Lg_添字 As Long
    Dim Lg_一時 As Long
    Va_表 = Fn_表読込(Ws_明細, 列_ピッキング明細.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_ピッキング明細.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_ピッキング明細.ピッキングNo)) = St_対象 Then
            Lg_件数 = Lg_件数 + 1
            ReDim Preserve Lg_結果(1 To Lg_件数)
            Lg_結果(Lg_件数) = Lg_行
        End If
    Next Lg_行
    If Lg_件数 = 0 Then
        Fn_明細行 = Empty
        Exit Function
    End If
    ' 棚番で並べる（挿入ソート。同じ棚番は明細の順）
    For Lg_行 = 2 To Lg_件数
        Lg_一時 = Lg_結果(Lg_行)
        Lg_添字 = Lg_行 - 1
        Do While Lg_添字 >= 1
            If StrComp(Fn_棚キー(Va_表(Lg_結果(Lg_添字), 列_ピッキング明細.棚番)), Fn_棚キー(Va_表(Lg_一時, 列_ピッキング明細.棚番)), vbBinaryCompare) <= 0 Then Exit Do
            Lg_結果(Lg_添字 + 1) = Lg_結果(Lg_添字)
            Lg_添字 = Lg_添字 - 1
        Loop
        Lg_結果(Lg_添字 + 1) = Lg_一時
    Next Lg_行
    For Lg_行 = 1 To Lg_件数
        Lg_結果(Lg_行) = Lg_結果(Lg_行) + Lg_開始行 - 1
    Next Lg_行
    Fn_明細行 = Lg_結果
End Function

' 並べ替え用の棚番（空は最後に来るよう "~"）
Private Function Fn_棚キー(ByVal Va_棚番 As Variant) As String
    If Fn_文字(Va_棚番) = "" Then Fn_棚キー = "~" Else Fn_棚キー = Fn_文字(Va_棚番)
End Function

' 明細を棚番順に表示する
Private Sub Pr_明細表示(ByVal Lg_行 As Long)
    Dim Va_行 As Variant
    Dim Va_品目 As Variant
    Dim Lg_添字 As Long
    Dim Lg_出力 As Long
    Call Pr_明細消去
    Ws_入力.Range(St_表示中セル).Value = St_対象
    Ws_入力.Range(St_内容セル).Value = Ws_頭.Cells(Lg_行, 列_ピッキング.種別).Value & "用　対象 " & Ws_頭.Cells(Lg_行, 列_ピッキング.対象伝票).Value & _
        "　" & Ws_頭.Cells(Lg_行, 列_ピッキング.内容).Value & IIf(Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.配送方法).Value) = "", "", "　" & Ws_頭.Cells(Lg_行, 列_ピッキング.配送方法).Value) & _
        "　状態：" & Ws_頭.Cells(Lg_行, 列_ピッキング.状態).Value
    Va_行 = Fn_明細行()
    If IsEmpty(Va_行) Then Exit Sub
    For Lg_添字 = 1 To UBound(Va_行)
        If Lg_添字 > Lg_明細最大 Then Exit For
        Lg_出力 = Lg_明細開始行 + Lg_添字 - 1
        Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.品番).Value))
        Ws_入力.Cells(Lg_出力, 列_明細表示.済).Value = Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.済).Value
        Ws_入力.Cells(Lg_出力, 列_明細表示.行No).Value = Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.行No).Value
        Ws_入力.Cells(Lg_出力, 列_明細表示.棚番).Value = IIf(Fn_文字(Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.棚番).Value) = "", "未設定", Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.棚番).Value)
        Ws_入力.Cells(Lg_出力, 列_明細表示.品番).Value = Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.品番).Value
        If Not IsEmpty(Va_品目) Then
            Ws_入力.Cells(Lg_出力, 列_明細表示.品名).Value = Va_品目(列_品目.品名)
            Ws_入力.Cells(Lg_出力, 列_明細表示.単位).Value = Va_品目(列_品目.単位)
        End If
        Ws_入力.Cells(Lg_出力, 列_明細表示.数量).Value = Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.数量).Value
        Ws_入力.Cells(Lg_出力, 列_明細表示.済).Interior.Color = RGB(255, 255, 204)
        Ws_入力.Cells(Lg_出力, 列_明細表示.済).HorizontalAlignment = xlCenter
    Next Lg_添字
End Sub

' 表示中の明細を消す
Private Sub Pr_明細消去()
    Ws_入力.Range(Ws_入力.Cells(Lg_明細開始行, 列_明細表示.済), Ws_入力.Cells(Lg_明細開始行 + Lg_明細最大 - 1, 列_明細表示.単位)).Clear
    Ws_入力.Range(St_表示中セル).Value = ""
    Ws_入力.Range(St_内容セル).Value = ""
End Sub

' 作業中か確かめる
Private Sub Pr_作業中チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.状態).Value) <> St_状態_作業中 Then
        Call Pr_入力エラー("ピッキング " & St_対象 & " は完了しています。")
    End If
End Sub

' すべての明細を済にする
Private Sub Pr_全行済(ByVal Lg_行 As Long)
    Dim Va_行 As Variant
    Dim Lg_添字 As Long
    Va_行 = Fn_明細行()
    If IsEmpty(Va_行) Then Exit Sub
    For Lg_添字 = 1 To UBound(Va_行)
        Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.済).Value = St_印
    Next Lg_添字
End Sub

' 画面の済の印を明細に保存する（表示中のピッキングと同じときだけ）
Private Sub Pr_済の保存(ByVal Lg_行 As Long)
    Dim Va_行 As Variant
    Dim Lg_添字 As Long
    Dim Lg_出力 As Long
    If St_表示中 <> St_対象 Then
        Call Pr_入力エラー("表示している明細が " & St_対象 & " ではありません。先に「読込」を押して、集めた行の「済」に○を付けてください。")
    End If
    Va_行 = Fn_明細行()
    If IsEmpty(Va_行) Then Exit Sub
    For Lg_添字 = 1 To UBound(Va_行)
        If Lg_添字 > Lg_明細最大 Then Exit For
        Lg_出力 = Lg_明細開始行 + Lg_添字 - 1
        If Fn_文字(Ws_入力.Cells(Lg_出力, 列_明細表示.済).Value) <> "" Then
            Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.済).Value = St_印
        Else
            Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.済).Value = ""
        End If
    Next Lg_添字
End Sub

' 全行が済か確かめる（済の印は保存済み）
Private Sub Pr_全済チェック(ByVal Lg_行 As Long)
    Dim Va_行 As Variant
    Dim Lg_添字 As Long
    Dim Lg_未 As Long
    Va_行 = Fn_明細行()
    If IsEmpty(Va_行) Then Exit Sub
    For Lg_添字 = 1 To UBound(Va_行)
        If Fn_文字(Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.済).Value) = "" Then Lg_未 = Lg_未 + 1
    Next Lg_添字
    If Lg_未 > 0 Then
        Call Pr_明細表示(Lg_行)
        Call Pr_入力エラー("まだ集めていない行が " & Lg_未 & " 行あります（済の印は保存しました）。全部集めてから確定してください。")
    End If
End Sub

' 出荷用なら、出荷数が今の在庫以下か確かめる
Private Sub Pr_出荷在庫チェック(ByVal Lg_行 As Long)
    Dim Lg_受注行 As Long
    Dim Va_品目 As Variant
    Dim Db_数量 As Double
    If Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.種別).Value) <> "出荷" Then Exit Sub
    Lg_受注行 = Fn_行検索(ThisWorkbook.Worksheets(St_シート_受注), 列_受注.次列 - 1, St_事業部, 列_受注.受注No, Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.対象伝票).Value))
    If Lg_受注行 = 0 Then Call Pr_入力エラー("受注が見つかりません。")
    Va_品目 = Fn_品目情報(St_事業部, Fn_文字(ThisWorkbook.Worksheets(St_シート_受注).Cells(Lg_受注行, 列_受注.品番).Value))
    Db_数量 = Fn_出荷数(Lg_行)
    If Db_数量 > Fn_数(Va_品目(列_品目.在庫数)) Then
        Call Pr_入力エラー("在庫が足りません（現在庫 " & Fn_表示数(Fn_数(Va_品目(列_品目.在庫数))) & "）。")
    End If
End Sub

' 出荷用ピッキングの出荷数（明細の1行目の数量）
Private Function Fn_出荷数(ByVal Lg_行 As Long) As Double
    Dim Va_行 As Variant
    Va_行 = Fn_明細行()
    Fn_出荷数 = Fn_数(Ws_明細.Cells(Va_行(1), 列_ピッキング明細.数量).Value)
End Function

' 確定の処理。戻り値：結果の文
'   製造用：明細ごとに払い出し、指示を完了、製造指図を着手にしてピッキングNoを外す
'   出荷用：指示を完了、出荷を計上して出荷Noを指示に残す
Private Function Fn_確定処理(ByVal Lg_行 As Long) As String
    Dim Ws_指図 As Worksheet
    Dim Va_行 As Variant
    Dim Lg_添字 As Long
    Dim Lg_対象行 As Long
    Dim St_対象伝票 As String
    Dim St_出荷No As String
    St_対象伝票 = Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.対象伝票).Value)
    If Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.種別).Value) = "製造" Then
        Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
        Lg_対象行 = Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_対象伝票)
        If Lg_対象行 = 0 Then Call Pr_入力エラー("製造指図 " & St_対象伝票 & " が見つかりません。")
        Va_行 = Fn_明細行()
        For Lg_添字 = 1 To UBound(Va_行)
            Call Pr_在庫移動(St_事業部, Fn_文字(Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.品番).Value), "払出", _
                -Fn_数(Ws_明細.Cells(Va_行(Lg_添字), 列_ピッキング明細.数量).Value), St_対象伝票)
        Next Lg_添字
        Ws_頭.Cells(Lg_行, 列_ピッキング.状態).Value = St_状態_完了
        Ws_指図.Cells(Lg_対象行, 列_製造指図.ピッキングNo).Value = ""
        Ws_指図.Cells(Lg_対象行, 列_製造指図.状態).Value = St_状態_着手
        Fn_確定処理 = "部材を払い出しました。製造指図 " & St_対象伝票 & " は着手になりました"
    Else
        Lg_対象行 = Fn_行検索(ThisWorkbook.Worksheets(St_シート_受注), 列_受注.次列 - 1, St_事業部, 列_受注.受注No, St_対象伝票)
        Ws_頭.Cells(Lg_行, 列_ピッキング.状態).Value = St_状態_完了
        St_出荷No = Fn_出荷計上(St_事業部, Lg_対象行, Fn_出荷数(Lg_行), Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.配送方法).Value), Dt_今日)
        Ws_頭.Cells(Lg_行, 列_ピッキング.出荷No).Value = St_出荷No
        Fn_確定処理 = "出荷しました（出荷No " & St_出荷No & "）"
    End If
End Function

' 取り消せるか（作業中だけ）
Private Sub Pr_取消チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.状態).Value) <> St_状態_作業中 Then
        Call Pr_入力エラー("完了したピッキングは取り消せません。")
    End If
End Sub

' 取り消してよいか確認する
Private Sub Pr_取消確認()
    Call Pr_削除確認("ピッキング指示 " & St_対象 & " を取り消します（ピッキング・明細の行を削除します）。")
End Sub

' 製造用なら製造指図のピッキングNoを外し、指示と明細を消す
Private Sub Pr_指示削除(ByVal Lg_行 As Long)
    Dim Ws_指図 As Worksheet
    Dim Lg_指図行 As Long
    If Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.種別).Value) = "製造" Then
        Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
        Lg_指図行 = Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, Fn_文字(Ws_頭.Cells(Lg_行, 列_ピッキング.対象伝票).Value))
        If Lg_指図行 > 0 Then Ws_指図.Cells(Lg_指図行, 列_製造指図.ピッキングNo).Value = ""
    End If
    Call Pr_ピッキング削除(St_事業部, St_対象)
End Sub

' 結果欄への表示
Private Sub Pr_読込結果表示()
    Call Pr_結果表示(Ws_入力, "ピッキング " & St_対象 & " を読み込みました。集めた行の「済」に○を付けて「確定」を押します")
End Sub

Private Sub Pr_操作結果表示(ByVal St_文 As String)
    Call Pr_結果表示(Ws_入力, "ピッキング " & St_対象 & "：" & St_文)
End Sub

' ピッキング一覧を書き直す
Private Sub Pr_一覧更新()
    Call Pr_一覧転記(Ws_入力, Lg_一覧行, St_シート_ピッキング, 列_ピッキング.次列 - 1, St_事業部, "ピッキング一覧")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力ピッキング再表示()
    If Not Fn_シートあり(St_シート_入力ピッキング) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_ピッキング シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力ピッキングシート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力ピッキング)
    Call Pr_入力シート枠(Ws_表, "ピッキング", _
        "倉庫の棚から品物を集める作業です。明細は棚番の順。集めた行の「済」に○を付けて確定すると、製造用は払出（着手）、出荷用は出荷になります。")
    Ws_表.Columns("B").ColumnWidth = 8
    Ws_表.Columns("F").ColumnWidth = 18
    Call Pr_入力見出し(Ws_表, "B5", "ピッキングNo を入れて「読込」")
    Call Pr_入力欄(Ws_表, St_対象セル, "No", "", "")
    Call Pr_ボタン作成(Ws_表, "E6", 1, "読込", "Pr_ピッキング読込", "ピッキング_読込")
    Call Pr_ボタン作成(Ws_表, "F6", 1, "全部済にする", "Pr_ピッキング全部済", "ピッキング_全部済")
    Call Pr_ボタン作成(Ws_表, "G6", 1, "確定", "Pr_ピッキング確定", "ピッキング_確定")
    Call Pr_ボタン作成(Ws_表, "I6", 1, "取消", "Pr_ピッキング取消", "ピッキング_取消")
    Ws_表.Range("B8").Value = "表示中"
    Ws_表.Range(Ws_表.Cells(Lg_明細見出し行, 列_明細表示.済), Ws_表.Cells(Lg_明細見出し行, 列_明細表示.単位)).Value = _
        Array("済", "行No", "棚番", "品番", "品名", "数量", "単位")
    With Ws_表.Range(Ws_表.Cells(Lg_明細見出し行, 列_明細表示.済), Ws_表.Cells(Lg_明細見出し行, 列_明細表示.単位))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
End Sub
