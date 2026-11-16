Attribute VB_Name = "Mod_製造"
Option Explicit
'==============================================================
' Mod_製造
' 【目的】入力_製造 シートの操作を受け持つ。
'         段階2：製造指図の作成（手動）、着手日の変更、製造指図の取消。
'         段階3：ピッキング指示（部材を確保）、部材なしの着手、完成の計上。
'         （手配No.で行う工程実績・外注に出す は Mod_製造実績）
' 【業務背景】HTML 版の「製造」タブ・「生産計画」タブと同じ。
'   ・完成予定日は、工程表の標準時間から自動で決まる（着手日＋所要日数）。
'   ・着手日を変えられるのは「計画」の指図だけ。完成予定も所要日数でずれる。
'   ・取り消せるのは「計画」の指図だけ。ピッキング指示が出ていればそれも消す。
'   ・ピッキング指示は「計画」の指図に出す。BOM の部材がすべて有効在庫で足りるときだけ。
'     出した時点で部材は「製造用確保」になり、ほかに回らない。払出は入力_ピッキングで確定する。
'   ・払い出す部材が無い品目（BOM が空）は、そのまま着手できる。
'   ・工程が1つも無い指図は、着手後に「完成を計上」で指示数を入庫する。
' 【前提】入力_製造 シートは初期化で作る（Pr_入力製造シート作成）。
'==============================================================

' ---- 入力_製造 シートの配置 ----
Private Const St_登録品番セル As String = "C6"      ' 製造指図作成：品目
Private Const St_登録数量セル As String = "C7"      ' 製造指図作成：数量
Private Const St_登録着手セル As String = "C8"      ' 製造指図作成：着手予定（今日から何日後）
Private Const St_対象指図セル As String = "C11"     ' 指図の操作：指図No
Private Const St_変更着手セル As String = "C12"     ' 指図の操作：新しい着手予定（今日から何日後）
Private Const Lg_一覧行 As Long = 40                ' 製造指図一覧を書く行（上は工程の欄と工程一覧）

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_製造 シート
Private Ws_指図 As Worksheet            ' 製造指図シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Dt_今日 As Date                 ' 今日
Private St_登録品番 As String           ' C6
Private Va_登録数量 As Variant          ' C7
Private Va_登録着手 As Variant          ' C8
Private St_対象指図 As String           ' C11
Private Va_変更着手 As Variant          ' C12

'==============================================================
' Pr_製造指図作成（メイン）  ボタン：入力_製造「製造指図を作成」
'==============================================================
Public Sub Pr_製造指図作成()
    Dim St_指図No As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_製造指図)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Call Pr_作成チェック                                    ' 1. 製造品か・数量・着手日を確認（誤った指図を作らない）→NGなら中断
    Call Pr_マスタ点検(St_事業部)                           ' 2. マスタの誤りを点検（工程の複写を狂わせない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    St_指図No = Fn_指図追加()                               ' 3. 製造指図を作る（工程を複写し手配No.を振る）
    Call Pr_作成結果表示(St_指図No)                         ' 4. 作った指図を結果欄に出す（次の操作に使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_着手日変更（メイン）  ボタン：入力_製造「着手日を変更」
'==============================================================
Public Sub Pr_着手日変更()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_製造指図)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象指図行()                                 ' 1. 指図No の行を探す（操作する指図を決める）→無ければ中断
    Call Pr_日程変更チェック(Lg_行)                         ' 2. 計画中か・日数が0〜60か確認（着手済の日程は動かさない）→NGなら中断
    Call Pr_マスタ点検(St_事業部)                           ' 3. マスタの誤りを点検（所要日数を狂わせない）→NGなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_日程書換(Lg_行)                                 ' 4. 着手予定と完成予定を書き換える（負荷の山をならす）
    Call Pr_操作結果表示("着手日を変更しました")            ' 5. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 6. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_製造指図取消（メイン）  ボタン：入力_製造「製造指図を取り消す」
'==============================================================
Public Sub Pr_製造指図取消()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_製造指図)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象指図行()                                 ' 1. 指図No の行を探す（操作する指図を決める）→無ければ中断
    Call Pr_取消チェック(Lg_行)                             ' 2. 計画中か確認（着手済の実績を消さない）→NGなら中断
    Call Pr_取消確認                                        ' 3. 削除してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_指図と付随削除(Lg_行)                           ' 4. ピッキング・工程・指図の行を消す（確保した部材も放す）
    Call Pr_操作結果表示("製造指図を取り消しました")        ' 5. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 6. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_製造ピッキング指示（メイン）  ボタン：入力_製造「ピッキング指示を出す」
'==============================================================
Public Sub Pr_製造ピッキング指示()
    Dim Lg_行 As Long
    Dim St_ピッキングNo As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_製造指図, St_シート_ピッキング)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象指図行()                                 ' 1. 指図No の行を探す（操作する指図を決める）→無ければ中断
    Call Pr_ピッキング可否チェック(Lg_行)                   ' 2. 計画中・部材が足りるか確認（欠品のまま着手しない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    St_ピッキングNo = Fn_部材ピッキング作成(Lg_行)          ' 3. 部材のピッキング指示を作る（部材を確保して棚から集める）
    Call Pr_ピッキング結果表示(St_ピッキングNo)             ' 4. 番号を結果欄に出す（入力_ピッキングで使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_部材なし着手（メイン）  ボタン：入力_製造「着手する（部材なし）」
'==============================================================
Public Sub Pr_部材なし着手()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_製造指図)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象指図行()                                 ' 1. 指図No の行を探す（操作する指図を決める）→無ければ中断
    Call Pr_部材なし着手チェック(Lg_行)                     ' 2. 計画中で部材が無い品目か確認（払出を飛ばさない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_状態書換(Lg_行, St_状態_着手)                   ' 3. 状態を着手にする（工程の実績を入れられるようにする）
    Call Pr_操作結果表示("着手しました")                    ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_完成計上（メイン）  ボタン：入力_製造「完成を計上」
'==============================================================
Public Sub Pr_完成計上()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_製造指図, St_シート_指図工程)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象指図行()                                 ' 1. 指図No の行を探す（操作する指図を決める）→無ければ中断
    Call Pr_完成計上チェック(Lg_行)                         ' 2. 着手中で残りの工程が無いか確認（工程の実績を飛ばさない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_指示数で完成(Lg_行)                             ' 3. 指示数を完成として入庫する（製品の在庫にする）
    Call Pr_操作結果表示("完成を計上しました")              ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力製造)
    Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_今日 = Date
    St_登録品番 = Fn_文字(Ws_入力.Range(St_登録品番セル).Value)
    Va_登録数量 = Ws_入力.Range(St_登録数量セル).Value
    Va_登録着手 = Ws_入力.Range(St_登録着手セル).Value
    St_対象指図 = Fn_文字(Ws_入力.Range(St_対象指図セル).Value)
    Va_変更着手 = Ws_入力.Range(St_変更着手セル).Value
End Sub

' 製造指図作成の入力を確かめる（製造品、数量、着手 0〜60日後）
Private Sub Pr_作成チェック()
    Dim Va_品目 As Variant
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    If IsEmpty(Va_品目) Then Call Pr_入力エラー("品目「" & St_登録品番 & "」が " & Fn_事業部名(St_事業部) & " の品目にありません。")
    If Not Fn_製造品か(Fn_文字(Va_品目(列_品目.区分))) Then Call Pr_入力エラー("製造指図を作れるのは製造品（中間品・加工品・完成品）だけです。")
    If Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目)) = 0 Then Call Pr_入力エラー(St_数量エラー)
    If Fn_日数変換(Va_登録着手, 0, 60) < 0 Then Call Pr_入力エラー("着手予定は0〜60の整数（今日から何日後）で入力してください。")
End Sub

' 製造指図を作る（受注には紐付けない）。戻り値：指図No
Private Function Fn_指図追加() As String
    Dim Va_品目 As Variant
    Dim Db_数量 As Double
    Dim Lg_着手 As Long
    Va_品目 = Fn_品目情報(St_事業部, St_登録品番)
    Db_数量 = Fn_数量変換(Va_登録数量, Fn_小数可か(Va_品目))
    Lg_着手 = Fn_日数変換(Va_登録着手, 0, 60)
    Fn_指図追加 = Fn_製造指図作成(St_事業部, St_登録品番, Db_数量, Lg_着手, _
        Lg_着手 + Fn_所要日数(St_事業部, St_登録品番, Db_数量), "", Dt_今日)
End Function

' 指図No の行を探す。戻り値：製造指図シートの行番号。空・見つからなければ中断
Private Function Fn_対象指図行() As Long
    Dim Lg_行 As Long
    If St_対象指図 = "" Then Call Pr_入力エラー("操作する指図No（例：MO-0001）を入力してください。")
    Lg_行 = Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_対象指図)
    If Lg_行 = 0 Then Call Pr_入力エラー("製造指図 " & St_対象指図 & " は " & Fn_事業部名(St_事業部) & " にありません。")
    Fn_対象指図行 = Lg_行
End Function

' 着手日を変えられるか（計画中、日数 0〜60）
Private Sub Pr_日程変更チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value) <> St_状態_計画 Then
        Call Pr_入力エラー("着手日を変えられるのは「計画」の製造指図だけです。")
    End If
    If Fn_日数変換(Va_変更着手, 0, 60) < 0 Then Call Pr_入力エラー("新しい着手予定は0〜60の整数（今日から何日後）で入力してください。")
End Sub

' 着手予定と完成予定を書き換える（完成予定＝着手予定＋所要日数）
Private Sub Pr_日程書換(ByVal Lg_行 As Long)
    Dim Lg_着手 As Long
    Lg_着手 = Fn_日数変換(Va_変更着手, 0, 60)
    Ws_指図.Cells(Lg_行, 列_製造指図.着手予定).Value = Dt_今日 + Lg_着手
    Ws_指図.Cells(Lg_行, 列_製造指図.完成予定).Value = Dt_今日 + Lg_着手 + _
        Fn_所要日数(St_事業部, Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.品番).Value), Fn_数(Ws_指図.Cells(Lg_行, 列_製造指図.数量).Value))
End Sub

' 取り消せるか（計画中だけ）
Private Sub Pr_取消チェック(ByVal Lg_行 As Long)
    If Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value) <> St_状態_計画 Then
        Call Pr_入力エラー("取り消せるのは「計画」の製造指図だけです（着手後は実績があるため）。")
    End If
End Sub

' 取り消してよいか確認する
Private Sub Pr_取消確認()
    Call Pr_削除確認("製造指図 " & St_対象指図 & " を取り消します（指図・工程・ピッキング指示の行を削除します）。")
End Sub

' ピッキング指示（あれば）・工程・指図の行を消す
Private Sub Pr_指図と付随削除(ByVal Lg_行 As Long)
    Call Pr_ピッキング削除(St_事業部, Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.ピッキングNo).Value))
    Call Pr_製造指図削除(St_事業部, St_対象指図)
End Sub

' 指図の BOM の部材と必要数（親の数量×員数）を、BOM シートの並びで返す
'   引数：Lg_行（製造指図シートの行）、Va_品番（出力）、Va_数量（出力）
'   戻り値：部材の数（0なら払い出す部材なし）
Private Function Fn_部材一覧(ByVal Lg_行 As Long, ByRef Va_品番 As Variant, ByRef Va_数量 As Variant) As Long
    Dim Va_BOM As Variant
    Dim Lg_BOM行 As Long
    Dim Lg_件数 As Long
    Dim St_親 As String
    Dim Db_親数量 As Double
    Dim St_品番() As String
    Dim Db_数量() As Double
    St_親 = Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.品番).Value)
    Db_親数量 = Fn_数(Ws_指図.Cells(Lg_行, 列_製造指図.数量).Value)
    Va_BOM = Fn_表読込(ThisWorkbook.Worksheets(St_シート_BOM), 列_BOM.次列 - 1)
    For Lg_BOM行 = 1 To Fn_行数(Va_BOM)
        If Fn_文字(Va_BOM(Lg_BOM行, 列_BOM.事業部)) = St_事業部 And Fn_文字(Va_BOM(Lg_BOM行, 列_BOM.親品番)) = St_親 Then
            Lg_件数 = Lg_件数 + 1
            ReDim Preserve St_品番(1 To Lg_件数)
            ReDim Preserve Db_数量(1 To Lg_件数)
            St_品番(Lg_件数) = Fn_文字(Va_BOM(Lg_BOM行, 列_BOM.子品番))
            Db_数量(Lg_件数) = Fn_丸め(Fn_数(Va_BOM(Lg_BOM行, 列_BOM.員数)) * Db_親数量)
        End If
    Next Lg_BOM行
    If Lg_件数 > 0 Then
        Va_品番 = St_品番
        Va_数量 = Db_数量
    End If
    Fn_部材一覧 = Lg_件数
End Function

' ピッキング指示を出せるか（計画中、まだ出していない、部材がある、部材がすべて有効在庫で足りる）
Private Sub Pr_ピッキング可否チェック(ByVal Lg_行 As Long)
    Dim Va_品番 As Variant
    Dim Va_数量 As Variant
    Dim Lg_添字 As Long
    Dim Db_有効 As Double
    Dim St_不足 As String
    Dim Va_品目 As Variant
    If Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value) <> St_状態_計画 Then
        Call Pr_入力エラー("ピッキング指示を出せるのは「計画」の製造指図だけです。")
    End If
    If Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.ピッキングNo).Value) <> "" Then
        Call Pr_入力エラー("すでにピッキング指示 " & Ws_指図.Cells(Lg_行, 列_製造指図.ピッキングNo).Value & " が出ています。入力_ピッキングで確定してください。")
    End If
    If Fn_部材一覧(Lg_行, Va_品番, Va_数量) = 0 Then
        Call Pr_入力エラー("払い出す部材がありません（BOM が空）。「着手する（部材なし）」を使ってください。")
    End If
    For Lg_添字 = LBound(Va_品番) To UBound(Va_品番)
        Db_有効 = Fn_有効在庫(St_事業部, CStr(Va_品番(Lg_添字)))
        If Db_有効 < Va_数量(Lg_添字) Then
            Va_品目 = Fn_品目情報(St_事業部, CStr(Va_品番(Lg_添字)))
            St_不足 = St_不足 & vbCrLf & "・" & Va_品目(列_品目.品名) & "　必要 " & Fn_表示数(CDbl(Va_数量(Lg_添字))) & " ／ 有効 " & Fn_表示数(Db_有効)
        End If
    Next Lg_添字
    If St_不足 <> "" Then Call Pr_入力エラー("部材が足りません：" & St_不足)
End Sub

' 部材のピッキング指示を作り、指図にピッキングNo を書く。戻り値：ピッキングNo
'   内容は HTML 版と同じ「品名 ×数量 の部材」。
Private Function Fn_部材ピッキング作成(ByVal Lg_行 As Long) As String
    Dim Va_品番 As Variant
    Dim Va_数量 As Variant
    Dim Va_品目 As Variant
    Dim St_番号 As String
    Call Fn_部材一覧(Lg_行, Va_品番, Va_数量)
    Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.品番).Value))
    St_番号 = Fn_ピッキング作成(St_事業部, "製造", St_対象指図, _
        Va_品目(列_品目.品名) & " ×" & Fn_表示数(Fn_数(Ws_指図.Cells(Lg_行, 列_製造指図.数量).Value)) & " の部材", Va_品番, Va_数量, "")
    Ws_指図.Cells(Lg_行, 列_製造指図.ピッキングNo).Value = St_番号
    Fn_部材ピッキング作成 = St_番号
End Function

' 部材なしで着手できるか（計画中で、BOM が空）
Private Sub Pr_部材なし着手チェック(ByVal Lg_行 As Long)
    Dim Va_品番 As Variant
    Dim Va_数量 As Variant
    If Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value) <> St_状態_計画 Then
        Call Pr_入力エラー("着手できるのは「計画」の製造指図だけです。")
    End If
    If Fn_部材一覧(Lg_行, Va_品番, Va_数量) > 0 Then
        Call Pr_入力エラー("この品目には部材があります。「ピッキング指示を出す」で部材を払い出すと着手になります。")
    End If
End Sub

' 製造指図の状態を書き換える
Private Sub Pr_状態書換(ByVal Lg_行 As Long, ByVal St_状態 As String)
    Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value = St_状態
End Sub

' 完成を計上できるか（着手中で、完了していない工程が無い）
Private Sub Pr_完成計上チェック(ByVal Lg_行 As Long)
    Dim Va_工程 As Variant
    Dim Lg_工程 As Long
    If Fn_文字(Ws_指図.Cells(Lg_行, 列_製造指図.状態).Value) <> St_状態_着手 Then
        Call Pr_入力エラー("完成を計上できるのは「着手」の製造指図だけです。")
    End If
    Va_工程 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_指図工程), 列_指図工程.次列 - 1)
    For Lg_工程 = 1 To Fn_行数(Va_工程)
        If Fn_文字(Va_工程(Lg_工程, 列_指図工程.事業部)) = St_事業部 And Fn_文字(Va_工程(Lg_工程, 列_指図工程.指図No)) = St_対象指図 _
           And Fn_文字(Va_工程(Lg_工程, 列_指図工程.状態)) <> St_状態_完了 Then
            Call Pr_入力エラー("まだ終わっていない工程があります。工程の実績は手配No.で入力してください（最終工程を入れると自動で完成します）。")
        End If
    Next Lg_工程
End Sub

' 指示数を完成として入庫する（工程の無い指図。HTML 版の「完成を計上する」）
Private Sub Pr_指示数で完成(ByVal Lg_行 As Long)
    Call Pr_工程完了後処理(St_事業部, St_対象指図, True, Fn_数(Ws_指図.Cells(Lg_行, 列_製造指図.数量).Value))
End Sub

' 結果欄への表示
Private Sub Pr_ピッキング結果表示(ByVal St_番号 As String)
    Call Pr_結果表示(Ws_入力, "製造指図 " & St_対象指図 & "：ピッキング指示 " & St_番号 & " を出しました。入力_ピッキングで集めて確定すると着手になります")
End Sub

Private Sub Pr_作成結果表示(ByVal St_指図No As String)
    Call Pr_結果表示(Ws_入力, "製造指図 " & St_指図No & " を作成しました（計画）")
    Ws_入力.Range(St_対象指図セル).Value = St_指図No
End Sub

Private Sub Pr_操作結果表示(ByVal St_文 As String)
    Call Pr_結果表示(Ws_入力, "製造指図 " & St_対象指図 & "：" & St_文)
End Sub

' 製造指図一覧を書き直す
Private Sub Pr_一覧更新()
    Call Pr_一覧転記(Ws_入力, Lg_一覧行, St_シート_製造指図, 列_製造指図.次列 - 1, St_事業部, "製造指図一覧")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力製造再表示()
    If Not Fn_シートあり(St_シート_入力製造) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_製造 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力製造シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力製造)
    Call Pr_入力シート枠(Ws_表, "製造", _
        "計画 → ピッキング指示 → 払出（着手）→ 工程ごとの実績 → 完成。完成予定日は工程表の標準時間から自動で決まります。")
    Call Pr_入力見出し(Ws_表, "B5", "製造指図作成（受注に紐付けない。受注生産の受注向けは入力_受注から）")
    Call Pr_入力欄(Ws_表, St_登録品番セル, "製造する品目（品番）", "", "選択肢_製造品")
    Call Pr_入力欄(Ws_表, St_登録数量セル, "数量", 10, "")
    Call Pr_入力欄(Ws_表, St_登録着手セル, "着手予定（今日から何日後）", 0, "")
    Call Pr_ボタン作成(Ws_表, "E6", 2, "製造指図を作成", "Pr_製造指図作成", "製造_作成")
    Call Pr_入力見出し(Ws_表, "B10", "指図の操作（指図No を入れてボタンを押す）")
    Call Pr_入力欄(Ws_表, St_対象指図セル, "指図No", "", "")
    Call Pr_入力欄(Ws_表, St_変更着手セル, "新しい着手予定（日後）", 0, "")
    Call Pr_ボタン作成(Ws_表, "E11", 2, "着手日を変更", "Pr_着手日変更", "製造_着手日")
    Call Pr_ボタン作成(Ws_表, "G11", 2, "製造指図を取り消す", "Pr_製造指図取消", "製造_取消")
    Call Pr_ボタン作成(Ws_表, "I11", 2, "ピッキング指示を出す", "Pr_製造ピッキング指示", "製造_ピッキング")
    Call Pr_ボタン作成(Ws_表, "E13", 2, "着手する（部材なし）", "Pr_部材なし着手", "製造_部材なし着手")
    Call Pr_ボタン作成(Ws_表, "G13", 2, "完成を計上（工程なし）", "Pr_完成計上", "製造_完成計上")
End Sub
