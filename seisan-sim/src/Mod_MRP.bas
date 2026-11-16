Attribute VB_Name = "Mod_MRP"
Option Explicit
'==============================================================
' Mod_MRP
' 【目的】MRP（資材所要量計算）を行い、手配案（製造指図・発注の案）を
'         MRP シートに出す。選んだ行、またはすべての行を手配する。
' 【業務背景】HTML 版の「MRP」タブと同じ計算をする。
'   1. 製造品の独立需要を集める
'      ・見込生産：需要＝未引当（受付）の受注の残数 ＋ まだ受注で埋まっていない見込
'                  （見込 － 受付の受注 － 引当済の受注）。受注と見込の多い方が需要になる
'      ・受注生産：受付の受注ごとに1つの需要（どの受注向けかを持つ）
'   2. 計画中でピッキング指示を出していない製造指図の部材を、子品目の需要にする
'   3. ローレベルコード（BOM 上の最も深い階層）の小さい順に品目を計算する
'      ・見込生産と購入品：不足＝総所要＋安全在庫－有効在庫－手配済。1ロットにまとめる
'      ・受注生産：受注ごとに、紐付き製造指図 → 在庫（紐付かない製造指図も含む）の順に
'                  充当し、残りを受注に紐付くロットにする（安全在庫は持たない）
'      ・製造品のロットは、工程表の所要日数を必要日から逆算して着手日を決め、
'        その部材を子品目の需要に足す（だから親から先に計算する）
'   4. 購入品は今日発注、入荷は今日＋リードタイム
' 【前提】マスタ点検を通っていること（循環が無い等）。
'         日付は「今日から何日後」の整数で計算し、表示のときに日付にする。
' 【並び順】HTML 版はデータを新しい順に持つ。同じ必要日の受注が並ぶときの順序を
'         HTML 版と同じにするため、受注・製造指図はシートの下（新しい行）から読む。
' 【処理の流れ（Pr_MRP計算）】
'   0. 変数格納 → 1. チェック → 2. マスタ点検 → 3. 計算 → 4. 書込 → 5. 表示更新
'==============================================================

' ---- MRP シートの配置 ----
Private Const St_安全在庫セル As String = "C5"      ' 安全在庫も確保するか（○）
Private Const Lg_案題名行 As Long = 7               ' 「■ 手配案」の行
Private Const Lg_案見出し行 As Long = 8             ' 手配案の列見出し
Private Const Lg_案開始行 As Long = 9               ' 手配案の1行目
Private Const Lg_日なし As Long = -999999           ' 必要日が無い（安全在庫の補充）ことを表す値

' 手配案の列（MRP シートの B 列から。実行の○はユーザーが入れる）
Private Enum 列_手配案
    実行 = 2            ' B：○を付けた行を「選んだ行を手配」で手配する
    種類                ' C：製造／発注
    品番                ' D
    品名                ' E
    数量                ' F
    単位                ' G
    着手日              ' H：製造は着手日、発注は発注日（今日）
    完成日              ' I：製造は完成日、発注は入荷予定日
    必要日              ' J：空なら安全在庫の補充
    受注No              ' K：受注生産のとき、どの受注向けか
    注意                ' L：必要日に間に合わない 等
    計算                ' M：不足の計算式
    次列
End Enum

' 計算結果（手配案）の項目（Fn_MRP計算 の戻り値の1つ目）
Private Enum 項_手配案
    種類 = 1            ' 製造／発注
    品番
    品名
    数量
    単位
    着手                ' 何日後
    完成                ' 何日後（発注は入荷）
    必要                ' 何日後（Lg_日なし なら必要日なし）
    受注No
    遅れ                ' True なら必要日に間に合わない
    受注生産            ' True なら受注生産の品目
    計算式
    次項
End Enum

' 計算の内訳（全品目）の項目（Fn_MRP計算 の戻り値の2つ目）
Private Enum 項_内訳
    品番 = 1
    品名
    段                  ' ローレベルコード
    総所要
    有効
    手配済
    安全
    不足
    方式
    次項
End Enum

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_MRP As Worksheet             ' MRP シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Bl_安全在庫 As Boolean          ' MRP シート C5 が○なら安全在庫も確保
Private Dt_今日 As Date                 ' 計算の基準日（今日）

' ---- MRP 計算の作業領域（Fn_MRP計算 の中だけで使う。呼ぶたびに作り直す）----
'      需要は計算の途中で増えていく（親のロットが子の需要を生む）ので、伸ばせる配列で持つ。
Private Lg_需要数 As Long
Private St_需要品番() As String
Private Db_需要量() As Double
Private Lg_需要日() As Long             ' 必要日（今日から何日後）
Private St_需要種類() As String         ' so＝受注、fc＝見込、dep＝親品目から
Private St_需要受注() As String         ' 受注生産のとき、どの受注向けか

'==============================================================
' Pr_MRP計算（メイン／目次）  ボタン：MRP シート「MRP計算」
'==============================================================
Public Sub Pr_MRP計算()
    Dim Va_結果 As Variant

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_MRP, St_シート_品目, St_シート_設定)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. シート・事業部・安全在庫の選択を格納（読み直さない）
    Call Pr_入力チェック                                    ' 1. 事業部の値を確認（誤った事業部で計算しない）→NGなら中断
    Call Pr_マスタ点検(St_事業部)                           ' 2. マスタの誤りを点検（計算を狂わせない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_安全在庫設定保存                                ' 2c. 安全在庫の選択を設定に残す（メニューの件数と合わせる）
    Va_結果 = Fn_MRP計算(St_事業部, Bl_安全在庫, Dt_今日)    ' 3. 所要量を計算して手配案を作る（HTMLと同じ計算）
    Call Pr_結果書込(Va_結果)                               ' 4. 手配案と内訳を MRP シートに書く（判断材料を見せる）
    Call Pr_計算結果表示(Va_結果)                           ' 4b. 件数を結果欄に出す（何件の手配が要るか伝える）
    Call Pr_画面最新化                                      ' 5. メニューの件数と選択肢を更新（状況を最新に）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_MRP選択手配（メイン／目次）  ボタン：MRP シート「選んだ行を手配」
'==============================================================
Public Sub Pr_MRP選択手配()
    Dim Va_結果 As Variant
    Dim Va_選択 As Variant
    Dim St_作成 As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_MRP, St_シート_品目, St_シート_設定)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. シート・事業部・安全在庫の選択を格納（読み直さない）
    Call Pr_入力チェック                                    ' 1. 事業部の値を確認（誤った事業部で手配しない）→NGなら中断
    Call Pr_マスタ点検(St_事業部)                           ' 2. マスタの誤りを点検（計算を狂わせない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Va_結果 = Fn_MRP計算(St_事業部, Bl_安全在庫, Dt_今日)    ' 3. 今のデータで計算し直す（表示後の変化に気づくため）
    Call Pr_表示との一致確認(Va_結果)                       ' 4. 表示中の手配案と同じか確認（古い案で手配しない）→違えば中断
    Va_選択 = Fn_選択行取得(Va_結果)                        ' 5. ○を付けた行を集める（選んだ行だけ手配）→0件なら中断
    St_作成 = Fn_手配実行(Va_結果, Va_選択)                 ' 6. 製造指図・発注を作る（不足を埋める）
    Va_結果 = Fn_MRP計算(St_事業部, Bl_安全在庫, Dt_今日)    ' 7. 手配後に計算し直す（残りの不足を見せる）
    Call Pr_結果書込(Va_結果)                               ' 8. 手配案と内訳を書き直す（最新の状態を見せる）
    Call Pr_手配結果表示(St_作成)                           ' 9. 作った伝票を結果欄に出す（何を作ったか伝える）
    Call Pr_画面最新化                                      ' 10. メニューの件数と選択肢を更新（状況を最新に）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_MRP一括手配（メイン／目次）  ボタン：MRP シート「まとめて手配」
'==============================================================
Public Sub Pr_MRP一括手配()
    Dim Va_結果 As Variant
    Dim Va_選択 As Variant
    Dim St_作成 As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_MRP, St_シート_品目, St_シート_設定)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. シート・事業部・安全在庫の選択を格納（読み直さない）
    Call Pr_入力チェック                                    ' 1. 事業部の値を確認（誤った事業部で手配しない）→NGなら中断
    Call Pr_マスタ点検(St_事業部)                           ' 2. マスタの誤りを点検（計算を狂わせない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_安全在庫設定保存                                ' 2c. 安全在庫の選択を設定に残す（メニューの件数と合わせる）
    Va_結果 = Fn_MRP計算(St_事業部, Bl_安全在庫, Dt_今日)    ' 3. 今のデータで計算する（HTMLの一括実行と同じ）
    Va_選択 = Fn_全行選択(Va_結果)                          ' 4. すべての案を選ぶ（まとめて手配）→0件なら中断
    St_作成 = Fn_手配実行(Va_結果, Va_選択)                 ' 5. 製造指図・発注を作る（不足を埋める）
    Va_結果 = Fn_MRP計算(St_事業部, Bl_安全在庫, Dt_今日)    ' 6. 手配後に計算し直す（残りの不足を見せる）
    Call Pr_結果書込(Va_結果)                               ' 7. 手配案と内訳を書き直す（最新の状態を見せる）
    Call Pr_手配結果表示(St_作成)                           ' 8. 作った伝票を結果欄に出す（何を作ったか伝える）
    Call Pr_画面最新化                                      ' 9. メニューの件数と選択肢を更新（状況を最新に）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_MRP = ThisWorkbook.Worksheets(St_シート_MRP)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Bl_安全在庫 = (Fn_文字(Ws_MRP.Range(St_安全在庫セル).Value) = St_印)
    Dt_今日 = Date
End Sub

' 現在の事業部が正しいか確かめる
Private Sub Pr_入力チェック()
    If St_事業部 <> St_加工 And St_事業部 <> St_組立 Then
        Call Pr_入力エラー("設定シートの「現在の事業部」が「加工」か「組立」になっていません。メニューで切り替えてください。")
    End If
End Sub

' 安全在庫の選択を設定シートに残す（メニューの「要手配」件数を同じ条件で数えるため）
Private Sub Pr_安全在庫設定保存()
    If Bl_安全在庫 Then
        Call Pr_設定値書込("MRP安全在庫", St_印)
    Else
        Call Pr_設定値書込("MRP安全在庫", "")
    End If
End Sub

'==============================================================
' MRP の計算（ほかのモジュールからも使う）
'==============================================================

' メニューの「要手配」の件数（設定の安全在庫の選択で計算する）
'   引数：St_対象事業部　戻り値：手配案の件数
Public Function Fn_要手配件数(ByVal St_対象事業部 As String) As Long
    Dim Va_結果 As Variant
    Va_結果 = Fn_MRP計算(St_対象事業部, Fn_文字(Fn_設定値("MRP安全在庫")) = St_印, Date)
    Fn_要手配件数 = Fn_行数(Va_結果(0))
End Function

' MRP を計算する（HTML 版の mrp）
'   引数：St_対象事業部、Bl_安全在庫を使う、Dt_基準日（今日）
'   戻り値：2要素の配列。(0)＝手配案（行×項_手配案。0件なら Empty）、(1)＝内訳（行×項_内訳）
Public Function Fn_MRP計算(ByVal St_対象事業部 As String, ByVal Bl_安全在庫を使う As Boolean, ByVal Dt_基準日 As Date) As Variant
    Dim Va_品目 As Variant, Va_BOM As Variant, Va_受注 As Variant, Va_見込 As Variant, Va_指図 As Variant
    Dim Ob_添字 As Object, Ob_受注引当 As Object, Ob_製造用 As Object, Ob_手配済 As Object, Ob_出来高 As Object
    Dim Lg_品目数 As Long, Lg_添字 As Long, Lg_行 As Long, Lg_子 As Long
    Dim Lg_品目行() As Long             ' この事業部の品目の、品目配列上の行（品目シートの並び順）
    Dim Lg_段() As Long                 ' ローレベルコード
    Dim Lg_順() As Long                 ' 計算する順（ローレベルコードの小さい順。同じなら品目シートの順）
    Dim Va_案行() As Variant           ' 手配案を1件ずつためる（1件＝項_手配案の1次元配列）
    Dim Va_一行 As Variant, Va_内訳() As Variant
    Dim Lg_案数 As Long
    Dim St_品番 As String, Bl_製造 As Boolean, Bl_受注生産 As Boolean
    Dim Db_受注計 As Double, Db_見込計 As Double, Db_未消込 As Double
    Dim Lg_受注最早 As Long, Lg_見込最早 As Long
    Dim Lg_需要() As Long, Lg_需要件数 As Long
    Dim Db_総所要 As Double, Db_有効 As Double, Db_手配済 As Double, Db_安全 As Double, Db_不足 As Double
    Dim Db_受注分 As Double, Db_見込分 As Double, Db_親分 As Double
    Dim Db_枠 As Double, Db_紐付 As Double, Db_不足分 As Double, Db_充当 As Double
    Dim Bl_実需要 As Boolean
    Dim Lg_ロット数 As Long
    Dim Db_ロット量() As Double, Lg_ロット日() As Long, St_ロット受注() As String
    Dim Db_ロット必要() As Double, Db_ロット紐付() As Double, Db_ロット充当() As Double
    Dim Lg_ロット As Long, Lg_日数 As Long, Lg_着手 As Long, Lg_完成 As Long, Bl_遅れ As Boolean
    Dim Ob_対象受注 As Object
    Dim St_状態 As String

    ' ---- 読み込み ----
    Va_品目 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_品目), 列_品目.次列 - 1)
    Va_BOM = Fn_表読込(ThisWorkbook.Worksheets(St_シート_BOM), 列_BOM.次列 - 1)
    Va_受注 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_受注), 列_受注.次列 - 1)
    Va_見込 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_見込), 列_見込.次列 - 1)
    Va_指図 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_製造指図), 列_製造指図.次列 - 1)
    Set Ob_受注引当 = Fn_受注引当集計(St_対象事業部)
    Set Ob_製造用 = Fn_製造用確保集計(St_対象事業部)
    Set Ob_手配済 = Fn_手配済集計(St_対象事業部)
    Set Ob_出来高 = Fn_指図出来高集計(St_対象事業部)
    Lg_需要数 = 0
    ReDim St_需要品番(1 To 16): ReDim Db_需要量(1 To 16): ReDim Lg_需要日(1 To 16)
    ReDim St_需要種類(1 To 16): ReDim St_需要受注(1 To 16)

    ' ---- この事業部の品目と、品番 → 添字 ----
    Set Ob_添字 = CreateObject("Scripting.Dictionary")
    Lg_品目数 = 0
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_対象事業部 Then
            Lg_品目数 = Lg_品目数 + 1
            ReDim Preserve Lg_品目行(1 To Lg_品目数)
            Lg_品目行(Lg_品目数) = Lg_行
            Ob_添字.Item(Fn_文字(Va_品目(Lg_行, 列_品目.品番))) = Lg_品目数
        End If
    Next Lg_行
    If Lg_品目数 = 0 Then
        Fn_MRP計算 = Array(Empty, Empty)
        Exit Function
    End If

    ' ---- ローレベルコード（HTML 版の llc：子の段 ＝ 親の段＋1 の最大。落ち着くまで繰り返す）----
    ReDim Lg_段(1 To Lg_品目数)
    Call Pr_段計算(St_対象事業部, Va_BOM, Ob_添字, Lg_段)
    ' 計算する順：段の小さい順（同じ段は品目シートの順を保つ＝安定な並べ替え）
    ReDim Lg_順(1 To Lg_品目数)
    For Lg_添字 = 1 To Lg_品目数
        Lg_順(Lg_添字) = Lg_添字
    Next Lg_添字
    Call Pr_安定並べ替え(Lg_順, Lg_段)

    ' ---- 1. 製造品の独立需要（受注・見込）----
    For Lg_添字 = 1 To Lg_品目数
        Lg_行 = Lg_品目行(Lg_添字)
        St_品番 = Fn_文字(Va_品目(Lg_行, 列_品目.品番))
        Bl_製造 = Fn_製造品か(Fn_文字(Va_品目(Lg_行, 列_品目.区分)))
        Bl_受注生産 = Bl_製造 And Fn_文字(Va_品目(Lg_行, 列_品目.生産方式)) = St_受注生産
        If Bl_受注生産 Then
            ' 受注生産：受付の受注ごとに需要（新しい受注から＝HTML 版の並び）
            For Lg_子 = Fn_行数(Va_受注) To 1 Step -1
                If Fn_受付受注か(Va_受注, Lg_子, St_対象事業部, St_品番) Then
                    Call Pr_需要追加(St_品番, Fn_数(Va_受注(Lg_子, 列_受注.数量)) - Fn_数(Va_受注(Lg_子, 列_受注.出荷済)), _
                        Fn_日数差(Va_受注(Lg_子, 列_受注.納期), Dt_基準日), "so", Fn_文字(Va_受注(Lg_子, 列_受注.受注No)))
                End If
            Next Lg_子
        ElseIf Bl_製造 Then
            ' 見込生産：受注が見込を消し込む
            Db_受注計 = 0: Lg_受注最早 = 999999
            For Lg_子 = 1 To Fn_行数(Va_受注)
                If Fn_受付受注か(Va_受注, Lg_子, St_対象事業部, St_品番) Then
                    Db_受注計 = Db_受注計 + Fn_数(Va_受注(Lg_子, 列_受注.数量)) - Fn_数(Va_受注(Lg_子, 列_受注.出荷済))
                    If Fn_日数差(Va_受注(Lg_子, 列_受注.納期), Dt_基準日) < Lg_受注最早 Then Lg_受注最早 = Fn_日数差(Va_受注(Lg_子, 列_受注.納期), Dt_基準日)
                End If
            Next Lg_子
            Db_見込計 = 0: Lg_見込最早 = 999999
            For Lg_子 = 1 To Fn_行数(Va_見込)
                If Fn_文字(Va_見込(Lg_子, 列_見込.事業部)) = St_対象事業部 And Fn_文字(Va_見込(Lg_子, 列_見込.品番)) = St_品番 Then
                    Db_見込計 = Db_見込計 + Fn_数(Va_見込(Lg_子, 列_見込.数量))
                    If Fn_日数差(Va_見込(Lg_子, 列_見込.期日), Dt_基準日) < Lg_見込最早 Then Lg_見込最早 = Fn_日数差(Va_見込(Lg_子, 列_見込.期日), Dt_基準日)
                End If
            Next Lg_子
            If Db_受注計 > 0 Then Call Pr_需要追加(St_品番, Db_受注計, Lg_受注最早, "so", "")
            Db_未消込 = Fn_丸め(Db_見込計 - Db_受注計 - Fn_辞書値(Ob_受注引当, St_品番))
            If Db_未消込 > 0 Then Call Pr_需要追加(St_品番, Db_未消込, Lg_見込最早, "fc", "")
        End If
    Next Lg_添字

    ' ---- 2. 計画中（ピッキング指示なし）の製造指図の部材 ----
    For Lg_行 = Fn_行数(Va_指図) To 1 Step -1
        If Fn_文字(Va_指図(Lg_行, 列_製造指図.事業部)) = St_対象事業部 _
           And Fn_文字(Va_指図(Lg_行, 列_製造指図.状態)) = St_状態_計画 _
           And Fn_文字(Va_指図(Lg_行, 列_製造指図.ピッキングNo)) = "" Then
            Lg_着手 = Fn_日数差(Va_指図(Lg_行, 列_製造指図.着手予定), Dt_基準日)
            If Lg_着手 < 0 Then Lg_着手 = 0
            Call Pr_子の需要追加(St_対象事業部, Va_BOM, Fn_文字(Va_指図(Lg_行, 列_製造指図.品番)), _
                Fn_数(Va_指図(Lg_行, 列_製造指図.数量)), Lg_着手)
        End If
    Next Lg_行

    ' ---- 3. 段の小さい順に、品目ごとに不足を計算 ----
    ReDim Va_内訳(1 To Lg_品目数, 1 To 項_内訳.次項 - 1)
    Lg_案数 = 0
    For Lg_添字 = 1 To Lg_品目数
        Lg_行 = Lg_品目行(Lg_順(Lg_添字))
        St_品番 = Fn_文字(Va_品目(Lg_行, 列_品目.品番))
        Bl_製造 = Fn_製造品か(Fn_文字(Va_品目(Lg_行, 列_品目.区分)))
        Bl_受注生産 = Bl_製造 And Fn_文字(Va_品目(Lg_行, 列_品目.生産方式)) = St_受注生産
        ' この品目の需要を、必要日の早い順に（同じ日は追加された順）
        Call Pr_品目の需要(St_品番, Lg_需要, Lg_需要件数)
        Db_総所要 = 0: Db_受注分 = 0: Db_見込分 = 0: Db_親分 = 0
        For Lg_子 = 1 To Lg_需要件数
            Db_総所要 = Db_総所要 + Db_需要量(Lg_需要(Lg_子))
            Select Case St_需要種類(Lg_需要(Lg_子))
                Case "so": Db_受注分 = Db_受注分 + Db_需要量(Lg_需要(Lg_子))
                Case "fc": Db_見込分 = Db_見込分 + Db_需要量(Lg_需要(Lg_子))
                Case Else: Db_親分 = Db_親分 + Db_需要量(Lg_需要(Lg_子))
            End Select
        Next Lg_子
        Db_総所要 = Fn_丸め(Db_総所要)
        Db_有効 = Fn_丸め(Fn_数(Va_品目(Lg_行, 列_品目.在庫数)) - Fn_辞書値(Ob_受注引当, St_品番) - Fn_辞書値(Ob_製造用, St_品番))
        Db_手配済 = Fn_辞書値(Ob_手配済, St_品番)
        Db_安全 = 0: Db_不足 = 0: Lg_ロット数 = 0
        ReDim Db_ロット量(1 To 1): ReDim Lg_ロット日(1 To 1): ReDim St_ロット受注(1 To 1)
        ReDim Db_ロット必要(1 To 1): ReDim Db_ロット紐付(1 To 1): ReDim Db_ロット充当(1 To 1)
        If Bl_受注生産 Then
            ' 受注生産：紐付き製造指図 → 在庫（紐付かない製造指図を含む）の順に充当
            Set Ob_対象受注 = CreateObject("Scripting.Dictionary")
            For Lg_子 = 1 To Lg_需要件数
                If St_需要受注(Lg_需要(Lg_子)) <> "" Then Ob_対象受注.Item(St_需要受注(Lg_需要(Lg_子))) = 1
            Next Lg_子
            Db_枠 = Db_有効
            For Lg_子 = 1 To Fn_行数(Va_指図)
                St_状態 = Fn_文字(Va_指図(Lg_子, 列_製造指図.状態))
                If Fn_文字(Va_指図(Lg_子, 列_製造指図.事業部)) = St_対象事業部 _
                   And Fn_文字(Va_指図(Lg_子, 列_製造指図.品番)) = St_品番 _
                   And (St_状態 = St_状態_計画 Or St_状態 = St_状態_着手) Then
                    If Not Ob_対象受注.Exists(Fn_文字(Va_指図(Lg_子, 列_製造指図.受注No))) Then
                        Db_枠 = Db_枠 + Fn_辞書値(Ob_出来高, Fn_文字(Va_指図(Lg_子, 列_製造指図.指図No)))
                    End If
                End If
            Next Lg_子
            Db_枠 = Fn_丸め(Db_枠)
            For Lg_子 = 1 To Lg_需要件数
                If St_需要受注(Lg_需要(Lg_子)) <> "" Then
                    Db_紐付 = Fn_紐付出来高(St_対象事業部, St_需要受注(Lg_需要(Lg_子)), St_品番)
                Else
                    Db_紐付 = 0
                End If
                Db_不足分 = Fn_丸め(Db_需要量(Lg_需要(Lg_子)) - Db_紐付)
                Db_充当 = Db_枠
                If Db_不足分 < Db_充当 Then Db_充当 = Db_不足分
                If Db_充当 < 0 Then Db_充当 = 0
                Db_枠 = Fn_丸め(Db_枠 - Db_充当)
                Db_不足分 = Fn_丸め(Db_不足分 - Db_充当)
                If Db_不足分 > Db_誤差 Then
                    Lg_ロット数 = Lg_ロット数 + 1
                    ReDim Preserve Db_ロット量(1 To Lg_ロット数): ReDim Preserve Lg_ロット日(1 To Lg_ロット数)
                    ReDim Preserve St_ロット受注(1 To Lg_ロット数): ReDim Preserve Db_ロット必要(1 To Lg_ロット数)
                    ReDim Preserve Db_ロット紐付(1 To Lg_ロット数): ReDim Preserve Db_ロット充当(1 To Lg_ロット数)
                    Db_ロット量(Lg_ロット数) = Fn_切上げ(Db_不足分)
                    Lg_ロット日(Lg_ロット数) = Lg_需要日(Lg_需要(Lg_子))
                    St_ロット受注(Lg_ロット数) = St_需要受注(Lg_需要(Lg_子))
                    Db_ロット必要(Lg_ロット数) = Db_需要量(Lg_需要(Lg_子))
                    Db_ロット紐付(Lg_ロット数) = Db_紐付
                    Db_ロット充当(Lg_ロット数) = Db_充当
                    Db_不足 = Fn_丸め(Db_不足 + Db_不足分)
                End If
            Next Lg_子
        Else
            ' 見込生産・購入品：総所要＋安全在庫－有効在庫－手配済 を1ロットに
            If Bl_安全在庫を使う Then Db_安全 = Fn_数(Va_品目(Lg_行, 列_品目.安全在庫))
            Db_不足 = Fn_丸め(Db_総所要 + Db_安全 - Db_有効 - Db_手配済)
            If Db_不足 < 0 Then Db_不足 = 0
            If Db_不足 > 0 Then
                Lg_ロット数 = 1
                Db_ロット量(1) = Fn_切上げ(Db_不足)
                If Lg_需要件数 > 0 Then Lg_ロット日(1) = Lg_需要日(Lg_需要(1)) Else Lg_ロット日(1) = Lg_日なし
                St_ロット受注(1) = ""
            End If
        End If
        ' 本当の需要があるか（安全在庫の補充だけなら、遅れの警告を出さない）
        Bl_実需要 = Bl_受注生産 Or Fn_丸め(Db_総所要 - Db_有効 - Db_手配済) > 0
        ' ロットごとに日程を決めて手配案にする
        For Lg_ロット = 1 To Lg_ロット数
            Bl_遅れ = False
            If Bl_製造 Then
                Lg_日数 = Fn_所要日数(St_対象事業部, St_品番, Db_ロット量(Lg_ロット))
                If Lg_ロット日(Lg_ロット) = Lg_日なし Then
                    Lg_着手 = 0: Lg_完成 = Lg_日数
                ElseIf Lg_ロット日(Lg_ロット) - Lg_日数 < 0 Then
                    Bl_遅れ = Bl_実需要: Lg_着手 = 0: Lg_完成 = Lg_日数
                Else
                    Lg_着手 = Lg_ロット日(Lg_ロット) - Lg_日数: Lg_完成 = Lg_ロット日(Lg_ロット)
                End If
                ' このロットの部材を子品目の需要に足す（子は段が大きいので、あとで計算される）
                Call Pr_子の需要追加(St_対象事業部, Va_BOM, St_品番, Db_ロット量(Lg_ロット), Lg_着手)
            Else
                Lg_着手 = 0
                Lg_完成 = CLng(Fn_数(Va_品目(Lg_行, 列_品目.LT)))
                Bl_遅れ = Bl_実需要 And Lg_ロット日(Lg_ロット) <> Lg_日なし And Lg_完成 > Lg_ロット日(Lg_ロット)
            End If
            Va_一行 = Fn_空の手配案()
            Va_一行(項_手配案.種類) = IIf(Bl_製造, "製造", "発注")
            Va_一行(項_手配案.品番) = St_品番
            Va_一行(項_手配案.品名) = Va_品目(Lg_行, 列_品目.品名)
            Va_一行(項_手配案.数量) = Db_ロット量(Lg_ロット)
            Va_一行(項_手配案.単位) = Va_品目(Lg_行, 列_品目.単位)
            Va_一行(項_手配案.着手) = Lg_着手
            Va_一行(項_手配案.完成) = Lg_完成
            Va_一行(項_手配案.必要) = Lg_ロット日(Lg_ロット)
            Va_一行(項_手配案.受注No) = St_ロット受注(Lg_ロット)
            Va_一行(項_手配案.遅れ) = Bl_遅れ
            Va_一行(項_手配案.受注生産) = Bl_受注生産
            If Bl_受注生産 Then
                Va_一行(項_手配案.計算式) = "必要 " & Db_ロット必要(Lg_ロット) & " － 紐付き製造 " & Db_ロット紐付(Lg_ロット) & _
                    " － 在庫充当 " & Db_ロット充当(Lg_ロット) & " ＝ 不足 " & Db_ロット量(Lg_ロット)
            Else
                Va_一行(項_手配案.計算式) = "総所要 " & Db_総所要 & " ＋ 安全 " & Db_安全 & " － 有効 " & Db_有効 & _
                    " － 手配済 " & Db_手配済 & " ＝ 不足 " & Db_不足 & Fn_内訳文(Db_受注分, Db_見込分, Db_親分)
            End If
            Lg_案数 = Lg_案数 + 1
            ReDim Preserve Va_案行(1 To Lg_案数)
            Va_案行(Lg_案数) = Va_一行
        Next Lg_ロット
        ' 内訳の行
        Va_内訳(Lg_添字, 項_内訳.品番) = St_品番
        Va_内訳(Lg_添字, 項_内訳.品名) = Va_品目(Lg_行, 列_品目.品名)
        Va_内訳(Lg_添字, 項_内訳.段) = Lg_段(Lg_順(Lg_添字))
        Va_内訳(Lg_添字, 項_内訳.総所要) = Db_総所要
        Va_内訳(Lg_添字, 項_内訳.有効) = Db_有効
        Va_内訳(Lg_添字, 項_内訳.手配済) = Db_手配済
        Va_内訳(Lg_添字, 項_内訳.安全) = IIf(Bl_受注生産, "－", Db_安全)
        Va_内訳(Lg_添字, 項_内訳.不足) = Db_不足
        Va_内訳(Lg_添字, 項_内訳.方式) = IIf(Bl_受注生産, St_受注生産, IIf(Bl_製造, St_見込生産, "購入"))
    Next Lg_添字

    If Lg_案数 = 0 Then
        Fn_MRP計算 = Array(Empty, Va_内訳)
    Else
        Fn_MRP計算 = Array(Fn_2次元化(Va_案行, Lg_案数), Va_内訳)
    End If
End Function

' 受付状態で、この事業部・品番の受注か
Private Function Fn_受付受注か(ByVal Va_受注 As Variant, ByVal Lg_行 As Long, ByVal St_対象事業部 As String, ByVal St_品番 As String) As Boolean
    Fn_受付受注か = (Fn_文字(Va_受注(Lg_行, 列_受注.事業部)) = St_対象事業部 _
        And Fn_文字(Va_受注(Lg_行, 列_受注.品番)) = St_品番 _
        And Fn_文字(Va_受注(Lg_行, 列_受注.状態)) = St_状態_受付)
End Function

' 総所要の内訳の文（受注・見込・親品目から）
Private Function Fn_内訳文(ByVal Db_受注分 As Double, ByVal Db_見込分 As Double, ByVal Db_親分 As Double) As String
    Dim St_文 As String
    If Db_受注分 <> 0 Then St_文 = St_文 & " ／ 受注 " & Fn_丸め(Db_受注分)
    If Db_見込分 <> 0 Then St_文 = St_文 & " ／ 見込 " & Fn_丸め(Db_見込分)
    If Db_親分 <> 0 Then St_文 = St_文 & " ／ 親品目から " & Fn_丸め(Db_親分)
    If St_文 <> "" Then St_文 = "（総所要の内訳" & Mid$(St_文, 3) & "）"
    Fn_内訳文 = St_文
End Function

' 需要を1件足す（HTML 版の add。数量が0以下なら足さない）
Private Sub Pr_需要追加(ByVal St_品番 As String, ByVal Db_量 As Double, ByVal Lg_日 As Long, ByVal St_種類 As String, ByVal St_受注 As String)
    If Db_量 <= 0 Then Exit Sub
    Lg_需要数 = Lg_需要数 + 1
    If Lg_需要数 > UBound(St_需要品番) Then
        ReDim Preserve St_需要品番(1 To Lg_需要数 * 2): ReDim Preserve Db_需要量(1 To Lg_需要数 * 2)
        ReDim Preserve Lg_需要日(1 To Lg_需要数 * 2): ReDim Preserve St_需要種類(1 To Lg_需要数 * 2)
        ReDim Preserve St_需要受注(1 To Lg_需要数 * 2)
    End If
    St_需要品番(Lg_需要数) = St_品番
    Db_需要量(Lg_需要数) = Fn_丸め(Db_量)
    Lg_需要日(Lg_需要数) = Lg_日
    St_需要種類(Lg_需要数) = St_種類
    St_需要受注(Lg_需要数) = St_受注
End Sub

' 親を Db_量 だけ作るときの子品目の所要を、需要（親品目から）に足す
Private Sub Pr_子の需要追加(ByVal St_対象事業部 As String, ByVal Va_BOM As Variant, ByVal St_親 As String, _
        ByVal Db_量 As Double, ByVal Lg_日 As Long)
    Dim Lg_行 As Long
    For Lg_行 = 1 To Fn_行数(Va_BOM)
        If Fn_文字(Va_BOM(Lg_行, 列_BOM.事業部)) = St_対象事業部 And Fn_文字(Va_BOM(Lg_行, 列_BOM.親品番)) = St_親 Then
            Call Pr_需要追加(Fn_文字(Va_BOM(Lg_行, 列_BOM.子品番)), Fn_数(Va_BOM(Lg_行, 列_BOM.員数)) * Db_量, Lg_日, "dep", "")
        End If
    Next Lg_行
End Sub

' 1品目の需要の番号を、必要日の早い順（同じ日は追加順）に並べて返す
'   引数：St_品番、Lg_結果（出力：需要の番号の配列）、Lg_件数（出力）
Private Sub Pr_品目の需要(ByVal St_品番 As String, ByRef Lg_結果() As Long, ByRef Lg_件数 As Long)
    Dim Lg_番 As Long
    Dim Lg_添字 As Long
    Dim Lg_一時 As Long
    Lg_件数 = 0
    ReDim Lg_結果(1 To 1)
    For Lg_番 = 1 To Lg_需要数
        If St_需要品番(Lg_番) = St_品番 Then
            Lg_件数 = Lg_件数 + 1
            ReDim Preserve Lg_結果(1 To Lg_件数)
            Lg_結果(Lg_件数) = Lg_番
        End If
    Next Lg_番
    ' 挿入ソート（同じ日の順序を保つ）
    For Lg_番 = 2 To Lg_件数
        Lg_一時 = Lg_結果(Lg_番)
        Lg_添字 = Lg_番 - 1
        Do While Lg_添字 >= 1
            If Lg_需要日(Lg_結果(Lg_添字)) <= Lg_需要日(Lg_一時) Then Exit Do
            Lg_結果(Lg_添字 + 1) = Lg_結果(Lg_添字)
            Lg_添字 = Lg_添字 - 1
        Loop
        Lg_結果(Lg_添字 + 1) = Lg_一時
    Next Lg_番
End Sub

' ローレベルコードを計算する（子の段 ＝ 親の段＋1 の最大。HTML 版と同じく最大50回）
Private Sub Pr_段計算(ByVal St_対象事業部 As String, ByVal Va_BOM As Variant, ByVal Ob_添字 As Object, ByRef Lg_段() As Long)
    Dim Bl_変化 As Boolean
    Dim Lg_回 As Long
    Dim Lg_行 As Long
    Dim Lg_親 As Long
    Dim Lg_子 As Long
    Bl_変化 = True
    Do While Bl_変化 And Lg_回 < 50
        Bl_変化 = False
        Lg_回 = Lg_回 + 1
        For Lg_行 = 1 To Fn_行数(Va_BOM)
            If Fn_文字(Va_BOM(Lg_行, 列_BOM.事業部)) = St_対象事業部 Then
                If Ob_添字.Exists(Fn_文字(Va_BOM(Lg_行, 列_BOM.親品番))) And Ob_添字.Exists(Fn_文字(Va_BOM(Lg_行, 列_BOM.子品番))) Then
                    Lg_親 = Ob_添字.Item(Fn_文字(Va_BOM(Lg_行, 列_BOM.親品番)))
                    Lg_子 = Ob_添字.Item(Fn_文字(Va_BOM(Lg_行, 列_BOM.子品番)))
                    If Lg_段(Lg_子) < Lg_段(Lg_親) + 1 Then
                        Lg_段(Lg_子) = Lg_段(Lg_親) + 1
                        Bl_変化 = True
                    End If
                End If
            End If
        Next Lg_行
    Loop
End Sub

' 番号の配列を、キーの小さい順に安定に並べ替える（挿入ソート）
Private Sub Pr_安定並べ替え(ByRef Lg_順() As Long, ByRef Lg_キー() As Long)
    Dim Lg_番 As Long
    Dim Lg_添字 As Long
    Dim Lg_一時 As Long
    For Lg_番 = LBound(Lg_順) + 1 To UBound(Lg_順)
        Lg_一時 = Lg_順(Lg_番)
        Lg_添字 = Lg_番 - 1
        Do While Lg_添字 >= LBound(Lg_順)
            If Lg_キー(Lg_順(Lg_添字)) <= Lg_キー(Lg_一時) Then Exit Do
            Lg_順(Lg_添字 + 1) = Lg_順(Lg_添字)
            Lg_添字 = Lg_添字 - 1
        Loop
        Lg_順(Lg_添字 + 1) = Lg_一時
    Next Lg_番
End Sub

' 手配案1件ぶんの空の配列を作る（呼ぶたびに新しい配列。前の行の中身を上書きしないため）
Private Function Fn_空の手配案() As Variant
    Dim Va_行() As Variant
    ReDim Va_行(1 To 項_手配案.次項 - 1)
    Fn_空の手配案 = Va_行
End Function

' 1件ずつためた手配案（1次元配列の配列）を、行×項目の2次元配列にまとめる
'   2次元配列は ReDim Preserve で行を伸ばせないので、1件ずつためてから最後にまとめる。
Private Function Fn_2次元化(ByVal Va_行の並び As Variant, ByVal Lg_件数 As Long) As Variant
    Dim Va_先() As Variant
    Dim Va_一行 As Variant
    Dim Lg_行 As Long
    Dim Lg_項 As Long
    ReDim Va_先(1 To Lg_件数, 1 To 項_手配案.次項 - 1)
    For Lg_行 = 1 To Lg_件数
        Va_一行 = Va_行の並び(Lg_行)
        For Lg_項 = 1 To 項_手配案.次項 - 1
            Va_先(Lg_行, Lg_項) = Va_一行(Lg_項)
        Next Lg_項
    Next Lg_行
    Fn_2次元化 = Va_先
End Function

'==============================================================
' MRP シートへの書き込みと手配
'==============================================================

' 手配案と内訳を MRP シートに書く
'   引数：Va_結果（Fn_MRP計算 の戻り値）
Private Sub Pr_結果書込(ByVal Va_結果 As Variant)
    Dim Va_案 As Variant
    Dim Va_内訳 As Variant
    Dim Va_出力() As Variant
    Dim Lg_行 As Long
    Dim Lg_件数 As Long
    Dim Lg_内訳行 As Long
    Va_案 = Va_結果(0)
    Va_内訳 = Va_結果(1)
    Lg_件数 = Fn_行数(Va_案)
    ' 前回の結果を消す
    Ws_MRP.Range(Ws_MRP.Cells(Lg_案題名行, 1), Ws_MRP.Cells(Ws_MRP.Rows.Count, 列_手配案.次列)).Clear
    Ws_MRP.Cells(Lg_案題名行, 2).Value = "■ 手配案（" & Lg_件数 & "件。" & Fn_事業部名(St_事業部) & "。手配する行の「実行」に○を付けて「選んだ行を手配」）"
    Ws_MRP.Cells(Lg_案題名行, 2).Font.Bold = True
    Ws_MRP.Range(Ws_MRP.Cells(Lg_案見出し行, 列_手配案.実行), Ws_MRP.Cells(Lg_案見出し行, 列_手配案.計算)).Value = _
        Array("実行", "種類", "品番", "品名", "数量", "単位", "着手・発注", "完成・入荷", "必要日", "受注No", "注意", "計算")
    Call Pr_見出し色(Lg_案見出し行, 列_手配案.実行, 列_手配案.計算)
    If Lg_件数 = 0 Then
        Ws_MRP.Cells(Lg_案開始行, 列_手配案.種類).Value = "不足はありません。手配は不要です。"
        Lg_内訳行 = Lg_案開始行 + 2
    Else
        ReDim Va_出力(1 To Lg_件数, 列_手配案.実行 To 列_手配案.計算)
        For Lg_行 = 1 To Lg_件数
            Va_出力(Lg_行, 列_手配案.実行) = ""
            Va_出力(Lg_行, 列_手配案.種類) = Va_案(Lg_行, 項_手配案.種類)
            Va_出力(Lg_行, 列_手配案.品番) = Va_案(Lg_行, 項_手配案.品番)
            Va_出力(Lg_行, 列_手配案.品名) = Va_案(Lg_行, 項_手配案.品名)
            Va_出力(Lg_行, 列_手配案.数量) = Va_案(Lg_行, 項_手配案.数量)
            Va_出力(Lg_行, 列_手配案.単位) = Va_案(Lg_行, 項_手配案.単位)
            Va_出力(Lg_行, 列_手配案.着手日) = Dt_今日 + Va_案(Lg_行, 項_手配案.着手)
            Va_出力(Lg_行, 列_手配案.完成日) = Dt_今日 + Va_案(Lg_行, 項_手配案.完成)
            If Va_案(Lg_行, 項_手配案.必要) = Lg_日なし Then
                Va_出力(Lg_行, 列_手配案.必要日) = "安全在庫の補充"
            Else
                Va_出力(Lg_行, 列_手配案.必要日) = Dt_今日 + Va_案(Lg_行, 項_手配案.必要)
            End If
            Va_出力(Lg_行, 列_手配案.受注No) = Va_案(Lg_行, 項_手配案.受注No)
            Va_出力(Lg_行, 列_手配案.注意) = Fn_注意文(Va_案, Lg_行)
            Va_出力(Lg_行, 列_手配案.計算) = Va_案(Lg_行, 項_手配案.計算式)
        Next Lg_行
        Ws_MRP.Range(Ws_MRP.Cells(Lg_案開始行, 列_手配案.実行), Ws_MRP.Cells(Lg_案開始行 + Lg_件数 - 1, 列_手配案.計算)).Value = Va_出力
        Ws_MRP.Range(Ws_MRP.Cells(Lg_案開始行, 列_手配案.着手日), Ws_MRP.Cells(Lg_案開始行 + Lg_件数 - 1, 列_手配案.必要日)).NumberFormat = "m/d"
        With Ws_MRP.Range(Ws_MRP.Cells(Lg_案開始行, 列_手配案.実行), Ws_MRP.Cells(Lg_案開始行 + Lg_件数 - 1, 列_手配案.実行))
            .Interior.Color = RGB(255, 255, 204)
            .HorizontalAlignment = xlCenter
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=St_印
        End With
        Lg_内訳行 = Lg_案開始行 + Lg_件数 + 1
    End If
    ' 計算の内訳（全品目）
    Ws_MRP.Cells(Lg_内訳行, 2).Value = "■ 計算の内訳（全品目。段＝ローレベルコード。小さい順に計算するので、親の手配数が子の所要に反映される）"
    Ws_MRP.Cells(Lg_内訳行, 2).Font.Bold = True
    Ws_MRP.Range(Ws_MRP.Cells(Lg_内訳行 + 1, 2), Ws_MRP.Cells(Lg_内訳行 + 1, 10)).Value = _
        Array("品番", "品名", "段", "総所要", "有効", "手配済", "安全", "不足", "方式")
    Call Pr_見出し色(Lg_内訳行 + 1, 2, 10)
    If Not IsEmpty(Va_内訳) Then
        Ws_MRP.Range(Ws_MRP.Cells(Lg_内訳行 + 2, 2), Ws_MRP.Cells(Lg_内訳行 + 1 + UBound(Va_内訳, 1), 10)).Value = Va_内訳
    End If
End Sub

' 手配案の注意の文（必要日に間に合わない、受注生産の向け先）
Private Function Fn_注意文(ByVal Va_案 As Variant, ByVal Lg_行 As Long) As String
    Dim St_文 As String
    If Va_案(Lg_行, 項_手配案.受注生産) Then
        If Va_案(Lg_行, 項_手配案.受注No) <> "" Then
            St_文 = Va_案(Lg_行, 項_手配案.受注No) & " 向け"
        Else
            St_文 = "親品目向け"
        End If
    End If
    If Va_案(Lg_行, 項_手配案.遅れ) Then
        If St_文 <> "" Then St_文 = St_文 & "／"
        St_文 = St_文 & "必要日に間に合わない"
    End If
    Fn_注意文 = St_文
End Function

' 表の見出しに色を付ける
Private Sub Pr_見出し色(ByVal Lg_行 As Long, ByVal Lg_始 As Long, ByVal Lg_終 As Long)
    With Ws_MRP.Range(Ws_MRP.Cells(Lg_行, Lg_始), Ws_MRP.Cells(Lg_行, Lg_終))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
End Sub

' 計算した件数を結果欄に出す
Private Sub Pr_計算結果表示(ByVal Va_結果 As Variant)
    Call Pr_結果表示(Ws_MRP, Fn_事業部名(St_事業部) & " の MRP を計算しました。手配案 " & Fn_行数(Va_結果(0)) & " 件" & _
        IIf(Bl_安全在庫, "（安全在庫も確保）", "（安全在庫は確保しない）"))
End Sub

' 表示中の手配案が、いま計算し直した結果と同じか確かめる。違えば書き直して中断
'   表示したあとに受注などが変わっていたら、古い案のまま手配しないようにするため。
Private Sub Pr_表示との一致確認(ByVal Va_結果 As Variant)
    Dim Va_案 As Variant
    Dim Lg_行 As Long
    Dim Lg_件数 As Long
    Dim Bl_一致 As Boolean
    Va_案 = Va_結果(0)
    Lg_件数 = Fn_行数(Va_案)
    Bl_一致 = True
    For Lg_行 = 1 To Lg_件数
        If Fn_文字(Ws_MRP.Cells(Lg_案開始行 + Lg_行 - 1, 列_手配案.品番).Value) <> Va_案(Lg_行, 項_手配案.品番) _
           Or Fn_数(Ws_MRP.Cells(Lg_案開始行 + Lg_行 - 1, 列_手配案.数量).Value) <> Va_案(Lg_行, 項_手配案.数量) _
           Or Fn_文字(Ws_MRP.Cells(Lg_案開始行 + Lg_行 - 1, 列_手配案.受注No).Value) <> Va_案(Lg_行, 項_手配案.受注No) Then
            Bl_一致 = False
            Exit For
        End If
    Next Lg_行
    ' 表示の方が行が多い（案が減った）ときも一致しない
    If Bl_一致 And Fn_文字(Ws_MRP.Cells(Lg_案開始行 + Lg_件数, 列_手配案.品番).Value) <> "" Then Bl_一致 = False
    If Lg_件数 = 0 Then Bl_一致 = False
    If Not Bl_一致 Then
        Call Pr_結果書込(Va_結果)
        MsgBox "表示していた手配案が、いまのデータで計算した結果と違っていました（または手配案がありません）。" & vbCrLf & _
               "最新の結果を表示したので、もう一度「実行」に○を付けてから押してください。", vbExclamation
        Err.Raise Lg_中断番号
    End If
End Sub

' ○を付けた行を集める
'   戻り値：手配案の行ごとの True/False の配列（1始まり）。1件も無ければ中断
Private Function Fn_選択行取得(ByVal Va_結果 As Variant) As Variant
    Dim Va_選択() As Variant
    Dim Lg_行 As Long
    Dim Lg_件数 As Long
    Dim Lg_選択数 As Long
    Lg_件数 = Fn_行数(Va_結果(0))
    ReDim Va_選択(1 To Lg_件数)
    For Lg_行 = 1 To Lg_件数
        Va_選択(Lg_行) = (Fn_文字(Ws_MRP.Cells(Lg_案開始行 + Lg_行 - 1, 列_手配案.実行).Value) <> "")
        If Va_選択(Lg_行) Then Lg_選択数 = Lg_選択数 + 1
    Next Lg_行
    If Lg_選択数 = 0 Then Call Pr_入力エラー("手配する行の「実行」欄に○を付けてください。")
    Fn_選択行取得 = Va_選択
End Function

' すべての行を選ぶ。手配案が0件なら中断
Private Function Fn_全行選択(ByVal Va_結果 As Variant) As Variant
    Dim Va_選択() As Variant
    Dim Lg_行 As Long
    If Fn_行数(Va_結果(0)) = 0 Then
        Call Pr_結果書込(Va_結果)
        Call Pr_入力エラー("不足はありません。手配は不要です。")
    End If
    ReDim Va_選択(1 To Fn_行数(Va_結果(0)))
    For Lg_行 = 1 To Fn_行数(Va_結果(0))
        Va_選択(Lg_行) = True
    Next Lg_行
    Fn_全行選択 = Va_選択
End Function

' 選んだ手配案ごとに、製造指図か発注を作る（HTML 版の exec）
'   戻り値：作った伝票番号を並べた文
Private Function Fn_手配実行(ByVal Va_結果 As Variant, ByVal Va_選択 As Variant) As String
    Dim Va_案 As Variant
    Dim Lg_行 As Long
    Dim St_作成 As String
    Va_案 = Va_結果(0)
    For Lg_行 = 1 To Fn_行数(Va_案)
        If Va_選択(Lg_行) Then
            If Va_案(Lg_行, 項_手配案.種類) = "製造" Then
                St_作成 = St_作成 & "、" & Fn_製造指図作成(St_事業部, CStr(Va_案(Lg_行, 項_手配案.品番)), CDbl(Va_案(Lg_行, 項_手配案.数量)), _
                    CLng(Va_案(Lg_行, 項_手配案.着手)), CLng(Va_案(Lg_行, 項_手配案.完成)), CStr(Va_案(Lg_行, 項_手配案.受注No)), Dt_今日)
            Else
                St_作成 = St_作成 & "、" & Fn_発注作成(St_事業部, CStr(Va_案(Lg_行, 項_手配案.品番)), CDbl(Va_案(Lg_行, 項_手配案.数量)), Dt_今日)
            End If
        End If
    Next Lg_行
    Fn_手配実行 = Mid$(St_作成, 2)
End Function

' 作った伝票を結果欄に出す
Private Sub Pr_手配結果表示(ByVal St_作成 As String)
    Call Pr_結果表示(Ws_MRP, "手配しました：" & St_作成)
End Sub

'==============================================================
' MRP シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_MRPシート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_MRP)
    Call Pr_入力シート枠(Ws_表, "MRP（所要量計算）", _
        "未引当の受注・見込・計画中の製造指図から必要数を出し、BOM を上から順に展開して不足を計算します。")
    Ws_表.Columns("E").ColumnWidth = 16
    Ws_表.Columns("L").ColumnWidth = 22
    Ws_表.Columns("M").ColumnWidth = 60
    Call Pr_入力欄(Ws_表, St_安全在庫セル, "安全在庫も確保する（○）", St_印, "")
    Ws_表.Range(St_安全在庫セル).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=St_印
    Call Pr_ボタン作成(Ws_表, "E5", 1, "MRP計算", "Pr_MRP計算", "MRP_計算")
    Call Pr_ボタン作成(Ws_表, "G5", 2, "選んだ行を手配", "Pr_MRP選択手配", "MRP_選択手配")
    Call Pr_ボタン作成(Ws_表, "J5", 2, "まとめて手配", "Pr_MRP一括手配", "MRP_一括手配")
End Sub
