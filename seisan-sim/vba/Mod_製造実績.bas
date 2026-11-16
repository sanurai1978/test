Attribute VB_Name = "Mod_製造実績"
Option Explicit
'==============================================================
' Mod_製造実績
' 【目的】入力_製造 シートの「工程（手配No.）」の欄の操作を受け持つ。
'         手配No.を読む（工程の内容と既定値を出す）、工程の実績登録、外注に出す。
'         あわせて、その製造指図の工程一覧を表示する。
' 【業務背景】HTML 版の製造指図の画面（工程ごとの実績入力・外注に出す）と同じ。
'   ・手配No.は工程1つごとの全社連番。製造指図書のバーコードを読んで入れる想定。
'   ・実績を入れられるのは、着手中の製造指図の「次の工程」（完了していない最初の工程）だけ。
'   ・良品＋不良 は投入数（1つ目の工程は指示数、2つ目以降は前の工程の良品数）と一致させる。
'   ・最終工程か、良品が0なら製造指図が完成し、良品が在庫に入る（受注生産は自動引当）。
'   ・外注の工程は、社内で実績を入れず「外注に出す」→ 入力_外注で受入。外注に出しても在庫は動かない。
' 【前提】入力_製造 シートの上半分は Mod_製造 が作る。この欄は Pr_製造実績欄作成 で足す。
'==============================================================

' ---- 入力_製造 シートの配置（工程の欄）----
Private Const St_手配Noセル As String = "C17"       ' 手配No.
Private Const St_良品セル As String = "C18"         ' 良品数
Private Const St_不良セル As String = "C19"         ' 不良数
Private Const St_時間セル As String = "C20"         ' 作業時間（分・段取を含む）
Private Const St_内容セル As String = "E18"         ' 読んだ工程の内容を表示する
Private Const St_指図Noセル As String = "C11"       ' 指図の操作の指図No（工程一覧の対象に使う。Mod_製造 の欄）
Private Const Lg_工程一覧行 As Long = 23            ' 工程一覧を書く行

' 工程一覧の項目（表示用）
Private Enum 項_工程表示
    工程順 = 1
    手配No
    工程名
    場所                ' 作業区名、または「外注：外注先」
    投入
    標準                ' 標準時間（分）。外注は「LT n日」
    状態
    良品
    不良
    時間
    次列
End Enum

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_製造 シート
Private Ws_工程 As Worksheet            ' 指図工程シート
Private Ws_指図 As Worksheet            ' 製造指図シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private Dt_今日 As Date                 ' 今日（外注の納期の基準）
Private St_手配No As String             ' C17
Private Va_良品 As Variant              ' C18
Private Va_不良 As Variant              ' C19
Private Va_時間 As Variant              ' C20
Private St_一覧指図 As String           ' C11（工程一覧の対象の指図No）

'==============================================================
' Pr_手配No読込（メイン）  ボタン：入力_製造「手配No.を読む」
'==============================================================
Public Sub Pr_手配No読込()
    Dim Lg_行 As Long
    Dim St_指図No As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_指図工程)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_工程行()                                     ' 1. 手配No.の工程を探す（どの工程か決める）→無ければ中断
    Call Pr_画面更新停止                                    ' 1b. 画面更新を止める（速くする）
    Call Pr_既定値表示(Lg_行)                               ' 2. 工程の内容と良品・時間の既定値を出す（入力の手間を減らす）
    St_指図No = Fn_工程の指図No(Lg_行)                      ' 3. この工程の指図Noを取る（工程一覧の対象にする）
    Call Pr_工程一覧更新(St_指図No)                         ' 4. その指図の工程一覧を出す（どこまで進んだか見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_工程実績登録（メイン）  ボタン：入力_製造「実績を登録」
'==============================================================
Public Sub Pr_工程実績登録()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_指図工程, St_シート_製造指図)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_工程行()                                     ' 1. 手配No.の工程を探す（どの工程か決める）→無ければ中断
    Call Pr_次工程チェック(Lg_行, False)                    ' 2. 着手中の次の社内工程か確認（順番を飛ばさない）→NGなら中断
    Call Pr_実績値チェック(Lg_行)                           ' 3. 良品＋不良＝投入数・時間を確認（数を合わせる）→NGなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_実績書込(Lg_行)                                 ' 4. 良品・不良・時間を書き工程を完了にする（実績を残す）
    Call Pr_完了後処理(Lg_行)                               ' 5. 最終工程なら完成・入庫・自動引当（製品にする）
    Call Pr_実績結果表示(Lg_行)                             ' 6. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 7. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_外注に出す（メイン）  ボタン：入力_製造「外注に出す」
'==============================================================
Public Sub Pr_外注に出す()
    Dim Lg_行 As Long
    Dim St_外注No As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力製造, St_シート_指図工程, St_シート_外注注文)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_工程行()                                     ' 1. 手配No.の工程を探す（どの工程か決める）→無ければ中断
    Call Pr_次工程チェック(Lg_行, True)                     ' 2. 着手中の次の外注工程か確認（順番を飛ばさない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    St_外注No = Fn_外注注文作成(Lg_行)                      ' 3. 外注注文を作り工程を外注中にする（外注先に預ける）
    Call Pr_外注結果表示(St_外注No)                         ' 4. 外注Noを結果欄に出す（入力_外注で受入に使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力製造)
    Set Ws_工程 = ThisWorkbook.Worksheets(St_シート_指図工程)
    Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_今日 = Date
    St_手配No = Fn_文字(Ws_入力.Range(St_手配Noセル).Value)
    Va_良品 = Ws_入力.Range(St_良品セル).Value
    Va_不良 = Ws_入力.Range(St_不良セル).Value
    Va_時間 = Ws_入力.Range(St_時間セル).Value
    St_一覧指図 = Fn_文字(Ws_入力.Range(St_指図Noセル).Value)
End Sub

' 手配No.の工程の行を探す。戻り値：指図工程シートの行。無い・他事業部なら中断
'   手配No.は全社で一意なので、事業部を問わず探してから、現在の事業部か確かめる。
Private Function Fn_工程行() As Long
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    If St_手配No = "" Then Call Pr_入力エラー("手配No.（7桁）を入力してください。")
    Va_表 = Fn_表読込(Ws_工程, 列_指図工程.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_指図工程.手配No)) = St_手配No Then
            If Fn_文字(Va_表(Lg_行, 列_指図工程.事業部)) <> St_事業部 Then
                Call Pr_入力エラー("手配No. " & St_手配No & " は " & Fn_事業部名(Fn_文字(Va_表(Lg_行, 列_指図工程.事業部))) & " の工程です。メニューで事業部を切り替えてください。")
            End If
            Fn_工程行 = Lg_行 + Lg_開始行 - 1
            Exit Function
        End If
    Next Lg_行
    Call Pr_入力エラー("手配No. " & St_手配No & " の工程はありません。")
End Function

' 指図の工程の行（指図工程シートの行番号）を工程順に返す
'   引数：St_指図No　戻り値：行番号の配列（1始まり）。工程が無ければ Empty
Private Function Fn_指図の工程行(ByVal St_指図No As String) As Variant
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim Lg_件数 As Long
    Dim Lg_結果() As Long
    Va_表 = Fn_表読込(Ws_工程, 列_指図工程.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_指図工程.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_指図工程.指図No)) = St_指図No Then
            Lg_件数 = Lg_件数 + 1
            ReDim Preserve Lg_結果(1 To Lg_件数)
            Lg_結果(Lg_件数) = Lg_行 + Lg_開始行 - 1
        End If
    Next Lg_行
    If Lg_件数 = 0 Then
        Fn_指図の工程行 = Empty
    Else
        Fn_指図の工程行 = Lg_結果
    End If
End Function

' 工程の投入数（1つ目の工程は指示数、2つ目以降は前の工程の良品数。HTML 版の moIn）
'   引数：Lg_行（指図工程シートの行）　戻り値：投入数
Private Function Fn_投入数(ByVal Lg_行 As Long) As Double
    Dim St_指図No As String
    Dim Va_行 As Variant
    Dim Lg_添字 As Long
    St_指図No = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value)
    Va_行 = Fn_指図の工程行(St_指図No)
    For Lg_添字 = 1 To UBound(Va_行)
        If Va_行(Lg_添字) = Lg_行 Then
            If Lg_添字 = 1 Then
                Fn_投入数 = Fn_数(Ws_指図.Cells(Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_指図No), 列_製造指図.数量).Value)
            Else
                Fn_投入数 = Fn_数(Ws_工程.Cells(Va_行(Lg_添字 - 1), 列_指図工程.良品).Value)
            End If
            Exit Function
        End If
    Next Lg_添字
End Function

' 標準時間（分）＝ 段取＋加工×投入数（HTML 版の opMin。外注は 0）
Private Function Fn_標準時間(ByVal Lg_行 As Long, ByVal Db_投入 As Double) As Double
    If Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.外注)) = St_印 Then
        Fn_標準時間 = 0
    Else
        Fn_標準時間 = Fn_丸め(Fn_数(Ws_工程.Cells(Lg_行, 列_指図工程.段取時間).Value) + Fn_数(Ws_工程.Cells(Lg_行, 列_指図工程.加工時間).Value) * Db_投入)
    End If
End Function

' 最終工程か（指図の工程の最後の行か）
Private Function Fn_最終工程か(ByVal Lg_行 As Long) As Boolean
    Dim Va_行 As Variant
    Va_行 = Fn_指図の工程行(Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value))
    Fn_最終工程か = (Va_行(UBound(Va_行)) = Lg_行)
End Function

' 工程の内容と、入力の既定値（良品＝投入数、不良＝0、時間＝標準時間を四捨五入）を出す
Private Sub Pr_既定値表示(ByVal Lg_行 As Long)
    Dim Db_投入 As Double
    Dim St_状態 As String
    Db_投入 = Fn_投入数(Lg_行)
    St_状態 = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value)
    Ws_入力.Range(St_内容セル).Value = Fn_工程の説明(Lg_行, Db_投入)
    If St_状態 = St_状態_未 Then
        Ws_入力.Range(St_良品セル).Value = Db_投入
        Ws_入力.Range(St_不良セル).Value = 0
        Ws_入力.Range(St_時間セル).Value = Int(Fn_標準時間(Lg_行, Db_投入) + 0.5)
    End If
    Ws_入力.Range(St_指図Noセル).Value = Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value
End Sub

' 工程の指図No
Private Function Fn_工程の指図No(ByVal Lg_行 As Long) As String
    Fn_工程の指図No = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value)
End Function

' 工程の説明の文（指図・工程・場所・投入・標準・状態）
Private Function Fn_工程の説明(ByVal Lg_行 As Long, ByVal Db_投入 As Double) As String
    Dim St_場所 As String
    If Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.外注)) = St_印 Then
        St_場所 = "外注：" & Ws_工程.Cells(Lg_行, 列_指図工程.外注先).Value & "（LT " & Ws_工程.Cells(Lg_行, 列_指図工程.外注LT).Value & "日）"
    Else
        St_場所 = Fn_作業区名(Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.作業区).Value)) & "・標準 " & Fn_表示数(Fn_標準時間(Lg_行, Db_投入)) & "分"
    End If
    Fn_工程の説明 = Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value & "　工程 " & Ws_工程.Cells(Lg_行, 列_指図工程.工程順).Value & " " & _
        Ws_工程.Cells(Lg_行, 列_指図工程.工程名).Value & "　" & St_場所 & "　投入 " & Fn_表示数(Db_投入) & "　状態：" & _
        Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value
End Function

' 作業区コードから作業区名
Private Function Fn_作業区名(ByVal St_コード As String) As String
    Dim Ws_作業区 As Worksheet
    Dim Lg_行 As Long
    Set Ws_作業区 = ThisWorkbook.Worksheets(St_シート_作業区)
    Lg_行 = Fn_行検索(Ws_作業区, 列_作業区.次列 - 1, St_事業部, 列_作業区.作業区, St_コード)
    If Lg_行 = 0 Then
        Fn_作業区名 = St_コード
    Else
        Fn_作業区名 = Fn_文字(Ws_作業区.Cells(Lg_行, 列_作業区.作業区名).Value)
    End If
End Function

' 着手中の製造指図の「次の工程」か確かめる。外注か社内かも合わせて確かめる
'   引数：Lg_行、Bl_外注（True＝外注に出す操作、False＝社内の実績登録）
Private Sub Pr_次工程チェック(ByVal Lg_行 As Long, ByVal Bl_外注 As Boolean)
    Dim St_指図No As String
    Dim Va_行 As Variant
    Dim Lg_添字 As Long
    Dim Lg_次 As Long
    Dim St_状態 As String
    St_指図No = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value)
    If Fn_文字(Ws_指図.Cells(Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_指図No), 列_製造指図.状態).Value) <> St_状態_着手 Then
        Call Pr_入力エラー("製造指図 " & St_指図No & " は着手していません（計画中なら部材を払い出すと着手、完成済みなら実績は入れられません）。")
    End If
    Va_行 = Fn_指図の工程行(St_指図No)
    For Lg_添字 = 1 To UBound(Va_行)
        If Fn_文字(Ws_工程.Cells(Va_行(Lg_添字), 列_指図工程.状態).Value) <> St_状態_完了 Then
            Lg_次 = Va_行(Lg_添字)
            Exit For
        End If
    Next Lg_添字
    If Lg_次 <> Lg_行 Then
        Call Pr_入力エラー("手配No. " & St_手配No & " はいま作業する工程ではありません（工程は順番に進めます。状態：" & Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value & "）。")
    End If
    St_状態 = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value)
    If Bl_外注 Then
        If Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.外注)) <> St_印 Then Call Pr_入力エラー("この工程は社内の工程です。「実績を登録」を使ってください。")
        If St_状態 = St_状態_外注中 Then Call Pr_入力エラー("この工程はすでに外注に出しています。戻ってきたら入力_外注で受入を計上してください。")
    Else
        If Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.外注)) = St_印 Then
            Call Pr_入力エラー("この工程は外注の工程です。「外注に出す」→ 入力_外注で受入、の順で進めます。")
        End If
    End If
End Sub

' 良品・不良・時間を確かめる（良品・不良は0以上の整数、良品＋不良＝投入数、時間は0以上の数）
Private Sub Pr_実績値チェック(ByVal Lg_行 As Long)
    Dim Db_投入 As Double
    If Not Fn_整数か(Va_良品) Or Not Fn_整数か(Va_不良) Then Call Pr_入力エラー("良品数・不良数は0以上の整数で入力してください。")
    Db_投入 = Fn_投入数(Lg_行)
    If CDbl(Va_良品) + CDbl(Va_不良) <> Db_投入 Then
        Call Pr_入力エラー("良品数＋不良数を投入数（" & Fn_表示数(Db_投入) & "）と一致させてください（いま " & Fn_表示数(CDbl(Va_良品) + CDbl(Va_不良)) & "）。")
    End If
    If Fn_文字(Va_時間) = "" Or Not IsNumeric(Va_時間) Then Call Pr_入力エラー("作業時間は0以上の数で入力してください。")
    If CDbl(Va_時間) < 0 Then Call Pr_入力エラー("作業時間は0以上の数で入力してください。")
End Sub

' 0以上の整数か
Private Function Fn_整数か(ByVal Va_値 As Variant) As Boolean
    Fn_整数か = False
    If Fn_文字(Va_値) = "" Or Not IsNumeric(Va_値) Then Exit Function
    Fn_整数か = (CDbl(Va_値) >= 0 And CDbl(Va_値) = Int(CDbl(Va_値)))
End Function

' 実績を書き、工程を完了にする
Private Sub Pr_実績書込(ByVal Lg_行 As Long)
    Ws_工程.Cells(Lg_行, 列_指図工程.良品).Value = CDbl(Va_良品)
    Ws_工程.Cells(Lg_行, 列_指図工程.不良).Value = CDbl(Va_不良)
    Ws_工程.Cells(Lg_行, 列_指図工程.作業時間).Value = Fn_丸め(CDbl(Va_時間))
    Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value = St_状態_完了
End Sub

' 工程が終わったあとの処理（最終工程か良品0なら完成・入庫・自動引当）
Private Sub Pr_完了後処理(ByVal Lg_行 As Long)
    Call Pr_工程完了後処理(St_事業部, Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value), Fn_最終工程か(Lg_行), CDbl(Va_良品))
End Sub

' 外注注文を作り、工程を「外注中」にする。戻り値：外注No
'   数量は投入数、納期は 今日＋外注LT（HTML 版と同じ）。
Private Function Fn_外注注文作成(ByVal Lg_行 As Long) As String
    Dim St_外注No As String
    Dim St_指図No As String
    St_指図No = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value)
    St_外注No = Fn_採番(St_事業部, "SC")
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_外注注文), Array(St_事業部, St_外注No, St_指図No, _
        Ws_工程.Cells(Lg_行, 列_指図工程.工程順).Value, Ws_工程.Cells(Lg_行, 列_指図工程.工程名).Value, _
        Ws_指図.Cells(Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_指図No), 列_製造指図.品番).Value, _
        Ws_工程.Cells(Lg_行, 列_指図工程.外注先).Value, Fn_投入数(Lg_行), Ws_工程.Cells(Lg_行, 列_指図工程.外注単価).Value, _
        Dt_今日 + Fn_数(Ws_工程.Cells(Lg_行, 列_指図工程.外注LT).Value), St_状態_外注中, "", ""))
    Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value = St_状態_外注中
    Fn_外注注文作成 = St_外注No
End Function

' 結果欄への表示
Private Sub Pr_実績結果表示(ByVal Lg_行 As Long)
    Dim St_指図No As String
    St_指図No = Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.指図No).Value)
    Call Pr_結果表示(Ws_入力, "手配No. " & St_手配No & " の実績を登録しました（" & St_指図No & " 工程 " & _
        Ws_工程.Cells(Lg_行, 列_指図工程.工程順).Value & "）。指図の状態：" & _
        Ws_指図.Cells(Fn_行検索(Ws_指図, 列_製造指図.次列 - 1, St_事業部, 列_製造指図.指図No, St_指図No), 列_製造指図.状態).Value)
    Ws_入力.Range(St_内容セル).Value = ""
End Sub

Private Sub Pr_外注結果表示(ByVal St_外注No As String)
    Call Pr_結果表示(Ws_入力, "手配No. " & St_手配No & " を外注に出しました（外注注文 " & St_外注No & "）。戻ってきたら入力_外注で受入を計上します")
    Ws_入力.Range(St_内容セル).Value = ""
End Sub

' 工程一覧を書き直す
'   引数：St_対象指図（空なら見出しだけ）
Private Sub Pr_工程一覧更新(ByVal St_対象指図 As String)
    Dim Va_行 As Variant
    Dim Va_出力() As Variant
    Dim Lg_添字 As Long
    Dim Lg_行 As Long
    Dim Bl_前が完了 As Boolean
    Dim Bl_次あり As Boolean
    Dim Db_投入 As Double
    Ws_入力.Range(Ws_入力.Cells(Lg_工程一覧行, 1), Ws_入力.Cells(Lg_工程一覧行 + 15, 12)).Clear
    Ws_入力.Cells(Lg_工程一覧行, 2).Value = "■ 工程一覧（製造指図：" & St_対象指図 & "。指図No の欄の指図を表示）"
    Ws_入力.Cells(Lg_工程一覧行, 2).Font.Bold = True
    Ws_入力.Range(Ws_入力.Cells(Lg_工程一覧行 + 1, 2), Ws_入力.Cells(Lg_工程一覧行 + 1, 11)).Value = _
        Array("工程", "手配No", "工程名", "作業区／外注先", "投入", "標準(分)", "状態", "良品", "不良", "時間(分)")
    With Ws_入力.Range(Ws_入力.Cells(Lg_工程一覧行 + 1, 2), Ws_入力.Cells(Lg_工程一覧行 + 1, 11))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    If St_対象指図 = "" Then Exit Sub
    Va_行 = Fn_指図の工程行(St_対象指図)
    If IsEmpty(Va_行) Then
        Ws_入力.Cells(Lg_工程一覧行 + 2, 2).Value = "（工程がありません）"
        Exit Sub
    End If
    ReDim Va_出力(1 To UBound(Va_行), 1 To 項_工程表示.次列 - 1)
    Bl_前が完了 = True
    For Lg_添字 = 1 To UBound(Va_行)
        Lg_行 = Va_行(Lg_添字)
        Va_出力(Lg_添字, 項_工程表示.工程順) = Ws_工程.Cells(Lg_行, 列_指図工程.工程順).Value
        Va_出力(Lg_添字, 項_工程表示.手配No) = Ws_工程.Cells(Lg_行, 列_指図工程.手配No).Value
        Va_出力(Lg_添字, 項_工程表示.工程名) = Ws_工程.Cells(Lg_行, 列_指図工程.工程名).Value
        ' 投入数は、前の工程が終わっていれば分かる（HTML 版と同じく、分からなければ空）
        If Bl_前が完了 Then
            Db_投入 = Fn_投入数(Lg_行)
            Va_出力(Lg_添字, 項_工程表示.投入) = Db_投入
        Else
            Va_出力(Lg_添字, 項_工程表示.投入) = ""
        End If
        If Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.外注)) = St_印 Then
            Va_出力(Lg_添字, 項_工程表示.場所) = "外注：" & Ws_工程.Cells(Lg_行, 列_指図工程.外注先).Value
            Va_出力(Lg_添字, 項_工程表示.標準) = "LT " & Ws_工程.Cells(Lg_行, 列_指図工程.外注LT).Value & "日"
        Else
            Va_出力(Lg_添字, 項_工程表示.場所) = Fn_作業区名(Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.作業区).Value))
            If Bl_前が完了 Then Va_出力(Lg_添字, 項_工程表示.標準) = Fn_標準時間(Lg_行, Db_投入) Else Va_出力(Lg_添字, 項_工程表示.標準) = ""
        End If
        Va_出力(Lg_添字, 項_工程表示.状態) = Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value
        If Not Bl_次あり And Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value) <> St_状態_完了 Then
            Va_出力(Lg_添字, 項_工程表示.状態) = Va_出力(Lg_添字, 項_工程表示.状態) & "（次）"
            Bl_次あり = True
        End If
        Va_出力(Lg_添字, 項_工程表示.良品) = Ws_工程.Cells(Lg_行, 列_指図工程.良品).Value
        Va_出力(Lg_添字, 項_工程表示.不良) = Ws_工程.Cells(Lg_行, 列_指図工程.不良).Value
        Va_出力(Lg_添字, 項_工程表示.時間) = Ws_工程.Cells(Lg_行, 列_指図工程.作業時間).Value
        Bl_前が完了 = (Fn_文字(Ws_工程.Cells(Lg_行, 列_指図工程.状態).Value) = St_状態_完了)
    Next Lg_添字
    Ws_入力.Range(Ws_入力.Cells(Lg_工程一覧行 + 2, 2), Ws_入力.Cells(Lg_工程一覧行 + 1 + UBound(Va_行), 11)).Value = Va_出力
End Sub

' 工程一覧を、指図No の欄の指図で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_工程一覧再表示()
    If Not Fn_シートあり(St_シート_入力製造) Then Exit Sub
    Call Pr_変数格納
    Call Pr_工程一覧更新(St_一覧指図)
End Sub

'==============================================================
' 入力_製造 シートに「工程（手配No.）」の欄を足す（Mod_初期化 から、Pr_入力製造シート作成 のあとに呼ぶ）
'==============================================================
Public Sub Pr_製造実績欄作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力製造)
    Call Pr_入力見出し(Ws_表, "B16", "工程の実績・外注（手配No.を入れて「手配No.を読む」→ 値を確かめてボタンを押す）")
    Call Pr_入力欄(Ws_表, St_手配Noセル, "手配No.", "", "")
    Call Pr_入力欄(Ws_表, St_良品セル, "良品数", "", "")
    Call Pr_入力欄(Ws_表, St_不良セル, "不良数", "", "")
    Call Pr_入力欄(Ws_表, St_時間セル, "作業時間（分・段取を含む）", "", "")
    Call Pr_ボタン作成(Ws_表, "E17", 2, "手配No.を読む", "Pr_手配No読込", "製造_手配No読込")
    Call Pr_ボタン作成(Ws_表, "G17", 2, "実績を登録", "Pr_工程実績登録", "製造_実績登録")
    Call Pr_ボタン作成(Ws_表, "I17", 2, "外注に出す", "Pr_外注に出す", "製造_外注に出す")
End Sub
