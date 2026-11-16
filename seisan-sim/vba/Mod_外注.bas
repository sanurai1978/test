Attribute VB_Name = "Mod_外注"
Option Explicit
'==============================================================
' Mod_外注
' 【目的】入力_外注 シートの操作（外注受入）を受け持つ。
' 【業務背景】HTML 版の「外注」タブと同じ。外注に出した仕掛品が戻ってきたら、
'   検査して良品・不良を入力する。良品＋不良 は外注に出した数と一致させる。
'   受入で工程が完了し、最終工程なら製造指図が完成する（社内工程と同じ扱い）。
'   仕掛品を預けているだけなので、外注に出す・受け入れるで在庫は動かない。
'   外注加工費（実績）＝ 外注単価 × 良品数。
' 【前提】入力_外注 シートは初期化で作る（Pr_入力外注シート作成）。
'         外注に出す操作は入力_製造（手配No.）で行う。
'==============================================================

' ---- 入力_外注 シートの配置 ----
Private Const St_対象セル As String = "C6"          ' 外注No
Private Const St_良品セル As String = "C7"          ' 良品数
Private Const St_不良セル As String = "C8"          ' 不良数
Private Const Lg_一覧行 As Long = 12                ' 外注注文一覧を書く行

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_外注 シート
Private Ws_外注 As Worksheet            ' 外注注文シート
Private Ws_工程 As Worksheet            ' 指図工程シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private St_対象 As String               ' C6
Private Va_良品 As Variant              ' C7
Private Va_不良 As Variant              ' C8

'==============================================================
' Pr_外注受入（メイン）  ボタン：入力_外注「受入を計上」
'==============================================================
Public Sub Pr_外注受入()
    Dim Lg_行 As Long
    Dim Lg_工程行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力外注, St_シート_外注注文, St_シート_指図工程)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_対象行()                                     ' 1. 外注No の行を探す（どの外注か決める）→無ければ中断
    Lg_工程行 = Fn_工程行(Lg_行)                            ' 2. 外注中の工程を探す（どの工程を完了するか決める）→NGなら中断
    Call Pr_数量チェック(Lg_行)                             ' 3. 良品＋不良＝外注数か確認（数を合わせる）→NGなら中断
    Call Pr_画面更新停止                                    ' 3b. 画面更新を止める（速くする）
    Call Pr_受入書込(Lg_行, Lg_工程行)                      ' 4. 外注注文と工程に受入の実績を書く（工程を完了にする）
    Call Pr_完了後処理(Lg_行, Lg_工程行)                    ' 5. 最終工程なら完成・入庫・自動引当（製品にする）
    Call Pr_受入結果表示(Lg_行)                             ' 6. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 7. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力外注)
    Set Ws_外注 = ThisWorkbook.Worksheets(St_シート_外注注文)
    Set Ws_工程 = ThisWorkbook.Worksheets(St_シート_指図工程)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    St_対象 = Fn_文字(Ws_入力.Range(St_対象セル).Value)
    Va_良品 = Ws_入力.Range(St_良品セル).Value
    Va_不良 = Ws_入力.Range(St_不良セル).Value
End Sub

' 外注No の行を探す。外注中でなければ中断
Private Function Fn_対象行() As Long
    Dim Lg_行 As Long
    If St_対象 = "" Then Call Pr_入力エラー("外注No（例：SC-0001）を入力してください。")
    Lg_行 = Fn_行検索(Ws_外注, 列_外注注文.次列 - 1, St_事業部, 列_外注注文.外注No, St_対象)
    If Lg_行 = 0 Then Call Pr_入力エラー("外注注文 " & St_対象 & " は " & Fn_事業部名(St_事業部) & " にありません。")
    If Fn_文字(Ws_外注.Cells(Lg_行, 列_外注注文.状態).Value) <> St_状態_外注中 Then
        Call Pr_入力エラー("外注注文 " & St_対象 & " はすでに受入済です。")
    End If
    Fn_対象行 = Lg_行
End Function

' 外注注文の工程（指図No＋工程順）の行を探す
Private Function Fn_工程行(ByVal Lg_行 As Long) As Long
    Dim Va_表 As Variant
    Dim Lg_工程 As Long
    Va_表 = Fn_表読込(Ws_工程, 列_指図工程.次列 - 1)
    For Lg_工程 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_工程, 列_指図工程.事業部)) = St_事業部 _
           And Fn_文字(Va_表(Lg_工程, 列_指図工程.指図No)) = Fn_文字(Ws_外注.Cells(Lg_行, 列_外注注文.指図No).Value) _
           And Fn_数(Va_表(Lg_工程, 列_指図工程.工程順)) = Fn_数(Ws_外注.Cells(Lg_行, 列_外注注文.工程順).Value) Then
            Fn_工程行 = Lg_工程 + Lg_開始行 - 1
            Exit Function
        End If
    Next Lg_工程
    Call Pr_入力エラー("外注注文の工程が製造指図に見つかりません。")
End Function

' 良品・不良を確かめる（0以上の整数、良品＋不良＝外注に出した数）
Private Sub Pr_数量チェック(ByVal Lg_行 As Long)
    Dim Db_数 As Double
    If Not Fn_整数か(Va_良品) Or Not Fn_整数か(Va_不良) Then Call Pr_入力エラー("良品数・不良数は0以上の整数で入力してください。")
    Db_数 = Fn_数(Ws_外注.Cells(Lg_行, 列_外注注文.数量).Value)
    If CDbl(Va_良品) + CDbl(Va_不良) <> Db_数 Then
        Call Pr_入力エラー("良品数＋不良数を外注に出した数（" & Fn_表示数(Db_数) & "）と一致させてください（いま " & Fn_表示数(CDbl(Va_良品) + CDbl(Va_不良)) & "）。")
    End If
End Sub

' 0以上の整数か
Private Function Fn_整数か(ByVal Va_値 As Variant) As Boolean
    Fn_整数か = False
    If Fn_文字(Va_値) = "" Or Not IsNumeric(Va_値) Then Exit Function
    Fn_整数か = (CDbl(Va_値) >= 0 And CDbl(Va_値) = Int(CDbl(Va_値)))
End Function

' 外注注文と工程に受入の実績を書く（工程の作業時間は 0。社内の時間を使わないため）
Private Sub Pr_受入書込(ByVal Lg_行 As Long, ByVal Lg_工程行 As Long)
    Ws_外注.Cells(Lg_行, 列_外注注文.状態).Value = St_状態_受入済
    Ws_外注.Cells(Lg_行, 列_外注注文.良品).Value = CDbl(Va_良品)
    Ws_外注.Cells(Lg_行, 列_外注注文.不良).Value = CDbl(Va_不良)
    Ws_工程.Cells(Lg_工程行, 列_指図工程.良品).Value = CDbl(Va_良品)
    Ws_工程.Cells(Lg_工程行, 列_指図工程.不良).Value = CDbl(Va_不良)
    Ws_工程.Cells(Lg_工程行, 列_指図工程.作業時間).Value = 0
    Ws_工程.Cells(Lg_工程行, 列_指図工程.状態).Value = St_状態_完了
End Sub

' 最終工程なら完成・入庫・自動引当
Private Sub Pr_完了後処理(ByVal Lg_行 As Long, ByVal Lg_工程行 As Long)
    Dim St_指図No As String
    Dim Va_表 As Variant
    Dim Lg_工程 As Long
    Dim Lg_最終 As Long
    St_指図No = Fn_文字(Ws_外注.Cells(Lg_行, 列_外注注文.指図No).Value)
    ' この指図の最後の工程の行
    Va_表 = Fn_表読込(Ws_工程, 列_指図工程.次列 - 1)
    For Lg_工程 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_工程, 列_指図工程.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_工程, 列_指図工程.指図No)) = St_指図No Then
            Lg_最終 = Lg_工程 + Lg_開始行 - 1
        End If
    Next Lg_工程
    Call Pr_工程完了後処理(St_事業部, St_指図No, Lg_最終 = Lg_工程行, CDbl(Va_良品))
End Sub

' 結果欄への表示
Private Sub Pr_受入結果表示(ByVal Lg_行 As Long)
    Call Pr_結果表示(Ws_入力, "外注注文 " & St_対象 & " の受入を計上しました（外注加工費 " & _
        Format$(Fn_数(Ws_外注.Cells(Lg_行, 列_外注注文.単価).Value) * CDbl(Va_良品), "#,##0") & " 円）")
End Sub

' 外注注文一覧を書き直す
Private Sub Pr_一覧更新()
    Call Pr_一覧転記(Ws_入力, Lg_一覧行, St_シート_外注注文, 列_外注注文.次列 - 1, St_事業部, "外注注文一覧")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力外注再表示()
    If Not Fn_シートあり(St_シート_入力外注) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_外注 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力外注シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力外注)
    Call Pr_入力シート枠(Ws_表, "外注", _
        "外注の工程は入力_製造（手配No.）で「外注に出す」と外注注文になります。戻ってきたらここで良品・不良を入れて受け入れます。在庫は動きません。")
    Call Pr_入力見出し(Ws_表, "B5", "外注受入")
    Call Pr_入力欄(Ws_表, St_対象セル, "外注No", "", "")
    Call Pr_入力欄(Ws_表, St_良品セル, "良品数", "", "")
    Call Pr_入力欄(Ws_表, St_不良セル, "不良数", 0, "")
    Call Pr_ボタン作成(Ws_表, "E6", 2, "受入を計上", "Pr_外注受入", "外注_受入")
End Sub
