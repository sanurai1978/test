Attribute VB_Name = "Mod_メニュー"
Option Explicit
'==============================================================
' Mod_メニュー
' 【目的】メニューシートの操作（事業部の切替、各シートへの移動、
'         件数の表示）を受け持つ。
' 【業務背景】HTML 版の上部（事業部の切替ボタンと「受注残・要手配・
'         製造中・出荷待ち」の件数）と、機能タブの代わり。
'         事業部を切り替えると、以後の操作はすべてその事業部が対象になる。
' 【前提】メニューのボタンは Mod_初期化 が作る。移動ボタンの名前は
'         「移動_シート名」にしてあり、押されたボタンの名前から行き先を決める。
'==============================================================

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_メニュー As Worksheet        ' メニューシート
Private St_移動先 As String             ' 押された移動ボタンの行き先シート名

'==============================================================
' Pr_加工事業部に切替（メイン）  ボタン：メニュー「加工事業部」
'==============================================================
Public Sub Pr_加工事業部に切替()
    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_メニュー, St_シート_設定)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. メニューシートを格納（各部品で読み直さない）
    Call Pr_事業部設定(St_加工)                              ' 1. 現在の事業部を加工にする（以後の操作の対象を決める）
    Call Pr_画面最新化                                      ' 2. 件数と選択肢を加工事業部にする（状況を見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_組立事業部に切替（メイン）  ボタン：メニュー「組立事業部」
'==============================================================
Public Sub Pr_組立事業部に切替()
    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_メニュー, St_シート_設定)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. メニューシートを格納（各部品で読み直さない）
    Call Pr_事業部設定(St_組立)                              ' 1. 現在の事業部を組立にする（以後の操作の対象を決める）
    Call Pr_画面最新化                                      ' 2. 件数と選択肢を組立事業部にする（状況を見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_シート移動（メイン）  ボタン：メニューの各シート名のボタン
'==============================================================
Public Sub Pr_シート移動()
    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_変数格納                                        ' 0. 押されたボタンから行き先を格納（どのシートへ行くか決める）
    Call Pr_移動先チェック                                  ' 1. 行き先のシートがあるか確認（無いシートへ行かない）→NGなら中断
    Call Pr_移動                                            ' 2. 行き先のシートを開く（一覧を見る・入力するため）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
'   St_移動先：ボタンから呼ばれたときだけ、ボタン名「移動_○○」の ○○ を入れる
Private Sub Pr_変数格納()
    Set Ws_メニュー = ThisWorkbook.Worksheets(St_シート_メニュー)
    St_移動先 = Fn_呼び出し元の移動先()
End Sub

' 押されたボタンの名前から行き先のシート名を取り出す（ボタン以外から呼ばれたら空）
Private Function Fn_呼び出し元の移動先() As String
    Dim St_名前 As String
    St_名前 = ""
    On Error Resume Next
    St_名前 = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(St_名前, Len(St_移動ボタン接頭)) = St_移動ボタン接頭 Then
        Fn_呼び出し元の移動先 = Mid$(St_名前, Len(St_移動ボタン接頭) + 1)
    Else
        Fn_呼び出し元の移動先 = ""
    End If
End Function

' 行き先のシートがあるか確かめる。無ければ中断
Private Sub Pr_移動先チェック()
    If St_移動先 = "" Or Not Fn_シートあり(St_移動先) Then
        MsgBox "移動先のシート「" & St_移動先 & "」がありません。", vbExclamation
        Err.Raise Lg_中断番号
    End If
End Sub

' 行き先のシートを開く
Private Sub Pr_移動()
    ThisWorkbook.Worksheets(St_移動先).Activate
    ThisWorkbook.Worksheets(St_移動先).Range("A1").Select
End Sub

' 現在の事業部を設定シートとメニューに書く
'   引数：St_事業部（加工／組立）
Private Sub Pr_事業部設定(ByVal St_事業部 As String)
    Call Pr_設定値書込("現在の事業部", St_事業部)
    Ws_メニュー.Range(St_メニュー_事業部セル).Value = Fn_事業部名(St_事業部)
End Sub

' 画面を最新にする（メニューの件数、入力シートの一覧、入力欄の選択肢）
'   各モジュールのメインが、データを変えたあとの最後に呼ぶ（Public）。
'   HTML 版は画面を描くたびに件数や一覧を作り直すので、それに合わせる。
Public Sub Pr_画面最新化()
    Call Pr_メニュー件数更新
    Call Pr_選択肢更新
    Call Pr_入力見込再表示
    Call Pr_入力受注再表示
    Call Pr_入力製造再表示
    Call Pr_入力発注再表示
    Call Pr_工程一覧再表示
    Call Pr_入力ピッキング再表示
    Call Pr_入力外注再表示
    Call Pr_入力出荷再表示
    Call Pr_入力在庫調整再表示
End Sub

' メニューの件数（受注残・要手配・製造中・出荷待ち）を現在の事業部の値で書き直す
'   要手配は MRP を計算して数える（設定「MRP安全在庫」の条件で）。
'   マスタに誤りがあって計算できないときは「マスタ要確認」と出す（メニューは止めない）。
Public Sub Pr_メニュー件数更新()
    Dim Ws_表示 As Worksheet
    Dim St_事業部 As String
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim Lg_受注残 As Long
    Dim Lg_出荷待ち As Long
    Dim Lg_製造中 As Long
    Dim St_状態 As String
    Set Ws_表示 = ThisWorkbook.Worksheets(St_シート_メニュー)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    ' 受注残（出荷済でない）と出荷待ち（引当済）
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_受注), 列_受注.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_受注.事業部)) = St_事業部 Then
            St_状態 = Fn_文字(Va_表(Lg_行, 列_受注.状態))
            If St_状態 <> St_状態_出荷済 Then Lg_受注残 = Lg_受注残 + 1
            If St_状態 = St_状態_引当済 Then Lg_出荷待ち = Lg_出荷待ち + 1
        End If
    Next Lg_行
    ' 製造中（計画・着手）
    Va_表 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_製造指図), 列_製造指図.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_製造指図.事業部)) = St_事業部 Then
            St_状態 = Fn_文字(Va_表(Lg_行, 列_製造指図.状態))
            If St_状態 = St_状態_計画 Or St_状態 = St_状態_着手 Then Lg_製造中 = Lg_製造中 + 1
        End If
    Next Lg_行
    Ws_表示.Range(St_メニュー_事業部セル).Value = Fn_事業部名(St_事業部)
    Ws_表示.Range(St_メニュー_受注残セル).Value = Lg_受注残
    Ws_表示.Range(St_メニュー_製造中セル).Value = Lg_製造中
    Ws_表示.Range(St_メニュー_出荷待ちセル).Value = Lg_出荷待ち
    Ws_表示.Range(St_メニュー_要手配セル).Value = Fn_要手配表示(St_事業部)
End Sub

' 要手配の件数（計算できなければ「マスタ要確認」）
Private Function Fn_要手配表示(ByVal St_事業部 As String) As Variant
    Dim Lg_件数 As Long
    On Error GoTo 計算不可
    Lg_件数 = Fn_要手配件数(St_事業部)
    Fn_要手配表示 = Lg_件数
    Exit Function
計算不可:
    Fn_要手配表示 = "マスタ要確認"
End Function

' 入力欄のドロップダウンの候補を、現在の事業部の品目で作り直す（隠しシート「選択肢」）
'   製品＝加工品・完成品、見込品＝見込生産の製品、製造品＝中間品も含む製造品、購入品＝材料・部品。
'   品目シートを直接編集しても、次の操作で候補が最新になるようにするため。
Public Sub Pr_選択肢更新()
    Dim Ws_選択肢 As Worksheet
    Dim Va_品目 As Variant
    Dim Lg_行 As Long
    Dim Lg_数(1 To 4) As Long
    Dim St_事業部 As String
    Dim St_区分 As String
    Dim St_品番 As String
    Dim Lg_列 As Long
    If Not Fn_シートあり(St_シート_選択肢) Then Exit Sub
    Set Ws_選択肢 = ThisWorkbook.Worksheets(St_シート_選択肢)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Va_品目 = Fn_表読込(ThisWorkbook.Worksheets(St_シート_品目), 列_品目.次列 - 1)
    Ws_選択肢.Cells.Clear
    Ws_選択肢.Range("A1:D1").Value = Array("製品", "見込品", "製造品", "購入品")
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_事業部 Then
            St_区分 = Fn_文字(Va_品目(Lg_行, 列_品目.区分))
            St_品番 = Fn_文字(Va_品目(Lg_行, 列_品目.品番))
            If Fn_製造品か(St_区分) Then
                Call Pr_候補追加(Ws_選択肢, 3, Lg_数(3), St_品番)
                If St_区分 <> "中間品" Then
                    Call Pr_候補追加(Ws_選択肢, 1, Lg_数(1), St_品番)
                    If Fn_文字(Va_品目(Lg_行, 列_品目.生産方式)) = St_見込生産 Then Call Pr_候補追加(Ws_選択肢, 2, Lg_数(2), St_品番)
                End If
            Else
                Call Pr_候補追加(Ws_選択肢, 4, Lg_数(4), St_品番)
            End If
        End If
    Next Lg_行
    ' 名前（入力欄の入力規則が参照する）を作り直す。候補が0件でも2行目を指す
    For Lg_列 = 1 To 4
        Call Pr_名前設定("選択肢_" & Ws_選択肢.Cells(1, Lg_列).Value, Ws_選択肢.Range(Ws_選択肢.Cells(2, Lg_列), Ws_選択肢.Cells(IIf(Lg_数(Lg_列) < 1, 1, Lg_数(Lg_列)) + 1, Lg_列)))
    Next Lg_列
End Sub

' 候補を1つ足す（列の次の行に書く）
Private Sub Pr_候補追加(ByVal Ws_選択肢 As Worksheet, ByVal Lg_列 As Long, ByRef Lg_件数 As Long, ByVal St_品番 As String)
    Lg_件数 = Lg_件数 + 1
    Ws_選択肢.Cells(Lg_件数 + 1, Lg_列).Value = St_品番
End Sub

' ブックの名前を作り直す
Private Sub Pr_名前設定(ByVal St_名前 As String, ByVal Rg_範囲 As Range)
    On Error Resume Next
    ThisWorkbook.Names(St_名前).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=St_名前, RefersTo:="='" & Rg_範囲.Worksheet.Name & "'!" & Rg_範囲.Address
End Sub
