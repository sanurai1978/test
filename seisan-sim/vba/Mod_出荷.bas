Attribute VB_Name = "Mod_出荷"
Option Explicit
'==============================================================
' Mod_出荷
' 【目的】入力_出荷 シートの操作（出荷のピッキング指示、出荷の取消）と、
'         出荷待ち（引当済の受注）の表示を受け持つ。
' 【業務背景】HTML 版の「出荷」タブと同じ。
'   ・出荷待ちの受注（引当済）に、出荷数と配送方法を決めてピッキング指示を出す。
'     出荷数を受注残より少なくすると分納になる。指示を出せるのは
'     受注残 － すでにピッキング中の数 まで、かつ今の在庫まで。
'   ・ピッキングを確定すると出荷になる（入力_ピッキング）。
'   ・出荷の取消：在庫を戻し、受注の出荷済を減らして「引当済」に戻す。
'     出荷で消し込んだ見込は戻さない（HTML 版と同じ）。
' 【前提】入力_出荷 シートは初期化で作る（Pr_入力出荷シート作成）。
'==============================================================

' ---- 入力_出荷 シートの配置 ----
Private Const St_受注セル As String = "C6"          ' 受注No
Private Const St_数量セル As String = "C7"          ' 今回の出荷数（空なら 受注残－ピッキング中）
Private Const St_配送セル As String = "C8"          ' 配送方法
Private Const St_出荷Noセル As String = "C11"       ' 出荷の取消：出荷No
Private Const Lg_待ち行 As Long = 14                ' 出荷待ちの表を書く行
Private Const St_配送方法一覧 As String = "自社便,路線便,宅配便"   ' HTML 版と同じ3種

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Ws_入力 As Worksheet            ' 入力_出荷 シート
Private Ws_受注 As Worksheet            ' 受注シート
Private Ws_出荷 As Worksheet            ' 出荷実績シート
Private St_事業部 As String             ' 設定「現在の事業部」
Private St_受注No As String             ' C6
Private Va_数量 As Variant              ' C7
Private St_配送方法 As String           ' C8
Private St_出荷No As String             ' C11

'==============================================================
' Pr_出荷ピッキング指示（メイン）  ボタン：入力_出荷「ピッキング指示を出す」
'==============================================================
Public Sub Pr_出荷ピッキング指示()
    Dim Lg_行 As Long
    Dim Db_数量 As Double
    Dim St_ピッキングNo As String

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力出荷, St_シート_受注, St_シート_ピッキング)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_出荷待ち受注行()                             ' 1. 引当済の受注を探す（出荷できる受注か決める）→NGなら中断
    Db_数量 = Fn_出荷数(Lg_行)                              ' 2. 出荷数を決めて確かめる（残数と在庫を超えない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    St_ピッキングNo = Fn_出荷ピッキング作成(Lg_行, Db_数量)  ' 3. 出荷用のピッキング指示を作る（製品を棚から集める）
    Call Pr_指示結果表示(St_ピッキングNo, Db_数量)          ' 4. 番号を結果欄に出す（入力_ピッキングで使う）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

'==============================================================
' Pr_出荷取消（メイン）  ボタン：入力_出荷「出荷を取り消す」
'==============================================================
Public Sub Pr_出荷取消()
    Dim Lg_行 As Long

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_入力出荷, St_シート_受注, St_シート_出荷実績)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. 入力欄と事業部を格納（各部品で読み直さない）
    Lg_行 = Fn_出荷行()                                     ' 1. 出荷No の行を探す（どの出荷か決める）→無ければ中断
    Call Pr_取消確認                                        ' 2. 取り消してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 2b. 画面更新を止める（速くする）
    Call Pr_出荷を戻す(Lg_行)                               ' 3. 在庫と受注の出荷済を戻し、出荷実績を消す（出荷前に戻す）
    Call Pr_取消結果表示                                    ' 4. 結果欄に出す（何をしたか伝える）
    Call Pr_画面最新化                                      ' 5. 件数・一覧・選択肢を更新（結果をその場で見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Set Ws_入力 = ThisWorkbook.Worksheets(St_シート_入力出荷)
    Set Ws_受注 = ThisWorkbook.Worksheets(St_シート_受注)
    Set Ws_出荷 = ThisWorkbook.Worksheets(St_シート_出荷実績)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    St_受注No = Fn_文字(Ws_入力.Range(St_受注セル).Value)
    Va_数量 = Ws_入力.Range(St_数量セル).Value
    St_配送方法 = Fn_文字(Ws_入力.Range(St_配送セル).Value)
    St_出荷No = Fn_文字(Ws_入力.Range(St_出荷Noセル).Value)
End Sub

' 受注No の行を探し、引当済か確かめる。戻り値：受注シートの行
Private Function Fn_出荷待ち受注行() As Long
    Dim Lg_行 As Long
    If St_受注No = "" Then Call Pr_入力エラー("出荷する受注No（例：SO-0003）を入力してください。")
    Lg_行 = Fn_行検索(Ws_受注, 列_受注.次列 - 1, St_事業部, 列_受注.受注No, St_受注No)
    If Lg_行 = 0 Then Call Pr_入力エラー("受注 " & St_受注No & " は " & Fn_事業部名(St_事業部) & " にありません。")
    If Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.状態).Value) <> St_状態_引当済 Then
        Call Pr_入力エラー("出荷できるのは「引当済」の受注だけです（いま：" & Ws_受注.Cells(Lg_行, 列_受注.状態).Value & "）。入力_受注で在庫を引き当ててください。")
    End If
    If InStr("," & St_配送方法一覧 & ",", "," & St_配送方法 & ",") = 0 Then
        Call Pr_入力エラー("配送方法は 自社便・路線便・宅配便 のどれかにしてください。")
    End If
    Fn_出荷待ち受注行 = Lg_行
End Function

' 今回の出荷数を決めて確かめる（空なら 受注残－ピッキング中）
'   戻り値：出荷数。受注残－ピッキング中 を超える、在庫を超える、数でない なら中断
Private Function Fn_出荷数(ByVal Lg_行 As Long) As Double
    Dim Va_品目 As Variant
    Dim Db_残 As Double
    Dim Db_中 As Double
    Dim Db_空き As Double
    Dim Db_数量 As Double
    Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.品番).Value))
    Db_残 = Fn_丸め(Fn_数(Ws_受注.Cells(Lg_行, 列_受注.数量).Value) - Fn_数(Ws_受注.Cells(Lg_行, 列_受注.出荷済).Value))
    Db_中 = Fn_出荷ピッキング中(St_事業部, St_受注No)
    Db_空き = Fn_丸め(Db_残 - Db_中)
    If Fn_文字(Va_数量) = "" Then
        Db_数量 = Db_空き
        If Db_数量 < 0 Then Db_数量 = 0
        Db_数量 = Fn_数量変換(Db_数量, Fn_小数可か(Va_品目))
    Else
        Db_数量 = Fn_数量変換(Va_数量, Fn_小数可か(Va_品目))
    End If
    If Db_数量 = 0 Then Call Pr_入力エラー(St_数量エラー)
    If Db_数量 > Db_空き Then
        Call Pr_入力エラー("ピッキング指示を出せるのは " & Fn_表示数(Db_空き) & " までです（受注残 " & Fn_表示数(Db_残) & "、うちピッキング中 " & Fn_表示数(Db_中) & "）。")
    End If
    If Db_数量 > Fn_数(Va_品目(列_品目.在庫数)) Then
        Call Pr_入力エラー("在庫が足りません（現在庫 " & Fn_表示数(Fn_数(Va_品目(列_品目.在庫数))) & "）。")
    End If
    Fn_出荷数 = Db_数量
End Function

' 出荷用のピッキング指示を作る。戻り値：ピッキングNo
'   内容は HTML 版と同じ「品名 数量単位 → 得意先」。
Private Function Fn_出荷ピッキング作成(ByVal Lg_行 As Long, ByVal Db_数量 As Double) As String
    Dim Va_品目 As Variant
    Dim St_品番 As String
    St_品番 = Fn_文字(Ws_受注.Cells(Lg_行, 列_受注.品番).Value)
    Va_品目 = Fn_品目情報(St_事業部, St_品番)
    Fn_出荷ピッキング作成 = Fn_ピッキング作成(St_事業部, "出荷", St_受注No, _
        Va_品目(列_品目.品名) & " " & Fn_表示数(Db_数量) & Va_品目(列_品目.単位) & " → " & Ws_受注.Cells(Lg_行, 列_受注.得意先).Value, _
        Array(St_品番), Array(Db_数量), St_配送方法)
End Function

' 出荷No の行を探す。戻り値：出荷実績シートの行
Private Function Fn_出荷行() As Long
    Dim Lg_行 As Long
    If St_出荷No = "" Then Call Pr_入力エラー("取り消す出荷No（例：SH-0001）を入力してください。")
    Lg_行 = Fn_行検索(Ws_出荷, 列_出荷実績.次列 - 1, St_事業部, 列_出荷実績.出荷No, St_出荷No)
    If Lg_行 = 0 Then Call Pr_入力エラー("出荷 " & St_出荷No & " は " & Fn_事業部名(St_事業部) & " にありません。")
    Fn_出荷行 = Lg_行
End Function

' 取り消してよいか確認する
Private Sub Pr_取消確認()
    Call Pr_削除確認("出荷 " & St_出荷No & " を取り消します（在庫を戻し、出荷実績の行を削除します）。")
End Sub

' 在庫を戻し、受注の出荷済を減らして引当済に戻し、出荷実績を消す
Private Sub Pr_出荷を戻す(ByVal Lg_行 As Long)
    Dim Lg_受注行 As Long
    Dim Db_数量 As Double
    Dim Db_出荷済 As Double
    Db_数量 = Fn_数(Ws_出荷.Cells(Lg_行, 列_出荷実績.数量).Value)
    Call Pr_在庫移動(St_事業部, Fn_文字(Ws_出荷.Cells(Lg_行, 列_出荷実績.品番).Value), "出荷取消", Db_数量, St_出荷No)
    Lg_受注行 = Fn_行検索(Ws_受注, 列_受注.次列 - 1, St_事業部, 列_受注.受注No, Fn_文字(Ws_出荷.Cells(Lg_行, 列_出荷実績.受注No).Value))
    If Lg_受注行 > 0 Then
        Db_出荷済 = Fn_丸め(Fn_数(Ws_受注.Cells(Lg_受注行, 列_受注.出荷済).Value) - Db_数量)
        If Db_出荷済 < 0 Then Db_出荷済 = 0
        Ws_受注.Cells(Lg_受注行, 列_受注.出荷済).Value = Db_出荷済
        Ws_受注.Cells(Lg_受注行, 列_受注.状態).Value = St_状態_引当済
    End If
    Call Pr_行削除(Ws_出荷, Array(Lg_行))
End Sub

' 結果欄への表示
Private Sub Pr_指示結果表示(ByVal St_ピッキングNo As String, ByVal Db_数量 As Double)
    Call Pr_結果表示(Ws_入力, "受注 " & St_受注No & "：" & Fn_表示数(Db_数量) & " のピッキング指示 " & St_ピッキングNo & _
        " を出しました。入力_ピッキングで集めて確定すると出荷になります")
End Sub

Private Sub Pr_取消結果表示()
    Call Pr_結果表示(Ws_入力, "出荷 " & St_出荷No & " を取り消しました")
End Sub

' 出荷待ち（引当済の受注）の表と、出荷実績の一覧を書き直す
Private Sub Pr_一覧更新()
    Dim Va_受注 As Variant
    Dim Lg_行 As Long
    Dim Lg_出力 As Long
    Dim Va_品目 As Variant
    Dim St_番号 As String
    Ws_入力.Range(Ws_入力.Cells(Lg_待ち行, 1), Ws_入力.Cells(Ws_入力.Rows.Count, 12)).Clear
    Ws_入力.Cells(Lg_待ち行, 2).Value = "■ 出荷待ち（引当済の受注。" & Fn_事業部名(St_事業部) & "）"
    Ws_入力.Cells(Lg_待ち行, 2).Font.Bold = True
    Ws_入力.Range(Ws_入力.Cells(Lg_待ち行 + 1, 2), Ws_入力.Cells(Lg_待ち行 + 1, 10)).Value = _
        Array("受注No", "品番", "品名", "得意先", "納期", "受注数", "出荷済", "残", "ピッキング中")
    With Ws_入力.Range(Ws_入力.Cells(Lg_待ち行 + 1, 2), Ws_入力.Cells(Lg_待ち行 + 1, 10))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    Lg_出力 = Lg_待ち行 + 2
    Va_受注 = Fn_表読込(Ws_受注, 列_受注.次列 - 1)
    For Lg_行 = Fn_行数(Va_受注) To 1 Step -1
        If Fn_文字(Va_受注(Lg_行, 列_受注.事業部)) = St_事業部 And Fn_文字(Va_受注(Lg_行, 列_受注.状態)) = St_状態_引当済 Then
            St_番号 = Fn_文字(Va_受注(Lg_行, 列_受注.受注No))
            Va_品目 = Fn_品目情報(St_事業部, Fn_文字(Va_受注(Lg_行, 列_受注.品番)))
            Ws_入力.Cells(Lg_出力, 2).Value = St_番号
            Ws_入力.Cells(Lg_出力, 3).Value = Va_受注(Lg_行, 列_受注.品番)
            If Not IsEmpty(Va_品目) Then Ws_入力.Cells(Lg_出力, 4).Value = Va_品目(列_品目.品名)
            Ws_入力.Cells(Lg_出力, 5).Value = Va_受注(Lg_行, 列_受注.得意先)
            Ws_入力.Cells(Lg_出力, 6).Value = Va_受注(Lg_行, 列_受注.納期)
            Ws_入力.Cells(Lg_出力, 6).NumberFormat = "yyyy/mm/dd"
            Ws_入力.Cells(Lg_出力, 7).Value = Va_受注(Lg_行, 列_受注.数量)
            Ws_入力.Cells(Lg_出力, 8).Value = Va_受注(Lg_行, 列_受注.出荷済)
            Ws_入力.Cells(Lg_出力, 9).Value = Fn_丸め(Fn_数(Va_受注(Lg_行, 列_受注.数量)) - Fn_数(Va_受注(Lg_行, 列_受注.出荷済)))
            Ws_入力.Cells(Lg_出力, 10).Value = Fn_出荷ピッキング中(St_事業部, St_番号)
            Lg_出力 = Lg_出力 + 1
        End If
    Next Lg_行
    If Lg_出力 = Lg_待ち行 + 2 Then
        Ws_入力.Cells(Lg_出力, 2).Value = "（出荷待ちはありません。入力_受注で在庫を引き当てると、ここに並びます）"
        Lg_出力 = Lg_出力 + 1
    End If
    Call Pr_一覧転記(Ws_入力, Lg_出力 + 1, St_シート_出荷実績, 列_出荷実績.次列 - 1, St_事業部, "出荷実績")
End Sub

' 入力シートの一覧を、現在の事業部で書き直す（メニューの Pr_画面最新化 から呼ぶ）
Public Sub Pr_入力出荷再表示()
    If Not Fn_シートあり(St_シート_入力出荷) Then Exit Sub
    Call Pr_変数格納
    Call Pr_一覧更新
End Sub

'==============================================================
' 入力_出荷 シートを作る（Mod_初期化 から呼ぶ）
'==============================================================
Public Sub Pr_入力出荷シート作成()
    Dim Ws_表 As Worksheet
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート_入力出荷)
    Call Pr_入力シート枠(Ws_表, "出荷", _
        "引当済の受注に出荷数を決めてピッキング指示を出します。出荷数を受注残より少なくすると分納。入力_ピッキングで確定すると出荷になります。")
    Call Pr_入力見出し(Ws_表, "B5", "出荷のピッキング指示")
    Call Pr_入力欄(Ws_表, St_受注セル, "受注No（引当済）", "", "")
    Call Pr_入力欄(Ws_表, St_数量セル, "今回の出荷数（空なら残り全部）", "", "")
    Call Pr_入力欄(Ws_表, St_配送セル, "配送方法", "路線便", "")
    Ws_表.Range(St_配送セル).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=St_配送方法一覧
    Call Pr_ボタン作成(Ws_表, "E6", 2, "ピッキング指示を出す", "Pr_出荷ピッキング指示", "出荷_指示")
    Call Pr_入力見出し(Ws_表, "B10", "出荷の取消")
    Call Pr_入力欄(Ws_表, St_出荷Noセル, "出荷No", "", "")
    Call Pr_ボタン作成(Ws_表, "E11", 2, "出荷を取り消す", "Pr_出荷取消", "出荷_取消")
End Sub
