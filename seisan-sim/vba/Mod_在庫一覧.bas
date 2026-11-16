Attribute VB_Name = "Mod_在庫一覧"
Option Explicit
'==============================================================
' Mod_在庫一覧
' 【目的】現在の事業部の品目ごとに、在庫・引当・有効在庫・手配済を
'         「在庫一覧」シートに書き出す。
' 【業務背景】HTML 版の「在庫」タブと同じ数字を出す。
'   有効在庫 ＝ 在庫 －（受注引当＋製造用確保）。
'   手配済   ＝ 製造中の見込み出来高＋発注残。
'   注意     ＝ 在庫0以下なら「欠品」、有効在庫が安全在庫より少なければ
'              「安全在庫割れ」。受注生産の品目は安全在庫を持たないので出さない。
' 【前提】初期化済み（必要なシートがそろっている）。
' 【処理の流れ】
'   0. 変数格納 → 1. 入力チェック → 2. 画面停止 → 3. 品目の計算列更新
'   → 4. 集計 → 5. 一覧書込 → 6. メニュー件数更新 → 7. 一覧を表示
'==============================================================

' ---- 一覧シートの配置 ----
Private Const St_題名セル As String = "A1"          ' 見出し（事業部名）
Private Const St_更新日時セル As String = "A2"      ' 更新日時
Private Const Lg_一覧見出し行 As Long = 4           ' 列見出しの行
Private Const Lg_一覧開始行 As Long = 5             ' データの開始行

' ---- モジュール変数（Pr_変数格納 で格納。複数の部品で使う値）----
Private Ws_在庫一覧 As Worksheet        ' 在庫一覧シート
Private Ws_品目 As Worksheet            ' 品目シート
Private St_事業部 As String             ' 設定「現在の事業部」（加工／組立）
Private Dt_実行日時 As Date             ' 実行を始めた日時（一覧に表示）

'==============================================================
' Pr_在庫一覧更新（メイン／目次）  ボタン：メニュー「在庫一覧」
'==============================================================
Public Sub Pr_在庫一覧更新()
    Dim Va_一覧 As Variant

    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_シート存在確認(Array(St_シート_在庫一覧, St_シート_品目, St_シート_設定)) ' 0a. シートの有無を確認（未初期化で動かさない）→無ければ中断
    Call Pr_変数格納                                        ' 0. シートと現在の事業部を格納（各部品で読み直さない）
    Call Pr_入力チェック                                    ' 1. 事業部の値を確認（誤った事業部で集計しない）→NGなら中断
    Call Pr_画面更新停止                                    ' 2. 画面更新を止める（速くする）
    Call Pr_品目計算列更新                                  ' 3. 標準原価・販売単価を再計算（マスタ編集を反映）
    Va_一覧 = Fn_一覧作成()                                 ' 4. 品目ごとの在庫・引当・有効在庫を集計（HTMLと同じ式）
    Call Pr_一覧書込(Va_一覧)                               ' 5. 在庫一覧シートに書く（一覧として見るため）
    Call Pr_画面最新化                                      ' 6. メニューの件数と選択肢を更新（状況を最新に）
    Call Pr_一覧表示                                        ' 7. 在庫一覧シートを開く（結果を見せる）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する（確認はしない。確認は Pr_入力チェック）
Private Sub Pr_変数格納()
    Set Ws_在庫一覧 = ThisWorkbook.Worksheets(St_シート_在庫一覧)
    Set Ws_品目 = ThisWorkbook.Worksheets(St_シート_品目)
    St_事業部 = Fn_文字(Fn_設定値("現在の事業部"))
    Dt_実行日時 = Now
End Sub

' 現在の事業部が「加工」か「組立」であることを確かめる。違えば中断
Private Sub Pr_入力チェック()
    If St_事業部 <> St_加工 And St_事業部 <> St_組立 Then
        MsgBox "設定シートの「現在の事業部」が「加工」か「組立」になっていません（いま：" & St_事業部 & "）。" & vbCrLf & _
               "メニューの事業部切替ボタンを押してください。", vbExclamation
        Err.Raise Lg_中断番号
    End If
End Sub

' 品目ごとの一覧の行を作る
'   戻り値：2次元配列（行×列_在庫一覧）。品目が0件なら Empty
'   使うモジュール変数：Ws_品目、St_事業部
Private Function Fn_一覧作成() As Variant
    Dim Va_品目 As Variant
    Dim Va_一覧 As Variant
    Dim Ob_受注引当 As Object
    Dim Ob_製造用 As Object
    Dim Ob_手配済 As Object
    Dim Lg_行 As Long
    Dim Lg_件数 As Long
    Dim Lg_出力行 As Long
    Dim St_品番 As String
    Dim Db_在庫 As Double
    Dim Db_引当 As Double
    Dim Db_有効 As Double
    Dim Bl_受注生産 As Boolean

    Va_品目 = Fn_表読込(Ws_品目, 列_品目.次列 - 1)
    ' この事業部の品目数を数える（出力配列の大きさを決めるため）
    Lg_件数 = 0
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_事業部 Then Lg_件数 = Lg_件数 + 1
    Next Lg_行
    If Lg_件数 = 0 Then
        Fn_一覧作成 = Empty
        Exit Function
    End If

    ' 引当・手配済を品番ごとに集計（Dictionary で品番から引けるようにする）
    Set Ob_受注引当 = Fn_受注引当集計(St_事業部)
    Set Ob_製造用 = Fn_製造用確保集計(St_事業部)
    Set Ob_手配済 = Fn_手配済集計(St_事業部)

    ReDim Va_一覧(1 To Lg_件数, 1 To 列_在庫一覧.次列 - 1)
    Lg_出力行 = 0
    For Lg_行 = 1 To Fn_行数(Va_品目)
        If Fn_文字(Va_品目(Lg_行, 列_品目.事業部)) = St_事業部 Then
            Lg_出力行 = Lg_出力行 + 1
            St_品番 = Fn_文字(Va_品目(Lg_行, 列_品目.品番))
            Db_在庫 = Fn_数(Va_品目(Lg_行, 列_品目.在庫数))
            Db_引当 = Fn_丸め(Fn_辞書値(Ob_受注引当, St_品番) + Fn_辞書値(Ob_製造用, St_品番))
            Db_有効 = Fn_丸め(Db_在庫 - Db_引当)
            Bl_受注生産 = Fn_製造品か(Fn_文字(Va_品目(Lg_行, 列_品目.区分))) _
                          And Fn_文字(Va_品目(Lg_行, 列_品目.生産方式)) = St_受注生産
            Va_一覧(Lg_出力行, 列_在庫一覧.品番) = St_品番
            Va_一覧(Lg_出力行, 列_在庫一覧.品名) = Va_品目(Lg_行, 列_品目.品名)
            Va_一覧(Lg_出力行, 列_在庫一覧.区分) = Va_品目(Lg_行, 列_品目.区分)
            Va_一覧(Lg_出力行, 列_在庫一覧.生産方式) = Va_品目(Lg_行, 列_品目.生産方式)
            Va_一覧(Lg_出力行, 列_在庫一覧.倉庫) = Va_品目(Lg_行, 列_品目.倉庫)
            Va_一覧(Lg_出力行, 列_在庫一覧.棚番) = Va_品目(Lg_行, 列_品目.棚番)
            Va_一覧(Lg_出力行, 列_在庫一覧.在庫) = Db_在庫
            Va_一覧(Lg_出力行, 列_在庫一覧.単位) = Va_品目(Lg_行, 列_品目.単位)
            Va_一覧(Lg_出力行, 列_在庫一覧.受注引当) = Fn_辞書値(Ob_受注引当, St_品番)
            Va_一覧(Lg_出力行, 列_在庫一覧.製造用確保) = Fn_辞書値(Ob_製造用, St_品番)
            Va_一覧(Lg_出力行, 列_在庫一覧.引当計) = Db_引当
            Va_一覧(Lg_出力行, 列_在庫一覧.有効在庫) = Db_有効
            Va_一覧(Lg_出力行, 列_在庫一覧.手配済) = Fn_辞書値(Ob_手配済, St_品番)
            ' 安全在庫と注意（受注生産は安全在庫を持たないので「－」、注意も出さない）
            If Bl_受注生産 Then
                Va_一覧(Lg_出力行, 列_在庫一覧.安全在庫) = "－"
                Va_一覧(Lg_出力行, 列_在庫一覧.注意) = ""
            Else
                Va_一覧(Lg_出力行, 列_在庫一覧.安全在庫) = Fn_数(Va_品目(Lg_行, 列_品目.安全在庫))
                Va_一覧(Lg_出力行, 列_在庫一覧.注意) = Fn_注意(Db_在庫, Db_有効, Fn_数(Va_品目(Lg_行, 列_品目.安全在庫)))
            End If
        End If
    Next Lg_行
    Fn_一覧作成 = Va_一覧
End Function

' 注意の文言（HTML 版の在庫タブと同じ判定順：欠品が先、次に安全在庫割れ）
Private Function Fn_注意(ByVal Db_在庫 As Double, ByVal Db_有効 As Double, ByVal Db_安全在庫 As Double) As String
    If Db_在庫 <= 0 Then
        Fn_注意 = "欠品"
    ElseIf Db_有効 < Db_安全在庫 Then
        Fn_注意 = "安全在庫割れ"
    Else
        Fn_注意 = ""
    End If
End Function

' 在庫一覧シートを書き直す（前回の内容を消してから書く）
'   引数：Va_一覧（Fn_一覧作成 の戻り値）
'   使うモジュール変数：Ws_在庫一覧、St_事業部、Dt_実行日時
Private Sub Pr_一覧書込(ByVal Va_一覧 As Variant)
    Dim Lg_列数 As Long
    Lg_列数 = 列_在庫一覧.次列 - 1
    Ws_在庫一覧.Cells.Clear
    Ws_在庫一覧.Range(St_題名セル).Value = "在庫一覧　" & Fn_事業部名(St_事業部)
    Ws_在庫一覧.Range(St_題名セル).Font.Bold = True
    Ws_在庫一覧.Range(St_題名セル).Font.Size = 14
    Ws_在庫一覧.Range(St_更新日時セル).Value = "更新：" & Format$(Dt_実行日時, "yyyy/mm/dd hh:nn") & _
        "　　有効在庫 ＝ 在庫 －（受注引当＋製造用確保）。手配済 ＝ 製造中の見込み出来高＋発注残。"
    ' 列見出し（並びは 列_在庫一覧 と同じ）
    Ws_在庫一覧.Range(Ws_在庫一覧.Cells(Lg_一覧見出し行, 1), Ws_在庫一覧.Cells(Lg_一覧見出し行, Lg_列数)).Value = _
        Array("品番", "品名", "区分", "生産方式", "倉庫", "棚番", "在庫", "単位", "受注引当", "製造用確保", "引当計", "有効在庫", "手配済", "安全在庫", "注意")
    With Ws_在庫一覧.Range(Ws_在庫一覧.Cells(Lg_一覧見出し行, 1), Ws_在庫一覧.Cells(Lg_一覧見出し行, Lg_列数))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    If IsEmpty(Va_一覧) Then
        Ws_在庫一覧.Cells(Lg_一覧開始行, 1).Value = "品目がありません"
        Exit Sub
    End If
    Ws_在庫一覧.Range(Ws_在庫一覧.Cells(Lg_一覧開始行, 1), _
        Ws_在庫一覧.Cells(Lg_一覧開始行 + UBound(Va_一覧, 1) - 1, Lg_列数)).Value = Va_一覧
    ' 数量の列は桁区切り、小数は必要なときだけ
    Ws_在庫一覧.Range(Ws_在庫一覧.Cells(Lg_一覧開始行, 列_在庫一覧.在庫), _
        Ws_在庫一覧.Cells(Lg_一覧開始行 + UBound(Va_一覧, 1) - 1, 列_在庫一覧.在庫)).NumberFormat = "#,##0.##"
    Ws_在庫一覧.Range(Ws_在庫一覧.Cells(Lg_一覧開始行, 列_在庫一覧.受注引当), _
        Ws_在庫一覧.Cells(Lg_一覧開始行 + UBound(Va_一覧, 1) - 1, 列_在庫一覧.安全在庫)).NumberFormat = "#,##0.##"
    Call Pr_注意の色付け(UBound(Va_一覧, 1))
    Ws_在庫一覧.Columns("A:O").AutoFit
End Sub

' 注意のある行を色で目立たせる（欠品＝赤、安全在庫割れ＝黄）
Private Sub Pr_注意の色付け(ByVal Lg_件数 As Long)
    Dim Lg_行 As Long
    Dim St_注意 As String
    For Lg_行 = Lg_一覧開始行 To Lg_一覧開始行 + Lg_件数 - 1
        St_注意 = Fn_文字(Ws_在庫一覧.Cells(Lg_行, 列_在庫一覧.注意).Value)
        If St_注意 = "欠品" Then
            Ws_在庫一覧.Cells(Lg_行, 列_在庫一覧.注意).Interior.Color = RGB(255, 199, 206)
        ElseIf St_注意 = "安全在庫割れ" Then
            Ws_在庫一覧.Cells(Lg_行, 列_在庫一覧.注意).Interior.Color = RGB(255, 235, 156)
        End If
    Next Lg_行
End Sub

' 在庫一覧シートを開く
Private Sub Pr_一覧表示()
    Ws_在庫一覧.Activate
    Ws_在庫一覧.Range("A1").Select
End Sub
