Attribute VB_Name = "Mod_初期化"
Option Explicit
'==============================================================
' Mod_初期化
' 【目的】1回実行すれば、シート・見出し・初期データ・メニューのボタンが
'         すべてできる初期化マクロ。メニューの「初期データに戻す」も同じ。
' 【業務背景】HTML 版の seed（初期データ）と同じ内容を入れる。
'         日付は HTML 版と同じく「初期化した日から何日後」で入れる
'         （例：受注 SO-0001 の納期＝初期化日＋10日）。
'         手配No.は HTML 版と同じ順（加工の MO-0001 → MO-0002 →
'         組立の MO-0001 → MO-0002、工程順）に 1000001 から振る。
' 【前提】・このモジュールと Mod_共通定義 ほか全モジュールを取り込み済み。
'         ・既存のシートはすべて消して作り直す（確認を出す）。
' 【注意】破壊的な操作：ブック内のシートをすべて削除して作り直す。
' 【処理の流れ】
'   0. 変数格納 → 1. 実行確認 → 2. 画面停止 → 3. シート作り直し
'   → 4. 見出し → 5. マスタ投入 → 6. 採番・設定 → 7. 伝票投入
'   → 8. 書式 → 9. メニュー・入力シート作成 → 10. 計算列・件数 → 11. 完了表示
'==============================================================

Private Const St_作業用シート As String = "作業用_初期化"   ' 全シート削除のあいだだけ置くシート

' ---- モジュール変数（Pr_変数格納 で格納）----
Private Dt_今日 As Date                 ' 初期化した日（初期データの日付の基準）
Private Bl_既存データあり As Boolean    ' 品目シートがすでにあるか（上書き確認を出すか）

'==============================================================
' Pr_初期化（メイン／目次）  ボタン：メニュー「初期データに戻す」
'   初回は VBE から実行する（Alt+F8 → Pr_初期化）。
'==============================================================
Public Sub Pr_初期化()
    On Error GoTo 終了処理                                  ' 中断・エラー・正常とも終了処理へ（後始末を1か所に）
    Call Pr_変数格納                                        ' 0. 今日の日付と既存データの有無を格納（初期データの基準）
    Call Pr_実行確認                                        ' 1. 全データを消してよいか確認（誤操作防止）→いいえなら中断
    Call Pr_画面更新停止                                    ' 2. 画面更新を止める（速くする）
    Call Pr_シート作り直し                                  ' 3. 全シートを消して決まった順に作る（構成をそろえる）
    Call Pr_見出し作成                                      ' 4. データシートの見出しを書く（列挙体と同じ並び）→不一致なら中断
    Call Pr_マスタ投入                                      ' 5. 品目・BOM・作業区・工程表を入れる（HTML の seed と同じ）
    Call Pr_採番設定投入                                    ' 6. 採番の初期値と設定を入れる（番号と事業部の起点）
    Call Pr_伝票投入                                        ' 7. 見込・受注・製造指図などを入れる（途中の業務状態を再現）
    Call Pr_書式設定                                        ' 8. 見出し色・日付書式・列幅をそろえる（一覧として見やすく）
    Call Pr_メニュー作成                                    ' 9. メニューの表示とボタンを作る（業務の入口）
    Call Pr_入力シート作成                                  ' 9b. 入力シートとMRPシートを作る（値を入れてボタンを押す画面）
    Call Pr_品目計算列更新                                  ' 10. 標準原価・販売単価を計算（品目の計算列を埋める）
    Call Pr_画面最新化                                      ' 10b. 件数・一覧・選択肢を表示（状況をひと目で見せる）
    Call Pr_完了表示                                        ' 11. メニューを開いて完了を伝える（次の操作を案内）

終了処理:
    Call Pr_終了処理(Err.Number, Err.Description)           ' 画面を戻しエラーなら表示（どの場合も必ず通る）
End Sub

' 複数の部品で使う値を格納する
Private Sub Pr_変数格納()
    Dt_今日 = Date
    Bl_既存データあり = Fn_シートあり(St_シート_品目)
End Sub

' 既存データがあれば、消してよいか確認する。「いいえ」なら中断
Private Sub Pr_実行確認()
    If Not Bl_既存データあり Then Exit Sub
    If MsgBox("両事業部の在庫・伝票・マスタの変更をすべて消して、初期データに戻します。" & vbCrLf & _
              "（ブック内のシートをすべて作り直します）" & vbCrLf & vbCrLf & "よろしいですか？", _
              vbYesNo + vbExclamation + vbDefaultButton2, "初期データに戻す") = vbNo Then
        Err.Raise Lg_中断番号
    End If
End Sub

' 全シートを消し、決まった順にシートを作る
'   Excel は最低1枚のシートが要るので、作業用シートを1枚置いてから消す。
Private Sub Pr_シート作り直し()
    Dim Ws_表 As Worksheet
    Dim Ws_作業 As Worksheet
    Dim Ws_前 As Worksheet
    Dim Va_名前 As Variant
    Dim Lg_添字 As Long
    Application.DisplayAlerts = False
    ' 作業用シートを先頭に置き、それ以外をすべて消す
    Set Ws_作業 = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    Ws_作業.Name = St_作業用シート & Format$(Now, "hhnnss")
    For Lg_添字 = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set Ws_表 = ThisWorkbook.Worksheets(Lg_添字)
        If Ws_表.Name <> Ws_作業.Name Then Ws_表.Delete
    Next Lg_添字
    ' 決まった順に作る（操作用 → マスタ → 伝票 → 履歴・採番・設定）
    Va_名前 = Fn_シート順()
    Set Ws_前 = Ws_作業
    For Lg_添字 = LBound(Va_名前) To UBound(Va_名前)
        Set Ws_表 = ThisWorkbook.Worksheets.Add(After:=Ws_前)
        Ws_表.Name = Va_名前(Lg_添字)
        Set Ws_前 = Ws_表
    Next Lg_添字
    Ws_作業.Delete
    Application.DisplayAlerts = True
End Sub

' ブック内のシートの並び（段階2以降で入力シートをここに足す）
Private Function Fn_シート順() As Variant
    Fn_シート順 = Array(St_シート_メニュー, St_シート_入力見込, St_シート_入力受注, St_シート_MRP, _
        St_シート_入力製造, St_シート_入力発注, St_シート_負荷一覧, St_シート_在庫一覧, _
        St_シート_品目, St_シート_BOM, St_シート_作業区, St_シート_工程表, _
        St_シート_見込, St_シート_受注, St_シート_製造指図, St_シート_指図工程, _
        St_シート_発注, St_シート_外注注文, St_シート_ピッキング, St_シート_ピッキング明細, _
        St_シート_出荷実績, St_シート_履歴, St_シート_採番, St_シート_設定, St_シート_選択肢)
End Function

' 各データシートの見出しを書く（並びは Mod_共通定義 の列挙体と同じ）
Private Sub Pr_見出し作成()
    Call Pr_見出し書込(St_シート_品目, 列_品目.次列 - 1, Array("事業部", "品番", "品名", "区分", "単位", "小数可", "生産方式", "倉庫", "棚番", "安全在庫", "LT(日)", "購入単価", "販売単価", "在庫数", "標準原価", "適用販売単価"))
    Call Pr_見出し書込(St_シート_BOM, 列_BOM.次列 - 1, Array("事業部", "親品番", "子品番", "員数"))
    Call Pr_見出し書込(St_シート_作業区, 列_作業区.次列 - 1, Array("事業部", "作業区", "作業区名", "賃率(円/時)", "能力(分/日)"))
    Call Pr_見出し書込(St_シート_工程表, 列_工程表.次列 - 1, Array("事業部", "品番", "工程順", "工程名", "作業区", "段取時間(分)", "加工時間(分/単位)", "外注", "外注先", "外注単価", "外注LT(日)"))
    Call Pr_見出し書込(St_シート_見込, 列_見込.次列 - 1, Array("事業部", "見込ID", "品番", "数量", "期日"))
    Call Pr_見出し書込(St_シート_受注, 列_受注.次列 - 1, Array("事業部", "受注No", "品番", "数量", "出荷済", "得意先", "納期", "状態"))
    Call Pr_見出し書込(St_シート_製造指図, 列_製造指図.次列 - 1, Array("事業部", "指図No", "品番", "数量", "状態", "着手予定", "完成予定", "受注No", "ピッキングNo", "出来高", "備考"))
    Call Pr_見出し書込(St_シート_指図工程, 列_指図工程.次列 - 1, Array("事業部", "指図No", "工程順", "手配No", "工程名", "作業区", "段取時間", "加工時間", "外注", "外注先", "外注単価", "外注LT", "状態", "良品", "不良", "作業時間(分)"))
    Call Pr_見出し書込(St_シート_発注, 列_発注.次列 - 1, Array("事業部", "発注No", "品番", "数量", "入荷予定", "状態"))
    Call Pr_見出し書込(St_シート_外注注文, 列_外注注文.次列 - 1, Array("事業部", "外注No", "指図No", "工程順", "工程名", "品番", "外注先", "数量", "単価", "納期", "状態", "良品", "不良"))
    Call Pr_見出し書込(St_シート_ピッキング, 列_ピッキング.次列 - 1, Array("事業部", "ピッキングNo", "種別", "対象伝票", "内容", "状態", "配送方法", "出荷No"))
    Call Pr_見出し書込(St_シート_ピッキング明細, 列_ピッキング明細.次列 - 1, Array("事業部", "ピッキングNo", "行No", "品番", "数量", "棚番", "済"))
    Call Pr_見出し書込(St_シート_出荷実績, 列_出荷実績.次列 - 1, Array("事業部", "出荷No", "受注No", "品番", "数量", "得意先", "出荷日", "配送方法", "販売単価"))
    Call Pr_見出し書込(St_シート_履歴, 列_履歴.次列 - 1, Array("事業部", "日時", "種別", "品番", "増減", "残", "伝票"))
    Call Pr_見出し書込(St_シート_採番, 列_採番.次列 - 1, Array("事業部", "種類", "最終番号"))
    Call Pr_見出し書込(St_シート_設定, 列_設定.次列 - 1, Array("項目", "値", "説明"))
End Sub

' 1枚のシートに見出しを書く。見出しの数が列挙体の列数と違えば中断
'   引数：St_シート名、Lg_列数（列挙体の 次列 - 1）、Va_見出し（0始まりの配列）
'   列を足したのに見出しを直し忘れる、の食い違いを初期化の時点で見つけるため。
Private Sub Pr_見出し書込(ByVal St_シート名 As String, ByVal Lg_列数 As Long, ByVal Va_見出し As Variant)
    Dim Ws_表 As Worksheet
    If UBound(Va_見出し) - LBound(Va_見出し) + 1 <> Lg_列数 Then
        MsgBox "シート「" & St_シート名 & "」の見出しの数（" & (UBound(Va_見出し) - LBound(Va_見出し) + 1) & _
               "）が列挙体の列数（" & Lg_列数 & "）と合いません。Mod_初期化 の見出しを直してください。", vbCritical
        Err.Raise Lg_中断番号
    End If
    Set Ws_表 = ThisWorkbook.Worksheets(St_シート名)
    Ws_表.Range(Ws_表.Cells(Lg_見出し行, 1), Ws_表.Cells(Lg_見出し行, Lg_列数)).Value = Va_見出し
End Sub

'==============================================================
' マスタ（HTML 版 seed0 の items・bom・wc・rt と同じ）
'==============================================================
Private Sub Pr_マスタ投入()
    ' ---- 品目（事業部, 品番, 品名, 区分, 単位, 在庫, 安全在庫, LT, 購入単価, 生産方式）----
    Call Pr_品目投入(St_加工, "M-001", "鋼板 SS400 t3.2", "材料", "kg", 400, 300, 5, 150, "")
    Call Pr_品目投入(St_加工, "M-002", "丸棒 S45C φ20", "材料", "kg", 240, 120, 7, 180, "")
    Call Pr_品目投入(St_加工, "P-101", "ブラケット", "加工品", "個", 35, 30, 2, "", St_見込生産)
    Call Pr_品目投入(St_加工, "P-102", "シャフト", "加工品", "個", 18, 30, 2, "", St_見込生産)
    Call Pr_品目投入(St_加工, "P-103", "ベースプレート", "加工品", "個", 8, 0, 3, "", St_受注生産)
    Call Pr_品目投入(St_組立, "A-301", "アームシステム", "完成品", "台", 0, 0, 3, "", St_受注生産)
    Call Pr_品目投入(St_組立, "A-201", "可動アーム", "完成品", "台", 6, 10, 2, "", St_見込生産)
    Call Pr_品目投入(St_組立, "A-202", "制御ユニット", "完成品", "台", 4, 5, 2, "", St_見込生産)
    Call Pr_品目投入(St_組立, "S-101", "駆動ユニット", "中間品", "個", 8, 10, 1, "", St_見込生産)
    Call Pr_品目投入(St_組立, "B-001", "モーター", "部品", "個", 50, 30, 14, 4200, "")
    Call Pr_品目投入(St_組立, "B-002", "制御基板", "部品", "枚", 40, 20, 21, 6800, "")
    Call Pr_品目投入(St_組立, "B-003", "ネジ M4", "部品", "本", 460, 200, 3, 3, "")
    Call Pr_品目投入(St_組立, "B-004", "アームフレーム", "部品", "個", 45, 20, 10, 2500, "")
    Call Pr_品目投入(St_組立, "B-005", "ケース", "部品", "個", 30, 15, 7, 1200, "")
    Call Pr_品目投入(St_組立, "B-006", "ハーネス", "部品", "本", 70, 30, 7, 380, "")

    ' ---- BOM（事業部, 親品番, 子品番, 員数）----
    Call Pr_BOM投入(St_加工, "P-101", "M-001", 1.2)
    Call Pr_BOM投入(St_加工, "P-102", "M-002", 0.8)
    Call Pr_BOM投入(St_加工, "P-103", "M-001", 3.5)
    Call Pr_BOM投入(St_組立, "A-301", "A-201", 1)
    Call Pr_BOM投入(St_組立, "A-301", "A-202", 1)
    Call Pr_BOM投入(St_組立, "A-301", "B-003", 4)
    Call Pr_BOM投入(St_組立, "A-201", "S-101", 1)
    Call Pr_BOM投入(St_組立, "A-201", "B-004", 2)
    Call Pr_BOM投入(St_組立, "A-201", "B-003", 8)
    Call Pr_BOM投入(St_組立, "A-202", "B-002", 1)
    Call Pr_BOM投入(St_組立, "A-202", "B-005", 1)
    Call Pr_BOM投入(St_組立, "A-202", "B-006", 2)
    Call Pr_BOM投入(St_組立, "A-202", "B-003", 6)
    Call Pr_BOM投入(St_組立, "S-101", "B-001", 1)
    Call Pr_BOM投入(St_組立, "S-101", "B-006", 1)
    Call Pr_BOM投入(St_組立, "S-101", "B-003", 4)

    ' ---- 作業区（事業部, 作業区, 作業区名, 賃率, 能力）。能力は全作業区 450分/日 ----
    Call Pr_作業区投入(St_加工, "WC-L", "レーザー切断", 6000)
    Call Pr_作業区投入(St_加工, "WC-B", "ベンダー曲げ", 4800)
    Call Pr_作業区投入(St_加工, "WC-T", "NC旋盤", 5400)
    Call Pr_作業区投入(St_加工, "WC-M", "マシニング", 6600)
    Call Pr_作業区投入(St_加工, "WC-Q", "検査", 3600)
    Call Pr_作業区投入(St_組立, "WC-S", "サブ組立", 3600)
    Call Pr_作業区投入(St_組立, "WC-A1", "組立ライン1", 4200)
    Call Pr_作業区投入(St_組立, "WC-A2", "組立ライン2", 4200)
    Call Pr_作業区投入(St_組立, "WC-E", "電装・配線", 4500)
    Call Pr_作業区投入(St_組立, "WC-I", "検査・調整", 4800)

    ' ---- 工程表（社内：事業部, 品番, 工程順, 工程名, 作業区, 段取, 加工）----
    ' ---- 　　　（外注：事業部, 品番, 工程順, 工程名, 外注先, 外注単価, 外注LT）----
    Call Pr_社内工程投入(St_加工, "P-101", 10, "レーザー切断", "WC-L", 20, 0.5)
    Call Pr_社内工程投入(St_加工, "P-101", 20, "曲げ", "WC-B", 15, 0.8)
    Call Pr_外注工程投入(St_加工, "P-101", 30, "メッキ処理", "光陽メッキ", 45, 3)
    Call Pr_社内工程投入(St_加工, "P-101", 40, "外観・寸法検査", "WC-Q", 5, 0.2)
    Call Pr_社内工程投入(St_加工, "P-102", 10, "旋削", "WC-T", 30, 3)
    Call Pr_外注工程投入(St_加工, "P-102", 20, "熱処理", "東都熱処理", 120, 2)
    Call Pr_社内工程投入(St_加工, "P-102", 30, "寸法検査", "WC-Q", 5, 0.5)
    Call Pr_社内工程投入(St_加工, "P-103", 10, "レーザー切断", "WC-L", 20, 1)
    Call Pr_社内工程投入(St_加工, "P-103", 20, "穴あけ・面取り", "WC-M", 25, 4)
    Call Pr_社内工程投入(St_加工, "P-103", 30, "寸法検査", "WC-Q", 5, 0.5)
    Call Pr_社内工程投入(St_組立, "S-101", 10, "サブ組立", "WC-S", 10, 6)
    Call Pr_社内工程投入(St_組立, "S-101", 20, "通電検査", "WC-I", 5, 2)
    Call Pr_社内工程投入(St_組立, "A-201", 10, "フレーム組立", "WC-A1", 15, 12)
    Call Pr_外注工程投入(St_組立, "A-201", 20, "フレーム塗装", "関東塗装工業", 600, 2)
    Call Pr_社内工程投入(St_組立, "A-201", 30, "駆動部取付", "WC-A1", 5, 8)
    Call Pr_社内工程投入(St_組立, "A-201", 40, "動作検査", "WC-I", 10, 5)
    Call Pr_社内工程投入(St_組立, "A-202", 10, "基板取付・配線", "WC-E", 10, 10)
    Call Pr_社内工程投入(St_組立, "A-202", 20, "ケース組付", "WC-A2", 5, 4)
    Call Pr_社内工程投入(St_組立, "A-202", 30, "通電検査", "WC-I", 5, 3)
    Call Pr_社内工程投入(St_組立, "A-301", 10, "総組立", "WC-A2", 20, 15)
    Call Pr_社内工程投入(St_組立, "A-301", 20, "配線・設定", "WC-E", 10, 10)
    Call Pr_社内工程投入(St_組立, "A-301", 30, "総合試験", "WC-I", 15, 12)
End Sub

' 品目を1行入れる。小数可・倉庫・棚番は HTML 版の決め方で埋める
'   小数可：単位が kg なら○（HTML 版は kg の品目だけ小数を入力できる）
'   倉庫  ：区分から（材料→材料倉庫、部品→部品倉庫、中間品→仕掛品倉庫、加工品・完成品→製品倉庫）
'   棚番  ：HTML 版の LOCS と同じ
Private Sub Pr_品目投入(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal St_品名 As String, _
        ByVal St_区分 As String, ByVal St_単位 As String, ByVal Db_在庫 As Double, ByVal Db_安全在庫 As Double, _
        ByVal Lg_LT As Long, ByVal Va_購入単価 As Variant, ByVal St_生産方式 As String)
    Dim St_小数可 As String
    If St_単位 = "kg" Then St_小数可 = St_印 Else St_小数可 = ""
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_品目), Array(St_事業部, St_品番, St_品名, St_区分, St_単位, _
        St_小数可, St_生産方式, Fn_初期倉庫(St_区分), Fn_初期棚番(St_品番), Db_安全在庫, Lg_LT, Va_購入単価, "", Db_在庫, "", ""))
End Sub

' 区分から倉庫を決める（HTML 版の WHDEF）
Private Function Fn_初期倉庫(ByVal St_区分 As String) As String
    Select Case St_区分
        Case "材料": Fn_初期倉庫 = "材料倉庫"
        Case "部品": Fn_初期倉庫 = "部品倉庫"
        Case "中間品": Fn_初期倉庫 = "仕掛品倉庫"
        Case Else: Fn_初期倉庫 = "製品倉庫"
    End Select
End Function

' 品番から棚番を決める（HTML 版の LOCS）
Private Function Fn_初期棚番(ByVal St_品番 As String) As String
    Select Case St_品番
        Case "M-001": Fn_初期棚番 = "A-01-01"
        Case "M-002": Fn_初期棚番 = "A-01-02"
        Case "P-101": Fn_初期棚番 = "C-02-01"
        Case "P-102": Fn_初期棚番 = "C-02-02"
        Case "P-103": Fn_初期棚番 = "C-03-01"
        Case "A-301": Fn_初期棚番 = "F-01-01"
        Case "A-201": Fn_初期棚番 = "F-01-02"
        Case "A-202": Fn_初期棚番 = "F-01-03"
        Case "S-101": Fn_初期棚番 = "D-02-01"
        Case "B-001": Fn_初期棚番 = "B-01-01"
        Case "B-002": Fn_初期棚番 = "B-01-02"
        Case "B-003": Fn_初期棚番 = "B-03-05"
        Case "B-004": Fn_初期棚番 = "B-02-01"
        Case "B-005": Fn_初期棚番 = "B-02-02"
        Case "B-006": Fn_初期棚番 = "B-03-01"
        Case Else: Fn_初期棚番 = ""
    End Select
End Function

' BOM を1行入れる
Private Sub Pr_BOM投入(ByVal St_事業部 As String, ByVal St_親 As String, ByVal St_子 As String, ByVal Db_員数 As Double)
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_BOM), Array(St_事業部, St_親, St_子, Db_員数))
End Sub

' 作業区を1行入れる（能力は HTML 版の CAP＝450分/日）
Private Sub Pr_作業区投入(ByVal St_事業部 As String, ByVal St_作業区 As String, ByVal St_名前 As String, ByVal Db_賃率 As Double)
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_作業区), Array(St_事業部, St_作業区, St_名前, Db_賃率, 450))
End Sub

' 社内工程を1行入れる
Private Sub Pr_社内工程投入(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal Lg_工程順 As Long, _
        ByVal St_工程名 As String, ByVal St_作業区 As String, ByVal Db_段取 As Double, ByVal Db_加工 As Double)
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_工程表), Array(St_事業部, St_品番, Lg_工程順, St_工程名, _
        St_作業区, Db_段取, Db_加工, "", "", "", ""))
End Sub

' 外注工程を1行入れる（社内の時間は 0、作業区は空）
Private Sub Pr_外注工程投入(ByVal St_事業部 As String, ByVal St_品番 As String, ByVal Lg_工程順 As Long, _
        ByVal St_工程名 As String, ByVal St_外注先 As String, ByVal Db_単価 As Double, ByVal Lg_LT As Long)
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_工程表), Array(St_事業部, St_品番, Lg_工程順, St_工程名, _
        "", 0, 0, St_印, St_外注先, Db_単価, Lg_LT))
End Sub

'==============================================================
' 採番と設定
'==============================================================
Private Sub Pr_採番設定投入()
    Dim Ws_採番 As Worksheet
    Dim Ws_設定 As Worksheet
    Set Ws_採番 = ThisWorkbook.Worksheets(St_シート_採番)
    Set Ws_設定 = ThisWorkbook.Worksheets(St_シート_設定)
    ' 採番：HTML 版 seed の seq と同じ最終番号。手配No.は伝票投入で 1000001 から振る
    Call Pr_採番行(Ws_採番, St_加工, Array("SO", 4, "MO", 2, "PO", 0, "SH", 0, "FC", 2, "SC", 1, "PK", 1))
    Call Pr_採番行(Ws_採番, St_組立, Array("SO", 3, "MO", 2, "PO", 1, "SH", 0, "FC", 2, "SC", 0, "PK", 1))
    Call Fn_行追記(Ws_採番, Array("全社", "手配No", 1000000))
    ' 設定（項目, 値, 説明）
    Call Fn_行追記(Ws_設定, Array("現在の事業部", St_加工, "加工／組立。メニューの切替ボタンで変わる"))
    Call Fn_行追記(Ws_設定, Array("MRP安全在庫", St_印, "○なら MRP で安全在庫も確保する（HTML 版の「安全在庫も確保する」）"))
    Call Fn_行追記(Ws_設定, Array("負荷表示日数", 14, "負荷一覧に出す日数（今日から）"))
    Call Fn_行追記(Ws_設定, Array("バーコードフォント", "Free 3 of 9", "CODE39 のフォント名。PC に無ければ番号だけ印字する"))
    Call Fn_行追記(Ws_設定, Array("PDF保存先", "", "空欄ならブックと同じフォルダの「帳票」フォルダ"))
    Call Fn_行追記(Ws_設定, Array("会社名", "株式会社 架空製作所", "帳票の発行者（架空）"))
    Call Fn_行追記(Ws_設定, Array("住所", "〒000-0000 架空県見本市サンプル町1-2-3", "帳票の発行者（架空）"))
    Call Fn_行追記(Ws_設定, Array("電話", "TEL 000-000-0000", "帳票の発行者（架空）"))
    Call Fn_行追記(Ws_設定, Array("登録番号", "登録番号 T0000000000000（見本）", "請求書に印字（架空）"))
    Call Fn_行追記(Ws_設定, Array("振込先", "見本銀行 サンプル支店　普通 0000000　カ）カクウセイサクショ", "請求書に印字（架空）"))
    Call Fn_行追記(Ws_設定, Array("初期化日", Dt_今日, "初期データの日付の基準日"))
End Sub

' 採番シートに1事業部ぶんの行を入れる
'   引数：Va_種類と番号（"SO", 4, "MO", 2, … の交互の配列）
Private Sub Pr_採番行(ByVal Ws_採番 As Worksheet, ByVal St_事業部 As String, ByVal Va_種類と番号 As Variant)
    Dim Lg_添字 As Long
    For Lg_添字 = LBound(Va_種類と番号) To UBound(Va_種類と番号) Step 2
        Call Fn_行追記(Ws_採番, Array(St_事業部, Va_種類と番号(Lg_添字), Va_種類と番号(Lg_添字 + 1)))
    Next Lg_添字
End Sub

'==============================================================
' 伝票（HTML 版 seed の so・mo・sc・po・fc・pk と同じ）
'   日付は Dt_今日 からの日数。行は作られた順（古いものが上）に入れる。
'==============================================================
Private Sub Pr_伝票投入()
    Dim Ws_見込 As Worksheet
    Dim Ws_受注 As Worksheet
    Dim Ws_指図 As Worksheet
    Set Ws_見込 = ThisWorkbook.Worksheets(St_シート_見込)
    Set Ws_受注 = ThisWorkbook.Worksheets(St_シート_受注)
    Set Ws_指図 = ThisWorkbook.Worksheets(St_シート_製造指図)

    ' ---- 見込（事業部, 見込ID, 品番, 数量, 期日）----
    Call Fn_行追記(Ws_見込, Array(St_加工, 1, "P-101", 150, Dt_今日 + 9))
    Call Fn_行追記(Ws_見込, Array(St_加工, 2, "P-102", 80, Dt_今日 + 8))
    Call Fn_行追記(Ws_見込, Array(St_組立, 1, "A-201", 20, Dt_今日 + 12))
    Call Fn_行追記(Ws_見込, Array(St_組立, 2, "A-202", 10, Dt_今日 + 10))

    ' ---- 受注（事業部, 受注No, 品番, 数量, 出荷済, 得意先, 納期, 状態）----
    Call Fn_行追記(Ws_受注, Array(St_加工, "SO-0001", "P-101", 200, 0, "東和機械", Dt_今日 + 10, St_状態_受付))
    Call Fn_行追記(Ws_受注, Array(St_加工, "SO-0002", "P-102", 60, 0, "北斗精工", Dt_今日 + 6, St_状態_受付))
    Call Fn_行追記(Ws_受注, Array(St_加工, "SO-0003", "P-103", 8, 0, "みなと工業", Dt_今日 + 3, St_状態_引当済))
    Call Fn_行追記(Ws_受注, Array(St_加工, "SO-0004", "P-103", 6, 0, "東和機械", Dt_今日 + 7, St_状態_受付))
    Call Fn_行追記(Ws_受注, Array(St_組立, "SO-0001", "A-201", 12, 0, "東和機械", Dt_今日 + 14, St_状態_受付))
    Call Fn_行追記(Ws_受注, Array(St_組立, "SO-0002", "A-301", 3, 0, "みなと工業", Dt_今日 + 21, St_状態_受付))
    Call Fn_行追記(Ws_受注, Array(St_組立, "SO-0003", "A-202", 4, 0, "北斗精工", Dt_今日 + 5, St_状態_引当済))

    ' ---- 製造指図（事業部, 指図No, 品番, 数量, 状態, 着手予定, 完成予定, 受注No, ピッキングNo, 出来高, 備考）----
    '      工程は工程表から複写し、手配No.を振る（加工 MO-0001 → MO-0002 → 組立 MO-0001 → MO-0002 の順）
    Call Fn_行追記(Ws_指図, Array(St_加工, "MO-0001", "P-102", 20, St_状態_計画, Dt_今日 + 1, Dt_今日 + 4, "", "", "", ""))
    Call Pr_指図工程複写(St_加工, "MO-0001", "P-102")
    Call Fn_行追記(Ws_指図, Array(St_加工, "MO-0002", "P-101", 40, St_状態_着手, Dt_今日 - 2, Dt_今日 + 3, "", "", "", ""))
    Call Pr_指図工程複写(St_加工, "MO-0002", "P-101")
    Call Fn_行追記(Ws_指図, Array(St_組立, "MO-0001", "S-101", 10, St_状態_着手, Dt_今日 - 1, Dt_今日 + 1, "", "", "", ""))
    Call Pr_指図工程複写(St_組立, "MO-0001", "S-101")
    Call Fn_行追記(Ws_指図, Array(St_組立, "MO-0002", "A-202", 6, St_状態_計画, Dt_今日 + 1, Dt_今日 + 2, "", "PK-0001", "", ""))
    Call Pr_指図工程複写(St_組立, "MO-0002", "A-202")

    ' ---- 途中まで進んだ工程（事業部, 指図No, 工程順, 状態, 良品, 不良, 作業時間）----
    Call Pr_工程状態設定(St_加工, "MO-0002", 10, St_状態_完了, 40, 0, 42)
    Call Pr_工程状態設定(St_加工, "MO-0002", 20, St_状態_完了, 40, 0, 50)
    Call Pr_工程状態設定(St_加工, "MO-0002", 30, St_状態_外注中, "", "", "")
    Call Pr_工程状態設定(St_組立, "MO-0001", 10, St_状態_完了, 10, 0, 74)

    ' ---- 外注注文（事業部, 外注No, 指図No, 工程順, 工程名, 品番, 外注先, 数量, 単価, 納期, 状態, 良品, 不良）----
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_外注注文), Array(St_加工, "SC-0001", "MO-0002", 30, "メッキ処理", "P-101", _
        "光陽メッキ", 40, 45, Dt_今日 + 2, St_状態_外注中, "", ""))

    ' ---- 発注（事業部, 発注No, 品番, 数量, 入荷予定, 状態）----
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_発注), Array(St_組立, "PO-0001", "B-002", 20, Dt_今日 + 9, St_状態_発注済))

    ' ---- ピッキング（事業部, ピッキングNo, 種別, 対象伝票, 内容, 状態, 配送方法, 出荷No）----
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング), Array(St_加工, "PK-0001", "出荷", "SO-0003", _
        "ベースプレート 8個 → みなと工業", St_状態_作業中, "路線便", ""))
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング), Array(St_組立, "PK-0001", "製造", "MO-0002", _
        "制御ユニット ×6 の部材", St_状態_作業中, "", ""))
    ' ---- ピッキング明細（事業部, ピッキングNo, 行No, 品番, 数量, 棚番, 済）----
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング明細), Array(St_加工, "PK-0001", 1, "P-103", 8, "C-03-01", ""))
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング明細), Array(St_組立, "PK-0001", 1, "B-002", 6, "B-01-02", St_印))
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング明細), Array(St_組立, "PK-0001", 2, "B-005", 6, "B-02-02", ""))
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング明細), Array(St_組立, "PK-0001", 3, "B-006", 12, "B-03-01", ""))
    Call Fn_行追記(ThisWorkbook.Worksheets(St_シート_ピッキング明細), Array(St_組立, "PK-0001", 4, "B-003", 36, "B-03-05", ""))
End Sub

' 指図工程の1行の状態と実績を書き換える（初期データで途中まで進んだ工程を作るため）
Private Sub Pr_工程状態設定(ByVal St_事業部 As String, ByVal St_指図No As String, ByVal Lg_工程順 As Long, _
        ByVal St_状態 As String, ByVal Va_良品 As Variant, ByVal Va_不良 As Variant, ByVal Va_時間 As Variant)
    Dim Ws_工程 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim Lg_シート行 As Long
    Set Ws_工程 = ThisWorkbook.Worksheets(St_シート_指図工程)
    Va_表 = Fn_表読込(Ws_工程, 列_指図工程.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_指図工程.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_指図工程.指図No)) = St_指図No _
           And Fn_数(Va_表(Lg_行, 列_指図工程.工程順)) = Lg_工程順 Then
            Lg_シート行 = Lg_行 + Lg_開始行 - 1
            Ws_工程.Cells(Lg_シート行, 列_指図工程.状態).Value = St_状態
            Ws_工程.Cells(Lg_シート行, 列_指図工程.良品).Value = Va_良品
            Ws_工程.Cells(Lg_シート行, 列_指図工程.不良).Value = Va_不良
            Ws_工程.Cells(Lg_シート行, 列_指図工程.作業時間).Value = Va_時間
            Exit Sub
        End If
    Next Lg_行
End Sub

'==============================================================
' 書式
'==============================================================
Private Sub Pr_書式設定()
    Dim Va_名前 As Variant
    Dim Va_名 As Variant
    ' データシート共通：見出しの色・太字、1行目の固定、オートフィルタ
    Va_名前 = Array(St_シート_品目, St_シート_BOM, St_シート_作業区, St_シート_工程表, St_シート_見込, St_シート_受注, _
        St_シート_製造指図, St_シート_指図工程, St_シート_発注, St_シート_外注注文, St_シート_ピッキング, _
        St_シート_ピッキング明細, St_シート_出荷実績, St_シート_履歴, St_シート_採番, St_シート_設定)
    For Each Va_名 In Va_名前
        Call Pr_表の書式(ThisWorkbook.Worksheets(CStr(Va_名)))
    Next Va_名
    ' 日付の列
    Call Pr_列書式(St_シート_見込, 列_見込.期日, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_受注, 列_受注.納期, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_製造指図, 列_製造指図.着手予定, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_製造指図, 列_製造指図.完成予定, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_発注, 列_発注.入荷予定, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_外注注文, 列_外注注文.納期, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_出荷実績, 列_出荷実績.出荷日, "yyyy/mm/dd")
    Call Pr_列書式(St_シート_履歴, 列_履歴.日時, "yyyy/mm/dd hh:mm")
    Call Pr_列書式(St_シート_設定, 列_設定.値, "General")
    ThisWorkbook.Worksheets(St_シート_設定).Cells(Fn_最終行(ThisWorkbook.Worksheets(St_シート_設定)), 列_設定.値).NumberFormat = "yyyy/mm/dd"
    ' マスタ：マクロが管理する列（在庫数）と計算列に色を付ける（直接書き換えない列の目印）
    Call Pr_列の色(St_シート_品目, 列_品目.在庫数, RGB(255, 242, 204))
    Call Pr_列の色(St_シート_品目, 列_品目.標準原価, RGB(237, 237, 237))
    Call Pr_列の色(St_シート_品目, 列_品目.適用販売単価, RGB(237, 237, 237))
    ' シート見出しの色（操作用＝青、マスタ＝緑、伝票＝橙、その他＝灰）
    Call Pr_タブ色(Array(St_シート_メニュー, St_シート_在庫一覧), RGB(91, 155, 213))
    Call Pr_タブ色(Array(St_シート_品目, St_シート_BOM, St_シート_作業区, St_シート_工程表), RGB(112, 173, 71))
    Call Pr_タブ色(Array(St_シート_見込, St_シート_受注, St_シート_製造指図, St_シート_指図工程, St_シート_発注, _
        St_シート_外注注文, St_シート_ピッキング, St_シート_ピッキング明細, St_シート_出荷実績), RGB(237, 125, 49))
    Call Pr_タブ色(Array(St_シート_履歴, St_シート_採番, St_シート_設定), RGB(165, 165, 165))
    ' 列幅
    For Each Va_名 In Va_名前
        ThisWorkbook.Worksheets(CStr(Va_名)).Columns("A:P").AutoFit
    Next Va_名
End Sub

' 表の見出しを整え、1行目を固定し、オートフィルタを付ける
Private Sub Pr_表の書式(ByVal Ws_表 As Worksheet)
    Dim Lg_列数 As Long
    Lg_列数 = Ws_表.Cells(Lg_見出し行, Ws_表.Columns.Count).End(xlToLeft).Column
    With Ws_表.Range(Ws_表.Cells(Lg_見出し行, 1), Ws_表.Cells(Lg_見出し行, Lg_列数))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    Ws_表.Range(Ws_表.Cells(Lg_見出し行, 1), Ws_表.Cells(Lg_見出し行, Lg_列数)).AutoFilter
    Ws_表.Activate
    ActiveWindow.FreezePanes = False
    Ws_表.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' 1列の表示形式をそろえる（見出しを除く）
Private Sub Pr_列書式(ByVal St_シート名 As String, ByVal Lg_列 As Long, ByVal St_書式 As String)
    With ThisWorkbook.Worksheets(St_シート名)
        .Range(.Cells(Lg_開始行, Lg_列), .Cells(.Rows.Count, Lg_列)).NumberFormat = St_書式
    End With
End Sub

' 1列の見出しに色を付ける（マクロが書く列の目印）
Private Sub Pr_列の色(ByVal St_シート名 As String, ByVal Lg_列 As Long, ByVal Lg_色 As Long)
    ThisWorkbook.Worksheets(St_シート名).Cells(Lg_見出し行, Lg_列).Interior.Color = Lg_色
End Sub

' シート見出しに色を付ける
Private Sub Pr_タブ色(ByVal Va_名前 As Variant, ByVal Lg_色 As Long)
    Dim Va_名 As Variant
    For Each Va_名 In Va_名前
        ThisWorkbook.Worksheets(CStr(Va_名)).Tab.Color = Lg_色
    Next Va_名
End Sub

'==============================================================
' メニュー
'==============================================================
Private Sub Pr_メニュー作成()
    Dim Ws_メニュー As Worksheet
    Set Ws_メニュー = ThisWorkbook.Worksheets(St_シート_メニュー)
    ActiveWindow.DisplayGridlines = True
    Ws_メニュー.Activate
    ActiveWindow.DisplayGridlines = False
    Ws_メニュー.Columns("A").ColumnWidth = 2
    Ws_メニュー.Columns("B:H").ColumnWidth = 14
    ' 題名と事業部
    Ws_メニュー.Range("B1").Value = "生産管理シミュレータ（Excel 版）"
    Ws_メニュー.Range("B1").Font.Size = 16
    Ws_メニュー.Range("B1").Font.Bold = True
    Ws_メニュー.Range("B2").Value = "架空の会社・架空のデータです。操作はすべて「現在の事業部」が対象です。"
    Ws_メニュー.Range("B3").Value = "現在の事業部"
    Ws_メニュー.Range(St_メニュー_事業部セル).Value = Fn_事業部名(St_加工)
    Ws_メニュー.Range(St_メニュー_事業部セル).Font.Bold = True
    Ws_メニュー.Range(St_メニュー_事業部セル).Font.Size = 14
    Call Pr_ボタン作成(Ws_メニュー, "E3", 1, "加工事業部", "Pr_加工事業部に切替", "切替_加工")
    Call Pr_ボタン作成(Ws_メニュー, "F3", 1, "組立事業部", "Pr_組立事業部に切替", "切替_組立")
    ' 件数
    Ws_メニュー.Range("C5:F5").Value = Array("受注残", "要手配", "製造中", "出荷待ち")
    Ws_メニュー.Range("C5:F5").Font.Bold = True
    Ws_メニュー.Range("C5:F6").HorizontalAlignment = xlCenter
    Ws_メニュー.Range("C6:F6").Font.Size = 16
    Ws_メニュー.Range(St_メニュー_要手配セル).Value = "未計算"
    ' 業務の流れ（左から順に。入力シートへ移動するボタンと、更新して表示するボタン）
    Ws_メニュー.Range("B8").Value = "■ 業務の流れ"
    Ws_メニュー.Range("B8").Font.Bold = True
    Call Pr_ボタン作成(Ws_メニュー, "B9", 1, "1. 見込", "Pr_シート移動", St_移動ボタン接頭 & St_シート_入力見込)
    Call Pr_ボタン作成(Ws_メニュー, "C9", 1, "2. 受注", "Pr_シート移動", St_移動ボタン接頭 & St_シート_入力受注)
    Call Pr_ボタン作成(Ws_メニュー, "D9", 1, "3. MRP", "Pr_シート移動", St_移動ボタン接頭 & St_シート_MRP)
    Call Pr_ボタン作成(Ws_メニュー, "E9", 1, "4. 負荷一覧", "Pr_負荷一覧更新", "業務_負荷一覧")
    Call Pr_ボタン作成(Ws_メニュー, "F9", 1, "5. 製造", "Pr_シート移動", St_移動ボタン接頭 & St_シート_入力製造)
    Call Pr_ボタン作成(Ws_メニュー, "G9", 1, "6. 発注", "Pr_シート移動", St_移動ボタン接頭 & St_シート_入力発注)
    Call Pr_ボタン作成(Ws_メニュー, "H9", 1, "在庫一覧", "Pr_在庫一覧更新", "業務_在庫一覧")
    Ws_メニュー.Range("B10").Value = "（負荷一覧・在庫一覧のボタンは、計算し直してから開きます）"
    ' マスタ（シートを直接編集）
    Ws_メニュー.Range("B12").Value = "■ マスタ（シートを直接編集する）"
    Ws_メニュー.Range("B12").Font.Bold = True
    Call Pr_移動ボタン列(Ws_メニュー, 13, Array(St_シート_品目, St_シート_BOM, St_シート_作業区, St_シート_工程表))
    ' データ（一覧として見る）
    Ws_メニュー.Range("B15").Value = "■ データ（一覧として見る。直接は書き換えない）"
    Ws_メニュー.Range("B15").Font.Bold = True
    Call Pr_移動ボタン列(Ws_メニュー, 16, Array(St_シート_見込, St_シート_受注, St_シート_製造指図, St_シート_指図工程, _
        St_シート_発注, St_シート_外注注文))
    Call Pr_移動ボタン列(Ws_メニュー, 18, Array(St_シート_ピッキング, St_シート_ピッキング明細, St_シート_出荷実績, _
        St_シート_履歴, St_シート_採番, St_シート_設定))
    ' その他
    Ws_メニュー.Range("B21").Value = "■ その他"
    Ws_メニュー.Range("B21").Font.Bold = True
    Call Pr_ボタン作成(Ws_メニュー, "B22", 2, "初期データに戻す", "Pr_初期化", "その他_初期化")
    Ws_メニュー.Range("B24").Value = "品目シートの黄色の列（在庫数）はマクロが書きます。灰色の列（標準原価・適用販売単価）は計算結果です。"
End Sub

' 入力シート・MRP シートを作る（配置は各モジュールが持つ）。選択肢シートは隠す
Private Sub Pr_入力シート作成()
    Call Pr_選択肢更新                  ' 入力欄のドロップダウンが参照する名前を先に作る
    Call Pr_入力見込シート作成
    Call Pr_入力受注シート作成
    Call Pr_MRPシート作成
    Call Pr_入力製造シート作成
    Call Pr_入力発注シート作成
    Call Pr_タブ色(Array(St_シート_入力見込, St_シート_入力受注, St_シート_MRP, St_シート_入力製造, St_シート_入力発注, St_シート_負荷一覧), RGB(91, 155, 213))
    ThisWorkbook.Worksheets(St_シート_選択肢).Visible = xlSheetHidden
End Sub

' 1行に移動ボタンを並べる（ボタン名は「移動_シート名」。Mod_メニュー の Pr_シート移動 が行き先に使う）
Private Sub Pr_移動ボタン列(ByVal Ws_メニュー As Worksheet, ByVal Lg_行 As Long, ByVal Va_名前 As Variant)
    Dim Lg_添字 As Long
    For Lg_添字 = LBound(Va_名前) To UBound(Va_名前)
        Call Pr_ボタン作成(Ws_メニュー, Ws_メニュー.Cells(Lg_行, 2 + Lg_添字 - LBound(Va_名前)).Address, 1, _
            CStr(Va_名前(Lg_添字)), "Pr_シート移動", St_移動ボタン接頭 & Va_名前(Lg_添字))
    Next Lg_添字
End Sub


' メニューを開いて完了を伝える
Private Sub Pr_完了表示()
    ThisWorkbook.Worksheets(St_シート_メニュー).Activate
    ThisWorkbook.Worksheets(St_シート_メニュー).Range("A1").Select
    MsgBox "初期データを入れました（基準日：" & Format$(Dt_今日, "yyyy/mm/dd") & "）。" & vbCrLf & _
           "メニューの「在庫一覧（更新）」で在庫を確認できます。", vbInformation
End Sub
