Attribute VB_Name = "Mod_共通定義"
Option Explicit
'==============================================================
' Mod_共通定義
' 【目的】すべてのモジュールが使う定義と小さな共通部品を1か所に置く。
'         ・中断番号（メインを途中で止めるための番号）
'         ・シート名
'         ・各シートの列の並び（列挙体）
'         ・丸め、表の読み書き、設定の読み書き、終了処理などの共通部品
' 【業務背景】生産管理シミュレータ（HTML 版）と同じ業務ルールで動く
'         Excel 版。HTML 版は全データを1かたまりで持つが、Excel 版は
'         本物のシステムのテーブル構造に近づけるため、データの種類ごとに
'         1シート・1行1件で持つ。どのシートも先頭列が「事業部」。
' 【前提】・データシートは1行目が見出し、2行目からデータ。
'         ・事業部の値は「加工」「組立」の2つ（表示は「加工事業部」など）。
'         ・列を足す・入れ替えるときは、このモジュールの列挙体と、
'           Mod_初期化 の見出し定義の2か所を直す（見出しの数が列挙体と
'           合わないと初期化が止まるようにしてある）。
' 【接頭辞】St_文字列 Lg_整数 Db_小数 Bl_真偽 Dt_日付 Va_バリアント
'           Ob_オブジェクト（Dictionary を含む） Co_コレクション
'           Ws_シート Wb_ブック Rg_セル範囲 Pr_Sub Fn_Function
'==============================================================

' ---- 中断（メインを途中で止める。エラー表示はしない）----
Public Const Lg_中断番号 As Long = vbObjectError + 1000

' ---- 表の行（どのデータシートも同じ）----
Public Const Lg_見出し行 As Long = 1        ' 1行目＝見出し
Public Const Lg_開始行 As Long = 2          ' 2行目からデータ

' ---- 計算の誤差（HTML 版の 1e-9 と同じ。切り上げで 2.0000000001 を 3 にしないため）----
Public Const Db_誤差 As Double = 0.000000001

' ---- 事業部（データシートの「事業部」列に入れる値）----
Public Const St_加工 As String = "加工"
Public Const St_組立 As String = "組立"

' ---- シート名 ----
Public Const St_シート_メニュー As String = "メニュー"
Public Const St_シート_在庫一覧 As String = "在庫一覧"
Public Const St_シート_品目 As String = "品目"
Public Const St_シート_BOM As String = "BOM"
Public Const St_シート_作業区 As String = "作業区"
Public Const St_シート_工程表 As String = "工程表"
Public Const St_シート_見込 As String = "見込"
Public Const St_シート_受注 As String = "受注"
Public Const St_シート_製造指図 As String = "製造指図"
Public Const St_シート_指図工程 As String = "指図工程"
Public Const St_シート_発注 As String = "発注"
Public Const St_シート_外注注文 As String = "外注注文"
Public Const St_シート_ピッキング As String = "ピッキング"
Public Const St_シート_ピッキング明細 As String = "ピッキング明細"
Public Const St_シート_出荷実績 As String = "出荷実績"
Public Const St_シート_履歴 As String = "履歴"
Public Const St_シート_採番 As String = "採番"
Public Const St_シート_設定 As String = "設定"
Public Const St_シート_入力見込 As String = "入力_見込"
Public Const St_シート_入力受注 As String = "入力_受注"
Public Const St_シート_MRP As String = "MRP"
Public Const St_シート_入力製造 As String = "入力_製造"
Public Const St_シート_入力発注 As String = "入力_発注"
Public Const St_シート_負荷一覧 As String = "負荷一覧"
Public Const St_シート_選択肢 As String = "選択肢"        ' 入力欄のドロップダウンの候補（隠しシート）

' ---- 入力シート共通の配置 ----
Public Const St_入力_結果セル As String = "B3"              ' 入力シートの「結果」を出すセル
Public Const St_数量エラー As String = "数量は1以上の整数で入力してください（小数可の品目は小数も入力できます）"

' ---- メニューシートの配置（Mod_初期化 が作り、Mod_メニュー が書き換える）----
Public Const St_メニュー_事業部セル As String = "C3"     ' 現在の事業部の表示
Public Const St_メニュー_受注残セル As String = "C6"     ' 件数：受注残（出荷済でない受注）
Public Const St_メニュー_要手配セル As String = "D6"     ' 件数：要手配（MRP の手配案の数）
Public Const St_メニュー_製造中セル As String = "E6"     ' 件数：製造中（計画・着手の製造指図）
Public Const St_メニュー_出荷待ちセル As String = "F6"   ' 件数：出荷待ち（引当済の受注）
Public Const St_移動ボタン接頭 As String = "移動_"      ' 移動ボタンの名前の頭

' ---- 状態の値（HTML 版の st と同じ言葉）----
Public Const St_状態_受付 As String = "受付"
Public Const St_状態_引当済 As String = "引当済"
Public Const St_状態_出荷済 As String = "出荷済"
Public Const St_状態_計画 As String = "計画"
Public Const St_状態_着手 As String = "着手"
Public Const St_状態_完成 As String = "完成"
Public Const St_状態_未 As String = "未"
Public Const St_状態_外注中 As String = "外注中"
Public Const St_状態_完了 As String = "完了"
Public Const St_状態_発注済 As String = "発注済"
Public Const St_状態_入荷済 As String = "入荷済"
Public Const St_状態_受入済 As String = "受入済"
Public Const St_状態_作業中 As String = "作業中"

' ---- 生産方式・印 ----
Public Const St_見込生産 As String = "見込生産"     ' MTS：見込で作って在庫を持つ
Public Const St_受注生産 As String = "受注生産"     ' MTO：受注ごとに作る（製造指図が受注に紐付く）
Public Const St_印 As String = "○"                 ' 小数可・外注・済 などの「はい」

'==============================================================
' 列の並び（列挙体）。次列 は「最後の列＋1」。列数は 次列 - 1。
'==============================================================

' 品目：品目マスタ（シートを直接編集する）
Public Enum 列_品目
    事業部 = 1          ' A：加工／組立
    品番                ' B：事業部の中で一意
    品名                ' C
    区分                ' D：材料・部品＝購入品、中間品・加工品・完成品＝製造品
    単位                ' E
    小数可              ' F：○なら小数の数量を入力できる（kg など）
    生産方式            ' G：製造品だけ。見込生産／受注生産
    倉庫                ' H：材料倉庫・部品倉庫・仕掛品倉庫・製品倉庫
    棚番                ' I：ピッキングリストはこの順に並ぶ
    安全在庫            ' J：見込生産と購入品で使う。受注生産では使わない
    LT                  ' K：リードタイム（日）。購入品の入荷予定に使う
    購入単価            ' L：購入品だけ（円）
    販売単価            ' M：製造品だけ。空欄なら標準原価の35%増し
    在庫数              ' N：マクロが管理（直接書き換えない。在庫調整を使う）
    標準原価            ' O：〔計算〕マクロが毎回書く
    適用販売単価        ' P：〔計算〕販売単価、空欄なら標準原価×1.35を10円単位に切り上げ
    次列
End Enum

' BOM：部品表（親1つあたり子をいくつ使うか）
Public Enum 列_BOM
    事業部 = 1          ' A
    親品番              ' B
    子品番              ' C
    員数                ' D：親1つあたりの子の数量
    次列
End Enum

' 作業区：社内で作業する場所と、その賃率・能力
Public Enum 列_作業区
    事業部 = 1          ' A
    作業区              ' B：作業区コード（WC-L など）
    作業区名            ' C
    賃率                ' D：円／時
    能力                ' E：分／日（初期値 450＝7.5時間）
    次列
End Enum

' 工程表：品目ごとの標準の工程（製造指図を作るときに複写する）
Public Enum 列_工程表
    事業部 = 1          ' A
    品番                ' B
    工程順              ' C：10, 20, 30 …
    工程名              ' D
    作業区              ' E：社内工程のとき
    段取時間            ' F：分／ロット（社内）
    加工時間            ' G：分／1単位（社内）
    外注                ' H：○なら外注工程
    外注先              ' I
    外注単価            ' J：円／1単位
    外注LT              ' K：日
    次列
End Enum

' 見込：見込生産の品目の「これだけ売れそう」
Public Enum 列_見込
    事業部 = 1          ' A
    見込ID              ' B：事業部ごとの連番（1, 2, …）
    品番                ' C
    数量                ' D：出荷すると納期の早い見込から減る
    期日                ' E：いつまでに
    次列
End Enum

' 受注
Public Enum 列_受注
    事業部 = 1          ' A
    受注No              ' B：SO-0001（事業部ごとの連番。両事業部に同じ番号がある）
    品番                ' C
    数量                ' D
    出荷済              ' E：分納の累計
    得意先              ' F
    納期                ' G
    状態                ' H：受付 → 引当済 → 出荷済
    次列
End Enum

' 製造指図（ヘッダ）
Public Enum 列_製造指図
    事業部 = 1          ' A
    指図No              ' B：MO-0001
    品番                ' C
    数量                ' D：指示数
    状態                ' E：計画 → 着手 → 完成
    着手予定            ' F
    完成予定            ' G
    受注No              ' H：受注生産のとき紐付く受注
    ピッキングNo        ' I：ピッキング指示中なら PK 番号
    出来高              ' J：完成したときの良品数
    備考                ' K：自動引当の結果など
    次列
End Enum

' 指図工程（製造指図の工程1つにつき1行。手配No.はこの行に振る）
Public Enum 列_指図工程
    事業部 = 1          ' A
    指図No              ' B
    工程順              ' C
    手配No              ' D：全社連番7桁（1000001～）。事業部をまたいで重複しない
    工程名              ' E
    作業区              ' F
    段取時間            ' G
    加工時間            ' H
    外注                ' I：○なら外注工程
    外注先              ' J
    外注単価            ' K
    外注LT              ' L
    状態                ' M：未 → (外注中) → 完了
    良品                ' N
    不良                ' O
    作業時間            ' P：分（段取を含む実績）
    次列
End Enum

' 発注（購入品）
Public Enum 列_発注
    事業部 = 1          ' A
    発注No              ' B：PO-0001
    品番                ' C
    数量                ' D
    入荷予定            ' E
    状態                ' F：発注済 → 入荷済
    次列
End Enum

' 外注注文（外注工程に出した仕掛品）
Public Enum 列_外注注文
    事業部 = 1          ' A
    外注No              ' B：SC-0001
    指図No              ' C
    工程順              ' D
    工程名              ' E
    品番                ' F
    外注先              ' G
    数量                ' H
    単価                ' I
    納期                ' J
    状態                ' K：外注中 → 受入済
    良品                ' L
    不良                ' M
    次列
End Enum

' ピッキング（ヘッダ）
Public Enum 列_ピッキング
    事業部 = 1          ' A
    ピッキングNo        ' B：PK-0001
    種別                ' C：製造／出荷
    対象伝票            ' D：製造なら指図No、出荷なら受注No
    内容                ' E
    状態                ' F：作業中 → 完了
    配送方法            ' G：出荷のとき
    出荷No              ' H：出荷を確定したときの出荷No
    次列
End Enum

' ピッキング明細
Public Enum 列_ピッキング明細
    事業部 = 1          ' A
    ピッキングNo        ' B
    行No                ' C
    品番                ' D
    数量                ' E
    棚番                ' F
    済                  ' G：○なら集め終わった
    次列
End Enum

' 出荷実績
Public Enum 列_出荷実績
    事業部 = 1          ' A
    出荷No              ' B：SH-0001
    受注No              ' C
    品番                ' D
    数量                ' E
    得意先              ' F
    出荷日              ' G
    配送方法            ' H
    販売単価            ' I：出荷した時点の単価を固定して持つ
    次列
End Enum

' 履歴（在庫の動き）
Public Enum 列_履歴
    事業部 = 1          ' A
    日時                ' B
    種別                ' C：払出・完成・入荷・出荷・出荷取消・調整
    品番                ' D
    増減                ' E
    残                  ' F：動いたあとの在庫数
    伝票                ' G
    次列
End Enum

' 採番（伝票番号の最終値）
Public Enum 列_採番
    事業部 = 1          ' A：加工／組立／全社（手配No.）
    種類                ' B：SO・MO・PO・SH・FC・SC・PK・手配No
    最終番号            ' C
    次列
End Enum

' 設定（項目と値）
Public Enum 列_設定
    項目 = 1            ' A
    値                  ' B
    説明                ' C
    次列
End Enum

' 在庫一覧（出力。現在の事業部の品目ごと）
Public Enum 列_在庫一覧
    品番 = 1            ' A
    品名                ' B
    区分                ' C
    生産方式            ' D：製造品だけ
    倉庫                ' E
    棚番                ' F
    在庫                ' G
    単位                ' H
    受注引当            ' I：引当済の受注の残数
    製造用確保          ' J：製造用ピッキングで確保した部材
    引当計              ' K：I＋J
    有効在庫            ' L：在庫－引当計（自由に使える数）
    手配済              ' M：製造中の見込み出来高＋発注残
    安全在庫            ' N
    注意                ' O：欠品／安全在庫割れ
    次列
End Enum

'==============================================================
' 共通部品
'==============================================================

' 小数2桁に丸める（HTML 版の r2 = Math.round(x*100)/100 と同じ。.5 は大きい方へ）
'   引数：Db_値　戻り値：丸めた値
'   VBA の Round は銀行型丸めで HTML と食い違うため使わない。
Public Function Fn_丸め(ByVal Db_値 As Double) As Double
    Fn_丸め = Int(Db_値 * 100 + 0.5) / 100
End Function

' 切り上げ（HTML 版の Math.ceil(x - 1e-9) と同じ。手配数を整数にするとき使う）
'   引数：Db_値　戻り値：切り上げた整数
Public Function Fn_切上げ(ByVal Db_値 As Double) As Double
    Fn_切上げ = -Int(-(Db_値 - Db_誤差))
End Function

' 事業部の表示名（加工 → 加工事業部）
Public Function Fn_事業部名(ByVal St_事業部 As String) As String
    Fn_事業部名 = St_事業部 & "事業部"
End Function

' 区分が製造品か（中間品・加工品・完成品は社内で作る。材料・部品は買う）
Public Function Fn_製造品か(ByVal St_区分 As String) As Boolean
    Fn_製造品か = (St_区分 = "中間品" Or St_区分 = "加工品" Or St_区分 = "完成品")
End Function

' セルの値を文字にそろえる（空・エラーは空文字。キーの比較に使う）
Public Function Fn_文字(ByVal Va_値 As Variant) As String
    If IsError(Va_値) Then
        Fn_文字 = ""
    ElseIf IsEmpty(Va_値) Or IsNull(Va_値) Then
        Fn_文字 = ""
    Else
        Fn_文字 = Trim$(CStr(Va_値))
    End If
End Function

' セルの値を数にそろえる（空・文字は 0。計算に使う）
Public Function Fn_数(ByVal Va_値 As Variant) As Double
    If IsError(Va_値) Then
        Fn_数 = 0
    ElseIf IsNumeric(Va_値) And Not IsEmpty(Va_値) Then
        Fn_数 = CDbl(Va_値)
    Else
        Fn_数 = 0
    End If
End Function

' 今日から何日後か（HTML 版の off。日付の列は日付で持つので差を取るだけ）
'   引数：Va_日付（セルの値）、Dt_今日　戻り値：日数（過去なら負）
Public Function Fn_日数差(ByVal Va_日付 As Variant, ByVal Dt_今日 As Date) As Long
    Fn_日数差 = CLng(Int(CDate(Va_日付)) - Int(Dt_今日))
End Function

' 日付を「m/d」で表す（HTML 版の md と同じ表示）
Public Function Fn_月日(ByVal Va_日付 As Variant) As String
    Fn_月日 = Month(CDate(Va_日付)) & "/" & Day(CDate(Va_日付))
End Function

' 表の最終行（A列で数える。データが無ければ見出し行を返す）
'   引数：Ws_表　戻り値：最終行番号
Public Function Fn_最終行(ByVal Ws_表 As Worksheet) As Long
    Dim Lg_行 As Long
    Lg_行 = Ws_表.Cells(Ws_表.Rows.Count, 1).End(xlUp).Row
    If Lg_行 < Lg_見出し行 Then Lg_行 = Lg_見出し行
    Fn_最終行 = Lg_行
End Function

' 表のデータ部分を2次元配列で一括して読む（1始まり。行×列）
'   引数：Ws_表、Lg_列数（列挙体の 次列 - 1）
'   戻り値：配列。データが0行なら Empty（呼ぶ側で IsEmpty を見る）
'   行数の多い読み書きは配列で一括（セルを1つずつ読むと遅いため）。
Public Function Fn_表読込(ByVal Ws_表 As Worksheet, ByVal Lg_列数 As Long) As Variant
    Dim Lg_最終 As Long
    Dim Va_表 As Variant
    Lg_最終 = Fn_最終行(Ws_表)
    If Lg_最終 < Lg_開始行 Then
        Fn_表読込 = Empty
        Exit Function
    End If
    Va_表 = Ws_表.Range(Ws_表.Cells(Lg_開始行, 1), Ws_表.Cells(Lg_最終, Lg_列数)).Value
    Fn_表読込 = Va_表
End Function

' 配列の行数（Empty なら 0）
Public Function Fn_行数(ByVal Va_表 As Variant) As Long
    If IsEmpty(Va_表) Then
        Fn_行数 = 0
    Else
        Fn_行数 = UBound(Va_表, 1)
    End If
End Function

' 表の末尾に1行追記する
'   引数：Ws_表、Va_行（0始まりの1次元配列。列挙体の並び順）
'   戻り値：書いた行番号（あとで同じ行を直すときに使う）
Public Function Fn_行追記(ByVal Ws_表 As Worksheet, ByVal Va_行 As Variant) As Long
    Dim Lg_行 As Long
    Dim Lg_数 As Long
    Lg_行 = Fn_最終行(Ws_表) + 1
    Lg_数 = UBound(Va_行) - LBound(Va_行) + 1
    Ws_表.Range(Ws_表.Cells(Lg_行, 1), Ws_表.Cells(Lg_行, Lg_数)).Value = Va_行
    Fn_行追記 = Lg_行
End Function

' 設定シートの値を読む（項目名で探す。無ければ空）
'   引数：St_項目　戻り値：値
'   各モジュールの Pr_変数格納 から呼ぶ（部品ごとに読み直さないため）。
Public Function Fn_設定値(ByVal St_項目 As String) As Variant
    Dim Ws_設定 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Set Ws_設定 = ThisWorkbook.Worksheets(St_シート_設定)
    Va_表 = Fn_表読込(Ws_設定, 列_設定.次列 - 1)
    Fn_設定値 = ""
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_設定.項目)) = St_項目 Then
            Fn_設定値 = Va_表(Lg_行, 列_設定.値)
            Exit Function
        End If
    Next Lg_行
End Function

' 設定シートに値を書く（項目名で探す。無ければ末尾に足す）
Public Sub Pr_設定値書込(ByVal St_項目 As String, ByVal Va_値 As Variant)
    Dim Ws_設定 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Set Ws_設定 = ThisWorkbook.Worksheets(St_シート_設定)
    Va_表 = Fn_表読込(Ws_設定, 列_設定.次列 - 1)
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_設定.項目)) = St_項目 Then
            Ws_設定.Cells(Lg_行 + Lg_開始行 - 1, 列_設定.値).Value = Va_値
            Exit Sub
        End If
    Next Lg_行
    Call Fn_行追記(Ws_設定, Array(St_項目, Va_値, ""))
End Sub

' シートがあるか
Public Function Fn_シートあり(ByVal St_名前 As String) As Boolean
    Dim Ws_表 As Worksheet
    Fn_シートあり = False
    For Each Ws_表 In ThisWorkbook.Worksheets
        If Ws_表.Name = St_名前 Then
            Fn_シートあり = True
            Exit Function
        End If
    Next Ws_表
End Function

' 必要なシートがそろっているか確かめる。無ければ初期化を促して中断
'   引数：Va_名前（シート名の配列）
Public Sub Pr_シート存在確認(ByVal Va_名前 As Variant)
    Dim Va_名 As Variant
    For Each Va_名 In Va_名前
        If Not Fn_シートあり(CStr(Va_名)) Then
            MsgBox "シート「" & Va_名 & "」がありません。" & vbCrLf & _
                   "先にメニューの「初期データに戻す」（または Pr_初期化）を実行してください。", vbExclamation
            Err.Raise Lg_中断番号
        End If
    Next Va_名
End Sub

' 画面更新・自動計算を止める（処理を速くし、途中の画面のちらつきを防ぐ）
Public Sub Pr_画面更新停止()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' 画面更新・自動計算を戻す
Public Sub Pr_画面更新再開()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' 終了処理：止めた設定を戻し、本当のエラーのときだけ内容を表示する
'   引数：Lg_番号（Err.Number）、St_内容（Err.Description）
'   正常（0）と中断（Lg_中断番号）は何も表示しない。中断の理由は、
'   中断を上げた側がすでにメッセージで伝えているため。
Public Sub Pr_終了処理(ByVal Lg_番号 As Long, ByVal St_内容 As String)
    Call Pr_画面更新再開
    If Lg_番号 <> 0 And Lg_番号 <> Lg_中断番号 Then
        MsgBox "処理を中断しました（エラー " & Lg_番号 & "）：" & vbCrLf & St_内容, vbExclamation
    End If
End Sub

' フォームのボタンを1つ置く
'   引数：Ws_置き場所、St_セル（左上）、Lg_幅のセル数、St_表示、St_マクロ名、St_ボタン名
Public Sub Pr_ボタン作成(ByVal Ws_置き場所 As Worksheet, ByVal St_セル As String, ByVal Lg_幅のセル数 As Long, _
        ByVal St_表示 As String, ByVal St_マクロ名 As String, ByVal St_ボタン名 As String)
    Dim Rg_位置 As Range
    Dim Ob_ボタン As Object
    Set Rg_位置 = Ws_置き場所.Range(St_セル).Resize(1, Lg_幅のセル数)
    Set Ob_ボタン = Ws_置き場所.Buttons.Add(Rg_位置.Left + 2, Rg_位置.Top + 1, Rg_位置.Width - 4, Rg_位置.Height + 6)
    Ob_ボタン.Caption = St_表示
    Ob_ボタン.OnAction = St_マクロ名
    Ob_ボタン.Name = St_ボタン名
End Sub

' 表の中から、事業部とキーが一致する行を探す
'   引数：Ws_表、Lg_列数、St_事業部、Lg_キー列、St_キー
'   戻り値：シート上の行番号。無ければ 0
Public Function Fn_行検索(ByVal Ws_表 As Worksheet, ByVal Lg_列数 As Long, ByVal St_事業部 As String, _
        ByVal Lg_キー列 As Long, ByVal St_キー As String) As Long
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Va_表 = Fn_表読込(Ws_表, Lg_列数)
    Fn_行検索 = 0
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 1)) = St_事業部 And Fn_文字(Va_表(Lg_行, Lg_キー列)) = St_キー Then
            Fn_行検索 = Lg_行 + Lg_開始行 - 1
            Exit Function
        End If
    Next Lg_行
End Function

' 品目マスタの1行を、列_品目 の番号で引ける1次元配列で返す
'   引数：St_事業部、St_品番　戻り値：配列（Va(列_品目.区分) のように使う）。無ければ Empty
Public Function Fn_品目情報(ByVal St_事業部 As String, ByVal St_品番 As String) As Variant
    Dim Ws_品目 As Worksheet
    Dim Lg_行 As Long
    Dim Lg_列 As Long
    Dim Va_行() As Variant
    Set Ws_品目 = ThisWorkbook.Worksheets(St_シート_品目)
    Lg_行 = Fn_行検索(Ws_品目, 列_品目.次列 - 1, St_事業部, 列_品目.品番, St_品番)
    If Lg_行 = 0 Or St_品番 = "" Then
        Fn_品目情報 = Empty
        Exit Function
    End If
    ReDim Va_行(1 To 列_品目.次列 - 1)
    For Lg_列 = 1 To 列_品目.次列 - 1
        Va_行(Lg_列) = Ws_品目.Cells(Lg_行, Lg_列).Value
    Next Lg_列
    Fn_品目情報 = Va_行
End Function

' 品目が小数の数量を扱えるか（小数可が○、または単位が kg。HTML 版の dec||unit==='kg'）
Public Function Fn_小数可か(ByVal Va_品目 As Variant) As Boolean
    Fn_小数可か = (Fn_文字(Va_品目(列_品目.小数可)) = St_印 Or Fn_文字(Va_品目(列_品目.単位)) = "kg")
End Function

' 入力された数量を確かめて返す（HTML 版の qtyOf）
'   引数：Va_値（セルの値）、Bl_小数可
'   戻り値：小数2桁に丸めた数量。0 以下・数でない・整数のみの品目で小数 のときは 0
Public Function Fn_数量変換(ByVal Va_値 As Variant, ByVal Bl_小数可 As Boolean) As Double
    Dim Db_値 As Double
    Fn_数量変換 = 0
    If Fn_文字(Va_値) = "" Or Not IsNumeric(Va_値) Then Exit Function
    Db_値 = CDbl(Va_値)
    If Db_値 <= 0 Then Exit Function
    If Not Bl_小数可 And Db_値 <> Int(Db_値) Then Exit Function
    Fn_数量変換 = Fn_丸め(Db_値)
End Function

' 入力された日数（今日から何日後）を確かめて返す
'   引数：Va_値、Lg_最小、Lg_最大　戻り値：日数。整数でない・範囲外は -1
Public Function Fn_日数変換(ByVal Va_値 As Variant, ByVal Lg_最小 As Long, ByVal Lg_最大 As Long) As Long
    Dim Db_値 As Double
    Fn_日数変換 = -1
    If Fn_文字(Va_値) = "" Or Not IsNumeric(Va_値) Then Exit Function
    Db_値 = CDbl(Va_値)
    If Db_値 <> Int(Db_値) Or Db_値 < Lg_最小 Or Db_値 > Lg_最大 Then Exit Function
    Fn_日数変換 = CLng(Db_値)
End Function

' 入力の不備を伝えて中断する（入力チェックの部品から呼ぶ）
'   引数：St_文（何が悪いか・どうすればよいか）
Public Sub Pr_入力エラー(ByVal St_文 As String)
    MsgBox St_文, vbExclamation, "入力を確認してください"
    Err.Raise Lg_中断番号
End Sub

' 取り消し・削除の前に確認する。「いいえ」なら中断
Public Sub Pr_削除確認(ByVal St_文 As String)
    If MsgBox(St_文 & vbCrLf & vbCrLf & "よろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認") = vbNo Then
        Err.Raise Lg_中断番号
    End If
End Sub

' 入力シートの「結果」欄に、いま行った処理の結果を書く
Public Sub Pr_結果表示(ByVal Ws_入力 As Worksheet, ByVal St_文 As String)
    Ws_入力.Range(St_入力_結果セル).Value = "結果：" & St_文 & "（" & Format$(Now, "hh:nn") & "）"
End Sub

' データシートの、現在の事業部の行を入力シートの下側に写す（新しいものを上に）
'   引数：Ws_出力、Lg_開始行（見出しを書く行）、St_元シート、Lg_列数、St_事業部、St_題名
'   入力シートで、いま操作した結果をその場で見られるようにするため（元のデータはデータシート）。
Public Sub Pr_一覧転記(ByVal Ws_出力 As Worksheet, ByVal Lg_開始行 As Long, ByVal St_元シート As String, _
        ByVal Lg_列数 As Long, ByVal St_事業部 As String, ByVal St_題名 As String)
    Dim Ws_元 As Worksheet
    Dim Va_表 As Variant
    Dim Va_出力() As Variant
    Dim Lg_行 As Long
    Dim Lg_列 As Long
    Dim Lg_件数 As Long
    Set Ws_元 = ThisWorkbook.Worksheets(St_元シート)
    Va_表 = Fn_表読込(Ws_元, Lg_列数)
    ' 前回の一覧を消す
    Ws_出力.Range(Ws_出力.Cells(Lg_開始行, 1), Ws_出力.Cells(Ws_出力.Rows.Count, Lg_列数 + 1)).Clear
    Ws_出力.Cells(Lg_開始行, 2).Value = "■ " & St_題名 & "（" & Fn_事業部名(St_事業部) & "。新しいものが上。元のデータは「" & St_元シート & "」シート）"
    Ws_出力.Cells(Lg_開始行, 2).Font.Bold = True
    ' 見出し（事業部の列を除く）
    For Lg_列 = 2 To Lg_列数
        Ws_出力.Cells(Lg_開始行 + 1, Lg_列).Value = Ws_元.Cells(Lg_見出し行, Lg_列).Value
    Next Lg_列
    With Ws_出力.Range(Ws_出力.Cells(Lg_開始行 + 1, 2), Ws_出力.Cells(Lg_開始行 + 1, Lg_列数))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    ' 現在の事業部の行を、下から（新しい順に）集める
    Lg_件数 = 0
    For Lg_行 = Fn_行数(Va_表) To 1 Step -1
        If Fn_文字(Va_表(Lg_行, 1)) = St_事業部 Then Lg_件数 = Lg_件数 + 1
    Next Lg_行
    If Lg_件数 = 0 Then
        Ws_出力.Cells(Lg_開始行 + 2, 2).Value = "（ありません）"
        Exit Sub
    End If
    ReDim Va_出力(1 To Lg_件数, 1 To Lg_列数 - 1)
    Lg_件数 = 0
    For Lg_行 = Fn_行数(Va_表) To 1 Step -1
        If Fn_文字(Va_表(Lg_行, 1)) = St_事業部 Then
            Lg_件数 = Lg_件数 + 1
            For Lg_列 = 2 To Lg_列数
                Va_出力(Lg_件数, Lg_列 - 1) = Va_表(Lg_行, Lg_列)
            Next Lg_列
        End If
    Next Lg_行
    Ws_出力.Range(Ws_出力.Cells(Lg_開始行 + 2, 2), Ws_出力.Cells(Lg_開始行 + 1 + Lg_件数, Lg_列数)).Value = Va_出力
    ' 表示形式は元のシートの2行目に合わせる（日付を数字で見せないため）
    For Lg_列 = 2 To Lg_列数
        Ws_出力.Range(Ws_出力.Cells(Lg_開始行 + 2, Lg_列), Ws_出力.Cells(Lg_開始行 + 1 + Lg_件数, Lg_列)).NumberFormat = _
            Ws_元.Cells(Lg_開始行, Lg_列).NumberFormat
    Next Lg_列
End Sub

' 入力シートの土台（題名・説明・結果欄）を作る（初期化から呼ぶ）
'   引数：Ws_入力、St_題名、St_説明
Public Sub Pr_入力シート枠(ByVal Ws_入力 As Worksheet, ByVal St_題名 As String, ByVal St_説明 As String)
    Ws_入力.Cells.Clear
    Ws_入力.Columns("A").ColumnWidth = 2
    Ws_入力.Columns("B").ColumnWidth = 24
    Ws_入力.Columns("C:L").ColumnWidth = 13
    Ws_入力.Range("B1").Value = St_題名
    Ws_入力.Range("B1").Font.Size = 14
    Ws_入力.Range("B1").Font.Bold = True
    Ws_入力.Range("B2").Value = St_説明
    Ws_入力.Range(St_入力_結果セル).Value = "結果："
    Ws_入力.Range(St_入力_結果セル).Font.Color = RGB(0, 97, 0)
End Sub

' 入力欄を1つ作る（見出しと、入力するセル。入力セルは黄色）
'   引数：Ws_入力、St_セル（入力するセル。見出しはその左）、St_見出し、Va_初期値、St_候補（ドロップダウンの名前。空なら無し）
Public Sub Pr_入力欄(ByVal Ws_入力 As Worksheet, ByVal St_セル As String, ByVal St_見出し As String, _
        ByVal Va_初期値 As Variant, ByVal St_候補 As String)
    Ws_入力.Range(St_セル).Offset(0, -1).Value = St_見出し
    Ws_入力.Range(St_セル).Value = Va_初期値
    Ws_入力.Range(St_セル).Interior.Color = RGB(255, 255, 204)
    Ws_入力.Range(St_セル).Borders.LineStyle = xlContinuous
    If St_候補 <> "" Then
        Ws_入力.Range(St_セル).Validation.Delete
        Ws_入力.Range(St_セル).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & St_候補
    End If
End Sub

' 入力シートの見出し（■ ○○）を書く
Public Sub Pr_入力見出し(ByVal Ws_入力 As Worksheet, ByVal St_セル As String, ByVal St_文 As String)
    Ws_入力.Range(St_セル).Value = "■ " & St_文
    Ws_入力.Range(St_セル).Font.Bold = True
End Sub

' データシートの行を消す（下の行から順に消す。行番号がずれないように）
'   引数：Ws_表、Va_行（シート上の行番号の配列）
Public Sub Pr_行削除(ByVal Ws_表 As Worksheet, ByVal Va_行 As Variant)
    Dim Lg_添字 As Long
    Dim Lg_比較 As Long
    Dim Va_一時 As Variant
    If IsEmpty(Va_行) Then Exit Sub
    ' 大きい行番号から順に並べる
    For Lg_添字 = LBound(Va_行) To UBound(Va_行) - 1
        For Lg_比較 = Lg_添字 + 1 To UBound(Va_行)
            If Va_行(Lg_比較) > Va_行(Lg_添字) Then
                Va_一時 = Va_行(Lg_添字)
                Va_行(Lg_添字) = Va_行(Lg_比較)
                Va_行(Lg_比較) = Va_一時
            End If
        Next Lg_比較
    Next Lg_添字
    For Lg_添字 = LBound(Va_行) To UBound(Va_行)
        Ws_表.Rows(Va_行(Lg_添字)).Delete
    Next Lg_添字
End Sub
