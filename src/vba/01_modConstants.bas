Attribute VB_Name = "modConstants"
Option Explicit

' ============================================================
' 定数定義モジュール
' シート名、Enum、共通型を集約管理する。
' 列番号のハードコーディングを避けるため、本モジュールでは
' 列"名称"のみを定義し、実際の列番号はヘッダー列マップ(Dictionary)
' から動的に取得すること。(modCommon.Fn_ヘッダー列マップ取得 を使用)
' ============================================================

' --- シート種別 ---
Public Enum Enm_シート種別
    シート_会社マスタ = 1
    シート_会員マスタ
    シート_拠出設定
    シート_月次実績
    シート_イベント履歴
    シート_株価マスタ
    シート_イベント入力
    シート_設定
    シート_操作ログ
    シート_スナップショット
End Enum

' --- イベント種別 ---
Public Enum Enm_イベント種別
    イベント_入会 = 1
    イベント_退会
    イベント_休会
    イベント_復帰
    イベント_口数変更
End Enum

' --- 処理状態(イベント入力/イベント履歴 共通) ---
Public Enum Enm_処理状態
    処理状態_未処理 = 0
    処理状態_処理済 = 1
    処理状態_エラー = 2
End Enum

' --- 会員状態(会員マスタ) ---
Public Enum Enm_会員状態
    会員状態_在籍 = 1
    会員状態_退会済 = 2
    会員状態_休会中 = 3
End Enum

' --- 拠出設定状態 ---
Public Enum Enm_拠出設定状態
    拠出設定_有効 = 1
    拠出設定_終了 = 2
End Enum

' --- 操作ログ エラー有無 ---
Public Enum Enm_ログ結果
    ログ_正常 = 0
    ログ_エラー = 1
End Enum

' --- 行位置に関する定数 ---
Public Const Lng_ヘッダー行 As Long = 1
Public Const Lng_データ開始行 As Long = 2

' --- 共通型 ---
' イベント入力シート1行分の検証済みデータを保持する型
Public Type Typ_イベント入力行
    Lng_行番号 As Long
    Str_MemberNo As String
    Str_氏名 As String
    Str_会社コード As String
    Lng_種別 As Long          ' Enm_イベント種別
    Dte_発生日 As Date
    Dbl_口数 As Double
    Bln_口数指定あり As Boolean
    Bln_新規会員 As Boolean
    Bln_エラーあり As Boolean
    Str_エラー内容 As String
End Type

' 1回の業務処理実行結果を保持する型
Public Type Typ_処理結果
    Lng_成功件数 As Long
    Lng_エラー件数 As Long
    Str_メッセージ As String
End Type

' ============================================================
' Fn_シート名取得
' 前提   : Enm_シート種別 が有効な値であること
' 入力   : Enm_対象 - シート種別
' 出力   : シート名(String)
' 例外   : 未定義の種別が渡された場合は空文字を返す
' 拡張余地: シート名を設定シートで管理し外部化することも可能
' ============================================================
Public Function Fn_シート名取得(ByVal Enm_対象 As Enm_シート種別) As String
    Select Case Enm_対象
        Case シート_会社マスタ: Fn_シート名取得 = "会社マスタ"
        Case シート_会員マスタ: Fn_シート名取得 = "会員マスタ"
        Case シート_拠出設定: Fn_シート名取得 = "拠出設定"
        Case シート_月次実績: Fn_シート名取得 = "月次実績"
        Case シート_イベント履歴: Fn_シート名取得 = "イベント履歴"
        Case シート_株価マスタ: Fn_シート名取得 = "株価マスタ"
        Case シート_イベント入力: Fn_シート名取得 = "イベント入力"
        Case シート_設定: Fn_シート名取得 = "設定"
        Case シート_操作ログ: Fn_シート名取得 = "操作ログ"
        Case シート_スナップショット: Fn_シート名取得 = "スナップショット"
        Case Else: Fn_シート名取得 = ""
    End Select
End Function

' ============================================================
' Fn_イベント種別文字列
' 入力: Lng_種別(Enm_イベント種別)
' 出力: 表示用日本語文字列
' ============================================================
Public Function Fn_イベント種別文字列(ByVal Lng_種別 As Long) As String
    Select Case Lng_種別
        Case イベント_入会: Fn_イベント種別文字列 = "入会"
        Case イベント_退会: Fn_イベント種別文字列 = "退会"
        Case イベント_休会: Fn_イベント種別文字列 = "休会"
        Case イベント_復帰: Fn_イベント種別文字列 = "休会から復帰"
        Case イベント_口数変更: Fn_イベント種別文字列 = "口数変更"
        Case Else: Fn_イベント種別文字列 = "不明"
    End Select
End Function

' ============================================================
' Fn_イベント種別変換
' 入力: Str_種別文字列 - イベント入力シートに記載された文字列
' 出力: Enm_イベント種別 に対応するLong値。該当なしは0
' ============================================================
Public Function Fn_イベント種別変換(ByVal Str_種別文字列 As String) As Long
    Select Case Trim(Str_種別文字列)
        Case "入会": Fn_イベント種別変換 = イベント_入会
        Case "退会": Fn_イベント種別変換 = イベント_退会
        Case "休会": Fn_イベント種別変換 = イベント_休会
        Case "休会から復帰": Fn_イベント種別変換 = イベント_復帰
        Case "口数変更": Fn_イベント種別変換 = イベント_口数変更
        Case Else: Fn_イベント種別変換 = 0
    End Select
End Function

' ============================================================
' Fn_処理状態文字列 / Fn_会員状態文字列 / Fn_拠出設定状態文字列
' 各Enumの表示用文字列変換
' ============================================================
Public Function Fn_処理状態文字列(ByVal Lng_状態 As Long) As String
    Select Case Lng_状態
        Case 処理状態_未処理: Fn_処理状態文字列 = "未処理"
        Case 処理状態_処理済: Fn_処理状態文字列 = "処理済"
        Case 処理状態_エラー: Fn_処理状態文字列 = "エラー"
        Case Else: Fn_処理状態文字列 = ""
    End Select
End Function

Public Function Fn_会員状態文字列(ByVal Lng_状態 As Long) As String
    Select Case Lng_状態
        Case 会員状態_在籍: Fn_会員状態文字列 = "在籍"
        Case 会員状態_退会済: Fn_会員状態文字列 = "退会済"
        Case 会員状態_休会中: Fn_会員状態文字列 = "休会中"
        Case Else: Fn_会員状態文字列 = ""
    End Select
End Function

Public Function Fn_拠出設定状態文字列(ByVal Lng_状態 As Long) As String
    Select Case Lng_状態
        Case 拠出設定_有効: Fn_拠出設定状態文字列 = "有効"
        Case 拠出設定_終了: Fn_拠出設定状態文字列 = "終了"
        Case Else: Fn_拠出設定状態文字列 = ""
    End Select
End Function
