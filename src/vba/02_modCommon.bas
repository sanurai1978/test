Attribute VB_Name = "modCommon"
Option Explicit

' ============================================================
' 共通モジュール
' 前提   : modConstants がロードされていること
' 入力   : 各Functionの引数を参照
' 出力   : 各Functionの戻り値を参照
' 例外   : シートが存在しない場合はエラーを呼び出し元に伝播する
' 拡張余地: ログ出力先を外部ファイルに変更する場合は
'           Pr_操作ログ出力 のみ改修すればよい
' ============================================================

Private Bln_元画面更新設定 As Boolean
Private Lng_元計算方法設定 As Long
Private Bln_元イベント設定 As Boolean
Private Bln_元アラート設定 As Boolean

' ------------------------------------------------------------
' Pr_環境抑止
' 前提: 各Public Subの先頭(手順1)で必ず呼び出すこと
' 入力: なし
' 出力: なし(現在の設定をモジュールレベル変数に退避)
' 例外: なし
' 拡張余地: なし
' ------------------------------------------------------------
Public Sub Pr_環境抑止()
    Bln_元画面更新設定 = Application.ScreenUpdating
    Lng_元計算方法設定 = Application.Calculation
    Bln_元イベント設定 = Application.EnableEvents
    Bln_元アラート設定 = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

' ------------------------------------------------------------
' Pr_環境復元
' 前提: Pr_環境抑止 と対で、各Public Subの末尾(手順8)で必ず呼び出す。
'       正常終了・異常終了(エラーハンドラ)いずれの経路からも呼ぶこと
' 入力: なし
' 出力: なし(退避しておいた設定に戻す)
' 例外: なし
' 拡張余地: なし
' ------------------------------------------------------------
Public Sub Pr_環境復元()
    Application.ScreenUpdating = Bln_元画面更新設定
    Application.Calculation = Lng_元計算方法設定
    Application.EnableEvents = Bln_元イベント設定
    Application.DisplayAlerts = Bln_元アラート設定
End Sub

' ------------------------------------------------------------
' Fn_シート取得
' 前提  : ブック内に対象シートが存在すること
' 入力  : Enm_種別 - Enm_シート種別
' 出力  : Worksheetオブジェクト
' 例外  : シートが存在しない場合は実行時エラー(呼び出し元でOn Error捕捉)
' 拡張余地: なし
' ------------------------------------------------------------
Public Function Fn_シート取得(ByVal Enm_種別 As Enm_シート種別) As Worksheet
    Dim Str_シート名 As String
    Str_シート名 = Fn_シート名取得(Enm_種別)
    Set Fn_シート取得 = ThisWorkbook.Worksheets(Str_シート名)
End Function

' ------------------------------------------------------------
' Fn_ヘッダー列マップ取得
' 前提  : 対象シートの1行目(Lng_ヘッダー行)に列見出しが入力されていること
' 入力  : Wsh_対象 - 対象ワークシート
' 出力  : Dictionary(キー:見出し文字列, 値:列番号Long)
' 例外  : 空白セルに達した時点で走査を終了する
' 拡張余地: 列の追加・入替が発生してもコード修正不要
' ------------------------------------------------------------
Public Function Fn_ヘッダー列マップ取得(ByVal Wsh_対象 As Worksheet) As Object
    Dim Dic_列 As Object
    Dim Lng_列 As Long
    Dim Str_見出し As String

    Set Dic_列 = CreateObject("Scripting.Dictionary")
    Lng_列 = 1

    Do
        Str_見出し = Trim(CStr(Wsh_対象.Cells(Lng_ヘッダー行, Lng_列).Value))
        If Str_見出し = "" Then Exit Do
        If Not Dic_列.Exists(Str_見出し) Then
            Dic_列.Add Str_見出し, Lng_列
        End If
        Lng_列 = Lng_列 + 1
    Loop

    Set Fn_ヘッダー列マップ取得 = Dic_列
End Function

' ------------------------------------------------------------
' Fn_最終行取得
' 前提  : キー列に隙間なくデータが入力されていること
' 入力  : Wsh_対象, Lng_キー列 - 存在確認に使う列番号
' 出力  : データが存在する最終行番号(データなしの場合は Lng_データ開始行 - 1)
' 例外  : なし
' 拡張余地: なし
' ------------------------------------------------------------
Public Function Fn_最終行取得(ByVal Wsh_対象 As Worksheet, ByVal Lng_キー列 As Long) As Long
    Dim Lng_行 As Long
    Lng_行 = Wsh_対象.Cells(Wsh_対象.Rows.Count, Lng_キー列).End(xlUp).Row
    If Lng_行 < Lng_データ開始行 Then
        Fn_最終行取得 = Lng_データ開始行 - 1
    Else
        Fn_最終行取得 = Lng_行
    End If
End Function

' ------------------------------------------------------------
' Fn_次の連番取得
' 前提  : ID列には数値の連番のみが入力されていること
' 入力  : Wsh_対象, Lng_ID列 - 連番が格納された列番号
' 出力  : 現在の最大値+1(データなしの場合は1)
' 例外  : なし
' 拡張余地: 採番方式を年月+連番などに変更する場合は本Functionのみ改修
' ------------------------------------------------------------
Public Function Fn_次の連番取得(ByVal Wsh_対象 As Worksheet, ByVal Lng_ID列 As Long) As Long
    Dim Lng_最終行 As Long
    Dim Lng_行 As Long
    Dim Lng_最大値 As Long

    Lng_最終行 = Fn_最終行取得(Wsh_対象, Lng_ID列)
    Lng_最大値 = 0

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If IsNumeric(Wsh_対象.Cells(Lng_行, Lng_ID列).Value) Then
            If CLng(Wsh_対象.Cells(Lng_行, Lng_ID列).Value) > Lng_最大値 Then
                Lng_最大値 = CLng(Wsh_対象.Cells(Lng_行, Lng_ID列).Value)
            End If
        End If
    Next Lng_行

    Fn_次の連番取得 = Lng_最大値 + 1
End Function

' ------------------------------------------------------------
' Pr_操作ログ出力
' 前提  : 操作ログシートの1行目に見出しが入力されていること
' 入力  : Str_処理名, Lng_件数, Enm_結果(Enm_ログ結果), Str_備考
' 出力  : 操作ログシート末尾に1行追加(削除・上書きは行わない)
' 例外  : なし
' 拡張余地: 出力先を外部ログファイルに変更する場合は本Subのみ改修
' ------------------------------------------------------------
Public Sub Pr_操作ログ出力(ByVal Str_処理名 As String, ByVal Lng_件数 As Long, _
                           ByVal Enm_結果 As Enm_ログ結果, ByVal Str_備考 As String)
    Dim Wsh_ログ As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh_ログ = Fn_シート取得(シート_操作ログ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh_ログ)
    Lng_行 = Fn_最終行取得(Wsh_ログ, Dic_列("実行日時")) + 1

    Wsh_ログ.Cells(Lng_行, Dic_列("実行日時")).Value = Now
    Wsh_ログ.Cells(Lng_行, Dic_列("処理名")).Value = Str_処理名
    Wsh_ログ.Cells(Lng_行, Dic_列("件数")).Value = Lng_件数
    Wsh_ログ.Cells(Lng_行, Dic_列("エラー有無")).Value = IIf(Enm_結果 = ログ_エラー, "エラー", "正常")
    Wsh_ログ.Cells(Lng_行, Dic_列("備考")).Value = Str_備考
End Sub

' ------------------------------------------------------------
' Fn_文字列有無判定
' 入力: Str_対象
' 出力: Trim後の文字列長が0より大きければTrue
' ------------------------------------------------------------
Public Function Fn_文字列有無判定(ByVal Str_対象 As String) As Boolean
    Fn_文字列有無判定 = (Len(Trim(Str_対象)) > 0)
End Function

' ------------------------------------------------------------
' Fn_年月文字列判定
' 入力: Str_対象 - "YYYYMM"形式かどうかを判定する文字列
' 出力: 6桁数値かつ月部分が01~12であればTrue
' ------------------------------------------------------------
Public Function Fn_年月文字列判定(ByVal Str_対象 As String) As Boolean
    Dim Lng_月 As Long
    If Len(Str_対象) <> 6 Then
        Fn_年月文字列判定 = False
        Exit Function
    End If
    If Not IsNumeric(Str_対象) Then
        Fn_年月文字列判定 = False
        Exit Function
    End If
    Lng_月 = CLng(Mid(Str_対象, 5, 2))
    Fn_年月文字列判定 = (Lng_月 >= 1 And Lng_月 <= 12)
End Function

' ------------------------------------------------------------
' Fn_年月末日取得
' 入力: Str_年月 - "YYYYMM"形式文字列
' 出力: 当該年月の末日(Date型)
' ------------------------------------------------------------
Public Function Fn_年月末日取得(ByVal Str_年月 As String) As Date
    Dim Lng_年 As Long, Lng_月 As Long
    Lng_年 = CLng(Left(Str_年月, 4))
    Lng_月 = CLng(Mid(Str_年月, 5, 2))
    Fn_年月末日取得 = DateSerial(Lng_年, Lng_月 + 1, 0)
End Function

' ------------------------------------------------------------
' Fn_設定値取得
' 前提: 設定シートが「設定キー」「設定値」列を持つこと
' 入力: Str_設定キー
' 出力: 設定シートに登録された設定値(見つからない場合は空文字)
' ------------------------------------------------------------
Public Function Fn_設定値取得(ByVal Str_設定キー As String) As String
    Dim Wsh_設定 As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long, Lng_最終行 As Long

    Set Wsh_設定 = Fn_シート取得(シート_設定)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh_設定)
    Lng_最終行 = Fn_最終行取得(Wsh_設定, Dic_列("設定キー"))

    Fn_設定値取得 = ""
    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Trim(CStr(Wsh_設定.Cells(Lng_行, Dic_列("設定キー")).Value)) = Str_設定キー Then
            Fn_設定値取得 = Trim(CStr(Wsh_設定.Cells(Lng_行, Dic_列("設定値")).Value))
            Exit Function
        End If
    Next Lng_行
End Function
