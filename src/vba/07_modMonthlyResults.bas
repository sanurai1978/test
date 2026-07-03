Attribute VB_Name = "modMonthlyResults"
Option Explicit

' ============================================================
' 月次実績更新モジュール
' 前提   : 拠出設定履歴と株価マスタを基に対象年月の実績を再計算する
' 入力   : Str_対象年月(省略時は設定シートの「処理対象年月」を使用)
' 出力   : 月次実績シートを更新(同一年月×MemberNo×会社コードは上書き、
'          物理削除は行わず何度実行しても結果が一致する=冪等)
' 例外   : 株価マスタに単価が存在しない会員はエラー件数として計上し
'          処理は継続する
' 拡張余地: 会社コードごとの拠出単価が異なる場合はFn_単価取得の
'           検索条件を拡張する
' ============================================================

' ------------------------------------------------------------
' Pr_月次実績更新実行
' 目次:
'   1. 環境抑止
'   2. 入力取得
'   3. 入力チェック
'   4. マスタ取得
'   5. 業務処理
'   6. 出力
'   7. ログ
'   8. 環境復元
' ------------------------------------------------------------
Public Sub Pr_月次実績更新実行(Optional ByVal Str_対象年月 As String = "")
    On Error GoTo ErrHandler

    Dim Dic_拠出設定 As Object
    Dim Wsh_会員 As Worksheet
    Dim Dic_会員列 As Object
    Dim Vnt_MemberNo As Variant
    Dim Lng_成功 As Long, Lng_エラー As Long
    Dim Str_会社コード As String
    Dim Dbl_単価 As Double
    Dim Lng_会員行 As Long

    ' 1. 環境抑止
    Pr_環境抑止

    ' 2. 入力取得
    If Not Fn_文字列有無判定(Str_対象年月) Then
        Str_対象年月 = Fn_設定値取得("処理対象年月")
    End If

    ' 3. 入力チェック
    If Not Fn_年月文字列判定(Str_対象年月) Then
        Pr_操作ログ出力 "月次実績更新", 0, ログ_エラー, "対象年月が不正です(" & Str_対象年月 & ")"
        Pr_環境復元
        MsgBox "対象年月が不正です。設定シートの「処理対象年月」を確認してください。", vbExclamation
        Exit Sub
    End If

    ' 4. マスタ取得
    Set Dic_拠出設定 = Fn_対象年月の拠出設定取得(Str_対象年月)
    Set Wsh_会員 = Fn_シート取得(シート_会員マスタ)
    Set Dic_会員列 = Fn_ヘッダー列マップ取得(Wsh_会員)

    Lng_成功 = 0
    Lng_エラー = 0

    ' 5. 業務処理 + 6. 出力(会員ごとに判定しながらUpsert)
    For Each Vnt_MemberNo In Dic_拠出設定.Keys
        Lng_会員行 = Fn_会員行検索(Wsh_会員, Dic_会員列, CStr(Vnt_MemberNo))
        If Lng_会員行 = 0 Then
            Lng_エラー = Lng_エラー + 1
        Else
            Str_会社コード = Trim(CStr(Wsh_会員.Cells(Lng_会員行, Dic_会員列("会社コード")).Value))
            Dbl_単価 = Fn_単価取得(Str_対象年月, Str_会社コード)

            If Dbl_単価 <= 0 Then
                Lng_エラー = Lng_エラー + 1
            Else
                Pr_月次実績Upsert Str_対象年月, CStr(Vnt_MemberNo), Str_会社コード, _
                    Dic_拠出設定(Vnt_MemberNo), Dbl_単価
                Lng_成功 = Lng_成功 + 1
            End If
        End If
    Next Vnt_MemberNo

    ' 7. ログ
    Pr_操作ログ出力 "月次実績更新(" & Str_対象年月 & ")", Lng_成功 + Lng_エラー, _
        IIf(Lng_エラー > 0, ログ_エラー, ログ_正常), _
        "成功:" & Lng_成功 & "件 エラー:" & Lng_エラー & "件"

    ' 8. 環境復元
    Pr_環境復元
    Exit Sub

ErrHandler:
    Pr_操作ログ出力 "月次実績更新", 0, ログ_エラー, "実行時エラー:" & Err.Number & " " & Err.Description
    Pr_環境復元
    MsgBox "月次実績更新でエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' ------------------------------------------------------------
' Fn_会員行検索
' 入力: Wsh(会員マスタ), Dic_列, Str_MemberNo
' 出力: 該当行番号(見つからない場合は0)
' ------------------------------------------------------------
Private Function Fn_会員行検索(ByVal Wsh As Worksheet, ByVal Dic_列 As Object, ByVal Str_MemberNo As String) As Long
    Dim Lng_最終行 As Long, Lng_行 As Long

    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("MemberNo"))
    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value)) = Str_MemberNo Then
            Fn_会員行検索 = Lng_行
            Exit Function
        End If
    Next Lng_行
    Fn_会員行検索 = 0
End Function

' ------------------------------------------------------------
' Fn_単価取得
' 入力: Str_対象年月, Str_会社コード
' 出力: 株価マスタに登録された単価(見つからない場合は0)
' ------------------------------------------------------------
Private Function Fn_単価取得(ByVal Str_対象年月 As String, ByVal Str_会社コード As String) As Double
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_株価マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("年月"))

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Trim(CStr(Wsh.Cells(Lng_行, Dic_列("年月")).Value)) = Str_対象年月 And _
           Trim(CStr(Wsh.Cells(Lng_行, Dic_列("会社コード")).Value)) = Str_会社コード Then
            Fn_単価取得 = CDbl(Wsh.Cells(Lng_行, Dic_列("単価")).Value)
            Exit Function
        End If
    Next Lng_行

    Fn_単価取得 = 0
End Function

' ------------------------------------------------------------
' Pr_月次実績Upsert
' 前提: 同一(年月, MemberNo, 会社コード)が既に存在する場合は上書きする
'       (再実行しても行が増殖しないための冪等性担保)
' 入力: Str_年月, Str_MemberNo, Str_会社コード, Dbl_口数, Dbl_単価
' 出力: 月次実績シートへの追加または更新
' ------------------------------------------------------------
Private Sub Pr_月次実績Upsert(ByVal Str_年月 As String, ByVal Str_MemberNo As String, _
                               ByVal Str_会社コード As String, ByVal Dbl_口数 As Double, ByVal Dbl_単価 As Double)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long
    Dim Lng_対象行 As Long

    Set Wsh = Fn_シート取得(シート_月次実績)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("年月"))
    Lng_対象行 = 0

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Trim(CStr(Wsh.Cells(Lng_行, Dic_列("年月")).Value)) = Str_年月 And _
           Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value)) = Str_MemberNo And _
           Trim(CStr(Wsh.Cells(Lng_行, Dic_列("会社コード")).Value)) = Str_会社コード Then
            Lng_対象行 = Lng_行
            Exit For
        End If
    Next Lng_行

    If Lng_対象行 = 0 Then
        Lng_対象行 = Lng_最終行 + 1
    End If

    Wsh.Cells(Lng_対象行, Dic_列("年月")).Value = Str_年月
    Wsh.Cells(Lng_対象行, Dic_列("MemberNo")).Value = Str_MemberNo
    Wsh.Cells(Lng_対象行, Dic_列("会社コード")).Value = Str_会社コード
    Wsh.Cells(Lng_対象行, Dic_列("口数")).Value = Dbl_口数
    Wsh.Cells(Lng_対象行, Dic_列("単価")).Value = Dbl_単価
    Wsh.Cells(Lng_対象行, Dic_列("評価額")).Value = Dbl_口数 * Dbl_単価
    Wsh.Cells(Lng_対象行, Dic_列("在籍状態")).Value = Fn_会員状態文字列(会員状態_在籍)
    Wsh.Cells(Lng_対象行, Dic_列("更新日時")).Value = Now
End Sub
