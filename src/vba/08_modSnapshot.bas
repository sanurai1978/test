Attribute VB_Name = "modSnapshot"
Option Explicit

' ============================================================
' スナップショット出力モジュール
' 前提   : 指定年月末時点の会社別会員一覧を出力する
'          在籍判定は「入会日<=月末 かつ (退会日が空 または 退会日>月末)」
'          を会員マスタの値でそのまま評価する(仕様書の定義のとおり)
' 入力   : Str_対象年月(省略時は設定シートの「処理対象年月」を使用)
' 出力   : 「スナップショット」シートの出力範囲をクリアして再出力する
'          (シート自体の削除は行わない)
' 例外   : 出力先シートが存在しない場合は操作ログにエラーを記録して終了
' 拡張余地: 会社マスタの会社名を結合して出力しているため、
'           他の会員属性を追加したい場合は出力列を追加するだけでよい
' ------------------------------------------------------------
' 非機能上の制約(重要):
'   会員マスタは「現在状態のみ」を保持する設計のため、本機能は
'   過去時点の会社所属や氏名変更を厳密には遡及できない。
'   過去のイベント履歴と完全に整合したスナップショットが必要な場合は
'   イベント履歴からの状態再構築(拡張余地)を別途検討すること。
' ============================================================

' ------------------------------------------------------------
' Pr_スナップショット出力実行
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
Public Sub Pr_スナップショット出力実行(Optional ByVal Str_対象年月 As String = "")
    On Error GoTo ErrHandler

    Dim Wsh_会員 As Worksheet, Wsh_会社 As Worksheet, Wsh_出力 As Worksheet
    Dim Dic_会員列 As Object, Dic_会社列 As Object, Dic_出力列 As Object
    Dim Dic_会社名 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long
    Dim Dte_基準日 As Date
    Dim Lng_出力行 As Long
    Dim Lng_件数 As Long
    Dim Str_会社コード As String

    ' 1. 環境抑止
    Pr_環境抑止

    ' 2. 入力取得
    If Not Fn_文字列有無判定(Str_対象年月) Then
        Str_対象年月 = Fn_設定値取得("処理対象年月")
    End If

    ' 3. 入力チェック
    If Not Fn_年月文字列判定(Str_対象年月) Then
        Pr_操作ログ出力 "スナップショット出力", 0, ログ_エラー, "対象年月が不正です(" & Str_対象年月 & ")"
        Pr_環境復元
        MsgBox "対象年月が不正です。設定シートの「処理対象年月」を確認してください。", vbExclamation
        Exit Sub
    End If
    Dte_基準日 = Fn_年月末日取得(Str_対象年月)

    ' 4. マスタ取得
    Set Wsh_会員 = Fn_シート取得(シート_会員マスタ)
    Set Dic_会員列 = Fn_ヘッダー列マップ取得(Wsh_会員)
    Set Wsh_会社 = Fn_シート取得(シート_会社マスタ)
    Set Dic_会社列 = Fn_ヘッダー列マップ取得(Wsh_会社)
    Set Dic_会社名 = Fn_会社名辞書取得(Wsh_会社, Dic_会社列)

    Set Wsh_出力 = Fn_シート取得(シート_スナップショット)
    Set Dic_出力列 = Fn_ヘッダー列マップ取得(Wsh_出力)
    Pr_出力範囲クリア Wsh_出力, Dic_出力列

    Lng_最終行 = Fn_最終行取得(Wsh_会員, Dic_会員列("MemberNo"))
    Lng_出力行 = Lng_データ開始行
    Lng_件数 = 0

    ' 5. 業務処理 + 6. 出力(在籍している会員のみ出力)
    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Fn_在籍判定( _
            Fn_日付取得(Wsh_会員.Cells(Lng_行, Dic_会員列("入会日")).Value), _
            Fn_日付取得(Wsh_会員.Cells(Lng_行, Dic_会員列("退会日")).Value), _
            Dte_基準日) Then

            Str_会社コード = Trim(CStr(Wsh_会員.Cells(Lng_行, Dic_会員列("会社コード")).Value))

            Wsh_出力.Cells(Lng_出力行, Dic_出力列("対象年月")).Value = Str_対象年月
            Wsh_出力.Cells(Lng_出力行, Dic_出力列("会社コード")).Value = Str_会社コード
            Wsh_出力.Cells(Lng_出力行, Dic_出力列("会社名")).Value = Fn_辞書値取得(Dic_会社名, Str_会社コード)
            Wsh_出力.Cells(Lng_出力行, Dic_出力列("MemberNo")).Value = _
                Wsh_会員.Cells(Lng_行, Dic_会員列("MemberNo")).Value
            Wsh_出力.Cells(Lng_出力行, Dic_出力列("氏名")).Value = _
                Wsh_会員.Cells(Lng_行, Dic_会員列("氏名")).Value
            Wsh_出力.Cells(Lng_出力行, Dic_出力列("現在口数")).Value = _
                Wsh_会員.Cells(Lng_行, Dic_会員列("現在口数")).Value

            Lng_出力行 = Lng_出力行 + 1
            Lng_件数 = Lng_件数 + 1
        End If
    Next Lng_行

    ' 7. ログ
    Pr_操作ログ出力 "スナップショット出力(" & Str_対象年月 & ")", Lng_件数, ログ_正常, "出力件数:" & Lng_件数 & "件"

    ' 8. 環境復元
    Pr_環境復元
    Exit Sub

ErrHandler:
    Pr_操作ログ出力 "スナップショット出力", 0, ログ_エラー, "実行時エラー:" & Err.Number & " " & Err.Description
    Pr_環境復元
    MsgBox "スナップショット出力でエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' ------------------------------------------------------------
' Fn_会社名辞書取得
' 入力: Wsh(会社マスタ), Dic_列
' 出力: Dictionary(キー:会社コード, 値:会社名)
' ------------------------------------------------------------
Private Function Fn_会社名辞書取得(ByVal Wsh As Worksheet, ByVal Dic_列 As Object) As Object
    Dim Dic_結果 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long
    Dim Str_コード As String

    Set Dic_結果 = CreateObject("Scripting.Dictionary")
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("会社コード"))

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        Str_コード = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("会社コード")).Value))
        If Fn_文字列有無判定(Str_コード) And Not Dic_結果.Exists(Str_コード) Then
            Dic_結果.Add Str_コード, Wsh.Cells(Lng_行, Dic_列("会社名")).Value
        End If
    Next Lng_行

    Set Fn_会社名辞書取得 = Dic_結果
End Function

' ------------------------------------------------------------
' Fn_辞書値取得 / Fn_日付取得(簡易ヘルパー)
' ------------------------------------------------------------
Private Function Fn_辞書値取得(ByVal Dic As Object, ByVal Str_キー As String) As String
    If Dic.Exists(Str_キー) Then
        Fn_辞書値取得 = Dic(Str_キー)
    Else
        Fn_辞書値取得 = ""
    End If
End Function

Private Function Fn_日付取得(ByVal Vnt_値 As Variant) As Date
    If IsDate(Vnt_値) Then
        Fn_日付取得 = CDate(Vnt_値)
    Else
        Fn_日付取得 = 0
    End If
End Function

' ------------------------------------------------------------
' Pr_出力範囲クリア
' 前提: ヘッダー行(1行目)は残し、データ行のみクリアする
' 入力: Wsh(スナップショット), Dic_列
' 出力: なし(データ行のセル内容をクリア)
' ------------------------------------------------------------
Private Sub Pr_出力範囲クリア(ByVal Wsh As Worksheet, ByVal Dic_列 As Object)
    Dim Lng_最終行 As Long
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("MemberNo"))
    If Lng_最終行 >= Lng_データ開始行 Then
        Wsh.Rows(Lng_データ開始行 & ":" & Lng_最終行).ClearContents
    End If
End Sub
