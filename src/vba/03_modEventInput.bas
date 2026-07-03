Attribute VB_Name = "modEventInput"
Option Explicit

' ============================================================
' イベント入力反映モジュール
' 前提   : イベント入力シートが唯一のユーザー編集シートである
' 入力   : イベント入力シート(処理状態が「未処理」または空欄の行)
' 出力   : 会員マスタ/拠出設定/イベント履歴を更新し
'          イベント入力シートの処理状態列を更新する
' 例外   : 行単位の業務エラーは処理状態=エラーとして記録し
'          後続行の処理は継続する。実行時エラーは操作ログに記録し中断する
' 拡張余地: イベント種別の追加は modConstants.Enm_イベント種別 と
'           本モジュールの Pr_業務処理 の Select Case に追加するだけでよい
' ============================================================

' ------------------------------------------------------------
' Pr_イベント入力反映実行
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
Public Sub Pr_イベント入力反映実行()
    On Error GoTo ErrHandler

    Dim Wsh_入力 As Worksheet
    Dim Dic_列 As Object
    Dim Lng_最終行 As Long
    Dim Lng_行 As Long
    Dim Typ_行 As Typ_イベント入力行
    Dim Dic_会員 As Object
    Dim Typ_結果 As Typ_処理結果

    ' 1. 環境抑止
    Pr_環境抑止

    ' 2. 入力取得
    Set Wsh_入力 = Fn_シート取得(シート_イベント入力)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh_入力)
    Lng_最終行 = Fn_最終行取得(Wsh_入力, Dic_列("MemberNo"))

    ' 4. マスタ取得(会員存在チェック用に先読み)
    Set Dic_会員 = Fn_会員マスタ辞書取得()

    Typ_結果.Lng_成功件数 = 0
    Typ_結果.Lng_エラー件数 = 0

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Fn_処理対象行判定(Wsh_入力, Dic_列, Lng_行) Then

            ' 3. 入力チェック
            Typ_行 = Fn_入力行チェック(Wsh_入力, Dic_列, Lng_行, Dic_会員)

            If Typ_行.Bln_エラーあり Then
                ' 6. 出力(エラー)
                Pr_入力行結果書込 Wsh_入力, Dic_列, Lng_行, 処理状態_エラー, Typ_行.Str_エラー内容
                Typ_結果.Lng_エラー件数 = Typ_結果.Lng_エラー件数 + 1
            Else
                ' 5. 業務処理
                If Pr_業務処理(Typ_行) Then
                    ' 会員マスタ辞書を最新化(入会等でMemberNoが増える場合に対応)
                    If Not Dic_会員.Exists(Typ_行.Str_MemberNo) Then
                        Dic_会員.Add Typ_行.Str_MemberNo, True
                    End If
                    ' 6. 出力(正常)
                    Pr_入力行結果書込 Wsh_入力, Dic_列, Lng_行, 処理状態_処理済, ""
                    Typ_結果.Lng_成功件数 = Typ_結果.Lng_成功件数 + 1
                Else
                    Pr_入力行結果書込 Wsh_入力, Dic_列, Lng_行, 処理状態_エラー, "業務処理でエラーが発生しました"
                    Typ_結果.Lng_エラー件数 = Typ_結果.Lng_エラー件数 + 1
                End If
            End If
        End If
    Next Lng_行

    ' 7. ログ
    Pr_操作ログ出力 "イベント入力反映", Typ_結果.Lng_成功件数 + Typ_結果.Lng_エラー件数, _
        IIf(Typ_結果.Lng_エラー件数 > 0, ログ_エラー, ログ_正常), _
        "成功:" & Typ_結果.Lng_成功件数 & "件 エラー:" & Typ_結果.Lng_エラー件数 & "件"

    ' 8. 環境復元
    Pr_環境復元
    Exit Sub

ErrHandler:
    Pr_操作ログ出力 "イベント入力反映", 0, ログ_エラー, "実行時エラー:" & Err.Number & " " & Err.Description
    Pr_環境復元
    MsgBox "イベント入力反映でエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' ------------------------------------------------------------
' Fn_処理対象行判定
' 前提: 処理状態が「処理済」の行は再実行時にスキップする(再実行可能の担保)
' 入力: Wsh, Dic_列, Lng_行
' 出力: MemberNoが入力済みかつ未処理の場合True、処理済の場合False
' ------------------------------------------------------------
Private Function Fn_処理対象行判定(ByVal Wsh As Worksheet, ByVal Dic_列 As Object, ByVal Lng_行 As Long) As Boolean
    Dim Str_MemberNo As String
    Dim Str_状態 As String

    Str_MemberNo = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value))
    Str_状態 = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("処理状態")).Value))

    If Not Fn_文字列有無判定(Str_MemberNo) Then
        Fn_処理対象行判定 = False
    ElseIf Str_状態 = Fn_処理状態文字列(処理状態_処理済) Then
        Fn_処理対象行判定 = False
    Else
        Fn_処理対象行判定 = True
    End If
End Function

' ------------------------------------------------------------
' Fn_入力行チェック
' 前提: 会員マスタ辞書(Dic_会員)が最新化されていること
' 入力: Wsh, Dic_列, Lng_行, Dic_会員
' 出力: Typ_イベント入力行(検証結果とエラー内容を格納)
' 例外: 必須項目未入力、イベント種別不正、既存/新規の整合性不正、
'       日付不正、口数不正 をそれぞれ判定しエラー内容に列挙する
' 拡張余地: チェック項目を追加する場合はCol_エラーへAddするだけでよい
' ------------------------------------------------------------
Private Function Fn_入力行チェック(ByVal Wsh As Worksheet, ByVal Dic_列 As Object, _
                                    ByVal Lng_行 As Long, ByVal Dic_会員 As Object) As Typ_イベント入力行
    Dim Typ_行 As Typ_イベント入力行
    Dim Str_種別文字列 As String
    Dim Col_エラー As New Collection
    Dim Vnt_項目 As Variant
    Dim Str_結合 As String

    Typ_行.Lng_行番号 = Lng_行
    Typ_行.Str_MemberNo = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value))
    Typ_行.Str_氏名 = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("氏名")).Value))
    Typ_行.Str_会社コード = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("会社コード")).Value))
    Str_種別文字列 = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("イベント種別")).Value))
    Typ_行.Lng_種別 = Fn_イベント種別変換(Str_種別文字列)

    If IsDate(Wsh.Cells(Lng_行, Dic_列("イベント発生日")).Value) Then
        Typ_行.Dte_発生日 = CDate(Wsh.Cells(Lng_行, Dic_列("イベント発生日")).Value)
    Else
        Col_エラー.Add "イベント発生日が未入力または不正です"
    End If

    If IsNumeric(Wsh.Cells(Lng_行, Dic_列("口数")).Value) And _
       Fn_文字列有無判定(CStr(Wsh.Cells(Lng_行, Dic_列("口数")).Value)) Then
        Typ_行.Dbl_口数 = CDbl(Wsh.Cells(Lng_行, Dic_列("口数")).Value)
        Typ_行.Bln_口数指定あり = True
    Else
        Typ_行.Bln_口数指定あり = False
    End If

    Typ_行.Bln_新規会員 = Not Dic_会員.Exists(Typ_行.Str_MemberNo)

    ' イベント種別チェック
    If Typ_行.Lng_種別 = 0 Then
        Col_エラー.Add "イベント種別が不正です(" & Str_種別文字列 & ")"
    End If

    ' 会社コードチェック
    If Not Fn_文字列有無判定(Typ_行.Str_会社コード) Then
        Col_エラー.Add "会社コードが未入力です"
    End If

    ' 種別ごとの整合性チェック
    Select Case Typ_行.Lng_種別
        Case イベント_入会
            If Not Typ_行.Bln_新規会員 Then
                Col_エラー.Add "入会イベントですが既に登録済みのMemberNoです"
            End If
            If Not Fn_文字列有無判定(Typ_行.Str_氏名) Then
                Col_エラー.Add "氏名が未入力です"
            End If
            If Not Typ_行.Bln_口数指定あり Then
                Col_エラー.Add "入会時の口数が未入力です"
            ElseIf Typ_行.Dbl_口数 <= 0 Then
                Col_エラー.Add "入会時の口数は正の数で入力してください"
            End If
        Case イベント_退会, イベント_休会, イベント_復帰
            If Typ_行.Bln_新規会員 Then
                Col_エラー.Add "未登録のMemberNoです"
            End If
        Case イベント_口数変更
            If Typ_行.Bln_新規会員 Then
                Col_エラー.Add "未登録のMemberNoです"
            End If
            If Not Typ_行.Bln_口数指定あり Then
                Col_エラー.Add "変更後の口数が未入力です"
            ElseIf Typ_行.Dbl_口数 <= 0 Then
                Col_エラー.Add "変更後の口数は正の数で入力してください"
            End If
    End Select

    Typ_行.Bln_エラーあり = (Col_エラー.Count > 0)
    If Typ_行.Bln_エラーあり Then
        Str_結合 = ""
        For Each Vnt_項目 In Col_エラー
            Str_結合 = Str_結合 & Vnt_項目 & "/ "
        Next Vnt_項目
        Typ_行.Str_エラー内容 = Left(Str_結合, Len(Str_結合) - 2)
    End If

    Fn_入力行チェック = Typ_行
End Function

' ------------------------------------------------------------
' Pr_業務処理
' 前提: Typ_行 は入力チェック済み(エラーなし)であること
' 入力: Typ_行(ByRef)
' 出力: True=正常終了、False=業務処理中にエラー発生
' 例外: On Errorで捕捉しFalseを返す(呼び出し元でエラー行として記録)
' 拡張余地: イベント種別追加時はCase文を追加する
' ------------------------------------------------------------
Private Function Pr_業務処理(ByRef Typ_行 As Typ_イベント入力行) As Boolean
    On Error GoTo ErrHandler
    Dim Lng_イベントNo As Long
    Dim Str_年月 As String

    Str_年月 = Format(Typ_行.Dte_発生日, "yyyymm")

    Select Case Typ_行.Lng_種別
        Case イベント_入会
            Lng_イベントNo = Fn_履歴追加(Typ_行.Str_MemberNo, イベント_入会, Typ_行.Dte_発生日, _
                Typ_行.Dbl_口数, "", Typ_行.Lng_行番号)
            Pr_入会登録 Typ_行.Str_MemberNo, Typ_行.Str_氏名, Typ_行.Str_会社コード, _
                Typ_行.Dte_発生日, Typ_行.Dbl_口数, Lng_イベントNo
            Pr_拠出設定新規 Typ_行.Str_MemberNo, Str_年月, Typ_行.Dbl_口数, Lng_イベントNo

        Case イベント_退会
            Lng_イベントNo = Fn_履歴追加(Typ_行.Str_MemberNo, イベント_退会, Typ_行.Dte_発生日, _
                0, "", Typ_行.Lng_行番号)
            Pr_退会登録 Typ_行.Str_MemberNo, Typ_行.Dte_発生日, Lng_イベントNo
            Pr_拠出設定終了 Typ_行.Str_MemberNo, Str_年月, Lng_イベントNo

        Case イベント_休会
            Lng_イベントNo = Fn_履歴追加(Typ_行.Str_MemberNo, イベント_休会, Typ_行.Dte_発生日, _
                0, "", Typ_行.Lng_行番号)
            Pr_休会登録 Typ_行.Str_MemberNo, Typ_行.Dte_発生日, Lng_イベントNo
            Pr_拠出設定終了 Typ_行.Str_MemberNo, Str_年月, Lng_イベントNo

        Case イベント_復帰
            Lng_イベントNo = Fn_履歴追加(Typ_行.Str_MemberNo, イベント_復帰, Typ_行.Dte_発生日, _
                Typ_行.Dbl_口数, "", Typ_行.Lng_行番号)
            Pr_復帰登録 Typ_行.Str_MemberNo, Typ_行.Dte_発生日, Lng_イベントNo
            Pr_拠出設定新規 Typ_行.Str_MemberNo, Str_年月, Fn_現在口数取得(Typ_行.Str_MemberNo), Lng_イベントNo

        Case イベント_口数変更
            Lng_イベントNo = Fn_履歴追加(Typ_行.Str_MemberNo, イベント_口数変更, Typ_行.Dte_発生日, _
                Typ_行.Dbl_口数, "", Typ_行.Lng_行番号)
            Pr_口数更新 Typ_行.Str_MemberNo, Typ_行.Dbl_口数, Lng_イベントNo
            Pr_拠出設定終了 Typ_行.Str_MemberNo, Str_年月, Lng_イベントNo
            Pr_拠出設定新規 Typ_行.Str_MemberNo, Str_年月, Typ_行.Dbl_口数, Lng_イベントNo
    End Select

    Pr_業務処理 = True
    Exit Function

ErrHandler:
    Pr_業務処理 = False
End Function

' ------------------------------------------------------------
' Pr_入力行結果書込
' 入力: Wsh, Dic_列, Lng_行, Enm_状態, Str_エラー内容
' 出力: イベント入力シートの処理状態/処理日時/エラー内容を更新
' ------------------------------------------------------------
Private Sub Pr_入力行結果書込(ByVal Wsh As Worksheet, ByVal Dic_列 As Object, ByVal Lng_行 As Long, _
                               ByVal Enm_状態 As Enm_処理状態, ByVal Str_エラー内容 As String)
    Wsh.Cells(Lng_行, Dic_列("処理状態")).Value = Fn_処理状態文字列(Enm_状態)
    Wsh.Cells(Lng_行, Dic_列("処理日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("エラー内容")).Value = Str_エラー内容
End Sub
