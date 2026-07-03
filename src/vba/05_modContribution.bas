Attribute VB_Name = "modContribution"
Option Explicit

' ============================================================
' 拠出設定更新モジュール
' 前提   : 拠出設定は履歴管理する(変更時は「旧レコード終了→新レコード追加」)
' 入力   : 各Public Pr_/Fn_の引数を参照
' 出力   : 拠出設定シートへの追加・更新(論理的な終了のみ、物理削除は行わない)
' 例外   : 対象の有効レコードが見つからない場合は何もしない
'          (再実行時に二重処理とならないようにするための安全策)
' 拡張余地: 口数以外(拠出単価等)の履歴管理項目を追加する場合は
'           列を追加しPr_拠出設定新規の引数を拡張する
' ============================================================

' ------------------------------------------------------------
' Fn_現在の拠出設定行取得
' 入力: Str_MemberNo
' 出力: 状態=有効 の行番号(見つからない場合は0)
' ------------------------------------------------------------
Private Function Fn_現在の拠出設定行取得(ByVal Str_MemberNo As String) As Long
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_拠出設定)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("設定No"))

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value)) = Str_MemberNo And _
           Trim(CStr(Wsh.Cells(Lng_行, Dic_列("状態")).Value)) = Fn_拠出設定状態文字列(拠出設定_有効) Then
            Fn_現在の拠出設定行取得 = Lng_行
            Exit Function
        End If
    Next Lng_行

    Fn_現在の拠出設定行取得 = 0
End Function

' ------------------------------------------------------------
' Pr_拠出設定新規
' 前提: 対象MemberNoに現在有効な拠出設定がないこと(呼び出し元の業務フローで担保)
' 入力: Str_MemberNo, Str_適用開始年月, Dbl_口数, Lng_イベントNo
' 出力: 拠出設定シート末尾に状態=有効の新規行を追加
' ------------------------------------------------------------
Public Sub Pr_拠出設定新規(ByVal Str_MemberNo As String, ByVal Str_適用開始年月 As String, _
                            ByVal Dbl_口数 As Double, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long
    Dim Lng_設定No As Long

    Set Wsh = Fn_シート取得(シート_拠出設定)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_最終行取得(Wsh, Dic_列("設定No")) + 1
    Lng_設定No = Fn_次の連番取得(Wsh, Dic_列("設定No"))

    Wsh.Cells(Lng_行, Dic_列("設定No")).Value = Lng_設定No
    Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value = Str_MemberNo
    Wsh.Cells(Lng_行, Dic_列("適用開始年月")).Value = Str_適用開始年月
    Wsh.Cells(Lng_行, Dic_列("適用終了年月")).Value = ""
    Wsh.Cells(Lng_行, Dic_列("口数")).Value = Dbl_口数
    Wsh.Cells(Lng_行, Dic_列("状態")).Value = Fn_拠出設定状態文字列(拠出設定_有効)
    Wsh.Cells(Lng_行, Dic_列("登録日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("登録元イベントNo")).Value = Lng_イベントNo
End Sub

' ------------------------------------------------------------
' Pr_拠出設定終了
' 前提: 対象MemberNoの現在有効な拠出設定が存在すること
' 入力: Str_MemberNo, Str_終了年月, Lng_イベントNo
' 出力: 該当行の適用終了年月・状態を更新(行の削除は行わない)
' 例外: 有効な拠出設定が存在しない場合は何もしない(再実行時の重複終了を防止)
' ------------------------------------------------------------
Public Sub Pr_拠出設定終了(ByVal Str_MemberNo As String, ByVal Str_終了年月 As String, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Lng_行 = Fn_現在の拠出設定行取得(Str_MemberNo)
    If Lng_行 = 0 Then Exit Sub

    Set Wsh = Fn_シート取得(シート_拠出設定)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)

    Wsh.Cells(Lng_行, Dic_列("適用終了年月")).Value = Str_終了年月
    Wsh.Cells(Lng_行, Dic_列("状態")).Value = Fn_拠出設定状態文字列(拠出設定_終了)
End Sub

' ------------------------------------------------------------
' Fn_対象年月の拠出設定取得
' 前提: 月次実績更新モジュールから呼び出される
' 入力: Str_対象年月
' 出力: Dictionary(キー:MemberNo, 値:口数Double) - 対象年月時点で
'       適用開始年月 <= 対象年月 かつ (適用終了年月が空 または 適用終了年月 > 対象年月)
'       を満たす拠出設定を持つ会員一覧
' ------------------------------------------------------------
Public Function Fn_対象年月の拠出設定取得(ByVal Str_対象年月 As String) As Object
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Dic_結果 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long
    Dim Str_MemberNo As String, Str_開始 As String, Str_終了 As String

    Set Wsh = Fn_シート取得(シート_拠出設定)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Set Dic_結果 = CreateObject("Scripting.Dictionary")
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("設定No"))

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        Str_MemberNo = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value))
        Str_開始 = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("適用開始年月")).Value))
        Str_終了 = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("適用終了年月")).Value))

        If Fn_文字列有無判定(Str_MemberNo) And Str_開始 <= Str_対象年月 And _
           (Not Fn_文字列有無判定(Str_終了) Or Str_終了 > Str_対象年月) Then
            If Not Dic_結果.Exists(Str_MemberNo) Then
                Dic_結果.Add Str_MemberNo, CDbl(Wsh.Cells(Lng_行, Dic_列("口数")).Value)
            End If
        End If
    Next Lng_行

    Set Fn_対象年月の拠出設定取得 = Dic_結果
End Function
