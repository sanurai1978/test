Attribute VB_Name = "modEventHistory"
Option Explicit

' ============================================================
' イベント履歴更新モジュール
' 前提   : イベント履歴は追記専用(物理削除・上書き禁止)
' 入力   : Fn_履歴追加 の引数を参照
' 出力   : イベント履歴シート末尾への1行追加
' 例外   : なし(呼び出し元でOn Errorにより実行時エラーを捕捉)
' 拡張余地: 反映結果の詳細(処理前後の値など)を残したい場合は
'           列を追加しWshへの書込を追加するだけでよい
' ============================================================

' ------------------------------------------------------------
' Fn_履歴追加
' 前提: 全イベントは削除・上書きせず必ず新規行として追加する
' 入力: Str_MemberNo, Lng_種別(Enm_イベント種別), Dte_発生日, Dbl_口数,
'       Str_備考, Lng_入力行番号
' 出力: 採番されたイベントNo
' ------------------------------------------------------------
Public Function Fn_履歴追加(ByVal Str_MemberNo As String, ByVal Lng_種別 As Long, _
                             ByVal Dte_発生日 As Date, ByVal Dbl_口数 As Double, _
                             ByVal Str_備考 As String, ByVal Lng_入力行番号 As Long) As Long
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long
    Dim Lng_イベントNo As Long

    Set Wsh = Fn_シート取得(シート_イベント履歴)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_最終行取得(Wsh, Dic_列("イベントNo")) + 1
    Lng_イベントNo = Fn_次の連番取得(Wsh, Dic_列("イベントNo"))

    Wsh.Cells(Lng_行, Dic_列("イベントNo")).Value = Lng_イベントNo
    Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value = Str_MemberNo
    Wsh.Cells(Lng_行, Dic_列("イベント種別")).Value = Fn_イベント種別文字列(Lng_種別)
    Wsh.Cells(Lng_行, Dic_列("イベント発生日")).Value = Dte_発生日
    Wsh.Cells(Lng_行, Dic_列("口数")).Value = Dbl_口数
    Wsh.Cells(Lng_行, Dic_列("反映日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("処理状態")).Value = Fn_処理状態文字列(処理状態_処理済)
    Wsh.Cells(Lng_行, Dic_列("備考")).Value = Str_備考
    Wsh.Cells(Lng_行, Dic_列("入力元行番号")).Value = Lng_入力行番号

    Fn_履歴追加 = Lng_イベントNo
End Function
