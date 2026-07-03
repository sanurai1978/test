Attribute VB_Name = "modMemberMaster"
Option Explicit

' ============================================================
' 会員マスタ更新モジュール
' 前提   : 会員マスタは現在状態のみを保持する(履歴はイベント履歴が担う)
' 入力   : 各Public Pr_/Fn_の引数を参照
' 出力   : 会員マスタシートの該当行を追加または更新する(物理削除は行わない)
' 例外   : 対象MemberNoが存在しない更新要求は何もせず終了する
' 拡張余地: 会員属性(部署等)を追加する場合はヘッダー列マップに
'           列を追加し、各Pr_更新系Subに引数を追加するだけでよい
' ============================================================

' ------------------------------------------------------------
' Fn_会員マスタ辞書取得
' 出力: Dictionary(キー:MemberNo, 値:行番号)
' ------------------------------------------------------------
Public Function Fn_会員マスタ辞書取得() As Object
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Dic_会員 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long
    Dim Str_MemberNo As String

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Set Dic_会員 = CreateObject("Scripting.Dictionary")
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("MemberNo"))

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        Str_MemberNo = Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value))
        If Fn_文字列有無判定(Str_MemberNo) And Not Dic_会員.Exists(Str_MemberNo) Then
            Dic_会員.Add Str_MemberNo, Lng_行
        End If
    Next Lng_行

    Set Fn_会員マスタ辞書取得 = Dic_会員
End Function

' ------------------------------------------------------------
' Fn_会員マスタ行検索
' 入力: Str_MemberNo
' 出力: 該当行番号(見つからない場合は0)
' ------------------------------------------------------------
Private Function Fn_会員マスタ行検索(ByVal Str_MemberNo As String) As Long
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_最終行 As Long, Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_最終行 = Fn_最終行取得(Wsh, Dic_列("MemberNo"))

    For Lng_行 = Lng_データ開始行 To Lng_最終行
        If Trim(CStr(Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value)) = Str_MemberNo Then
            Fn_会員マスタ行検索 = Lng_行
            Exit Function
        End If
    Next Lng_行

    Fn_会員マスタ行検索 = 0
End Function

' ------------------------------------------------------------
' Fn_現在口数取得
' 入力: Str_MemberNo
' 出力: 会員マスタに登録されている現在口数(未登録時は0)
' ------------------------------------------------------------
Public Function Fn_現在口数取得(ByVal Str_MemberNo As String) As Double
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_会員マスタ行検索(Str_MemberNo)

    If Lng_行 = 0 Then
        Fn_現在口数取得 = 0
    Else
        Fn_現在口数取得 = CDbl(Wsh.Cells(Lng_行, Dic_列("現在口数")).Value)
    End If
End Function

' ------------------------------------------------------------
' Fn_在籍判定
' 前提: 入会日 <= 基準日 かつ (退会日が空 または 退会日 > 基準日)
' 入力: Dte_入会日, Dte_退会日(未入力は 0 として扱う), Dte_基準日
' 出力: 在籍していればTrue
' ------------------------------------------------------------
Public Function Fn_在籍判定(ByVal Dte_入会日 As Date, ByVal Dte_退会日 As Date, ByVal Dte_基準日 As Date) As Boolean
    Dim Bln_退会日空 As Boolean
    Bln_退会日空 = (Dte_退会日 = 0)

    Fn_在籍判定 = (Dte_入会日 <= Dte_基準日) And (Bln_退会日空 Or Dte_退会日 > Dte_基準日)
End Function

' ------------------------------------------------------------
' Pr_入会登録
' 前提: MemberNoが会員マスタに未登録であること(呼び出し元でチェック済み)
' 入力: Str_MemberNo, Str_氏名, Str_会社コード, Dte_入会日, Dbl_口数, Lng_イベントNo
' 出力: 会員マスタ末尾に新規行を追加
' ------------------------------------------------------------
Public Sub Pr_入会登録(ByVal Str_MemberNo As String, ByVal Str_氏名 As String, _
                        ByVal Str_会社コード As String, ByVal Dte_入会日 As Date, _
                        ByVal Dbl_口数 As Double, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_最終行取得(Wsh, Dic_列("MemberNo")) + 1

    Wsh.Cells(Lng_行, Dic_列("MemberNo")).Value = Str_MemberNo
    Wsh.Cells(Lng_行, Dic_列("氏名")).Value = Str_氏名
    Wsh.Cells(Lng_行, Dic_列("会社コード")).Value = Str_会社コード
    Wsh.Cells(Lng_行, Dic_列("会員状態")).Value = Fn_会員状態文字列(会員状態_在籍)
    Wsh.Cells(Lng_行, Dic_列("入会日")).Value = Dte_入会日
    Wsh.Cells(Lng_行, Dic_列("退会日")).Value = ""
    Wsh.Cells(Lng_行, Dic_列("休会開始日")).Value = ""
    Wsh.Cells(Lng_行, Dic_列("休会終了日")).Value = ""
    Wsh.Cells(Lng_行, Dic_列("現在口数")).Value = Dbl_口数
    Wsh.Cells(Lng_行, Dic_列("最終更新日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("最終更新イベントNo")).Value = Lng_イベントNo
End Sub

' ------------------------------------------------------------
' Pr_退会登録
' 前提: MemberNoが会員マスタに登録済みであること
' 入力: Str_MemberNo, Dte_退会日, Lng_イベントNo
' 出力: 該当行の会員状態・退会日・最終更新情報を更新
' ------------------------------------------------------------
Public Sub Pr_退会登録(ByVal Str_MemberNo As String, ByVal Dte_退会日 As Date, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_会員マスタ行検索(Str_MemberNo)
    If Lng_行 = 0 Then Exit Sub

    Wsh.Cells(Lng_行, Dic_列("会員状態")).Value = Fn_会員状態文字列(会員状態_退会済)
    Wsh.Cells(Lng_行, Dic_列("退会日")).Value = Dte_退会日
    Wsh.Cells(Lng_行, Dic_列("最終更新日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("最終更新イベントNo")).Value = Lng_イベントNo
End Sub

' ------------------------------------------------------------
' Pr_休会登録
' 前提: MemberNoが会員マスタに登録済みであること
' 入力: Str_MemberNo, Dte_休会開始日, Lng_イベントNo
' 出力: 該当行の会員状態・休会開始日・最終更新情報を更新
' ------------------------------------------------------------
Public Sub Pr_休会登録(ByVal Str_MemberNo As String, ByVal Dte_休会開始日 As Date, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_会員マスタ行検索(Str_MemberNo)
    If Lng_行 = 0 Then Exit Sub

    Wsh.Cells(Lng_行, Dic_列("会員状態")).Value = Fn_会員状態文字列(会員状態_休会中)
    Wsh.Cells(Lng_行, Dic_列("休会開始日")).Value = Dte_休会開始日
    Wsh.Cells(Lng_行, Dic_列("休会終了日")).Value = ""
    Wsh.Cells(Lng_行, Dic_列("最終更新日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("最終更新イベントNo")).Value = Lng_イベントNo
End Sub

' ------------------------------------------------------------
' Pr_復帰登録
' 前提: MemberNoが会員マスタに登録済みであること
' 入力: Str_MemberNo, Dte_復帰日, Lng_イベントNo
' 出力: 該当行の会員状態・休会終了日・最終更新情報を更新
' ------------------------------------------------------------
Public Sub Pr_復帰登録(ByVal Str_MemberNo As String, ByVal Dte_復帰日 As Date, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_会員マスタ行検索(Str_MemberNo)
    If Lng_行 = 0 Then Exit Sub

    Wsh.Cells(Lng_行, Dic_列("会員状態")).Value = Fn_会員状態文字列(会員状態_在籍)
    Wsh.Cells(Lng_行, Dic_列("休会終了日")).Value = Dte_復帰日
    Wsh.Cells(Lng_行, Dic_列("最終更新日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("最終更新イベントNo")).Value = Lng_イベントNo
End Sub

' ------------------------------------------------------------
' Pr_口数更新
' 前提: MemberNoが会員マスタに登録済みであること
' 入力: Str_MemberNo, Dbl_新口数, Lng_イベントNo
' 出力: 該当行の現在口数・最終更新情報を更新
' ------------------------------------------------------------
Public Sub Pr_口数更新(ByVal Str_MemberNo As String, ByVal Dbl_新口数 As Double, ByVal Lng_イベントNo As Long)
    Dim Wsh As Worksheet
    Dim Dic_列 As Object
    Dim Lng_行 As Long

    Set Wsh = Fn_シート取得(シート_会員マスタ)
    Set Dic_列 = Fn_ヘッダー列マップ取得(Wsh)
    Lng_行 = Fn_会員マスタ行検索(Str_MemberNo)
    If Lng_行 = 0 Then Exit Sub

    Wsh.Cells(Lng_行, Dic_列("現在口数")).Value = Dbl_新口数
    Wsh.Cells(Lng_行, Dic_列("最終更新日時")).Value = Now
    Wsh.Cells(Lng_行, Dic_列("最終更新イベントNo")).Value = Lng_イベントNo
End Sub
