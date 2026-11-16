Attribute VB_Name = "Mod_採番"
Option Explicit
'==============================================================
' Mod_採番
' 【目的】伝票番号と手配No.を振る。
' 【業務背景】HTML 版と同じ番号の振り方にする。
'   ・伝票番号（SO・MO・PO・SH・SC・PK）は事業部ごとの連番で「SO-0001」の形。
'     両事業部に同じ番号があり、事業部＋番号で見分ける。
'   ・見込ID（FC）は事業部ごとの連番の数字（1, 2, …）。
'   ・手配No.は製造指図の工程1つごとに振る全社連番（7桁、1000001～）。
'     事業部をまたいで重複しないので、手配No.だけで工程が特定できる
'     （製造指図書のバーコードを読んで実績を入れる使い方のため）。
' 【前提】採番シートに「事業部｜種類｜最終番号」の行がある（初期化で作る）。
'         行が無い種類は 0 から始める（行を足す）。
' 【流れ】最終番号を読む → 1足して書き戻す → 形を整えて返す
'==============================================================

Private Const St_全社 As String = "全社"          ' 手配No.の行の事業部欄
Private Const St_種類_手配 As String = "手配No"   ' 手配No.の行の種類欄

' 伝票番号を1つ振る（例：Fn_採番("加工", "SO") → "SO-0005"）
'   引数：St_事業部（加工／組立）、St_種類（SO・MO・PO・SH・SC・PK）
'   戻り値：「種類-4桁」の番号
Public Function Fn_採番(ByVal St_事業部 As String, ByVal St_種類 As String) As String
    Dim Lg_番号 As Long
    Lg_番号 = Fn_次の番号(St_事業部, St_種類)
    Fn_採番 = St_種類 & "-" & Format$(Lg_番号, "0000")
End Function

' 見込IDを1つ振る（事業部ごとの連番の数字）
Public Function Fn_見込ID採番(ByVal St_事業部 As String) As Long
    Fn_見込ID採番 = Fn_次の番号(St_事業部, "FC")
End Function

' 手配No.を1つ振る（全社連番。7桁の文字で返す）
Public Function Fn_手配No採番() As String
    Fn_手配No採番 = CStr(Fn_次の番号(St_全社, St_種類_手配))
End Function

' 採番シートの最終番号に1足して書き戻し、その番号を返す
'   引数：St_事業部、St_種類　戻り値：新しい番号
Private Function Fn_次の番号(ByVal St_事業部 As String, ByVal St_種類 As String) As Long
    Dim Ws_採番 As Worksheet
    Dim Va_表 As Variant
    Dim Lg_行 As Long
    Dim Lg_番号 As Long
    Set Ws_採番 = ThisWorkbook.Worksheets(St_シート_採番)
    Va_表 = Fn_表読込(Ws_採番, 列_採番.次列 - 1)
    ' 該当する行を探して1足す
    For Lg_行 = 1 To Fn_行数(Va_表)
        If Fn_文字(Va_表(Lg_行, 列_採番.事業部)) = St_事業部 And Fn_文字(Va_表(Lg_行, 列_採番.種類)) = St_種類 Then
            Lg_番号 = CLng(Fn_数(Va_表(Lg_行, 列_採番.最終番号))) + 1
            Ws_採番.Cells(Lg_行 + Lg_開始行 - 1, 列_採番.最終番号).Value = Lg_番号
            Fn_次の番号 = Lg_番号
            Exit Function
        End If
    Next Lg_行
    ' 行が無ければ足す（手配No.は 1000001 から、それ以外は 1 から）
    If St_種類 = St_種類_手配 Then
        Lg_番号 = 1000001
    Else
        Lg_番号 = 1
    End If
    Call Fn_行追記(Ws_採番, Array(St_事業部, St_種類, Lg_番号))
    Fn_次の番号 = Lg_番号
End Function
