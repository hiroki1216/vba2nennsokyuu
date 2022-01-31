Attribute VB_Name = "Module1"
'グローバル変数の定義

Public optbtnOkResult As Boolean 'オプションボタンの真偽値(はい)
Public optbtnNoResult As Boolean 'オプションボタンの真偽値(いいえ 同世帯加入者なし)
Public optbtnNotResult As Boolean 'オプションボタンの真偽値(いいえ 基準日時点加入者あり)

Public ConvertToday As String '今日の日付(文字列)
Public today As Date '今日の日付
Public thisYear As Long '現年度
Public objectYear As Long '遡及年度
Public convertToWareki As Long '和暦変換用

Public goBackAbleDate As Date '遡及可能年月日
Public goBackAbleDateComparison As Date '遡及可能年月日(比較用)

Public firstDeadline As Date '遡及年度の第一期納付期限

Public initDate As Date '遡及年度の開始日
Public finDate As Date '遡及年度の終了日


'本日の年月日から年度を求める関数
Function fiscalYear(ByVal today As Date) As Long
    If Month(today) >= 4 Then
        fiscalYear = Year(today)
    Else
        fiscalYear = Year(today) - 1
    End If
End Function



