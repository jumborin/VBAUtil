Option Explicit

' 日付操作を行うユーティリティクラス

' 本日日付を8桁の文字列(yyyymmdd)で返す。
Function getNowDateToString() as String
  getNowDateToString = Year(Date) & Month(Date) & Day(Date)
End Function

' 現在時刻を6桁の文字列(HHmmss)で返す。
Function getNowDateTimeToString() as String
  getNowDateTimeToString = Hour(Time) & Minute(Time) & Second(Time)
End Function

' 本日の曜日を文字列で返却する
Function getWeekToString() as String
  getWeekToString = WeekdayName(Weekday(Date))
End Function
