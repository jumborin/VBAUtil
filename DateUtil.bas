Option Explicit

' 8桁の日付を文字列で返す。
Function getNowDateToString() as String
  getNowDateToString = Year(Date) & Month(Date) & Day(Date)
End Function

' 14桁の現在日時を文字列で返す。
Function getNowDateTimeToString() as String
  getNowDateTimeToString = Hour(Time) & Minute(Time) & Second(Time)
End Function

' 本日の曜日を文字列で返却する
Function getWeekToString() as String
  getWeekToString = WeekdayName(Weekday(Date))
End Function
