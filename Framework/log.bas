Option Explicit

' 内容：ログ処理のモジュール

' ログファイル
Const lofFile = ActiveWorkbook.Path & "\log.log"

' ログレベルEnum
Public Enum LogLevel
  DEBUG = 1
  INFO = 2
  WARN = 3
  ERROR = 4
  CRIT = 5
  ALERT = 6
  EMERG = 7
End Enum

' Enumから文字列を返す。
Private Sub getEnumToString(ByVal logLevel as Integer)
  Select Case logLevel
  Case logLevel.DEBUG
    getEnumToString = "DEBUG"
  Case logLevel.INFO
    getEnumToString = "INFO"
  Case logLevel.WARN
    getEnumToString = "WARN"
  Case logLevel.ERROR
    getEnumToString = "ERROR"
  Case logLevel.CRIT
    getEnumToString = "CRIT"
  Case logLevel.ALERT
    getEnumToString = "ALERT"
  Case logLevel.EMERG
    getEnumToString = "EMERG"
  End Select
End Sub


' ログファイル出力(出力形式：現在時刻_ログレベル_メッセージ)
Sub ファイル出力(ByVal logLevel as LogLevel,ByVal message as String)
  If Dir(lofFile) <> "" Then
    Open logFile For Append As #1
      Print #1, Now & "_" & getEnumToString(logLevel) & "_" & message
    Close #1
  Else
    Open logFile For Output As #1
      Print #1, Now & "_" & getEnumToString(logLevel) & "_" & message
    Close #1
  End If
End Sub


