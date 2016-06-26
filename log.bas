' ログファイル
Const lofFile = ActiveWorkbook.Path & "\log.log"

' ログファイル出力
Sub ファイル出力(ByVal str as String)
  If Dir(lofFile) <> "" Then
    Open logFile For Append As #1
      Print #1, str
    Close #1
  Else
    Open logFile For Output As #1
      Print #1, str
    Close #1
  End If
End Sub
