Option Explicit


Dim startTime as Single
Dim endTime As Single
Dim processTime As Long

' 開始時に実行する
Sub startProcess()
  startTime = Now
  Application.ScreenUpdating = False
End Sub

' 終了時に実行する
Sub endProcess()
  Application.ScreenUpdating = True
  endTime = Now
  processTime = endTime - startTime
  Debug.Print("処理時間は" & Format(processTime, "hh時間nn分ss秒") & "でした")
  MsgBox("処理が完了しました。")
End Sub
