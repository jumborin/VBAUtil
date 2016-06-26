Option Explicit


Dim startTime as Single
Dim endTime As Single
Dim processTime As Long

' 開始時に実行する
Sub startProcess()
  startTime = Timer
  Application.ScreenUpdating = False
End Sub

' 終了時に実行する
Sub endProcess()
  Application.ScreenUpdating = True
  MsgBox("処理が完了しました。")
  endTime = Timer
  processTime = endTime - startTime
  Debug.Print("処理時間は" & processTime & "秒でした")
End Sub
