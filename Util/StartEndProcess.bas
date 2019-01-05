Option Explicit

' 開始時刻
Dim startTime as Single

' 終了時刻
Dim endTime As Single

' 処理時間
Dim processTime As Long

' 開始時に実行することで、開始時刻をセットし、画面表示を非表示にする。
Sub startProcess()
  startTime = Now
  Application.ScreenUpdating = False
End Sub

' 終了時に実行することで、終了時刻をデバッグログに出力し、終了メッセージをポップアップ表示する。
Sub endProcess()
  Application.ScreenUpdating = True
  endTime = Now
  processTime = endTime - startTime
  Debug.Print("処理時間は" & Format(processTime, "hh時間nn分ss秒") & "でした")
  MsgBox("処理が完了しました。")
End Sub
