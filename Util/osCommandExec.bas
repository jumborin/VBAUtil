' 引数で指定したDOSコマンドを実行する
Sub osCommandExec(ByVal cmd as String)
  Dim ShellObj
  Set ShellObj = CreateObject("WScript.Shell")
  ShellObj.Run cmd, 0, True
End Sub
