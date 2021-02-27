Option Explicit

' CSVファイルを操作するためのユーティリティ処理をまとめたクラス

' Const CSVFileName = "data.csv"
' Const CSVFilePath = ActiveWorkbook.Path & "\"

' ファイル存在チェック処理
' 引数で渡したファイルが存在しているかをチェックし、結果をBooleanで返却する
' True：引数で渡したファイルが存在している
' False：引数で渡したファイルが存在しない
Function isExistFile(ByVal fileName as String) as Boolean
  If Dir(fileName) <> "" then
    isExistFile = True
  Else
    isExistFile = False
  End IF
End Function

' 変数で渡したコレクションの中身を引数で渡したパスのCSVファイルに書き出す処理
Sub writeCSV(ByVal list as Collection,ByVal CSVFileFullPath as String)
  Dim outputData as Variant
  
  Open CSVFileFullPath For Output As #1
    For each outputData in list
      Print #1,Join(outputData,",")
    Next outputData
  Close #1
  
  MsgBox CSVFileFullPath & "にCSVファイルとして書き出しました"
End Sub

' 引数で渡したCSVファイルを読み込み、コレクションで返却する。
Function readCSV(ByVal CSVFileName as String) as Collection
  Dim buf as String
  Dim list as New Collection
  Open CSVFileName For Input As #2
    Do While Not EOF(5)
      Line Input #2, buf
      list.add(buf)
    Loop
  Close #2
  Set readCSV = list
End Function
