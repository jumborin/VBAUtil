Option Explicit

' CSVファイル操作クラス

' Const CSVFileName = "data.csv"
' Const CSVFilePath = ActiveWorkbook.Path & "\"

' ファイル存在チェック
Function isExistFile(ByVal fileName as String) as Boolean
  If Dir(fileName) <> "" then
    isExistFile = True
  Else
    isExistFile = False
  End IF
End Function

' CSV書き出し
Sub writeCSV(ByVal list as collection,ByVal CSVFileFullPath as String)
  Dim outputData as Variant
  
  Open CSVFileFullPath For Output As #1
    For each outputData in list
      Print #1,Join(outputData,",")
      Print #1,vbcrlf
    Next outputData
  Close #1
  
  MsgBox CSVFileFullPath & "にCSVファイルとして書き出しました"
End Sub

' CSVファイルを読み込み、コレクションで返却する。
Function readCSV(ByVal CSVFileName as String) as Collection
  Dim buf as String
  Dim list as New Collection
  Open CSVFileName For Input As #2
    Do While Not EOF(5)
      Line Input #2, buf
      list.add(buf)
    Loop
  Close #2
  readCSV = list
End Function
