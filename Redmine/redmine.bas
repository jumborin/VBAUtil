Option Explicit

Const CSVFileName = "C:\VBAUtil"
Const シート名 = "Sheet1"

Sub 実行()
  Dim buf as String
  Dim ドメイン as New Domain
  Dim 行No as Integer
  
  行No = 2
  Call ドメイン.出力先シートセット(シート名)
  Call ドメイン.タイトル行を出力(出力ワークシート,1)
  Open CSVFileName For Input As #1
    Do While Not EOF(5)
      Line Input #1, buf
      Call ドメイン.データセット(Split(buf,","))
      Call ドメイン.セルに出力(出力ワークシート,行No)
      行No = 行No + 1
    Loop
  Close #1
End Sub
