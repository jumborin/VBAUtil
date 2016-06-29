Option Explicit

Const CSVFileName = "C:\VBAUtil"
Const シート名 = "Sheet1"

Sub 実行()
  Dim buf as String
  Dim ドメイン as New Domain
  Dim 行No as Integer
  Const タイトル行 = 1
  行No = タイトル行
  
  Call ドメイン.出力先シートセット(シート名)
  Call ドメイン.タイトル行を出力(出力ワークシート,タイトル行)
  Open CSVFileName For Input As #1
    Do While Not EOF(5)
      行No = 行No + 1
      Line Input #1, buf
      Call ドメイン.データセット(Split(buf,","))
      Call ドメイン.セルに出力(出力ワークシート,行No)
    Loop
  Close #1
  Call シート設定(ドメイン.フィールド数)
End Sub

Private Sub シート設定(ByVal 列数 as Integer)
  Dim 列No as Integer
  WorkSheets(シート名).Cells(1,1).Value = "課題一覧(" & Date & "" & Time & "時点)"
  For 列No = 1 to 列数
    WorkSheets(シート名).Columns(列No).AutoFit
  Next 列No
End Sub
