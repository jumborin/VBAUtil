' 右端シート追加
Sub sheetAdd(ByVal fileName as String)
  Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = fileName
End Sub

' 指定シート削除
Sub sheetDelete(ByVal sheetName as String)
    Application.DisplayAlerts = False
    Workbooks(sheetName).Delete
    Application.DisplayAlerts = True
End Sub

' シート名変更
Sub sheetChange(ByVal src as String, ByVal dist as String)
  Workbooks(src).Name = dist
End Sub

' シート有無を検索し、結果を返す(有：True,無：False)
Function isExistSheet(ByVal sheetName as String) as Boolean
  Dim workbook As Workbook
  For Each workbook In Workbooks
    If workbook.Name = sheetName Then
      isExistSheet = True
    End If
  Next workbook
  isExistSheet = False
End Function

' シートの一覧を作成し、コレクションに入れて返す
Function getSheetList() as Collection
  Dim list as New Collection
  Dim sheetNo as Integer
  For sheetNo = 1 To Sheets.Count
　　list.Add(Sheets(sheetNo).Name)
　Next sheetNo
　getSheetList = list
End Function
