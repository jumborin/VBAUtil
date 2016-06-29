Option Explicit

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

' 最終行取得
Function getLastRow(ByVal sheetName as String) as Integer
  Dim columnNo as Integer
  Dim maxRowNo as Long
  maxRowNo = 1
  For columnNo = 1 to 30
    If maxRowNo < Workbooks(sheetName).Cells(Rows.Count, 1).End(xlUp).Row then
      maxRowNo = Workbooks(sheetName).Cells(Rows.Count, 1).End(xlUp).Row
    End If
  Next columnNo
  getLastRow = maxRowNo
End Function

' 最終列取得
Function getLastColumn() as Integer
  Dim rowNo as Integer
  Dim maxColumnNo as Long
  maxColumnNo = 1
  
  For rowNo = 1 to 30
    If maxColumnNo < Workbooks(sheetName).Cells(Rows.Count, 1).End(xlUp).Row then
      maxColumnNo = Workbooks(sheetName).Cells(Rows.Count, 1).End(xlUp).Row
    End If
  Next rowNo
  
  getLastColumn = maxColumnNo
End Function

