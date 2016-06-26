Option Explicit

' プリント設定を行う。
Sub printConfig(ByVal sheetName as String, ByVal titleRowNo as Integer)
  ' ページ数設定
  Worksheets(sheetName).PageSetup.Zoom = False
  Worksheets(sheetName).PageSetup.FitToPagesTall = 1
  Worksheets(sheetName).PageSetup.FitToPagesWide = 1
  ' 改ページ表示
  Worksheets(sheetName).View = xlPageBreakPreview
  ' タイトル行設定
  Worksheets(sheetName).PageSetup.PrintTitleRows = "$1:$" & titleRowNo
  ' ヘッダーフッター設定
  WorkSheets(sheetName).PageSetup.LeftHeader = ""
  WorkSheets(sheetName).PageSetup.CenterHeader = ""
  WorkSheets(sheetName).PageSetup.RightHeader = ""
  WorkSheets(sheetName).PageSetup.LeftFooter = "印刷日：&D"
  WorkSheets(sheetName).PageSetup.CenterFooter = "&P/&N"
  WorkSheets(sheetName).PageSetup.RightFooter = ""
End Sub
