Option Explicit

' プリント設定を行う。
Sub printConfig(ByVal sheetName as String, ByVal titleRowNo as Integer)
  Worksheets(sheetName).PageSetup.Zoom = False
  Worksheets(sheetName).PageSetup.FitToPagesTall = 1
  Worksheets(sheetName).PageSetup.FitToPagesWide = 1
  Worksheets(sheetName).View = xlPageBreakPreview
  Worksheets(sheetName).PageSetup.PrintTitleRows = "$1:$" & titleRowNo
End Sub
