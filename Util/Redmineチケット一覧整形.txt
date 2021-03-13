Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicrosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' 罫線を引くマクロ
' 前提条件：issueがあるシートがアクティブになること
' ===========================================================
Sub Redmine取り込みマクロ()
  Dim 最終行,最終列 as Integer
  最終行 = Cells(1, 1).End(xlDown).Row
  最終列 = Cells(1, 1).End(xlToRight).Column

  ' 罫線を引く
  Range(Cells(1,1),Cells(最終行,最終列)).Borders.Weight = xlContinuous

End Sub