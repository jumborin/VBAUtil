Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicrosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' CSVファイルを整形するマクロ
' 前提条件：対象シートがアクティブになること
' ===========================================================
Sub Redmine取り込みマクロ()
  Dim 最終行,最終列 as Integer
  最終行 = Cells(1, 1).End(xlDown).Row
  最終列 = Cells(1, 1).End(xlToRight).Column

  ' 罫線を引く
  Range(Cells(1,1),Cells(最終行,最終列)).Borders.Weight = xlContinuous
  
  ' タイトル行に背景色をつける
  Range(Cells(1,1),Cells(1,最終列)).Interior.ColorIndex = 3

  ' 列幅を調整する
  Columns.EntireColumn.Autofit
  
  ' オートフィルタがついていない場合のみオートフィルタを設定する
  If ActiveSheet.AutoFilterMode = False Then
    Range(Cells(1,1),Cells(最終行,最終列)).AutoFilter
  End If
  
End Sub