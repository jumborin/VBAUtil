Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicrosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' 次に開いた人が読みやすいように全シートでA1を選択した状態にするマクロ
' ===========================================================
Sub A1Select()
  
	' 変数の宣言
	Dim シートオブジェクト as Worksheet
	
	For Each シートオブジェクト In Sheets
	  シートオブジェクト.Cells(1,1).Select
	Next
End Sub
