Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' 目次としてハイパーリンク付きのシート一覧を作成するマクロ
' 前提条件：目次シートが存在すること
' ===========================================================
Sub 目次作成_シート一覧()

	Const 目次シート名 = "目次"
	Const 出力列No = 1
	
	
	Dim シートオブジェクト as Worksheet
	Dim 行No as Integer
	行No=1
	
	REM クリア
	Sheets(目次シート名).COLUMNs("A:A").Clear
	
	REM ハイパーリンク付き目次一覧を作成
	For Each シートオブジェクト In Sheets
		ActiveSheet.Hyperlinks.Add _
		  Anchor:=Sheets(目次シート名).Cells(行No,出力列No) _
		  , Address:=ThisWorkbook.FullName _
		  , SubAddress:=シートオブジェクト.Name & "!A1" _
		  , TextToDisplay:=シートオブジェクト.Name
		行No=行No+1
	Next
End Sub

