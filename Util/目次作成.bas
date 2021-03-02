Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' 目次としてハイパーリンク付きのシート一覧を作成するマクロ
' 前提条件：目次シートが存在すること
' ===========================================================
Sub 目次作成_シート一覧()

	' 定数処理
	Const 目次シート名 = "目次"
	Const 出力列No = 1
	
	' 変数の宣言
	Dim シートオブジェクト as Worksheet
	Dim 行No as Integer
	Dim isSheetFlag as Boolean
	
	' 変数の初期値設定
	行No=1
	isSheetFlag = false
	
	' シートが存在しない場合は、シートを追加
	For Each シートオブジェクト In Sheets
		if シートオブジェクト.Name = 目次シート名 then
			isSheetFlag =true
			' 目次シートをクリア
			Sheets(目次シート名).COLUMNs("A:A").Clear
			Exit for
		end if
	Next
	if isSheetFlag=false then
		ThisWorkbook.Worksheets.Add
		ThisWorkbook.ActiveSheet.Name = 目次シート名
	end if
	
	' ハイパーリンク付き目次一覧を作成
	For Each シートオブジェクト In Sheets
		ActiveSheet.Hyperlinks.Add _
		  Anchor:=Sheets(目次シート名).Cells(行No,出力列No) _
		  , Address:=ThisWorkbook.FullName _
		  , SubAddress:=シートオブジェクト.Name & "!A1" _
		  , TextToDisplay:=シートオブジェクト.Name
		行No=行No+1
	Next
	
	' 列幅を文字列に合わせて調整
    Columns("A:A").EntireColumn.AutoFit
    
End Sub

