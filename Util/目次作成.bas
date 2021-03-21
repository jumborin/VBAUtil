Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicrosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' 目次としてハイパーリンク付きのシート一覧を作成するマクロ
' 前提条件：目次シートが存在すること
' ===========================================================
Sub 目次作成_シート一覧()

    ' 定数処理
    Const 目次シート名 = "目次"
    Const 出力列No = 1
    Const 出力列 = "A:A"
    
    ' 変数の宣言
    Dim シートオブジェクト As Worksheet
    Dim 行No As Integer
    Dim isSheetFlag As Boolean
    
    ' 変数の初期値設定
    行No = 1
    isSheetFlag = False
    
    ' シートが存在しない場合は、シートを追加
    For Each シートオブジェクト In Sheets
        If シートオブジェクト.Name = 目次シート名 Then
            isSheetFlag = True
            ' 目次シートをクリア
            Sheets(目次シート名).Columns(出力列).Clear
            Exit For
        End If
    Next
    If isSheetFlag = False Then
        ThisWorkbook.Worksheets.Add
        ThisWorkbook.ActiveSheet.Name = 目次シート名
    End If
    
    ' ハイパーリンク付き目次一覧を作成
    For Each シートオブジェクト In Sheets
        ' ハイパーリンク作成
        ActiveSheet.Hyperlinks.Add _
          Anchor:=Sheets(目次シート名).Cells(行No, 出力列No) _
          , Address:=ThisWorkbook.FullName _
          , SubAddress:=シートオブジェクト.Name & "!A1" _
          , TextToDisplay:=シートオブジェクト.Name
          
        ' 罫線を引く
        Sheets(目次シート名).Cells(行No, 出力列No).Borders.Weight = xlContinuous
        行No = 行No + 1
    Next
    
    ' 列幅を文字列に合わせて調整
    Columns(出力列).EntireColumn.AutoFit
    
End Sub
