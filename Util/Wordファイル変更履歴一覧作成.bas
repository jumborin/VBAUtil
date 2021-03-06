Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicosoftOfficeで実施する場合、削除すること。

Option Explicit

Const 変更履歴シート名 = "変更履歴"

' 
' Wordファイルの更新履歴をExcelで一覧化する処理
' 
Sub 実行()

	' 表示処理を省略
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 変数の宣言
    Dim ファイル名 As String
    Dim ワードファイル As Word.Document
    
    ' 定数の宣言
    Const 検索Path = "C:\VirtualE"
    
    ' Excel内に変更履歴シートを作成
    ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = 変更履歴シート名
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 1).Value = "No"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 2).Value = "作成日"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 3).Value = "作成者"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 4).Value = "ファイル名"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 5).Value = "ページ"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 6).Value = "行"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 7).Value = "変更種別"
    ThisWorkbook.Worksheets(変更履歴シート名).Cells(1, 8).Value = "変更内容"

	' ワードファイル処理のための準備
    Dim ワードアプリ As New Word.Application
    ワードアプリ.DisplayAlerts = wdAlertsNone
    ワードアプリ.Visible = True
    
    ' 1ファイルずつ処理
    ファイル名 = Dir(検索Path & "\*.doc*")
    Do While ファイル名 <> ""
        Set ワードファイル = ワードアプリ.Documents.Open(検索Path & "\" & ファイル名)
        Call 変更履歴データ出力(ワードファイル)
        ワードファイル.Close
        ファイル名 = Dir()
    Loop
    
    ' 列幅を調整
    Columns("A:Z").EntireColumn.AutoFit
    
    ' 後処理
    ワードアプリ.Quit
    Set ワードアプリ = Nothing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub 変更履歴データ出力(ByVal ワードファイル As Word.Document)
    Dim 行,コメント番号 As Long
    
    On Error Resume Next
    
    行 = ThisWorkbook.Worksheets(変更履歴シート名).Cells(1048576, 1).End(xlUp).Row + 1
    For コメント番号 = 1 To ワードファイル.Revisions.Count
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 1).Value = "=ROW()-1" 'No
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 2).Value = ワードファイル.Revisions(コメント番号).Date   '作成日
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 3).Value = ワードファイル.Revisions(コメント番号).Author '作成者
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 4).Value = ワードファイル.Name 'ファイル名
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 5).Value = ワードファイル.Revisions(コメント番号).Range.Information(wdActiveEndAdjustedPageNumber) 'ページ
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 6).Value = ワードファイル.Revisions(コメント番号).Range.Information(wdFirstCharacterLineNumber)    '行数
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 7).Value = 変更履歴タイプ取得(ワードファイル.Revisions(コメント番号).Type)        '変更種別
        ThisWorkbook.Worksheets(変更履歴シート名).Cells(行, 8).Value = Mid(ワードファイル.Revisions(コメント番号).Range, 1, 255) '変更内容
        行 = 行 + 1
    Next コメント番号
End Sub

Private Function 変更履歴タイプ取得(ByVal タイプ As Integer) As String
    Select Case タイプ
        Case wdNoRevision: 変更履歴タイプ取得 = "変更なし"
        Case wdRevisionConflict: 変更履歴タイプ取得 = "競合"
        Case wdRevisionDelete: 変更履歴タイプ取得 = "削除"
        Case wdRevisionDisplayField: 変更履歴タイプ取得 = "フィールド表示の変更"
        Case wdRevisionInsert: 変更履歴タイプ取得 = "挿入"
        Case wdRevisionParagraphNumber: 変更履歴タイプ取得 = "段落番号の変更"
        Case wdRevisionParagraphProperty: 変更履歴タイプ取得 = "段落のプロパティの変更"
        Case wdRevisionProperty: 変更履歴タイプ取得 = "プロパティの変更"
        Case wdRevisionReconcile: 変更履歴タイプ取得 = "解決された競合"
        Case wdRevisionReplace: 変更履歴タイプ取得 = "置換"
        Case wdRevisionSectionProperty: 変更履歴タイプ取得 = "セクションのプロパティの変更"
        Case wdRevisionStyle: 変更履歴タイプ取得 = "スタイルの変更"
        Case wdRevisionStyleDefinition: 変更履歴タイプ取得 = "スタイル定義の変更"
        Case wdRevisionTableProperty: 変更履歴タイプ取得 = "表のプロパティの変更"
        Case wdRevisionCellDeletion: 変更履歴タイプ取得 = "表のセルの削除"
        Case wdRevisionCellInsertion: 変更履歴タイプ取得 = "表のセルの挿入"
        Case wdRevisionCellMerge: 変更履歴タイプ取得 = "表のセルの結合"
        Case wdRevisionMovedFrom: 変更履歴タイプ取得 = "内容の移動元"
        Case wdRevisionMovedTo: 変更履歴タイプ取得 = "内容の移動先"
    End Select
End Function
