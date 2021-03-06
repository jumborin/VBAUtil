Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicosoftOfficeで実施する場合、削除すること。

Option Explicit

Const コメント一覧シート名 = "コメント一覧"

' 
' WordファイルのコメントをExcelで一覧化する処理
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
    
    ' Excel内にコメント一覧シートを作成
    ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = コメント一覧シート名
    ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1, 1).Value = "No"
    ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1, 2).Value = "ファイル名"
    ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1, 3).Value = "ページ"
    ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1, 4).Value = "行"
    ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1, 5).Value = "作成者"
    ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1, 6).Value = "コメント内容"

	' ワードファイル処理のための準備
    Dim ワードアプリ As New Word.Application
    ワードアプリ.DisplayAlerts = wdAlertsNone
    ワードアプリ.Visible = True
    
    ' 1ファイルずつ処理
    ファイル名 = Dir(検索Path & "\*.doc*")
    Do While ファイル名 <> ""
        Set ワードファイル = ワードアプリ.Documents.Open(検索Path & "\" & ファイル名)
        Call コメント一覧データ出力(ワードファイル)
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

Private Sub コメント一覧データ出力(ByVal ワードファイル As Word.Document)
    Dim 行,コメント番号 As Long
    
    On Error Resume Next
    
    行 = ThisWorkbook.Worksheets(コメント一覧シート名).Cells(1048576, 1).End(xlUp).Row + 1
    For コメント番号 = 1 To ワードファイル.Comments.Count
        ThisWorkbook.Worksheets(コメント一覧シート名).Cells(行, 1).Value = "=ROW()-1"
        ThisWorkbook.Worksheets(コメント一覧シート名).Cells(行, 2).Value = ワードファイル.Name
        ThisWorkbook.Worksheets(コメント一覧シート名).Cells(行, 3).Value = ワードファイル.Comments(コメント番号).Scope.Information(wdActiveEndAdjustedPageNumber)
        ThisWorkbook.Worksheets(コメント一覧シート名).Cells(行, 4).Value = ワードファイル.Comments(コメント番号).Scope.Information(wdFirstCharacterLineNumber)
        ThisWorkbook.Worksheets(コメント一覧シート名).Cells(行, 5).Value = ワードファイル.Comments(コメント番号).Author
        ThisWorkbook.Worksheets(コメント一覧シート名).Cells(行, 6).Value = ワードファイル.Comments(コメント番号).Range
        行 = 行 + 1
    Next コメント番号
End Sub
