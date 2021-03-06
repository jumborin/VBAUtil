Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicosoftOfficeで実施する場合、削除すること。

Option Explicit

' 
' Wordファイルの更新履歴とかコメントをExcelで一覧化する処理のメインメソッド
' 
Sub 実行()

	' 表示処理を省略
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 変数の宣言
    Dim ファイル名, 検索Path As String
    Dim ワードファイル As Word.Document
    Dim コメント As New コメント
    Dim 変更履歴 As New 変更履歴
    Const 検索Path = ""
    検索Path = ThisWorkbook.Sheets(1).Cells(1, 2).Text
    ファイル名 = Dir(検索Path & "\*.doc*")
    
    ' 別モジュールのメソッドを呼び出ししてExcel内にシートを作成
    Call コメント.コメント一覧シート設定
    Call 変更履歴.変更履歴シート設定

	' ワードファイル処理のための準備
    Dim ワードアプリ As New Word.Application
    ワードアプリ.DisplayAlerts = wdAlertsNone
    ワードアプリ.Visible = True
    
    ' 1ファイルずつ処理
    Do While ファイル名 <> ""
        Set ワードファイル = ワードアプリ.Documents.Open(検索Path & "\" & ファイル名)
        Call コメント.コメント一覧データ出力(ワードファイル)
        Call 変更履歴.変更履歴データ出力(ワードファイル)
        ワードファイル.Close
        ファイル名 = Dir()
    Loop
    
    ' 後処理
    ワードアプリ.Quit
    Set ワードアプリ = Nothing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
