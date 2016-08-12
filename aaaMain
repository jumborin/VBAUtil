Option Explicit

Sub 実行()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim ファイル名, 検索Path As String
    Dim ワードファイル As Word.Document
    Dim コメント As New コメント
    Dim 変更履歴 As New 変更履歴
    
    Call コメント.シート設定
    Call 変更履歴.シート設定
    検索Path = ThisWorkbook.Sheets(1).Cells(1, 2).Text
    ファイル名 = Dir(検索Path & "\*.doc*")
   
    Dim ワードアプリ As New Word.Application
    ワードアプリ.DisplayAlerts = wdAlertsNone
    ワードアプリ.Visible = True
    
    Do While ファイル名 <> ""
        Set ワードファイル = ワードアプリ.Documents.Open(検索Path & "\" & ファイル名)
        Call コメント.データ出力(ワードファイル)
        Call 変更履歴.データ出力(ワードファイル)
        ワードファイル.Close
        ファイル名 = Dir()
    Loop
    
    ワードアプリ.Quit
    Set ワードアプリ = Nothing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
