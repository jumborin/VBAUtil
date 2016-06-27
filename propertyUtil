' エクセルファイルの個人情報を削除する。
Sub clearProperties(ByVal wb as Workbook)
  wb.Open
  Application.Username = " "
  wb.BuiltinDocumentProperties.Item("Author").Value = Empty
  wb.BuiltinDocumentProperties.Item("Last Author").Value = Empty
  wb.BuiltinDocumentProperties.Item("Company").Value = Empty
  wb.BuiltinDocumentProperties.Item("Manager").Value = Empty
  wb.Save
  MsgBox ("個人情報の削除が完了しました。")
  wb.Close
End Sub
