# vba script to combine multiple workbooks together (seperate excel files) to form one master workbook

Sub GetSheets()
Path = "C:\Users\dt\Desktop\dt kte\"
Filename = Dir(Path & "*.xls")
  Do While Filename <> ""
  Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
     For Each Sheet In ActiveWorkbook.Sheets
     Sheet.Copy After:=ThisWorkbook.Sheets(1)
  Next Sheet
     Workbooks(Filename).Close
     Filename = Dir()
  Loop
End Sub