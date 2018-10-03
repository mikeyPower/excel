#The below code will help split an excel workbook containning multiple sheets (tabs) and save them as individual files


Sub SplitTest()
    Dim Sht As Worksheet
    Dim fName As String
    Dim ShtCountBk1 As Integer
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ShtCountBk1 = IIf(Sheets.Count Mod 2 = 1, Sheets.Count / 2 + 0.5, Sheets.Count / 2)
    Set neww = Workbooks.Add
    For Each Sht In ThisWorkbook.Worksheets
        i = i + 1
        If i > ShtCountBk1 Then
            fName = Replace(ThisWorkbook.Name, ".xls", "")
            neww.SaveAs ThisWorkbook.Path & "\" & fName & " (1).xls"
            Set neww = Workbooks.Add
            i = 1
        End If
        Sht.Copy after:=Worksheets(neww.Sheets.Count)
        If i = 1 Then
            For Each ws In Worksheets
                If ws.Name <> Sht.Name Then
                    ws.Delete
                End If
            Next ws
        End If
    Next Sht
    fName = Replace(ThisWorkbook.Name, ".xls", "")
    neww.SaveAs ThisWorkbook.Path & "\" & fName & " (2).xls"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub