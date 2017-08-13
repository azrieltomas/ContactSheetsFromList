Sub create_postal()
    Dim contactBook, newBook As Workbook
    Dim ws, postalsh As Worksheet
    Dim i As Integer
    

    Set contactBook = ThisWorkbook
    Set newBook = Workbooks.Add
    Set postalsh = ActiveSheet
    i = 2

    'headers
    Cells(1, 1).Value = "Company Name"
    Cells(1, 2).Value = "Contact Name"
    Cells(1, 3).Value = "Postal Address"
    Cells(1, 4).Value = "Phone Number"
    Rows(1).Font.Bold = True

    For Each ws In contactBook.Worksheets
        If ws.Name <> "MAIN" Then
            If ws.Name <> "TEMPLATE" Then
                If ws.Cells(11, 3).Value <> "" Then
                    newBook.Sheets(postalsh.Name).Cells(i, 1).Value = ws.Cells(2, 2).Value
                    newBook.Sheets(postalsh.Name).Cells(i, 2).Value = ws.Cells(4, 3).Value
                    newBook.Sheets(postalsh.Name).Cells(i, 3).Value = ws.Cells(11, 3).Value
                    newBook.Sheets(postalsh.Name).Cells(i, 4).Value = ws.Cells(6, 3).Value
                    i = i + 1
                End If
            End If
        End If
    Next ws

    newBook.Sheets(postalsh.Name).Columns("A:F").AutoFit
End Sub

