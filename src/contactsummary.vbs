Sub contact_details()
    Dim mainBook, contactBook As Workbook
    Dim ws, contactSh As Worksheet
    Dim i As Integer
          
    Set mainBook = ThisWorkbook
    Set contactBook = Workbooks.Add
    Set contactSh = Worksheets.Add
    contactSh.Name = "Contact Details"
    
    'headers
    contactSh.Activate
    contactSh.Range(Cells(1, 2), Cells(1, 5)).Merge
    contactSh.Cells(1, 2).Value = "CONTACT DETAILS"
    contactSh.Cells(2, 2).Value = "COMPANY"
    contactSh.Cells(2, 3).Value = "CONTACT"
    contactSh.Cells(2, 4).Value = "EMAIL"
    contactSh.Cells(2, 5).Value = "MOBILE"
    
    i = 3
    
    'copy data
    For Each ws In mainBook.Worksheets
        If ws.Name <> "MAIN" Then
            If ws.Name <> "TEMPLATE" Then
                contactSh.Cells(i, 2).Value = ws.Cells(2, 2).Value
                contactSh.Cells(i, 3).Value = ws.Cells(4, 3).Value
                contactSh.Cells(i, 4).Value = ws.Cells(5, 3).Value
                contactSh.Cells(i, 5).Value = ws.Cells(6, 3).Value
                i = i + 1
            End If
        End If
    Next ws
    
    'formatting
    contactSh.Activate
    With contactSh.Range(Cells(1, 2), Cells(i - 1, 5))
        .Font.Name = Calibri
        .Font.Size = 10
        .Font.Underline = xlUnderlineStyleNone
        .Font.Color = RGB(0, 0, 0)
        .BorderAround Weight:=xlMedium
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 0
    End With
    With contactSh.Range(Cells(1, 2), Cells(2, 5))
        .Font.Size = 12
        .Font.FontStyle = "Bold"
        .Interior.Color = RGB(226, 239, 218)
        .EntireColumn.AutoFit
        .RowHeight = 24.75
    End With
    contactSh.Columns(1).ColumnWidth = 2.43
    
End Sub

