Function last_row(ByVal testColumn As Long) As Integer
    'Return last row for passed column
    Dim returnVal As Integer
    With ActiveSheet
        returnVal = .Cells(.Rows.Count, testColumn).End(xlUp).Row
    End With
    last_row = returnVal
End Function

Sub accred_dump()
    Dim contactBook, newBook As Workbook
    Dim ws, detailsSh As Worksheet
    Dim i, j, lastRow, lastRowB As Integer
    
    newonly = MsgBox("New entries only?", vbYesNoCancel)
    If newonly = vbCancel Then
        Exit Sub
    Else

        Set contactBook = ThisWorkbook
        Set newBook = Workbooks.Add
        Set detailsSh = ActiveSheet
        i = 2
    
        'headers
        Cells(1, 1).Value = "Full Name"
        Cells(1, 2).Value = "Company"
        Cells(1, 3).Value = "Details"
        Rows(1).Font.Bold = True
    
        For Each ws In contactBook.Worksheets
            If ws.Name <> "MAIN" Then
                If ws.Name <> "TEMPLATE" Then
                    
                    'where do all the names sit?
                    ws.Activate
                    'use a val to determine if these have already been sent
                    If newonly = vbNo Then
                        lastRow = last_row(2) - 2
                        For j = 14 To lastRow
                            If ws.Cells(j, 2) = "" Then
                                lastRowB = j - 1
                                Exit For
                            Else
                                lastRowB = lastRow
                            End If
                        Next j
                        For j = 14 To lastRowB
                            newBook.Sheets(detailsSh.Name).Cells(i, 1).Value = ws.Cells(j, 2).Value
                            newBook.Sheets(detailsSh.Name).Cells(i, 2).Value = ws.Cells(2, 2).Value
                            newBook.Sheets(detailsSh.Name).Cells(i, 3).Value = ws.Cells(j, 3).Value
                            i = i + 1
                        Next j
                    ElseIf newonly = vbYes Then
                        If ws.Cells(2, 5).Value = "" Then
                            lastRow = last_row(2) - 2
                            For j = 14 To lastRow
                                If ws.Cells(j, 2) = "" Then
                                    lastRowB = j - 1
                                    Exit For
                                Else
                                    lastRowB = lastRow
                                End If
                            Next j
                            For j = 14 To lastRowB
                                newBook.Sheets(detailsSh.Name).Cells(i, 1).Value = ws.Cells(j, 2).Value
                                newBook.Sheets(detailsSh.Name).Cells(i, 2).Value = ws.Cells(2, 2).Value
                                newBook.Sheets(detailsSh.Name).Cells(i, 3).Value = ws.Cells(j, 3).Value
                                i = i + 1
                            Next j
                            ws.Cells(2, 5).Value = "Entry Reported"
                        End If
                    End If
                End If
            End If
        Next ws
    End If

    newBook.Sheets(detailsSh.Name).Columns("A:C").AutoFit
    contactBook.Sheets("MAIN").Activate
    newBook.Activate
End Sub

Sub request_reset()
    Dim contactBook As Workbook
    Dim ws As Worksheet
    
    Set contactBook = ThisWorkbook
    
    For Each ws In contactBook.Worksheets
        If ws.Name <> "MAIN" Then
            If ws.Name <> "TEMPLATE" Then
                ws.Cells(4, 5).Value = ""
            End If
        End If
    Next ws
End Sub



