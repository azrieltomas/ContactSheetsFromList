Function sheet_check(ByVal sheetname As String) As Boolean
    Dim returnBool As Boolean
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = sheetname Then
            returnBool = True
            Exit For
        Else
            returnBool = False
        End If
    Next ws

    sheet_check = returnBool
End Function

Sub format_fix(ByVal lastRow As Integer)
    'remove colours and borders first
    With Sheets("MAIN").Columns("B")
        .Borders.LineStyle = xlNone
    End With
    With Sheets("MAIN").Range("B1:B" & lastRow)
        .Font.Name = Calibri
        .Font.Size = 10
        .Font.Underline = xlUnderlineStyleNone
        .Font.Color = RGB(0, 0, 0)
        .BorderAround Weight:=xlMedium
        .Borders(xlInsideHorizontal).Weight = xlThin
        .HorizontalAlignment = xlCenter
        .Interior.ColorIndex = 0
    End With
    With Sheets("MAIN").Range("B1").Font
        .Size = 12
        .FontStyle = "Bold"
    End With
End Sub

Sub print_setup()
    Dim lastRow As Integer
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        'skip MAIN
        If ws.Name <> "MAIN" Then
            ws.Activate
            ws.Columns("A").ColumnWidth = 22
            ws.Columns("B").ColumnWidth = 300
            ws.Columns("C").ColumnWidth = 300
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            ws.Range("A2:C" & lastRow).Select
            ActiveSheet.PageSetup.PrintArea = Selection.Address
        End If
    Next ws
    'return to MAIN
    Sheets("MAIN").Activate
End Sub

Sub create_single()
    Dim xName, nameClean As String
    Dim i, lastRow, bandRow As Integer
    Dim templateSh, mainSh As Worksheet

    Set templateSh = Sheets("TEMPLATE")
    Set mainSh = Sheets("MAIN")

    'get new Name
    xName = WorksheetFunction.Proper(InputBox("Enter name"))
    nameClean = Left(Replace(UCase(xName), "/", "_"), 24)
    
    'error checking
    If xName = "" Then
        MsgBox "No value entered"
        End
    End If
    If sheet_check(nameClean) Then
        MsgBox "Sheet for " & xName & " has already been created"
        End
    End If

    With Sheets("MAIN")
        'Now to add the name to the list and sort
        lastRow = .Cells(Rows.Count, 2).End(xlUp).Row + 1
        .Range("B" & lastRow).Value = xName
        .Columns("B").Sort key1:=Range("B2"), Header:=xlYes
        'For loop to determine where that cell sits
        For i = 2 To lastRow
            If .Range("B" & i).Value = xName Then
                bandRow = i
            End If
        Next i
    End With

    'Copy the template to the correct position
    templateSh.Copy after:=Sheets(bandRow)
    ActiveSheet.Name = nameClean
    'change title on sheet
    Sheets(nameClean).Range("B2").Value = xName
    'hyperlink cell on MAIN to new sheet
    mainSh.Activate
    mainSh.Hyperlinks.Add Range("B" & bandRow), "", "'" & nameClean & "'!A1"
    'Formatting fixes
    format_fix (lastRow)
End Sub

Sub create_entries()
    Dim i, j As Integer
    Dim lastRow As Integer
    Dim xName, nameClean As String
    Dim templateSh, mainSh As Worksheet

    Set templateSh = Sheets("TEMPLATE")
    Set mainSh = Sheets("MAIN")

    'get last entry row
    lastRow = mainSh.Cells(Rows.Count, 2).End(xlUp).Row

    'Sort first
    mainSh.Columns("B").Sort key1:=Range("B2"), Header:=xlYes

     For i = 2 To lastRow
        xName = mainSh.Range("B" & i).Value
        'remove invalid / char and truncate to 24 characters
        nameClean = Left(Replace(UCase(xName), "/", "_"), 24)
        'check if it already exists and error out
        If sheet_check(nameClean) Then
            MsgBox "Sheet for " & xName & " has already been created" & vbNewLine & _
                "Please use the manual add function"
            End
        End If
    
        templateSh.Copy after:=Sheets(Sheets.Count)
        'need to remove invalid characters :\/?*[]'
        ActiveSheet.Name = nameClean
        'change title on sheet
        Sheets(nameClean).Range("B2").Value = xName

        'hyperlink cell on MAIN to new sheet
        mainSh.Activate
        mainSh.Hyperlinks.Add Range("B" & i), "", "'" & nameClean & "'!A1"
     Next i

     'Hyperlinks look crap
     format_fix (lastRow)
End Sub

Sub delete_sheets()
    Dim ws, mainSh As Worksheet
    Dim i, lastRow As Integer

    Set mainSh = Sheets("MAIN")
    
    If MsgBox("Are you sure you want to delete all sheets", vbYesNo, "Confirm") = vbYes Then
        For Each ws In ActiveWorkbook.Worksheets
            'Shutup, OR doesn't work
            If ws.Name <> "MAIN" Then
                If ws.Name <> "TEMPLATE" Then
                    Application.DisplayAlerts = False
                    ws.Delete
                    Application.DisplayAlerts = True
                End If
            End If
        Next ws
        'delete hyperlinks
        mainSh.Activate
        lastRow = mainSh.Cells(Rows.Count, 2).End(xlUp).Row
        For i = 2 To lastRow
            Range("B" & i).Hyperlinks.Delete
        Next i
        format_fix (lastRow)
    End If
End Sub