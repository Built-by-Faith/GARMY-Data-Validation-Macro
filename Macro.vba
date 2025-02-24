Sub Run_GARMY_Checks()

    ' Declare variables to store worksheet, last row, and error messages
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim duplicateFoundC As Boolean
    Dim duplicateFoundG As Boolean
    Dim numbersFoundH As Boolean
    Dim msg As String
    Dim duplicateRowsC As String
    Dim duplicateRowsG As String
    Dim errorRowsH As String
    Dim errorRowsJ As String
    Dim pfWs As Worksheet
    Dim txtBox As MSForms.TextBox
    
    ' Set worksheets to be used in the macro
    Set ws = ThisWorkbook.Sheets("Import Template") ' Worksheet containing data to be checked
    Set pfWs = ThisWorkbook.Sheets("Process Functions") ' Worksheet where error message will be displayed
    
    ' Initialize message and row variables
    msg = "" ' Initialize error message
    duplicateRowsC = "" ' Initialize string to store rows with duplicate NIINs in Column C
    duplicateRowsG = "" ' Initialize string to store rows with duplicate REFNUMs in Column G
    errorRowsH = "" ' Initialize string to store rows with numbers in Column H
    errorRowsJ = "" ' Initialize string to store rows with non-numerical characters or blank cells in Column J
    
    ' Set cell backgrounds to no fill in Columns C, G, and J
    For i = 2 To ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        ' Clear any existing fill colors in Columns C, G, and J
        ws.Cells(i, "C").Interior.ColorIndex = 0
        ws.Cells(i, "G").Interior.ColorIndex = 0
        ws.Cells(i, "J").Interior.ColorIndex = 0
    Next i
    
    ' Get the last row with data in Column H (HA.UNITISHA)
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Check for numbers in Column H (HA.UNITISHA)
    numbersFoundH = False ' Initialize flag to track if numbers are found in Column H
    For i = 2 To lastRow
        ' Check if cell contains a number
        If ws.Cells(i, "H").Value Like "*[0-9]*" Then
            ' Highlight cell with number in red
            ws.Cells(i, "H").Interior.Color = RGB(255, 0, 0)
            numbersFoundH = True ' Set flag to True
            errorRowsH = errorRowsH & i & ", " ' Add row number to error message
        End If
    Next i
    
    ' Add message if numbers are found in Column H, listing row numbers
    If numbersFoundH Then
        ' Remove trailing comma and space from error message
        errorRowsH = Left(errorRowsH, Len(errorRowsH) - 2)
        ' Add error message to main message string
        msg = msg & "Numbers found in Column HA.UNITISHA in Rows: " & errorRowsH & vbCrLf
    End If
    
    ' Check for duplicates in Column C (HA.NIINSNHA)
    duplicateFoundC = False ' Initialize flag to track if duplicates are found in Column C
    For i = 2 To lastRow
        ' Check if cell is not empty
        If ws.Cells(i, "C").Value <> "" Then
            ' Check for duplicates in remaining rows
            For j = i + 1 To lastRow
                ' Check if cell value matches another cell value
                If ws.Cells(i, "C").Value = ws.Cells(j, "C").Value Then
                    ' Highlight cells with duplicate values in red
                    ws.Cells(i, "C").Interior.Color = RGB(255, 0, 0)
                    ws.Cells(j, "C").Interior.Color = RGB(255, 0, 0)
                    duplicateFoundC = True ' Set flag to True
                    ' Add row numbers to error message
                    If InStr(duplicateRowsC, CStr(i)) = 0 Then duplicateRowsC = duplicateRowsC & i & ", "
                    If InStr(duplicateRowsC, CStr(j)) = 0 Then duplicateRowsC = duplicateRowsC & j & ", "
                End If
            Next j
        End If
    Next i
    
    ' Add message if duplicates are found in Column C, listing row numbers
    If duplicateFoundC Then
        ' Remove trailing comma and space from error message
        duplicateRowsC = Left(duplicateRowsC, Len(duplicateRowsC) - 2)
        ' Add error message to main message string
        msg = msg & "The duplicate NIINs in Column C are in Rows: " & duplicateRowsC & vbCrLf
    End If
    
    ' Check for duplicates in Column G (HA.REFNUMHA)
    duplicateFoundG = False ' Initialize flag to track if duplicates are found in Column G
    For i = 2 To lastRow
        ' Check if cell is not empty
        If ws.Cells(i, "G").Value <> "" Then
            ' Check for duplicates in remaining rows
            For j = i + 1 To lastRow
                ' Check if cell value matches another cell value
                If ws.Cells(i, "G").Value = ws.Cells(j, "G").Value Then
                    ' Highlight cells with duplicate values in red
                    ws.Cells(i, "G").Interior.Color = RGB(255, 0, 0)
                    ws.Cells(j, "G").Interior.Color = RGB(255, 0, 0)
                    duplicateFoundG = True ' Set flag to True
                    ' Add row numbers to error message
                    If InStr(duplicateRowsG, CStr(i)) = 0 Then duplicateRowsG = duplicateRowsG & i & ", "
                    If InStr(duplicateRowsG, CStr(j)) = 0 Then duplicateRowsG = duplicateRowsG & j & ", "
                End If
            Next j
        End If
    Next i
    
    ' Add message if duplicates are found in Column G, listing row numbers
    If duplicateFoundG Then
        ' Remove trailing comma and space from error message
        duplicateRowsG = Left(duplicateRowsG, Len(duplicateRowsG) - 2)
        ' Add error message to main message string
        msg = msg & "The duplicate REFNUMs in Column G are in Rows: " & duplicateRowsG & Chr(10)
    End If
    
    ' Check for non-numerical characters or blank cells in Column J
    For i = 3 To lastRow
        ' Check if cell contains non-numerical characters or is blank
        If IsEmpty(ws.Cells(i, "J").Value) Or Not IsNumeric(ws.Cells(i, "J").Value) Then
            ' Highlight cell with non-numerical characters or blank cell in red
            ws.Cells(i, "J").Interior.Color = RGB(255, 0, 0)
            errorRowsJ = errorRowsJ & i & ", " ' Add row number to error message
        End If
    Next i
    
    ' Add message if non-numerical characters or blank cells are found in Column J, listing row numbers
    If errorRowsJ <> "" Then
        ' Remove trailing comma and space from error message
        errorRowsJ = Left(errorRowsJ, Len(errorRowsJ) - 2)
        ' Add error message to main message string
        msg = msg & "Error - There are non-numerical characters or blank cells in Column J in Rows: " & errorRowsJ & Chr(10)
    End If
    
    ' Display the error message in a text box on the Process Functions sheet
    If pfWs.OLEObjects.Count > 0 Then
        ' Check if text box already exists
        For Each oleObj In pfWs.OLEObjects
            ' Check if text box is a Forms.TextBox.1 object
            If oleObj.progID = "Forms.TextBox.1" Then
                Set txtBox = oleObj.Object
                Exit For
            End If
        Next oleObj
    Else
        ' Create a new text box if one doesn't exist
        Set txtBox = pfWs.OLEObjects.Add(ClassType:="Forms.TextBox.1", Left:=pfWs.Range("A15").Left, Top:=pfWs.Range("A15").Top, Width:=300, Height:=200).Object
        With txtBox
            ' Set text box properties
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
            .Font.Name = "Calibri"
            .Font.Size = 11
            .MultiLine = True
            .WordWrap = True
        End With
    End If
    
    ' Display error message in text box
    If msg = "" Then
        ' Display "No errors found" message if no errors were found
        txtBox.Text = "No errors found."
    Else
        ' Display error message
        txtBox.Text = msg
    End If
End Sub

