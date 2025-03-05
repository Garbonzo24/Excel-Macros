Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Disable events and screen updates
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Define the source worksheet with error check
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Daily ar report")
    If ws Is Nothing Then
        MsgBox "Source sheet 'Daily ar report' not found. Check the name.", vbCritical
        GoTo CleanUp
    End If
    On Error GoTo ErrorHandler
    
    ' Debug: Log the pasted range
    Debug.Print "Pasted range: " & Target.Address
    
    ' Define the data range (assuming it starts at A1)
    Dim dataRange As Range
    Set dataRange = ws.Range("A1").CurrentRegion
    
    ' Check if paste intersects with data range
    If Not Intersect(Target, dataRange) Is Nothing Then
        Debug.Print "Data range detected: " & dataRange.Address
        
        ' Clear existing filters
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
            Debug.Print "Cleared existing filter"
        End If
        
        ' Verify column P exists
        If dataRange.Columns.Count >= 16 Then
            ' Apply filter to column P for "Blake Caldwell"
            dataRange.AutoFilter Field:=16, Criteria1:="Blake Caldwell"
            Debug.Print "Filter applied to column P"
            
            ' Delete rows not matching the filter (skip header row)
            Dim rngToDelete As Range
            Dim row As Range
            Dim firstDataRow As Long
            firstDataRow = dataRange.Rows(1).Row + 1 ' Assumes header in row 1
            
            For Each row In dataRange.Offset(1).Resize(dataRange.Rows.Count - 1).Rows
                If row.EntireRow.Hidden Then
                    If rngToDelete Is Nothing Then
                        Set rngToDelete = row
                    Else
                        Set rngToDelete = Union(rngToDelete, row)
                    End If
                End If
            Next row
            
            If Not rngToDelete Is Nothing Then
                rngToDelete.EntireRow.Delete
                Debug.Print "Deleted non-matching rows"
            End If
            
            ' Clear the filter
            ws.AutoFilterMode = False
            Debug.Print "Filter cleared"
            
            ' Recalculate data range after deletion
            Set dataRange = ws.Range("A1").CurrentRegion
            
            ' Convert column B in "Daily ar report" to numbers (skip header)
            Dim colBRange As Range
            Set colBRange = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row)
            For Each cell In colBRange
                If Not IsEmpty(cell) Then
                    If IsNumeric(cell.Value) Then
                        cell.Value = CDbl(cell.Value) ' Convert to double (number)
                    End If
                End If
            Next cell
            Debug.Print "Converted Daily ar report column B to numbers"
            
            ' Add totals row for columns D-M (4-13) with positive numbers only
            Dim lastRow As Long
            lastRow = dataRange.Rows.Count + 1 ' New row after data
            
            ' Add "Totals" label in column A of new row
            ws.Cells(lastRow, 1).Value = "Totals"
            
            ' Calculate sum of positive numbers for columns D (4) to M (13)
            Dim col As Long
            For col = 4 To 13
                Dim sumPositive As Double
                sumPositive = 0
                
                ' Sum only positive numbers in this column
                Dim r As Long
                For r = 2 To lastRow - 1 ' Start at row 2 (skip header), end before totals row
                    Dim cellValue As Variant
                    cellValue = ws.Cells(r, col).Value
                    If IsNumeric(cellValue) Then
                        If cellValue > 0 Then
                            sumPositive = sumPositive + cellValue
                        End If
                    End If
                Next r
                
                ' Place the sum in the totals row
                ws.Cells(lastRow, col).Value = sumPositive
            Next col
            Debug.Print "Totals row added for columns D-M"
            
            ' Copy totals to "Progress reports" sheet with error check
            Dim wsProgress As Worksheet
            On Error Resume Next
            Set wsProgress = ThisWorkbook.Sheets("Progress reports")
            If wsProgress Is Nothing Then
                MsgBox "Target sheet 'Progress reports' not found. Check the name.", vbCritical
                GoTo CleanUp
            End If
            On Error GoTo ErrorHandler
            
            ' Find the last row in "Progress reports" and add new row
            Dim progressLastRow As Long
            progressLastRow = wsProgress.Cells(wsProgress.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Add current date to column A
            wsProgress.Cells(progressLastRow, 1).Value = Date ' Today’s date
            
            ' Copy totals from D-M (4-13) to Progress reports B-K (2-11)
            wsProgress.Range(wsProgress.Cells(progressLastRow, 2), wsProgress.Cells(progressLastRow, 11)).Value = _
                ws.Range(ws.Cells(lastRow, 4), ws.Cells(lastRow, 13)).Value
            Debug.Print "Totals copied to Progress reports, row " & progressLastRow & ", columns B-K"
            
            ' XLOOKUP: Match column B in "Notes" to column B in "Daily ar report", return column L to "Notes" column O
            Dim wsNotes As Worksheet
            On Error Resume Next
            Set wsNotes = ThisWorkbook.Sheets("Notes")
            If wsNotes Is Nothing Then
                MsgBox "Sheet 'Notes' not found. Check the name.", vbCritical
                GoTo CleanUp
            End If
            On Error GoTo ErrorHandler
            
            ' Define ranges for lookup
            Dim notesRange As Range
            Dim dailyRange As Range
            Set notesRange = wsNotes.Range("B2:B" & wsNotes.Cells(wsNotes.Rows.Count, 2).End(xlUp).Row) ' Column B in Notes
            Set dailyRange = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row) ' Column B in Daily ar report
            
            ' Perform XLOOKUP for each value in Notes column B
            Dim noteCell As Range
            For Each noteCell In notesRange
                If Not IsEmpty(noteCell) Then
                    Dim lookupResult As Variant
                    lookupResult = Application.VLookup(noteCell.Value, _
                        ws.Range("B2:L" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row), _
                        11, False) ' 11 = column L relative to B (B=1, L=11)
                    
                    ' Place result in column O (15th column) of Notes
                    If Not IsError(lookupResult) Then
                        wsNotes.Cells(noteCell.Row, 15).Value = lookupResult
                    Else
                        wsNotes.Cells(noteCell.Row, 15).Value = "" ' Blank if no match
                    End If
                End If
            Next noteCell
            Debug.Print "XLOOKUP completed for Notes column O"
        Else
            MsgBox "Error: Data range has only " & dataRange.Columns.Count & " columns. Needs 16 (up to P).", vbExclamation
        End If
    Else
        Debug.Print "Paste outside data range: " & dataRange.Address
    End If
    
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & vbNewLine & "Check sheet names and data range.", vbCritical
    GoTo CleanUp
End Sub
