Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Disable events and screen updates
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Define the source worksheet with error check
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("First day")
    If ws Is Nothing Then
        MsgBox "Source sheet 'First day' not found. Check the name.", vbCritical
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
            firstDataRow = dataRange.Rows(1).row + 1 ' Assumes header in row 1
            
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
            
            ' Convert column B in "First day" to numbers (skip header)
            Dim colBRange As Range
            Set colBRange = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).row)
            For Each cell In colBRange
                If Not IsEmpty(cell) Then
                    If IsNumeric(cell.Value) Then
                        cell.Value = CDbl(cell.Value) ' Convert to double (number)
                    End If
                End If
            Next cell
            Debug.Print "Converted First day column B to numbers"
            
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
            
            ' Copy to "Progress reports" sheet with error check
            Dim wsProgress As Worksheet
            On Error Resume Next
            Set wsProgress = ThisWorkbook.Sheets("Progress reports")
            If wsProgress Is Nothing Then
                MsgBox "Target sheet 'Progress reports' not found. Check the name.", vbCritical
                GoTo CleanUp
            End If
            On Error GoTo ErrorHandler
            
            ' Set rows for Progress reports
            Dim progressFirstRow As Long
            Dim progressTotalsRow As Long
            progressFirstRow = 1    ' First row from source goes to row 1
            progressTotalsRow = 2   ' Totals row goes to row 2
            
            ' Add current date to column A for both rows
            wsProgress.Cells(progressFirstRow, 1).Value = Date ' Today’s date for first row
            wsProgress.Cells(progressTotalsRow, 1).Value = Date ' Today’s date for totals row
            
            ' Copy first row from D-M (4-13) to Progress reports B-K (2-11) in row 1
            wsProgress.Range(wsProgress.Cells(progressFirstRow, 2), wsProgress.Cells(progressFirstRow, 11)).Value = _
                ws.Range(ws.Cells(1, 4), ws.Cells(1, 13)).Value
            
            ' Copy totals from D-M (4-13) to Progress reports B-K (2-11) in row 2
            wsProgress.Range(wsProgress.Cells(progressTotalsRow, 2), wsProgress.Cells(progressTotalsRow, 11)).Value = _
                ws.Range(ws.Cells(lastRow, 4), ws.Cells(lastRow, 13)).Value
            
            Debug.Print "First row copied to Progress reports row " & progressFirstRow & ", columns B-K"
            Debug.Print "Totals copied to Progress reports row " & progressTotalsRow & ", columns B-K"
            
            ' Copy to Notes sheet, keeping all data initially then clearing O-R and adding headers
            Dim wsNotes As Worksheet
            On Error Resume Next
            Set wsNotes = ThisWorkbook.Sheets("Notes")
            If wsNotes Is Nothing Then
                MsgBox "Target sheet 'Notes' not found. Check the name.", vbCritical
                GoTo CleanUp
            End If
            On Error GoTo ErrorHandler
            
            ' Clear previous content in Notes sheet
            wsNotes.Cells.Clear
            
            ' Get source dimensions
            Dim lastCol As Long
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            ' Copy all columns from First day to Notes
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy
            wsNotes.Cells(1, 1).PasteSpecial xlPasteAll
            
            ' Clear columns O-R (15-18) in Notes sheet and add headers
            If lastCol >= 15 Then
                Dim clearRange As Range
                Set clearRange = wsNotes.Range(wsNotes.Cells(1, 15), wsNotes.Cells(lastRow, 18))
                clearRange.ClearContents
                
                ' Add headers to columns O-R in row 1
                wsNotes.Cells(1, 15).Value = "today"
                wsNotes.Cells(1, 16).Value = "Last update"
                wsNotes.Cells(1, 17).Value = "notes"
                wsNotes.Cells(1, 18).Value = "Feb notes"
            End If
            
            Application.CutCopyMode = False
            Debug.Print "Copied all data to Notes sheet, cleared columns O-R, and added headers"
        End If
    End If

CleanUp:
    ' Re-enable events and screen updates
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub
