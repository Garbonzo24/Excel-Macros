Private Sub Worksheet_Change(ByVal Target As Range)
    ' Disable events and screen updates to prevent flickering and loops
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Define the worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Check if the change occurred in the data range (starting at A1)
    If Not Intersect(Target, ws.Range("A1").CurrentRegion) Is Nothing Then
        ' Clear existing filters
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
        
        ' Define the data range dynamically (assuming headers in row 1)
        Dim dataRange As Range
        Set dataRange = ws.Range("A1").CurrentRegion
        
        ' Apply filter to column P (16th column) for "Blake Caldwell"
        dataRange.AutoFilter Field:=16, Criteria1:="Blake Caldwell"
    End If
    
    ' Re-enable events and screen updates
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
