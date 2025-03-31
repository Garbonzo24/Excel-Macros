Sub SendAREmails()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim customerName As String
    Dim customerNum As String
    Dim email As String
    Dim pastDueBalance As Double
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Notes") ' Replace "Sheet1" with your sheet name
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Create Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Loop through each row
    For i = 2 To lastRow ' Assuming row 1 is headers
        customerName = ws.Cells(i, 1).Value   ' Column A
        customerNum = ws.Cells(i, 2).Value    ' Column B
        email = ws.Cells(i, 3).Value          ' Column C
        pastDueBalance = ws.Cells(i, 12).Value ' Column L
        
        ' Check if past due balance is positive (overdue)
        If pastDueBalance > 100000 And email <> "" Then
            ' Create new email
            Set OutlookMail = OutlookApp.CreateItem(0)
            
            With OutlookMail
                .To = email
                .Subject = "Payment Reminder: Overdue Balance for Customer #" & customerNum
                .Body = "Dear " & customerName & "," & vbCrLf & vbCrLf & _
                        "This is a reminder that your account (Customer #" & customerNum & ") " & _
                        "has an overdue balance of $" & Format(pastDueBalance, "#,##0.00") & "." & vbCrLf & _
                        "Please settle the outstanding amount at your earliest convenience." & vbCrLf & _
                        "Contact us at [Your Contact Info] if you have questions or need payment details." & vbCrLf & vbCrLf & _
                        "Thank you," & vbCrLf & _
                        "[Your Name]" & vbCrLf & _
                        "[Your Company]"
                .Display ' Use .Display instead of .Send to review emails before sending
            End With
            
            ' Optional: Mark as sent in Excel
            ws.Cells(i, 13).Value = "Email Sent " & Now ' Writes to Column M
        End If
    Next i
    
    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    MsgBox "Emails have been processed!", vbInformation

End Sub
