Sub Test5()
    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim SentItemsFolder As Object
    Dim TargetMailbox As Object
    Dim MailItem As Object
    Dim FoundMail As Object
    Dim ws As Worksheet
    Dim logWorkbook As Workbook
    Dim logSheet As Worksheet
    Dim lastRow As Long
    Dim logRow As Long
    Dim i As Long
    Dim candidateID As String
    Dim candidateName As String
    Dim searchSubject As String
    Dim filteredItems As Object
    Dim recipient As String
    Dim logPath As String
    Dim currentDateTime As String
    Dim sendFromAddress As String
    Dim startDate As Date
    Dim endDate As Date
    Dim fso As Object
    Dim emailsSent As Boolean
    Dim filter As String
    
    ' Initialize flags
    emailsSent = False

    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Create a new instance of Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Get the Namespace object
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    
    ' Access the specific mailbox by name
    On Error Resume Next
    Set TargetMailbox = OutlookNamespace.Folders("abc@4001.xyz.com")
    On Error GoTo 0
    
    If TargetMailbox Is Nothing Then
        MsgBox "Mailbox not found or inaccessible.", vbExclamation
        Exit Sub
    End If
    
    ' Get the Sent Items folder
    Set SentItemsFolder = TargetMailbox.Folders("Sent Items")
    
    If SentItemsFolder Is Nothing Then
        MsgBox "Unable to access Sent Items folder.", vbExclamation
        Exit Sub
    End If
    
    ' Date range to limit search to the last 5 days
    startDate = Date - 5
    endDate = Date
    
    ' Create a log workbook
    Set logWorkbook = Workbooks.Add
    Set logSheet = logWorkbook.Sheets(1)
    logSheet.Name = "LogSheet"
    logSheet.Cells(1, 1).Value = "Candidate ID"
    logSheet.Cells(1, 2).Value = "Candidate Name"
    logSheet.Cells(1, 3).Value = "Status"
    
    ' Get the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    logRow = 2
    
    ' Loop through each row in the worksheet
    For i = 2 To lastRow
        candidateID = ws.Cells(i, 1).Value
        candidateName = ws.Cells(i, 2).Value
        searchSubject = ws.Cells(i, 3).Value
        
        ' Use a filter to search within the date range for faster results
        filter = "[ReceivedTime] >= '" & Format(startDate, "yyyy-mm-dd hh:mm:ss") & "' AND " & _
                 "[ReceivedTime] <= '" & Format(endDate, "yyyy-mm-dd hh:mm:ss") & "'"
        
        ' Restrict search to the date range for faster results
        Set filteredItems = SentItemsFolder.Items.Restrict(filter)
        filteredItems.Sort "[ReceivedTime]", True
        
        ' Check if any emails are found
        If filteredItems.Count > 0 Then
            ' Loop through the filtered items to match the subject or body
            For Each FoundMail In filteredItems
                If InStr(FoundMail.Subject, candidateID) > 0 Or InStr(FoundMail.Body, candidateID) > 0 Then
                    ' Get recipient email and send follow-up
                    recipient = FoundMail.Recipients.Item(1).Address
                    Set MailItem = OutlookApp.CreateItem(0)
                    With MailItem
                        .To = recipient
                        .Subject = "RE: " & FoundMail.Subject
                        .Body = FoundMail.Body
                        .SentOnBehalfOfName = "abc@4001.xyz.com" ' Send from mailbox
                        .Send
                    End With
                    
                    ' Log success
                    logSheet.Cells(logRow, 1).Value = candidateID
                    logSheet.Cells(logRow, 2).Value = candidateName
                    logSheet.Cells(logRow, 3).Value = "Email sent successfully"
                    
                    emailsSent = True
                    Exit For
                End If
            Next FoundMail
        Else
            ' Log failure
            logSheet.Cells(logRow, 1).Value = candidateID
            logSheet.Cells(logRow, 2).Value = candidateName
            logSheet.Cells(logRow, 3).Value = "Email not found"
        End If
        
        logRow = logRow + 1
    Next i
    
    ' Set the log file path with current date and time
    currentDateTime = Format(Now, "yyyymmdd_hhmmss")
    logPath = Environ("USERPROFILE") & "\Documents\Logs\Log_" & currentDateTime & ".xlsx"
    
    ' Save log file
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ("USERPROFILE") & "\Documents\Logs\") Then
        fso.CreateFolder Environ("USERPROFILE") & "\Documents\Logs\"
    End If
    logWorkbook.SaveAs logPath
    logWorkbook.Close SaveChanges:=False
    
    ' Cleanup
    Set OutlookApp = Nothing
    Set OutlookNamespace = Nothing
    Set TargetMailbox = Nothing
    Set SentItemsFolder = Nothing
    Set fso = Nothing
End Sub


