Option Explicit 'This line ensures that all variables in the code must be explicitly declared. It helps prevent typos and makes the code easier to understand and maintain.

' Declare variables to store recipient and attachment information.
' These variables are declared outside any specific procedure, making them accessible from anywhere within the code module.
Dim recipients As Collection ' Holds a collection of recipients. Each recipient will be stored as a dictionary, allowing access to their information using keys like "Email", "CC", "BCC", etc.
Dim attachmentPaths As Collection ' Holds collection of file paths, each representing an attachment to be included in the emails.

' This event procedure is executed when the UserForm is initialized (loaded).
Private Sub UserForm_Initialize()
    ' Set fixed width and height for the UserForm.
    ' This ensures the form maintains a consistent size and appearance.
    BulkMailerForm.Width = 335 ' Set the width of the UserForm to 335 pixels.
    BulkMailerForm.Height = 400 ' Set the height of the UserForm to 400 pixels.

    LoadSignatures ' Call the LoadSignatures procedure to populate the signature dropdown list when the form initializes.
    Set recipients = New Collection ' Initialize the recipients collection. This creates an empty collection to store recipient data later.
    Set attachmentPaths = New Collection ' Initialize the attachmentPaths collection. This creates an empty collection to store attachment file paths later.
End Sub

' This procedure loads available email signatures from the user's signature folder and populates a combobox on the form.
Private Sub LoadSignatures()
    Dim signatureDir As String ' Declare a variable to store the path to the user's signature folder.
    Dim fso As Object ' Declare a variable to interact with the file system.
    Dim file As Object ' Declare a variable to represent a file.

    signatureDir = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Signatures\" ' Construct the path to the user's signature folder. This path is standard for most Windows systems.
    Set fso = CreateObject("Scripting.FileSystemObject") ' Create a FileSystemObject. This object allows you to work with files and folders.

    ' Check if the signature directory exists.
    If fso.FolderExists(signatureDir) Then ' If the signature folder exists, proceed to load signatures.
        ' Loop through each file in the signature directory.
        For Each file In fso.GetFolder(signatureDir).Files ' Iterate through each file within the signature folder.
            If file.Name Like "*.htm" Then ' Check if the file has an ".htm" extension, indicating it's an HTML signature file.
                cmbSignatures.AddItem file.Name ' If the file is an HTML signature, add its name to the signatures combobox (cmbSignatures).
            End If
        Next file ' Move to the next file in the signature directory.
    End If ' End the check for the signature directory's existence.
End Sub

' This procedure handles the click event of the "Load Text" button.
' It allows the user to load the content of a text file into the email body textbox.
Private Sub btnLoadText_Click()
    Dim filePath As String ' Declare a variable to store the selected file path.
    filePath = Application.GetOpenFilename("Text files (*.txt), *.txt") ' Open the file dialog box, allowing the user to select a text file. The selected file's path is stored in the filePath variable.

    ' Check if a file was selected.
    If filePath <> "False" Then ' If the user selects a file (filePath is not "False"), proceed to load its content.
        txtEmailContent.Text = ReadFile(filePath) ' Read the content of the selected file using the ReadFile function and set it as the text of the email body textbox (txtEmailContent).
    End If ' End the check for file selection.
End Sub

' This procedure handles the click event of the "Save Text" button.
' It allows the user to save the current content of the email body textbox to a text file.
Private Sub btnSaveText_Click()
    Dim filePath As String ' Declare a variable to store the chosen file path for saving.
    filePath = Application.GetSaveAsFilename("Text files (*.txt), *.txt") ' Open the "Save As" dialog box, allowing the user to choose a location and name for saving the file as a text file.

    ' Check if a file path was provided.
    If filePath <> "False" Then ' If the user provides a valid file path (filePath is not "False"), proceed to save the content.
        WriteFile filePath, txtEmailContent.Text ' Call the WriteFile function to write the content of the email body textbox (txtEmailContent) to the specified file path.
        MsgBox "Template saved successfully!", vbInformation, "Bulk Mailer" ' Display a message box informing the user that the template has been saved successfully.
    End If ' End the check for a valid file path.
End Sub

' This procedure handles the click event of the "Load Recipients" button.
' It allows the user to load recipient data from a CSV file.
Private Sub btnLoadRecipients_Click()
    Dim filePath As String ' store the selected file path.
    Dim ws As Worksheet '  worksheet containing the recipient data.
    Dim lastRow As Long ' last row containing data in the worksheet.
    Dim lastCol As Long ' last column containing data
    Dim i As Long '  iteratie through rows.
    Dim j As Long ' iterate through columns.
    Dim headers As Collection ' Declare a collection to store the column headers from the CSV file.

    filePath = Application.GetOpenFilename("CSV files (*.csv), *.csv") ' Open the file dialog box, filtering for CSV files and allowing the user to select a file.

    ' Check if a file was selected.
    If filePath <> "False" Then ' If the user selects a file (filePath is not "False"), proceed to load recipient data.
        Set ws = Workbooks.Open(filePath).Sheets(1) ' Open the selected CSV file and set the first worksheet to the ws variable.
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row containing data in column A. This assumes the first column ("A") has data in all used rows.
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Find the last column containing data in the first row. This assumes the first row contains headers.
        Set recipients = New Collection ' Re-initialize the recipients collection to clear any previous data.
        Set headers = New Collection ' Initialize the headers collection to store column headers.

        ' Read headers from the first row.
        For j = 1 To lastCol ' Iterate through each column in the first row (header row).
            headers.Add ws.Cells(1, j).Value ' Add the value of each cell in the first row to the headers collection.
        Next j

        ' Read recipient data from the worksheet.
        For i = 2 To lastRow ' Iterate through each row starting from the second row (data rows).
            Dim recipient As Object ' Declare a variable to represent a single recipient.
            Set recipient = CreateObject("Scripting.Dictionary") ' Create a dictionary object to store the recipient's data. Each key in the dictionary will be a header, and the value will be the corresponding data for that recipient.

            ' Populate the recipient dictionary with data from each column.
            For j = 1 To headers.Count ' Iterate through each column in the current row.
                recipient.Add headers(j), ws.Cells(i, j).Value ' Add a new key-value pair to the recipient dictionary. The key is the header from the headers collection, and the value is the cell value from the current row and column.
            Next j

            recipients.Add recipient ' Add the recipient dictionary to the recipients collection.
        Next i

        ws.Parent.Close False ' Close the CSV file without saving changes.

        ' Load variables into cmbVariables combobox.
        cmbVariables.Clear ' Clear any existing items in the cmbVariables combobox.
        Dim key As Variant ' Declare a variable to iterate through the keys of the recipient dictionary.

        ' Add each header from the first recipient's data as an item in the cmbVariables combobox.
        For Each key In recipients(1).Keys ' Iterate through each key in the first recipient's dictionary.
            cmbVariables.AddItem key ' Add the key as an item in the cmbVariables combobox. This allows the user to select these variables to insert into the email body.
        Next key
    End If ' End the check for file selection.
End Sub

' This procedure handles the double-click event of the cmbVariables combobox.
Private Sub cmbVariables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    InsertVariableIntoBody ' Call the InsertVariableIntoBody procedure to insert the selected variable into the email body.
End Sub

' This procedure handles the KeyDown event of the cmbVariables combobox.
Private Sub cmbVariables_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ' Check if the Enter key is pressed.
        InsertVariableIntoBody ' If the Enter key is pressed, call the InsertVariableIntoBody procedure.
    End If
End Sub

' This procedure inserts the selected variable from the cmbVariables combobox into the email body textbox (txtEmailContent).
Private Sub InsertVariableIntoBody()
    Dim variable As String ' Declare a variable to store the selected variable.
    variable = cmbVariables.Text ' Get the selected text from the cmbVariables combobox and store it in the variable variable.

    ' Check if a variable is selected.
    If variable <> "" Then ' If a variable is selected (the variable is not empty), proceed to insert it.
        txtEmailContent.SelText = "{" & variable & "}" ' Insert the selected variable enclosed in curly braces "{}" into the email body textbox (txtEmailContent) at the current cursor position or replacing the selected text.
    End If ' End the check for variable selection.
End Sub


' This procedure handles the click event of the "Preview Email" button.
' It generates a preview of the email using the first recipient's data.
Private Sub btnPreviewEmail_Click()
    ' Input validation: Check if recipients are loaded.
    If recipients.Count = 0 Then
        MsgBox "No recipients loaded to preview the email.", vbExclamation, "Bulk Mailer" ' Display an error message if no recipients are loaded.
        Exit Sub '
    End If

    ' Input validation: Check if the subject is entered.
    If txtSubject.Text = "" Then
        MsgBox "Please enter or generate an email subject.", vbExclamation, "Missing Subject" ' Display an error message if the subject is empty.
        Exit Sub ' 
    End If

    ' Input validation: Check if a signature is selected.
    If cmbSignatures.ListIndex = -1 Then
        MsgBox "Please select an email signature.", vbExclamation, "Missing Signature" ' Display an error message if no signature is selected.
        Exit Sub '
    End If

    ' Declare variables for Outlook objects and email content.
    Dim outlook As Object '
    Dim emailItem As Object ' email being created.
    Dim signature As String ' store the HTML content of the selected signature.
    Dim signatureDir As String ' stores the path to the user's signature folder.
    Dim signatureHtmlPath As String ' store the full path to the selected signature file.
    Dim key As Variant ' A variable to iterate through the keys (headers) of the recipient's data.

    ' Get the signature content.
    signatureDir = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Signatures\" ' Construct the path to the user's signature folder.
    signatureHtmlPath = signatureDir & cmbSignatures.Value ' Construct the full path to the selected signature file.
    signature = ReadFile(signatureHtmlPath) ' Read the content of the signature file and store it in the signature variable.

    ' Prepare the email content using the first recipient's data.
    Dim emailBody As String ' A variable to store the email body content.
    emailBody = txtEmailContent.Text ' Initialize the email body with the content from the email body textbox (txtEmailContent).
    Dim emailSubject As String ' A variable to store the email subject.
    emailSubject = txtSubject.Text ' Initialize the email subject with the content from the subject textbox (txtSubject).

    ' Replace variables in the email body and subject with the first recipient's data.
    For Each key In recipients(1).Keys ' Iterate through each key (header) in the first recipient's data.
        emailBody = Replace(emailBody, "{" & key & "}", recipients(1)(key)) ' Replace any occurrence of the variable enclosed in curly braces (e.g., "{Name}") with the corresponding value from the recipient's data.
        emailSubject = Replace(emailSubject, "{" & key & "}", recipients(1)(key)) ' Do the same replacement for the email subject.
    Next key

    ' Create and display the preview email.
    Set outlook = CreateObject("Outlook.Application") ' Create an Outlook application object.
    Set emailItem = outlook.CreateItem(0) ' Create a new email item.

    ' Set email details.
    emailItem.Subject = emailSubject ' Set the email subject.
    emailItem.HTMLBody = Replace(emailBody, vbCrLf, "<br>") & signature ' Set the email body as HTML, replacing line breaks with "<br>" tags, and appending the signature.
    emailItem.To = recipients(1)("Email") ' Set the "To" recipient.
    If recipients(1)("CC") <> "" Then emailItem.CC = recipients(1)("CC") ' Set the "CC" recipients if available.
    If recipients(1)("BCC") <> "" Then emailItem.BCC = recipients(1)("BCC") ' Set the "BCC" recipients if available.

    ' Add attachments if any.
    Dim j As Long ' A counter variable for iterating through attachments.
    For j = 1 To attachmentPaths.Count
        emailItem.Attachments.Add attachmentPaths(j) ' Add each attachment from the attachmentPaths collection to the email.
    Next j

    emailItem.Display ' Display the email preview.
End Sub

' This procedure handles the click event of the "Add Attachment" button.
' It allows the user to select multiple files to attach to the emails.
Private Sub btnAddAttachment_Click()
    Dim filePath As Variant ' Declare a variable to store the selected file paths. This is declared as a Variant to handle multiple file selections.
    filePath = Application.GetOpenFilename("All files (*.*), *.*", MultiSelect:=True) ' Open the file dialog box, allowing the user to select multiple files.

    ' Check if files were selected.
    If IsArray(filePath) Then ' If the user selects multiple files, filePath will be an array, so check if it's an array.
        Dim i As Long ' A counter variable for iterating through the selected files.
        For i = LBound(filePath) To UBound(filePath) ' Iterate through each selected file path in the filePath array.
            attachmentPaths.Add filePath(i) ' Add each selected file path to the attachmentPaths collection.
        Next i
        MsgBox "Attachments added successfully!", vbInformation, "Attachments" ' Display a message box confirming the attachments have been added.
    End If ' End the check for file selections.
End Sub

' This procedure handles the click event of the "Send Emails" button.
' It sends personalized emails to each recipient in the recipients collection.
Private Sub btnSendEmails_Click()
    ' Input validation: Check if the subject is entered.
    If txtSubject.Text = "" Then
        MsgBox "Please enter or generate an email subject.", vbExclamation, "Missing Subject"
        Exit Sub
    End If

    ' Input validation: Check if a signature is selected.
    If cmbSignatures.ListIndex = -1 Then
        MsgBox "Please select an email signature.", vbExclamation, "Missing Signature"
        Exit Sub
    End If

    ' Declare variables for Outlook objects and email content.
    Dim outlook As Object ' A variable to represent the Outlook application.
    Dim emailItem As Object ' A variable to represent the email being created.
    Dim signature As String ' A variable to store the HTML content of the selected signature.
    Dim signatureDir As String ' A variable to store the path to the user's signature folder.
    Dim signatureHtmlPath As String ' A variable to store the full path to the selected signature file.
    Dim i As Long ' A counter variable for iterating through recipients.

    ' Get the signature content.
    signatureDir = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Signatures\" ' Construct the path to the user's signature folder.
    signatureHtmlPath = signatureDir & cmbSignatures.Value ' Construct the full path to the selected signature file.
    signature = ReadFile(signatureHtmlPath) ' Read the content of the signature file and store it in the signature variable.

    Set outlook = CreateObject("Outlook.Application") ' Create an Outlook application object.

    ' Loop through each recipient and send the email.
    For i = 1 To recipients.Count ' Iterate through each recipient in the recipients collection.
        Set emailItem = outlook.CreateItem(0) ' Create a new email item for the current recipient.

        ' Dynamically generate the email subject.
        emailItem.Subject = txtSubject.Text ' Initialize the email subject with the content from the subject textbox (txtSubject).
        Dim key As Variant ' A variable to iterate through the keys (headers) of the recipient's data.
        For Each key In recipients(i).Keys ' Iterate through each key (header) in the current recipient's data.
            emailItem.Subject = Replace(emailItem.Subject, "{" & key & "}", recipients(i)(key)) ' Replace any occurrence of the variable enclosed in curly braces (e.g., "{Name}") with the corresponding value from the current recipient's data.
        Next key

        ' Dynamically generate the email body.
        Dim emailBody As String ' A variable to store the email body content.
        emailBody = txtEmailContent.Text ' Initialize the email body with the content from the email body textbox (txtEmailContent).
        For Each key In recipients(i).Keys ' Iterate through each key (header) in the current recipient's data.
            emailBody = Replace(emailBody, "{" & key & "}", recipients(i)(key)) ' Replace any occurrence of the variable enclosed in curly braces (e.g., "{Name}") with the corresponding value from the current recipient's data.
        Next key

        ' Set the email body as HTML, replacing line breaks with "<br>" tags, and appending the signature.
        emailItem.HTMLBody = Replace(emailBody, vbCrLf, "<br>") & signature

        ' Set recipient fields.
        emailItem.To = recipients(i)("Email") ' Set the "To" recipient.
        If recipients(i)("CC") <> "" Then emailItem.CC = recipients(i)("CC") ' Set the "CC" recipients if available.
        If recipients(i)("BCC") <> "" Then emailItem.BCC = recipients(i)("BCC") ' Set the "BCC" recipients if available.

        ' Attach files.
        Dim j As Long ' A counter variable for iterating through attachments.
        For j = 1 To attachmentPaths.Count
            emailItem.Attachments.Add attachmentPaths(j) ' Add each attachment from the attachmentPaths collection to the email.
        Next j

        emailItem.Send ' Send the email.
    Next i ' Move to the next recipient.

    MsgBox "Emails sent successfully!", vbInformation, "Bulk Mailer" ' Display a message box confirming the emails have been sent.
End Sub

' This function reads the content of a text file and returns it as a string.
Private Function ReadFile(filePath As String) As String
    Dim fso As Object ' A variable to represent the FileSystemObject.
    Dim ts As Object ' A variable to represent the TextStream object for reading the file.

    Set fso = CreateObject("Scripting.FileSystemObject") ' Create a FileSystemObject.
    Set ts = fso.OpenTextFile(filePath, 1) ' Open the file for reading.
    ReadFile = ts.ReadAll ' Read the entire content of the file and assign it to the function's return value.
    ts.Close ' Close the TextStream object.
End Function

' This procedure writes the provided content to a file.
Private Sub WriteFile(filePath As String, content As String)
    Dim fso As Object ' A variable to represent the FileSystemObject.
    Dim ts As Object ' A variable to represent the TextStream object for writing to the file.

    Set fso = CreateObject("Scripting.FileSystemObject") ' Create a FileSystemObject.
    Set ts = fso.CreateTextFile(filePath, True) ' Create a new file for writing.
    ts.Write content ' Write the provided content to the file.
    ts.Close ' Close the TextStream object.
End Sub