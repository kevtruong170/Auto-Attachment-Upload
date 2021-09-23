'Validating Sender's email address
Function validateSender(Sender As String) As Integer
    
    'On 0, sender's email address is verified, else, script will terminate.
    If InStr(Sender, "@hotmail.com") Then
        validateSender = 0
    Else
        validateSender = 1
    End If
    
End Function

'Validating the file extension of the attachments
Function validateFileExt(fName As String) As Integer
    
    'Valid files will return 0 and be inputted into the corresponding folder
    If UCase(fName) Like "*PNG*" Or UCase(fName) Like "*JPG*" Or UCase(fName) Like "*HEIC*" Or UCase(fName) Like "*JPEG*" Then
        validateFileExt = 0
    Else
        validateFileExt = 1
    End If
    
End Function

'Functionality to create a folder when user responds, yes.
Function makeDir(newPath As String) As Integer

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    makeDir = 2
    On Error GoTo endFunc
    fs.createFolder (newPath)
    makeDir = 0

endFunc:
End Function

Public Sub SavePhotos(Message As Outlook.MailItem)

    'Use if statement to restrict it to specific users of email addresses
    If (validateSender(Message.SenderEmailAddress) = 0) Then
        
        'Initialization of variables
        Dim oAttachment As Outlook.Attachment
        Dim FolderDest As String
        Dim Path As String
        
        'Path of folder location (not actual folder destination, but where it is located).
        Path = "S:\Pictures\Shipping Pictures\"
    
        'Concatenating strings of actual folder destination where photos will be placed.
        FolderDest = Path & Message.Subject & "\"
        
        Dim FolderExists As String
                
        FolderExists = Dir(FolderDest, vbDirectory)
        
        Dim isError As Integer
        
        'Validating the existence of the directory
        If FolderExists = "" Then
            Dim Response As Integer
            Response = MsgBox("Error, folder directory for """ & Message.Subject & """ does not exist. Would you like to create a new folder?", vbQuestion + vbYesNo)
            
            If (Response = vbYes) Then
                isError = makeDir(FolderDest)
                If (isError = 2) Then
                    GoTo ErrorHandler
                End If
                GoTo NewFolder
            ElseIf (Response = vbNo) Then
                MsgBox "No folder was created and no pictures were uploaded."
            End If
        Else
        
            
            
NewFolder:
            If (Message.Attachments.Count = 0) Then
                isError = 3
                GoTo ErrorHandler
            End If
            
            'Validating and transferring each attachment to folder destination.
            For Each oAttachment In Message.Attachments
                isError = 1
                If (validateFileExt(oAttachment.DisplayName) = 0) Then
                    On Error GoTo ErrorHandler
                    oAttachment.SaveAsFile FolderDest & oAttachment.DisplayName
                    isError = 0
                End If
            Next

'if isError = 1, there is an error in uploading file
ErrorHandler:
            If (isError = 1) Then
                MsgBox "Error in uploading images."
            ElseIf (isError = 2) Then
                MsgBox "Error when creating folder, invalid folder title. Please check message subject."
            ElseIf (isError = 3) Then
                MsgBox "Message has no attachments, folder exists but is empty."
            Else
                MsgBox "Pictures have been uploaded to " + FolderDest + " email was deleted."
                Message.delete
            End If
        End If
        
    End If

End Sub
