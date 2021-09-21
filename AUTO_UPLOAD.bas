Attribute VB_Name = "Module1"
'Validating Sender's email address
Function validateSender(Sender As String) As Integer
    
    'On 0, sender's email address is verified, else, script will terminate.
    If InStr(Sender, "@uoguelph.ca") Or InStr(Sender, "@safdrives.com") Then
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

Public Sub SavePhotos(Message As Outlook.MailItem)
    
    'Use if statement to restrict it to specific users of email addresses
    If (validateSender(Message.SenderEmailAddress) = 0) Then
        
        'Initialization of variables
        Dim oAttachment As Outlook.Attachment
        Dim FolderDest As String
        Dim Path As String
        
        'Path of folder location (not actual folder destination, but where it is located).
        Path = "C:\Users\Kevin\Desktop\TEST\"
    
        'Concatenating strings of actual folder destination where photos will be placed.
        FolderDest = Path & Message.Subject & "\"
        
        Dim FolderExists As String
                
        FolderExists = Dir(FolderDest, vbDirectory)
        
        'Validating the existence of the directory
        If FolderExists = "" Then
            MsgBox "Error, folder directory for """ & Message.Subject & """ does not exist. Nothing was uploaded. Please check exact spelling and formatting of subject."
        Else
        
            Dim isError As Integer
            isError = 1
            
            'Validating and transferring each attachment to folder destination.
            For Each oAttachment In Message.Attachments
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
            Else
                MsgBox "Pictures have been uploaded to " + FolderDest
            End If
        End If
        
    End If

End Sub
