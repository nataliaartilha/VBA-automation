Attribute VB_Name = "Módulo1"
Public Sub SaveAttachmentsToDisk(MItem As Outlook.MailItem)
Dim oAttachment As Outlook.Attachment
Dim sSaveFolder As String
sSaveFolder = "G:\depto\RENDA\Natalia Artilha\Historico_agora\"
For Each oAttachment In MItem.Attachments
oAttachment.SaveAsFile sSaveFolder & oAttachment.FileName
Next
End Sub
