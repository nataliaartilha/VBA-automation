Attribute VB_Name = "Módulo7"
Sub ImportTableToExcel()
Dim xOutlook    As New Outlook.Application
Dim xMailItem   As MailItem
Dim xTable      As Word.Table
Dim xDoc        As Word.Document
Dim xWs         As Worksheet
Dim i           As Integer
Dim xRow        As Integer

'Columns("T:W").Delete
On Error Resume Next
Range("A1").Select
Set xWs = ThisWorkbook.ActiveSheet
xRow = 1

For Each xMailItem In xOutlook.Application.ActiveExplorer.Selection
    Set xDoc = xMailItem.GetInspector.WordEditor

    For i = 1 To xDoc.Tables.Count
        Set xTable = xDoc.Tables(i)
        
        xTable.Range.Copy
        xWs.Paste
        xRow = xRow + xTable.Rows.Count + 1
        xWs.Range("A" & CStr(xRow)).Select
    Next
Next
ActiveSheet.Pictures.Delete
End Sub
'teste apagar imagens

Sub apagar_imagens()
Dim pic As Object
    For Each pic In ActiveSheet.Pictures
        pic.Delete
    Next pic
End Sub

Sub apagar()
ActiveSheet.Pictures.Delete
End Sub
