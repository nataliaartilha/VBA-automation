Attribute VB_Name = "Módulo8"
Sub automation_2()

Application.ScreenUpdating = False
'Applicationd.DisplayAlerts = False

Dim PathName2, Filename2 As String
Dim linha, x As Integer
'Workbooks("VOLUME NEGOCIADO BBG").Close SaveChanges:=True
'Workbooks("Base Relatório").Activate
'Worksheets("SERVICE ORDER").Activate
'Worksheets("SERVICE ORDER").Range("A1").Select

'x = Sheets("SERVICE ORDER").Cells(Rows.Count, 21).End(xlUp).Row
'MsgBox (x)
'ultimalinha = Range("A1").End(xlDown).Row
'x = Range("A" & Cells.Rows.Count).End(xlUp).Row
'x = Range("A1").End(xlDown).Row
'Worksheets("PYTHON").Cells(ultimalinha + 1, 2).Select
'ultimalinha = Worksheets("SERVICE ORDER").Cells(Worksheets("SERVICE ORDER").Rows.Count.End(xlUp).Row
Workbooks("VOLUME NE BBG").Close SaveChanges:=True
'PathName2 = "G:GOCIADO\depto\RENDA\Formador de Mercado\Relatórios Gerenciais\Controle  Fator\"
'Filename2 = "Boletas CITRIX.xlsm"
linha = 1
'Workbooks.Open Filename:=PathName2 & Filename2
'Worksheets("SERVICE ORDER").Activate
'Cells(linha, 3).Select

If Cells(linha, 3).Value = Workbooks("Base Relatório").Worksheets("RELATÓRIO 5 CORRETORAS").Cells(2, 14) Then
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("Base Relatório").Worksheets("SERVICE ORDER").Cells(ultimalinha + 1, 1).Select
    Worksheets("PYTHON").Cells(ultimalinha + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Else
    linha = linha + 1
    
Application.ScreenUpdating = True
'Applicationd.DisplayAlerts = True
End If

End Sub


