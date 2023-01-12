Attribute VB_Name = "Módulo17"
'nova futura
Sub Todas_futura()
Worksheets("futura").Activate
Call ImportTableToExcel
Call copiar_colar_futura
Call tx_cliente
Call colar_di_futura
Call Colar_PU_DI_futura
Call SalvarAba_futura
Call Enviar_email_futura
Call copiar_python_calculadora_2
End Sub


Sub copiar_colar_futura()
Dim linha, linha2 As Integer

linha = 2
linha2 = 2
'taxa cliente
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 7).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 9).PasteSpecial
    
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2


linha2 = 2
'taxa emissão
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 6).Copy
    
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10).PasteSpecial
    'Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10).Style = "Percent"
    'Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10) = Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10) / 100
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'quantidade
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 4).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 7).PasteSpecial
    
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2

'fazer a contraparte
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha2, 11) = "nova futura"
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'contraparte
Worksheets("futura").Cells(1, 11) = "Contraparte"
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 11).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 5).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend

linha = 2
linha2 = 2
Worksheets("futura").Cells(1, 10) = "Indexador"
'fazer a indexador
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 10) = "CDI"
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'indxador
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 10).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 6).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'data vencimento
While Worksheets("futura").Cells(linha, 1).Value <> ""
    Worksheets("futura").Cells(linha, 3).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(28, 2).PasteSpecial
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(29, 2).Copy
     Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 8).PasteSpecial xlPasteValues
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
End Sub
Sub tx_cliente()
Dim linha As Integer
linha = 2
Worksheets("CALCULADORA_2").Activate
Range("O1:O9").Copy
Range("N1:N9").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
While Worksheets("CALCULADORA_2").Cells(linha, 4) <> ""
    Worksheets("CALCULADORA_2").Cells(linha, 14).Select
    Selection.Copy
    Worksheets("CALCULADORA_2").Cells(linha, 9).PasteSpecial xlPasteValues
    linha = linha + 1
Wend
End Sub
Sub colar_di_futura()
Dim linha1, linha2 As Integer
Application.CutCopyMode = True
linha1 = 2
linha2 = 1
While Worksheets("CALCULADORA_2").Cells(linha1, 5).Value <> ""
    Call CopiarPU
    Worksheets("CALCULADORA_2").Cells(1, 2).Value = Worksheets("CALCULADORA_2").Cells(linha1, 4).Value
    Call CopiarPU
    Worksheets("CALCULADORA_2").Cells(linha1, 11).Value = Worksheets("CALCULADORA_2").Cells(15, 3).Value
    Worksheets("CALCULADORA_2").Cells(linha1, 12).Value = Worksheets("CALCULADORA_2").Cells(19, 3).Value
    'Worksheets("CALCULADORA").Cells(linha1, 14).Value = Worksheets("CALCULADORA").Cells(19, 2).Style = "Percent"
    linha1 = linha1 + 1
    linha2 = linha2 + 1

Wend
Application.CutCopyMode = False


End Sub

Sub Colar_PU_DI_futura()
Dim linha1 As Integer
linha1 = 2

Worksheets("futura").Cells(1, 12) = "DI"
Worksheets("futura").Cells(1, 9) = "PU"
While Worksheets("CALCULADORA_2").Cells(linha1, 5).Value <> ""
    Worksheets("futura").Cells(linha1, 12) = Worksheets("CALCULADORA_2").Cells(linha1, 11).Value
    Worksheets("futura").Cells(linha1, 9) = Worksheets("CALCULADORA_2").Cells(linha1, 12).Value
    Worksheets("futura").Cells(linha1, 12).Style = "Percent"
    linha1 = linha1 + 1
'Worksheets("CALCULADORA_2").Cells(linha1, 11).Value = Worksheets("CALCULADORA_2").Cells(15, 3).Value
'Worksheets("CALCULADORA_2").Cells(linha1, 12).Value = Worksheets("CALCULADORA_2").Cells(19, 3).Value
Wend
End Sub

'exporta aba e exlui as macros
Sub SalvarAba_futura()
Dim ultimalinha As Integer
Dim x As Integer
'Impede que o Excel atualize a tela
Application.ScreenUpdating = False
'Impede que o Excel exiba alertas
Application.DisplayAlerts = False
Worksheets("futura").Activate
Range("D1").Select
'Worksheets("PYTHON").Range("B2:I100").ClearContents
Selection.End(xlDown).Select
'x.Select
Selection.End(xlDown).Select
x = Selection.Row

'ultimalinha = Sheets("futura").Cells(Rows.Count, 4).End(xlDown).Row
'MsgBox (x)
'ultimalinha2 = Sheets("futura").Cells(Rows.Count, 4).End(xlUp).Row
'linha_excluir = Sheets("futura").Cells(ultimalinha + 1, 2)

'Seta uma variável para se referir a nova pasta de trabalho
Dim NovoWB As Workbook
'Cria esta nova aba
Set NovoWB = Workbooks.Add(xlWBATWorksheet)
With NovoWB
'Copia a aba atual para o novo arquivo, como a segunda aba
ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count)
'Deleta a primeira aba do arquivo criado (Aba em branco)
.Worksheets(1).Delete
.Worksheets("futura").Columns("J:Q").Delete
.Worksheets("futura").Rows(x).Delete
.Worksheets("futura").Rows(x).Delete
.Worksheets("futura").Rows(x).Delete
.Worksheets("futura").Rows(x).Delete
.Worksheets("futura").Rows(x).Delete
.Worksheets("futura").Rows(x).Delete
.Worksheets("futura").Rows(x).Delete
'.Worksheets("futura").Rows(linha_excluir).Delete



'.Worksheets("genial").Columns("S:Z").Delete
'.Worksheets("genial").Rows("12:32").Delete
'.Worksheets("genial").Rows("1:8").Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & "boleta_futura" & ".xlsx"
'Fecha o novo arquivo
'Workbooks("boleta_agora").Columns("T:Z").Delete
.Close SaveChanges:=True
End With



'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub
Sub Enviar_email_futura()
Dim txtFileName, nomearq, nomeRel, nomeemail As String
Dim saudacao As String


'Range(Selection, Selection.End(xlToRight)).Select
'Range(Selection, Selection.End(xlDown)).Select
'tabela = Selection

If Hour(Now) < 12 Then
saudacao = "Bom dia."
ElseIf Hour(Now) >= 12 And Hour(Now) <= 18 Then
saudacao = "Boa tarde, prezados!"
ElseIf Hour(Now) > 18 Then
saudacao = "Boa noite, prezados!"
End If
nomeemail = "Renda Fixa - Aplicações"
Diretorio = "G:\depto\RENDA\Natalia Artilha\"
Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)
nomeRel = "boleta_futura"
With Email
.display
.To = "danilo.silva@futurainvestimentos.com.br"
.cc = "mesarf@futurainvestimentos.com.br;captacao@fator.com.br"
.Subject = nomeemail
.HTMLBody = saudacao & Chr(12) & Chr(12) & "Operação realizada!" & Chr(12) & Chr(12) & "Segue PU no arquivo em anexo." & Chr(12) & Chr(12) & "Atenciosamente," & .HTMLBody
.Attachments.Add (Diretorio & nomeRel & ".xlsx")
'Email.send
End With



End Sub


Sub copiar_python_calculadora_2()
Dim linha, ultimalinha As Integer
linha = 2
Worksheets("PYTHON").Activate
'Worksheets("PYTHON").Range("B2:I100").ClearContents
ultimalinha = Sheets("PYTHON").Cells(Rows.Count, 2).End(xlUp).Row

While Worksheets("CALCULADORA_2").Cells(linha, 5) <> ""
    'Worksheets("python").Range("A"&ultimalinha) = Worksheets("CALCULADORA").Range("D"&linha:"L"&linha).Value
    Worksheets("CALCULADORA_2").Activate
    Worksheets("CALCULADORA_2").Cells(linha, 5).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Worksheets("PYTHON").Activate
    
    Worksheets("PYTHON").Cells(ultimalinha + 1, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    linha = linha + 1
    ultimalinha = ultimalinha + 1
Wend
End Sub


