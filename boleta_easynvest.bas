Attribute VB_Name = "Módulo8"
'Boleta Easynvest
Sub Todes_easy()
Call ImportTableToExcel
Call copiar_colar_easy
Call colar_di_2
Call SalvarAba_easy
Call Enviar_email_easy
Call copiar_python
End Sub

'1-Sub ImportTableToExcel()

Sub copiar_colar_easy()

Dim linha As Integer

linha = 2
Application.CutCopyMode = True
Worksheets("CALCULADORA").Range("E2:N9").ClearContents
'data vencimento
While Worksheets("easynvest").Cells(linha, 5).Value <> ""
    Worksheets("easynvest").Cells(linha, 5).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 8).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'taxa cliente
While Worksheets("easynvest").Cells(linha, 7).Value <> ""
    Worksheets("easynvest").Cells(linha, 7).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 9).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'taxa emissão
While Worksheets("easynvest").Cells(linha, 6).Value <> ""
    Worksheets("easynvest").Cells(linha, 6).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 10).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'quantidade
While Worksheets("easynvest").Cells(linha, 8).Value <> ""
    Worksheets("easynvest").Cells(linha, 8).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 7).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'DI
While Worksheets("easynvest").Cells(linha, 9).Value <> ""
    Worksheets("easynvest").Cells(linha, 9).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 11).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'PU
While Worksheets("easynvest").Cells(linha, 10).Value <> ""
    Worksheets("easynvest").Cells(linha, 10).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 12).PasteSpecial
    linha = linha + 1
Wend
linha = 2
Worksheets("easynvest").Range("L1") = "Contraparte"
'fazer a contraparte
While Worksheets("easynvest").Cells(linha, 1).Value <> ""
    Worksheets("easynvest").Cells(linha, 12) = "Easynvest"
    linha = linha + 1
Wend
linha = 2
'contraparte
While Worksheets("easynvest").Cells(linha, 12).Value <> ""
    Worksheets("easynvest").Cells(linha, 12).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 5).PasteSpecial
    linha = linha + 1
Wend
'cdi indexador
linha = 2
While Worksheets("easynvest").Cells(linha, 3).Value <> ""
    Worksheets("easynvest").Cells(linha, 3).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 6).PasteSpecial
    linha = linha + 1
Wend
Application.CutCopyMode = False
End Sub
Sub colar_di_2()

Application.CutCopyMode = True
linha1 = 2
linha2 = 1
While Worksheets("CALCULADORA").Cells(linha1, 4).Value <> ""
    Worksheets("CALCULADORA").Cells(1, 2).Value = Worksheets("CALCULADORA").Cells(linha1, 4).Value
    Worksheets("CALCULADORA").Cells(15, 2).Value = Worksheets("CALCULADORA").Cells(linha1, 11).Value
    Worksheets("CALCULADORA").Cells(linha1, 14).Value = Worksheets("CALCULADORA").Cells(19, 2).Value
    linha1 = linha1 + 1
    linha2 = linha2 + 1
Wend
Application.CutCopyMode = False
End Sub

'exporta aba e exlui as macros
Sub SalvarAba_easy()
'Impede que o Excel atualize a tela
Application.ScreenUpdating = False
'Impede que o Excel exiba alertas
Application.DisplayAlerts = False

'Seta uma variável para se referir a nova pasta de trabalho
Dim NovoWB As Workbook
'Cria esta nova aba
Set NovoWB = Workbooks.Add(xlWBATWorksheet)
With NovoWB
'Copia a aba atual para o novo arquivo, como a segunda aba
ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count)
'Deleta a primeira aba do arquivo criado (Aba em branco)
.Worksheets(1).Delete
.Worksheets("easynvest").Columns("L:Z").Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & "boleta_easynvest" & ".xlsx"
'Fecha o novo arquivo
'Workbooks("boleta_agora").Columns("T:Z").Delete
.Close SaveChanges:=True
End With



'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub

Sub Enviar_email_easy()
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
nomeemail = "Aplicações NuInvest - " & Format(Worksheets("CALCULADORA").Range("B7"), "dd/mm/yyyy") & " - BANCO FATOR"
Diretorio = "G:\depto\RENDA\Natalia Artilha\"
Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)
nomeRel = "boleta_easynvest"
With Email
.display
.To = "juliani.cardoso@nubank.com.br;mnegro@fator.com.br;BancoFatorTesouraria@fator.com.br"
.cc = "renda.fixa@easynvest.com.br"
.Subject = nomeemail
.HTMLBody = saudacao & Chr(12) & Chr(12) & "Operação realizada!" & Chr(12) & Chr(12) & "Segue PU no arquivo em anexo." & Chr(12) & Chr(12) & "Atenciosamente," & .HTMLBody
.Attachments.Add (Diretorio & nomeRel & ".xlsx")
'Email.send
End With



End Sub
