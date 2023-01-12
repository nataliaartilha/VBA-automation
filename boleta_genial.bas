Attribute VB_Name = "Módulo9"
'tem que estar na aba genial e no email da genial
Sub todos_genial_1()
Worksheets("genial").Activate
Call ImportTableToExcel
Call copiar_colar_genial
'apagar espaços
End Sub
'tem que estar na aba genial
Sub todos_genial_2()
Worksheets("genial").Activate
Call colar_di_genial
Call SalvarAba_genial
Call Enviar_email_genial
Call copiar_python
End Sub


'genial
Sub copiar_colar_genial()
Dim x As String
Dim linha, linha2 As Integer
linha2 = 2
linha = 10
Application.CutCopyMode = True
Worksheets("CALCULADORA").Range("E2:N9").ClearContents
'data vencimento
While Worksheets("genial").Cells(linha, 7).Value <> ""
    Worksheets("genial").Cells(linha, 7).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 8).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
'taxa cliente
While Worksheets("genial").Cells(linha, 11).Value <> ""
    Worksheets("genial").Cells(linha, 11).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 9).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
'taxa emissão
While Worksheets("genial").Cells(linha, 10).Value <> ""
    Worksheets("genial").Cells(linha, 10).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 10).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
'quantidade
While Worksheets("genial").Cells(linha, 12).Value <> ""
    Worksheets("genial").Cells(linha, 12).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 7).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
'DI
While Worksheets("genial").Cells(linha, 17).Value <> ""
    Worksheets("genial").Cells(linha, 17).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 11).PasteSpecial
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 11).Style = "Percent"
    
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
'PU
While Worksheets("genial").Cells(linha, 15).Value <> ""
    Worksheets("genial").Cells(linha, 15).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 12).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
Worksheets("genial").Range("S9") = "Contraparte"
'fazer a contraparte
While Worksheets("genial").Cells(linha, 5).Value <> ""
    Worksheets("genial").Cells(linha2, 19) = "Genial"
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 10
linha2 = 2
'contraparte
While Worksheets("genial").Cells(linha, 19).Value <> ""
    Worksheets("genial").Cells(linha, 19).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 5).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
'cdi indexador
linha = 10
linha2 = 2
While Worksheets("genial").Cells(linha, 9).Value <> ""
    Worksheets("genial").Cells(linha, 9).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 6).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
Application.CutCopyMode = False
End Sub
Sub Clear()
Dim rLocal As Range

'Atribuição das variáveis
  Set rLocal = Worksheets("CALCULADORA").Range("D2:L9")

    rLocal.Replace What:=" ", Replacement:=""
End Sub

Sub teste()
Dim letra, x

linha2 = 2
x = Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 6).Value
letra = CStr(x)
Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha2, 6).Value = Trim(letra)
End Sub

Sub colar_di_genial()
Dim linha1, linha2 As Integer
Application.CutCopyMode = True
linha1 = 2
linha2 = 1
While Worksheets("CALCULADORA").Cells(linha1, 4).Value <> ""
    Worksheets("CALCULADORA").Cells(1, 2).Value = Worksheets("CALCULADORA").Cells(linha1, 4).Value
    Worksheets("CALCULADORA").Cells(15, 2).Value = Worksheets("CALCULADORA").Cells(linha1, 3).Value
    Worksheets("CALCULADORA").Cells(linha1, 14).Value = Worksheets("CALCULADORA").Cells(19, 2).Value
    linha1 = linha1 + 1
    linha2 = linha2 + 1
Wend
Application.CutCopyMode = False
End Sub
'exporta aba e exlui as macros
Sub SalvarAba_genial()
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
.Worksheets("genial").Columns("S:Z").Delete
'.Worksheets("genial").Columns("S:Z").Delete
.Worksheets("genial").Rows("12:32").Delete
.Worksheets("genial").Rows("1:8").Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & "boleta_genial" & ".xlsx"
'Fecha o novo arquivo
'Workbooks("boleta_agora").Columns("T:Z").Delete
.Close SaveChanges:=True
End With



'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub
Sub Enviar_email_genial()
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
nomeemail = "EMISSÃO BANCO FATOR S.A. | GENIAL INVESTIMENTOS " & Format(Worksheets("CALCULADORA").Range("B7"), "dd/mm/yyyy")
Diretorio = "G:\depto\RENDA\Natalia Artilha\"
Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)
nomeRel = "boleta_genial"
With Email
.display
.To = "traderf@genial.com.vc;captacao@fator.com.br"
.cc = "rendafixa@genial.com.vc"
.Subject = nomeemail
.HTMLBody = saudacao & Chr(12) & Chr(12) & "Operação realizada!" & Chr(12) & Chr(12) & "Segue PU no arquivo em anexo." & Chr(12) & Chr(12) & "Atenciosamente," & .HTMLBody
.Attachments.Add (Diretorio & nomeRel & ".xlsx")
'Email.send
End With



End Sub

