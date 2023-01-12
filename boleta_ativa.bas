Attribute VB_Name = "Módulo11"
Sub Todos_ativa()
Call AbreMaisRecenteNovo_e_copia_e_cola_2
Call copiar_colar_ativa
Call colar_di_ativa
Call Colar_PU_DI_ativa
Call SalvarAba_atival
Call Enviar_email_ativa
Call copiar_python_calculadora_2
End Sub




'ativa
Sub AbreMaisRecenteNovo_e_copia_e_cola_2()
Application.ScreenUpdating = False
'Applicationd.DisplayAlerts = False

  Dim arqSys As FileSystemObject
  Dim objArq As File
  Dim minhaPasta
  Dim nomearq As String
  Dim dataArq As Date
Workbooks("Captação CDB - Calculadora.nova versao").Activate
Worksheets("CALCULADORA_2").Range("E2:M9").ClearContents
Worksheets("ativa").Range("A1:Z100").ClearContents

        Const Diret As String = "G:\depto\RENDA\Natalia Artilha\Historico_ativa"
        Set arqSys = New FileSystemObject
        Set minhaPasta = arqSys.GetFolder(Diret)
        dataArq = DateSerial(1900, 1, 1)
For Each objArq In minhaPasta.Files
    If objArq.DateLastModified > dataArq Then
        dataArq = objArq.DateLastModified
        nomearq = objArq
    End If
Next objArq
        ActiveWorkbook.FollowHyperlink Address:=nomearq
        Set arqSys = Nothing
        Set minhaPasta = Nothing
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("ativa").Activate
Range("A1").PasteSpecial

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub
Sub copiar_colar_ativa()
Dim linha, linha2 As Integer

linha = 2
linha2 = 2
'taxa cliente
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha, 3).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 9).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'taxa emissão
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha, 4).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10).PasteSpecial
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10).Style = "Percent"
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10) = Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 10) / 100
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'quantidade
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    x = Worksheets("ativa").Cells(linha, 5).Value / 1000
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 7) = x
    
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2

'fazer a contraparte
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha2, 6) = "ativa"
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'contraparte
Worksheets("ativa").Cells(1, 6) = "Contraparte"
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha, 6).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 5).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend

linha = 2
linha2 = 2

'fazer a indexador
Worksheets("ativa").Cells(1, 7) = "Indexador"
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha2, 7) = "CDI"
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'indxador
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha, 7).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 6).PasteSpecial
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
'data vencimento
While Worksheets("ativa").Cells(linha, 1).Value <> ""
    Worksheets("ativa").Cells(linha, 2).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(28, 2).PasteSpecial
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(29, 2).Copy
     Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA_2").Cells(linha2, 8).PasteSpecial xlPasteValues
    linha = linha + 1
    linha2 = linha2 + 1
Wend
linha = 2
linha2 = 2
End Sub

Sub colar_di_ativa()
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

Sub Colar_PU_DI_ativa()
Dim linha1 As Integer
linha1 = 2

Worksheets("ativa").Cells(1, 8) = "DI"
Worksheets("ativa").Cells(1, 9) = "PU"
While Worksheets("CALCULADORA_2").Cells(linha1, 5).Value <> ""
    Worksheets("ativa").Cells(linha1, 8) = Worksheets("CALCULADORA_2").Cells(linha1, 11).Value
    Worksheets("ativa").Cells(linha1, 9) = Worksheets("CALCULADORA_2").Cells(linha1, 12).Value
    Worksheets("ativa").Cells(linha1, 8).Style = "Percent"
    linha1 = linha1 + 1
'Worksheets("CALCULADORA_2").Cells(linha1, 11).Value = Worksheets("CALCULADORA_2").Cells(15, 3).Value
'Worksheets("CALCULADORA_2").Cells(linha1, 12).Value = Worksheets("CALCULADORA_2").Cells(19, 3).Value
Wend
End Sub

'exporta aba e exlui as macros
Sub SalvarAba_atival()
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
.Worksheets("ativa").Columns("K:M").Delete
'.Worksheets("genial").Columns("S:Z").Delete
'.Worksheets("genial").Rows("12:32").Delete
'.Worksheets("genial").Rows("1:8").Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & "boleta_ativa" & ".xlsx"
'Fecha o novo arquivo
'Workbooks("boleta_agora").Columns("T:Z").Delete
.Close SaveChanges:=True
End With



'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub
Sub Enviar_email_ativa()
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
nomeemail = "Aplicacao BANCO FATOR"
Diretorio = "G:\depto\RENDA\Natalia Artilha\"
Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)
nomeRel = "boleta_ativa"
With Email
.display
.To = "rendafixa@ativainvestimentos.com.br;captacao@fator.com.br;middlerendafixa@ativainvestimentos.com.br"
'.cc = ""
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
