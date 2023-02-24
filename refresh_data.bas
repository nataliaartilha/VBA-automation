Attribute VB_Name = "Módulo7"
'automatizar o outro processo
Sub automation_1()
Dim PathName, Filename As String

Application.ScreenUpdating = False
'Applicationd.DisplayAlerts = False

PathName = "G:\depto\RENDA\Formador de Mercado\FUNDOS\"
Filename = "VOLUME NEGOCIADO BBG.xlsx"
Workbooks.Open Filename:=PathName & Filename
Workbooks("VOLUME NEGOCIADO BBG").Activate
Application.OnTime (Now + TimeValue("00:00:20")), "automation_3"
'Application.Run "RefreshAllWorkbooks"
'Application.Run "RefreshAllStaticData"
'Application.Run "RefreshEntireWorksheet"
'Application.Run "RefreshAllWorkbooks"
'Application.Run "RefreshAllStaticData"
Application.ScreenUpdating = True
'Applicationd.DisplayAlerts = True

End Sub
Sub automation_3()
Workbooks("VOLUME NEGOCIADO BBG").Close SaveChanges:=True
Workbooks("Base Relatório").Activate
ThisWorkbook.RefreshAll
'Application.Run "RefreshAllWorkbooks"
'Application.Run "RefreshAllStaticData"
MsgBox ("Atualizou as bases")
'Application.OnTime (Now + TimeValue("00:00:20")), "BODB"



End Sub




