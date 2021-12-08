Attribute VB_Name = "Módulo3"
Sub enviaremail()
Dim ultima_feeder As Long
Dim ultima_cabotagem As Long
Dim ultima_linha As Long
Dim Signature As Variant
Dim OutMail As Object
Set objeto_outlook = CreateObject("Outlook.Application")

Set Email = objeto_outlook.createitem(0)

nome = "Feeder_pendentes"
nome2 = "Norte_pendentes"
nome3 = "Nordeste_pendentes"
nome4 = "Sul_Sudeste_pendentes"

Application.ScreenUpdating = False

If Sheets("Arquivos").Visible = False Then
    Sheets("Arquivos").Visible = True
End If
If Sheets("Norte").Visible = False Then
    Sheets("Norte").Visible = True
End If
If Sheets("Nordeste").Visible = False Then
    Sheets("Nordeste").Visible = True
End If
If Sheets("SulSudeste").Visible = False Then
    Sheets("SulSudeste").Visible = True
End If
'Caso feeder
ThisWorkbook.Sheets("Dinâmica").Activate
ThisWorkbook.Sheets("Dinâmica").Range("D14").Select
Selection.ShowDetail = True
ultima_feeder = Range("A2").End(xlDown).Row
Range("B1:F" & ultima_feeder).Copy
ThisWorkbook.Sheets("Arquivos").Range("A1").PasteSpecial xlPasteValues
ThisWorkbook.Sheets("Arquivos").Activate
ActiveSheet.Copy
With ActiveWorkbook
.SaveAs ThisWorkbook.Path & "\" & nome & ".xlsx"
ActiveWorkbook.Close
End With
ThisWorkbook.Sheets("Arquivos").Activate
Range("A1:E1000").ClearContents

'Casos cabotagem
ThisWorkbook.Sheets("Dinâmica").Activate
ThisWorkbook.Sheets("Dinâmica").Range("C14").Select
Selection.ShowDetail = True
ultima_cabotagem = Range("A2").End(xlDown).Row
Range("B1:F" & ultima_cabotagem).Copy
ThisWorkbook.Sheets("Arquivos").Range("A1").PasteSpecial xlPasteValues
ThisWorkbook.Sheets("Arquivos").Activate
Range("A1:E1").Select
Selection.AutoFilter
'NORTE
ThisWorkbook.Sheets("Arquivos").Range("B1").AutoFilter Field:=2, Criteria1:="MAO", Operator:=xlOr, Criteria2:="VLC"
ultima_linha = Range("A2").End(xlDown).Row
Range("A1:E" & ultima_linha).Copy
ThisWorkbook.Sheets("Norte").Range("A1").PasteSpecial xlPasteValues
ThisWorkbook.Sheets("Norte").Activate
ActiveSheet.Copy
With ActiveWorkbook
.SaveAs ThisWorkbook.Path & "\" & nome2 & ".xlsx"
ActiveWorkbook.Close
End With
ThisWorkbook.Sheets("Arquivos").Activate
If ThisWorkbook.Sheets("Arquivos").AutoFilterMode Or ThisWorkbook.Sheets("Arquivos").FilterMode Then
    ThisWorkbook.Sheets("Arquivos").ShowAllData
End If

'NORDESTE

ThisWorkbook.Sheets("Arquivos").Range("B1").AutoFilter Field:=2, Criteria1:=Array("PEC", "SUA", "SSA"), Operator:=xlFilterValues
ultima_linha = Range("A2").End(xlDown).Row
Range("A1:E" & ultima_linha).Copy
ThisWorkbook.Sheets("Nordeste").Range("A1").PasteSpecial xlPasteValues
ThisWorkbook.Sheets("Nordeste").Activate
ActiveSheet.Copy
With ActiveWorkbook
.SaveAs ThisWorkbook.Path & "\" & nome3 & ".xlsx"
ActiveWorkbook.Close
End With
ThisWorkbook.Sheets("Arquivos").Activate
If ThisWorkbook.Sheets("Arquivos").AutoFilterMode Or ThisWorkbook.Sheets("Arquivos").FilterMode Then
    ThisWorkbook.Sheets("Arquivos").ShowAllData
End If

'Sul-Sudeste

ThisWorkbook.Sheets("Arquivos").Range("B1").AutoFilter Field:=2, Criteria1:=Array("RIG", "SPB", "IOA", "SSZ"), Operator:=xlFilterValues
ultima_linha = Range("A2").End(xlDown).Row
Range("A1:E" & ultima_linha).Copy
ThisWorkbook.Sheets("SulSudeste").Range("A1").PasteSpecial xlPasteValues
ThisWorkbook.Sheets("SulSudeste").Activate
ActiveSheet.Copy
With ActiveWorkbook
.SaveAs ThisWorkbook.Path & "\" & nome4 & ".xlsx"
ActiveWorkbook.Close
End With
ThisWorkbook.Sheets("Arquivos").Activate
If ThisWorkbook.Sheets("Arquivos").AutoFilterMode Or ThisWorkbook.Sheets("Arquivos").FilterMode Then
    ThisWorkbook.Sheets("Arquivos").ShowAllData
End If
ThisWorkbook.Sheets("Arquivos").Range("A1:E1000").ClearContents

If Sheets("Arquivos").Visible = True Then
    Sheets("Arquivos").Visible = False
End If
If Sheets("Norte").Visible = True Then
    Sheets("Norte").Visible = False
End If
If Sheets("Nordeste").Visible = True Then
    Sheets("Nordeste").Visible = False
End If
If Sheets("SulSudeste").Visible = True Then
    Sheets("SulSudeste").Visible = False
End If


Application.ScreenUpdating = True



Email.display
Email.to = "email_ficticio_email.com.br"
Email.cc = "email_ficticio_email.com.br"
Email.Subject = "Faturamento Pendente - Feeder"
Email.Body = "Olá!" & Chr(10) & Chr(10) & "Poderiam por gentileza verificar os casos pendentes em anexo?" & Chr(10) & assinatura
Email.Attachments.Add (ThisWorkbook.Path & "\Feeder_pendentes.xlsx")
HTMLBody = Signature




Set objeto_outlook = CreateObject("Outlook.Application")

Set Email = objeto_outlook.createitem(0)
Email.display

Email.to = "email_ficticio_email.com.br"
Email.cc = "email_ficticio_email.com.br"
Email.Subject = "Faturamento Pendente - Norte"
Email.Body = "Olá!" & Chr(10) & Chr(10) & "Poderiam por gentileza verificar os casos pendentes em anexo?"
Email.Attachments.Add (ThisWorkbook.Path & "\Norte_pendentes.xlsx")

Set objeto_outlook = CreateObject("Outlook.Application")

Set Email = objeto_outlook.createitem(0)
Email.display

Email.to = "email_ficticio_email.com.br"
Email.cc = "email_ficticio_email.com.br"
Email.Subject = "Faturamento Pendente - Nordeste"
Email.Body = "Olá!" & Chr(10) & Chr(10) & "Poderiam por gentileza verificar os casos pendentes em anexo?"
Email.Attachments.Add (ThisWorkbook.Path & "\Nordeste_pendentes.xlsx")

Set objeto_outlook = CreateObject("Outlook.Application")

Set Email = objeto_outlook.createitem(0)
Email.display

Email.to = "email_ficticio_email.com.br"
Email.cc = "email_ficticio_email.com.br"
Email.Subject = "Faturamento Pendente - Sul/Sudeste"
Email.Body = "Olá!" & Chr(10) & Chr(10) & "Poderiam por gentileza verificar os casos pendentes em anexo?"
Email.Attachments.Add (ThisWorkbook.Path & "\Sul_Sudeste_pendentes.xlsx")



End Sub

