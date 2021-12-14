Attribute VB_Name = "Módulo3"
Sub exporta_dados()

Dim wkOrigem1 As Worksheet
Dim wkOrigem2 As Worksheet
Dim wkDestino1 As Worksheet
Dim wkDestino2 As Worksheet
Dim wkDestino3 As Worksheet
Dim wkDestino4 As Worksheet
Dim wkDestino5 As Worksheet
Dim x As Long
Dim z As Long
Dim c As Long
Dim mydate
Dim nome As String
mydate = Date

'Nome da planilha variando de acordo com a data atual
nome_data = Format(Date, "DDMM")
nome_primera = "TMSxDocsys_"
nome_final = ".xlsm"
nome_certo = nome_primera & nome_data & nome_final


Application.ScreenUpdating = False


Workbooks.Open Filename:="V:\PM_ALCT\_Public\COM - Plantão\DOCSYS - MultiTMS\2021\12_Dezembro\Docsys_Dezembro.xlsx"
Workbooks.Open Filename:="V:\PM_ALCT\_Public\COM - Plantão\DOCSYS - MultiTMS\2021\12_Dezembro\PROCV_KARININHA.xlsx"

Set wkOrigem1 = Workbooks(nome_certo).Worksheets("Dinâmica")
Set wkOrigem2 = Workbooks(nome_certo).Worksheets("DacsTransfer")
Set wkDestino1 = Workbooks("Docsys_Dezembro.xlsx").Worksheets("Tracking")
Set wkDestino2 = Workbooks("Docsys_Dezembro.xlsx").Worksheets("DacsTransfer")
Set wkDestino3 = Workbooks("PROCV_KARININHA.xlsx").Worksheets("Controle_Dacs_transfer")
Set wkDestino4 = Workbooks("PROCV_KARININHA.xlsx").Worksheets("Cross_Check_Dacs_Transfer")
Set wkDestino5 = Workbooks("PROCV_KARININHA.xlsx").Worksheets("Erro")

'Verificando a ultima linha preenchida
With wkDestino1
z = 1
verificaCel2 = wkDestino1.Cells(z, 2).Value
Do While verificaCel2 <> ""
    z = z + 1
    verificaCel2 = wkDestino1.Cells(z, 2).Value
Loop
w = z



End With
'verificando a ultima linha preenchida
With wkOrigem1

x = 8
verificaCel = wkOrigem1.Cells(x, 9).Value
Do While verificaCel <> ""
    x = x + 1
    verificaCel = wkOrigem1.Cells(x, 9).Value
Loop
y = x - 2

'Copiando os dados da planilha wkOrigem1 para a planilha wkDestino1
ultima_linha = "L" & y
primeira_linha_nova = "B" & w
wkOrigem1.Range("I8:" + ultima_linha).Copy
wkDestino1.Range(primeira_linha_nova).PasteSpecial xlPasteValues

End With

'Verificando a ultima linha da planilha wkOrigem2
With wkOrigem2
c = 7
verifica_erro = wkOrigem2.Cells(c, 6).Value
Do While verifica_erro <> ""
    c = c + 1
    verifica_erro = wkOrigem2.Cells(c, 6).Value
Loop


m = c - 1

'Copiando os dados da planilha wkOrigem2 e colocando na planilha wkDestino4
ultima_linha2 = "G" & m
wkOrigem2.Range("A7:" + ultima_linha2).Copy
wkDestino4.Range("A2").PasteSpecial xlPasteValues
End With



'Copiando os dados da planilha wkDestino2 e colocando na planilha wkDestino3

'Inserir filtros
With wkDestino2
wkDestino2.Range("L1").AutoFilter Field:=12, Criteria1:=Array("Verificar", "TI", "Livrar de erro"), Operator:=xlFilterValues
d = 2
verificar_erros = wkDestino2.Cells(d, 4).Value
Do While verificar_erros <> ""
    d = d + 1
    verificar_erros = wkDestino2.Cells(d, 4).Value
Loop


ultima_linha3 = "N" & d
wkDestino2.Range("D2:" + ultima_linha3).Copy
wkDestino3.Range("B2").PasteSpecial xlPasteValues
If wkDestino2.AutoFilterMode Or wkDestino2.FilterMode Then
    wkDestino2.ShowAllData
End If

End With
With wkDestino1

a = 2
verificaCel_novo = wkDestino1.Cells(a, 2).Value
Do While verificaCel_novo <> ""
    a = a + 1
    verificaCel_novo = wkDestino1.Cells(a, 2).Value
    
Loop


b = a - 1


'preenchendo com a data atual planilha wkDestino1
Do While w <= b
    wkDestino1.Cells(w, 1).Value = mydate
    w = w + 1
Loop

End With

'Filtrando os casos #N/D da planilha wkDestino4 e colocando na planilha wkDestino5
With wkDestino4
    wkDestino4.Range("I1").AutoFilter Field:=9, Criteria1:="#N/D"
    ultima_linha = wkDestino4.Range("I2").End(xlDown).Row
    wkDestino4.Range("A2:G" & ultima_linha).Copy
    wkDestino5.Range("A1").PasteSpecial xlPasteValues
End With
    
'Copiando os dados da planilha wkDestino5 e colocando na wkDestino2
'Atualizando as formulas da planilha wkDestino2
With wkDestino5
    ultima_linha_novo = wkDestino5.Range("A1").End(xlDown).Row
    wkDestino5.Range("A1:G" & ultima_linha_novo).Copy
    ultima_linha_controle = wkDestino2.Range("D1").End(xlDown).Row
    linha_certa = ultima_linha_controle + 1
    inserir = "D" & linha_certa
    wkDestino2.Range(inserir).PasteSpecial xlPasteValues
    ultima_linha_controle_formulas = wkDestino2.Range("D1").End(xlDown).Row
    wkDestino2.Range("C2").Copy
    colar_formula1 = "C" & linha_certa
    colar_formula2 = "C" & ultima_linha_controle_formulas
    wkDestino2.Range(colar_formula1, [colar_formula2]).PasteSpecial xlPasteFormulas
    wkDestino2.Range("B2").Copy
    colar_formula3 = "B" & linha_certa
    colar_formula4 = "B" & ultima_linha_controle_formulas
    wkDestino2.Range(colar_formula3, [colar_formula4]).PasteSpecial xlPasteFormulas
    wkDestino2.Range("L2").Copy
    colar_formula5 = "L" & linha_certa
    colar_formula6 = "L" & ultima_linha_controle_formulas
    wkDestino2.Range(colar_formula5, [colar_formula6]).PasteSpecial xlPasteFormulas
    colar_data = "A" & linha_certa
    colar_data2 = "A" & ultima_linha_controle_formulas
    wkDestino2.Range(colar_data, [colar_data2]) = mydate
    
End With

'Atualizando as formulas da planilha wkDestino1
With wkDestino1
    ultima_linha_controle_tracking = wkDestino1.Range("F1").End(xlDown).Row
    linha_certa_tracking = ultima_linha_controle_tracking + 1
    ultima_linha_controle_formulas_tracking = wkDestino1.Range("E1").End(xlDown).Row
    wkDestino1.Range("F2").Copy
    colar_formula7 = "F" & linha_certa_tracking
    colar_formula8 = "F" & ultima_linha_controle_formulas_tracking
    wkDestino1.Range(colar_formula7, [colar_formula8]).PasteSpecial xlPasteFormulas
End With



Workbooks("PROCV_KARININHA.xlsx").Close SaveChanges:=False
Application.ScreenUpdating = True

End Sub

