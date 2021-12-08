Attribute VB_Name = "Módulo1"
Sub Rodar_relatorio()
Dim x As Long
Dim y As Long
Dim z As Long
Dim ultima As Long

Application.ScreenUpdating = False

x = 2

verificaCel = Sheets("Base_Solange").Cells(x, 38).Value
Do While verificaCel <> ""
    x = x + 1
    verificaCel = Sheets("Base_Solange").Cells(x, 38).Value
Loop

y = x - 1
coluna_linha = "AN" & y
Sheets("Base_Solange").Range("AN2").Copy Destination:=Range("AN3:" + coluna_linha)
Sheets("Base_Solange").Range("AN3:" + coluna_linha).Copy
Range("AN3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AO" & y
Sheets("Base_Solange").Range("AO2").Copy Destination:=Range("AO3:" + coluna_linha)
Sheets("Base_Solange").Range("AO3:" + coluna_linha).Copy
Range("AO3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AP" & y
Sheets("Base_Solange").Range("AP2").Copy Destination:=Range("AP3:" + coluna_linha)
Sheets("Base_Solange").Range("AP3:" + coluna_linha).Copy
Range("AP3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AQ" & y
Sheets("Base_Solange").Range("AQ2").Copy Destination:=Range("AQ3:" + coluna_linha)
Sheets("Base_Solange").Range("AQ3:" + coluna_linha).Copy
Range("AQ3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AR" & y
Sheets("Base_Solange").Range("AR2").Copy Destination:=Range("AR3:" + coluna_linha)
Sheets("Base_Solange").Range("AR3:" + coluna_linha).Copy
Range("AR3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AS" & y
Sheets("Base_Solange").Range("AS2").Copy Destination:=Range("AS3:" + coluna_linha)
Sheets("Base_Solange").Range("AS3:" + coluna_linha).Copy
Range("AS3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AT" & y
Sheets("Base_Solange").Range("AT2").Copy Destination:=Range("AT3:" + coluna_linha)
Sheets("Base_Solange").Range("AT3:" + coluna_linha).Copy
Range("AT3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AU" & y
Sheets("Base_Solange").Range("AU2").Copy Destination:=Range("AU3:" + coluna_linha)
Sheets("Base_Solange").Range("AU3:" + coluna_linha).Copy
Range("AU3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AV" & y
Sheets("Base_Solange").Range("AV2").Copy Destination:=Range("AV3:" + coluna_linha)
Sheets("Base_Solange").Range("AV3:" + coluna_linha).Copy
Range("AV3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AW" & y
Sheets("Base_Solange").Range("AW2").Copy Destination:=Range("AW3:" + coluna_linha)
Sheets("Base_Solange").Range("AW3:" + coluna_linha).Copy
Range("AW3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AX" & y
Sheets("Base_Solange").Range("AX2").Copy Destination:=Range("AX3:" + coluna_linha)
Sheets("Base_Solange").Range("AX3:" + coluna_linha).Copy
Range("AX3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AY" & y
Sheets("Base_Solange").Range("AY2").Copy Destination:=Range("AY3:" + coluna_linha)
Sheets("Base_Solange").Range("AY3:" + coluna_linha).Copy
Range("AY3:" + coluna_linha).PasteSpecial xlPasteValues


coluna_linha = "AZ" & y
Sheets("Base_Solange").Range("AZ2").Copy Destination:=Range("AZ3:" + coluna_linha)
Sheets("Base_Solange").Range("AZ3:" + coluna_linha).Copy
Range("AZ3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "BA" & y
Sheets("Base_Solange").Range("BA2").Copy Destination:=Range("BA3:" + coluna_linha)
Sheets("Base_Solange").Range("BA3:" + coluna_linha).Copy
Range("BA3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "BB" & y
Sheets("Base_Solange").Range("BB2").Copy Destination:=Range("BB3:" + coluna_linha)
Sheets("Base_Solange").Range("BB3:" + coluna_linha).Copy
Range("BB3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "BC" & y
Sheets("Base_Solange").Range("BC2").Copy Destination:=Range("BC3:" + coluna_linha)
Sheets("Base_Solange").Range("BC3:" + coluna_linha).Copy
Range("BC3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "BD" & y
Sheets("Base_Solange").Range("BD2").Copy Destination:=Range("BD3:" + coluna_linha)
Sheets("Base_Solange").Range("BD3:" + coluna_linha).Copy
Range("BD3:" + coluna_linha).PasteSpecial xlPasteValues

coluna_linha = "AU" & y
Sheets("Base_Solange").Range("AU2").Copy Destination:=Range("AU3:" + coluna_linha)
Sheets("Base_Solange").Range("AU3:" + coluna_linha).Copy
Range("AU3:" + coluna_linha).PasteSpecial xlPasteValues


ThisWorkbook.Sheets("Dinâmica").PivotTables("Tabela dinâmica1").PivotCache.Refresh
ThisWorkbook.Sheets("Dinâmica").PivotTables("Tabela dinâmica2").PivotCache.Refresh
ThisWorkbook.Sheets("Dinâmica").PivotTables("Tabela dinâmica3").PivotCache.Refresh
ThisWorkbook.Sheets("Dacs Transfer").PivotTables("PivotTable1").PivotCache.Refresh

Application.ScreenUpdating = True



End Sub

