Attribute VB_Name = "Módulo2"
Sub Limpar()

Dim ultima As Long
Dim ultima2 As Long

If Range("A2").Value <> "" Then
    ultima = Sheets("Base_Solange").Cells(Rows.Count, 1).End(xlUp).Row
    lugar1 = "AM" & ultima
    Range("A2:" + lugar1).ClearContents
End If

If Range("AN3").Value <> "" Then
    ultima2 = Sheets("Base_Solange").Cells(Rows.Count, 40).End(xlUp).Row
    lugar2 = "BD" & ultima2
    Range("AN3:" + lugar2).ClearContents
End If


End Sub


