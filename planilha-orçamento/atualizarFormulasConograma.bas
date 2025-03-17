' Sub atualizarFormulasCronograma()
'     Dim wsCronograma As Worksheet
'     Dim wsMemorial As Worksheet
'     Dim ultimaLinha As Long
'     Dim ultimaColuna As Long
'     Dim i As Long, j As Long
    
'     Set wsCronograma = ThisWorkbook.Sheets("CRONOGRAMA")
'     Set wsMemorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")
    
'     ' Encontra a última linha preenchida no cronograma
'     ultimaLinha = wsCronograma.Cells(Rows.Count, 1).End(xlUp).Row
'     ultimaColuna = wsCronograma.Cells(51, Columns.Count).End(xlToLeft).Column

'     ' Atualiza as fórmulas de rastreamento
'     For i = 52 To ultimaLinha Step 2  ' Pegando apenas as linhas que correspondem ao MEMORIAL ORÇ
'         wsCronograma.Cells(i, 8).Formula = "=IFERROR(INDEX('MEMORIAL ORÇ'!H:H, " & (i / 2) & "), """")"
'     Next i

'     MsgBox "Fórmulas de rastreamento atualizadas no CRONOGRAMA!", vbInformation
' End Sub

Sub atualizarFormulasCronograma()
    Dim wsCronograma As Worksheet
    Dim wsMemorial As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaColuna As Long
    Dim i As Long, j As Long
    Dim colBase As Long

    Set wsCronograma = ThisWorkbook.Sheets("CRONOGRAMA")
    Set wsMemorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")

    ' Encontra a última linha preenchida no cronograma
    ultimaLinha = wsCronograma.Cells(Rows.Count, 1).End(xlUp).Row
    ultimaColuna = wsCronograma.Cells(51, Columns.Count).End(xlToLeft).Column

    ' Define a coluna base no MEMORIAL ORÇ (coluna onde começa 15 DIAS)
    colBase = 8 ' Supondo que "15 DIAS" começa na coluna H (8ª coluna)

    ' Atualiza as fórmulas de rastreamento para cada coluna de prazo (15 DIAS, 30 DIAS, 45 DIAS...)
    For j = colBase To ultimaColuna
        For i = 52 To ultimaLinha Step 2  ' Pegando apenas as linhas correspondentes no MEMORIAL ORÇ
            wsCronograma.Cells(i, j).Formula = "=IFERROR(INDEX('MEMORIAL ORÇ'!" & Cells(1, j).Address(False, False) & ":" & Cells(Rows.Count, j).Address(False, False) & ", " & (i / 2) & "), """")"
        Next i
    Next j

    MsgBox "Fórmulas de rastreamento atualizadas no CRONOGRAMA!", vbInformation
End Sub
