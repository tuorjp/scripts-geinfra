Sub Formatar()
' GOINFRA_FORMATAR_COMPOSIÇÃO Macro

Dim numero As Variant

numero = Application.InputBox("1 - GOINFRA DES" & Chr(10) & "2 - GOINFRA ONE" & Chr(10) & "3 - GOINFRA MATERIAL DES" & Chr(10) & "4 - GOINFRA MATERIAL ONE" & Chr(10) & "5 - GOINFRA MAO DE OBRA DES" & Chr(10) & "6 - GOINFRA MAO DE OBRA ONE" & Chr(10) & "7 - SINAPI ANALITICA DES" & Chr(10) & "8 - SINAPI ANALITICA ONE", "Escolha nº Tabela", "Digite Aqui")

 ' Verifica se o usuário cancelou ou não inseriu nada
If numero = False Or numero = "" Then
    MsgBox "É necessário inserir um número para continuar.", vbExclamation, "Erro"
    Exit Sub
End If

' Converte a entrada para número
If Not IsNumeric(numero) Then
    MsgBox "Por favor, insira um número válido.", vbExclamation, "Erro"
    Exit Sub
End If

numero = CInt(numero)

Dim tipo As Integer
tipo = 0
Dim Plan As String
Plan = Plan

Select Case numero
    Case Is = "1"
        Sheets("GOINFRA DES").Select
        tipo = 1
        Plan = "GOINFRA DES"
    Case Is = "2"
        Sheets("GOINFRA ONE").Select
        tipo = 1
        Plan = "GOINFRA ONE"
    Case Is = "3"
        Sheets("GOINFRA MATERIAL DES").Select
        tipo = 2
        Plan = "GOINFRA MATERIAL DES"
    Case Is = "4"
        Sheets("GOINFRA MATERIAL ONE").Select
        tipo = 2
        Plan = "GOINFRA MATERIAL ONE"
    Case Is = "5"
        Sheets("GOINFRA MAO DE OBRA DES").Select
        tipo = 3
        Plan = "GOINFRA MAO DE OBRA DES"
    Case Is = "6"
        Sheets("GOINFRA MAO DE OBRA ONE").Select
        tipo = 3
        Plan = "GOINFRA MAO DE OBRA ONE"
    Case Is = "7"
        Sheets("SINAPI ANALITICA DES").Select
        tipo = 4
        Plan = "SINAPI ANALITICA DES"
    Case Is = "8"
        Sheets("SINAPI ANALITICA ONE").Select
        tipo = 4
        Plan = "SINAPI ANALITICA ONE"
End Select

Cells.Select
Selection.Delete Shift:=xlUp
Cells.Select
ActiveSheet.DrawingObjects.Select
Selection.Delete

Dim arquivo As String
arquivo = Application.GetOpenFilename(, , "Abrir arquivo")

If arquivo = "False" Or arquivo = "" Then
    MsgBox "É necessário selecionar um arquivo para continuar.", vbExclamation, "Erro"
    Exit Sub
End If

Workbooks.Open arquivo

Range("CC1") = ActiveWorkbook.Name
Range("CC1").Select
Dim Fechar As String
Fechar = Range("CC1")
        
Cells.Select
Selection.Copy
Windows("MACRO-DE-FORMATAÇÃO-PLANILHAS-REFERENCIAIS.xlsm") _
.Activate

Range("A1").Select
ActiveSheet.Paste
    
Select Case tipo
    Case Is = "1" 'GOINFRA ANÁLITICA
        Cells.Select
        ActiveSheet.DrawingObjects.Select
        Selection.Delete
        Selection.SpecialCells(xlCellTypeLastCell).Select
        Range(Selection, Cells(1)).Select
        Selection.UnMerge
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlToLeft
        Columns("A:F").Select
        Range("F1").Activate
        Columns("A:F").EntireColumn.AutoFit
        Intersect(Sheets(Plan).Cells, Sheets(Plan).Range("A:F").SpecialCells(xlCellTypeBlanks, xlErrors).EntireRow).Delete
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("A:F").Select
        Range("F1").Activate
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.EntireRow.Delete
    Case Is = "2" 'GOINFRA MATERIAL
        Cells.Select
        ActiveSheet.DrawingObjects.Select
        Selection.Delete
        Selection.SpecialCells(xlCellTypeLastCell).Select
        Range(Selection, Cells(1)).Select
        Selection.UnMerge
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlToLeft
        Columns("A:D").Select
        Range("D1").Activate
        Intersect(Sheets(Plan).Cells, Sheets(Plan).Range("A:D").SpecialCells(xlCellTypeBlanks, xlErrors).EntireRow).Delete
        Cells.Select
        Cells.EntireColumn.AutoFit
        Columns("E:E").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Delete Shift:=xlToLeft
    Case Is = "3" 'GOINFRA MÃO DE OBRA
        Cells.Select
        ActiveSheet.DrawingObjects.Select
        Selection.Delete
        Selection.SpecialCells(xlCellTypeLastCell).Select
        Range(Selection, Cells(1)).Select
        Selection.UnMerge
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlToLeft
        Columns("A:F").Select
        Range("F1").Activate
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Intersect(Sheets(Plan).Cells, Sheets(Plan).Range("A:F").SpecialCells(xlCellTypeBlanks, xlErrors).EntireRow).Delete
        Cells.Select
        Cells.EntireColumn.AutoFit
    Case Is = "4" 'SINAPI ANALITICA
        Rows("1:5").Select
        Range("A5").Activate
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlUp
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
        Columns("A:F").Select
        Selection.Delete Shift:=xlToLeft
        Columns("D:D").Select
        Selection.Delete Shift:=xlToLeft
        Columns("E:L").Select
        Selection.Delete Shift:=xlToLeft
        Columns("F:F").Select
        Selection.Delete Shift:=xlToLeft
        Columns("G:G").Select
        Selection.Delete Shift:=xlToLeft
        Columns("H:H").Select
        Selection.Delete Shift:=xlToLeft
        Columns("I:I").Select
        Selection.Delete Shift:=xlToLeft
        Columns("J:N").Select
        Selection.Delete Shift:=xlToLeft
        Columns("D:D").Select
        Selection.Cut
        Range("J1").Select
        ActiveSheet.Paste
        Columns("D:D").Select
        Selection.Delete Shift:=xlToLeft
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        ' Selection.SpecialCells(xlCellTypeBlanks).Select
        ' Selection.EntireRow.Delete

        Cells.Select
        Cells.EntireColumn.AutoFit
        Columns("B:B").Select
        Selection.ColumnWidth = 30
        ActiveWindow.LargeScroll ToRight:=-1
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("D:D").Select
        Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("E:E").Select
        Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("F:F").Select
        Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("G:G").Select
        Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("H:H").Select
        Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("I:I").Select
        Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        'Range("A1").Select

        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.EntireRow.Delete
End Select
 
Range("A1").Select

Windows(Fechar).Activate
Cells.Select
ActiveWorkbook.Saved = True
ActiveWorkbook.Close

End Sub
