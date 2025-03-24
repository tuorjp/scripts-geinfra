Sub adicionarColuna()
    On Error GoTo TratarErro

    ThisWorkbook.Save
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim memorial As Worksheet
    Dim cronograma As Worksheet
    Dim primeiraColunaMemorial As Integer
    Dim ultimaColunaMemorial As Integer
    Dim primeiraLinhaMemorial As Integer
    Dim primeiraLinhaCronograma As Integer
    Dim primeiraColunaCronograma As Integer
    Dim rangeFormatadaCronograma As Range
    Dim ultimaLinhaCronograma As Range
    Dim ultimaColunaCronograma As Integer
    Dim colunaCronogramaTotalComBDI As Integer
    Dim quantidadeDeColunasInserir As Integer
    Dim colunaDescMemorialDeCalc As Integer

    Set memorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")
    Set cronograma = ThisWorkbook.Sheets("CRONOGRAMA")

    primeiraLinhaMemorial = 27
    ultimaColunaMemorial = memorial.Cells(1, memorial.Columns.Count).End(xlToLeft).Column
    Set ultimaLinhaCronograma = cronograma.Range("G:G").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If ultimaLinhaCronograma Is Nothing Then Err.Raise vbObjectError + 1, , "Não foi possível encontrar a última linha do cronograma."

    Set rangeFormatadaCronograma = cronograma.Range("Q51:R" & ultimaLinhaCronograma.Row - 1)
    rangeFormatadaCronograma.MergeCells = False

    ultimaColunaCronograma = cronograma.Cells(1, cronograma.Columns.Count).End(xlToLeft).Column
    primeiraColunaMemorial = 8
    primeiraLinhaCronograma = 54
    primeiraColunaCronograma = 16

    quantidadeDeColunasInserir = Application.InputBox(Prompt:="Número de colunas a inserir:", Title:="Colunas", Type:=1)
    If quantidadeDeColunasInserir <= 0 Then Exit Sub

    'Encontra a coluna DESCRIÇÃO - MEMORIAL DE CALCULO
    For i = 1 To ultimaColunaMemorial
         If memorial.Cells(25, i).MergeCells Then
            If memorial.Cells(25, i).MergeArea.Cells(1, 1).Value = "DESCRIÇÃO - MEMORIAL DE CALCULO" Then
                colunaDescMemorialDeCalc = i
                Exit For
            End If
        Else
            If memorial.Cells(25, i).Value = "DESCRIÇÃO - MEMORIAL DE CALCULO" Then
                colunaDescMemorialDeCalc = i
                Exit For
            End If
        End If
    Next i

    If colunaDescMemorialDeCalc = 0 Then Err.Raise vbObjectError + 2, , "Não foi encontrada a coluna 'DESCRIÇÃO - MEMORIAL DE CALCULO'."

    'Encontra a coluna TOTAL COM BDI
    For i = 1 To ultimaColunaCronograma
        If cronograma.Cells(25, i).MergeCells Then
            If cronograma.Cells(25, i).MergeArea.Cells(1, 1).Value = "TOTAL COM" Then
                colunaCronogramaTotalComBDI = i
                Exit For
            End If
        Else
            If cronograma.Cells(25, i).Value = "TOTAL COM" Then
                colunaCronogramaTotalComBDI = i
                Exit For
            End If
        End If
    Next i

    If colunaCronogramaTotalComBDI = 0 Then Err.Raise vbObjectError + 3, , "Não foi encontrada a coluna 'TOTAL COM'."

    'Insere colunas no MEMORIAL
    For i = 1 To quantidadeDeColunasInserir
        memorial.Columns(colunaDescMemorialDeCalc - 1).Insert Shift:=xlToRight
        memorial.Range("A:A").Copy Destination:=memorial.Cells(51, colunaDescMemorialDeCalc - 1).EntireColumn
    Next i

    'Insere colunas no CRONOGRAMA
    For i = 1 To quantidadeDeColunasInserir
        cronograma.Columns(colunaCronogramaTotalComBDI - 1).Insert Shift:=xlToRight
        cronograma.Columns(colunaCronogramaTotalComBDI - 2).Insert Shift:=xlToRight
        rangeFormatadaCronograma.Copy Destination:=cronograma.Cells(1, colunaCronogramaTotalComBDI - 1)
    Next i

Finalizar:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    Exit Sub

TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro no Procedimento"
    Resume Finalizar
End Sub
