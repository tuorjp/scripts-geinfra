Sub removerColunas()
    On Error GoTo TratarErro

    'ThisWorkbook.Save
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim memorial As Worksheet
    Dim cronograma As Worksheet

    Dim primeiraColunaMemorial As Integer
    Dim primeiraColunaCronograma As Integer
    Dim ultimaColunaMemorial As Integer
    Dim ultimaColunaCronograma As Integer
    
    Dim primeiraLinhaMemorial As Integer
    Dim primeiraLinhaCronograma As Integer
    Dim ultimaLinhaCronograma As Range
    
    Dim colunaCronogramaTotalComBDI As Integer
    Dim colunaDescMemorialDeCalc As Integer
    Dim quantidadeDeColunasApagar As Integer

    Set memorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")
    Set cronograma = ThisWorkbook.Sheets("CRONOGRAMA")

    primeiraLinhaMemorial = 27
    ultimaColunaMemorial = memorial.Cells(1, memorial.Columns.Count).End(xlToLeft).Column
    Set ultimaLinhaCronograma = cronograma.Range("G:G").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If ultimaLinhaCronograma Is Nothing Then Err.Raise vbObjectError + 1, , "Não foi possível encontrar a última linha do cronograma."

    ultimaColunaCronograma = cronograma.Cells(1, cronograma.Columns.Count).End(xlToLeft).Column
    primeiraColunaMemorial = 8
    primeiraLinhaCronograma = 54
    primeiraColunaCronograma = 16

    quantidadeDeColunasApagar = Application.InputBox(Prompt:="Número de colunas a inserir:", Title:="Colunas", Type:=1)
    If quantidadeDeColunasApagar <= 0 Then Exit Sub

    Dim colunaNaoApagarMemorial As Integer
    colunaNaoApagarMemorial = 0

    Dim ultimaColunaMemorialDetectada As Integer

    '=== Encontra a coluna "DESCRIÇÃO - MEMORIAL DE CALCULO" no MEMORIAL ===
    'Apaga colunas no MEMORIAL
    For i = 1 To quantidadeDeColunasApagar
        ultimaColunaMemorialDetectada = memorial.Cells(25, memorial.Columns.Count).End(xlToLeft).Column ' Detecta a última coluna preenchida

        For k = 1 To ultimaColunaMemorialDetectada
            Dim valorCelulaMemorial As String
            
            If memorial.Cells(25, k).MergeCells Then
                valorCelulaMemorial = Trim(CStr(memorial.Cells(25, k).MergeArea.Cells(1, 1).Value))
            Else
                valorCelulaMemorial = Trim(CStr(memorial.Cells(25, k).Value))
            End If
            
            If StrComp(valorCelulaMemorial, "NÃO APAGAR", vbTextCompare) = 0 Then
                colunaNaoApagarMemorial = k
                Exit For
            End If
        Next k

        If colunaNaoApagarMemorial = 0 Then Err.Raise vbObjectError + 2, , "Não foi encontrada a coluna 'NÃO APAGAR' no MEMORIAL."
        colunaDescMemorialDeCalc = colunaNaoApagarMemorial - 3

        'Verificando se há algo a excluir
        Dim verificarValorDaCelulaASerExcluida As String
        verificarValorDaCelulaASerExcluida = Trim(CStr(memorial.Cells(25, i).Value))
        
        If StrComp(verificarValorDaCelulaASerExcluida, "QTD", vbTextCompare) = 0 Then
            Exit Sub
        End If

        memorial.Columns(colunaDescMemorialDeCalc - 1).Delete Shift:=xlToLeft
    Next i

    Dim colunaNaoApagarCronograma As Integer
    colunaNaoApagarCronograma = 0

    For i = 1 To quantidadeDeColunasApagar
        '=== Encontra a coluna "NÃO APAGAR" e "TOTAL COM BDI no CRONOGRAMA" ===
        Dim ultimaColunaCronogramaDetectada As Integer
        ultimaColunaCronogramaDetectada = cronograma.Cells(51, cronograma.Columns.Count).End(xlToLeft).Column ' Detecta a última coluna preenchida

        'Percorre a planilha para buscar um valor, e retorna qual célula tem o valor 
        For j = 1 To ultimaColunaCronogramaDetectada
            'Variável que armazena o valor das células iteradas
            Dim valorCelulaCronograma As String
            
            If cronograma.Cells(51, j).MergeCells Then
                valorCelulaCronograma = Trim(CStr(cronograma.Cells(51, j).MergeArea.Cells(1, 1).Value))
            Else
                valorCelulaCronograma = Trim(CStr(cronograma.Cells(51, j).Value))
            End If
            
            If StrComp(valorCelulaCronograma, "NÃO APAGAR", vbTextCompare) = 0 Then
                'Se a célula conter a string, sua localização é armazenado
                colunaNaoApagarCronograma = j
                Exit For
            End If
        Next j

        If colunaNaoApagarCronograma = 0 Then Err.Raise vbObjectError + 3, , "Não foi encontrada a coluna 'NÃO APAGAR' no CRONOGRAMA."
        colunaCronogramaTotalComBDI = colunaNaoApagarCronograma - 3

        'Deleta as últimas colunas (quantidadeDeColunasApagar) tendo como base a
        'colunaCronogramaTotalComBDI para achar as colunas a serem excluídas
        cronograma.Columns(colunaCronogramaTotalComBDI - 1).Delete Shift:=xlToLeft
        cronograma.Columns(colunaCronogramaTotalComBDI - 2).Delete Shift:=xlToLeft
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
