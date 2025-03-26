Sub adicionarColuna()
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
    Dim rangeFormatadaCronograma As Range
    Dim quantidadeDeColunasInserir As Integer
    
    Set memorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")
    Set cronograma = ThisWorkbook.Sheets("CRONOGRAMA")

    primeiraLinhaMemorial = 27
    ultimaColunaMemorial = memorial.Cells(1, memorial.Columns.Count).End(xlToLeft).Column
    
    'Variável que vai receber a célula que contém a string LAST ROW
    Set ultimaLinhaCronograma = cronograma.Range("G:G").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If ultimaLinhaCronograma Is Nothing Then Err.Raise vbObjectError + 1, , "Não foi possível encontrar a última linha do cronograma."

    'Células que contém a formatação usada nas colunas
    Set rangeFormatadaCronograma = cronograma.Range("E51:F" & ultimaLinhaCronograma.Row - 1)
    rangeFormatadaCronograma.MergeCells = False

    ultimaColunaCronograma = cronograma.Cells(1, cronograma.Columns.Count).End(xlToLeft).Column
    primeiraColunaMemorial = 8
    primeiraLinhaCronograma = 54
    primeiraColunaCronograma = 16

    quantidadeDeColunasInserir = Application.InputBox(Prompt:="Número de colunas a inserir:", Title:="Colunas", Type:=1)
    If quantidadeDeColunasInserir <= 0 Then Exit Sub

    '=== Encontra a coluna DESCRIÇÃO - MEMORIAL DE CALCULO no MEMORIAL ===
    Dim colunaNaoApagarMemorial As Integer
    colunaNaoApagarMemorial = 0

    Dim ultimaColunaMemorialDetectada As Integer
    ultimaColunaMemorialDetectada = memorial.Cells(25, memorial.Columns.Count).End(xlToLeft).Column ' Detecta a última coluna preenchida

    For i = 1 To ultimaColunaMemorialDetectada
        'Variável que recebe a célula que contém a string "NÃO APAGAR"
        Dim valorCelulaMemorial As String
        
        If memorial.Cells(25, i).MergeCells Then
            valorCelulaMemorial = Trim(CStr(memorial.Cells(25, i).MergeArea.Cells(1, 1).Value))
        Else
            valorCelulaMemorial = Trim(CStr(memorial.Cells(25, i).Value))
        End If
        
        If StrComp(valorCelulaMemorial, "NÃO APAGAR", vbTextCompare) = 0 Then
            'Índice da string "NÃO APAGAR"
            colunaNaoApagarMemorial = i
            Exit For
        End If
    Next i

    If colunaNaoApagarMemorial = 0 Then Err.Raise vbObjectError + 2, , "Não foi encontrada a coluna 'NÃO APAGAR' no MEMORIAL."
    'Encontra a coluna DESCRIÇÃO MEMORIAL que é a última coluna
    'antes dela é que se insere ou se apaga uma coluna
    colunaDescMemorialDeCalc = colunaNaoApagarMemorial - 3

    'Insere colunas no MEMORIAL
    For i = 1 To quantidadeDeColunasInserir
        'Insere a coluna nova
        memorial.Columns(colunaDescMemorialDeCalc).Insert Shift:=xlToRight
        'Copia a formatação da coluna padrão
        memorial.Range("A:A").Copy Destination:=memorial.Cells(51, colunaDescMemorialDeCalc).EntireColumn
    Next i

    For i = 1 To quantidadeDeColunasInserir

        '=== Encontra a coluna TOTAL COM BDI no CRONOGRAMA ===
        Dim colunaNaoApagarCronograma As Integer
        colunaNaoApagarCronograma = 0

        Dim ultimaColunaCronogramaDetectada As Integer
        ultimaColunaCronogramaDetectada = cronograma.Cells(51, cronograma.Columns.Count).End(xlToLeft).Column ' Detecta a última coluna preenchida

        For j = 1 To ultimaColunaCronogramaDetectada
            Dim valorCelulaCronograma As String
            
            If cronograma.Cells(51, j).MergeCells Then
                valorCelulaCronograma = Trim(CStr(cronograma.Cells(51, j).MergeArea.Cells(1, 1).Value))
            Else
                valorCelulaCronograma = Trim(CStr(cronograma.Cells(51, j).Value))
            End If
            
            If StrComp(valorCelulaCronograma, "NÃO APAGAR", vbTextCompare) = 0 Then
                colunaNaoApagarCronograma = j
                Exit For
            End If
        Next j

        If colunaNaoApagarCronograma = 0 Then Err.Raise vbObjectError + 3, , "Não foi encontrada a coluna 'NÃO APAGAR' no CRONOGRAMA."
        colunaCronogramaTotalComBDI = colunaNaoApagarCronograma - 3

        cronograma.Columns(colunaCronogramaTotalComBDI).Insert Shift:=xlToRight
        cronograma.Columns(colunaCronogramaTotalComBDI).Insert Shift:=xlToRight
        
        'Copia o conteúdo e formatação
        rangeFormatadaCronograma.Copy
        cronograma.Cells(51, colunaCronogramaTotalComBDI).PasteSpecial Paste:=xlPasteAll

        'Copia o tamanho das colunas
        cronograma.Columns(rangeFormatadaCronograma.Column).Copy
        cronograma.Columns(colunaCronogramaTotalComBDI).PasteSpecial Paste:=xlPasteColumnWidths
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
