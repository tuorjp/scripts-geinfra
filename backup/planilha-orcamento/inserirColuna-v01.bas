Sub inserirColunaCronograma()
    'Salvar arquivo antes de realizar alterações
    ThisWorkbook.Save

    'Desativa a execução eventos excel, atualização de tela e o cálculo automático
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    'Declaração da planilha ativa
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CRONOGRAMA")

    'Declaração das variáveis
    Dim ultimaColuna As Integer
    Dim contador As Integer
    Dim primeiraColunaDias As Integer
    Dim i As Integer
    Dim novoValor As Integer
    contador = 0
    primeiraColunaDias = 0 'Para armazenar a posição da primeira ocorrência de "DIAS"

    'Encontra a última coluna preenchida na linha 51
    ultimaColuna = ws.Cells(51, ws.Columns.Count).End(xlToLeft).Column
    
    'Identifica a primeira ocorrência de "DIAS" e conta quantos existem a partir da coluna G
    For i = 7 To ultimaColuna
        Dim valorCelula As String
        'Converte o valor da célula para String, por que pode ser Empty
        valorCelula = CStr(ws.Cells(51, i).Value)
        
        If InStr(1, valorCelula, "DIAS", vbTextCompare) > 0 Then
            If primeiraColunaDias = 0 Then primeiraColunaDias = i
            contador = contador + 1
        End If
    Next i

    'Se não encontrou "DIAS", não faz nada
    If primeiraColunaDias = 0 Then
        MsgBox "Nenhuma coluna 'DIAS' encontrada.", vbExclamation, "Erro"
        Exit Sub
    End If

    'Insere as novas colunas antes da primeira coluna "DIAS"
    ws.Columns(primeiraColunaDias - 1).Insert Shift:=xlToRight
    ws.Columns(primeiraColunaDias - 1).Insert Shift:=xlToRight

    'ws.Range(ws.Cells(1, primeiraColunaDias - 1), ws.Cells(ws.Rows.Count, primeiraColunaDias)).MergeCells = False

    Dim ultimaLinhaCronograma As Range
    Set ultimaLinhaCronograma = ws.Range("G:G").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    
    'Tirar o merge para não dar problema na hora de colar
    ws.Range("Q51:" & "R" & ultimaLinhaCronograma.Row - 1).MergeCells = False

    'Copia até a última linha preenchida
    ws.Range("E51:F" & ultimaLinhaCronograma.Row - 1).Copy
    ws.Cells(51, primeiraColunaDias - 1).PasteSpecial Paste:=xlPasteAll

    'Atualiza os valores das colunas "DIAS"
    novoValor = 15 'Sempre começa com 15

    For i = primeiraColunaDias To ws.Cells(51, ws.Columns.Count).End(xlToRight).Column
        If InStr(1, ws.Cells(51, i).Value, "DIAS", vbTextCompare) > 0 Then
            ws.Cells(51, i).Value = novoValor & " DIAS"
            novoValor = novoValor + 15 'Incrementa de 15 em 15
        End If
    Next i

    'Inserir uma coluna em memorial para cada duas no cronograma
    '!VERIFICAR SE SERÁ POSSÍVEL VINCULAR MEMORIAL AO CRONOGRAMA AUTOMATICAMENTE
    Dim wsMemorial As Worksheet
    Set wsMemorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")

    Dim colunaDestino As Long
    Dim valorExistente As Boolean
    Dim novaCelula As Range
    colunaDestino = 8

    For Each celula In ws.Range(ws.Cells(51, 1), ws.Cells(51, ultimaColuna))
        If InStr(1, CStr(celula.Value), "DIAS", vbTextCompare) > 0 Then
            valor = Trim(celula.Value)

            'Verifica se valor existe em memorial
            valorExistente = False
            For Each destino In wsMemorial.Range(wsMemorial.Cells(25, 8), wsMemorial.Cells(25, wsMemorial.Columns.Count))
                If Trim(destino.value) = valor Then
                    valorExistente = True
                    Exit For
                End IF
            Next destino

            ' Se não existir, adicionar na próxima linha disponível na coluna H
            If Not valorExistente Then
                If colunaDestino = 8 Then
                    wsMemorial.Cells(25, colunaDestino).Value = valor
                End If
                If colunaDestino > 8 Then
                    wsMemorial.Columns(colunaDestino).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                    wsMemorial.Cells(25, colunaDestino).Value = valor
                End If
                 wsMemorial.Columns(colunaDestino).ColumnWidth = 20
                 Set novaCelula = wsMemorial.Cells(25, colunaDestino)
                 novaCelula.Interior.Color = RGB(0, 32, 96)
                 novaCelula.Font.Color = RGB(255, 255, 255)
                 
                 colunaDestino = colunaDestino + 1
            End If
        End If
    next celula

    'Reativa as configurações desativadas
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
End Sub
