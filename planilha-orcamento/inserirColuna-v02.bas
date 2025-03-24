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

    'Reativa as configurações desativadas
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
End Sub
