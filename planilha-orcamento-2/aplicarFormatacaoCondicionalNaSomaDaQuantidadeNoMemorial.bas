Sub aplicarFormatacaoCondicionalNaSomaDaQuantidadeNoMemorial()
    Dim ws As Worksheet
    Dim cmb As Object
    Dim tipoValor As String
    Dim ultimaLinha As Integer
    Dim ultimaColuna As Integer
    Dim intervaloH As Range
    Dim cel As Range

    ' Definir a planilha
    Set ws = ThisWorkbook.Sheets("MEMORIAL ORÇ")

    ' Capturar valor da ComboBox
    Set cmb = ws.OLEObjects("cmbTipoValor").Object
    tipoValor = Trim(LCase(cmb.Value))

    ' Validar seleção
    If tipoValor <> "quantidade" And tipoValor <> "porcentagem" Then
        MsgBox "Erro: Escolha 'QUANTIDADE' ou 'PORCENTAGEM' na ComboBox!", vbExclamation
        Exit Sub
    End If

    ' Encontrar a última linha antes de "LAST ROW" na coluna B
    Dim ultimaCelula As Range
    Set ultimaCelula = ws.Range("B:B").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    
    If Not ultimaCelula Is Nothing Then
        ultimaLinha = ultimaCelula.Row - 1
    Else
        MsgBox "Erro: 'LAST ROW' não encontrada!", vbExclamation
        Exit Sub
    End If

    ' Encontrar a última coluna antes de "DESCRIÇÃO - MEMORIAL DE CALCULO" na linha 25
    Dim ultimaCelulaColuna As Range
    Set ultimaCelulaColuna = ws.Rows(25).Find("DESCRIÇÃO - MEMORIAL DE CALCULO", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns)

    If Not ultimaCelulaColuna Is Nothing Then
        ultimaColuna = ultimaCelulaColuna.Column - 1
    Else
        MsgBox "Erro: 'DESCRIÇÃO - MEMORIAL DE CALCULO' não encontrada na linha 25!", vbExclamation
        Exit Sub
    End If

    ' Definir o intervalo da coluna H (H28 até última linha antes de LAST ROW)
    Set intervaloH = ws.Range("H28:H" & ultimaLinha)

    ' Remover formatação condicional existente
    intervaloH.FormatConditions.Delete

    ' Se for QUANTIDADE, aplicar nova formatação condicional
    If tipoValor = "quantidade" Then
        For Each cel In intervaloH
             Dim formulaCondicional As String
            Dim enderecoSoma As String
            Dim enderecoH As String

            ' Obter o intervalo correto para soma
            enderecoSoma = ws.Range(cel.Offset(0, 1), ws.Cells(cel.Row, ultimaColuna)).Address(0, 0)

            ' Obter a célula da coluna H
            enderecoH = cel.Address(0, 0)

            ' Criar a fórmula no formato correto
            'formulaCondicional = "=SUM(" & enderecoSoma & ")<>" & enderecoH
            'formulaCondicional = Replace("=SUM(" & enderecoSoma & ")<>" & enderecoH, ",", Application.International(xlListSeparator))
            formulaCondicional = "=SOMA(" & enderecoSoma & ")<>" & enderecoH

            Debug.Print "Linha: " & cel.Row & " | Fórmula: " & formulaCondicional

            ' Aplicar a formatação condicional corretamente
            'With cel.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaCondicional)
            With cel.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaCondicional, ",", Application.DecimalSeparator))
                .Font.Color = RGB(255, 0, 0) ' Vermelho
                .Font.Bold = True
            End With
            Debug.Print "Cor " & cel.Font.Color
        Next cel
    End If

End Sub
