' Insere fórmulas de rastreamento no cronograma com base nos valores do memorial
Sub insereFormulaRastreamento()
    Dim memorial As Worksheet
    Dim cronograma As Worksheet

    Dim linhaCronograma As Integer
    Dim colunaMemorial As Integer
    Dim colunaCronograma As Integer

    Dim ultimaLinhaMemorial As Integer
    Dim ultimaLinhaCronograma As Integer
    Dim primeiraColunaMemorial As Integer
    Dim ultimaColunaMemorial As Integer
    Dim ultimaColunaCronograma As Integer

    Set memorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")
    Set cronograma = ThisWorkbook.Sheets("CRONOGRAMA")

    ' Obtém o tipo de entrada do usuário da ComboBox
    Dim cmb As Object
    Set cmb = memorial.OLEObjects("cmbTipoValor").Object
    Dim tipoValor As String
    tipoValor = Trim(LCase(cmb.Value))

    ' Se o tipo não for válido, exibe erro e encerra
    If tipoValor <> "quantidade" And tipoValor <> "porcentagem" Then
        MsgBox "Erro: Escolha 'QUANTIDADE' ou 'PORCENTAGEM' na ComboBox!", vbExclamation
        Exit Sub
    End If

    ' Encontra a última linha válida no Memorial (antes da linha "LAST ROW")
    Dim ultimaLinha As Range
    Set ultimaLinha = memorial.Range("B:B").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If Not ultimaLinha Is Nothing Then
        ultimaLinhaMemorial = ultimaLinha.Row - 1
    Else
        MsgBox "Erro: 'LAST ROW' não encontrada no Memorial!", vbExclamation
        Exit Sub
    End If

    ' Encontra a última linha válida no Cronograma (antes da linha "LAST ROW")
    Set ultimaLinha = cronograma.Range("G:G").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If Not ultimaLinha Is Nothing Then
        ultimaLinhaCronograma = ultimaLinha.Row - 1
    Else
        MsgBox "Erro: 'LAST ROW' não encontrada no Cronograma!", vbExclamation
        Exit Sub
    End If

    ' Encontra os limites de colunas no Memorial
    primeiraColunaMemorial = 9 ' Começa depois da coluna 8
    For colunaMemorial = primeiraColunaMemorial To memorial.Cells(25, memorial.Columns.Count).End(xlToLeft).Column
        If memorial.Cells(25, colunaMemorial).Value = "DESCRIÇÃO - MEMORIAL DE CALCULO" Then
            ultimaColunaMemorial = colunaMemorial - 1 ' Pegamos a anterior
            Exit For
        End If
    Next colunaMemorial

    ' Encontra a última coluna válida no Cronograma (antes de "NÃO APAGAR" na linha 51)
    Dim ultimaColuna As Range
    Set ultimaColuna = cronograma.Rows(51).Find("NÃO APAGAR", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns)
    If Not ultimaColuna Is Nothing Then
        ultimaColunaCronograma = ultimaColuna.Column - 5 ' Pega 5 colunas antes
    Else
        MsgBox "Erro: 'NÃO APAGAR' não encontrada na linha 51 do Cronograma!", vbExclamation
        Exit Sub
    End If

    ' Loop pelas colunas de interesse no Memorial
    For colunaMemorial = primeiraColunaMemorial To ultimaColunaMemorial
        colunaCronograma = (colunaMemorial - primeiraColunaMemorial) * 2 + 17 ' Ajusta a posição inicial

        ' Garante que não ultrapasse o limite das colunas no Cronograma
        If colunaCronograma > ultimaColunaCronograma Then Exit For

        ' Loop pelas linhas do Cronograma
        For linhaCronograma = 55 To ultimaLinhaCronograma Step 2
            Dim linhaMemorial As Integer
            If cronograma.Cells(linhaCronograma, 8).MergeCells Then
                linhaMemorial = cronograma.Cells(linhaCronograma, 8).MergeArea.Cells(1, 1).Value
            Else
                linhaMemorial = cronograma.Cells(linhaCronograma, 8).Value
            End If

            ' Verifica se a linha no Memorial é válida
            If IsNumeric(linhaMemorial) And linhaMemorial >= 28 And linhaMemorial <= ultimaLinhaMemorial Then
                ' Obtém o valor digitado no Memorial
                Dim valorMemorial As Double
                valorMemorial = memorial.Cells(linhaMemorial, colunaMemorial).Value

                ' Verifica se há um valor válido
                If Trim(CStr(valorMemorial)) <> "" And valorMemorial <> 0 Then
                    Dim valorFinal As Double

                    ' Se o tipo for quantidade, divide pelo valor da coluna H (QTD) para obter a porcentagem
                    If tipoValor = "quantidade" Then
                        Dim valorQtd As Double
                        valorQtd = memorial.Cells(linhaMemorial, 8).Value ' Coluna H

                        ' Evita erro de divisão por zero
                        If valorQtd <> 0 Then
                            valorFinal = valorMemorial / valorQtd
                        Else
                            valorFinal = 0 ' Define 0 caso o valor em QTD seja 0
                        End If
                    Else
                        ' Se for porcentagem, mantém o valor original
                        valorFinal = valorMemorial
                    End If

                    ' Exibe os valores no Debug
                    Debug.Print "Linha MEMORIAL: " & linhaMemorial & ", Coluna MEMORIAL: " & colunaMemorial
                    Debug.Print "Linha CRONOGRAMA: " & linhaCronograma & ", Coluna CRONOGRAMA: " & colunaCronograma
                    Debug.Print "Valor original: " & valorMemorial & " | Valor final: " & valorFinal

                    ' Insere o valor no Cronograma como porcentagem (se for quantidade)
                    If tipoValor = "quantidade" Then
                        cronograma.Cells(linhaCronograma, colunaCronograma).Value = valorFinal
                        cronograma.Cells(linhaCronograma, colunaCronograma).NumberFormat = "0.00%" ' Exibe como porcentagem
                    Else
                        cronograma.Cells(linhaCronograma, colunaCronograma).Formula = _
                            "='MEMORIAL ORÇ'!" & memorial.Cells(linhaMemorial, colunaMemorial).Address(False, False)
                    End If
                End If
            End If
        Next linhaCronograma
    Next colunaMemorial

End Sub
