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

            Debug.Print " "
            Debug.Print "Linha MEMORIAL: " & linhaMemorial & ", Coluna MEMORIAL: " & colunaMemorial
            Debug.Print "Linha CRONOGRAMA: " & linhaCronograma & ", Coluna CRONOGRAMA: " & colunaCronograma

            ' Verifica se a linha no Memorial é válida
            If IsNumeric(linhaMemorial) And linhaMemorial >= 28 And linhaMemorial <= ultimaLinhaMemorial Then
                ' Verifica se há um valor válido
                If Trim(CStr(memorial.Cells(linhaMemorial, colunaMemorial).Value)) <> "" Then
                    Debug.Print "Linha MEMORIAL LOOP: " & linhaMemorial & ", Coluna MEMORIAL: " & colunaMemorial
                    Debug.Print "Linha CRONOGRAMA LOOP: " & linhaCronograma & ", Coluna CRONOGRAMA: " & colunaCronograma
                    Debug.Print "Inserindo fórmula em cronograma: " & linhaCronograma & ":" & colunaCronograma

                    ' Insere a fórmula apenas na primeira linha do bloco
                    cronograma.Cells(linhaCronograma, colunaCronograma).Formula = _
                        "='MEMORIAL ORÇ'!" & memorial.Cells(linhaMemorial, colunaMemorial).Address(False, False)
                End If
            End If
        Next linhaCronograma
    Next colunaMemorial

    MsgBox "Rastreamento concluído!"
End Sub
