Sub insereFormulaRastreamento()
    Dim memorial As Worksheet
    Dim cronograma As Worksheet

    Dim linhaCronograma As Integer
    Dim colunaMemorial As Integer
    Dim colunaCronograma As Integer

    Dim ultimaLinhaMemorial As Integer
    Dim ultimaLinhaCronograma As Integer
    Dim colunaQTD As Integer, colunaDescricaoMemorial As Integer

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

    ' Encontra as colunas de interesse no Memorial
    For colunaMemorial = 1 To memorial.Cells(25, memorial.Columns.Count).End(xlToLeft).Column
        Dim valorCelula As String

        If memorial.Cells(25, colunaMemorial).MergeCells Then
            valorCelula = Trim(CStr(memorial.Cells(25, colunaMemorial).MergeArea.Cells(1, 1).Value))
        Else
            valorCelula = Trim(CStr(memorial.Cells(25, colunaMemorial).Value))
        End If

        If valorCelula = "QTD" Then
            colunaQTD = colunaMemorial
        ElseIf valorCelula = "DESCRIÇÃO - MEMORIAL DE CALCULO" Then
            colunaDescricaoMemorial = colunaMemorial
        End If
    Next colunaMemorial

    ' Loop pelas colunas de interesse
    For colunaMemorial = colunaQTD To colunaDescricaoMemorial
        colunaCronograma = (colunaMemorial - colunaQTD) * 2 + 17  ' Inicia na coluna Q (17)
        For linhaCronograma = 55 To ultimaLinhaCronograma Step 2

            ' Verifica o número da linha correspondente no Memorial a partir da coluna H (número 8)
            Dim linhaMemorial As Integer
            If cronograma.Cells(linhaCronograma, 8).MergeCells Then
                linhaMemorial = cronograma.Cells(linhaCronograma, 8).MergeArea.Cells(1, 1).Value
            Else
                linhaMemorial = cronograma.Cells(linhaCronograma, 8).Value
            End If

            ' Verifica se há um número válido no cronograma
            If IsNumeric(linhaMemorial) And linhaMemorial >= 28 And linhaMemorial <= ultimaLinhaMemorial Then

                ' Verifica se a célula do Memorial tem valor
                If Trim(CStr(memorial.Cells(linhaMemorial, colunaMemorial).Value)) <> "" Then
                    Debug.Print "Linha MEMORIAL LOOP: " & linhaMemorial & ", Coluna MEMORIAL: " & colunaMemorial
                    Debug.Print "Linha CRONOGRAMA LOOP: " & linhaCronograma & ", Coluna CRONOGRAMA: " & colunaCronograma
                    Debug.Print "Inserindo fórmula em cronograma: " & linhaCronograma & ":" & colunaCronograma
                    Debug.Print "Value: " & linhaMemorial
                    Debug.Print " "

                    ' Insere a fórmula apenas na primeira linha do bloco
                    'cronograma.Cells(linhaCronograma, colunaCronograma).formula = _
                        "='MEMORIAL ORÇ'!" & memorial.Cells(linhaMemorial, colunaMemorial).Address(False, False)
                End If
            End If
        Next linhaCronograma
    Next colunaMemorial

    MsgBox "Rastreamento concluído!"
End Sub
