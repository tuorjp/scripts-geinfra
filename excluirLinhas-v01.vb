Sub excluirLinha()
    Dim inputLinhas As String
    Dim linhasExcluir As Variant
    Dim i As Long, linhaInicial As Long, linhaFinal As Long
    Dim ws As Worksheet, wsMemorial As Worksheet
    
    Set ws = ActiveSheet
    Set wsMemorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")
    Set wsCronograma = ThisWorkbook.Sheets("CRONOGRAMA")
    Set wsTerc = ThisWorkbook.Sheets("SERV. TERCEIRIZAÇÃO")
    Set wsCurvaAbc = ThisWorkbook.Sheets("CURVA ABC_ITENS DE RELEVÂNCIA")
    
    Application.ScreenUpdating = False
    
    'Abre a caixa de diálogo para o usuário inserir as linhas
    inputLinhas = InputBox("Digite as linhas a serem excluídas (ex: 3 para excluir a linha 3 ou 3-10 para excluir o intervalo de 3 a 10)", "Excluir Linhas")
    
    'Verifica se o usuário cancelou ou não digitou nada
    If inputLinhas = "" Then
        
        MsgBox "Nenhuma entrada fornecida. Script Encerrado."
        Exit Sub
    
    End If
    
    'Divide a entrada com base em espaços
    linhasExcluir = Split(inputLinhas, " ")
    
    'Loop para processar cada parte da entrada
    For i = LBound(linhasExcluir) To UBound(linhasExcluir)
    
        'Verifica se é um intervalo (ex: 3-1)
        If InStr(linhasExcluir(i), "-") > 0 Then

            Dim intervalo As Variant
            intervalo = Split(linhasExcluir(i), "-")
            
            If UBound(intervalo) = 1 Then
                
                linhaInicial = CLng(intervalo(0))
                linhaFinal = CLng(intervalo(1))

                'Verifica se a linha a ser excluída é menor ou igual à linha 27
                If linhaFinal <= 27 Then
                    MsgBox "Não é possível excluir linhas acima da linha 27"
                    Exit Sub
                ElseIf linhaInicial <= 27 Then
                    MsgBox "Não é possível excluir linhas acima da linha 27"
                    Exit Sub
                EndIf
                
                'Verifica se o intervalo final é o maior
                If linhaFinal < linhaInicial Then
                    
                    Dim Temp As Long
                    Temp = linhaInicial
                    linhaInicial = linhaFinal
                    linhaFinal = Temp
                    
                End If
                
                'Exclui as linhas no intervalo especificado
                For linha = linhaFinal To linhaInicial Step -1
                    
                    ws.Rows(linha & ":" & linha).Delete
                    wsMemorial.Rows(linha & ":" & linha).Delete
                    wsTerc.Rows(linha & ":" & linha).Delete
                    wsCurvaAbc.Rows(linha & ":" & linha).Delete
                    'A exclusão aqui é feita na mesma linha, por que ao deletar uma linha,
                    'a posterior sobe, assumindo a mesma posição da já excluída 
                    wsCronograma.Rows(2 * linha - 1 & ":" & 2 * linha - 1).Delete
                    wsCronograma.Rows(2 * linha - 1 & ":" & 2 * linha - 1).Delete
                    'wsCronograma.Rows(2 * linha & ":" & 2 * linha).Delete
                    
                    'Copia o modelo de A19:H19 e cola na linha removida
                    ws.Range("A19:H19").Copy
                    ws.Cells(linha, 1).PasteSpecial Paste:=xlPasteAll
                    ws.Cells(linha, 1).EntireRow.Hidden = False
                    
                Next linha
                
            End If
        
        Else
            
            'Se for apenas um número, exclui essa linha
            On Error Resume Next
                If CLng(linhasExcluir(i)) <= 27 Then

                    MsgBox "Não é possível excluir linhas acima da linha 27"
                    Exit Sub

                EndIf

                ws.Rows(CLng(linhasExcluir(i))).Delete
                wsMemorial.Rows(CLng(linhasExcluir(i))).Delete
                wsTerc.Rows(CLng(linhasExcluir(i))).Delete
                wsCurvaAbc.Rows(CLng(linhasExcluir(i))).Delete
                wsCronograma.Rows(2 * CLng(linhasExcluir(i)) - 1).Delete
                wsCronograma.Rows(2 * CLng(linhasExcluir(i))).Delete
                
                'Copia o modelo de A19:H19 e cola na linha removida
                ws.Range("A19:H19").Copy
                ws.Cells(linha, 1).PasteSpecial Paste:=xlPasteAll
                ws.Cells(linha, 1).EntireRow.Hidden = False
                
            On Error GoTo 0
            
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Linhas excluídas com sucesso!"
End Sub
