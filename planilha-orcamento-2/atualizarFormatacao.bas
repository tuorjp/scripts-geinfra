Sub atualizarFormatacao()
    Dim memorial As Worksheet
    Set memorial = ThisWorkbook.Sheets("MEMORIAL ORÇ")

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

    ' Atualiza a formatação da célula A6 conforme o tipo selecionado
    With memorial.Range("A6")
        If tipoValor = "quantidade" Then
            .NumberFormat = "0.00" ' Formato numérico normal
        Else
            .NumberFormat = "0.00%" ' Formato percentual
        End If
    End With

    ' Encontra os limites de colunas no Memorial
    Dim primeiraColunaMemorial As Integer
    Dim ultimaColunaMemorial As Integer
    primeiraColunaMemorial = 9 ' Começa depois da coluna 8

    Dim colunaMemorial As Integer
    For colunaMemorial = primeiraColunaMemorial To memorial.Cells(25, memorial.Columns.Count).End(xlToLeft).Column
        If memorial.Cells(25, colunaMemorial).Value = "DESCRIÇÃO - MEMORIAL DE CALCULO" Then
            ultimaColunaMemorial = colunaMemorial - 1 ' Pegamos a anterior
            Exit For
        End If
    Next colunaMemorial

    ' Encontra a última linha válida no Memorial (antes da linha "LAST ROW")
    Dim ultimaLinha As Range
    Dim ultimaLinhaMemorial As Integer
    Set ultimaLinha = memorial.Range("B:B").Find("LAST ROW", LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If Not ultimaLinha Is Nothing Then
        ultimaLinhaMemorial = ultimaLinha.Row - 1
    Else
        MsgBox "Erro: 'LAST ROW' não encontrada no Memorial!", vbExclamation
        Exit Sub
    End If

    ' Aplica formatação ao intervalo do Memorial (Linha 6 e Linhas de Dados)
    Dim intervaloMemorialA As Range
    Dim intervaloMemorial As Range
    Dim intervaloLinha6 As Range

    Set intervaloMemorialA = memorial.Range("A28:A" & ultimaLinhaMemorial)
    Set intervaloMemorial = memorial.Range(memorial.Cells(28, primeiraColunaMemorial), memorial.Cells(ultimaLinhaMemorial, ultimaColunaMemorial))
    Set intervaloLinha6 = memorial.Range(memorial.Cells(6, primeiraColunaMemorial), memorial.Cells(6, ultimaColunaMemorial))

    If tipoValor = "quantidade" Then
        intervaloMemorial.NumberFormat = "0.00"
        intervaloMemorialA.NumberFormat = "0.00"
        intervaloLinha6.NumberFormat = "0.00"
    Else
        intervaloMemorial.NumberFormat = "0.00%"
        intervaloMemorialA.NumberFormat = "0.00%"
        intervaloLinha6.NumberFormat = "0.00%"
    End If

End Sub
