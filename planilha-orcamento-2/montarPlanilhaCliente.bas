Sub montarPlanilhaCliente()
    Dim planilhaOrigem As Workbook
    Dim caminhoPlanilhaDestino As String
    Dim planilhaDestino As Workbook
    Dim abaOrigem As Worksheet
    Dim abaDestino As Worksheet
    Dim nomeAba As String
    Dim i As Integer
    Dim planilha As Worksheet
    Dim planilhaExiste As Boolean
    Dim cell As Range

    ' Define o caminho da planilha de destino
    caminhoPlanilhaDestino = "C:\teste\Arquivo.xlsx"

    ' Define a planilha de origem
    Set planilhaOrigem = ThisWorkbook

    ' Tentar abrir a planilha de destino, se não existir, cria uma nova
    On Error Resume Next
    Set planilhaDestino = Workbooks.Open(caminhoPlanilhaDestino)
    If planilhaDestino Is Nothing Then
        Set planilhaDestino = Workbooks.Add
        planilhaDestino.SaveAs caminhoPlanilhaDestino
    End If
    On Error GoTo 0

    ' Lista de abas que serão copiadas
    Dim abasParaCopiar As Variant
    abasParaCopiar = Array("CAPA", "EST. DE CUSTOS", "MEMORIAL ORÇ", "CRONOGRAMA")

    ' Percorrer cada aba e tomar decisão com base no nome da aba
    For i = LBound(abasParaCopiar) To UBound(abasParaCopiar)
        nomeAba = abasParaCopiar(i)

        ' Verifica se a aba existe na planilha de origem
        On Error Resume Next
        Set abaOrigem = planilhaOrigem.Sheets(nomeAba)
        On Error GoTo 0

        If Not abaOrigem Is Nothing Then
            planilhaExiste = False
            For Each planilha In planilhaDestino.Sheets
                If planilha.Name = nomeAba Then
                    planilhaExiste = True
                    Exit For ' Encerra o loop ao encontrar a aba
                End If
            Next planilha

            ' Se a aba não existir na planilha de destino, cria uma nova
            If Not planilhaExiste Then
                Set abaDestino = planilhaDestino.Worksheets.Add(After:=planilhaDestino.Worksheets(planilhaDestino.Worksheets.Count))
                abaDestino.Name = nomeAba
            Else
                Set abaDestino = planilhaDestino.Sheets(nomeAba)
            End If

            ' --- TRATAR CÉLULAS MESCLADAS ---
            ' Desmescla todas as células da aba de destino antes de copiar
            If abaDestino.UsedRange.MergeCells Then
                abaDestino.Cells.UnMerge
            End If

            ' Copia os formatos das células, incluindo mesclagens
            abaOrigem.Cells.Copy
            abaDestino.Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

            ' Agora mescla as células conforme a aba de origem
            For Each cell In abaOrigem.UsedRange
                If cell.MergeCells Then
                    abaDestino.Range(cell.MergeArea.Address).Merge
                End If
            Next cell

            ' Converte as fórmulas em valores apenas para "EST. DE CUSTOS"
            If nomeAba = "EST. DE CUSTOS" Or nomeAba = "MEMORIAL ORÇ" Or nomeAba = "CRONOGRAMA" Then
                ' Desmescla todas as células antes de colar os valores
                abaDestino.Cells.UnMerge
                
                ' Copia apenas os valores, removendo as fórmulas
                abaOrigem.Cells.Copy
                abaDestino.Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

                ' Mescla novamente as células conforme a aba de origem
                For Each cell In abaOrigem.UsedRange
                    If cell.MergeCells Then
                        abaDestino.Range(cell.MergeArea.Address).Merge
                    End If
                Next cell
            End If

            ' Bloquear todas as células da aba
            abaDestino.Cells.Locked = True
            abaDestino.Protect Password:="UEG", AllowFiltering:=True

            Application.CutCopyMode = False
        Else
            MsgBox "A aba '" & nomeAba & "' não foi encontrada na planilha de origem.", vbExclamation
        End If
    Next i

    ' Salva e fecha a planilha de destino
    planilhaDestino.Save
    planilhaDestino.Close
    MsgBox "Cópia concluída com sucesso!", vbInformation
End Sub
