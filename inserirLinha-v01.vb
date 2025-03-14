Sub inserirLinha()
    'Macro Inserir Linha - insere nova linha, a partir de parâmetros indicados pelo usuário.
    'Salvar arquivo antes de realizar alterações
    ThisWorkbook.Save
    
    'Desativa eventos, atualização de tela e cálculo automático
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'Variáveis
    Dim rowNum As Integer
    Dim linhas As Integer
    Dim numero As Integer
    
    On Error Resume Next
    
    'Solicita entrada do usuário
    rowNum = Application.InputBox(Prompt:="Número da linha para inserir acima:", Title:="Linhas", Type:=1)

    'Verificações do input digitado
    If rowNum = False Then
        MsgBox "Script Cancelado" 
        Exit Sub
    End If

    If Not isNumeric(rowNum) Or rowNum = 0 Then
        MsgBox "Digite um número válido"
        Exit Sub
    ElseIf rowNum <= 27 Then
        MsgBox "Insira um Número maior que 27"
        Exit Sub
    End If

    linhas = Application.InputBox(Prompt:="Quantidade de linhas:", Title:="Quantidade", Type:=1)

    'Verificação do input digitado
    If linhas = False Then
        MsgBox "Script Cancelado" 
        Exit Sub
    End If

    If Not isNumeric(linhas) Or linhas <= 0 Then
        MsgBox "Insira um Número de linhas válido"
        Exit Sub
    End If

    'Inserir linhas na aba "EST. DE CUSTOS"
    For i = 1 To linhas
        Rows(rowNum + i - 1 & ":" & rowNum + i - 1).Insert Shift:=xlDown
    Next i
    
    'Solicita tipo de linha a ser copiada
    numero = Application.InputBox("1 - Título" & Chr(10) & _
                                 "2 - Subtítulo" & Chr(10) & _
                                 "3 - Itens" & Chr(10) & _
                                 "4 - Branco", "Escolha nº Tabela")
    
    'Verificação do input digitado
    If numero = False Then
        MsgBox "Script Cancelado" 
        Exit Sub
    End If

    If Not IsNumeric(numero) Or numero < 1 Or numero > 4 Then
        MsgBox "Escolha um número entre 1 e 4."
        Exit Sub
    End If
    
    'Seleciona a linha base conforme o tipo escolhido na primeira ABA EST. DE CUSTOS
    Select Case numero
        Case 1: Rows("4:4").Select   'Título
        Case 2: Rows("5:5").Select   'Subtítulo
        Case 3: Rows("6:6").Select   'Itens
        Case 4: Rows("7:7").Select   'Branco
    End Select
    
    'Copia e cola a linha selecionada nas novas linhas inseridas
    For i = 1 To linhas
        Selection.Copy
        Rows(rowNum + i - 1 & ":" & rowNum + i - 1).Select
        ActiveSheet.Paste
        Selection.EntireRow.Hidden = False
    Next i
    
    'Ajusta formatação
    Range("A4:H4").Select
    Selection.Copy
    Cells(rowNum + linhas, 1).EntireRow.AutoFit
    ActiveSheet.Paste
    Selection.EntireRow.Hidden = False
    
    'Cria um array(vetor) com os nomes de todas as abas e processa a inserção de linhas nelas
    'dentro do loop For Each
    Dim aba As Variant
    Dim abas As Variant
    abas = Array("MEMORIAL ORÇ", "SERV. TERCEIRIZAÇÃO", "CURVA ABC_ITENS DE RELEVÂNCIA", "CRONOGRAMA")
    
    For Each aba In abas
        Sheets(aba).Select
        
        'Seleciona linha base conforme o tipo escolhido na aba atual, lá em cima foi feito somente na aba EST. DE CUSTOS
        Select Case numero
            Case 1: Rows("4:4").Select
            Case 2: Rows("5:5").Select
            Case 3: Rows("6:6").Select
            Case 4: Rows("7:7").Select
        End Select

        If aba = "CRONOGRAMA" Then
            Dim linhaACopiar As Range
            Dim linhaEstCusto As Long
            Dim linhaCronograma1 As Long
            Dim linhaCronograma2 As Long

            Select Case numero
                Case 1: Set linhaACopiar = Rows("23:24")
                Case 2: Set linhaACopiar = Rows("25:26")
                Case 3: Set linhaACopiar = Rows("27:28")
                Case 4: Set linhaACopiar = Rows("29:30")
                Case Else
                    MsgBox "Número inválido!", vbExclamation
                    Exit Sub
            End Select

            'Para cada linha em EST. DE CUSTOS, insere duas em CRONOGRAMA
            For i =1 To linhas
                linhaEstCusto = rowNum + i - 1
                linhaCronograma1 = 2 * linhaEstCusto - 1
                linhaCronograma2 = 2 * linhaEstCusto
                
                'Insere duas linhas vazias para acomodar a cópia
                Rows(linhaCronograma1 & ":" & linhaCronograma2).Insert Shift:=xlDown

                'Copia as duas linhas e cola no local correto
                linhaACopiar.Copy Destination:=Rows(linhaCronograma1 & ":" & linhaCronograma2)
            Next i
        Else
            'Insere e copia as linhas
            For i = 1 To linhas
                Rows(rowNum + i - 1 & ":" & rowNum + i - 1).Insert Shift:=xlDown
                Selection.Copy
                Rows(rowNum + i - 1 & ":" & rowNum + i - 1).Select
                ActiveSheet.Paste
                Selection.EntireRow.Hidden = False
            Next i
        End If
    Next aba
    
    'Retorna para aba principal e reativa funcionalidades do Excel
    Sheets("EST. DE CUSTOS").Select
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False

End Sub
