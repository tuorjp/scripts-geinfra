Sub Imprimir()

    Application.Dialogs(xlDialogPrinterSetup).Show
    Sheets("CAPA").Range("D5:N51").PrintPreview

End Sub

Sub Imprimir_Com_Aba_Temporaria()
    
    Dim wsTemp As Worksheet, wsCapa As Worksheet, wsEstCusto As Worksheet

    Set wsTemp = ThisWorkBook.Sheets.Add
    wsTemp.Name = "TEMP_IMPRESSAO"

    Set wsCapa = Sheets("CAPA")
    Set wsEstCusto = Sheets("EST. DE CUSTOS")

    wsCapa.Range("D5:N51").Copy wsTemp.cells(1, 1)
    wsEstCusto.Range("K12:AD47").Copy wsTemp.cells(60, 1)

    wsTemp.printOut Copies := 1, Collate := True

    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True

End Sub
