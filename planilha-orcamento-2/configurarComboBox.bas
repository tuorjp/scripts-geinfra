Sub configurarComboBox()
    Dim cmb As Object
    Set cmb = ThisWorkbook.Sheets("MEMORIAL ORÇ").OLEObjects("cmbTipoValor").Object
    
    ' Limpa a ComboBox antes de adicionar opções
    cmb.Clear
    cmb.AddItem "QUANTIDADE"
    cmb.AddItem "PORCENTAGEM"
    
    ' Define um valor padrão
    cmb.Value = "QUANTIDADE"
End Sub
