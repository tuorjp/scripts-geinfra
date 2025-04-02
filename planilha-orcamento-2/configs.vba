
'No diretório: Microsoft Excel Objects/EstaPastaDeTrabalho
Private Sub Workbook_Open()
    Call configurarComboBox
End Sub

'No diretório: Planilha4(MEMORIAL ORÇ)
Private Sub cmbTipoValor_Change()
    Call atualizarFormatacao
    Call aplicarFormatacaoCondicional
End Sub
