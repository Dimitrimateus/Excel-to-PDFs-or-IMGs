Sub AutoAjustarColunasSelecionadas()
    ' ==============================================================================
    ' Macro: AutoAjustarColunasSelecionadas
    ' Propósito: Simula o "duplo clique" nas bordas das colunas, ajustando a
    ' largura de todas as colunas selecionadas para o tamanho exato do conteúdo.
    ' ==============================================================================
    
    ' Ignora erros de seleção inválida
    On Error Resume Next
    
    ' O método AutoFit faz exatamente a mesma ação do "duplo clique"
    Selection.Columns.AutoFit
    
    On Error GoTo 0
End Sub
