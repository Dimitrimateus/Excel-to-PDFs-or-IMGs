Sub AutoAjustarColunasInteligente()
    ' ==============================================================================
    ' Macro: AutoAjustarColunasInteligente
    ' Propósito: Encontra o menor tamanho possível que comporte o título E todo o 
    ' conteúdo da coluna, adicionando um respiro mínimo para evitar o erro "###".
    ' ==============================================================================
    
    Dim col As Range
    ' 1.5 é o menor valor seguro para evitar que filtros escondam o texto
    Const PADDING_MINIMO As Double = 0.75
    
    On Error Resume Next
    
    ' Garantimos que o Excel olhe para a COLUNA INTEIRA (EntireColumn)
    ' e não apenas para as células que você selecionou com o mouse.
    Selection.EntireColumn.AutoFit
    
    ' Adiciona o respiro mínimo necessário para renderização de tela e filtros
    For Each col In Selection.Columns
        If col.ColumnWidth > 0 Then
            col.ColumnWidth = col.ColumnWidth + PADDING_MINIMO
        End If
    Next col
    
    On Error GoTo 0
End Sub