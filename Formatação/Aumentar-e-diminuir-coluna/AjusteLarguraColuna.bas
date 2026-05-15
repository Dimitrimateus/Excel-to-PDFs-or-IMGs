Sub AumentarLarguraColuna()
    Dim col As Range
    ' Ignora erro caso a coluna chegue ao limite máximo do Excel (255)
    On Error Resume Next
    
    ' Loop para funcionar mesmo se você selecionar várias colunas de uma vez
    For Each col In Selection.Columns
        col.ColumnWidth = col.ColumnWidth + 1 ' Aumenta em 1 unidade
    Next col
    
    On Error GoTo 0
End Sub

Sub DiminuirLarguraColuna()
    Dim col As Range
    ' Ignora erro para não bugar o código se a largura for zero
    On Error Resume Next
    
    For Each col In Selection.Columns
        ' Verifica se a coluna já está muito fina para evitar erros negativos
        If col.ColumnWidth > 1 Then
            col.ColumnWidth = col.ColumnWidth - 1 ' Diminui em 1 unidade
        Else
            col.ColumnWidth = 0.1
        End If
    Next col
    
    On Error GoTo 0
End Sub
