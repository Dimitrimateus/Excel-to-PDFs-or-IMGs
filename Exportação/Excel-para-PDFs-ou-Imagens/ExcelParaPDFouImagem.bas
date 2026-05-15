Option Explicit

' ==============================================================================
' CONFIGURAÇÕES GERAIS
' ==============================================================================
Const HEADER_ROWS As String = "$1:$1"   ' Linha de cabeçalho para repetir no PDF
Const LIMITE_LINHAS As Long = 60        ' Até 60 linhas = FOTO. Acima disso = PDF.

Sub Exportar_PDFs_Ou_Fotos_Dinamico()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim pastaDestino As String
    Dim dlg As FileDialog
    Dim itensUnicos As Object
    Dim valoresColuna As Variant
    Dim i As Long
    Dim item As Variant
    Dim nomeArquivoBase As String
    Dim prefixo As String
    Dim linhasVisiveis As Long
    Dim rngSelecao As Range
    Dim colunaFiltro As Long
    
    On Error GoTo TratarErro
    
    ' --- 1. PREPARAÇÃO E ESCOLHA DA COLUNA ---
    Set ws = ActiveSheet ' Agora funciona em qualquer aba que estiver aberta
    
    ' Pede para o usuário clicar na coluna que será o filtro
    On Error Resume Next
    Set rngSelecao = Application.InputBox( _
        Prompt:="Clique em qualquer célula da COLUNA que você deseja usar para separar os arquivos." & vbCrLf & vbCrLf & "(Ex: Clique na coluna de Vendedores, Cidades, etc.)", _
        Title:="Selecionar Coluna de Filtro", _
        Type:=8)
    On Error GoTo TratarErro
    
    ' Se o usuário cancelar a seleção, encerra o código
    If rngSelecao Is Nothing Then Exit Sub
    
    colunaFiltro = rngSelecao.Column
    
    prefixo = InputBox("Digite o prefixo para os arquivos (ex: Relatorio):", "Prefixo do Arquivo", "Relatorio")
    If prefixo = "" Then Exit Sub ' Se cancelar, encerra
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ultimaLinha = ws.Cells(ws.Rows.Count, colunaFiltro).End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Nenhum dado encontrado na coluna selecionada!", vbExclamation
        GoTo Encerrar
    End If
    
    Call ConfigurarPagina(ws)
    
    ' Escolher pasta de destino
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Selecione a pasta para salvar os arquivos"
    If dlg.Show <> -1 Then GoTo Encerrar
    pastaDestino = dlg.SelectedItems(1)
    
    ' --- 2. MAPEAMENTO DE ITENS ÚNICOS ---
    Set itensUnicos = CreateObject("Scripting.Dictionary")
    valoresColuna = ws.Range(ws.Cells(2, colunaFiltro), ws.Cells(ultimaLinha, colunaFiltro)).Value
    
    For i = 1 To UBound(valoresColuna, 1)
        If valoresColuna(i, 1) <> "" Then
            If Not itensUnicos.Exists(valoresColuna(i, 1)) Then
                itensUnicos.Add valoresColuna(i, 1), 1
            End If
        End If
    Next i
    
    ' Remove filtros existentes antes de começar
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' --- 3. GERAÇÃO DOS ARQUIVOS (LOOP) ---
    For Each item In itensUnicos.Keys
        ' Filtra a base inteira usando a coluna escolhida
        ws.Range("A1").CurrentRegion.AutoFilter Field:=colunaFiltro, Criteria1:=item
        
        ' Conta linhas visíveis (ignorando o cabeçalho)
        linhasVisiveis = Application.WorksheetFunction.Subtotal(103, ws.Range(ws.Cells(2, colunaFiltro), ws.Cells(ultimaLinha, colunaFiltro)))
        nomeArquivoBase = pastaDestino & "\" & LimparNome(prefixo & "_" & item)
        
        ' Deleta arquivos antigos com o mesmo nome para não dar conflito
        On Error Resume Next
        Kill nomeArquivoBase & ".pdf"
        Kill nomeArquivoBase & ".png"
        On Error GoTo TratarErro
        
        ' Decide se vai ser PNG ou PDF
        If linhasVisiveis <= LIMITE_LINHAS Then
            Call ExportarComoImagemSuprema(ws, nomeArquivoBase & ".png")
        Else
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                                   Filename:=nomeArquivoBase & ".pdf", _
                                   Quality:=xlQualityStandard
        End If
    Next item
    
    ' --- 4. FINALIZAÇÃO ---
    ws.AutoFilterMode = False
    MsgBox "Tudo pronto! Arquivos gerados com sucesso na pasta selecionada.", vbInformation

Encerrar:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

TratarErro:
    MsgBox "Ops! Ocorreu um erro: " & Err.Description, vbCritical
    Resume Encerrar
End Sub

' ==============================================================================
' FUNÇÃO PARA TIRAR "PRINT" (COMBO SUPREMO CONTRA TELAS PRETAS)
' ==============================================================================
Private Sub ExportarComoImagemSuprema(wsOrigem As Worksheet, caminhoCompleto As String)
    Dim rngFiltrado As Range
    Dim rngTemp As Range
    Dim wsTemp As Worksheet
    Dim chartObj As ChartObject
    Dim tentativa As Integer
    Dim falhou As Boolean
    Dim ultimaLinhaTemp As Long
    Dim ultimaColunaTemp As Long
    
    Set rngFiltrado = wsOrigem.AutoFilter.Range
    
    Application.ScreenUpdating = False
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "TempPrint_" & Format(Now, "hhmmss")
    
    On Error GoTo LimparTemp
    
    rngFiltrado.Copy
    With wsTemp.Range("A1")
        .PasteSpecial Paste:=xlPasteColumnWidths
        .PasteSpecial Paste:=xlPasteAll
    End With
    Application.CutCopyMode = False
    
    ultimaLinhaTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    ultimaColunaTemp = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
    Set rngTemp = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(ultimaLinhaTemp, ultimaColunaTemp))
    
    ActiveWindow.DisplayGridlines = False
    wsTemp.Activate
    rngTemp.Cells(1, 1).Select
    ActiveWindow.Zoom = 100
    
    Application.ScreenUpdating = True
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01"))
    
    On Error Resume Next
    For tentativa = 1 To 10
        Err.Clear
        rngTemp.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
        If Err.Number = 0 Then Exit For
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01") / 10)
    Next tentativa
    
    If Err.Number <> 0 Then falhou = True
    On Error GoTo LimparTemp
    
    If falhou Then
        MsgBox "O Windows engasgou e se recusou a copiar o trecho: " & caminhoCompleto, vbExclamation
        GoTo LimparTemp
    End If
    
    Application.ScreenUpdating = False
    
    Set chartObj = wsTemp.ChartObjects.Add(Left:=0, Top:=0, Width:=rngTemp.Width, Height:=rngTemp.Height)
    With chartObj
        .Chart.ChartArea.Format.Line.Visible = msoFalse
        .Chart.ChartArea.Format.Fill.Visible = msoTrue
        .Chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Activate
        .Chart.Paste
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
        .Chart.Export Filename:=caminhoCompleto, FilterName:="PNG"
        .Delete
    End With
    
LimparTemp:
    Application.DisplayAlerts = False
    On Error Resume Next
    wsTemp.Delete
    Application.DisplayAlerts = True
    wsOrigem.Activate
    Application.ScreenUpdating = False
End Sub

' ==============================================================================
' FUNÇÕES AUXILIARES
' ==============================================================================
Private Sub ConfigurarPagina(ws As Worksheet)
    With ws.PageSetup
        .PrintTitleRows = HEADER_ROWS
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
End Sub

Private Function LimparNome(ByVal texto As String) As String
    Dim caracteresInvalidos As Variant
    Dim i As Long
    caracteresInvalidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(caracteresInvalidos) To UBound(caracteresInvalidos)
        texto = Replace(texto, caracteresInvalidos(i), "-") ' Troquei o ponto por hífen para evitar bugar extensões
    Next i
    LimparNome = texto
End Function

