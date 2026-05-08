Option Explicit

' ==============================================================================
' CONFIGURAÇÕES RÁPIDAS
' ==============================================================================
Const DATA_SHEET As String = "Sheet1"   ' Nome da aba onde estão seus dados
Const ZONE_COLUMN As Long = 4           ' Coluna D (4) onde ficam as Zonas
Const HEADER_ROWS As String = "$1:$1"   ' Linha de cabeçalho para repetir no PDF
Const LIMITE_LINHAS As Long = 60        ' Até 60 linhas = FOTO. Acima disso = PDF.

Sub Exportar_PDFs_Ou_Fotos_Por_Zona()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim pastaDestino As String
    Dim dlg As FileDialog
    Dim zonasUnicas As Object
    Dim valoresZona As Variant
    Dim i As Long
    Dim zona As Variant
    Dim nomeArquivoBase As String
    Dim prefixo As String
    Dim linhasVisiveis As Long
    
    On Error GoTo TratarErro
    
    ' --- 1. PREPARAÇÃO ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DATA_SHEET)
    On Error GoTo TratarErro
    If ws Is Nothing Then
        MsgBox "Aba '" & DATA_SHEET & "' não encontrada!", vbCritical
        GoTo Encerrar
    End If
    
    prefixo = InputBox("Digite o prefixo para os arquivos:", "Prefixo", "Relatorio")
    If prefixo = "" Then GoTo Encerrar
    
    ultimaLinha = ws.Cells(ws.Rows.Count, ZONE_COLUMN).End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Nenhum dado encontrado!", vbExclamation
        GoTo Encerrar
    End If
    
    Call ConfigurarPagina(ws)
    
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    If dlg.Show <> -1 Then GoTo Encerrar
    pastaDestino = dlg.SelectedItems(1)
    
    ' --- 2. MAPEAMENTO DE ZONAS ---
    Set zonasUnicas = CreateObject("Scripting.Dictionary")
    valoresZona = ws.Range(ws.Cells(2, ZONE_COLUMN), ws.Cells(ultimaLinha, ZONE_COLUMN)).Value
    
    For i = 1 To UBound(valoresZona, 1)
        If valoresZona(i, 1) <> "" Then
            If Not zonasUnicas.Exists(valoresZona(i, 1)) Then
                zonasUnicas.Add valoresZona(i, 1), 1
            End If
        End If
    Next i
    
    ' --- 3. GERAÇÃO DOS ARQUIVOS (LOOP) ---
    For Each zona In zonasUnicas.Keys
        ws.Range("A1").AutoFilter Field:=ZONE_COLUMN, Criteria1:=zona
        
        linhasVisiveis = Application.WorksheetFunction.Subtotal(103, ws.Range(ws.Cells(2, ZONE_COLUMN), ws.Cells(ultimaLinha, ZONE_COLUMN)))
        nomeArquivoBase = pastaDestino & "\" & LimparNome(prefixo & "_" & zona)
        
        On Error Resume Next
        Kill nomeArquivoBase & ".pdf"
        Kill nomeArquivoBase & ".png"
        On Error GoTo TratarErro
        
        If linhasVisiveis <= LIMITE_LINHAS Then
            Call ExportarComoImagemSuprema(ws, nomeArquivoBase & ".png")
        Else
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                                   Filename:=nomeArquivoBase & ".pdf", _
                                   Quality:=xlQualityStandard
        End If
    Next zona
    
    ' --- 4. FINALIZAÇÃO ---
    ws.AutoFilterMode = False
    MsgBox "Tudo pronto! Arquivos gerados com sucesso.", vbInformation
    
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
    
    ' 1. Copia o intervalo (O Excel ignora linhas ocultas automaticamente)
    Set rngFiltrado = wsOrigem.AutoFilter.Range
    
    ' 2. Cria a aba temporária limpa
    Application.ScreenUpdating = False
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "TempPrint_" & Format(Now, "hhmmss")
    
    On Error GoTo LimparTemp
    
    ' 3. Cola os dados (Unificando a tabela sem as linhas ocultas do filtro)
    rngFiltrado.Copy
    With wsTemp.Range("A1")
        .PasteSpecial Paste:=xlPasteColumnWidths
        .PasteSpecial Paste:=xlPasteAll
    End With
    Application.CutCopyMode = False
    
    ' Descobre as bordas reais da nova tabela limpa
    ultimaLinhaTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    ultimaColunaTemp = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
    Set rngTemp = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(ultimaLinhaTemp, ultimaColunaTemp))
    
    ' Prepara o visual (fundo branco)
    ActiveWindow.DisplayGridlines = False
    wsTemp.Activate
    rngTemp.Cells(1, 1).Select
    ActiveWindow.Zoom = 100
    
    Application.ScreenUpdating = True
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01"))
    
    ' 4. Tira a foto usando xlPrinter (Fuga do limite de largura do monitor)
    On Error Resume Next
    For tentativa = 1 To 10 ' Aumentamos a insistência para vencer o bloqueio do Windows
        Err.Clear
        rngTemp.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
        If Err.Number = 0 Then Exit For
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01") / 10)
    Next tentativa
    
    If Err.Number <> 0 Then falhou = True
    On Error GoTo LimparTemp
    
    If falhou Then
        MsgBox "O Windows engasgou e se recusou a copiar a zona: " & caminhoCompleto, vbExclamation
        GoTo LimparTemp
    End If
    
    Application.ScreenUpdating = False
    
    ' 5. Cria o Gráfico e Força o Fundo a ser BRANCO (Mata as linhas pretas)
    Set chartObj = wsTemp.ChartObjects.Add(Left:=0, Top:=0, Width:=rngTemp.Width, Height:=rngTemp.Height)
    With chartObj
        .Chart.ChartArea.Format.Line.Visible = msoFalse
        
        ' Força o fundo branco puro
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
    ' 6. Autodestruição da aba temporária
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
        texto = Replace(texto, caracteresInvalidos(i), ".")
    Next i
    LimparNome = texto
End Function