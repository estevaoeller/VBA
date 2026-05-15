Sub AuditarDependentes()

    ' ============================================================
    ' Objetivo:
    ' Rastrear e auditar as dependências (internas e externas) de 
    ' um intervalo específico de células ("B28:N71") na aba "Opex_UE".
    '
    ' Ações realizadas:
    ' 1. Cria ou limpa uma aba de relatório chamada "Auditoria".
    ' 2. Verifica cada célula do intervalo em busca de fórmulas ou valores.
    ' 3. Identifica células dependentes na mesma aba (internas).
    ' 4. Identifica células dependentes em outras abas (externas).
    ' 5. Registra o endereço, fórmula/valor e os dependentes no relatório.
    ' 6. Aplica formatação visual (cabeçalhos, cores e auto-filtro).
    ' ============================================================

    Dim wsOrigem As Worksheet
    Dim wsAudit As Worksheet
    Dim rngAudit As Range
    Dim cell As Range
    Dim linha As Long
    Dim depRange As Range
    Dim depCell As Range
    Dim depsInternas As String
    Dim depsExternas As String
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsOrigem = ThisWorkbook.Sheets("Opex_UE")
    Set rngAudit = wsOrigem.Range("B28:N71")
    
    ' Criar ou limpar aba de auditoria
    On Error Resume Next
    Set wsAudit = ThisWorkbook.Sheets("Auditoria")
    On Error GoTo 0
    If wsAudit Is Nothing Then
        Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Change Log"))
        wsAudit.Name = "Auditoria"
        wsAudit.TabColor = RGB(255, 192, 0)
    Else
        wsAudit.Cells.Clear
    End If
    
    ' Cabeçalho
    With wsAudit
        .Range("A1").Value = "Célula"
        .Range("B1").Value = "Aba"
        .Range("C1").Value = "Fórmula / Valor"
        .Range("D1").Value = "Dependentes (mesma aba)"
        .Range("E1").Value = "Dependentes (outras abas)"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(31, 56, 100)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
    End With
    
    linha = 2
    wsOrigem.Activate  ' necessário para ShowDependents funcionar
    
    For Each cell In rngAudit
        If cell.HasFormula Or cell.Value <> "" Then
        
            depsInternas = ""
            depsExternas = ""
            
            ' --- Dependentes mesma aba via objeto nativo ---
            On Error Resume Next
            Set depRange = Nothing
            Set depRange = cell.Dependents
            On Error GoTo 0
            
            If Not depRange Is Nothing Then
                For Each depCell In depRange
                    depsInternas = depsInternas & depCell.Address(False, False) & " "
                Next depCell
            End If
            
            ' --- Dependentes cross-sheet via setas de auditoria ---
            On Error Resume Next
            cell.ShowDependents True
            i = 1
            Do
                Err.Clear
                Set depCell = cell.NavigateArrow(True, i)
                If Err.Number <> 0 Then Exit Do
                If depCell.Worksheet.Name <> wsOrigem.Name Then
                    depsExternas = depsExternas & depCell.Worksheet.Name & "!" & depCell.Address(False, False) & " "
                End If
                i = i + 1
            Loop
            cell.ShowDependents False
            On Error GoTo 0
            
            ' Registra apenas se tiver fórmula ou dependentes
            If cell.HasFormula Or depsInternas <> "" Or depsExternas <> "" Then
                wsAudit.Cells(linha, "A").Value = cell.Address(False, False)
                wsAudit.Cells(linha, "B").Value = wsOrigem.Name
                wsAudit.Cells(linha, "C").Value = IIf(cell.HasFormula, cell.Formula, cell.Value)
                wsAudit.Cells(linha, "D").Value = Trim(depsInternas)
                wsAudit.Cells(linha, "E").Value = Trim(depsExternas)
                linha = linha + 1
            End If
        End If
    Next cell
    
    wsAudit.Columns("A:E").AutoFit
    wsAudit.Range("A1:E1").AutoFilter
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Auditoria concluída. " & (linha - 2) & " células registradas.", vbInformation
    wsAudit.Activate
End Sub