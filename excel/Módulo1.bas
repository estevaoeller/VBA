Attribute VB_Name = "Calcular_CP"




Sub CALC_CPeTJ_SV()
    ' Definir a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")

    ' Definir as células de trabalho como variáveis
    Dim modoDaPlanilha As Range:         Set modoDaPlanilha = ws.Range("H37")
    Dim modoPlanilhaOpcao As Range:      Set modoPlanilhaOpcao = ws.Range("F59")
    
    Dim vplAlvo As Range:                Set vplAlvo = ws.Range("D45")
    Dim contraprestacaoCenario As Range: Set contraprestacaoCenario = ws.Range("H41")
    
    Dim tjAlvo As Range:                 Set tjAlvo = ws.Range("D46")
    Dim txjurosatvfinCenario As Range:   Set txjurosatvfinCenario = ws.Range("G41")
    
    Dim resultadoMacro As Range:         Set resultadoMacro = ws.Range("B59")

    ' Mudar a planilha para o modo de cálculo
    modoDaPlanilha.Value = modoPlanilhaOpcao.Value

    ' Definir variáveis para apresentar o tempo e a data do cálculo
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date

    ' Iniciar o temporizador
    startTime = Timer

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' O Solver funciona de forma mais confiável na planilha ativa
    ws.Activate
    
    ' Executar o Solver para encontrar as duas variáveis simultaneamente
    SolverReset
    ' Define o objetivo principal: vplAlvo = 0, mudando as duas células de cenário
    SolverOk SetCell:=vplAlvo.Address, MaxMinVal:=3, ValueOf:=0, ByChange:=contraprestacaoCenario.Address & "," & txjurosatvfinCenario.Address
    ' Adiciona a restrição: tjAlvo = 0
    SolverAdd CellRef:=tjAlvo.Address, Relation:=2, FormulaText:="0"
    ' Resolve sem exibir a caixa de diálogo do Solver
    SolverSolve UserFinish:=True

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ' Deslocar o histórico de resultados para baixo (mantendo 15 registros: de B59 até B73)
    ws.Range("B60:B73").Value = ws.Range("B59:B72").Value
    
    ' Limpa a 16ª posição (B74) caso algum dado tenha escapado
    ws.Range("B74").ClearContents

    ' Escrever o novo tempo de execução e a data na primeira célula de resultado (B59)
    resultadoMacro.Value = "CP+TJ-SV: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Notificar o usuário com uma caixa de mensagem
    ' MsgBox "This code ran successfully on " & CurrentDate & " in " & MinutesElapsed & " minutes", vbInformation

    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Lidar com erros e garantir que a tela volte a atualizar
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
End Sub



Sub CALC_CPeTJ_GS()
    ' Definir a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")

    ' Definir as células de trabalho como variáveis
    Dim modoDaPlanilha As Range:         Set modoDaPlanilha = ws.Range("H37")
    Dim modoPlanilhaOpcao As Range:      Set modoPlanilhaOpcao = ws.Range("F59")
    
    Dim vplAlvo As Range:                Set vplAlvo = ws.Range("D45")
    Dim contraprestacaoCenario As Range: Set contraprestacaoCenario = ws.Range("H41")
    
    Dim tjAlvo As Range:                 Set tjAlvo = ws.Range("D46")
    Dim txjurosatvfinCenario As Range:   Set txjurosatvfinCenario = ws.Range("G41")
    
    Dim resultadoMacro As Range:         Set resultadoMacro = ws.Range("B59")

    ' Mudar a planilha para o modo de cálculo
    modoDaPlanilha.Value = modoPlanilhaOpcao.Value

    ' Definir variáveis para apresentar o tempo e a data do cálculo
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date
    Dim i As Integer

    ' Iniciar o temporizador
    startTime = Timer

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Executar o Atingir Meta (Goal Seek) de forma iterativa para estabilizar os dois valores
    For i = 1 To 5
        vplAlvo.GoalSeek Goal:=0, ChangingCell:=contraprestacaoCenario
        tjAlvo.GoalSeek Goal:=0, ChangingCell:=txjurosatvfinCenario
    Next i

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ' Deslocar o histórico de resultados para baixo (mantendo 15 registros: de B59 até B73)
    ws.Range("B60:B73").Value = ws.Range("B59:B72").Value
    
    ' Limpa a 16ª posição (B74) caso algum dado tenha escapado
    ws.Range("B74").ClearContents

    ' Escrever o novo tempo de execução e a data na primeira célula de resultado (B59)
    resultadoMacro.Value = "CP+TJ-SV: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Notificar o usuário com uma caixa de mensagem
    ' MsgBox "This code ran successfully on " & CurrentDate & " in " & MinutesElapsed & " minutes", vbInformation

    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Lidar com erros e garantir que a tela volte a atualizar
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
End Sub




Sub CALC_CP_SPOLVER()
    ' Definir a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")

    ' Definir as células de trabalho como variáveis
    Dim modoDaPlanilha As Range:         Set modoDaPlanilha = ws.Range("H37")
    Dim modoPlanilhaOpcao As Range:      Set modoPlanilhaOpcao = ws.Range("F59")
    Dim vplAlvo As Range:                Set vplAlvo = ws.Range("D45")
    Dim contraprestacaoCenario As Range: Set contraprestacaoCenario = ws.Range("H41")
    Dim resultadoMacro As Range:         Set resultadoMacro = ws.Range("B59")

    ' Mudar a planilha para o modo de cálculo
    modoDaPlanilha.Value = modoPlanilhaOpcao.Value

    ' Definir variáveis para apresentar o tempo e a data do cálculo
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date

    ' Iniciar o temporizador
    startTime = Timer

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Executar o Atingir Meta (Goal Seek)
    vplAlvo.GoalSeek Goal:=0, ChangingCell:=contraprestacaoCenario

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ' Deslocar o histórico de resultados para baixo (mantendo 15 registros: de B59 até B73)
    ws.Range("B60:B73").Value = ws.Range("B59:B72").Value
    
    ' Limpa a 16ª posição (B74) caso algum dado tenha escapado
    ws.Range("B74").ClearContents

    ' Escrever o novo tempo de execução e a data na primeira célula de resultado (B59)
    resultadoMacro.Value = "CP: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Notificar o usuário com uma caixa de mensagem
    ' MsgBox "This code ran successfully on " & CurrentDate & " in " & MinutesElapsed & " minutes", vbInformation

    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Lidar com erros e garantir que a tela volte a atualizar
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
End Sub

Sub CALC_TJ_SPOLVER()
    ' Definir a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")

    ' Definir as células de trabalho como variáveis
    Dim modoDaPlanilha As Range:         Set modoDaPlanilha = ws.Range("H37")
    Dim modoPlanilhaOpcao As Range:      Set modoPlanilhaOpcao = ws.Range("F59")
    Dim tjAlvo As Range:                 Set tjAlvo = ws.Range("D46")
    Dim txjurosatvfinCenario As Range:   Set txjurosatvfinCenario = ws.Range("G41")
    Dim resultadoMacro As Range:         Set resultadoMacro = ws.Range("B59")

    ' Mudar a planilha para o modo de cálculo
    modoDaPlanilha.Value = modoPlanilhaOpcao.Value

    ' Definir variáveis para apresentar o tempo e a data do cálculo
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date

    ' Iniciar o temporizador
    startTime = Timer

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Executar o Atingir Meta (Goal Seek)
    tjAlvo.GoalSeek Goal:=0, ChangingCell:=txjurosatvfinCenario

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ' Deslocar o histórico de resultados para baixo (mantendo 15 registros: de B59 até B73)
    ws.Range("B60:B73").Value = ws.Range("B59:B72").Value
    
    ' Limpa a 16ª posição (B74) caso algum dado tenha escapado
    ws.Range("B74").ClearContents

    ' Escrever o novo tempo de execução e a data na primeira célula de resultado (B59)
    resultadoMacro.Value = "TJ: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Notificar o usuário com uma caixa de mensagem
    ' MsgBox "This code ran successfully on " & CurrentDate & " in " & MinutesElapsed & " minutes", vbInformation

    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Lidar com erros e garantir que a tela volte a atualizar
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
End Sub






Sub AtivoFinanc()
'
' AtivoFinanc Macro
'
    Range("D36").Select
    Application.CutCopyMode = False
        Sheets("Resultados").Range("D36").GoalSeek Goal:=0, ChangingCell:=Sheets("Resultados").Range("G42")
    Range("A1").Select
    
End Sub




Sub CalcularContraprestacaoSolver()


    ' Defina a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")

    ' Mudar a planilha para o modo de cálculo
    ws.Range("H37").Value = ws.Range("D50").Value

    ' Declaração de variáveis
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date
    Dim ResultCell As Range
    Dim i As Long

    ' Iniciar o temporizador
    startTime = Timer

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    For i = 1 To 5
        ' GoalSeek para D13
        ws.Range("D13").GoalSeek Goal:=0, ChangingCell:=ws.Range("H42")
        
        ' GoalSeek para D34
        ' ws.Range("D36").GoalSeek Goal:=0, ChangingCell:=ws.Range("g42")
    Next i

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ' Definir a célula de resultado
    ' AJUSTAR QUANDO MUDAR DE LUGAR
    Set ResultCell = ws.Range("B44")

    ' Escrever o tempo de execução e a data na célula de resultado
    ResultCell.Value = "CP: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Notificar o usuário com uma caixa de mensagem
    ' MsgBox "This code ran successfully on " & CurrentDate & " in " & MinutesElapsed & " minutes", vbInformation

    ' Recalcular a planilha
    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    ' Restaurar as configurações do Excel
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Lidar com erros (substitua ou adicione conforme necessário)
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
    Resume Next
End Sub

Sub Copy_to_table()
    Dim searchValue As Variant
    Dim searchRange As Range
    Dim sourceRange1 As Range
    Dim sourceRange2 As Range
    Dim sourceRange3 As Range
    Dim destinationRange1 As Range
    Dim destinationRange2 As Range
    Dim destinationRange3 As Range
    Dim i As Integer
    
    ' Set the value to search in range S
    searchValue = Range("I34").Value
    
    ' Set the search range (S11 to S19)
    Set searchRange = Range("E41:E43")
    
    ' Set the source ranges (U9 and V9)
    Set sourceRange1 = Range("H40")
    Set sourceRange2 = Range("G40")
    Set sourceRange3 = Range("H34")
    
    ' Set the destination ranges (U11 to U19 and V11 to V19)
    Set destinationRange1 = Range("H41:H43")
    Set destinationRange2 = Range("G41:G43")
    Set destinationRange3 = Range("F41:F43")
    
    ' Loop through the search range to find the matching value
    For i = 1 To searchRange.Rows.Count
        If searchRange.Cells(i).Value = searchValue Then
            ' Copy the corresponding values from source ranges to destination ranges
            destinationRange1.Cells(i).Value = sourceRange1.Value
            sourceRange2.Copy
            destinationRange2.Cells(i).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
' Desabilitado, não estava funcionando. Copiava o valor por cima da fórmula
'            destinationRange2.Cells(i).Value = sourceRange2.Value
            destinationRange3.Cells(i).Value = sourceRange3.Value
            Exit For ' Exit the loop once a match is found
        End If
    Next i
End Sub





Sub Rodar_Cenarios()
    ' Rodar_Cenarios Macro
    
    DoEvents
    
    ' Set the calculation mode to the first scenario
    Sheets("Resultados").Range("H35").Value = Sheets("Resultados").Range("D47").Value
    
' Mostra na tela?
    DoEvents
    
    ' Loop through scenarios
    Dim scenarioRange As Range
    Dim scenarioCell As Range
    Set scenarioRange = Sheets("Resultados").Range("H47:H49")
    
    For Each scenarioCell In scenarioRange
        ' Select the current scenario
        Sheets("Resultados").Range("H34").Value = scenarioCell.Value
        
        ' Run the calculation macro
        Application.Run "PPP_ESCOLAS_SP_v9.xlsm!CalcularContraprestacao"
        
        ' Run the copy to table macro
        Application.Run "PPP_ESCOLAS_SP_v9.xlsm!Copy_to_table"
    Next scenarioCell
    
    ' Set the calculation mode back to the result
    Sheets("Resultados").Range("H35").Value = Sheets("Resultados").Range("D48").Value
End Sub



Sub Rodar_Contra_TJ()

    ' Defina a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")
    Dim ResultCell           As Range: Set ResultCell = ws.Range("B46")
    
    ' Modo da planilha
    Dim Modo_Planilha        As Range: Set Modo_Planilha = ws.Range("H37")
    Dim Modo_Planilha_Nome   As Range: Set Modo_Planilha_Nome = ws.Range("D50")
        
    ' Declaração das células como variáveis
    Dim VPL                  As Range: Set VPL = ws.Range("D13")
    Dim Tarifa               As Range: Set Tarifa = ws.Range("H42")
        
    Dim Saldo_Ativo_Fin      As Range: Set Saldo_Ativo_Fin = ws.Range("D36")
    Dim TJ                   As Range: Set TJ = ws.Range("G42")
  
    ' Mudar a planilha para o modo de cálculo
    Modo_Planilha.Value = Modo_Planilha_Nome.Value

    ' Declaração de variáveis
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date
    Dim i As Long

    ' Iniciar o temporizador
    startTime = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    For i = 1 To 5
        VPL.GoalSeek Goal:=0, ChangingCell:=Tarifa
        Saldo_Ativo_Fin.GoalSeek Goal:=0, ChangingCell:=TJ
    Next i

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ResultCell.Value = "CP+TJ: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Recalcular a planilha
    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    ' Restaurar as configurações do Excel
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Lidar com erros (substitua ou adicione conforme necessário)
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
    Resume Next
End Sub

Sub CP_Fixa()
'
' CP_Fixa Macro
'

    Dim wsRes As Worksheet
    
    Dim ParcelaFixaMaxima As Range
    Dim Tarifa As Range
    Dim OpexZerado As Range
    Dim VPL As Range
    Dim ResultCell As Range
    
    Dim startTime As Double
    Dim elapsedTime As String
    Dim currentDateTime As Date
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set wsRes = ThisWorkbook.Sheets("Resultados")
    
    ' Células na mesma aba
    Set ParcelaFixaMaxima = wsRes.Range("F42")
    Set Tarifa = wsRes.Range("H42")
    Set OpexZerado = wsRes.Range("H38")
    Set VPL = wsRes.Range("D13")
    Set ResultCell = wsRes.Range("B45")
    
    ' Início
    startTime = Timer
    
    ' Zera H42
    Tarifa.Value = 0
    
    ' Coloca OPEX zerado = 0
    OpexZerado.Value = 0
    
    ' Goal Seek
    VPL.GoalSeek Goal:=0, ChangingCell:=ParcelaFixaMaxima
    
    ' Restaura OPEX zerado = 1
    OpexZerado.Value = 1
    
    ' Tempo decorrido
    elapsedTime = Format((Timer - startTime) / 86400, "hh:mm:ss")
    currentDateTime = Now
    
    ResultCell.Value = "PF Máx.: Resultado em " & elapsedTime & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation



End Sub



Sub Rodar_TJ()



    ' Defina a planilha de trabalho
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")

    ' Mudar a planilha para o modo de cálculo
    ws.Range("H37").Value = ws.Range("D50").Value

    ' Declaração de variáveis
    Dim startTime As Double
    Dim MinutesElapsed As String
    Dim currentDateTime As Date
    Dim ResultCell As Range
    Dim i As Long

    ' Iniciar o temporizador
    startTime = Timer

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

'    For i = 1 To 5
        
        ' GoalSeek para D34
        ws.Range("D36").GoalSeek Goal:=0, ChangingCell:=ws.Range("g42")
'    Next i

    ' Determinar quanto tempo o código levou para ser executado
    MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' Obter a data e hora atual
    currentDateTime = Now

    ' Definir a célula de resultado
    Set ResultCell = ws.Range("B45")

    ' Escrever o tempo de execução e a data na célula de resultado
    ResultCell.Value = "TJ: Resultado em " & MinutesElapsed & " - " & Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")

    ' Notificar o usuário com uma caixa de mensagem
    ' MsgBox "This code ran successfully on " & CurrentDate & " in " & MinutesElapsed & " minutes", vbInformation

    ' Recalcular a planilha
    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")

    ' Restaurar as configurações do Excel
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Lidar com erros (substitua ou adicione conforme necessário)
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
    Resume Next
End Sub
