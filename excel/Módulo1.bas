Attribute VB_Name = "Modulo1"

Private Const RESULTADOS_SHEET As String = "Resultados"
Private Const CALC_MODE_CELL As String = "H37"
Private Const CALC_MODE_OPTION_CELL As String = "F59"
Private Const RESULT_HISTORY_CELL As String = "B59"

Sub VPLZERO()
Attribute VPLZERO.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("D45").GoalSeek Goal:=0, ChangingCell:=Range("H41")
    Range("D37").GoalSeek Goal:=0, ChangingCell:=Range("G41")
End Sub

Sub CALC_CP_GS()
    RunResultadoGoalSeek "CP-GS", True, False, 1
End Sub

Sub CALC_TJ_GS()
    RunResultadoGoalSeek "TJ-GS", False, True, 1
End Sub

Sub CALC_CPeTJ_GS()
    RunResultadoGoalSeek "CP+TJ-GS", True, True, 5
End Sub

Sub CALC_CP_SV()
    RunResultadoGoalSeek "CP-SV", True, False, 1
End Sub

Sub CALC_TJ_SV()
    RunResultadoGoalSeek "TJ-SV", False, True, 1
End Sub

Sub CALC_CPeTJ_SV()
    Dim ws As Worksheet
    Dim startTime As Double

    Set ws = ThisWorkbook.Sheets(RESULTADOS_SHEET)
    PrepareResultadosMode ws

    startTime = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ws.Activate

    SolverReset
    SolverOk SetCell:=ws.Range("D45").Address, _
             MaxMinVal:=3, _
             ValueOf:=0, _
             ByChange:=ws.Range("H41").Address & "," & ws.Range("G41").Address
    SolverAdd CellRef:=ws.Range("D46").Address, Relation:=2, FormulaText:="0"
    SolverSolve UserFinish:=True

    FinishResultadoCalculation ws, "CP+TJ-SV", startTime
    Exit Sub

ErrorHandler:
    HandleCalculationError
End Sub

Sub VPLZERO_SOLVER()
    Application.Run "Solver.xlam!SolverReset"
    Application.Run "Solver.xlam!SolverOk", Range("D45"), 3, 0, Range("H41")
    Application.Run "Solver.xlam!SolverSolve", True

    Application.Run "Solver.xlam!SolverReset"
    Application.Run "Solver.xlam!SolverOk", Range("D37"), 3, 0, Range("G41")
    Application.Run "Solver.xlam!SolverSolve", True
End Sub

Sub AtivoFinanc()
    Range("D36").Select
    Application.CutCopyMode = False
    Sheets("Resultados").Range("D36").GoalSeek Goal:=0, ChangingCell:=Sheets("Resultados").Range("G42")
    Range("A1").Select
End Sub

Sub CalcularContraprestacaoSolver()
    Dim ws As Worksheet
    Dim startTime As Double

    Set ws = ThisWorkbook.Sheets("Resultados")
    ws.Range("H37").Value = ws.Range("D50").Value

    startTime = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim i As Long
    For i = 1 To 5
        ws.Range("D13").GoalSeek Goal:=0, ChangingCell:=ws.Range("H42")
    Next i

    WriteSingleResult ws, ws.Range("B44"), "CP", startTime
    Exit Sub

ErrorHandler:
    HandleCalculationError
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

    searchValue = Range("I34").Value

    Set searchRange = Range("E41:E43")
    Set sourceRange1 = Range("H40")
    Set sourceRange2 = Range("G40")
    Set sourceRange3 = Range("H34")
    Set destinationRange1 = Range("H41:H43")
    Set destinationRange2 = Range("G41:G43")
    Set destinationRange3 = Range("F41:F43")

    For i = 1 To searchRange.Rows.Count
        If searchRange.Cells(i).Value = searchValue Then
            destinationRange1.Cells(i).Value = sourceRange1.Value
            sourceRange2.Copy
            destinationRange2.Cells(i).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            destinationRange3.Cells(i).Value = sourceRange3.Value
            Exit For
        End If
    Next i
End Sub

Sub Rodar_Cenarios()
    Dim scenarioRange As Range
    Dim scenarioCell As Range

    DoEvents
    Sheets("Resultados").Range("H35").Value = Sheets("Resultados").Range("D47").Value
    DoEvents

    Set scenarioRange = Sheets("Resultados").Range("H47:H49")

    For Each scenarioCell In scenarioRange
        Sheets("Resultados").Range("H34").Value = scenarioCell.Value
        Application.Run "PPP_ESCOLAS_SP_v9.xlsm!CalcularContraprestacao"
        Application.Run "PPP_ESCOLAS_SP_v9.xlsm!Copy_to_table"
    Next scenarioCell

    Sheets("Resultados").Range("H35").Value = Sheets("Resultados").Range("D48").Value
End Sub

Sub Rodar_Contra_TJ()
    Dim ws As Worksheet
    Dim startTime As Double
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Resultados")
    ws.Range("H37").Value = ws.Range("D50").Value

    startTime = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    For i = 1 To 5
        ws.Range("D13").GoalSeek Goal:=0, ChangingCell:=ws.Range("H42")
        ws.Range("D36").GoalSeek Goal:=0, ChangingCell:=ws.Range("G42")
    Next i

    WriteSingleResult ws, ws.Range("B46"), "CP+TJ", startTime
    Exit Sub

ErrorHandler:
    HandleCalculationError
End Sub

Sub CP_Fixa()
    Dim wsRes As Worksheet
    Dim startTime As Double

    Set wsRes = ThisWorkbook.Sheets("Resultados")

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    startTime = Timer

    wsRes.Range("H42").Value = 0
    wsRes.Range("H38").Value = 0
    wsRes.Range("D13").GoalSeek Goal:=0, ChangingCell:=wsRes.Range("F42")
    wsRes.Range("H38").Value = 1

    WriteSingleResult wsRes, wsRes.Range("B45"), "PF Max.", startTime
    Exit Sub

ErrorHandler:
    HandleCalculationError
End Sub

Sub Rodar_TJ()
    Dim ws As Worksheet
    Dim startTime As Double

    Set ws = ThisWorkbook.Sheets("Resultados")
    ws.Range("H37").Value = ws.Range("D50").Value

    startTime = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ws.Range("D36").GoalSeek Goal:=0, ChangingCell:=ws.Range("G42")

    WriteSingleResult ws, ws.Range("B45"), "TJ", startTime
    Exit Sub

ErrorHandler:
    HandleCalculationError
End Sub

Private Sub RunResultadoGoalSeek(ByVal resultLabel As String, _
                                 ByVal calculateCP As Boolean, _
                                 ByVal calculateTJ As Boolean, _
                                 ByVal iterations As Long)
    Dim ws As Worksheet
    Dim startTime As Double
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(RESULTADOS_SHEET)
    PrepareResultadosMode ws

    startTime = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    For i = 1 To iterations
        If calculateCP Then
            ws.Range("D45").GoalSeek Goal:=0, ChangingCell:=ws.Range("H41")
        End If

        If calculateTJ Then
            ws.Range("D46").GoalSeek Goal:=0, ChangingCell:=ws.Range("G41")
        End If
    Next i

    FinishResultadoCalculation ws, resultLabel, startTime
    Exit Sub

ErrorHandler:
    HandleCalculationError
End Sub

Private Sub PrepareResultadosMode(ByVal ws As Worksheet)
    ws.Range(CALC_MODE_CELL).Value = ws.Range(CALC_MODE_OPTION_CELL).Value
End Sub

Private Sub FinishResultadoCalculation(ByVal ws As Worksheet, _
                                       ByVal resultLabel As String, _
                                       ByVal startTime As Double)
    ws.Range("B60:B73").Value = ws.Range("B59:B72").Value
    ws.Range("B74").ClearContents
    ws.Range(RESULT_HISTORY_CELL).Value = BuildResultMessage(resultLabel, startTime)

    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")
    Application.ScreenUpdating = True
End Sub

Private Sub WriteSingleResult(ByVal ws As Worksheet, _
                              ByVal resultCell As Range, _
                              ByVal resultLabel As String, _
                              ByVal startTime As Double)
    resultCell.Value = BuildResultMessage(resultLabel, startTime)

    ws.Calculate
    Beep
    Application.Wait Now + TimeValue("00:00:01")
    Application.ScreenUpdating = True
End Sub

Private Function BuildResultMessage(ByVal resultLabel As String, ByVal startTime As Double) As String
    BuildResultMessage = resultLabel & ": Resultado em " & _
                         Format((Timer - startTime) / 86400, "hh:mm:ss") & _
                         " - " & Format(Now, "dd/mm/yyyy hh:mm:ss")
End Function

Private Sub HandleCalculationError()
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & Err.Description, vbExclamation
End Sub

