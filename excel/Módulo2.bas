Attribute VB_Name = "Módulo2"
Sub ExecutarCenariosWACC()

    Dim i As Integer
    Dim x As Integer
    Dim resultado As Range
    Dim ws As Worksheet
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    Dim formattedTime As String
    Dim currentDateTime As String
    Dim totalStartTime As Double
    Dim totalEndTime As Double
    Dim totalElapsedTime As Double
    Dim formattedTotalTime As String
    
    ' Define a planilha ativa
    Set ws = ActiveSheet
    
    ' Pergunte ao usuˆrio o nìmero mˆximo de cenˆrios a serem executados
    x = InputBox("Insira o nìmero mˆximo de cenˆrios a serem executados:", "Nìmero de Cenˆrios")
    
    ' Inicie a contagem de tempo total
    totalStartTime = Timer
    
    ' Certifique-se de que x _ um nìmero vˆlido
    If IsNumeric(x) And x > 0 Then
        For i = 1 To x
            ' Defina a c_lula Y60 com o nìmero do cenˆrio atual
            ws.Cells(60, 25).Value = i ' Coluna Y _ a 25é coluna
            
            ' Inicie a contagem de tempo
            startTime = Timer
            
            ' Execute a macro Rodar_Contra_TJ
            Dim j As Integer
            For j = 1 To 5
                ' GoalSeek para D13
                ws.Range("D13").GoalSeek Goal:=0, ChangingCell:=ws.Range("H42")
                
                ' GoalSeek para D36
                ws.Range("D36").GoalSeek Goal:=0, ChangingCell:=ws.Range("G42")
            Next j
            
            ' Pare a contagem de tempo
            endTime = Timer
            elapsedTime = endTime - startTime
            
            ' Formate o tempo decorrido como hh:mm:ss
            formattedTime = Format(elapsedTime \ 3600, "00") & ":" & _
                            Format((elapsedTime Mod 3600) \ 60, "00") & ":" & _
                            Format(elapsedTime Mod 60, "00")
            
            ' Obtenha a data e hora atuais
            currentDateTime = Format(Now, "dd/mm/yyyy hh:mm:ss")
            
            ' Registre o tempo de execu†o na coluna AE da linha correspondente ao cenˆrioo
            ws.Cells(60 + i, 31).Value = formattedTime & " - " & currentDateTime ' Coluna AE _ a 31é coluna
            
            ' Copie o resultado das c_lulas AB60:AD60
            Set resultado = ws.Range("AB60:AD60")
            
            ' Cole o resultado como valores na linha correta correspondente ao cenˆrio atual
            ws.Range("AB" & 60 + i & ":AD" & 60 + i).Value = resultado.Value
        Next i
        
        ' Pare a contagem de tempo total
        totalEndTime = Timer
        totalElapsedTime = totalEndTime - totalStartTime
        
        ' Formate o tempo total decorrido como hh:mm:ss
        formattedTotalTime = "Tempo total: " & Format(totalElapsedTime \ 3600, "00") & ":" & _
                             Format((totalElapsedTime Mod 3600) \ 60, "00") & ":" & _
                             Format(totalElapsedTime Mod 60, "00")
                             
        ' Registre o tempo total de execu†o em AE600
        ws.Cells(60, 31).Value = formattedTotalTime & " - " & currentDateTime
    Else
        MsgBox "Por favor, insira um nìmero vˆlido de cenˆrios.", vbExclamation
    End If
    
End Sub


