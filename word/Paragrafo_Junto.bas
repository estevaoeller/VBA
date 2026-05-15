Attribute VB_Name = "JuntoProx"

Sub MACRO_JUNTO_PROX()
Attribute MACRO_JUNTO_PROX.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.MACRO_JUNTO_PROX"
'
' MACRO_JUNTO_PROX Macro
'
'
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .KeepWithNext = True
    End With
End Sub

