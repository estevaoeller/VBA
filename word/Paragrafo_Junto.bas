Attribute VB_Name = "JuntoProx"

Sub MACRO_JUNTO_PROX()
Attribute MACRO_JUNTO_PROX.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.MACRO_JUNTO_PROX"
'
' MACRO_JUNTO_PROX Macro
'
'
   ' ============================================================
    ' Objetivo:
    ' Aplica a formatação "Manter com o próximo" ao parágrafo
    ' selecionado, evitando que ele fique isolado no final da página,
    ' separado do texto ou elemento seguinte (ideal para títulos).
    '
    ' Ações realizadas:
    ' 1. Desativa o espaçamento automático antes do parágrafo.
    ' 2. Desativa o espaçamento automático depois do parágrafo.
    ' 3. Habilita a propriedade "Manter com o próximo" (KeepWithNext).
    ' ============================================================

    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .KeepWithNext = True
    End With
End Sub
