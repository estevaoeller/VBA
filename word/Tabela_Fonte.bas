Attribute VB_Name = "RodapeNormalizar"

Sub MACRO_RODAPE_NORMALIZAR()
Attribute MACRO_RODAPE_NORMALIZAR.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.MACRO_RODAPE_NORMALIZAR"
'
' MACRO_RODAPE_NORMALIZAR Macro
'
'
    ' ============================================================
    ' Objetivo:
    ' Normalizar a formatação de parágrafos usados como rodapé,
    ' fonte ou notas de rodapé abaixo de tabelas e figuras.
    '
    ' Ações realizadas:
    ' 1. Remove recuos de margem (esquerda, direita e primeira linha).
    ' 2. Define o espaçamento antes do texto como 0pt e depois como 12pt.
    ' 3. Aplica espaçamento entre linhas Simples e centraliza o texto.
    ' 4. Remove formatações indesejadas de quebra de página e fluxo de texto.
    ' ============================================================

    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 12
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With

End Sub
