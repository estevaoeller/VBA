Attribute VB_Name = "Tabela_Normalizar_V1"




Sub Macro_TABELA_NORMALIZAR_V0()
'
' Macro_TABELA_NORMALIZAR Macro
'
'
    ' ============================================================
    ' Objetivo:
    ' Normalizar a formatação de parágrafos e células de uma tabela,
    ' padronizando recuos, espaçamentos, alinhamento e tipo de fonte.
    '
    ' Ações realizadas:
    ' 1. Zera os recuos de margem e o espaçamento antes/depois do parágrafo.
    ' 2. Aplica espaçamento simples e remove quebras de página indesejadas.
    ' 3. Seleciona a célula onde o cursor está posicionado.
    ' 4. Centraliza o conteúdo horizontalmente e verticalmente.
    ' 5. Aplica a fonte Arial Narrow, tamanho 9.
    ' 6. Garante que qualquer formatação em itálico seja removida/alternada.
    ' ============================================================

    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Font.Size = 9
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Italic = wdToggle
    Selection.Font.Italic = wdToggle
    
    
    
    
    
    
    
    
    
    
End Sub
