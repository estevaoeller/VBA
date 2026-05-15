Attribute VB_Name = "RodapeNormalizar-fontefgv"

Sub MACRO_RODAPE_NORMALIZAR()
Attribute MACRO_RODAPE_NORMALIZAR.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.MACRO_RODAPE_NORMALIZAR"
'
' MACRO_RODAPE_NORMALIZAR Macro
'
'
    ' ============================================================
    ' Objetivo:
    ' Inserir o texto da fonte e normalizar a formatação do parágrafo
    ' de rodapé (geralmente usado logo abaixo de tabelas ou figuras).
    '
    ' Ações realizadas pela macro:
    ' 1. Digita automaticamente o texto padronizado (ex: "Fonte: Elaboração Própria").
    ' 2. Aplica a fonte e o tamanho configurados (ex: Arial Narrow, 9).
    ' 3. Centraliza o parágrafo horizontalmente.
    ' 4. Aplica os espaçamentos antes e depois do texto (ex: 0pt e 12pt).
    ' 5. Remove recuos de margem e limpa formatações indesejadas de quebra de página.
    ' ============================================================

    ' --- DECLARAÇÃO DAS VARIÁVEIS ---
    Dim textoFonte As String
    Dim espacamentoAntes As Single
    Dim espacamentoDepois As Single
    Dim alinhamentoParagrafo As Long
    Dim tamanhoFonte As Single
    Dim nomeFonte As String
    
    ' --- CONFIGURAÇÕES PRINCIPAIS (Altere os valores aqui) ---
    textoFonte = "Fonte: Elaboração Própria"
    espacamentoAntes = 0
    espacamentoDepois = 12
    alinhamentoParagrafo = wdAlignParagraphCenter
    tamanhoFonte = 9
    nomeFonte = "Arial Narrow"
    ' ---------------------------------------------------------
    
    ' Aplica a fonte e digita o texto
    Selection.Font.Name = nomeFonte
    Selection.Font.Size = tamanhoFonte
    Selection.TypeText Text:=textoFonte

    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = espacamentoAntes
        .SpaceBeforeAuto = False
        .SpaceAfter = espacamentoDepois
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = alinhamentoParagrafo
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
