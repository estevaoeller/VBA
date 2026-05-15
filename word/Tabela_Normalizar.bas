Attribute VB_Name = "Tabela_Normalizar"

Sub Macro_TABELA_NORMALIZAR()
Attribute Macro_TABELA_NORMALIZAR.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro_TABELA_NORMALIZAR"
'
' Macro_TABELA_NORMALIZAR Macro
'

    ' ============================================================
    ' MACRO_TABELA_NORMALIZAR
    '
    ' Objetivo:
    ' Normalizar a tabela onde o cursor estiver posicionado,
    ' aplicando padr�o de fonte, alinhamento, espa�amento,
    ' altura das linhas e formata��o do cabe�alho.
    '
    ' Regras aplicadas:
    ' 1. Atua sobre a primeira tabela encontrada na sele��o.
    ' 2. Remove it�lico de toda a tabela.
    ' 3. Define fonte Arial Narrow, tamanho 9.
    ' 4. Centraliza horizontalmente o conte�do das c�lulas.
    ' 5. Centraliza verticalmente o conte�do das c�lulas.
    ' 6. Define altura de 0,5 cm para todas as linhas.
    ' 7. Define altura de 0,9 cm para a primeira linha.
    ' 8. Define a primeira linha como cabe�alho repetido.
    ' 9. Aplica negrito ao cabe�alho.
    ' ============================================================

    Dim tbl As Table

    ' ------------------------------------------------------------
    ' 1. Verifica se o cursor ou a sele��o est� dentro de uma tabela
    ' ------------------------------------------------------------
    ' Se o cursor n�o estiver dentro de uma tabela, a macro � interrompida.
    ' Isso evita erro ao tentar aplicar comandos de tabela fora de uma tabela.
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "Coloque o cursor dentro da tabela ou selecione a tabela antes de executar a macro.", vbExclamation
        Exit Sub
    End If

    ' ------------------------------------------------------------
    ' 2. Identifica a tabela que será formatada
    ' ------------------------------------------------------------
    ' Selection.Tables(1) pega a primeira tabela associada à seleção atual.
    ' Se o cursor estiver dentro de uma tabela, essa será a tabela usada.
    Set tbl = Selection.Tables(1)

    ' ------------------------------------------------------------
    ' 3. Normaliza a formatação dos parágrafos da tabela inteira
    ' ------------------------------------------------------------
    ' Esta parte remove recuos, espaçamentos antes/depois,
    ' define espaçamento simples e centraliza os parágrafos.
    With tbl.Range.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)

        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False

        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter

        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .OutlineLevel = wdOutlineLevelBodyText
    End With

    ' ------------------------------------------------------------
    ' 4. Define a fonte da tabela inteira
    ' ------------------------------------------------------------
    ' Aplica Arial Narrow, tamanho 9,
    ' remove itálico e remove negrito da tabela inteira.
    ' O negrito será reaplicado depois apenas ao cabeçalho.
    With tbl.Range.Font
        .Name = "Arial Narrow"
        .Size = 9
        .Italic = False
        .Bold = False
    End With

    ' ------------------------------------------------------------
    ' 5. Centraliza verticalmente o conteúdo das células
    ' ------------------------------------------------------------
    ' Faz com que o texto fique no centro vertical da célula,
    ' e não colado na parte superior ou inferior.
    tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter

    ' ------------------------------------------------------------
    ' 6. Define a altura padrão das linhas da tabela
    ' ------------------------------------------------------------
    ' Aplica altura exata de 0,5 cm para todas as linhas.
    ' Como é altura exata, o Word pode cortar conteúdo se houver texto demais.
    With tbl.Rows
        .HeightRule = wdRowHeightAtLeast
        .Height = CentimetersToPoints(0.5)
    End With

    ' ------------------------------------------------------------
    ' 7. Formata a primeira linha como cabeçalho da tabela
    ' ------------------------------------------------------------
    ' A primeira linha recebe tratamento específico:
    ' - repete como cabeçalho nas páginas seguintes;
    ' - tem altura de 0,9 cm;
    ' - fica em negrito;
    ' - permanece sem itálico;
    ' - fica centralizada horizontal e verticalmente.
    With tbl.Rows(1)
        .HeadingFormat = True
        .HeightRule = wdRowHeightExactly
        .Height = CentimetersToPoints(0.9)

        .Range.Font.Bold = True
        .Range.Font.Italic = False

        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.SpaceAfter = 0
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        .Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With

    
    
End Sub
