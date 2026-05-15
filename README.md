# Macros 365

Colecao de macros VBA para automatizar rotinas recorrentes no Microsoft 365, com foco em modelos financeiros no Excel e padronizacao de documentos no Word.

O projeto reune modulos `.bas` versionaveis e um add-in Excel (`MacrosEE.xlam`), facilitando manutencao, reaproveitamento e auditoria das macros usadas no dia a dia.

## O que tem aqui

### Excel

Macros voltadas a modelos financeiros e planilhas de analise:

- Calculo por `GoalSeek` para zerar VPL, contraprestacao e taxa.
- Rotinas com Solver para cenarios em que ha mais de uma variavel de decisao.
- Execucao em lote de cenarios de WACC, com registro de tempo e consolidacao de resultados.
- Copia automatica de resultados para tabelas de controle.
- Auditoria de dependentes em intervalos especificos, incluindo dependencias na mesma aba e em outras abas.

### Word

Macros para acelerar formatacao padronizada de relatorios:

- Normalizacao de tabelas com fonte, alinhamento, espacamento e cabecalho repetido.
- Insercao e formatacao de fonte/rodape abaixo de tabelas e figuras.
- Aplicacao de "manter com o proximo" em paragrafos, util para titulos e elementos que nao devem ficar separados.

## Estrutura do repositorio

```text
.
|-- excel/
|   |-- auditoria.bas
|   |-- Modulo1.bas
|   `-- Modulo2.bas
|-- word/
|   |-- Paragrafo_Junto.bas
|   |-- Tabela_Fonte.bas
|   |-- Tabela_Fonte_FGV.bas
|   |-- Tabela_Normalizar.bas
|   `-- Tabela_Normalizar_v0.bas
|-- MacrosEE.xlam
`-- README.md
```

## Como usar

### Importar os modulos VBA

1. Abra o Excel ou Word.
2. Pressione `Alt + F11` para abrir o Editor do VBA.
3. No projeto desejado, clique com o botao direito em `Modules`.
4. Selecione `Import File...`.
5. Escolha o arquivo `.bas` correspondente.
6. Salve o arquivo como habilitado para macros (`.xlsm`, `.docm`) ou instale como add-in quando aplicavel.

### Usar o add-in do Excel

1. Abra o Excel.
2. Acesse `File > Options > Add-ins`.
3. Em `Manage: Excel Add-ins`, clique em `Go...`.
4. Clique em `Browse...` e selecione `MacrosEE.xlam`.
5. Ative o add-in na lista.

> Algumas macros dependem de nomes especificos de abas e celulas, como `Resultados`, `Opex_UE`, `Change Log`, `D45`, `H41`, entre outras. Antes de rodar em outra planilha, confira se o layout esperado existe.

## Principais macros

| Arquivo | Macro | Funcao |
| --- | --- | --- |
| `excel/Modulo1.bas` | `CALC_CP_GS`, `CALC_TJ_GS`, `CALC_CPeTJ_GS` | Calculam contraprestacao e/ou taxa por GoalSeek. |
| `excel/Modulo1.bas` | `CALC_CP_SV`, `CALC_TJ_SV`, `CALC_CPeTJ_SV` | Executam calculos equivalentes com apoio do Solver. |
| `excel/Modulo1.bas` | `Rodar_Cenarios` | Percorre cenarios e copia resultados para tabela. |
| `excel/Modulo2.bas` | `ExecutarCenariosWACC` | Executa multiplos cenarios WACC e registra tempo/resultados. |
| `excel/auditoria.bas` | `AuditarDependentes` | Gera uma aba `Auditoria` com dependentes internos e externos. |
| `word/Tabela_Normalizar.bas` | `Macro_TABELA_NORMALIZAR` | Padroniza a tabela selecionada ou onde o cursor estiver. |
| `word/Tabela_Fonte.bas` | `MACRO_RODAPE_NORMALIZAR` | Normaliza paragrafos de fonte/rodape. |
| `word/Tabela_Fonte_FGV.bas` | `MACRO_RODAPE_NORMALIZAR` | Insere `Fonte: Elaboracao Propria` e aplica formatacao. |
| `word/Paragrafo_Junto.bas` | `MACRO_JUNTO_PROX` | Aplica "manter com o proximo" ao paragrafo selecionado. |

## Requisitos

- Microsoft Excel e/ou Word com suporte a VBA.
- Macros habilitadas no arquivo ou add-in.
- Solver instalado/ativado no Excel para as rotinas que usam `SolverReset`, `SolverOk`, `SolverAdd` e `SolverSolve`.

Para ativar o Solver:

1. No Excel, va em `File > Options > Add-ins`.
2. Em `Manage: Excel Add-ins`, clique em `Go...`.
3. Marque `Solver Add-in`.
4. Confirme em `OK`.

## Cuidados antes de executar

- Faca uma copia da planilha/documento antes de rodar macros que alteram celulas, tabelas ou formatacao.
- Verifique se os nomes das abas e os enderecos de celulas batem com o modelo usado.
- Em macros com `GoalSeek` ou Solver, confirme se as formulas estao calculando corretamente antes da execucao em lote.
- Evite rodar macros de auditoria em intervalos muito grandes sem testar primeiro, pois isso pode deixar o Excel lento.

## Desenvolvimento

Os arquivos `.bas` sao a fonte principal para controle de versao. Ao alterar uma macro pelo Editor do VBA, exporte novamente o modulo correspondente para manter o repositorio atualizado.

Fluxo sugerido:

1. Edite/teste a macro no Excel ou Word.
2. Exporte o modulo `.bas`.
3. Revise o diff no Git.
4. Atualize este README se a macro mudar de comportamento.

## Licenca

Uso interno/pessoal. Defina uma licenca formal caso o projeto seja compartilhado publicamente.
