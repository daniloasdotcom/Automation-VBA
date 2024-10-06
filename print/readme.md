# VBA Print Scripts

Este repositório contém uma coleção de códigos VBA relacionados a tarefas de impressão no Excel. Cada script serve a um propósito específico relacionado à impressão de dados organizados em planilhas. Abaixo estão as descrições dos arquivos e suas funcionalidades.

## Códigos

### 1. `printDataSpecifiq.bas`
- **Descrição**: Gera páginas de impressão com base em um intervalo específico de registros inserido pelo usuário.
- **Funcionalidade**: 
    - O usuário insere o registro inicial e final nas células `G3` e `G4` na aba de impressão.
    - O código copia os dados de uma planilha chamada "Banco de Dados" e imprime na aba "Aba de Impressão".
    - Imprime as páginas no formato preto e branco.
- **Útil para**: Imprimir apenas um intervalo específico de registros de uma base de dados.

### 2. `printDataAll.bas`
- **Descrição**: Percorre todas as linhas de uma planilha e gera uma página de impressão para cada linha.
- **Funcionalidade**: 
    - Copia o cabeçalho e os dados de cada linha da planilha "Banco de Dados".
    - Imprime cada conjunto de dados na aba de impressão, uma linha por vez.
- **Útil para**: Imprimir todas as linhas de uma planilha de dados em um formato padronizado.

### 3. `printData.bas`
- **Descrição**: Imprime todas as páginas de uma planilha de dados.
- **Funcionalidade**: 
    - Automatiza o processo de copiar cabeçalhos e dados para uma área de impressão.
    - Define uma área de impressão fixa e imprime todas as páginas de acordo com os dados da planilha.
- **Útil para**: Automação completa de impressão de uma planilha inteira.

## Como usar
1. Baixe o arquivo desejado.
2. Importe o arquivo `.bas` no editor VBA (Alt + F11 no Excel).
3. Execute a macro para o processo de impressão desejado.

## Contribuições
Sinta-se à vontade para contribuir com melhorias ou novas funcionalidades abrindo uma pull request.

