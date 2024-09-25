Attribute VB_Name = "Módulo1"
Sub GerarPaginasParaTodosRegistrosPreservandoFormatacaoPretoEBranco()
    Dim wsDados As Worksheet
    Dim wsImpressao As Worksheet
    Dim linhaDados As Range
    Dim linhaCabecalho As Range
    Dim i As Integer, linhaAtual As Integer
    Dim ultimaLinha As Long
    
    ' Definir a aba de dados e a aba de impressão
    Set wsDados = ThisWorkbook.Sheets("Banco de Dados") ' Nome correto da aba de dados
    Set wsImpressao = ThisWorkbook.Sheets("Aba de Impressão") ' Nome correto da aba de impressão
    
    ' Selecionar a linha do cabeçalho (assumindo que o cabeçalho está na linha 3, começando na coluna B)
    Set linhaCabecalho = wsDados.Range("B3:W3") ' Cabeçalho agora vai de B a W
    
    ' Encontrar a última linha de dados na aba "Banco de Dados"
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    
    ' Loop para percorrer todas as linhas de dados (a partir da linha 4 até a última linha)
    For linhaAtual = 4 To ultimaLinha
        ' Limpar a aba de impressão (somente os dados, preservando formatação)
        wsImpressao.Range("B6:C" & (5 + linhaCabecalho.Columns.Count)).ClearContents
        
        ' Selecionar a linha de dados atual (a partir da coluna B até a coluna W agora)
        Set linhaDados = wsDados.Range("B" & linhaAtual & ":W" & linhaAtual)
        
        ' Copiar os cabeçalhos e os dados para a aba de impressão
        For i = 1 To linhaCabecalho.Columns.Count
            ' Inserir o cabeçalho na coluna 2 (B) da aba de impressão
            wsImpressao.Cells(6 + (i - 1), 2).Value = linhaCabecalho.Cells(1, i).Value
            
            ' Inserir os dados correspondentes na coluna 3 (C) da aba de impressão
            wsImpressao.Cells(6 + (i - 1), 3).Value = linhaDados.Cells(1, i).Value
        Next i
        
        ' Definir a nova área de impressão (de A1 até D28)
        wsImpressao.PageSetup.PrintArea = wsImpressao.Range("A1:D28").Address
        
        ' Configurar impressão em preto e branco
        wsImpressao.PageSetup.BlackAndWhite = True
        
        ' Imprimir a página
        wsImpressao.PrintOut
        
        ' Exibir mensagem informando o registro processado (pode ser removido ou ajustado conforme desejado)
        MsgBox "Dados da linha " & linhaAtual - 3 & " foram gerados e impressos!", vbInformation
    Next linhaAtual
End Sub



