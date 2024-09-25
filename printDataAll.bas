Attribute VB_Name = "M�dulo1"
Sub GerarPaginasParaTodosRegistrosPreservandoFormatacaoPretoEBranco()
    Dim wsDados As Worksheet
    Dim wsImpressao As Worksheet
    Dim linhaDados As Range
    Dim linhaCabecalho As Range
    Dim i As Integer, linhaAtual As Integer
    Dim ultimaLinha As Long
    
    ' Definir a aba de dados e a aba de impress�o
    Set wsDados = ThisWorkbook.Sheets("Banco de Dados") ' Nome correto da aba de dados
    Set wsImpressao = ThisWorkbook.Sheets("Aba de Impress�o") ' Nome correto da aba de impress�o
    
    ' Selecionar a linha do cabe�alho (assumindo que o cabe�alho est� na linha 3, come�ando na coluna B)
    Set linhaCabecalho = wsDados.Range("B3:W3") ' Cabe�alho agora vai de B a W
    
    ' Encontrar a �ltima linha de dados na aba "Banco de Dados"
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    
    ' Loop para percorrer todas as linhas de dados (a partir da linha 4 at� a �ltima linha)
    For linhaAtual = 4 To ultimaLinha
        ' Limpar a aba de impress�o (somente os dados, preservando formata��o)
        wsImpressao.Range("B6:C" & (5 + linhaCabecalho.Columns.Count)).ClearContents
        
        ' Selecionar a linha de dados atual (a partir da coluna B at� a coluna W agora)
        Set linhaDados = wsDados.Range("B" & linhaAtual & ":W" & linhaAtual)
        
        ' Copiar os cabe�alhos e os dados para a aba de impress�o
        For i = 1 To linhaCabecalho.Columns.Count
            ' Inserir o cabe�alho na coluna 2 (B) da aba de impress�o
            wsImpressao.Cells(6 + (i - 1), 2).Value = linhaCabecalho.Cells(1, i).Value
            
            ' Inserir os dados correspondentes na coluna 3 (C) da aba de impress�o
            wsImpressao.Cells(6 + (i - 1), 3).Value = linhaDados.Cells(1, i).Value
        Next i
        
        ' Definir a nova �rea de impress�o (de A1 at� D28)
        wsImpressao.PageSetup.PrintArea = wsImpressao.Range("A1:D28").Address
        
        ' Configurar impress�o em preto e branco
        wsImpressao.PageSetup.BlackAndWhite = True
        
        ' Imprimir a p�gina
        wsImpressao.PrintOut
        
        ' Exibir mensagem informando o registro processado (pode ser removido ou ajustado conforme desejado)
        MsgBox "Dados da linha " & linhaAtual - 3 & " foram gerados e impressos!", vbInformation
    Next linhaAtual
End Sub



