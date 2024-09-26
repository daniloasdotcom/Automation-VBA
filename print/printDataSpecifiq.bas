Attribute VB_Name = "Módulo3"
Sub GerarPaginasParaIntervaloDeRegistrosPorColunaAPretoEBranco()
    Dim wsDados As Worksheet
    Dim wsImpressao As Worksheet
    Dim linhaDados As Range
    Dim linhaCabecalho As Range
    Dim i As Integer, linhaAtual As Long
    Dim ultimaLinha As Long
    Dim registroInicial As Long
    Dim registroFinal As Long
    Dim celulaRegistroInicial As Range
    Dim celulaRegistroFinal As Range
    
    ' Definir a aba de dados e a aba de impressão
    Set wsDados = ThisWorkbook.Sheets("Banco de Dados") ' Nome correto da aba de dados
    Set wsImpressao = ThisWorkbook.Sheets("Aba de Impressão") ' Nome correto da aba de impressão
    
    ' Selecionar a linha do cabeçalho (assumindo que o cabeçalho está na linha 3, começando na coluna B)
    Set linhaCabecalho = wsDados.Range("B3:W3") ' Cabeçalho agora vai de B a W
    
    ' Encontrar a última linha de dados na aba "Banco de Dados"
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    
    ' Ler o intervalo de registros inserido nas células G3 (registro inicial) e G4 (registro final) da Aba de Impressão
    If IsNumeric(wsImpressao.Range("G3").Value) And IsNumeric(wsImpressao.Range("G4").Value) Then
        registroInicial = CLng(wsImpressao.Range("G3").Value)
        registroFinal = CLng(wsImpressao.Range("G4").Value)
    Else
        MsgBox "Por favor, insira valores numéricos válidos nas células G3 e G4 da Aba de Impressão.", vbCritical
        Exit Sub
    End If
    
    ' Verificar se o registro inicial e final estão na coluna A
    Set celulaRegistroInicial = wsDados.Columns("A").Find(What:=registroInicial, LookIn:=xlValues, LookAt:=xlWhole)
    Set celulaRegistroFinal = wsDados.Columns("A").Find(What:=registroFinal, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Verificar se os registros foram encontrados
    If celulaRegistroInicial Is Nothing Or celulaRegistroFinal Is Nothing Then
        MsgBox "Intervalo inválido. Certifique-se de que os registros existem na coluna A.", vbCritical
        Exit Sub
    End If
    
    ' Validar o intervalo inserido
    If registroInicial < 1 Or registroFinal > wsDados.Cells(ultimaLinha, "A").Value Or registroInicial > registroFinal Then
        MsgBox "Intervalo inválido. Por favor, insira um intervalo válido entre 1 e " & wsDados.Cells(ultimaLinha, "A").Value, vbCritical
        Exit Sub
    End If
    
    ' Loop para percorrer as linhas de dados dentro do intervalo inserido (usando as referências das células encontradas na coluna A)
    For linhaAtual = celulaRegistroInicial.Row To celulaRegistroFinal.Row
        ' Limpar a aba de impressão (somente os dados, preservando formatação)
        wsImpressao.Range("B6:C" & (5 + linhaCabecalho.Columns.Count)).ClearContents
        
        ' Selecionar a linha de dados atual (a partir da coluna B agora)
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
        MsgBox "Dados do registro " & wsDados.Cells(linhaAtual, "A").Value & " foram gerados e impressos!", vbInformation
    Next linhaAtual
End Sub


