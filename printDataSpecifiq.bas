Attribute VB_Name = "M�dulo3"
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
    
    ' Definir a aba de dados e a aba de impress�o
    Set wsDados = ThisWorkbook.Sheets("Banco de Dados") ' Nome correto da aba de dados
    Set wsImpressao = ThisWorkbook.Sheets("Aba de Impress�o") ' Nome correto da aba de impress�o
    
    ' Selecionar a linha do cabe�alho (assumindo que o cabe�alho est� na linha 3, come�ando na coluna B)
    Set linhaCabecalho = wsDados.Range("B3:W3") ' Cabe�alho agora vai de B a W
    
    ' Encontrar a �ltima linha de dados na aba "Banco de Dados"
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    
    ' Ler o intervalo de registros inserido nas c�lulas G3 (registro inicial) e G4 (registro final) da Aba de Impress�o
    If IsNumeric(wsImpressao.Range("G3").Value) And IsNumeric(wsImpressao.Range("G4").Value) Then
        registroInicial = CLng(wsImpressao.Range("G3").Value)
        registroFinal = CLng(wsImpressao.Range("G4").Value)
    Else
        MsgBox "Por favor, insira valores num�ricos v�lidos nas c�lulas G3 e G4 da Aba de Impress�o.", vbCritical
        Exit Sub
    End If
    
    ' Verificar se o registro inicial e final est�o na coluna A
    Set celulaRegistroInicial = wsDados.Columns("A").Find(What:=registroInicial, LookIn:=xlValues, LookAt:=xlWhole)
    Set celulaRegistroFinal = wsDados.Columns("A").Find(What:=registroFinal, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Verificar se os registros foram encontrados
    If celulaRegistroInicial Is Nothing Or celulaRegistroFinal Is Nothing Then
        MsgBox "Intervalo inv�lido. Certifique-se de que os registros existem na coluna A.", vbCritical
        Exit Sub
    End If
    
    ' Validar o intervalo inserido
    If registroInicial < 1 Or registroFinal > wsDados.Cells(ultimaLinha, "A").Value Or registroInicial > registroFinal Then
        MsgBox "Intervalo inv�lido. Por favor, insira um intervalo v�lido entre 1 e " & wsDados.Cells(ultimaLinha, "A").Value, vbCritical
        Exit Sub
    End If
    
    ' Loop para percorrer as linhas de dados dentro do intervalo inserido (usando as refer�ncias das c�lulas encontradas na coluna A)
    For linhaAtual = celulaRegistroInicial.Row To celulaRegistroFinal.Row
        ' Limpar a aba de impress�o (somente os dados, preservando formata��o)
        wsImpressao.Range("B6:C" & (5 + linhaCabecalho.Columns.Count)).ClearContents
        
        ' Selecionar a linha de dados atual (a partir da coluna B agora)
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
        MsgBox "Dados do registro " & wsDados.Cells(linhaAtual, "A").Value & " foram gerados e impressos!", vbInformation
    Next linhaAtual
End Sub


