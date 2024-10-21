Attribute VB_Name = "Módulo10"
Sub ValidarVendasAgendadoEDistintos()
    Dim wsBase As Worksheet
    Dim wsAgendado As Worksheet
    Dim ultimaLinhaBase As Long, ultimaLinhaAgendado As Long
    Dim i As Long, j As Long
    Dim pdvBase As String, produtosBase As String
    Dim pdvAgendado As String, produtoAgendado As String
    Dim quantidadeVendida As Double
    Dim produtosDistintosVendidos As Object
    Dim missao As String
    Dim dictPDV As Object
    Dim produtosArray() As String
    
    ' Inicializar dicionários
    Set dictPDV = CreateObject("Scripting.Dictionary")
    Set produtosDistintosVendidos = CreateObject("Scripting.Dictionary")
    
    ' Definir as planilhas
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    Set wsAgendado = ActiveWorkbook.Worksheets("Agendado")
    
    ' Encontrar a última linha de ambas as planilhas
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaAgendado = wsAgendado.Cells(wsAgendado.Rows.Count, "A").End(xlUp).Row
    
    ' Copiar dados das planilhas para arrays
    Dim baseData As Variant
    Dim agendadoData As Variant
    
    baseData = wsBase.Range("A2:O" & ultimaLinhaBase).Value ' Copiar dados da Base
    agendadoData = wsAgendado.Range("A2:V" & ultimaLinhaAgendado).Value ' Copiar dados de Agendado
    
    ' Criar dicionário para armazenar as vendas do Agendado por PDV e Produto
    For j = 1 To UBound(agendadoData, 1)
        pdvAgendado = agendadoData(j, 1) ' PDV do Agendado
        produtoAgendado = agendadoData(j, 19) ' Produto do Agendado (Coluna S)
        
        If Not dictPDV.exists(pdvAgendado) Then
            Set dictPDV(pdvAgendado) = CreateObject("Scripting.Dictionary")
        End If
        
        If Not dictPDV(pdvAgendado).exists(produtoAgendado) Then
            dictPDV(pdvAgendado)(produtoAgendado) = agendadoData(j, 22) ' Quantidade vendida na coluna V
        Else
            dictPDV(pdvAgendado)(produtoAgendado) = dictPDV(pdvAgendado)(produtoAgendado) + agendadoData(j, 22)
        End If
    Next j
    
    ' Processar os dados da Base
    For i = 1 To UBound(baseData, 1)
        pdvBase = baseData(i, 6) ' PDV da base
        produtosBase = baseData(i, 4) ' Produtos da base (separados por vírgula)
        missao = baseData(i, 3) ' Missão na coluna C da Base
        
        ' Separar os produtos da base
        produtosArray = Split(produtosBase, ",")
        
        ' Inicializar a quantidade de venda e produtos distintos
        quantidadeVendida = 0
        produtosDistintosVendidos.RemoveAll
        
        ' Verificar se o PDV existe no dicionário
        If dictPDV.exists(pdvBase) Then
            For Each produto In produtosArray
                produto = Trim(produto)
                If dictPDV(pdvBase).exists(produto) Then
                    quantidadeVendida = quantidadeVendida + dictPDV(pdvBase)(produto)
                    
                    ' Adicionar produto à lista de distintos
                    If Not produtosDistintosVendidos.exists(produto) Then
                        produtosDistintosVendidos.Add produto, Nothing
                    End If
                End If
            Next produto
        End If
        
        ' Preencher a quantidade vendida na coluna J
        wsBase.Cells(i + 1, 10).Value = quantidadeVendida
        
        ' Verificar se a missão contém "distinto" ou "distintos"
        If InStr(missao, "distinto") > 0 Or InStr(missao, "distintos") > 0 Then
            ' Preencher a coluna N com a quantidade de produtos distintos
            wsBase.Cells(i + 1, 14).Value = produtosDistintosVendidos.Count
            
            ' Preencher a coluna O com a lista de produtos distintos vendidos
            If produtosDistintosVendidos.Count > 0 Then
                Dim listaProdutos As String
                listaProdutos = Join(produtosDistintosVendidos.keys, ", ")
                wsBase.Cells(i + 1, 15).Value = listaProdutos
            Else
                wsBase.Cells(i + 1, 15).Value = 0
            End If
        Else
            ' Se a missão não contiver "distinto" ou "distintos", preencher a coluna N com 0
            wsBase.Cells(i + 1, 14).Value = 0
        End If
    Next i
    
    ' Informar que a macro terminou
    MsgBox "Vendas agendadas concluídas com sucesso!"
End Sub

