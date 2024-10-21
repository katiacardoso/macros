Attribute VB_Name = "Módulo4"
Sub ValidarQuantidadeCompradaMes()
    Dim wsBase As Worksheet
    Dim wsCompraMês As Worksheet
    Dim ultimaLinhaBase As Long, ultimaLinhaCompraMês As Long
    Dim i As Long, j As Long
    Dim pdvBase As String, produtosBase As String
    Dim pdvCompraMês As String, produtoCompraMês As String
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
    Set wsCompraMês = ActiveWorkbook.Worksheets("03.05.09")
    
    ' Encontrar a última linha de ambas as planilhas
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaCompraMês = wsCompraMês.Cells(wsCompraMês.Rows.Count, "A").End(xlUp).Row
    
    ' Copiar dados das planilhas para arrays
    Dim baseData As Variant
    Dim compraMesData As Variant
    
    baseData = wsBase.Range("A2:P" & ultimaLinhaBase).Value ' Copiar dados da Base
    compraMesData = wsCompraMês.Range("A2:G" & ultimaLinhaCompraMês).Value ' Copiar dados de CompraMês
    
    ' Percorrer a planilha CompraMês e armazenar os dados no dicionário
    For j = 1 To UBound(compraMesData, 1)
        pdvCompraMês = compraMesData(j, 1) ' PDV do CompraMês
        produtoCompraMês = compraMesData(j, 4) ' Produto do CompraMês
        
        ' Verificar se a célula contém um número, caso contrário, atribuir 0
        If IsNumeric(compraMesData(j, 7)) Then
            quantidadeVendida = compraMesData(j, 7) ' Quantidade vendida no CompraMês
        Else
            quantidadeVendida = 0 ' Se não for numérico, atribuir 0
        End If
        
        ' Armazenar no dicionário usando PDV como chave
        If Not dictPDV.exists(pdvCompraMês) Then
            Set dictPDV(pdvCompraMês) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Armazenar a quantidade vendida e os produtos distintos
        If Not dictPDV(pdvCompraMês).exists(produtoCompraMês) Then
            dictPDV(pdvCompraMês)(produtoCompraMês) = quantidadeVendida
        Else
            dictPDV(pdvCompraMês)(produtoCompraMês) = dictPDV(pdvCompraMês)(produtoCompraMês) + quantidadeVendida
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
        
        ' Preencher a quantidade vendida na coluna K
        wsBase.Cells(i + 1, 11).Value = quantidadeVendida
        
        ' Verificar se a missão contém "distinto" ou "distintos"
        If InStr(missao, "distinto") > 0 Or InStr(missao, "distintos") > 0 Then
            ' Preencher a coluna P com a quantidade de produtos distintos
            wsBase.Cells(i + 1, 16).Value = produtosDistintosVendidos.Count
            
            ' Preencher a coluna Q com a lista de produtos distintos vendidos
            If produtosDistintosVendidos.Count > 0 Then
                Dim listaProdutos As String
                listaProdutos = Join(produtosDistintosVendidos.keys, ", ")
                wsBase.Cells(i + 1, 17).Value = listaProdutos
            Else
                wsBase.Cells(i + 1, 17).Value = 0
            End If
        Else
            ' Se a missão não contiver "distinto" ou "distintos", preencher a coluna P com 0
            wsBase.Cells(i + 1, 16).Value = 0
        End If
    Next i
    
    ' Informar que a macro terminou
    MsgBox "Quantidade de Compra no mês calculada com sucesso!"
End Sub

