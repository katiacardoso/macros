Attribute VB_Name = "M�dulo4"
Sub ValidarQuantidadeCompradaMes()
    Dim wsBase As Worksheet
    Dim wsCompraM�s As Worksheet
    Dim ultimaLinhaBase As Long, ultimaLinhaCompraM�s As Long
    Dim i As Long, j As Long
    Dim pdvBase As String, produtosBase As String
    Dim pdvCompraM�s As String, produtoCompraM�s As String
    Dim quantidadeVendida As Double
    Dim produtosDistintosVendidos As Object
    Dim missao As String
    Dim dictPDV As Object
    Dim produtosArray() As String
    
    ' Inicializar dicion�rios
    Set dictPDV = CreateObject("Scripting.Dictionary")
    Set produtosDistintosVendidos = CreateObject("Scripting.Dictionary")
    
    ' Definir as planilhas
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    Set wsCompraM�s = ActiveWorkbook.Worksheets("03.05.09")
    
    ' Encontrar a �ltima linha de ambas as planilhas
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaCompraM�s = wsCompraM�s.Cells(wsCompraM�s.Rows.Count, "A").End(xlUp).Row
    
    ' Copiar dados das planilhas para arrays
    Dim baseData As Variant
    Dim compraMesData As Variant
    
    baseData = wsBase.Range("A2:P" & ultimaLinhaBase).Value ' Copiar dados da Base
    compraMesData = wsCompraM�s.Range("A2:G" & ultimaLinhaCompraM�s).Value ' Copiar dados de CompraM�s
    
    ' Percorrer a planilha CompraM�s e armazenar os dados no dicion�rio
    For j = 1 To UBound(compraMesData, 1)
        pdvCompraM�s = compraMesData(j, 1) ' PDV do CompraM�s
        produtoCompraM�s = compraMesData(j, 4) ' Produto do CompraM�s
        
        ' Verificar se a c�lula cont�m um n�mero, caso contr�rio, atribuir 0
        If IsNumeric(compraMesData(j, 7)) Then
            quantidadeVendida = compraMesData(j, 7) ' Quantidade vendida no CompraM�s
        Else
            quantidadeVendida = 0 ' Se n�o for num�rico, atribuir 0
        End If
        
        ' Armazenar no dicion�rio usando PDV como chave
        If Not dictPDV.exists(pdvCompraM�s) Then
            Set dictPDV(pdvCompraM�s) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Armazenar a quantidade vendida e os produtos distintos
        If Not dictPDV(pdvCompraM�s).exists(produtoCompraM�s) Then
            dictPDV(pdvCompraM�s)(produtoCompraM�s) = quantidadeVendida
        Else
            dictPDV(pdvCompraM�s)(produtoCompraM�s) = dictPDV(pdvCompraM�s)(produtoCompraM�s) + quantidadeVendida
        End If
    Next j
    
    ' Processar os dados da Base
    For i = 1 To UBound(baseData, 1)
        pdvBase = baseData(i, 6) ' PDV da base
        produtosBase = baseData(i, 4) ' Produtos da base (separados por v�rgula)
        missao = baseData(i, 3) ' Miss�o na coluna C da Base
        
        ' Separar os produtos da base
        produtosArray = Split(produtosBase, ",")
        
        ' Inicializar a quantidade de venda e produtos distintos
        quantidadeVendida = 0
        produtosDistintosVendidos.RemoveAll
        
        ' Verificar se o PDV existe no dicion�rio
        If dictPDV.exists(pdvBase) Then
            For Each produto In produtosArray
                produto = Trim(produto)
                If dictPDV(pdvBase).exists(produto) Then
                    quantidadeVendida = quantidadeVendida + dictPDV(pdvBase)(produto)
                    
                    ' Adicionar produto � lista de distintos
                    If Not produtosDistintosVendidos.exists(produto) Then
                        produtosDistintosVendidos.Add produto, Nothing
                    End If
                End If
            Next produto
        End If
        
        ' Preencher a quantidade vendida na coluna K
        wsBase.Cells(i + 1, 11).Value = quantidadeVendida
        
        ' Verificar se a miss�o cont�m "distinto" ou "distintos"
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
            ' Se a miss�o n�o contiver "distinto" ou "distintos", preencher a coluna P com 0
            wsBase.Cells(i + 1, 16).Value = 0
        End If
    Next i
    
    ' Informar que a macro terminou
    MsgBox "Quantidade de Compra no m�s calculada com sucesso!"
End Sub

