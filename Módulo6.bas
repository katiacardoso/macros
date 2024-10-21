Attribute VB_Name = "M�dulo11"
Sub checkSku()
    Dim wsBase As Worksheet
    Dim ultimaLinhaBase As Long
    Dim i As Long
    Dim produtosDistintosDia As String
    Dim produtosDistintosAgendado As String
    Dim produtosDistintosMes As String
    Dim produtosTotalDistintos As Collection
    Dim listaProdutosTotal As String
    Dim produto As Variant
    
    ' Definir a planilha
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    
    ' Encontrar a �ltima linha da planilha Base
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    
    ' Percorrer cada linha da planilha Base
    For i = 2 To ultimaLinhaBase
        ' Obter os produtos distintos vendidos no Dia, Agendado e M�s
        produtosDistintosDia = wsBase.Cells(i, "M").Value
        produtosDistintosAgendado = wsBase.Cells(i, "O").Value
        produtosDistintosMes = wsBase.Cells(i, "Q").Value
        
        
        ' Se as 3 vari�veis forem iguais a 0, retornar 0 na coluna S e pular a listagem
        If produtosDistintosDia = "0" And produtosDistintosAgendado = "0" And produtosDistintosMes = "0" Then
            wsBase.Cells(i, "S").Value = 0
            wsBase.Cells(i, "T").Value = ""
        Else
        
            ' Inicializar a cole��o para armazenar os produtos distintos
            Set produtosTotalDistintos = New Collection
            
            ' Fun��o para adicionar produtos distintos � cole��o, desconsiderando "0"
            Call AdicionarProdutosDistintos(produtosTotalDistintos, produtosDistintosDia)
            Call AdicionarProdutosDistintos(produtosTotalDistintos, produtosDistintosAgendado)
            Call AdicionarProdutosDistintos(produtosTotalDistintos, produtosDistintosMes)
            
            ' Construir a lista de produtos distintos
            listaProdutosTotal = ""
            
            For Each produto In produtosTotalDistintos
                If listaProdutosTotal = "" Then
                    listaProdutosTotal = produto
                Else
                    listaProdutosTotal = listaProdutosTotal & ", " & produto
                End If
            Next produto
                        
            listaProdutosTotal = Trim(listaProdutosTotal) ' Remover espa�os em branco
                    
            ' Retornar a quantidade de produtos distintos na coluna S
            wsBase.Cells(i, "S").Value = produtosTotalDistintos.Count
            
            ' Retornar a lista de produtos distintos na coluna T
            wsBase.Cells(i, "T").Value = listaProdutosTotal
        End If
    Next i
    
    ' Informar que a macro terminou
    MsgBox "Check SKU conclu�da!"
End Sub

' Fun��o para adicionar produtos distintos � cole��o
Sub AdicionarProdutosDistintos(ByRef colecaoProdutos As Collection, produtos As String)
    Dim produtosArray() As String
    Dim produto As String
    Dim k As Long
    
    ' Verificar se a string de produtos n�o est� vazia e n�o cont�m "0"
    If produtos <> "" And produtos <> "0" Then
        ' Separar os produtos por v�rgula
        produtosArray = Split(produtos, ",")
        
        ' Percorrer os produtos e adicion�-los � cole��o, garantindo que sejam distintos
        For k = LBound(produtosArray) To UBound(produtosArray)
            produto = Trim(produtosArray(k)) ' Remover espa�os em branco
            On Error Resume Next
            colecaoProdutos.Add produto, CStr(produto) ' Adicionar apenas se for distinto
            On Error GoTo 0
        Next k
    End If
End Sub

