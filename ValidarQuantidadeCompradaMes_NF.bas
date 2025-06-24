' Macro feita inicialmente para validar tarefas de Marketplace a partir da exportação de dados do Bees Force

Sub ValidarQuantidadeCompradaMes_NF()
    Dim wsBase As Worksheet
    Dim wsNotasFiscais As Worksheet
    Dim ultimaLinhaBase As Long, ultimaLinhaNF As Long
    Dim i As Long, j As Long
    Dim pdvBase As String, produtosBase As String
    Dim pdvNF As String, produtoNF As String
    Dim quantidadeVendida As Double
    Dim produtosDistintosVendidos As Object
    Dim missao As String
    Dim dictPDV As Object
    Dim produtosArray() As String
    Dim statusNota As String, operacaoNota As Variant

    ' ==== Colunas Base (A=1, ..., M=13, O=15, P=16, Q=17, S=19, X=24, Y=25) ====
    Const COL_MISSAO As Long = 22        ' Coluna L
    Const COL_PRODUTOS As Long = 28      ' Coluna O
    Const COL_PDV_BASE As Long = 9       ' Coluna A
    'Const COL_QTD_FINAL As Long = 30     ' Coluna S
    Const COL_QTD_DIST As Long = 30      ' Coluna X
    Const COL_LISTA_PROD As Long = 31    ' Coluna Y

    ' ==== Colunas Notas Fiscais ("03.02.37") ====
    Const COL_PDV_NF As Long = 13         ' Coluna A
    Const COL_OPERACAO As Long = 3       ' Coluna D
    Const COL_STATUS As Long = 10        ' Coluna K
    Const COL_PRODUTO_NF As Long = 16    ' Coluna Q
    Const COL_QTD_NF As Long = 20        ' Coluna U

    ' Inicializar dicionï¿½rios
    Set dictPDV = CreateObject("Scripting.Dictionary")
    Set produtosDistintosVendidos = CreateObject("Scripting.Dictionary")

    ' Definir as planilhas
    Set wsBase = ThisWorkbook.Worksheets("Base")
    Set wsNotasFiscais = ThisWorkbook.Worksheets("03.02.37")

    ' Obter ï¿½ltimas linhas
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaNF = wsNotasFiscais.Cells(wsNotasFiscais.Rows.Count, "A").End(xlUp).Row
    
    ' Limpar colunas COL_QTD_DIST e COL_LISTA_PROD a partir da linha 2
    wsBase.Range(wsBase.Cells(2, COL_QTD_DIST), wsBase.Cells(ultimaLinhaBase, COL_QTD_DIST)).ClearContents
    wsBase.Range(wsBase.Cells(2, COL_LISTA_PROD), wsBase.Cells(ultimaLinhaBase, COL_LISTA_PROD)).ClearContents

    
    ' Copiar dados para arrays
    Dim baseData As Variant
    Dim nfData As Variant
    baseData = wsBase.Range("A2:AF" & ultimaLinhaBase).Value
    nfData = wsNotasFiscais.Range("A2:Z" & ultimaLinhaNF).Value

    ' Construir dicionï¿½rio de PDV x Produto a partir das notas fiscais vï¿½lidas
    For j = 1 To UBound(nfData, 1)
        statusNota = Trim(nfData(j, COL_STATUS))
        operacaoNota = nfData(j, COL_OPERACAO)

        If statusNota = "A" And (operacaoNota = 1 Or operacaoNota = 2) Then
            pdvNF = Trim(nfData(j, COL_PDV_NF))
            produtoNF = Trim(nfData(j, COL_PRODUTO_NF))

            If IsNumeric(nfData(j, COL_QTD_NF)) Then
                quantidadeVendida = CDbl(nfData(j, COL_QTD_NF))
            Else
                quantidadeVendida = 0
            End If

            If Not dictPDV.exists(pdvNF) Then
                Set dictPDV(pdvNF) = CreateObject("Scripting.Dictionary")
            End If

            If Not dictPDV(pdvNF).exists(produtoNF) Then
                dictPDV(pdvNF)(produtoNF) = quantidadeVendida
            Else
                dictPDV(pdvNF)(produtoNF) = dictPDV(pdvNF)(produtoNF) + quantidadeVendida
            End If
        End If
    Next j

    ' Processar dados da planilha Base
    For i = 1 To UBound(baseData, 1)
        pdvBase = Trim(baseData(i, COL_PDV_BASE))
        produtosBase = baseData(i, COL_PRODUTOS)
        missao = baseData(i, COL_MISSAO)

        produtosArray = Split(produtosBase, ",")
        quantidadeVendida = 0
        produtosDistintosVendidos.RemoveAll

        If dictPDV.exists(pdvBase) Then
            For Each produto In produtosArray
                produto = Trim(produto)
                If dictPDV(pdvBase).exists(produto) Then
                    quantidadeVendida = quantidadeVendida + dictPDV(pdvBase)(produto)
                    If Not produtosDistintosVendidos.exists(produto) Then
                        produtosDistintosVendidos.Add produto, Nothing
                    End If
                End If
            Next produto
        End If

        ' Atualizar valores na planilha Base
        'wsBase.Cells(i + 1, COL_QTD_FINAL).Value = quantidadeVendida

        If InStr(LCase(missao), "distinto") > 0 Then
            wsBase.Cells(i + 1, COL_QTD_DIST).Value = produtosDistintosVendidos.Count
            If produtosDistintosVendidos.Count > 0 Then
                wsBase.Cells(i + 1, COL_LISTA_PROD).Value = Join(produtosDistintosVendidos.keys, ", ")
            Else
                wsBase.Cells(i + 1, COL_LISTA_PROD).Value = 0
            End If
        Else
            wsBase.Cells(i + 1, COL_QTD_DIST).Value = 0
        End If
    Next i

    MsgBox "Quantidade de compras por notas fiscais (status A, operaï¿½ï¿½es 1 e 2) processada com sucesso!"
End Sub

