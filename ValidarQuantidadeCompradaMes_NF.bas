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
    Dim dataCriacaoTarefa As Date, dataVendaNF As Date

    ' ==== Colunas Base ==== 
    Const COL_DATA_TAREFA As Long = 2     ' Coluna B
    Const COL_PDV_BASE As Long = 9        ' Coluna I
    Const COL_MISSAO As Long = 22         ' Coluna V
    Const COL_PRODUTOS As Long = 28       ' Coluna AB
    Const COL_QTD_DIST As Long = 30       ' Coluna AD
    Const COL_LISTA_PROD As Long = 31     ' Coluna AE

    ' ==== Colunas Notas Fiscais ==== 
    Const COL_PDV_NF As Long = 13         ' Coluna M
    Const COL_OPERACAO As Long = 3        ' Coluna C
    Const COL_STATUS As Long = 10         ' Coluna J
    Const COL_DATA_VENDA As Long = 6      ' Coluna F
    Const COL_PRODUTO_NF As Long = 16     ' Coluna P
    Const COL_QTD_NF As Long = 20         ' Coluna T

    Set dictPDV = CreateObject("Scripting.Dictionary")
    Set produtosDistintosVendidos = CreateObject("Scripting.Dictionary")

    Set wsBase = ThisWorkbook.Worksheets("Base")
    Set wsNotasFiscais = ThisWorkbook.Worksheets("03.02.37")

    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaNF = wsNotasFiscais.Cells(wsNotasFiscais.Rows.Count, "A").End(xlUp).Row

    wsBase.Range(wsBase.Cells(2, COL_QTD_DIST), wsBase.Cells(ultimaLinhaBase, COL_QTD_DIST)).ClearContents
    wsBase.Range(wsBase.Cells(2, COL_LISTA_PROD), wsBase.Cells(ultimaLinhaBase, COL_LISTA_PROD)).ClearContents

    Dim baseData As Variant
    Dim nfData As Variant
    baseData = wsBase.Range("A2:AF" & ultimaLinhaBase).Value
    nfData = wsNotasFiscais.Range("A2:Z" & ultimaLinhaNF).Value

    ' Percorrer dados da planilha Base
    For i = 1 To UBound(baseData, 1)
        pdvBase = Trim(baseData(i, COL_PDV_BASE))
        produtosBase = baseData(i, COL_PRODUTOS)
        missao = baseData(i, COL_MISSAO)
        produtosArray = Split(produtosBase, ",")
        quantidadeVendida = 0
        produtosDistintosVendidos.RemoveAll

        ' Verificar se a data de criação da tarefa é válida
        If IsDate(baseData(i, COL_DATA_TAREFA)) Then
            dataCriacaoTarefa = baseData(i, COL_DATA_TAREFA)
        Else
            GoTo ProximoItem ' pula se não tiver data válida
        End If

        ' Percorrer notas fiscais
        For j = 1 To UBound(nfData, 1)
            statusNota = Trim(nfData(j, COL_STATUS))
            operacaoNota = nfData(j, COL_OPERACAO)

            If statusNota = "A" And (operacaoNota = 1 Or operacaoNota = 2) Then
                pdvNF = Trim(nfData(j, COL_PDV_NF))
                produtoNF = Trim(nfData(j, COL_PRODUTO_NF))

                ' Validar PDV igual
                If pdvBase = pdvNF Then
                    ' Validar data da venda
                    If IsDate(nfData(j, COL_DATA_VENDA)) Then
                        dataVendaNF = nfData(j, COL_DATA_VENDA)

                        If dataVendaNF >= dataCriacaoTarefa Then
                            ' Validar se o produto da nota está na lista da tarefa
                            For Each produto In produtosArray
                                produto = Trim(produto)
                                If produto = produtoNF Then
                                    If IsNumeric(nfData(j, COL_QTD_NF)) Then
                                        quantidadeVendida = quantidadeVendida + CDbl(nfData(j, COL_QTD_NF))
                                    End If
                                    If Not produtosDistintosVendidos.exists(produto) Then
                                        produtosDistintosVendidos.Add produto, Nothing
                                    End If
                                End If
                            Next produto
                        End If
                    End If
                End If
            End If
        Next j

        ' Atualizar planilha Base
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

ProximoItem:
    Next i

    MsgBox "Validação de compras com data a partir da criação da tarefa concluída com sucesso!"
End Sub


