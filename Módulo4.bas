Attribute VB_Name = "M�dulo1"
Sub ValidarVendasDiaEDistintos()
    Dim wsBase As Worksheet
    Dim wsDia As Worksheet
    Dim ultimaLinhaBase As Long, ultimaLinhaDia As Long
    Dim i As Long, j As Long, k As Long
    Dim pdvBase As String, produtosBase As String
    Dim pdvDia As String, produtoDia As String
    Dim quantidadeVendida As Long
    Dim quantidadeDistintos As Long
    Dim missao As String
    Dim quantidadeDia As Long
    Dim codigosDobrar As Variant
    Dim produtosArray() As String
    Dim baseData As Variant, diaData As Variant
    Dim produtosDistintosVendidos As Collection
    Dim listaProdutos As String
    
    ' Desativar atualiza��es de tela e c�lculo autom�tico para melhorar a performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Definir os c�digos que devem ter a quantidade dobrada
    codigosDobrar = Array("988", "2538", "982")
    
    ' Definir as planilhas
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    Set wsDia = ActiveWorkbook.Worksheets("Dia")
    
    ' Encontrar a �ltima linha de ambas as planilhas
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaDia = wsDia.Cells(wsDia.Rows.Count, "A").End(xlUp).Row
    
    ' Carregar os dados da planilha em arrays para processamento em mem�ria
    baseData = wsBase.Range("A1:X" & ultimaLinhaBase).Value
    diaData = wsDia.Range("A1:I" & ultimaLinhaDia).Value
    
    ' Percorrer cada linha da planilha Base
    For i = 2 To ultimaLinhaBase
        pdvBase = baseData(i, 6) ' PDV da base (Coluna F)
        produtosBase = baseData(i, 4) ' Produtos da base (Coluna D)
        missao = baseData(i, 3) ' Miss�o (Coluna C)
        
        ' Separar os produtos da base
        produtosArray = Split(produtosBase, ",")
        
        quantidadeVendida = 0 ' Inicializar a quantidade vendida para o PDV atual
        quantidadeDistintos = 0 ' Inicializar a quantidade distinta para o PDV atual
        Set produtosDistintosVendidos = New Collection ' Cole��o para armazenar produtos distintos vendidos
        
        ' Inicializar a mensagem padr�o na coluna X
        baseData(i, 24) = "N�o Venda, Falta ou Digitado" ' Coluna X
        
        ' Percorrer a planilha Dia para comparar os PDVs e produtos
        For j = 2 To ultimaLinhaDia
            pdvDia = diaData(j, 1) ' PDV do dia (Coluna A)
            produtoDia = diaData(j, 8) ' Produto do dia (Coluna H)
            quantidadeDia = diaData(j, 9) ' Quantidade vendida do dia (Coluna I)
            
            ' Verificar se o PDV e o produto correspondem e se a coluna E � igual a 1
            If pdvBase = pdvDia And diaData(j, 5) = 1 Then
                For k = LBound(produtosArray) To UBound(produtosArray)
                    If Trim(produtosArray(k)) = produtoDia Then
                        ' Adicionar o produto � lista de produtos distintos vendidos
                        On Error Resume Next
                        produtosDistintosVendidos.Add produtoDia, CStr(produtoDia)
                        On Error GoTo 0
                        
                        ' Verificar se o produto est� na lista de c�digos para dobrar a quantidade
                        If IsInArray(produtoDia, codigosDobrar) Then
                            quantidadeVendida = quantidadeVendida + (quantidadeDia * 2)
                        Else
                            quantidadeVendida = quantidadeVendida + quantidadeDia
                        End If
                        
                        ' Atualizar a mensagem indicando que houve venda
                        baseData(i, 24) = "Venda Realizada" ' Coluna X
                    End If
                Next k
            End If
        Next j
        
        ' Preencher a coluna I com a quantidade vendida
        baseData(i, 9) = quantidadeVendida ' Coluna I
        
        ' Verificar se a miss�o cont�m "distinto" ou "distintos"
        If InStr(missao, "distinto") > 0 Or InStr(missao, "distintos") > 0 Then
            ' Preencher a coluna L com a quantidade de produtos distintos
            baseData(i, 12) = produtosDistintosVendidos.Count ' Coluna L
            
            ' Preencher a coluna M com a lista de produtos distintos vendidos
            If produtosDistintosVendidos.Count > 0 Then
                listaProdutos = ""
                For Each produtoDia In produtosDistintosVendidos
                    If listaProdutos = "" Then
                        listaProdutos = produtoDia
                    Else
                        listaProdutos = listaProdutos & ", " & produtoDia
                    End If
                Next produtoDia
                baseData(i, 13) = listaProdutos ' Coluna M
            Else
                baseData(i, 13) = 0 ' Coluna M
            End If
        Else
            baseData(i, 12) = 0 ' Coluna L
        End If
    Next i
    
    ' Copiar os dados modificados de volta para a planilha Base
    wsBase.Range("A1:X" & ultimaLinhaBase).Value = baseData
    
    ' Reativar atualiza��es de tela e c�lculos autom�ticos
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Informar que a macro terminou
    MsgBox "Vendas dias Calculadas"
End Sub

' Fun��o auxiliar para verificar se o produto est� no array de c�digos para dobrar a quantidade
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

