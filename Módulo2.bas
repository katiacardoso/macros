Attribute VB_Name = "M�dulo6"
Sub ValidarQuantidadeMissao()
    Dim wsBase As Worksheet
    Dim ultimaLinhaBase As Long
    Dim missao As String, pdv As String
    Dim produtosArray() As String
    Dim i As Long
    Dim resultado As Long
    
    ' Definir a planilha base
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    
    ' Encontrar a �ltima linha
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    
    ' Percorrer cada linha da planilha Base
    For i = 2 To ultimaLinhaBase
        missao = wsBase.Cells(i, "C").Value ' Coluna "Missao"
        pdv = wsBase.Cells(i, "F").Value ' PDV na coluna B
        
        If InStr(missao, "qualquer produto") And InStr(missao, "caixas") > 0 Then
            resultado = ExtrairNumeroDeCaixas(missao)
        ' Verificar a estrutura da frase
        ElseIf InStr(missao, "qualquer produto") > 0 Then
            ' Verificar se o cliente comprou qualquer produto
            resultado = ValidarCompraDeProdutos(wsBase.Cells(i, "D").Value) ' Supondo produtos concatenados em D
        ElseIf InStr(missao, "caixa") > 0 Then
            ' Extrair o n�mero de caixas a partir da miss�o
            resultado = ExtrairNumeroDeCaixas(missao)
        ElseIf InStr(missao, "produtos distintos") > 0 Then
            ' Extrair o n�mero de produtos distintos
            resultado = ValidarProdutosDistintos(wsBase.Cells(i, "D").Value, missao)
        Else
            resultado = 0
        End If
        
        ' Preencher a nova coluna com o resultado
        wsBase.Cells(i, "H").Value = resultado
    Next i
    MsgBox "Quantidade de Tarefas Calculado"
End Sub

' Fun��o para validar se o cliente comprou qualquer produto
Function ValidarCompraDeProdutos(produtosConcatenados As String) As Long
    If produtosConcatenados <> "" Then
        ValidarCompraDeProdutos = 1
    Else
        ValidarCompraDeProdutos = 0
    End If
End Function


' Fun��o para extrair o n�mero de caixas da miss�o
Function ExtrairNumeroDeCaixas(missao As String) As Long
    Dim palavras() As String
    Dim i As Long
    
    ' Separar a miss�o em palavras
    palavras = Split(missao, " ")
    
    ' Procurar o n�mero de caixas
    For i = LBound(palavras) To UBound(palavras)
        If IsNumeric(palavras(i)) Then
            ExtrairNumeroDeCaixas = CLng(palavras(i))
            Exit Function
        End If
    Next i
    
    ExtrairNumeroDeCaixas = 0
End Function
' Fun��o para validar a compra de produtos distintos
Function ValidarProdutosDistintos(produtosConcatenados As String, missao As String) As Long
    Dim produtos() As String
    Dim qtdDistintos As Long
    Dim qtdNaMissao As Long
    
    ' Extrair o n�mero de produtos distintos solicitados na miss�o
    qtdNaMissao = ExtrairNumeroDeCaixas(missao)
    
    ' Separar os produtos concatenados
    produtos = Split(produtosConcatenados, ",")
    
    ' Verificar a quantidade de produtos distintos
    qtdDistintos = UBound(produtos) - LBound(produtos) + 1
    
    ' Validar
    If qtdDistintos >= qtdNaMissao Then
        ValidarProdutosDistintos = qtdNaMissao
    Else
        ValidarProdutosDistintos = 0
    End If
End Function
