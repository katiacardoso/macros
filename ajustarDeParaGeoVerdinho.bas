Sub MapearSKUCategorias()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim categoriaDict As Object
    Dim i As Long, j As Long
    Dim categoria As Variant, sku As String
    Dim resultado As String
    Dim wsResultado As Worksheet
    Dim resposta As VbMsgBoxResult

    ' Configurar a planilha ativa
    Set ws = ThisWorkbook.Sheets("Depara_SKU")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Inicializar o dicionário para armazenar as categorias e SKUs
    Set categoriaDict = CreateObject("Scripting.Dictionary")

    ' Loop pelas linhas e colunas para mapear SKUs por categorias
    For i = 2 To lastRow ' Começa na linha 2 para ignorar o cabeçalho
        sku = ws.Cells(i, 1).Value ' SKU está na coluna A
        For j = 2 To lastCol ' Verifica as categorias nas colunas
            categoria = ws.Cells(i, j).Value
            If Not IsEmpty(categoria) And categoria <> "-" Then
                If Not categoriaDict.exists(categoria) Then
                    categoriaDict.Add categoria, sku
                Else
                    categoriaDict(categoria) = categoriaDict(categoria) & ", " & sku
                End If
            End If
        Next j
    Next i

    ' Verificar se a aba "Resultado" já existe
    On Error Resume Next
    Set wsResultado = ThisWorkbook.Sheets("Resultado")
    On Error GoTo 0

    If Not wsResultado Is Nothing Then
        resposta = MsgBox("A aba 'Resultado' já existe. Deseja sobrescrevê-la?", vbYesNo + vbQuestion, "Confirmação")
        If resposta = vbYes Then
            Application.DisplayAlerts = False
            wsResultado.Delete
            Application.DisplayAlerts = True
            Set wsResultado = ThisWorkbook.Sheets.Add
            wsResultado.Name = "Resultado"
        Else
            MsgBox "Operação cancelada pelo usuário.", vbExclamation
            Exit Sub
        End If
    Else
        Set wsResultado = ThisWorkbook.Sheets.Add
        wsResultado.Name = "Resultado"
    End If

    ' Escrever os resultados na aba "Resultado"
    i = 1
    For Each categoria In categoriaDict.Keys
        resultado = categoria & ": " & categoriaDict(categoria)
        wsResultado.Cells(i, 1).Value = resultado
        i = i + 1
    Next categoria

    ' Transformar texto em colunas com delimitador ":"
    wsResultado.Columns("A").TextToColumns Destination:=wsResultado.Columns("A"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:=":"

    ' Remover linhas que contenham apenas números na coluna A
    lastRow = wsResultado.Cells(wsResultado.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 1 Step -1
        If IsNumeric(wsResultado.Cells(i, "A").Value) Then
            wsResultado.Rows(i).Delete
        End If
    Next i

    ' Adiciona um comentário sobre possível lentidão após o preenchimento
    MsgBox "Os valores foram adicionados na aba 'Resultado'. Dependendo da quantidade de dados, a performance do Excel pode ser afetada.", vbExclamation

    MsgBox "Mapeamento concluído! Verifique a aba 'Resultado'.", vbInformation
End Sub


