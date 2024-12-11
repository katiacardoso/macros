Sub ClientesNaoCompradores()

    Dim wsCaixa As Worksheet
    Dim wsControle As Worksheet
    Dim wsNaoCompradores As Worksheet
    Dim ultimaLinhaCaixa As Long
    Dim ultimaLinhaControle As Long
    Dim ultimaLinhaNaoCompradores As Long
    Dim dictCompradores As Object
    Dim i As Long, cliente As String
    Dim encontrou As Boolean

    ' Definir as planilhas
    Set wsCaixa = ThisWorkbook.Worksheets("03.05.09 Cliente - Caixa")
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    
    ' Criar/limpar aba para clientes não compradores
    On Error Resume Next
    Set wsNaoCompradores = ThisWorkbook.Worksheets("Clientes Não Compradores")
    If wsNaoCompradores Is Nothing Then
        Set wsNaoCompradores = ThisWorkbook.Worksheets.Add
        wsNaoCompradores.Name = "Clientes Não Compradores"
    Else
        wsNaoCompradores.Cells.Clear
    End If
    On Error GoTo 0

    ' Criar um dicionário para armazenar os clientes compradores
    Set dictCompradores = CreateObject("Scripting.Dictionary")

    ' Obter última linha da aba "03.05.09 Cliente - Caixa"
    ultimaLinhaCaixa = wsCaixa.Cells(wsCaixa.Rows.Count, "C").End(xlUp).Row

    ' Passar pelos registros e adicionar clientes compradores ao dicionário
    For i = 2 To ultimaLinhaCaixa
        If wsCaixa.Cells(i, "D").Value = "Mondelez" Then
            cliente = wsCaixa.Cells(i, "C").Value
            If Not dictCompradores.exists(cliente) Then
                dictCompradores.Add cliente, True
            End If
        End If
    Next i

    ' Obter última linha da aba "Controle"
    ultimaLinhaControle = wsControle.Cells(wsControle.Rows.Count, "A").End(xlUp).Row

    ' Verificar clientes da aba "Controle" e identificar não compradores
    ultimaLinhaNaoCompradores = 1
    wsNaoCompradores.Cells(1, 1).Value = "Clientes Não Compradores"
    For i = 2 To ultimaLinhaControle
        cliente = wsControle.Cells(i, "A").Value
        If Not dictCompradores.exists(cliente) Then
            ultimaLinhaNaoCompradores = ultimaLinhaNaoCompradores + 1
            wsNaoCompradores.Cells(ultimaLinhaNaoCompradores, 1).Value = cliente
        End If
    Next i

    ' Mensagem de conclusão
    MsgBox "Lista de clientes não compradores gerada com sucesso!", vbInformation

End Sub
