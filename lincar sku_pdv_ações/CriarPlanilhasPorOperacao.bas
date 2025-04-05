Sub CriarPlanilhasPorOperacao()

    Dim wsBase As Worksheet
    Dim wsNova As Worksheet
    Dim ultimaLinha As Long, i As Long, j As Long
    Dim dictOperacoes As Object
    Dim operacao As Variant
    Dim produtos As Variant, pdv As String, tarefa As String, quantidade As String
    Dim regex As Object, matches As Object
    Dim novaLinha As Long
    Dim linhaAtual As Long
    Dim nomeAba As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsBase = ThisWorkbook.Sheets("BASEE")
    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

    ' Criar dicionário para armazenar operações únicas
    Set dictOperacoes = CreateObject("Scripting.Dictionary")

    ' Identificar operações únicas na coluna N
    For i = 2 To ultimaLinha
        operacao = Trim(wsBase.Cells(i, "N").Value)
        If Len(operacao) > 0 Then
            If Not dictOperacoes.exists(operacao) Then
                dictOperacoes.Add operacao, True
            End If
        End If
    Next i

    ' Expressão regular para extrair a quantidade
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "(\d+)\s*caixas"
    End With

    ' Loop em cada operação
    For Each operacao In dictOperacoes.Keys

        ' Criar nome da nova aba sem caracteres inválidos
        nomeAba = "acob_" & Replace(Replace(Replace(Replace(operacao, "\", "_"), "/", "_"), "*", "_"), "?", "_")
        
        ' Apagar aba existente com esse nome, se houver
        On Error Resume Next
        ThisWorkbook.Sheets(nomeAba).Delete
        On Error GoTo 0
        
        Set wsNova = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNova.Name = nomeAba
        
        ' Cabeçalho
        wsNova.Range("A1").Value = "produto"
        wsNova.Range("B1").Value = "pdv"
        wsNova.Range("C1").Value = "quantidade"
        novaLinha = 2

        ' Varredura nas linhas da BASE para essa operação
        For i = 2 To ultimaLinha
            If Trim(wsBase.Cells(i, "N").Value) = operacao Then
                If wsBase.Cells(i, "E").Value <> "" And wsBase.Cells(i, "M").Value <> "" Then
                    produtos = Split(wsBase.Cells(i, "E").Value, ",")
                    pdv = wsBase.Cells(i, "M").Value
                    tarefa = wsBase.Cells(i, "T").Value
                    quantidade = "1"
                    
                    If regex.test(tarefa) Then
                        Set matches = regex.Execute(tarefa)
                        quantidade = matches(0).SubMatches(0)
                    End If
                    
                    For j = LBound(produtos) To UBound(produtos)
                        wsNova.Cells(novaLinha, 1).Value = Trim(produtos(j))
                        wsNova.Cells(novaLinha, 2).Value = pdv
                        wsNova.Cells(novaLinha, 3).Value = quantidade
                        novaLinha = novaLinha + 1
                    Next j
                End If
            End If
        Next i
    Next operacao

    Application.ScreenUpdating = True
    MsgBox "Planilhas criadas com sucesso para cada operação!", vbInformation

End Sub

