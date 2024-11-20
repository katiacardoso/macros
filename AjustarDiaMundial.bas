Sub AdicionarColunasEConcatenar()
    Dim wsBase As Worksheet
    Dim wsBaseClientes As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Define as abas
    Set wsBase = ThisWorkbook.Sheets("Base")
    ' Abre a planilha da base de clientes (substitua pelo caminho completo)
    Set wsBaseClientes = Workbooks.Open("C:\Users\k.lorena\Documents\DM\BaseClientes\012011.xlsx").Sheets("baseClientes")

    ' Encontra a última linha da aba Base
    lastRow = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

    ' Verifica se o usuário deseja inserir as colunas
    If MsgBox("Deseja inserir as colunas? (Sim para inserir)", vbYesNo) = vbYes Then
        ' Insere as colunas
        wsBase.Columns("B:B").Insert Shift:=xlToRight
        wsBase.Columns("J:K").Insert Shift:=xlToRight
        
        ' Nomeia e formata as novas colunas
        With wsBase
            .Cells(1, "B").Value = "UNB_PDV"
            .Cells(1, "J").Value = "NOME FANTASIA"
            .Cells(1, "K").Value = "VISITA"
            
            ' Copia a formatação da coluna A para as novas colunas
            .Range("B1:K1").Interior.Color = .Range("A1").Interior.Color
            .Range("B1:K1").Font.Color = .Range("A1").Font.Color
            ' ... outras propriedades de formatação que desejar copiar
        End With
        
    End If

    ' Renomeia as colunas se necessário
    For col = 1 To wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
        If wsBase.Cells(1, col).Value = "dim_estrutura[cod_unb]" Then
            wsBase.Cells(1, col).Value = "UNB"
        ElseIf wsBase.Cells(1, col).Value = "dim_estrutura[comercial]" Then
            wsBase.Cells(1, col).Value = "COMERCIAL"
        ElseIf wsBase.Cells(1, col).Value = "dim_estrutura[supercom]" Then
            wsBase.Cells(1, col).Value = "SUPERCOM"
        ElseIf wsBase.Cells(1, col).Value = "dim_estrutura[operacao]" Then
            wsBase.Cells(1, col).Value = "OPERAÇÃO"
        ElseIf wsBase.Cells(1, col).Value = "dim_estrutura[tipooperacao]" Then
            wsBase.Cells(1, col).Value = "TIPO"
        ElseIf wsBase.Cells(1, col).Value = "dim_estrutura[cod_gv]" Then
            wsBase.Cells(1, col).Value = "GV"
        ElseIf wsBase.Cells(1, col).Value = "dim_estrutura[cod_setor]" Then
            wsBase.Cells(1, col).Value = "SETOR"
        ElseIf wsBase.Cells(1, col).Value = "fato_tasks_kpis[cod_pdv]" Then
            wsBase.Cells(1, col).Value = "PDV"
        ElseIf wsBase.Cells(1, col).Value = "fato_tasks_kpis[task_text]" Then
            wsBase.Cells(1, col).Value = "TAREFA"
        End If
               
    Next col

    ' Concatena os valores nas colunas A e I na coluna B
    For i = 2 To lastRow
        wsBase.Cells(i, "B").Value = wsBase.Cells(i, "A").Value & "_" & wsBase.Cells(i, "I").Value
    Next i

    ' Busca e retorna os valores da aba baseClientes
    For i = 2 To lastRow
        wsBase.Cells(i, "J").Value = Application.VLookup(wsBase.Cells(i, "B").Value, wsBaseClientes.Range("A:K"), 10, False)
        wsBase.Cells(i, "K").Value = Application.VLookup(wsBase.Cells(i, "B").Value, wsBaseClientes.Range("A:K"), 11, False)
    Next i

End Sub
