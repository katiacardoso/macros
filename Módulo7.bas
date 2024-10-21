Attribute VB_Name = "Módulo12"
Sub ValidarVolumeETarefas()
    Dim wsBase As Worksheet
    Dim ultimaLinhaBase As Long
    Dim i As Long
    Dim missao As String
    Dim volumeDia As Double, skuDia As Double
    Dim volumeAgendado As Double, skuAgendado As Double
    Dim volumeMes As Double, checkSku As Double
    Dim quantidadeTarefa As Double
    Dim somaVolume As Double
    Dim valorU As Long
    Dim valorV As Long
    Dim resultadoW As Long
    
    ' Definir a planilha
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    
    ' Encontrar a última linha da planilha Base
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    
    ' Percorrer cada linha da planilha Base
    For i = 2 To ultimaLinhaBase
        ' Obter valores das células
        missao = wsBase.Cells(i, "C").Value ' Missão
        quantidadeTarefa = wsBase.Cells(i, "H").Value ' Quantidade da Tarefa
        volumeDia = wsBase.Cells(i, "I").Value ' Volume do Dia
        volumeAgendado = wsBase.Cells(i, "J").Value ' Volume Agendado
        volumeMes = wsBase.Cells(i, "K").Value ' Volume do Mês
        checkSku = wsBase.Cells(i, "S").Value ' check sku
        skuDia = wsBase.Cells(i, "L").Value ' Sku Dia
        skuAgendado = wsBase.Cells(i, "N").Value ' Sku Agendado
        
        ' 1. Calcular a soma total de volumeDia, volumeAgendado e volumeMes e colocar na coluna R
        somaVolume = volumeDia + volumeAgendado + volumeMes
        wsBase.Cells(i, "R").Value = somaVolume
        
        ' 2. Verificar se a missão contém "distinto" ou "distintos"
        If InStr(1, missao, "distinto", vbTextCompare) > 0 Or InStr(1, missao, "distintos", vbTextCompare) > 0 Then
            wsBase.Cells(i, "U").Value = 0
            ' 3. Validar se skuDia e skuAgendado são iguais a 0
            If skuDia = 0 And skuAgendado = 0 Then
                wsBase.Cells(i, "V").Value = 0
            ElseIf skuDia > 0 Or skuAgendado > 0 And checkSku >= quantidadeTarefa Then
                wsBase.Cells(i, "V").Value = 1
            Else
                wsBase.Cells(i, "V").Value = 0
            End If
            
        Else
            ' 3. Validar se volumeDia e volumeAgendado são iguais a 0
            If volumeDia = 0 And volumeAgendado = 0 Then
                wsBase.Cells(i, "U").Value = 0
            ElseIf volumeDia > 0 Or volumeAgendado > 0 And somaVolume >= quantidadeTarefa Then
                wsBase.Cells(i, "U").Value = 1
            Else
                wsBase.Cells(i, "U").Value = 0
            End If
        End If
        
        ' 4. Comparar a coluna U com a coluna V e retornar na coluna W
        valorU = wsBase.Cells(i, "U").Value
        valorV = wsBase.Cells(i, "V").Value
        
        If valorU = 1 Or valorV = 1 Then
            resultadoW = 1
        Else
            resultadoW = 0
        End If
        
        ' Colocar o resultado final na coluna W
        wsBase.Cells(i, "W").Value = resultadoW
    Next i
    
    ' Informar que a macro terminou
    MsgBox "Validação concluída!"
End Sub

