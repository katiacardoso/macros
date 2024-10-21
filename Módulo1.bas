Attribute VB_Name = "Módulo1"
Sub ConcatenarUNB_PDV()
    Dim wsBase As Worksheet
    Dim wsDia As Worksheet
    Dim wsAgendado As Worksheet
    Dim wsMes As Worksheet
    Dim ultimaLinhaBase As Long, ultimaLinhaDia As Long
    Dim ultimaLinhaAgendado As Long, ultimaLinhaMes As Long
    Dim i As Long, j As Long
    Dim pdvBase As String, produtosBase As String
    Dim pdvDia As String, pdvAgendado As String, pdvMes As String
    Dim unbDia As String, unbAgendado As String, unbMes As String
    Dim quantidadeVendida As Long
    
    ' Definir as planilhas
    Set wsBase = ActiveWorkbook.Worksheets("Base")
    Set wsDia = ActiveWorkbook.Worksheets("Dia")
    Set wsAgendado = ActiveWorkbook.Worksheets("Agendado")
    Set wsMes = ActiveWorkbook.Worksheets("03.05.09") ' Nome da aba
    
    ' Encontrar a última linha de todas as planilhas
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaDia = wsDia.Cells(wsDia.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaAgendado = wsAgendado.Cells(wsAgendado.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaMes = wsMes.Cells(wsMes.Rows.Count, "A").End(xlUp).Row
    
        ' --- Arquivo Dia ---
        For j = 2 To ultimaLinhaDia
            ' Determinar o código UNB com base na coluna B
            Select Case wsDia.Cells(j, "B").Value
                Case 1
                    unbDia = "323527"
                Case 2
                    unbDia = "878928"
                Case 3
                    unbDia = "970751"
                Case 4
                    unbDia = "1017039"
            End Select
            ' Concatenar UNB e código PDV da coluna F
            pdvDia = unbDia & "_" & wsDia.Cells(j, "F").Value
            wsDia.Cells(j, "A").Value = pdvDia
        Next j
        
        ' --- Arquivo Agendado ---
        For j = 2 To ultimaLinhaAgendado
            ' Concatenar UNB da coluna C e código PDV da coluna D
            unbAgendado = wsAgendado.Cells(j, "C").Value
            pdvAgendado = unbAgendado & "_" & wsAgendado.Cells(j, "D").Value
            wsAgendado.Cells(j, "A").Value = pdvAgendado
        Next j
        
        ' --- Arquivo Mês (03.05.09) ---
        For j = 2 To ultimaLinhaMes
            ' Determinar o código UNB com base na coluna S
            Select Case wsMes.Cells(j, "S").Value
                Case 1
                    unbMes = "323527"
                Case 2
                    unbMes = "878928"
                Case 3
                    unbMes = "970751"
                Case 4
                    unbMes = "1017039"
            End Select
            ' Concatenar UNB e código PDV da coluna B
            pdvMes = unbMes & "_" & wsMes.Cells(j, "B").Value
            wsMes.Cells(j, "A").Value = pdvMes
        Next j
    
    ' Informar que a macro terminou
    MsgBox "Concatenação concluída!"
End Sub


