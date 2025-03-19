Sub ExportarPDFPorSetor()
    Dim ws As Worksheet
    Dim setor As Range
    Dim setores As Object
    Dim caminho As String
    Dim nomeArquivo As String
    Dim setorValor As Variant
    Dim setorMin As Integer, setorMax As Integer
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Definir intervalo desejado (altere conforme necessário)
    setorMin = 111
    setorMax = 118
    
    ' Criar um objeto para armazenar setores únicos dentro do intervalo
    Set setores = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha da coluna F
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Percorrer a coluna F e armazenar setores ï¿½nicos dentro do intervalo definido
    For Each setor In ws.Range("F2:F" & ultimaLinha)
        If IsNumeric(setor.Value) Then ' Verifica se ï¿½ um nï¿½mero
            If setor.Value >= setorMin And setor.Value <= setorMax Then
                If Not setores.exists(CStr(setor.Value)) Then
                    setores.Add CStr(setor.Value), Nothing
                End If
            End If
        End If
    Next setor
    
    ' Caminho da pasta de Documentos do usuï¿½rio
    caminho = Environ("USERPROFILE") & "\Documents\"
    
    ' Filtrar e gerar PDFs apenas para os setores do intervalo
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each setorValor In setores.keys
        ' Aplicar o filtro na coluna F
        ws.Range("F1").AutoFilter Field:=6, Criteria1:=setorValor
        
        ' Definir nome do arquivo
        nomeArquivo = "Setor_" & setorValor & ".pdf"
        
        ' Exportar como PDF
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=caminho & nomeArquivo, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    Next setorValor
    
    ' Remover Filtro
    ws.AutoFilterMode = False
    
    ' Restaurar atualizações
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Exportação concluídda! Os PDFs estï¿½o na pasta Documentos.", vbInformation
End Sub
