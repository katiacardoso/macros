Sub CorrigirCaracteres()
    Dim celula As Range
    ' Define o intervalo como toda a coluna G, ajustando para a última linha com dados
    Dim ultimaLinha As Long
    ultimaLinha = Application.WorksheetFunction.Min(Cells(Rows.Count, "G").End(xlUp).Row, Cells(Rows.Count, "H").End(xlUp).Row)
    ' Formata a coluna H como texto antes de processar os dados
    Range("H2:H" & ultimaLinha).NumberFormat = "@"
    For Each celula In Range("G2:G" & ultimaLinha) ' Inicia em G2 e vai até a última linha com dados na coluna G
        celula.Value = Replace(celula.Value, "ðŸš¨ ", "")
        celula.Value = Replace(celula.Value, "ðŸŸ¡ ", "")
        celula.Value = Replace(celula.Value, "Ã•", "Õ")
        celula.Value = Replace(celula.Value, "Ãª", "ê")
        celula.Value = Replace(celula.Value, "Ã¡", "á")
        celula.Value = Replace(celula.Value, "ã", "Á")
        celula.Value = Replace(celula.Value, "Á-", "í")
        celula.Value = Replace(celula.Value, "Ãº", "ú")
        celula.Value = Replace(celula.Value, "seÃ§Ã£o", "sessão")
        celula.Value = Replace(celula.Value, "PromoÃ§Ã£o", "Promoção")
        celula.Value = Replace(celula.Value, "disponÃ­vel", "disponível")
        celula.Value = Replace(celula.Value, "DisponÃ­vel", "Disponível")
        celula.Value = Replace(celula.Value, "nÃ£o", "não")
        celula.Value = Replace(celula.Value, "ZÃ‰", "Zé")
        celula.Value = Replace(celula.Value, "INSTÃVEL ", "INSTÁVEL ")
        celula.Value = Replace(celula.Value, "ESTÃVEL ", "ESTÁVEL ")
        celula.Value = Replace(celula.Value, "ÃGUA", "ÁGUA")
        celula.Value = Replace(celula.Value, "CachaÃ§a", "Cachaça")
        celula.Value = Replace(celula.Value, "Ãgua", "Água")
        ' Adicione mais substituições conforme necessário
    Next celula
    
    For Each celula In Range("H2:H" & ultimaLinha)
        ' Verifica se a célula contém um valor (não está vazia)
        If Not IsEmpty(celula) Then
            ' Remove os colchetes e aspas duplas
            celula.Value = Replace(celula.Value, "[", "")
            celula.Value = Replace(celula.Value, "]", "")
            celula.Value = Replace(celula.Value, """", "")
        End If
    Next celula
  
     MsgBox "Caracteres especiais ajustados!"
End Sub
