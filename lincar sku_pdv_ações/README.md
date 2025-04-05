Assim executa para mostrar o resultado: 10 

If regex.test(tarefa) Then
    Set matches = regex.Execute(tarefa)
    quantidade = matches(0).SubMatches(0) ' <- aqui é o correto!
Else
    quantidade = "1"
End If

Assim executa para mostrar o resultado : 10 caixas

If regex.test(tarefa) Then
    Set matches = regex.Execute(tarefa)
    quantidade = matches(0).SubMatches(0)
Else
    quantidade = "1" ' Valor padrão se não encontrar
End If
