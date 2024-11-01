# Macros para validação de tarefas
Macro para botão acionador das macros

      Sub Apuração()
      
         Call ConcatenarUNB_PDV 
         Call ValidarQuantidadeMissao
         Call ValidarQuantidadeCompradaMes
         Call ValidarVendasDiaEDistintos
         Call ValidarVendasAgendadoEDistintos
         Call checkSku
         Call ValidarVolumeETarefas
      End Sub

Descrição dos módulos:


1 - ConcatenarUNB_PDV()

2 - ValidarQuantidadeMissao()

3 - ValidarQuantidadeCompradaMes()

4 - ValidarVendasDiaEDistintos()

5 - ValidarVendasAgendadoEDistintos()

6 - checkSku()

7 - ValidarVolumeETarefas()
