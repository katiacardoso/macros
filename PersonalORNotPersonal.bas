'--------Declarar planilha para ser usada DENTRO do Personal.xlsb


'Definir as planilhas com base na pasta de trabalho ativa
Dim wbAtivo As Workbook
Set wbAtivo = ActiveWorkbook

Set wsBase = wbAtivo.Sheets("Export")
Set wsOrigem = wbAtivo.Sheets("012011")
Set wsDePara = wbAtivo.Sheets("de_para")




'---------Declarar planilha para ser usada FORA do Personal.xlsb

Dim wsBase As Worksheet
Dim wsOrigem As Worksheet
Dim wsDePara As Worksheet

' Definir as planilhas
    Set wsBase = ThisWorkbook.Sheets("Base")
    Set wsOrigem = ThisWorkbook.Sheets("012011")
    Set wsDePara = ThisWorkbook.Sheets("de_para")
