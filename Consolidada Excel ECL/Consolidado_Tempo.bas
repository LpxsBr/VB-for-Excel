Attribute VB_Name = "Módulo2"
Sub at()
'
' Dev by: LpxsBr
' eCOCOLab
'
' declaração da variavel linha para iniciar uma contagem

Dim linha As Integer

linha = 1

' Estrutura de repetição para que o processo ocorra enquanto e quando a linha de referencia esteja preenchida

While Cells(linha + 1, 1) <> ""

' Preenchimento da Celula I2 em diante com o somase do Tempo por referencia
' Ocorre um =somase a cada ref de linha preenchida com ativação de botão
' A celula A2 (+pulos) recebe o valor dado por somase das Referencias de tempo na planilha Tempo

Cells(linha + 1, 9).Value = "=SUMIF(Tempo!C[-8],Consolidado!RC[-8],Tempo!C[-6])"

' O conjunto Cells(linha+1, 9) e o incremento linha = linha + 1 (linha++) garante o pulo de linha

linha = linha + 1

    
Wend

' Mensagem
MsgBox ("Planilha atualizada")

End Sub



