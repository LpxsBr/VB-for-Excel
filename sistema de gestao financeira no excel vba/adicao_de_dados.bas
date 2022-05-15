Sub add()

Dim data_ocorreencia As Date

    ' recebendo as variaveis do campo de adição
    
    'referencia range é a celula do excel que ta sendo usada
    ' por exemplo, Range("A1") é referente a celula de excel A1
    ' posicionada na coluna A e na linha 1
    
    tipo_ocorrencia = Range("B5")
    data_ocorreencia = Range("C5")
    descricao_ocorrencia = Range("D5")
    categoria_ocorrencia = Range("E5")
    valor_ocorrencia = Range("F5")
    
    ' tratando exibição do dado para letra maiscula up case
    
    case_tipo_ocorrencia = UCase(tipo_ocorrencia)
    
    'transferindo as primeiras variaveis para o local de historico
    
    Range("J4").End(xlDown).Offset(1, 0).Value = case_tipo_ocorrencia
    Range("K4").End(xlDown).Offset(1, 0).Value = data_ocorrencia
    Range("L4").End(xlDown).Offset(1, 0).Value = descricao_ocorrencia
    Range("M4").End(xlDown).Offset(1, 0).Value = categoria_ocorrencia
    
    'teste if, elseif pra verificar como a ultima variavel deve se alocar
    ' se o tipo for igual a D de Despesa, o valor da ocorrencia será multiplicado
    'por 1 para se tornar negativo no historico
    ' caso não seja, apenas imprimirá o valor sem alterações
    
    If case_tipo_ocorrencia = "D" Then
        Range("N4").End(xlDown).Offset(1, 0).Value = valor_ocorrencia * -1
    ElseIf case_tipo_ocorrencia = "R" Then
        Range("N4").End(xlDown).Offset(1, 0).Value = valor_ocorrencia
    End If
    
    'limpa o campo de adição
    
End Sub
