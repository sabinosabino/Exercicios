Attribute VB_Name = "Randomico"
Private Function GenerateRnd(inicial As Integer, final As Integer) As Integer
    'Para gerar n�meros aleat�rios com VBA (Visual Basic for Applications), voc� pode usar a fun��o Rnd.
    'Esta fun��o gera um n�mero aleat�rio fracion�rio entre 0 e 1, por exemplo:
    
    'Se voc� deseja gerar um n�mero aleat�rio inteiro entre um intervalo espec�fico, pode usar a fun��o Int.
    'Por exemplo, para gerar um n�mero aleat�rio inteiro entre 1 e 10:
    'aleatorio = Int(10 * Rnd) + 1
    
    'Se voc� deseja gerar um n�mero aleat�rio dentro de um intervalo espec�fico, pode usar a seguinte f�rmula:
    '(superior - inferior + 1) * Rnd + inferior
    'Por exemplo, para gerar um n�mero aleat�rio entre 10 e 20:
    'aleatorio = (20 - 10 + 1) * Rnd + 10
    '� importante lembrar de sempre inicializar a semente do gerador de n�meros aleat�rios com a fun��o Randomize,
    'para que os n�meros gerados sejam realmente aleat�rios. Voc� pode fazer isso no in�cio do seu c�digo, por exemplo:
    'Randomize
    GenerateRnd = Int((final - inicial + 1) * Rnd + inicial)
End Function
Private Function getArrayAletarios(inicial As Integer, final As Integer, qtdItens As Integer) As Variant
    If qtdItens > final Then
        MsgBox "QtdItens n�o pode ser maior do que limite", vbCritical, "Aten��o"
        Exit Function
    End If
    Dim arr() As Integer
    ReDim arr(qtdItens - 1)
    Dim i As Integer
    Dim x As Integer
    Dim a As Integer
    Do Until arr(qtdItens - 1) > 0
        a = GenerateRnd(inicial, final)
        If i = 0 Then
            arr(i) = a
            i = i + 1
        Else
            If Not filterArray(arr, a) Then
                arr(i) = a
                i = i + 1
            End If
        End If
    Loop
    getArrayAletarios = arr
End Function
Private Function filterArray(arr() As Integer, value As Integer) As Boolean
    Dim numero As Variant
    For Each numero In arr
        If numero = value Then
            filterArray = True
            Exit For
        End If
    Next
End Function

Public Sub getNumerosAletarios(inicial As Integer, final As Integer, qtdItens As Integer)
On Error GoTo 1
    Dim arr() As Integer
    Dim item As Variant
    Dim i As Integer
    arr = getArrayAletarios(inicial, final, qtdItens)
    arr = Classificar(arr)
    For Each item In arr
        Debug.Print "Numero " & i + 1 & ":"; arr(i)
        i = i + 1
    Next
    Exit Sub
1:
    
End Sub
Private Function Classificar(arr() As Integer) As Variant
    'Por exemplo, o seguinte c�digo implementa o
    'bubble sort para ordenar um array de n�meros de forma crescente:
    
    Dim trocou As Boolean
    Dim x As Integer
    Dim qtd As Integer
    Dim aux As Integer
    Dim arrNumeros() As Integer
    
    arrNumeros = arr
    qtd = UBound(arrNumeros) - 1
    
    For x = 0 To qtd
        For i = 0 To UBound(arrNumeros) - 1
            If arrNumeros(i) > arrNumeros(i + 1) Then
                aux = arrNumeros(i)
                arrNumeros(i) = arrNumeros(i + 1)
                arrNumeros(i + 1) = aux
            End If
        Next
    Next x
    
    Classificar = arrNumeros

End Function
