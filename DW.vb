Public Function EW() As String
    Dim e As Integer
    Dim s As Integer
    Dim lanzamientos As String
    Dim resultado As String
    Dim i As Integer
    Dim dado As Integer

    ' Solicitar el número de lanzamientos
    Do
        e = InputBox("Ingresa el número de lanzamientos (1-100):", "EW")
        If e < 1 Or e > 100 Then
            MsgBox "Número inválido. Debe ser entre 1 y 100."
        End If
    Loop While e < 1 Or e > 100

    ' Realizar los lanzamientos y calcular la suma
    s = 0
    lanzamientos = ""
    For i = 1 To e
        dado = WorksheetFunction.RandBetween(1, 6)
        s = s + dado
        lanzamientos = lanzamientos & dado & ", "
    Next i

    ' Eliminar la última coma y espacio
    If Len(lanzamientos) > 0 Then
        lanzamientos = Left(lanzamientos, Len(lanzamientos) - 2)
    End If

    ' Mostrar los resultados
    MsgBox "Lanzamientos: " & lanzamientos & vbCrLf & vbCrLf & "Suma total: " & s

    ' Determinar el resultado
    Select Case s
        Case Is = 100
            resultado = "¡Éxito! Has alcanzado exactamente 100 puntos."
        Case Is < 100
            resultado = "Te faltan " & (100 - s) & " puntos para llegar a 100."
        Case Else
            resultado = "Te has pasado por " & (s - 100) & " puntos del objetivo de 100."
    End Select

    ' Mostrar el resultado final
    MsgBox resultado

    ' Devolver el resultado
    EW = resultado
End Function
