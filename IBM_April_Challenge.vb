Option Explicit
Dim matriz() As String
Dim revision() As String
Dim Setn, Intok as integer
Dim size As Long



Sub programa()

Setn = 9
IntoK = 5
size = Application.WorksheetFunction.Combin(Setn, IntoK)
progAfin
challenge
revisar

End Sub



Sub progAfin()

Dim A, B, C, D, E As Integer
Dim contador As Integer

ReDim matriz(size, 2)
contador = 1

For A = 1 To 5
	For B = 2 To 6
		For C = 3 To 7
			For D = 4 To 8
				For E = 5 To 9
					If A < B And A < C And B < C And A < D And B < D And C < D And A < E And B < E And C < E And D < E Then
						If A <> B And A <> C And B <> C And A <> D And B <> D And C <> D And A <> E And B <> E And C <> E And D <> E Then
							matriz(contador, 1) = (CStr(A) + CStr(B) + CStr(C) + CStr(D) + CStr(E))
                            matriz(contador, 0) = "0"
                            contador = contador + 1
						Endif
					Endif
				Next
			Next
		Next
	Next
Next

End Sub




Sub challenge()

Dim iteraciones, i, j As Long
Dim numeros() As Integer
Dim localSize, pos As Integer
Dim final As String

ReDim numeros(Setn)

For iteraciones = 1 To size * 15
    localSize = Setn
    pos = 0
    For i = 1 To Setn
        numeros(i) = i
    Next
    While localSize > (Setn - Intok)
        For j = 1 To iteraciones
            If pos = Setn Then
                pos = 1
            Else
                pos = pos + 1
            End If
            If numeros(pos) = 0 Then
                j = j - 1
            End If
        Next
        numeros(pos) = 0
        localSize = localSize - 1
    Wend
    For i = 1 To Setn
        If numeros(i) <> 0 Then
            final = (CStr(final) + CStr(i))
        End If
    Next
    llenar (final)
    final = ""
Next
End Sub



Sub llenar(valor)
Dim paso As Integer

For paso = 1 To size
    If matriz(paso, 1) = valor Then
        matriz(paso, 0) = CStr(Int(matriz(paso, 0)) + 1)
    End If
Next

End Sub

Sub revisar()
Dim paso, contador As Integer
ReDim revision(size)

contador = 1

For paso = 1 To size
    If matriz(paso, 0) = "0" Then
        revision(contador) = matriz(paso, 1)
        contador = contador + 1
    End If
Next

End Sub