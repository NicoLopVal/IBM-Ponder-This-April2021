Option Explicit
Dim matriz() As String
Dim revision() As String

Sub programa()

progAfin
challenge
revisar

End Sub



Sub progAfin()

Dim A, B, C, D, E As Integer
Dim F, G, H, I, J, K As Integer
Dim contador As Integer

ReDim matriz(1365, 2)
contador = 1

For A = 1 To 5
    For B = 2 To 6
        For C = 3 To 7
            For D = 4 To 8
                For E = 5 To 9
                    For F = 6 To 10
                        For G = 7 To 11
							for H = 8 to 12
								for I = 9 to 13
									for J = 10 to 14
										for K = 11 to 15
											If A < B And A < C And A < D And A < E And A < F And A < G And A < H And A < I And A < J And A < K And B < C And B < D And B < E And B < F And B < G And B < H And B < I And B < J And B < K And C < D And C < E And C < F And C < G And C < H And C < I And C < J And C < K And D < E And D < F And D < G And D < H And D < I And D < J And D < K And E < F And E < G And E < H And E < I And E < J And E < K And F < G And F < H And F < I And F < J And F < K And G < H And G < I And G < J And G < K And H < I And H < J And H < K And I < J And I < K And J < K then
												If A <> B And A <> C And A <> D And A <> E And A <> F And A <> G And A <> H And A <> I And A <> J And A <> K And B <> C And B <> D And B <> E And B <> F And B <> G And B <> H And B <> I And B <> J And B <> K And C <> D And C <> E And C <> F And C <> G And C <> H And C <> I And C <> J And C <> K And D <> E And D <> F And D <> G And D <> H And D <> I And D <> J And D <> K And E <> F And E <> G And E <> H And E <> I And E <> J And E <> K And F <> G And F <> H And F <> I And F <> J And F <> K And G <> H And G <> I And G <> J And G <> K And H <> I And H <> J And H <> K And I <> J And I <> K And J <> K then
													matriz(contador, 1) = (CStr(A) + CStr(B) + CStr(C) + CStr(D) + CStr(E) + CStr(F) + CStr(G) + CStr(H) + CStr(I) + CStr(J) + CStr(K))
													matriz(contador, 0) = "0"
													contador = contador + 1
												endif
											End If
										next
									next
								next
							next
                        Next
                    Next
                Next
            Next
        Next
    Next
Next

End Sub

Sub challenge()
Dim iteraciones, i, j As Integer
Dim numeros() As Integer
Dim size, pos As Integer
Dim final As String

ReDim numeros(15)

For iteraciones = 1 To 32000
    size = 15
    pos = 0
    For i = 1 To 15
        numeros(i) = i
    Next
    While size > 11
        For j = 1 To iteraciones
            If pos = 15 Then
                pos = 1
            Else
                pos = pos + 1
            End If
            If numeros(pos) = 0 Then
                j = j - 1
            End If
        Next
        numeros(pos) = 0
        size = size - 1
    Wend
    For i = 1 To 15
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

For paso = 1 To 1365
    If matriz(paso, 1) = valor Then
        matriz(paso, 0) = CStr(Int(matriz(paso, 0)) + 1)
    End If
Next

End Sub

Sub revisar()
Dim paso, contador As Integer
ReDim revision(1365)

contador = 1

For paso = 1 To 1365
    If matriz(paso, 0) = "0" Then
        revision(contador) = matriz(paso, 1)
        contador = contador + 1
    End If
Next

End Sub