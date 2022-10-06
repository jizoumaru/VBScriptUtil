Option Explicit

Function I32(n)
	If n > 2147483647 Then
		I32 = n - 4294967296
	ElseIf n < -2147483648 Then
		I32 = n + 4294967296
	Else
		I32 = n
	End If
End Function

Function U32(n)
	If n < 0 Then
		U32 = n + 4294967296
	Else
		U32 = n
	End If
End Function

Function Shr(v, n)
	Shr = I32(Fix(U32(v) / (2 ^ n)))
End Function

Function Shl(v, n)
	If n = 0 Then
		Shl = I32(v)
		Exit Function
	End If
	Shl = I32((v And ((2 ^ (32 - n)) - 1)) * (2 ^ n))
End Function

Function Rotl(v, n)
	Rotl = Shl(v, n) Or Shr(v, 32 - n)
End Function

Class ClassSalsa20
	Private State
	
	Private Sub Class_Initialize()
		ReDim State(15)
	End Sub

	Sub Init(b)
		State( 1) = ToI32(b,  0)
		State( 2) = ToI32(b,  4)
		State( 3) = ToI32(b,  8)
		State( 4) = ToI32(b, 12)
		State(11) = ToI32(b, 16)
		State(12) = ToI32(b, 20)
		State(13) = ToI32(b, 24)
		State(14) = ToI32(b, 28)
		State( 0) = 1634760805
		State( 5) = 857760878
		State(10) = 2036477234
		State(15) = 1797285236
		State( 6) = ToI32(b, 32)
		State( 7) = ToI32(b, 36)
		State( 8) = 0
		State( 9) = 0
	End Sub

	Function ToI32(a, i)
		ToI32 = a(i + 0) Or Shl(a(i + 1), 8) Or Shl(a(i + 2), 16) Or Shl(a(i + 3), 24)
	End Function
	
	Sub Encrypt(m)
		Dim j
		j = 0
		
		Dim b
		ReDim b(63)
		
		Dim l
		For l = UBound(m) To 0 Step -64
			Call Stir(b)
			
			State(8) = I32(State(8) + 1)
			
			If State(8) = 0 Then
				State(9) = I32(State(9) + 1)
			End If
			
			Dim i
			For i = 0 To 63
				m(j) = m(j) Xor b(i)
				j = j + 1
			Next
		Next
	End Sub

	Sub Stir(b)
		Dim x
		ReDim x(15)

		Dim i
		
		For i = 0 To 15
			x(i) = State(i)
		Next

		For i = 20 To 1 Step -2
			x( 4) = x( 4) Xor Rotl(I32(x( 0) + x(12)),  7)
			x( 8) = x( 8) Xor Rotl(I32(x( 4) + x( 0)),  9)
			x(12) = x(12) Xor Rotl(I32(x( 8) + x( 4)), 13)
			x( 0) = x( 0) Xor Rotl(I32(x(12) + x( 8)), 18)
			x( 9) = x( 9) Xor Rotl(I32(x( 5) + x( 1)),  7)
			x(13) = x(13) Xor Rotl(I32(x( 9) + x( 5)),  9)
			x( 1) = x( 1) Xor Rotl(I32(x(13) + x( 9)), 13)
			x( 5) = x( 5) Xor Rotl(I32(x( 1) + x(13)), 18)
			x(14) = x(14) Xor Rotl(I32(x(10) + x( 6)),  7)
			x( 2) = x( 2) Xor Rotl(I32(x(14) + x(10)),  9)
			x( 6) = x( 6) Xor Rotl(I32(x( 2) + x(14)), 13)
			x(10) = x(10) Xor Rotl(I32(x( 6) + x( 2)), 18)
			x( 3) = x( 3) Xor Rotl(I32(x(15) + x(11)),  7)
			x( 7) = x( 7) Xor Rotl(I32(x( 3) + x(15)),  9)
			x(11) = x(11) Xor Rotl(I32(x( 7) + x( 3)), 13)
			x(15) = x(15) Xor Rotl(I32(x(11) + x( 7)), 18)
			x( 1) = x( 1) Xor Rotl(I32(x( 0) + x( 3)),  7)
			x( 2) = x( 2) Xor Rotl(I32(x( 1) + x( 0)),  9)
			x( 3) = x( 3) Xor Rotl(I32(x( 2) + x( 1)), 13)
			x( 0) = x( 0) Xor Rotl(I32(x( 3) + x( 2)), 18)
			x( 6) = x( 6) Xor Rotl(I32(x( 5) + x( 4)),  7)
			x( 7) = x( 7) Xor Rotl(I32(x( 6) + x( 5)),  9)
			x( 4) = x( 4) Xor Rotl(I32(x( 7) + x( 6)), 13)
			x( 5) = x( 5) Xor Rotl(I32(x( 4) + x( 7)), 18)
			x(11) = x(11) Xor Rotl(I32(x(10) + x( 9)),  7)
			x( 8) = x( 8) Xor Rotl(I32(x(11) + x(10)),  9)
			x( 9) = x( 9) Xor Rotl(I32(x( 8) + x(11)), 13)
			x(10) = x(10) Xor Rotl(I32(x( 9) + x( 8)), 18)
			x(12) = x(12) Xor Rotl(I32(x(15) + x(14)),  7)
			x(13) = x(13) Xor Rotl(I32(x(12) + x(15)),  9)
			x(14) = x(14) Xor Rotl(I32(x(13) + x(12)), 13)
			x(15) = x(15) Xor Rotl(I32(x(14) + x(13)), 18)
		Next
		
		For i = 0 To 15
			x(i) = I32(x(i) + State(i))
			b(i * 4 + 0) = x(i) And &HFF
			b(i * 4 + 1) = Shr(x(i),  8) And &HFF
			b(i * 4 + 2) = Shr(x(i), 16) And &HFF
			b(i * 4 + 3) = Shr(x(i), 24) And &HFF
		Next
	End Sub
End Class

Function Seq(n)
	Dim a
	ReDim a(n - 1)
	
	Dim i
	For i = 0 To UBound(a)
		a(i) = i And &HFF
	Next
	
	Seq = a
End Function

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub

Sub Main()
	Dim salsa
	Set salsa = New ClassSalsa20
	
	Call salsa.Init(Seq(40))
	
	Dim m
	m = Seq(65536)
	
	Call salsa.Encrypt(m)
	
	Dim i
	For i = 0 To UBound(m)
		Call WriteLine(m(i))
	Next
End Sub

Call Main()
