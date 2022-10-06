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

Class ClassChaCha8
	Sub QuarterRound(x, a, b, c, d)
		x(a) = I32(x(a) + x(b))
		x(d) = Rotl(x(d) Xor x(a),16)
		x(c) = I32(x(c) + x(d))
		x(b) = Rotl(x(b) Xor x(c),12)
		x(a) = I32(x(a) + x(b))
		x(d) = Rotl(x(d) Xor x(a), 8)
		x(c) = I32(x(c) + x(d))
		x(b) = Rotl(x(b) Xor x(c), 7)
	End Sub
	
	Function ToI32(a, i)
		ToI32 = a(i + 0) Or Shl(a(i + 1), 8) Or Shl(a(i + 2), 16) Or Shl(a(i + 3), 24)
	End Function
	
	Private State
	
	Private Sub Class_Initialize()
		ReDim State(15)
	End Sub

	Sub Init(b)
		State( 4) = ToI32(b,  0)
		State( 5) = ToI32(b,  4)
		State( 6) = ToI32(b,  8)
		State( 7) = ToI32(b, 12)
		State( 8) = ToI32(b, 16)
		State( 9) = ToI32(b, 20)
		State(10) = ToI32(b, 24)
		State(11) = ToI32(b, 28)
		State( 0) = 1634760805
		State( 1) = 857760878
		State( 2) = 2036477234
		State( 3) = 1797285236
		State(12) = 0
		State(13) = 0
		State(14) = ToI32(b, 32)
		State(15) = ToI32(b, 36)
	End Sub

	Sub Encrypt(m)
		For i = 0 To UBound(State)
			Call WriteLine(i & " " & State(i))
		Next
		
		Dim j
		j = 0
		
		Dim b
		ReDim b(63)
		
		Dim l
		For l = UBound(m) To 0 Step -64
			Call Stir(b)
			
			State(12) = I32(State(12) + 1)
			
			If State(12) = 0 Then
				State(13) = I32(State(13) + 1)
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

		For i = 8 To 1 Step -2
			Call QuarterRound(x, 0, 4,  8, 12)
			Call QuarterRound(x, 1, 5,  9, 13)
			Call QuarterRound(x, 2, 6, 10, 14)
			Call QuarterRound(x, 3, 7, 11, 15)
			Call QuarterRound(x, 0, 5, 10, 15)
			Call QuarterRound(x, 1, 6, 11, 12)
			Call QuarterRound(x, 2, 7,  8, 13)
			Call QuarterRound(x, 3, 4,  9, 14)
		Next
		
		For i = 0 To 15
			x(i) = I32(x(i) + State(i))
			b(i * 4 + 0) = x(i) And &HFF
			b(i * 4 + 1) = x(i) \ &H100 And &HFF
			b(i * 4 + 2) = x(i) \ &H10000 And &HFF
			b(i * 4 + 3) = x(i) \ &H1000000 And &HFF
		Next
	End Sub
End Class

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub

Sub Main()
	Dim chacha
	Set chacha = New ClassChaCha8
	
	Dim b
	ReDim b(39)
	
	Dim i
	For i = 0 To UBound(b)
		b(i) = i
	Next
	
	Call chacha.Init(b)
	
	Dim m
	ReDim m(65535)
	
	For i = 0 To UBound(m)
		m(i) = i And &HFF
	Next
	
	For i = 0 To UBound(m)
		Call WriteLine("m " & m(i))
	Next
	
	Call chacha.Encrypt(m)
	
	For i = 0 To UBound(m)
		Call WriteLine("c " & m(i))
	Next
End Sub

Call Main()
