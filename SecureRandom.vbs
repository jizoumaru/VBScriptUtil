Option Explicit

Function Shr(v, n)
	Shr = I32(Fix(U32(v) / (2 ^ n)))
End Function

Function Shl(v, n)
	Shl = I32((v And ((2 ^ (32 - n)) - 1)) * (2 ^ n))
End Function

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

Class Salsa20
	Function Rotate(v, n)
		Rotate = Shl(v, n) Or Shr(v, 32 - n)
	End Function

	Function ToI32(a, i)
		ToI32 = a(i + 0) Or Shl(a(i + 1), 8) Or Shl(a(i + 2), 16) Or Shl(a(i + 3), 24)
	End Function
	
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

	Sub Encrypt(m, c)
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
				c(j) = m(j) Xor b(i)
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
			x( 4) = x( 4) Xor Rotate(I32(x( 0) + x(12)),  7)
			x( 8) = x( 8) Xor Rotate(I32(x( 4) + x( 0)),  9)
			x(12) = x(12) Xor Rotate(I32(x( 8) + x( 4)), 13)
			x( 0) = x( 0) Xor Rotate(I32(x(12) + x( 8)), 18)
			x( 9) = x( 9) Xor Rotate(I32(x( 5) + x( 1)),  7)
			x(13) = x(13) Xor Rotate(I32(x( 9) + x( 5)),  9)
			x( 1) = x( 1) Xor Rotate(I32(x(13) + x( 9)), 13)
			x( 5) = x( 5) Xor Rotate(I32(x( 1) + x(13)), 18)
			x(14) = x(14) Xor Rotate(I32(x(10) + x( 6)),  7)
			x( 2) = x( 2) Xor Rotate(I32(x(14) + x(10)),  9)
			x( 6) = x( 6) Xor Rotate(I32(x( 2) + x(14)), 13)
			x(10) = x(10) Xor Rotate(I32(x( 6) + x( 2)), 18)
			x( 3) = x( 3) Xor Rotate(I32(x(15) + x(11)),  7)
			x( 7) = x( 7) Xor Rotate(I32(x( 3) + x(15)),  9)
			x(11) = x(11) Xor Rotate(I32(x( 7) + x( 3)), 13)
			x(15) = x(15) Xor Rotate(I32(x(11) + x( 7)), 18)
			x( 1) = x( 1) Xor Rotate(I32(x( 0) + x( 3)),  7)
			x( 2) = x( 2) Xor Rotate(I32(x( 1) + x( 0)),  9)
			x( 3) = x( 3) Xor Rotate(I32(x( 2) + x( 1)), 13)
			x( 0) = x( 0) Xor Rotate(I32(x( 3) + x( 2)), 18)
			x( 6) = x( 6) Xor Rotate(I32(x( 5) + x( 4)),  7)
			x( 7) = x( 7) Xor Rotate(I32(x( 6) + x( 5)),  9)
			x( 4) = x( 4) Xor Rotate(I32(x( 7) + x( 6)), 13)
			x( 5) = x( 5) Xor Rotate(I32(x( 4) + x( 7)), 18)
			x(11) = x(11) Xor Rotate(I32(x(10) + x( 9)),  7)
			x( 8) = x( 8) Xor Rotate(I32(x(11) + x(10)),  9)
			x( 9) = x( 9) Xor Rotate(I32(x( 8) + x(11)), 13)
			x(10) = x(10) Xor Rotate(I32(x( 9) + x( 8)), 18)
			x(12) = x(12) Xor Rotate(I32(x(15) + x(14)),  7)
			x(13) = x(13) Xor Rotate(I32(x(12) + x(15)),  9)
			x(14) = x(14) Xor Rotate(I32(x(13) + x(12)), 13)
			x(15) = x(15) Xor Rotate(I32(x(14) + x(13)), 18)
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

Class SecureRandom
	Private Salsa
	Private Buffer
	Private Index
	
	Private Sub Class_Initialize()
		Set Salsa = New Salsa20
		ReDim Buffer(1023)
		Index = UBound(Buffer) + 1
		Call InitKey()
		Call InitNonce()
	End Sub

	Sub InitKey()
		Dim s
		s = Replace(GetUUID(), "-", "")
		
		Dim i
		For i = 0 To 31
			Buffer(i) = CByte("&H" & Mid(s, i + 1, 1))
		Next
	End Sub
	
	Sub InitNonce()
		Dim d
		d = Now
		
		Dim t
		t = Timer()
		
		Buffer(32) = Year(d) \ 100 Mod 100
		Buffer(33) = Year(d) Mod 100
		Buffer(34) = Month(d)
		Buffer(35) = Day(d)
		Buffer(36) = Fix(t) \ 60 \ 60
		Buffer(37) = Fix(t) \ 60 Mod 60
		Buffer(38) = Fix(t) Mod 60
		Buffer(39) = Fix((t - Fix(t)) * 100)
	End Sub

	Function GetUUID()
		Dim wmi
		Set wmi = GetObject("winmgmts:\\.\root\cimv2")
		
		Dim items
		Set items = wmi.ExecQuery("Select * from Win32_ComputerSystemProduct")
		
		Dim item
		Set item = items.ItemIndex(0)
		
		GetUUID = item.UUID
	End Function
	
	Sub Fill()
		Call Salsa.Init(Buffer)
		Call Salsa.Encrypt(Buffer, Buffer)
		Index = 40
	End Sub
	
	Function NextByte()
		If Index > UBound(Buffer) Then
			Call Fill()
		End If
		
		NextByte = Buffer(Index)
		Index = Index + 1
	End Function

	Function NextInt()
		If Index > UBound(Buffer) - 4 Then
			Call Fill()
		End If
		
		NextInt = Buffer(Index) _
			Or Buffer(Index + 1) * &H100 _
			Or Buffer(Index + 2) * &H10000 _
			Or I32(Buffer(Index + 3) * &H1000000)
			
		Index = Index + 4
	End Function
End Class

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub

Sub Main()
	Dim random
	Set random = New SecureRandom
	
	Dim i
	For i = 1 To 100
		Call WriteLine(random.NextInt())
	Next
End Sub

Call Main()
