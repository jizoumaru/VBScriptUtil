Option Explicit

Call Main()

Sub Main()
	Dim r, i, t
	
	Set r = New SecureRandom
	
	t = Timer()
	
	For i = 1 To 1000
		Call WriteLine(r.NextInt())
	Next
	
	t = Timer() - t
	Call WriteLine(t)
End Sub

Function Shr(v, n)
	Shr = I32(Fix(U32(v) / (2 ^ n)))
End Function

Function Shl(v, n)
	If n = 0 Then
		Shl = v
	Else
		Shl = I32((v And ((2 ^ (32 - n)) - 1)) * (2 ^ n))
	End If
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
	If n > 4294967295 Then
		U32 = n - 4294967296
	ElseIf n < 0 Then
		U32 = n + 4294967296
	Else
		U32 = n
	End If
End Function

Function Rotl(v, n)
	Rotl = Shl(v, n) Or Shr(v, 32 - n)
End Function

Function I8ToI32(a, i)
	I8ToI32 = a(i + 0) _
		Or a(i + 1) * &H100 _
		Or a(i + 2) * &H10000 _
		Or I32(a(i + 3) * &H1000000)
End Function

Function GetUUID()
	Dim wmi, items, item
	Set wmi = GetObject("winmgmts:\\.\root\cimv2")
	Set items = wmi.ExecQuery("Select * from Win32_ComputerSystemProduct")
	Set item = items.ItemIndex(0)
	GetUUID = item.UUID
End Function

Class ChaCha8
	Private State
	
	Private Sub Class_Initialize()
		ReDim State(15)
	End Sub

	Sub Init(b)
		State( 4) = I8ToI32(b,  0)
		State( 5) = I8ToI32(b,  4)
		State( 6) = I8ToI32(b,  8)
		State( 7) = I8ToI32(b, 12)
		State( 8) = I8ToI32(b, 16)
		State( 9) = I8ToI32(b, 20)
		State(10) = I8ToI32(b, 24)
		State(11) = I8ToI32(b, 28)
		State( 0) = 1634760805
		State( 1) = 857760878
		State( 2) = 2036477234
		State( 3) = 1797285236
		State(12) = 0
		State(13) = 0
		State(14) = I8ToI32(b, 32)
		State(15) = I8ToI32(b, 36)
	End Sub
	
	Sub Encrypt(m)
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
			b(i * 4 + 1) = Shr(x(i),  8) And &HFF
			b(i * 4 + 2) = Shr(x(i), 16) And &HFF
			b(i * 4 + 3) = Shr(x(i), 24) And &HFF
		Next
	End Sub
	
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
End Class

Class SecureRandom
	Private Cipher
	Private Buffer
	Private Index
	
	Private Sub Class_Initialize()
		Set Cipher = New ChaCha8
		ReDim Buffer(1023)
		Call Cipher.Init(Seed())
		Call ReKey()
		Call ReKey()
	End Sub

	Function Seed()
		Dim a, s, i, d, t
		
		ReDim a(39)
		s = Replace(GetUUID(), "-", "")
		
		For i = 0 To 31
			a(i) = CByte("&H" & Mid(s, i + 1, 1))
		Next

		d = Now
		t = Timer()
		
		a(32) = Year(d) \ 100 Mod 100
		a(33) = Year(d) Mod 100
		a(34) = Month(d)
		a(35) = Day(d)
		a(36) = Fix(t) \ 60 \ 60
		a(37) = Fix(t) \ 60 Mod 60
		a(38) = Fix(t) Mod 60
		a(39) = Fix((t - Fix(t)) * 100)
		
		Seed = a
	End Function
	
	Sub ReKey()
		Dim i

		Call Cipher.Encrypt(Buffer)
		Call Cipher.Init(Buffer)
		
		For i = 0 To 39
			Buffer(i) = 0
		Next
		
		Index = i
	End Sub
	
	Function NextByte()
		If Index > UBound(Buffer) Then
			Call ReKey()
		End If
		
		NextByte = Buffer(Index)
		Index = Index + 1
	End Function

	Function NextInt()
		If Index > UBound(Buffer) - 4 Then
			Call ReKey()
		End If
		
		NextInt = I8ToI32(Buffer, Index)
		Index = Index + 4
	End Function
	
	Function NextBound(bound)
		Const m = &H7FFFFFFF
		Dim r, n
		
		r = (m - bound + 1) Mod bound
		
		Do
			n = NextInt() And m
			
			If r <= n Then
				Exit Do
			End If
		Loop
		
		NextBound = n Mod bound
	End Function
	
	Sub Shuffle(a)
		Dim i, j, t
		For i = UBound(a) To 1 Step -1
			j = NextBound(i + 1)
			t = a(i)
			a(i) = a(j)
			a(j) = t
		Next
	End Sub
End Class

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub
