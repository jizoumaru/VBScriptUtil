Option Explicit

Call Main()

Sub Main()
	Dim r, a, i
	
	Set r = New CipherRandom
	a = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
	
	For i = 1 To 10
		Call r.Shuffle(a)
		Call WriteLine(Join(a, ","))
	Next
End Sub

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
	Function Encrypt(m)
		Dim s, j, b, l, i
		
		ReDim s(15)
		s( 4) = I8ToI32(m,  0)
		s( 5) = I8ToI32(m,  4)
		s( 6) = I8ToI32(m,  8)
		s( 7) = I8ToI32(m, 12)
		s( 8) = I8ToI32(m, 16)
		s( 9) = I8ToI32(m, 20)
		s(10) = I8ToI32(m, 24)
		s(11) = I8ToI32(m, 28)
		s( 0) = 1634760805
		s( 1) = 857760878
		s( 2) = 2036477234
		s( 3) = 1797285236
		s(12) = 0
		s(13) = 0
		s(14) = I8ToI32(m, 32)
		s(15) = I8ToI32(m, 36)

		j = 0
		ReDim b(63)
		
		For l = UBound(m) To 0 Step -64
			Call Stir(s, b)
			
			s(12) = U32(s(12) + 1)
			
			If s(12) = 0 Then
				s(13) = U32(s(13) + 1)
			End If
			
			For i = 0 To 63
				m(j) = m(j) Xor b(i)
				j = j + 1
			Next
		Next
		
		Encrypt = 40
	End Function

	Sub Stir(s, b)
		Dim x, i
		ReDim x(15)

		For i = 0 To 15
			x(i) = s(i)
		Next

		For i = 8 To 1 Step -2
			Call QR(x, 0, 4,  8, 12)
			Call QR(x, 1, 5,  9, 13)
			Call QR(x, 2, 6, 10, 14)
			Call QR(x, 3, 7, 11, 15)
			Call QR(x, 0, 5, 10, 15)
			Call QR(x, 1, 6, 11, 12)
			Call QR(x, 2, 7,  8, 13)
			Call QR(x, 3, 4,  9, 14)
		Next
		
		For i = 0 To 15
			x(i) = I32(x(i) + s(i))
			b(i * 4 + 0) = x(i) And &HFF
			b(i * 4 + 1) = Shr(x(i),  8) And &HFF
			b(i * 4 + 2) = Shr(x(i), 16) And &HFF
			b(i * 4 + 3) = Shr(x(i), 24) And &HFF
		Next
	End Sub
	
	Sub QR(x, a, b, c, d)
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

Class CipherRandom
	Private Cipher
	Private Buffer
	Private Index
	
	Private Sub Class_Initialize()
		Set Cipher = New ChaCha8
		ReDim Buffer(1023)
		Call Init(Buffer)
		Index = UBound(Buffer) + 1
	End Sub

	Sub Init(a)
		Dim s, i, d, t
		
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
	End Sub
	
	Function NextByte()
		If Index > UBound(Buffer) Then
			Index = Cipher.Encrypt(Buffer)
		End If
		
		NextByte = Buffer(Index)
		Index = Index + 1
	End Function

	Function NextInt()
		If Index > UBound(Buffer) - 4 Then
			Index = Cipher.Encrypt(Buffer)
		End If
		
		NextInt = I8ToI32(Buffer, Index)
		Index = Index + 4
	End Function
	
	Function NextBound(bound)
		Const m = &H7FFFFFFF
		Dim r, n
		
		r = (m - bound + 1) Mod bound
		
		Do
			n = NextInt()
			
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
