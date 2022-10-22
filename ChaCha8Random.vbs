Option Explicit

Call Main()

Sub Main()
	Dim r
	Set r = New CipherRandom

	Dim t
	t = Timer()
	
	Dim i
	For i = 1 To 100
		Dim n
		n = r.NextInt()
		Call WriteLine(n)
	Next
	
	t = Timer() - t
	Call WriteLine(t)
End Sub

Sub TestChaCha8()
	Dim expected
	expected = Array(64,224,168,233,24,129,61,173,32,184,132,188,36,243,206,82,222,86,162,251,48,170,140,74,39,2,171,177,205,38,41,228,159,42,44,55,67,23,231,68,168,198,192,156,38,54,64,240,201,157,223,251,66,130,59,175,142,38,35,25,172,110,71,76,195,191,18,103,159,133,246,149,115,144,42,83,73,100,72,125,252,191,124,64,129,233,6,131,184,110,216,185,83,86,208,53,11,85,21,141,204,169,193,2,170,163,93,120,195,9,33,120,239,10,210,150,91,184,252,155,23,120,181,213,245,111,59,223,79,254,104,17,181,44,66,185,95,79,123,64,135,193,224,147,40,41,137,233,178,238,210,40,154,8,35,217,186,221,79,88,234,170,196,237,184,217,220,161,188,31,196,194,214,188,151,80,207,80,84,141,9,81,226,40,36,181,173,95,7,91,90,42,30,31,204,100,19,6,2,216,108,159,125,101,9,148,51,79,197,125,34,58,157,17,227,164,135,64,230,237,4,16,148,62,70,151,255,252,151,19,251,130,95,166,182,204,147,5,20,11,130,3,207,160,77,194,68,20,99,209,108,232,223,139,193,27,55,237,111,91,244,67,38,233,76,86,40,177,202,144,65,81,215,6,244,134,226,245,40,82,136,86,49,184,44,179,47,178,204,189,182,158,59,71,11,91,204,117,78,100,134,28,55,77,212,228,248,14,84,81,171,228,245,145,127,192,31,177,76,4,179,149,175,129,24,214,45,149,79,98,219,243,196,253,220,228,60,120,215,182,82,130,235,220,59,107,71,170,254,188,145,203,246,50,36,91,4,120,214,110,159,212,15,119,197,84,239,162,200,13,53,159,174,150,41,126,28,33,39,252,84,226,190,169,234,213,156,216,64,161,210,105,219,178,146,114,23,209,40,223,20,179,19,93,41,120,122,193,149,248,158,69,73,0,35,15,17,14,82,104,166,164,36,181,210,41,208,183,130,41,192,104,39,70,101,117,171,131,65,166,25,175,157,117,112,104,181,31,144,89,2,23,187,17,193,203,82,148,137,49,168,218,55,108,149,80,181,2,170,205,188,104,231,11,184,171,177,124,26,244,21,215,125,209,25,104,207,235,215,126,12,73,77,65,34,8,176,110,90,244,66,138,49,191,29,92,112,23,241,215,126,182,229,121,213,203,45,202,95,146,179,125,120,158,181,99,150,44,88,214,46,62,84,183,209,126,151,23,169,254,253,11,29,125,81,41,170,33,22,211,235,208,71,54,159,181,247,105,39,7,127,225,54,196,95,147,17,13,184,185,167,162,97,96,238,76,222,80,92,125,65,178,158,136,171,161,148,213,227,125,170,230,137,142,101,193,66,229,114,10,174,236,146,108,127,108,13,223,254,115,28,62,192,169,81,158,249,96,226,57,230,117,250,91,154,38,12,116,118,171,233,113,79,162,12,71,192,47,14,234,179,136,140,77,188,4,18,12,174,160,144,190,36,215,122,195,63,170,103,241,81,107,208,36,51,69,213,69,135,178,72,137,215,83,155,240,235,171,20,120,1,88,48,63,218,124,44,140,66,231,174,247,109,104,14,65,140,30,88,43,54,92,147,163,75,219,118,87,133,7,212,230,30,61,66,236,226,11,97,22,187,148,22,77,17,120,236,232,79,155,61,115,214,142,23,31,143,235,129,27,92,91,141,71,115,221,157,195,190,37,115,241,81,148,15,74,226,157,48,182,17,148,216,69,204,110,134,145,230,115,64,110,241,210,152,226,159,164,127,15,153,190,122,34,147,189,107,191,228,135,248,49,126,44,161,205,193,46,168,240,60,239,119,237,71,189,53,67,187,178,64,81,245,242,163,157,156,82,170,101,195,176,39,237,150,152,198,154,120,224,246,71,67,146,173,12,43,110,70,47,205,225,81,48,223,68,233,184,7,221,49,140,58,191,86,233,210,4,198,39,184,156,118,65,12,95,91,110,175,135,52,102,81,67,153,102,164,57,168,175,253,225,19,26,212,215,117,40,246,122,254,248,74,181,113,109,214,163,160,75,42,47,76,101,168,77,204,247,160,238,88,118,71,163,143,168,108,254,157,165,104,139,66,20,168,31,81,193,25,83,32,36,53,152,194,108,133,128,117,72,114,60,164,244,178,178,211,252,45,218,253,63,102,63,224,227,81,136,144,200,118,10,252,248,72,83,147,65,93,165,35,55,106,44,193,51,251,133,218,164,25,164,186,185,33,87,79,19,234,14,223,127,99,157,76,31,168,190,19,76)
	
	Dim chacha
	Set chacha = New ChaCha8
	
	Call chacha.SetKey(Seq(32))
	Call chacha.SetIV(Seq(8))
	
	Dim message
	message = seq(1000)
	Call chacha.Encrypt(message)
	
	Call WriteLine(ArrayEquals(expected, message))
End Sub

Function ArrayEquals(a, b)
	If UBound(a) <> UBound(b) Then
		ArrayEquals = False
		Exit Function
	End If
	
	Dim i
	For i = 0 To UBound(a)
		If a(i) <> b(i) Then
			ArrayEquals = False
			Exit Function
		End If
	Next
	
	ArrayEquals = True
End Function

Function Seq(n)
	Dim a
	ReDim a(n - 1)
	
	Dim i
	For i = 0 To UBound(a)
		a(i) = i And &HFF
	Next
	
	Seq = a
End Function

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
	I8ToI32 = a(i + 0) Or Shl(a(i + 1), 8) Or Shl(a(i + 2), 16) Or Shl(a(i + 3), 24)
End Function

Function Slice(src, offset, count)
	Dim dest
	ReDim dest(count - 1)
	
	Dim i
	For i = 0 To UBound(dest)
		dest(i) = src(offset + i)
	Next
	
	Slice = dest
End Function

Class ChaCha8
	Private State
	
	Private Sub Class_Initialize()
		ReDim State(15)
	End Sub
	
	Sub SetKey(b)
		State( 0) = 1634760805
		State( 1) = 857760878
		State( 2) = 2036477234
		State( 3) = 1797285236
		State( 4) = I8ToI32(b,  0)
		State( 5) = I8ToI32(b,  4)
		State( 6) = I8ToI32(b,  8)
		State( 7) = I8ToI32(b, 12)
		State( 8) = I8ToI32(b, 16)
		State( 9) = I8ToI32(b, 20)
		State(10) = I8ToI32(b, 24)
		State(11) = I8ToI32(b, 28)
	End Sub
	
	Sub SetIV(b)
		State(12) = 0
		State(13) = 0
		State(14) = I8ToI32(b, 0)
		State(15) = I8ToI32(b, 4)
	End Sub
	
	Sub Encrypt(m)
		Dim j
		j = 0
		
		Dim b
		ReDim b(63)
		
		Dim l
		For l = UBound(m) To 0 Step -64
			Call Stir(b)
			Call CountUp()
			
			Dim i
			For i = 0 To 63
				If j > UBound(m) Then
					Exit For
				End If
				m(j) = m(j) Xor b(i)
				j = j + 1
			Next
		Next
	End Sub
	
	Sub CountUp()
		State(12) = I32(State(12) + 1)
		
		If State(12) = 0 Then
			State(13) = I32(State(13) + 1)
		End If
	End Sub

	Sub Stir(b)
		Dim x
		ReDim x(15)

		Dim i
		
		For i = 0 To 15
			x(i) = State(i)
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
			x(i) = I32(x(i) + State(i))
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

		Dim seed
		Set seed = New SecureSeed
		Call Cipher.SetKey(seed.Seed(32))
		Call Cipher.SetIV(seed.Seed(8))
		
		Call Fill()
		Call Fill()
	End Sub
	
	Sub Fill()
		Call Cipher.Encrypt(Buffer)
		Call Cipher.SetKey(Slice(Buffer, 0, 32))
		Call Cipher.SetIV(Slice(Buffer, 32, 8))
		
		Dim i
		For i = 0 To 39
			Buffer(i) = 0
		Next

		Index = 40
	End Sub

	Function NextInt()
		If Index > UBound(Buffer) - 4 Then
			Call Fill()
		End If
		
		NextInt = I8ToI32(Buffer, Index)
		Index = Index + 4
	End Function
	
	Function NextBound(bound)
		Const m = &H7FFFFFFF
		
		Dim r
		r = (m - bound + 1) Mod bound
		
		Do
			Dim n
			n = NextInt() And m
			
			If r <= n Then
				NextBound = n Mod bound
				Exit Function
			End If
		Loop
	End Function
	
	Sub Shuffle(a)
		Dim i
		For i = UBound(a) To 1 Step -1
			Dim j
			j = NextBound(i + 1)
			
			Dim t
			t = a(i)
			a(i) = a(j)
			a(j) = t
		Next
	End Sub
End Class

Class SecureSeed
	Private Rijndael
	Private Buffer
	Private Index
	
	Private Sub Class_Initialize()
		Set Rijndael = CreateObject("System.Security.Cryptography.RijndaelManaged")
		Buffer = ""
		Index = LenB(Buffer) + 1
	End Sub
	
	Function GetByte()
		If Index > LenB(Buffer) Then
			Call Rijndael.GenerateKey()
			Buffer = CStr(Rijndael.Key)
			Index = 1
		End If
		
		GetByte = AscB(MidB(Buffer, Index, 1))
		Index = Index + 1
	End Function
	
	Function Seed(n)
		Dim a
		ReDim a(n - 1)
		
		Dim i
		For i = 0 To UBound(a)
			a(i) = GetByte()
		Next

		Seed = a
	End Function
End Class

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub
