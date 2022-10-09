Option Explicit

Call Main()

Sub Main()
	Dim rc4
	Set rc4 = New ClassRc4
	
	Dim key
	key = Seq(32)
	
	Dim msg
	msg = Seq(256)
	
	Call rc4.Encrypt(key, msg)

	Dim i
	For i = 0 To UBound(msg)
		Call WriteLine(msg(i))
	Next
End Sub

Function Seq(n)
	Dim a
	ReDim a(n - 1)
	
	Dim i
	For i = 0 To UBound(a)
		a(i) = i And &HFF
	Next
	
	Seq = a
End Function

Class ClassRc4
	Sub Encrypt(key, msg)
		Dim s, i, j, k, t
		
		ReDim s(255)
		
		For i = 0 To UBound(s)
			s(i) = i
		Next
		
		j = 0
		
		For i = 0 To UBound(s)
			j = (j + s(i) + key(i And &H1F)) And &HFF
			t = s(i)
			s(i) = s(j)
			s(j) = t
		Next
		
		j = 0
		k = 0
		
		For i = 0 To UBound(msg)
			j = (j + 1) And &HFF
			k = (k + s(j)) And &HFF
			t = s(j)
			s(j) = s(k)
			s(k) = t
			msg(i) = msg(i) Xor s((s(j) + s(k)) And &HFF)
		Next
	End Sub
End Class

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub
