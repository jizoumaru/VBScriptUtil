Option Explicit

Const U32_MAX = 4294967296
Const I32_MAX = 2147483647
Const I32_MIN = -2147483648

Call Main()

Sub Main()
	Dim xsh
	Set xsh = New Xoshiro128ss
	Call xsh.Init(1234, 5678, 91012, 3456)
	
	Dim i
	For i = 1 To 1000000
		Dim n
		n = xsh.Generate()
		If i Mod 10000 = 0 Then
			Call WriteLine(n)
		End If
	Next
End Sub

Class Xoshiro128ss
	Private JUMP_
	Private status_
	
	Public Sub Init(a, b, c, d)
		JUMP_ = Array(&H8764000b&, &Hf542d2d3&, &H6fa035c3&, &H77f2db5b&)
		status_ = Array(a, b, c, d)
		Call Jump()
	End Sub
	
	Public Function Generate()
		Generate = Add(Rotl(Add(status_(0), status_(3)), 7), status_(0))
		Dim t
		t = Shl(status_(1), 9)
		status_(2) = status_(2) Xor status_(0)
		status_(3) = status_(3) Xor status_(1)
		status_(1) = status_(1) Xor status_(2)
		status_(0) = status_(0) Xor status_(3)
		status_(2) = status_(2) Xor t
		status_(3) = Rotl(status_(3), 11)
	End Function
	
	Private Function Rotl(x, k)
		Rotl = Shl(x, k) Or Shr(x, (32 - k))
	End Function
	
	Public Sub Jump()
		Dim s0
		s0 = 0
		
		Dim s1
		s1 = 0
		
		Dim s2
		s2 = 0
		
		Dim s3
		s3 = 0

		Dim i
		For i = 0 To UBound(JUMP_)
			Dim b
			For b = 0 To 31
				If (JUMP_(i) And Shl(1, b)) <> 0 Then
					s0 = s0 Xor status_(0)
					s1 = s1 Xor status_(1)
					s2 = s2 Xor status_(2)
					s3 = s3 Xor status_(3)
				End If
				
				Call Generate()
			Next
		Next

		status_(0) = s0
		status_(1) = s1
		status_(2) = s2
		status_(3) = s3
	End Sub
End Class

Function Add(x, y)
	If x + y > I32_MAX Then
		Add = x + y - U32_MAX
	ElseIf x + y < I32_MIN Then
		Add = x + y + U32_MAX
	Else
		Add = x + y
	End If
End Function

Function I32(val)
	If I32_MAX < val Then
		I32 = val - U32_MAX
	ElseIf val < I32_MIN Then
		I32 = val + U32_MAX
	Else
		I32 = val
	End If
End Function

Function U32(val)
	If val < 0 Then
		U32 = val + U32_MAX
	Else
		U32 = val
	End If
End Function

Function Remainder(a, b)
	Remainder = a - Fix(a / b) * b
End Function

Function Shr(val, num)
	Shr = Fix(U32(val) / (2 ^ num))
End Function

Function Shl(val, num)
	If num = 0 Then
		Shl = val
	Else
		Shl = I32((val And ((2 ^ (32 - num)) - 1)) * (2 ^ num))
	End If
End Function

Sub WriteLine(a)
	Call WScript.StdOut.WriteLine(a)
End Sub
