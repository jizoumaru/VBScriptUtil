Option Explicit

Const U32_MAX = 4294967296
Const I32_MAX = 2147483647
Const I32_MIN = -2147483648

Call Main()

Sub Main()
	Dim xa
	Set xa = New XorshiftAdd
	Call xa.Init(1234)
	
	Dim i
	For i = 1 To 1000000
		Call WriteLine(i & ":" & U32(xa.Generate()))
	Next
End Sub

Class XorshiftAdd
	Private status_
	
	Public Sub Init(seed)
		Dim status
		status = Array(seed, 0, 0, 0)
		
		Dim i
		
		For i = 1 To 7
			status(i And 3) = status(i And 3) Xor Add(i, Mul(1812433253, _
					status((i - 1) And 3) Xor Shr(status((i - 1) And 3), 30)))
		Next
		
		For i = 1 To 8
			status = NextState(status)
		Next
		
		status_ = status
	End Sub

	Public Function Generate()
		status_ = NextState(status_)
		Generate = Add(status_(3), status_(2))
	End Function
	
	Private Function NextState(status)
		Dim t
		t = status(0)
		t = t Xor Shl(t, 15)
		t = t Xor Shr(t, 18)
		t = t Xor Shl(status(3), 11)
		NextState = Array(status(1), status(2), status(3), t)
	End Function
End Class

Function Mul(ByVal x, ByVal y)
	Dim r
	r = 0
	
	Do
		If x = 0 Then
			Exit Do
		End If
		
		If y = 0 Then
			Exit Do
		End If
		
		If (y And 1) <> 0 Then
			r = r + x
		End If
		
		x = Shl(x, 1)
		y = Shr(y, 1)
	Loop
	
	If r < I32_MIN Then
		r = U32_MAX - 1 - Remainder(-1 - r, U32_MAX)
	ElseIf I32_MAX < r Then
		r = Remainder(r, U32_MAX)
	End If
	
	Mul = I32(r)
End Function

Function Add(a, b)
	Dim c
	c = a + b
	
	If c < I32_MIN Then
		Add = c + U32_MAX
	ElseIf I32_MAX < c Then
		Add = c - U32_MAX
	Else
		Add = c
	End If
End Function

Function I32(val)
	If I32_MAX < val Then
		I32 = val - U32_MAX
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
	Shl = I32((val And ((2 ^ (32 - num)) - 1)) * (2 ^ num))
End Function

Sub WriteLine(a)
	Call WScript.StdOut.WriteLine(a)
End Sub
