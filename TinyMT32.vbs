Option Explicit

Const U32_MAX = 4294967296
Const I32_MAX = 2147483647
Const I32_MIN = -2147483648

Call Main()

Sub Main()
	Dim mt
	Set mt = New TinyMT32
	Call mt.Init(1)
	
	Dim i
	For i = 1 To 50
		Call WriteLine(i & ":" & U32(mt.Generate()))
	Next
End Sub

Class TinyMT32
	Private param_
	Private status_
	
	Public Sub Init(seed)
		Dim param
		param = Array(seed, &H8f7011ee, &Hfc78ff1f, &H3793fdff)
	
		Dim status
		status = param
		
		Dim i
		
		For i = 1 To 7
			status(i And 3) = status(i And 3) Xor I32(i + Mul(1812433253, _
					status((i - 1) And 3) Xor Shr(status((i - 1) And 3), 30)))
		Next
		
		For i = 1 To 8
			status = NextState(status, param)
		Next
		
		param_ = param
		status_ = status
	End Sub

	Public Function Generate()
		status_ = NextState(status_, param_)
		Generate = Termper(status_, param_)
	End Function
	
	Private Function NextState(status, param)
		Dim x
		x = (status(0) And &H7FFFFFFF) Xor status(1) Xor status(2)
		x = x Xor Shl(x, 1)
		
		Dim y
		y = status(3)
		y = y Xor Shr(y, 1) Xor x
		
		NextState = Array( _
			status(1), _
			status(2) Xor (-(y And 1) And param(1)), _
			x Xor Shl(y, 10) Xor (-(y And 1) And param(2)), _
			y)
	End Function
	
	Private Function Termper(status, param)
		Dim a
		a = I32(status(0) + Shr(status(2), 8))
		Termper = status(3) Xor a Xor (-(a And 1) And param(3))
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
	
	Mul = I32(Remainder(r, U32_MAX))
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
	Shl = I32((val And ((2 ^ (32 - num)) - 1)) * (2 ^ num))
End Function

Sub WriteLine(a)
	Call WScript.StdOut.WriteLine(a)
End Sub
