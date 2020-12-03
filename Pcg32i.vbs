Option Explicit

Main

Sub Main()
	Dim pcg
	Set pcg = New Pcg32i
	Call pcg.Init(1306043654 Xor CLng(Timer() * 1000), 1599302943 Xor CLng(Timer() * 1000))
	
	Dim i
	For i = 1 To 10
		Call WScript.StdOut.WriteLine(pcg.Generate())
	Next
End Sub

Class Pcg32i
	Private State
	Private Inc
	
	Sub Init(State_, Inc_)
		State = 0
		Inc = Shl(State_, 1) Or 1
		Call Generate()
		State = I32(State + State_)
		Call Generate()
	End Sub
	
	Function Generate()
		Dim old
		old = State
		
		State = I32(Mul(State, 747796405) + Inc)
		
		Dim word
		word = Mul(Shr(old, I32(Shr(old, 28) + 4)) Xor old, 277803737)
		
		Generate = Shr(word, 22) Xor word
	End Function
End Class

Const U32_MAX = 4294967296
Const I32_MAX = 2147483647
Const I32_MIN = -2147483648

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
