Option Explicit

Const I32_MAX = 2147483647
Const U32_MAX = 4294967296

Call Main()

Sub Main()
	Dim xs
	Set xs = New Xorshift
	
	Dim i
	For i = 1 To 10
		Call WriteLine(i & ":" & U32(xs.Generate()))
	Next
End Sub

Class Xorshift
	Private n
	
	Private Sub Class_Initialize()
		n = &H92D68CA2
	End Sub
	
	Public Function Generate()
		n = n Xor Shl(n, 13)
		n = n Xor Shr(n, 17)
		n = n Xor Shl(n, 5)
		Generate = n
	End Function
End Class

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

Function Shr(val, num)
	Shr = Fix(U32(val) / (2 ^ num))
End Function

Function Shl(val, num)
	Shl = I32((val And ((2 ^ (32 - num)) - 1)) * (2 ^ num))
End Function

Sub WriteLine(a)
	Call WScript.StdOut.WriteLine(a)
End Sub
