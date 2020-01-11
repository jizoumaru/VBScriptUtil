Option Explicit

Const U32_MAX = 4294967296
Const I32_MAX = 2147483647
Const I32_MIN = -2147483648

Call Main()

Sub Main()
	Dim mt
	Set mt = New MT19937
	Call mt.InitKeys(Array(&H123, &H234, &H345, &H456))
	
	Dim i
	For i = 1 To 10
		Call WriteLine(i & ":" & U32(mt.Generate()))
	Next
End Sub

Class MT19937
	Private N
	Private M
	Private MATRIX_A
	Private mt
	Private mti
	Private mag01

	Private Sub Class_Initialize()
		N = 624
		M = 397
		MATRIX_A = &H9908B0DF
		ReDim mt(N - 1)
		mti = 0
		mag01 = Array(0, MATRIX_A)
	End Sub

	Public Sub InitSeed(seed)
		mt(0) = seed
		
		Dim i
		For i = 1 To N - 1
			mt(i) = Add(Mul(1812433253, mt(i - 1) Xor Shr(mt(i - 1), 30)), i)
		Next
		
		mti = N
	End Sub

	Public Sub InitKeys(keys)
		Call InitSeed(19650218)
		
		Dim i
		i = 1
		
		Dim j
		j = 0
		
		Dim k
		
		For k = 1 To N
			mt(i) = Add(Add((mt(i) Xor Mul(mt(i - 1) Xor Shr(mt(i - 1), 30), 1664525)), keys(j)), j)
			i = i + 1
			j = j + 1
			
			If i >= N Then
				mt(0) = mt(N - 1)
				i = 1
			End If
		
			If j > UBound(keys) Then
				j = 0
			End If
		Next
		
		For k = 0 To N - 1 - 1
			mt(i) = Subtract(mt(i) Xor Mul(mt(i - 1) Xor Shr(mt(i - 1), 30), 1566083941), i)
			i = i + 1
			
			If i >= N Then
				mt(0) = mt(N - 1)
				i = 1
			End If
		Next

		mt(0) = &H80000000
	End Sub

	Public Function Generate()
		If mti >= N Then
			Dim i
			Dim j
			Dim k
			
			j = M
			For i = 0 To N - M - 1
				k = (mt(i) And &H80000000) Or (mt(i + 1) And &H7FFFFFFF)
				mt(i) = mt(j) Xor Shr(k, 1) Xor mag01(k And 1)
				j = j + 1
			Next
			
			j = 0
			For i = N - M To N - 1 - 1
				k = (mt(i) And &H80000000) Or (mt(i + 1) And &H7FFFFFFF)
				mt(i) = mt(j) Xor Shr(k, 1) Xor mag01(k And 1)
				j = j + 1
			Next
			
			k = (mt(N - 1) And &H80000000) Or (mt(0) And &H7FFFFFFF)
			mt(N - 1) = mt(M - 1) Xor Shr(k, 1) Xor mag01(k And 1)
			mti = 0
		End If

		Dim y
		y = mt(mti)
		y = y Xor Shr(y, 11)
		y = y Xor (Shl(y, 7) And &H9D2C5680)
		y = y Xor (Shl(y, 15) And &HEFC60000)
		y = y Xor Shr(y, 18)
		
		mti = mti + 1
		Generate = y
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

Function Subtract(a, b)
	Dim c
	c = a - b
	
	If c < I32_MIN Then
		Subtract = c + U32_MAX
	ElseIf I32_MAX < c Then
		Subtract = c - U32_MAX
	Else
		Subtract = c
	End If
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
