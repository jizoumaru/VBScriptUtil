Option Explicit

Call Main()

Sub Main()
	Const size = 8
	
	Dim s
	Set s = New ArrayStack

	Dim i
	For i = 1 To size
		Call s.Add(i)
	Next
	
	Do While s.Count > 0
		Call WriteLine(s.Remove() & ":" & s.ToString())
	Loop
End Sub

Class ArrayStack
	Private xs
	Private n
	
	Public Property Get Count()
		Count = n
	End Property
	
	Private Sub Class_Initialize()
		ReDim xs(0)
		n = 0
	End Sub
	
	Public Sub Add(x)
		If UBound(xs) < n Then
			ReDim Preserve xs(n * 2 - 1)
		End If
		xs(n) = x
		n = n + 1
	End Sub
	
	Public Function Remove()
		If 0 < n Then
			n = n - 1
			Remove = xs(n)
			xs(n) = Empty
		Else
			Remove = Empty
		End If
	End Function
	
	Public Function ToString()
		ToString = "(" & Join(xs, ", ") & ")"
	End Function
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
