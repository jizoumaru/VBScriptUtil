Option Explicit

Call Main()

Sub Main()
	Dim s
	Set s = New ArrayList

	Dim i
	For i = 1 To 15
		Call s.Add(Empty)
	Next
	
	For i = 0 To s.Count - 1
		s(i) = i
	Next
	
	For i = 0 To s.Count - 1
		Call WriteLine(s(i))
	Next
End Sub

Class ArrayList
	Private xs
	Private n
	
	Public Property Let Item(i, x)
		xs(i) = x
	End Property
	
	Public Default Property Get Item(i)
		Item = xs(i)
	End Property
	
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
	
	Public Function ToString()
		ToString = "(" & Join(xs, ", ") & ")"
	End Function
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
