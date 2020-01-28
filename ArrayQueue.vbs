Option Explicit

Call Main()

Sub Main()
	Const size = 8
	
	Dim i
	For i = 0 To size
		Dim q
		Set q = New ArrayQueue

		Dim n
		n = 10
		
		Dim j
		
		For j = 1 To size
			Call q.Add(n)
			n = n + 1
		Next
		
		For j = 1 To i
			Call q.Remove()
		Next
		
		For j = 1 To i + 24
			Call q.Add(n)
			n = n + 1
		Next
		
		Call WriteLine(q.ToString())
	Next
End Sub

Class ArrayQueue
	Private xs
	Private n
	Private f
	Private l
	
	Public Property Get Count()
		Count = n
	End Property
	
	Private Sub Class_Initialize()
		ReDim xs(0)
		n = 0
		f = LBound(xs)
		l = UBound(xs)
	End Sub
	
	Public Sub Add(x)
		If UBound(xs) < n Then
			ReDim Preserve xs(n * 2 - 1)
			If l < f Then
				Dim i
				If l < n \ 2 Then
					For i = LBound(xs) To l
						xs(i + n) = xs(i)
						xs(i) = Empty
					Next
					l = l + n
				Else
					For i = n - 1 To f Step -1
						xs(i + n) = xs(i)
						xs(i) = Empty
					Next
					f = f + n
				End If
			End If
		End If
		l = (l + 1) And UBound(xs)
		xs(l) = x
		n = n + 1
	End Sub
	
	Public Function Remove()
		If 0 < n Then
			Remove = xs(f)
			xs(f) = Empty
			f = (f + 1) And UBound(xs)
			n = n - 1
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
