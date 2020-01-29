Option Explicit

Call Main()

Sub Main()
	Dim s
	Set s = New LinkedStack

	Dim i
	For i = 1 To 10
		Call s.Add(i)
	Next
	
	Do While s.Count > 0
		Call WriteLine(s.Remove())
	Loop
End Sub

Class LinkedStack
	Private Tail
	Private Num
	
	Public Property Get Count()
		Count = Num
	End Property
	
	Private Sub Class_Initialize()
		Set Tail = Nothing
		Num = 0
	End Sub
	
	Private Sub Class_Terminate()
		Dim i
		For i = 1 To Num
			Call Remove()
		Next
	End Sub
	
	Public Sub Add(val)
		Set Tail = Node(val, Tail)
		Num = Num + 1
	End Sub
	
	Public Function Remove()
		If 0 < Num Then
			Remove = Tail.Val
			Set Tail = Tail.Nx
			Num = Num - 1
		Else
			Remove = Empty
		End If
	End Function
	
	Private Function Node(val, nx)
		Dim n
		Set n = New LinkedStackNode
		n.Val = val
		Set n.Nx = nx
		Set Node = n
	End Function
End Class

Class LinkedStackNode
	Public Val
	Public Nx
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
