Option Explicit

Call Main()

Sub Main()
	Dim s
	Set s = New LinkedQueue

	Dim i
	For i = 1 To 10
		Call s.Add(i)
	Next
	
	Do While s.Count > 0
		Call WriteLine(s.Remove())
	Loop
End Sub

Class LinkedQueue
	Private Ring
	Private Num
	
	Public Property Get Count()
		Count = Num
	End Property
	
	Private Sub Class_Initialize()
		Set Ring = Node(Empty, Nothing, Nothing)
		Set Ring.Pr = Ring
		Set Ring.Nx = Ring
		Num = 0
	End Sub
	
	Private Sub Class_Terminate()
		Dim i
		For i = 1 To Num
			Call Remove()
		Next
		Ring.Nx = Empty
		Ring.Pr = Empty
		Ring = Empty
	End Sub
	
	Public Sub Add(val)
		Set Ring.Pr.Nx = Node(val, Ring.Pr, Ring)
		Set Ring.Pr = Ring.Pr.Nx
		Num = Num + 1
	End Sub
	
	Public Function Remove()
		If 0 < Num Then
			Remove = Ring.Nx.Val
			Set Ring.Nx = Ring.Nx.Nx
			Set Ring.Nx.Pr = Ring
			Num = Num - 1
		Else
			Remove = Empty
		End If
	End Function
	
	Private Function Node(val, pr, nx)
		Dim n
		Set n = New LinkedQueueNode
		n.Val = val
		Set n.Pr = pr
		Set n.Nx = nx
		Set Node = n
	End Function
End Class

Class LinkedQueueNode
	Public Val
	Public Pr
	Public Nx
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
