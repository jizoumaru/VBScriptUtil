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
	Private Head
	Private Tail
	Private Num
	
	Public Property Get Count()
		Count = Num
	End Property
	
	Private Sub Class_Initialize()
		Num = 0
	End Sub
	
	Private Sub Class_Terminate()
		Dim i
		For i = 1 To Num
			Call Remove()
		Next
	End Sub
	
	Public Sub Add(val)
		Dim n
		Set n = New LinkedQueueNode
		n.Val = val
		
		If Num = 0 Then
			Set Head = n
			Set Tail = n
		Else
			Set Tail.Nx = n
			Set Tail = n
		End If
		Num = Num + 1
	End Sub
	
	Public Function Remove()
		If 0 < Num Then
			Remove = Head.Val
			Num = Num - 1
			If Num = 0 Then
				Head = Empty
				Tail = Empty
			Else
				Set Head = Head.Nx
			End If
		Else
			Remove = Empty
		End If
	End Function
End Class

Class LinkedQueueNode
	Public Val
	Public Nx
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
