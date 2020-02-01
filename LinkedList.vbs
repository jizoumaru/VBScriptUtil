Option Explicit

Call Main()

Sub Main()
	Dim s
	Set s = New LinkedList
	
	Dim i
	For i = 1 To 10
		Call s.AddTail(i)
	Next

	Do While s.Count > 0
		Call WriteLine(s.RemoveTail())
	Loop
End Sub

Class LinkedList
	Private Ring
	Private Num
	
	Public Property Get Count()
		Count = Num
	End Property
	
	Private Sub Class_Initialize()
		Set Ring = New LinkedListNode
		Ring.Val = "root"
		Set Ring.Pr = Ring
		Set Ring.Nx = Ring
		Num = 0
	End Sub
	
	Private Sub Class_Terminate()
		Dim i
		For i = 1 To Num
			Set Ring.Nx = Ring.Nx.Nx
			Set Ring.Nx.Pr = Ring
		Next
		Ring.Nx = Empty
		Ring.Pr = Empty
		Ring = Empty
	End Sub
	
	Public Sub AddHead(val)
		Dim n
		Set n = New LinkedListNode
		n.Val = val
		Set n.Pr = Ring
		Set n.Nx = Ring.Nx
		Set Ring.Nx.Pr = n
		Set Ring.Nx = n
		Num = Num + 1
	End Sub
	
	Public Function RemoveHead()
		If Num = 0 Then
			RemoveHead = Empty
		Else
			RemoveHead = Ring.Nx.Val
			Set Ring.Nx = Ring.Nx.Nx
			Set Ring.Nx.Pr = Ring
			Num = Num - 1
		End If
	End Function
	
	Public Sub AddTail(val)
		Dim n
		Set n = New LinkedListNode
		n.Val = val
		Set n.Pr = Ring.Pr
		Set n.Nx = Ring
		Set Ring.Pr.Nx = n
		Set Ring.Pr = n
		Num = Num + 1
	End Sub
	
	Public Function RemoveTail()
		If Num = 0 Then
			RemoveTail = Empty
		Else
			RemoveTail = Ring.Pr.Val
			Set Ring.Pr = Ring.Pr.Pr
			Set Ring.Pr.Nx = Ring
			Num = Num - 1
		End If
	End Function
End Class

Class LinkedListNode
	Public Val
	Public Pr
	Public Nx
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
