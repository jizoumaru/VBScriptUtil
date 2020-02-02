Option Explicit

Call Main()

Sub Main()
	Const n = 10
	
	Dim h
	Set h = New HashTable
	Call h.Init(n)
	
	Dim i
	For i = 1 To n
		h("key" & i) = "value" & i
	Next
	
	For i = 1 To n
		Call WriteLine(h("key" & i))
	Next
End Sub

Class HashTable
	Private Table
	Private Num
	
	Public Sub Init(size)
		ReDim Table(size * 2 - 1)
		Num = 0
	End Sub

	Public Default Property Get Item(x)
		Call Find(Table(Hash(x)), x, Item)
	End Property
	
	Public Property Let Item(x, y)
		Call Add(Table(Hash(x)), x, y)
	End Property
	
	Private Sub Add(bucket, x, y)
		If IsEmpty(bucket) Then
			bucket = Array(x, y)
			Num = Num + 1
		Else
			Dim i
			For i = 0 To UBound(bucket) Step 2
				If bucket(i) = x Then
					bucket(i + 1) = y
					Exit Sub
				End If
			Next
			
			Dim n
			n = UBound(bucket)
			
			ReDim Preserve bucket(n + 2)
			bucket(n + 1) = x
			bucket(n + 2) = y
			
			Num = Num + 1
		End If
	End Sub

	Private Sub Find(bucket, x, y)
		If IsArray(bucket) Then
			Dim i
			For i = 0 To UBound(bucket) Step 2
				If bucket(i) = x Then
					y = bucket(i + 1)
					Exit Sub
				End If
			Next
		End If
		y = Empty
	End Sub
	
	Private Function Hash(val)
		Dim h
		If VarType(val) = vbString Then
			h = 5381
			
			Dim i
			For i = 1 To Len(val)
				h = (&H1FFFFFF And h) * 33 + AscW(Mid(val, i, 1))
			Next
		Else
			h = val
		End If
		
		Dim n
		n = (UBound(Table) + 1) \ 2
		
		h = h Mod n
		
		If h < 0 Then
			h = h + n
		End If
		
		Hash = h
	End Function
End Class

Sub WriteLine(val)
	Call WScript.StdOut.WriteLine(val)
End Sub
