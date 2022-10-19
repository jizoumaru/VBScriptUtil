Option Explicit

Call Main()

Sub Main()
	Dim r
	Set r = New SecureRandom

	Dim i
	For i = 1 To 100
		Call WriteLine(r.NextI32())
	Next
End Sub

Class SecureRandom
	Private Rijndael
	Private Buffer
	Private Index
	
	Private Sub Class_Initialize()
		Set Rijndael = CreateObject("System.Security.Cryptography.RijndaelManaged")
		Buffer = ""
		Index = LenB(Buffer) + 1
	End Sub
	
	Public Function NextByte()
		If Index > LenB(Buffer) Then
			Call Rijndael.GenerateKey()
			Buffer = Rijndael.Key
			Index = 1
		End If
		NextByte = AscB(MidB(Buffer, Index, 1))
		Index = Index + 1
	End Function
	
	Public Function NextI32()
		NextI32 = I32(NextByte() _
			+ NextByte() * &H100 _
			+ NextByte() * &H10000 _
			+ NextByte() * &H1000000)
	End Function
	
	Private Function I32(n)
		If n > &H7FFFFFFF Then
			I32 = n - 4294967296
		Else
			I32 = n
		End If
	End Function
End Class

Sub WriteLine(s)
	Call WScript.StdOut.WriteLine(s)
End Sub
