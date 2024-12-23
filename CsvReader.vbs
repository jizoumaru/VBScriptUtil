Option Explicit

Call Main()

Sub Main()
	Dim Reader
	Set Reader = New CsvReaderClass
	Call Reader.OpenFile("sample.txt")

	Do Until Reader.AtEndOfStream
		Call WriteLine(FormatRecord(Reader.Read()))
	Loop
End Sub

Sub WriteLine(value)
	Call WScript.StdOut.WriteLine(value)
End Sub

Sub CsvReader_Test()
	Call AssertEquals("", CsvReader_Read(""))
	
	Call AssertEquals("()",  CsvReader_Read("" & """"))
	Call AssertEquals("(|)", CsvReader_Read("" & ","))
	Call AssertEquals("()",  CsvReader_Read("" & vbCr))
	Call AssertEquals("()",  CsvReader_Read("" & vbLf))
	Call AssertEquals("(A)", CsvReader_Read("" & "A"))
	
	Call AssertEquals("(|)",  CsvReader_Read("," & """"))
	Call AssertEquals("(||)", CsvReader_Read("," & ","))
	Call AssertEquals("(|)",  CsvReader_Read("," & vbCr))
	Call AssertEquals("(|)",  CsvReader_Read("," & vbLf))
	Call AssertEquals("(|A)", CsvReader_Read("," & "A"))
	
	Call AssertEquals("(Aq)", CsvReader_Read("A" & """"))
	Call AssertEquals("(A|)", CsvReader_Read("A" & ","))
	Call AssertEquals("(A)",  CsvReader_Read("A" & vbCr))
	Call AssertEquals("(A)",  CsvReader_Read("A" & vbLf))
	Call AssertEquals("(AA)", CsvReader_Read("A" & "A"))
	
	Call AssertEquals("()",  CsvReader_Read("""" & """"))
	Call AssertEquals("(,)", CsvReader_Read("""" & ","))
	Call AssertEquals("(r)", CsvReader_Read("""" & vbCr))
	Call AssertEquals("(n)", CsvReader_Read("""" & vbLf))
	Call AssertEquals("(A)", CsvReader_Read("""" & "A"))
	
	Call AssertEquals("()()",  CsvReader_Read(vbCr & """"))
	Call AssertEquals("()(|)", CsvReader_Read(vbCr & ","))
	Call AssertEquals("()()",  CsvReader_Read(vbCr & vbCr))
	Call AssertEquals("()",    CsvReader_Read(vbCr & vbLf))
	Call AssertEquals("()(A)", CsvReader_Read(vbCr & "A"))
End Sub

Function CsvReader_Read(Str)
	Dim StringReader
	Set StringReader = New StringReaderClass
	Call StringReader.Init(Str)
	
	Dim CsvReader
	Set CsvReader = New CsvReaderClass
	Call CsvReader.Init(StringReader)
	
	Dim Result
	Result = Array()
	
	Do Until CsvReader.AtEndOfStream
		ReDim Preserve Result(UBound(Result) + 1)
		Result(UBound(Result)) = FormatRecord(CsvReader.Read())
	Loop
	
	CsvReader_Read = Join(Result, "")
End Function

Function FormatRecord(Arr)
	Dim I
	For I = 0 To UBound(Arr)
		Arr(I) = Replace(Arr(I), """", "q")
		Arr(I) = Replace(Arr(I), vbCr, "r")
		Arr(I) = Replace(Arr(I), vbLf, "n")
	Next
	
	FormatRecord = "(" & Join(Arr, "|") & ")"
End Function

Sub StringReader_Test()
	Dim Reader
	Set Reader = New StringReaderClass

	Call Reader.Init("")
	Call AssertEquals(True, Reader.AtEndOfStream)
	
	Call Reader.Init("a")
	Call AssertEquals(False, Reader.AtEndOfStream)
	Call AssertEquals("a", Reader.Read(1))
	Call AssertEquals(True, Reader.AtEndOfStream)

	Call Reader.Init("a")
	Call AssertEquals("a", Reader.Read(2))
	Call AssertEquals(True, Reader.AtEndOfStream)

	Call Reader.Init("abcde")
	Call AssertEquals("ab", Reader.Read(2))
	Call AssertEquals("cd", Reader.Read(2))
	Call AssertEquals("e", Reader.Read(2))
	Call AssertEquals(True, Reader.AtEndOfStream)
End Sub

Sub AssertEquals(ExpectedValue, ActualValue)
	If ExpectedValue = ActualValue Then
		Call WriteLine("OK: " & ExpectedValue)
	Else
		Call WriteLine("ERROR: " & ExpectedValue & ", " & ActualValue)
	End If
End Sub

Class StringReaderClass
	Private Value
	Private Index
	
	Public Sub Init(InValue)
		Value = InValue
		Index = 1
	End Sub
	
	Public Property Get AtEndOfStream()
		AtEndOfStream = Len(Value) < Index
	End Property
	
	Public Function Read(Count)
		If Count < Len(Value) - Index + 1 Then
			Read = Mid(Value, Index, Count)
			Index = Index + Count
		Else
			Read = Mid(Value, Index)
			Index = Len(Value) + 1
		End If
	End Function
End Class

Class CsvReaderClass
	Private Stream
	Private Buffer
	Private Offset

	Public Sub OpenFile(File)
		Dim FileSystem
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		Call Init(FileSystem.OpenTextFile(File))
	End Sub
	
	Public Sub Init(InStream)
		Set Stream = InStream
		Offset = 1
		Buffer = ""
	End Sub
	
	Public Property Get AtEndOfStream()
		If Len(Buffer) < Offset Then
			AtEndOfStream = Stream.AtEndOfStream
		Else
			AtEndOfStream = False
		End If
	End Property
	
	Private Sub AddValue(Record, Buffer, FromIndex, ToIndex)
		ReDim Preserve Record(UBound(Record) + 1)

		If Mid(Buffer, FromIndex, 1) <> """" Then
			Record(UBound(Record)) = Mid(Buffer, FromIndex, ToIndex - FromIndex)
			Exit Sub
		End If
		
		Dim Escaped
		Escaped = True
		
		Dim I
		For I = FromIndex + 1 To ToIndex - 1
			If Mid(Buffer, I, 1) = """" Then
				Escaped = Not Escaped
			Else
				If Not Escaped Then
					Exit For
				End If
			End If
		Next
		
		If Escaped Then
			Record(UBound(Record)) = Replace(Mid(Buffer, FromIndex + 1, I - FromIndex - 1), """""", """")
		Else
			Record(UBound(Record)) = Replace(Mid(Buffer, FromIndex + 1, I - 1 - FromIndex - 1), """""", """") & Mid(Buffer, I, ToIndex - I)
		End If
	End Sub
	
	Public Function Read()
		Const START = 0
		Const FIELD = 1
		Const ESCAPED = 2
		Const PLAIN = 3
		Const CR = 4

		Read = Array()
		
		Dim Index
		Index = Offset
		
		Dim State
		State = START
		
		Do
			For Index = Index To Len(Buffer)
				Dim C
				C = Mid(Buffer, Index, 1)
			
				If State = CR Then
					If C = vbLf Then
						Call AddValue(Read, Buffer, Offset, Index - 1)
						Offset = Index + 1
						Exit Function
					Else
						Call AddValue(Read, Buffer, Offset, Index - 1)
						Offset = Index
						Exit Function
					End If
				ElseIf State = ESCAPED Then
					If C = """" Then
						State = FIELD
					End If
				Else
					If C = """" Then
						If State <> PLAIN Then
							State = ESCAPED
						End If
					ElseIf C = "," Then
						Call AddValue(Read, Buffer, Offset, Index)
						Offset = Index + 1
						State = FIELD
					ElseIf C = vbCr Then
						State = CR
					ElseIf C = vbLf Then
						Call AddValue(Read, Buffer, Offset, Index)
						Offset = Index + 1
						Exit Function
					Else
						State = PLAIN
					End If
				End If
			Next

			If Stream.AtEndOfStream Then
				If State = CR Then
					Call AddValue(Read, Buffer, Offset, Index - 1)
				ElseIf State <> START Then
					Call AddValue(Read, Buffer, Offset, Index)
				End If
				
				Buffer = ""
				Exit Function
			End If
			
			Buffer = Mid(Buffer, Offset) & Stream.Read(1)
			Index = Index - Offset + 1
			Offset = Offset - Offset + 1
		Loop
	End Function
End Class
