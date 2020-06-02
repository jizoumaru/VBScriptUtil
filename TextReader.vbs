Option Explicit

Class TextReader
	Private Stream_
	Private Buf_
	Private Idx_
	Private BufSize_
	Public Current
	
	Public Sub Init(file, n)
		Const iomode_reading = 1
		Const create_no = False
		Const format_unicode = -1
		
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")

		Set Stream_ = fso.OpenTextFile(file, iomode_reading, create_no, format_unicode)
		BufSize_ = n
		Buf_ = ""
		Idx_ = 1
		Current = ""
	End Sub
	
	Public Function ReadAll()
		Dim a
		ReDim a(3)
		
		Dim i
		i = 0
		
		Do Until Stream_.AtEndOfStream
			If UBound(a) < i Then
				ReDim Preserve a(i + i - 1)
			End If
			a(i) = Stream_.Read(BufSize_)
			i = i + 1
		Loop
		
		ReDim Preserve a(i - 1)
		ReadAll = Join(a, "")
	End Function
	
	Public Function MoveNext()
		If Len(Buf_) < Idx_ Then
			If Stream_.AtEndOfStream Then
				MoveNext = False
				Exit Function
			End If
			Idx_ = 1
			Buf_ = Stream_.Read(BufSize_)
		End If
		
		Dim a
		a = ""
		
		Do
			Dim cr
			cr = InStr(Idx_, Buf_, vbCr)

			Dim lf
			lf = InStr(Idx_, Buf_, vbLf)
			
			If (0 < cr) Or (0 < lf) Then
				If (lf = 0) Or ((0 < cr) And (cr < lf)) Then
					Current = a & Mid(Buf_, Idx_, cr - Idx_)

					If cr < Len(Buf_) Then
						If cr + 1 = lf Then
							Idx_ = lf + 1
						Else
							Idx_ = cr + 1
						End If
					Else
						If Stream_.AtEndOfStream Then
							Idx_ = cr + 1
						Else
							Idx_ = 1
							Buf_ = Stream_.Read(BufSize_)
							If Mid(Buf_, Idx_, 1) = vbLf Then
								Idx_ = Idx_ + 1
							End If
						End If
					End If
				Else
					Current = a & Mid(Buf_, Idx_, lf - Idx_)
					Idx_ = lf + 1
				End If

				MoveNext = True
				Exit Function
			End If
			
			a = a & Mid(Buf_, Idx_, Len(Buf_) + 1 - Idx_)
			
			If Stream_.AtEndOfStream Then
				Idx_ = Len(Buf_) + 1
				Exit Do
			End If
			
			Idx_ = 1
			Buf_ = Stream_.Read(BufSize_)
		Loop
		
		Current = a
		MoveNext = True
	End Function
	
	Private Sub Class_Terminate()
		Call Stream_.Close()
		Set Stream_ = Nothing
	End Sub
End Class
