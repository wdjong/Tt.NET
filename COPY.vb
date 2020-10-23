Option Strict Off
Option Explicit On
Module modCopy
	'This program was created using BASIC in 1986 by Walter de Jong.
	'It was converted to C in 1989.
	'It was converted to VB3 (Win 3.x) in 1995
	'It was converted to VB4 (Win95/NT) in Dec 1997
	'April 2000: Locked version, consistent windows, modal
	'Things to do:
	
	'   tURN OFF TEST WHEN USING MENU OPTIONS.. PARTICULARLY WHERE MODAL FORM DISPLAYED
	
	'   password protected preferences
	'   preferences stored in registry
	'   check for lines longer than 65 characters
	'   indicate in database  'screen testlength and editing settings applying during test
	'   prevent changes to settings during test?
	'   scroll original during typing of copy
	
	'General guidelines
	'Windows resizable
	'Dialog boxes modal
	'Titlebar display false for dialog boxes 'make sure all variables are declared
	'
	'scope project
	Public tesLen As Short ' Length of test (seconds)
	Public tRemains As Short 'seconds remaining (timer1 FileStart)
	Public allowEdits As Boolean 'if allowing BS then true
	Public fil As String 'test file
	Public userName As String 'Name of user supplied in frmTitle

	'Scope module
	Dim linpos1 As Short 'Position in
	Dim wrd1 As String 'word from copy
	Dim eof1 As Boolean 'end of file flag for copy file
	Dim linpos2 As Short 'Position in
	Dim wrd2 As String 'word from original
	Dim errcount As Short 'error count with compare method
	Dim errcoun1 As Short
	Dim words As Double 'number of standard words
	Dim wpm As Double 'words per minute
	Dim acc As Double 'accuracy compare method
	Dim ac2 As Double 'accuracy spell check method
	
	Public Const szTitle As Short = 700 'size of title bar
	
	' MsgBox parameters
	Public Const MB_OK As Short = 0 ' OK button only
	Public Const MB_OKCANCEL As Short = 1 ' OK and Cancel buttons
	Public Const MB_ABORTRETRYIGNORE As Short = 2 ' Abort, Retry, and Ignore buttons
	Public Const MB_YESNOCANCEL As Short = 3 ' Yes, No, and Cancel buttons
	Public Const MB_YESNO As Short = 4 ' Yes and No buttons
	Public Const MB_RETRYCANCEL As Short = 5 ' Retry and Cancel buttons
	Public Const MB_ICONSTOP As Short = 16 ' Critical message
	Public Const MB_ICONQUESTION As Short = 32 ' Warning query
	Public Const MB_ICONEXCLAMATION As Short = 48 ' Warning message
	Public Const MB_ICONINFORMATION As Short = 64 ' Information message
	Public Const MB_APPLMODAL As Short = 0 ' Application Modal Message Box
	Public Const MB_DEFBUTTON1 As Short = 0 ' First button is default
	Public Const MB_DEFBUTTON2 As Short = 256 ' Second button is default
	Public Const MB_DEFBUTTON3 As Short = 512 ' Third button is default
	Public Const MB_SYSTEMMODAL As Short = 4096 'System Modal
	
	Sub CheckFiles() 'Spelling Check Method
		FrmResult.DefInstance.LstError.Items.Clear()
		FrmResult.DefInstance.LstError.Items.Add("--Spell Check--")
		eof1 = False 'go through copy again
		linpos1 = 1 'position in copy
		errcoun1 = 0 'reinitialise error count
		GetFile1Word() 'a word from the copy
		Do While Not eof1
			If InStr(FrmOrig.DefInstance.TxtOrig.Text, wrd1) = 0 Then 'word not found
				errcoun1 += errcoun1 + 1
				FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": not found")
			End If
			GetFile1Word()
		Loop 
	End Sub

	Sub ChooseCopy()
		Dim fd As OpenFileDialog = New OpenFileDialog()
		Dim fileOpenName As String
		Dim aLine As String 'a line of text from file

		fd.Title = "Choose copy"
		fd.InitialDirectory = VB6.GetPath 'look in same directory
		fd.Filter = "All files (*.*)|*.*"
		fd.FilterIndex = 1 '*.*
		fd.RestoreDirectory = True
		fd.FileName = "temp"

		If fd.ShowDialog() = DialogResult.OK Then
			fileOpenName = fd.FileName
			If fileOpenName <> "" Then
				On Error GoTo diskerr
				FileOpen(1, fileOpenName, OpenMode.Input)
				FrmTT.DefInstance.TxtCopy.Text = ""
				aLine = LineInput(1)
				tesLen = Val(Left(aLine, 1)) * 60 'backward compatibility
				If tesLen > 0 Then 'old version of test stored teslen at start of test
					aLine = Right(aLine, Len(aLine) - 1)
				Else
					tesLen = 5 * 60 'if it's not stored there then assume 5 minutes
				End If
				FrmTT.DefInstance.TxtCopy.Text = FrmTT.DefInstance.TxtCopy.Text & aLine & vbCrLf
				Do While Not EOF(1)
					aLine = LineInput(1)
					FrmTT.DefInstance.TxtCopy.Text = FrmTT.DefInstance.TxtCopy.Text & aLine & vbCrLf
				Loop
				FileClose(1)
				On Error GoTo 0
				If Right(FrmTT.DefInstance.TxtCopy.Text, 2) = vbCrLf Then
					FrmTT.DefInstance.TxtCopy.Text = Left(FrmTT.DefInstance.TxtCopy.Text, Len(FrmTT.DefInstance.TxtCopy.Text) - 2)
				End If
			End If
		End If
		GoTo endChoose
diskerr:
		MsgBox("Error accessing copy.")
		Resume Next
endChoose:
	End Sub

	Sub CompareFiles() 'Compare files method
		eof1 = False
		linpos1 = 1
		linpos2 = 1
		errcount = 0
		GetFile1Word() 'get a word from copy
		Do While Not eof1 And errcount < words / 2
			GetFile2Word() 'get a word from original
			If wrd1 <> wrd2 Then ErrorFound()
			GetFile1Word()
		Loop 
		If eof1 And wrd1 <> "" Then 'check last word
			GetFile2Word()
			If wrd1 <> Left(wrd2, Len(wrd1)) Then 'partial word not equal
				errcount += 1 'error found
				FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": (eof c) 1 word")
			End If
		End If
	End Sub
	
	Sub ErrorFound()
		Dim finish As Boolean
		Dim d As Short
		Dim e As Short
		Dim old1, old2 As String
		Dim tmp1, tmp2 As String
		
		finish = False
		Do While Not finish
			errcount += 1 'assume an error

			'An error of case alone (caps lock?)
			If LCase(wrd1) = LCase(wrd2) Then
				FrmResult.DefInstance.LstError.Items.Add(wrd1 & " should be " & wrd2)
				Exit Do
			End If
			
			'Is all but one letter ok?
			If Len(wrd1) > 1 And Len(wrd1) = Len(wrd2) Then
				e = 0
				For d = 1 To Len(wrd1)
					If Mid(wrd1, d, 1) = Mid(wrd2, d, 1) Then e += 1
				Next d
				If e = Len(wrd1) - 1 Then
					FrmResult.DefInstance.LstError.Items.Add(wrd1 & " should be " & wrd2)
					Exit Do
				End If
			End If
			
			'*********************** Read a word
			old1 = wrd1
			GetFile1Word() 'see below for poss. eof condition
			old2 = wrd2
			GetFile2Word()
			
			'Is the next word okay?
			If wrd1 = wrd2 Then
				FrmResult.DefInstance.LstError.Items.Add(old1 & " should be " & old2)
				Exit Do
			End If

			'Is next word of copy -> EOF (c.f. compare files for eof handling)
			If eof1 Then 'we've reached the end of the copy old1$ is wrong
				If wrd1 <> Left(wrd2, Len(wrd1)) Then 'wrd1$ is wrong
					errcount += 1
					FrmResult.DefInstance.LstError.Items.Add(old1 & ": (eof a) 1 word")
					FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": (eof a) 1 word")
				Else 'wrd1$ is right only old1$ is wrong
					FrmResult.DefInstance.LstError.Items.Add(old1 & ": (eof b) 1 word")
				End If
				Exit Do
			End If

			' /* space in wrong place? */
			tmp1 = old1 & wrd1
			tmp2 = old2 & wrd2
			If tmp1 = tmp2 Then
				FrmResult.DefInstance.LstError.Items.Add(old1 & " " & wrd1 & ": space in wrong place")
				Exit Do
			End If
			
			' /* space added */
			If old2 = tmp1 Then
				FrmResult.DefInstance.LstError.Items.Add(old1 & " " & wrd1 & ": extra space")
				GetFile1Word()
				If wrd1 <> wrd2 Then
					errcount += 1
					FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": ... and next wrong")
				End If
				Exit Do
			End If
			
			'/* space missing */
			If old1 = tmp2 Then
				GetFile2Word()
				FrmResult.DefInstance.LstError.Items.Add(old1 & ": space missing")
				If wrd1 <> wrd2 Then
					errcount += 1
					FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": ... and next wrong")
				End If
				Exit Do
			End If
			
			'/* are the words swapped? */
			If old1 = wrd2 And old2 = wrd1 Then
				FrmResult.DefInstance.LstError.Items.Add(old1 & " " & wrd1 & ": swapped")
				Exit Do
			End If
			
			'/* extra word? */
			If wrd1 = old2 Then
				FrmResult.DefInstance.LstError.Items.Add(old1 & ": extra word")
				GetFile1Word()
				If wrd1 <> wrd2 Then
					errcount += 1
					FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": ... and next wrong")
				End If
				Exit Do
			End If
			
			'/* missing word? */
			'old1: found  wrd1: his
			'old2: we     wrd2: found     wrd2: his
			If wrd2 = old1 Then
				FrmResult.DefInstance.LstError.Items.Add(old2 & ": word missing")
				GetFile2Word()
				If wrd1 <> wrd2 Then
					errcount += 1
					FrmResult.DefInstance.LstError.Items.Add(wrd1 & ": ... and next wrong")
				End If
				Exit Do
			End If
			
			FrmResult.DefInstance.LstError.Items.Add(old1 & ": unknown error (" & old2 & ")")
		Loop 
break: 
	End Sub
	
	Sub FindSpeed()
		words = Len(FrmTT.DefInstance.TxtCopy.Text) / 5 'standard word = 5 char
		wpm = Int(words / ((tesLen - tRemains) / 60) * 100) / 100
		FrmResult.DefInstance.TxSpeed.Text = CStr(wpm)
	End Sub
	
	Sub GetFile1Word()
		wrd1 = GetNext1()
		If linpos1 > Len(FrmTT.DefInstance.TxtCopy.Text) Then eof1 = True
	End Sub
	
	Sub GetFile2Word() 'straight from the original
		wrd2 = GetNext2() 'get the next typed word
		If linpos2 > Len(FrmOrig.DefInstance.TxtOrig.Text) Then 'we've reached the end of the original
			'eof2 = True 'we're currently at the last word of the original 'what's the point
			linpos2 = 1 'loop around - start comparing from the start of the original again
			wrd2 = GetNext2() ' get the first word
		End If
	End Sub

	Function GetNext2() As String
		Dim InWord As Boolean
		Dim c As String
		Dim wrd As String

		InWord = False
		If linpos2 > Len(FrmOrig.DefInstance.TxtOrig.Text) Then 'no more words
			GetNext2 = ""
			Exit Function
		Else
			c = Mid(FrmOrig.DefInstance.TxtOrig.Text, linpos2, 1)
		End If
		Do While Not InWord
			If c <= " " Then 'skip white space
				linpos2 += 1
				If linpos2 > Len(FrmOrig.DefInstance.TxtOrig.Text) Then
					GetNext2 = "" 'no more words
					Exit Function
				Else
					c = Mid(FrmOrig.DefInstance.TxtOrig.Text, linpos2, 1)
				End If
			Else
				InWord = True
			End If
		Loop
		wrd = ""
		Do While InWord
			wrd += c
			linpos2 += 1
			If linpos2 > Len(FrmOrig.DefInstance.TxtOrig.Text) Then
				GetNext2 = wrd
				Exit Function
			Else
				c = Mid(FrmOrig.DefInstance.TxtOrig.Text, linpos2, 1)
			End If
			If c <= " " Then 'white space indicates end of word
				InWord = False
			End If
		Loop
		GetNext2 = wrd
	End Function
	Function GetNext1() As String
		Dim InWord As Boolean
		Dim c As String
		Dim wrd As String

		wrd = "" 'clear the word
		InWord = False 'we are not yet in a word
		If linpos1 > Len(FrmTT.DefInstance.TxtCopy.Text) Then
			GetNext1 = "" 'when end reached return empty
			Exit Function
		End If
		c = Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1, 1) 'look at a char
		If linpos1 < Len(FrmTT.DefInstance.TxtCopy.Text) - 1 Then 'check for 2 spaces
			If Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1, 2) = "  " Then '2 spaces
				If linpos1 > 1 Then 'not at beginning of file
					If Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1 - 1, 1) > " " Then 'must follow a letter
						If Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1 - 1, 1) <> "." Then
							errcount += 1
							FrmResult.DefInstance.LstError.Items.Add(": > 1 space after word")
							linpos1 += 2
							c = Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1, 1)
						End If
					End If
				End If
			End If
		End If
		Do While Not InWord 'skip white space
			If c <= " " Then
				linpos1 += 1
				If linpos1 > Len(FrmTT.DefInstance.TxtCopy.Text) Then
					GetNext1 = "" 'return empty no word found
					Exit Function
				Else
					c = Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1, 1)
				End If
			Else
				InWord = True
			End If
		Loop
		Do While InWord
			wrd += c
			linpos1 += 1
			If linpos1 > Len(FrmTT.DefInstance.TxtCopy.Text) Then
				GetNext1 = wrd
				Exit Function
			Else
				c = Mid(FrmTT.DefInstance.TxtCopy.Text, linpos1, 1)
			End If
			If c <= " " Then
				InWord = False
			End If
		Loop
		GetNext1 = wrd
	End Function

	Public Function IsEditing(ByRef aChar As Short) As Boolean
		IsEditing = False
		Select Case aChar
			Case System.Windows.Forms.Keys.Back
				IsEditing = True
			Case System.Windows.Forms.Keys.PageUp '0x21    PAGE UP key.
				IsEditing = True
			Case System.Windows.Forms.Keys.PageDown '0x22    PAGE DOWN key.
				IsEditing = True
			Case System.Windows.Forms.Keys.End '0x23    END key.
				IsEditing = True
			Case System.Windows.Forms.Keys.Home '0x24    HOME key.
				IsEditing = True
			Case System.Windows.Forms.Keys.Left '0x25    LEFT ARROW key.
				IsEditing = True
			Case System.Windows.Forms.Keys.Up '0x26    UP ARROW key.
				IsEditing = True
			Case System.Windows.Forms.Keys.Right '0x27    RIGHT ARROW key.
				IsEditing = True
			Case System.Windows.Forms.Keys.Down '0x28    DOWN ARROW key.
				IsEditing = True
			Case System.Windows.Forms.Keys.Delete
				IsEditing = True
			Case System.Windows.Forms.Keys.Insert
				IsEditing = True
		End Select
	End Function

	Sub MarkTest()
		If FrmTT.DefInstance.TxtCopy.Text <> "" Then
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor 'Hourglass
			FindSpeed()
			FrmResult.DefInstance.LstError.Items.Clear()
			CompareFiles() 'Sequential method
			acc = Int(((words - errcount) / words * 100) * 100) / 100
			If (acc < 75) Then
				CheckFiles() 'Spelling check method
				ac2 = Int(((words - errcoun1) / words * 100) * 100) / 100
				If ac2 < 0 Then ac2 = 0
				FrmResult.DefInstance.TxAccuracy.Text = CStr(ac2)
			Else
				FrmResult.DefInstance.TxAccuracy.Text = CStr(acc)
			End If
			FrmResult.DefInstance.TxTime.Text = CStr(tesLen - tRemains)
			If allowEdits Then FrmResult.DefInstance.ChkAllow.CheckState = System.Windows.Forms.CheckState.Checked Else FrmResult.DefInstance.ChkAllow.CheckState = System.Windows.Forms.CheckState.Unchecked
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			FrmTT.DefInstance.TxtCopy.ReadOnly = True
			FrmResult.DefInstance.ShowDialog()
		End If
	End Sub

	Sub NewCopy()
		StopTime()
		tRemains = 0
		FrmTT.DefInstance.TxtCopy.ReadOnly = False 'locked after first test is finished
		FrmTT.DefInstance.TxtCopy.Text = "" 'clear copy
	End Sub

End Module