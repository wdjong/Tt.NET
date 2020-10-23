Option Strict Off
Option Explicit On
Module modOrig
	
	Public Sub ChooseOriginal()
		Dim aLine As String 'a line of text from file
		Dim oldFil As String

		Dim fd As OpenFileDialog = New OpenFileDialog()
		Dim strFileName As String

		fd.Title = "Choose original"
		fd.InitialDirectory = VB6.GetPath 'look in same directory
		fd.Filter = "Test (test*.*)|test*.*|All files (*.*)|*.*"
		fd.FilterIndex = 1 'test*.*
		fd.RestoreDirectory = True

		If fd.ShowDialog() = DialogResult.OK Then
			oldFil = fil 'save
			fil = fd.FileName
			FrmOrig.DefInstance.Text = "Original: " & fil
			On Error GoTo diskerr
			FileOpen(2, fil, OpenMode.Input)
			FrmOrig.DefInstance.TxtOrig.Text = ""
			Do While Not EOF(2)
				aLine = LineInput(2)
				FrmOrig.DefInstance.TxtOrig.Text = FrmOrig.DefInstance.TxtOrig.Text & aLine & Chr(13) & Chr(10)
			Loop
			On Error GoTo 0
			FrmOrig.DefInstance.Show()
			FrmTT.DefInstance.TxtCopy.Focus()
		End If
		GoTo endChoose

diskerr: 
		MsgBox("Error accessing original.")
		fil = ""
		Resume endChoose
		
endChoose: 
		FileClose(2)
		On Error GoTo 0
		
	End Sub
End Module