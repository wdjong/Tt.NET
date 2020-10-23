Option Strict Off
Option Explicit On
Module modHistory
	
	Sub AddHistoryRecord()
        Dim testLength As Double
        Dim testPos As Short
        Dim myDataRow As DataRow

        Try
            myDataRow = FrmHistory.DefInstance.DataSet1.Tables(0).NewRow()
            myDataRow("TestDate") = Now
            myDataRow("Speed") = Val(FrmResult.DefInstance.TxSpeed.Text)
            myDataRow("Accuracy") = Val(FrmResult.DefInstance.TxAccuracy.Text)
            myDataRow("TestName") = My.Settings.testee 'current user
            testLength = Int((tesLen - tRemains) / 60 * 100) / 100 'in minutes
            myDataRow("TestLen") = testLength
            testPos = InStr(LCase(fil), "\test") 'finding test number from filename
            If testPos > 0 Then
                myDataRow("TestNum") = Mid(fil, testPos + 1, Len(fil) - testPos + 2)
            End If
            FrmHistory.DefInstance.DataSet1.Tables(0).Rows.Add(myDataRow)
            FrmHistory.DefInstance.Result1TableAdapter.Update(FrmHistory.DefInstance.DataSet1.Tables(0))
        Catch ex As Exception
            MsgBox("AddHistoryRecord: Writing to database: " & ex.Message)
        End Try
	End Sub
End Module