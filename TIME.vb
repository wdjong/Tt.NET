Option Strict Off
Option Explicit On
Module modTime
	Sub startTime()
		FrmTime.DefInstance.Show() 'show guage
		tRemains = tesLen 'time = current test length in secs
        FrmTime.DefInstance.Gauge1.Maximum = tesLen 'setup up guage
		FrmTime.DefInstance.Gauge1.Value = tRemains
		FrmTT.DefInstance.Timer1.Enabled = True 'start clock
	End Sub
	
	
	Sub StopTime()
		FrmTime.DefInstance.Hide() 'turn off guage
		FrmTT.DefInstance.Timer1.Enabled = False 'turn off timer
	End Sub
End Module