Option Strict Off
Option Explicit On
Friend Class DlgOptions
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents ChkBS As System.Windows.Forms.CheckBox
	Public WithEvents CmdOk As System.Windows.Forms.Button
	Public WithEvents TxtTestlen As System.Windows.Forms.TextBox
	Public WithEvents LblBS As System.Windows.Forms.Label
	Public WithEvents LblTestlen As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DlgOptions))
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.ChkBS = New System.Windows.Forms.CheckBox
		Me.CmdOk = New System.Windows.Forms.Button
		Me.TxtTestlen = New System.Windows.Forms.TextBox
		Me.LblBS = New System.Windows.Forms.Label
		Me.LblTestlen = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'chkBS
		'
		Me.ChkBS.BackColor = System.Drawing.SystemColors.Control
		Me.ChkBS.Cursor = System.Windows.Forms.Cursors.Default
		Me.ChkBS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ChkBS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ChkBS.Location = New System.Drawing.Point(160, 72)
		Me.ChkBS.Name = "chkBS"
		Me.ChkBS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ChkBS.Size = New System.Drawing.Size(25, 25)
		Me.ChkBS.TabIndex = 2
		'
		'cmdOk
		'
		Me.CmdOk.BackColor = System.Drawing.SystemColors.Control
		Me.CmdOk.Cursor = System.Windows.Forms.Cursors.Default
		Me.CmdOk.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CmdOk.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CmdOk.Location = New System.Drawing.Point(296, 96)
		Me.CmdOk.Name = "cmdOk"
		Me.CmdOk.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CmdOk.Size = New System.Drawing.Size(57, 25)
		Me.CmdOk.TabIndex = 3
		Me.CmdOk.TabStop = False
		Me.CmdOk.Text = "Ok"
		'
		'txtTestlen
		'
		Me.TxtTestlen.AcceptsReturn = True
		Me.TxtTestlen.AutoSize = False
		Me.TxtTestlen.BackColor = System.Drawing.SystemColors.Window
		Me.TxtTestlen.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtTestlen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TxtTestlen.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtTestlen.Location = New System.Drawing.Point(160, 24)
		Me.TxtTestlen.MaxLength = 0
		Me.TxtTestlen.Name = "txtTestlen"
		Me.TxtTestlen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtTestlen.Size = New System.Drawing.Size(129, 19)
		Me.TxtTestlen.TabIndex = 1
		Me.TxtTestlen.Text = ""
		'
		'lblBS
		'
		Me.LblBS.BackColor = System.Drawing.SystemColors.Control
		Me.LblBS.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblBS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblBS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblBS.Location = New System.Drawing.Point(24, 72)
		Me.LblBS.Name = "lblBS"
		Me.LblBS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblBS.Size = New System.Drawing.Size(121, 25)
		Me.LblBS.TabIndex = 4
		Me.LblBS.Text = "Allow editing"
		'
		'lblTestlen
		'
		Me.LblTestlen.BackColor = System.Drawing.SystemColors.Control
		Me.LblTestlen.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblTestlen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblTestlen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblTestlen.Location = New System.Drawing.Point(24, 24)
		Me.LblTestlen.Name = "lblTestlen"
		Me.LblTestlen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblTestlen.Size = New System.Drawing.Size(105, 25)
		Me.LblTestlen.TabIndex = 0
		Me.LblTestlen.Text = "Test length (seconds)"
		'
		'frmOptions
		'
		Me.AcceptButton = Me.CmdOk
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ClientSize = New System.Drawing.Size(368, 134)
		Me.Controls.Add(Me.ChkBS)
		Me.Controls.Add(Me.CmdOk)
		Me.Controls.Add(Me.TxtTestlen)
		Me.Controls.Add(Me.LblBS)
		Me.Controls.Add(Me.LblTestlen)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Location = New System.Drawing.Point(119, 147)
		Me.Name = "frmOptions"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Options"
		Me.ResumeLayout(False)

	End Sub
#End Region
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As DlgOptions
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As DlgOptions
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New DlgOptions()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region

	Private Sub CmdOk_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOk.Click
		Dim newLen As Short

		On Error Resume Next
		newLen = Val(TxtTestlen.Text)
		If newLen > 0 Then
			tesLen = newLen
			My.Settings.testLength = tesLen
		End If
		On Error GoTo 0
		My.Settings.allowEdits = ChkBS.Checked

		Me.Close()
	End Sub

	Private Sub FrmOptions_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Turn off any existing test
		StopTime()
		tRemains = 0 'reset clock
		FrmTT.DefInstance.TxtCopy.ReadOnly = False 'locked after first test is finished
		FrmTT.DefInstance.TxtCopy.Text = "" 'clear copy
		CentreForm(Me)

		TxtTestlen.Text = My.Settings.testLength
		ChkBS.Checked = My.Settings.allowEdits
	End Sub

	Private Sub MnuFileExit_Click()
		Me.Close()
	End Sub

End Class