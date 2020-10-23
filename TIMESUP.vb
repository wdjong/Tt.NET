Option Strict Off
Option Explicit On
Friend Class FrmTimesUp
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
    Public WithEvents TxTestlen As System.Windows.Forms.TextBox
    Public WithEvents CmdMark As System.Windows.Forms.Button
    Public WithEvents Label1_3 As System.Windows.Forms.Label
    Public WithEvents Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents LblTestlen As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmTimesUp))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TxTestlen = New System.Windows.Forms.TextBox
        Me.CmdMark = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.Label1_3 = New System.Windows.Forms.Label
        Me.Label1_1 = New System.Windows.Forms.Label
        Me.LblTestlen = New System.Windows.Forms.Label
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txTestlen
        '
        Me.TxTestlen.AutoSize = False
        Me.TxTestlen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxTestlen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxTestlen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxTestlen.Location = New System.Drawing.Point(136, 144)
        Me.TxTestlen.MaxLength = 0
        Me.TxTestlen.Name = "txTestlen"
        Me.TxTestlen.ReadOnly = True
        Me.TxTestlen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxTestlen.Size = New System.Drawing.Size(113, 19)
        Me.TxTestlen.TabIndex = 0
        Me.TxTestlen.TabStop = False
        Me.TxTestlen.Text = "Text1"
        '
        'cmdMark
        '
        Me.CmdMark.BackColor = System.Drawing.SystemColors.Control
        Me.CmdMark.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdMark.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdMark.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdMark.Location = New System.Drawing.Point(264, 144)
        Me.CmdMark.Name = "cmdMark"
        Me.CmdMark.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdMark.Size = New System.Drawing.Size(57, 25)
        Me.CmdMark.TabIndex = 1
        Me.CmdMark.TabStop = False
        Me.CmdMark.Text = "Ok"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Label1_3)
        Me.Frame1.Controls.Add(Me.Label1_1)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(305, 113)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        '
        '_Label1_3
        '
        Me.Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me.Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1_3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me.Label1_3, CType(3, Short))
        Me.Label1_3.Location = New System.Drawing.Point(16, 64)
        Me.Label1_3.Name = "_Label1_3"
        Me.Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1_3.Size = New System.Drawing.Size(233, 25)
        Me.Label1_3.TabIndex = 5
        Me.Label1_3.Text = "You can't type any more now."
        '
        '_Label1_1
        '
        Me.Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1_1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me.Label1_1, CType(1, Short))
        Me.Label1_1.Location = New System.Drawing.Point(16, 24)
        Me.Label1_1.Name = "_Label1_1"
        Me.Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1_1.Size = New System.Drawing.Size(129, 25)
        Me.Label1_1.TabIndex = 4
        Me.Label1_1.Text = "Your time is up. "
        '
        'lblTestlen
        '
        Me.LblTestlen.BackColor = System.Drawing.SystemColors.Control
        Me.LblTestlen.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblTestlen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTestlen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblTestlen.Location = New System.Drawing.Point(32, 144)
        Me.LblTestlen.Name = "lblTestlen"
        Me.LblTestlen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblTestlen.Size = New System.Drawing.Size(97, 17)
        Me.LblTestlen.TabIndex = 2
        Me.LblTestlen.Text = "Test length (secs)"
        '
        'FrmTimesUp
        '
        Me.AcceptButton = Me.CmdMark
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(336, 186)
        Me.Controls.Add(Me.TxTestlen)
        Me.Controls.Add(Me.CmdMark)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.LblTestlen)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(186, 104)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTimesUp"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Time's up!"
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FrmTimesUp
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmTimesUp
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmTimesUp()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Private Sub CmdMark_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdMark.Click
        FrmTimesUp.DefInstance.Hide()
        If fil = "" Then 'no original specified
            ChooseOriginal()
        End If
        MarkTest()
    End Sub


    Private Sub CmdMark_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmdMark.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Space Then
            MsgBox("You can stop typing now.")
            KeyAscii = 0
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub FrmTimesUp_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Beep()
		CentreForm(FrmTimesUp.DefInstance)
		txTestlen.Text = CStr(tesLen - tRemains)
	End Sub
End Class