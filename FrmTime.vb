Option Strict Off
Option Explicit On
Friend Class FrmTime
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
    'Public WithEvents Gauge1 As AxMSComctlLib.AxProgressBar
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents Gauge1 As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmTime))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Gauge1 = New System.Windows.Forms.ProgressBar
        Me.SuspendLayout()
        '
        'Gauge1
        '
        Me.Gauge1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Gauge1.Location = New System.Drawing.Point(80, -8)
        Me.Gauge1.Name = "Gauge1"
        Me.Gauge1.Size = New System.Drawing.Size(424, 32)
        Me.Gauge1.TabIndex = 0
        '
        'FrmTime
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(592, 22)
        Me.Controls.Add(Me.Gauge1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(22, 119)
        Me.Name = "FrmTime"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Time"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmTime
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmTime
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmTime()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 

    Private Sub FrmTime_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FrmTT.DefInstance.Top) + VB6.PixelsToTwipsY(FrmTT.DefInstance.Height) - 360)
        Left = FrmTT.DefInstance.Left
        Width = FrmTT.DefInstance.Width
        Height = VB6.TwipsToPixelsY(720)
    End Sub

    Private Sub FrmTime_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FrmTT.DefInstance.Top) + VB6.PixelsToTwipsY(FrmTT.DefInstance.Height) - 360)
        Left = FrmTT.DefInstance.Left
        Width = FrmTT.DefInstance.Width
        Height = VB6.TwipsToPixelsY(720)
    End Sub

    Private Sub Gauge1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Gauge1.Enter
        FrmTT.DefInstance.TxtCopy.Focus()
    End Sub
End Class