Option Strict Off
Option Explicit On
Friend Class FrmOrig
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
    Public WithEvents TxtOrig As System.Windows.Forms.TextBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmOrig))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TxtOrig = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'txtOrig
        '
        Me.TxtOrig.AcceptsReturn = True
        Me.TxtOrig.AutoSize = False
        Me.TxtOrig.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtOrig.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtOrig.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtOrig.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtOrig.Location = New System.Drawing.Point(0, 0)
        Me.TxtOrig.MaxLength = 0
        Me.TxtOrig.Multiline = True
        Me.TxtOrig.Name = "txtOrig"
        Me.TxtOrig.ReadOnly = True
        Me.TxtOrig.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtOrig.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtOrig.Size = New System.Drawing.Size(446, 397)
        Me.TxtOrig.TabIndex = 0
        Me.TxtOrig.Text = "txtOrig"
        '
        'FrmOrig
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(446, 397)
        Me.Controls.Add(Me.TxtOrig)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(8, Byte), CType(0, Byte), CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(46, 101)
        Me.Name = "FrmOrig"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Original"
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FrmOrig
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmOrig
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmOrig()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Private Sub FrmOrig_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Program originally designed for use with hardcopy original
        FrmOrig.DefInstance.Top = FrmTT.DefInstance.Top
        FrmOrig.DefInstance.Left = FrmTT.DefInstance.Left
        FrmOrig.DefInstance.Width = FrmTT.DefInstance.Width
        FrmOrig.DefInstance.Height = FrmTT.DefInstance.Height
    End Sub

    Private Sub FrmOrig_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = 2 Then 'maximizedtxtCopy.Top = 0
            WindowState = System.Windows.Forms.FormWindowState.Normal 'normal
            Me.Top = 0 'problem with maximizing
            Me.Left = 0 'when viewing original window
            Me.Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
            Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - szTitle) 'for taskbar
        End If
        TxtOrig.Top = 0
        TxtOrig.Left = 0
        TxtOrig.Width = ClientRectangle.Width
        TxtOrig.Height = ClientRectangle.Height
        'Calculate font size 20 converts twips to points
        '65 characters per line.
        '1.5 is ratio of horizontal to vertical
        TxtOrig.Font = VB6.FontChangeSize(TxtOrig.Font, VB6.PixelsToTwipsX(TxtOrig.Width) / 20 / 65 * 1.5)
    End Sub

    Private Sub FrmOrig_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        FrmTT.DefInstance.Top = Me.Top 'aTop unarrange window
        FrmTT.DefInstance.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) + VB6.PixelsToTwipsY(FrmTT.DefInstance.Height)) 'aHeight
        fil = "" 'clear original selection
    End Sub

    Private Sub TxtOrig_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtOrig.Enter
        On Error GoTo txtOrig_GotFocus_error
        FrmTT.DefInstance.TxtCopy.Focus()
        On Error GoTo 0
        GoTo txtOrig_GotFocus_exit

txtOrig_GotFocus_error:
        Resume Next 'this happened when marking a test i'd edited from an opened copy.
        'modal frmResults showing.

txtOrig_GotFocus_exit:

    End Sub
End Class