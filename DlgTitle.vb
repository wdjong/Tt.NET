Option Strict Off
Option Explicit On
Friend Class FrmTitle
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
    Public WithEvents CmdContinue As System.Windows.Forms.Button
    Public WithEvents TxtName As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents LblName As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmTitle))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CmdContinue = New System.Windows.Forms.Button
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.LblName = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdContinue
        '
        Me.CmdContinue.BackColor = System.Drawing.SystemColors.Control
        Me.CmdContinue.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdContinue.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdContinue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdContinue.Location = New System.Drawing.Point(288, 248)
        Me.CmdContinue.Name = "cmdContinue"
        Me.CmdContinue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdContinue.Size = New System.Drawing.Size(57, 25)
        Me.CmdContinue.TabIndex = 1
        Me.CmdContinue.TabStop = False
        Me.CmdContinue.Text = "Ok"
        '
        'txtName
        '
        Me.TxtName.AcceptsReturn = True
        Me.TxtName.AutoSize = False
        Me.TxtName.BackColor = System.Drawing.SystemColors.Window
        Me.TxtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtName.Location = New System.Drawing.Point(48, 208)
        Me.TxtName.MaxLength = 0
        Me.TxtName.Name = "txtName"
        Me.TxtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtName.Size = New System.Drawing.Size(249, 19)
        Me.TxtName.TabIndex = 0
        Me.TxtName.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(24, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(297, 41)
        Me.Label2.TabIndex = 3
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblName
        '
        Me.LblName.BackColor = System.Drawing.SystemColors.Control
        Me.LblName.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblName.Location = New System.Drawing.Point(48, 184)
        Me.LblName.Name = "lblName"
        Me.LblName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblName.Size = New System.Drawing.Size(137, 17)
        Me.LblName.TabIndex = 4
        Me.LblName.Text = "Enter your name"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(80, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(161, 33)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "TypeTest"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'FrmTitle
        '
        Me.AcceptButton = Me.CmdContinue
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(360, 304)
        Me.Controls.Add(Me.CmdContinue)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LblName)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(18, Byte), CType(0, Byte), CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(29, 152)
        Me.Name = "FrmTitle"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Sign-in"
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FrmTitle
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmTitle
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmTitle()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Private Sub CmdContinue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdContinue.Click
        'FrmTitle.DefInstance.Hide()
        Me.Close()
        'form closing event fires
    End Sub

    Private Sub FrmTitle_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Label2.Text = "Welcome to TypeTest. You can use this to give yourself a standard AS2708 style test." '"You can get your own copy of TypeTest through the Internet" & vbCrLf & " (currently http://users.tpg.com.au/dejong58/typetest). "
        'If TTlocked Then
        '    Label2.Text = Label2.Text & vbCrLf & "The program is currently locked."
        'End If
        CentreForm(Me)
    End Sub

    Private Sub FrmTitle_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        'Label1.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) / 5)
        'Label1.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) / 6)
        'Label1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) - VB6.PixelsToTwipsX(Label1.Left) * 2)
        'Label2.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) * 2 / 5)
        'Label2.Left = Label1.Left
        'Label2.Width = Label1.Width
        'lblName.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) * 3 / 5)
        'lblName.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) / 3)
        'lblName.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) / 3)
        'txtName.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblName.Top) + VB6.PixelsToTwipsY(lblName.Height))
        'txtName.Left = lblName.Left
        'txtName.Width = lblName.Width
        'cmdContinue.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtName.Left) + VB6.PixelsToTwipsX(txtName.Width) - VB6.PixelsToTwipsX(cmdContinue.Width))
        'cmdContinue.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) * 4 / 5)
    End Sub

    Private Sub FrmTitle_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.TxtName.Text = "" Then
            Me.TxtName.Text = "Anon"
        End If
        FrmTT.DefInstance.Text = "TypeTest: " & Me.TxtName.Text
        userName = Me.TxtName.Text
    End Sub
End Class