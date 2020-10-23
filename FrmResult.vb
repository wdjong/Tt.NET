Option Strict Off
Option Explicit On
Friend Class FrmResult
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
    Public WithEvents ChkAllow As System.Windows.Forms.CheckBox
    Public WithEvents TxTime As System.Windows.Forms.TextBox
    Public WithEvents LstError As System.Windows.Forms.ListBox
    Public WithEvents TxAccuracy As System.Windows.Forms.TextBox
    Public WithEvents CmdOk As System.Windows.Forms.Button
    Public WithEvents TxSpeed As System.Windows.Forms.TextBox
    Public WithEvents LblAllow As System.Windows.Forms.Label
    Public WithEvents LblTime As System.Windows.Forms.Label
    Public WithEvents LblList As System.Windows.Forms.Label
    Public WithEvents LblAccuracy As System.Windows.Forms.Label
    Public WithEvents LblSpeed As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmResult))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ChkAllow = New System.Windows.Forms.CheckBox
        Me.TxTime = New System.Windows.Forms.TextBox
        Me.LstError = New System.Windows.Forms.ListBox
        Me.TxAccuracy = New System.Windows.Forms.TextBox
        Me.CmdOk = New System.Windows.Forms.Button
        Me.TxSpeed = New System.Windows.Forms.TextBox
        Me.LblAllow = New System.Windows.Forms.Label
        Me.LblTime = New System.Windows.Forms.Label
        Me.LblList = New System.Windows.Forms.Label
        Me.LblAccuracy = New System.Windows.Forms.Label
        Me.LblSpeed = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'chkAllow
        '
        Me.ChkAllow.BackColor = System.Drawing.SystemColors.Control
        Me.ChkAllow.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkAllow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkAllow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkAllow.Location = New System.Drawing.Point(304, 296)
        Me.ChkAllow.Name = "chkAllow"
        Me.ChkAllow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkAllow.Size = New System.Drawing.Size(113, 25)
        Me.ChkAllow.TabIndex = 10
        Me.ChkAllow.TabStop = False
        '
        'txTime
        '
        Me.TxTime.AcceptsReturn = True
        Me.TxTime.AutoSize = False
        Me.TxTime.BackColor = System.Drawing.SystemColors.Control
        Me.TxTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxTime.Location = New System.Drawing.Point(88, 296)
        Me.TxTime.MaxLength = 0
        Me.TxTime.Name = "txTime"
        Me.TxTime.ReadOnly = True
        Me.TxTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxTime.Size = New System.Drawing.Size(60, 24)
        Me.TxTime.TabIndex = 9
        Me.TxTime.TabStop = False
        Me.TxTime.Text = "Text1"
        '
        'lstError
        '
        Me.LstError.BackColor = System.Drawing.SystemColors.Control
        Me.LstError.Cursor = System.Windows.Forms.Cursors.Default
        Me.LstError.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstError.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LstError.ItemHeight = 14
        Me.LstError.Location = New System.Drawing.Point(24, 72)
        Me.LstError.Name = "lstError"
        Me.LstError.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LstError.Size = New System.Drawing.Size(393, 200)
        Me.LstError.TabIndex = 5
        Me.LstError.TabStop = False
        '
        'TxAccuracy
        '
        Me.TxAccuracy.AcceptsReturn = True
        Me.TxAccuracy.AutoSize = False
        Me.TxAccuracy.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxAccuracy.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxAccuracy.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxAccuracy.Location = New System.Drawing.Point(304, 24)
        Me.TxAccuracy.MaxLength = 0
        Me.TxAccuracy.Name = "TxAccuracy"
        Me.TxAccuracy.ReadOnly = True
        Me.TxAccuracy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxAccuracy.Size = New System.Drawing.Size(113, 19)
        Me.TxAccuracy.TabIndex = 4
        Me.TxAccuracy.TabStop = False
        Me.TxAccuracy.Text = ""
        '
        'CmdOk
        '
        Me.CmdOk.BackColor = System.Drawing.SystemColors.Control
        Me.CmdOk.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdOk.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdOk.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdOk.Location = New System.Drawing.Point(360, 328)
        Me.CmdOk.Name = "CmdOk"
        Me.CmdOk.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdOk.Size = New System.Drawing.Size(57, 25)
        Me.CmdOk.TabIndex = 3
        Me.CmdOk.TabStop = False
        Me.CmdOk.Text = "Ok"
        '
        'TxSpeed
        '
        Me.TxSpeed.AcceptsReturn = True
        Me.TxSpeed.AutoSize = False
        Me.TxSpeed.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxSpeed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxSpeed.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxSpeed.Location = New System.Drawing.Point(96, 24)
        Me.TxSpeed.MaxLength = 0
        Me.TxSpeed.Name = "TxSpeed"
        Me.TxSpeed.ReadOnly = True
        Me.TxSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxSpeed.Size = New System.Drawing.Size(113, 19)
        Me.TxSpeed.TabIndex = 2
        Me.TxSpeed.TabStop = False
        Me.TxSpeed.Text = ""
        '
        'lblAllow
        '
        Me.LblAllow.BackColor = System.Drawing.SystemColors.Control
        Me.LblAllow.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblAllow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAllow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblAllow.Location = New System.Drawing.Point(232, 296)
        Me.LblAllow.Name = "lblAllow"
        Me.LblAllow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblAllow.Size = New System.Drawing.Size(64, 25)
        Me.LblAllow.TabIndex = 8
        Me.LblAllow.Text = "Allow Edits"
        '
        'lblTime
        '
        Me.LblTime.BackColor = System.Drawing.SystemColors.Control
        Me.LblTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblTime.Location = New System.Drawing.Point(24, 296)
        Me.LblTime.Name = "lblTime"
        Me.LblTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblTime.Size = New System.Drawing.Size(65, 25)
        Me.LblTime.TabIndex = 7
        Me.LblTime.Text = "Time (secs)"
        '
        'lblList
        '
        Me.LblList.BackColor = System.Drawing.SystemColors.Control
        Me.LblList.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblList.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblList.Location = New System.Drawing.Point(24, 56)
        Me.LblList.Name = "lblList"
        Me.LblList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblList.Size = New System.Drawing.Size(89, 17)
        Me.LblList.TabIndex = 6
        Me.LblList.Text = "Errors"
        '
        'lblAccuracy
        '
        Me.LblAccuracy.BackColor = System.Drawing.SystemColors.Control
        Me.LblAccuracy.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblAccuracy.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAccuracy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblAccuracy.Location = New System.Drawing.Point(232, 24)
        Me.LblAccuracy.Name = "lblAccuracy"
        Me.LblAccuracy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblAccuracy.Size = New System.Drawing.Size(65, 17)
        Me.LblAccuracy.TabIndex = 1
        Me.LblAccuracy.Text = "Accuracy"
        '
        'lblSpeed
        '
        Me.LblSpeed.BackColor = System.Drawing.SystemColors.Control
        Me.LblSpeed.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblSpeed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblSpeed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblSpeed.Location = New System.Drawing.Point(24, 24)
        Me.LblSpeed.Name = "lblSpeed"
        Me.LblSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblSpeed.Size = New System.Drawing.Size(65, 17)
        Me.LblSpeed.TabIndex = 0
        Me.LblSpeed.Text = "Speed"
        '
        'FrmResult
        '
        Me.AcceptButton = Me.CmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(441, 375)
        Me.Controls.Add(Me.ChkAllow)
        Me.Controls.Add(Me.TxTime)
        Me.Controls.Add(Me.LstError)
        Me.Controls.Add(Me.TxAccuracy)
        Me.Controls.Add(Me.CmdOk)
        Me.Controls.Add(Me.TxSpeed)
        Me.Controls.Add(Me.LblAllow)
        Me.Controls.Add(Me.LblTime)
        Me.Controls.Add(Me.LblList)
        Me.Controls.Add(Me.LblAccuracy)
        Me.Controls.Add(Me.LblSpeed)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(91, 100)
        Me.Name = "FrmResult"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Results"
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FrmResult
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmResult
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmResult()
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
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        AddHistoryRecord()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        FrmResult.DefInstance.Hide()
    End Sub

    Private Sub CmdOk_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmdOk.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Space Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub FrmResult_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        CentreForm(FrmResult.DefInstance)
    End Sub


    'Event FrmResult.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub FrmResult_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        'If ScaleWidth > 6735 And ScaleHeight > 5505 Then
        Dim vgrid As Short

        vgrid = VB6.PixelsToTwipsY(ClientRectangle.Height) / 20

        LblSpeed.Top = VB6.TwipsToPixelsY(2 * vgrid) '360
        LblSpeed.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) / 20) '150
        LblSpeed.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) / 20) '285
        LblSpeed.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) / 5) '975
        TxSpeed.Top = LblSpeed.Top
        TxSpeed.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(LblSpeed.Left) + VB6.PixelsToTwipsX(LblSpeed.Width))
        TxSpeed.Height = LblSpeed.Height
        TxSpeed.Width = LblSpeed.Width

        LblAccuracy.Top = LblSpeed.Top
        LblAccuracy.Height = LblSpeed.Height
        LblAccuracy.Width = LblSpeed.Width
        LblAccuracy.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) - VB6.PixelsToTwipsX(LblSpeed.Left) - VB6.PixelsToTwipsX(LblSpeed.Width) * 2)
        TxAccuracy.Top = LblAccuracy.Top
        TxAccuracy.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(LblAccuracy.Left) + VB6.PixelsToTwipsX(LblAccuracy.Width))
        TxAccuracy.Height = LblAccuracy.Height
        TxAccuracy.Width = LblAccuracy.Width

        LblList.Top = VB6.TwipsToPixelsY(4 * vgrid) '720
        LblList.Left = LblSpeed.Left
        LblList.Height = VB6.TwipsToPixelsY(1 * vgrid)
        LblList.Width = LblSpeed.Width
        LstError.Top = VB6.TwipsToPixelsY(5 * vgrid)
        LstError.Left = LblList.Left
        LstError.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) - VB6.PixelsToTwipsX(LstError.Left) * 2)
        LstError.Height = VB6.TwipsToPixelsY(8 * vgrid)

        LblTime.Top = VB6.TwipsToPixelsY(14 * vgrid)
        LblTime.Left = LblSpeed.Left
        LblTime.Height = LblSpeed.Height
        LblTime.Width = LblSpeed.Width
        TxTime.Top = LblTime.Top
        TxTime.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(LblTime.Left) + VB6.PixelsToTwipsX(LblTime.Width))
        TxTime.Height = LblTime.Height
        TxTime.Width = LblTime.Width

        LblAllow.Top = LblTime.Top
        LblAllow.Left = LblAccuracy.Left
        LblAllow.Height = LblSpeed.Height
        LblAllow.Width = LblSpeed.Width
        ChkAllow.Top = LblTime.Top
        ChkAllow.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(LblAllow.Left) + VB6.PixelsToTwipsX(LblAllow.Width))

        CmdOk.Width = VB6.TwipsToPixelsX(0.5 * VB6.PixelsToTwipsX(LblSpeed.Width))
        CmdOk.Height = VB6.TwipsToPixelsY(1.5 * vgrid)
        CmdOk.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) - VB6.PixelsToTwipsX(CmdOk.Width) - VB6.PixelsToTwipsX(LblSpeed.Left))
        CmdOk.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) - VB6.PixelsToTwipsY(CmdOk.Height) - VB6.PixelsToTwipsY(LblSpeed.Top))
        'End If
    End Sub
End Class