Option Strict Off
Option Explicit On
Imports System.Drawing.Printing
Imports System.IO 'for streamreader

Friend Class FrmTT
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
    Public WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents TxtCopy As System.Windows.Forms.TextBox
    'Public WithEvents CD1 As MSComDlg.CommonDialog 'AxMSComDlg.AxCommonDialog
    Public WithEvents MnuFileExit As System.Windows.Forms.MenuItem
    Public WithEvents MnuFile As System.Windows.Forms.MenuItem
    Public WithEvents MnuOrigChoose As System.Windows.Forms.MenuItem
    Public WithEvents MnuOrigView As System.Windows.Forms.MenuItem
    Public WithEvents MnuOriginalHide As System.Windows.Forms.MenuItem
    Public WithEvents MnuOrigPrint As System.Windows.Forms.MenuItem
    Public WithEvents MnuOrig As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopyNew As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopyOpen As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopyPreferences As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopyMark As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopySave As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopyPrint As System.Windows.Forms.MenuItem
    Public WithEvents MnuCopy As System.Windows.Forms.MenuItem
    Public WithEvents MnuResultChange As System.Windows.Forms.MenuItem
    Public WithEvents MnuResultHistory As System.Windows.Forms.MenuItem
    Public WithEvents CmdResult As System.Windows.Forms.MenuItem
    Public WithEvents MnuHelpContents As System.Windows.Forms.MenuItem
    Public WithEvents MnuHelpAbout As System.Windows.Forms.MenuItem
    Public WithEvents MnuHelp As System.Windows.Forms.MenuItem
    Public MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents PrintDialog1 As PrintDialog
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents Printer As System.Drawing.Printing.PrintDocument

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTT))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TxtCopy = New System.Windows.Forms.TextBox()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MnuFile = New System.Windows.Forms.MenuItem()
        Me.MnuFileExit = New System.Windows.Forms.MenuItem()
        Me.MnuOrig = New System.Windows.Forms.MenuItem()
        Me.MnuOrigChoose = New System.Windows.Forms.MenuItem()
        Me.MnuOrigView = New System.Windows.Forms.MenuItem()
        Me.MnuOriginalHide = New System.Windows.Forms.MenuItem()
        Me.MnuOrigPrint = New System.Windows.Forms.MenuItem()
        Me.MnuCopy = New System.Windows.Forms.MenuItem()
        Me.MnuCopyNew = New System.Windows.Forms.MenuItem()
        Me.MnuCopyOpen = New System.Windows.Forms.MenuItem()
        Me.MnuCopyPreferences = New System.Windows.Forms.MenuItem()
        Me.MnuCopyMark = New System.Windows.Forms.MenuItem()
        Me.MnuCopySave = New System.Windows.Forms.MenuItem()
        Me.MnuCopyPrint = New System.Windows.Forms.MenuItem()
        Me.CmdResult = New System.Windows.Forms.MenuItem()
        Me.MnuResultChange = New System.Windows.Forms.MenuItem()
        Me.MnuResultHistory = New System.Windows.Forms.MenuItem()
        Me.MnuHelp = New System.Windows.Forms.MenuItem()
        Me.MnuHelpContents = New System.Windows.Forms.MenuItem()
        Me.MnuHelpAbout = New System.Windows.Forms.MenuItem()
        Me.Printer = New System.Drawing.Printing.PrintDocument()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'TxtCopy
        '
        Me.TxtCopy.AcceptsReturn = True
        Me.TxtCopy.AcceptsTab = True
        Me.TxtCopy.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCopy.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtCopy.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtCopy.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtCopy.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCopy.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtCopy.Location = New System.Drawing.Point(0, 0)
        Me.TxtCopy.MaxLength = 0
        Me.TxtCopy.Multiline = True
        Me.TxtCopy.Name = "TxtCopy"
        Me.TxtCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtCopy.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtCopy.Size = New System.Drawing.Size(573, 387)
        Me.TxtCopy.TabIndex = 0
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuFile, Me.MnuOrig, Me.MnuCopy, Me.CmdResult, Me.MnuHelp})
        '
        'MnuFile
        '
        Me.MnuFile.Index = 0
        Me.MnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuFileExit})
        Me.MnuFile.Text = "&File"
        '
        'MnuFileExit
        '
        Me.MnuFileExit.Index = 0
        Me.MnuFileExit.Text = "E&xit"
        '
        'MnuOrig
        '
        Me.MnuOrig.Index = 1
        Me.MnuOrig.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuOrigChoose, Me.MnuOrigView, Me.MnuOriginalHide, Me.MnuOrigPrint})
        Me.MnuOrig.Text = "&Original"
        '
        'MnuOrigChoose
        '
        Me.MnuOrigChoose.Index = 0
        Me.MnuOrigChoose.Text = "&Choose"
        '
        'MnuOrigView
        '
        Me.MnuOrigView.Index = 1
        Me.MnuOrigView.Text = "&View"
        '
        'MnuOriginalHide
        '
        Me.MnuOriginalHide.Index = 2
        Me.MnuOriginalHide.Text = "&Hide"
        '
        'MnuOrigPrint
        '
        Me.MnuOrigPrint.Index = 3
        Me.MnuOrigPrint.Text = "&Print"
        '
        'MnuCopy
        '
        Me.MnuCopy.Index = 2
        Me.MnuCopy.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuCopyNew, Me.MnuCopyOpen, Me.MnuCopyPreferences, Me.MnuCopyMark, Me.MnuCopySave, Me.MnuCopyPrint})
        Me.MnuCopy.Text = "&Copy"
        '
        'MnuCopyNew
        '
        Me.MnuCopyNew.Index = 0
        Me.MnuCopyNew.Text = "&New"
        '
        'MnuCopyOpen
        '
        Me.MnuCopyOpen.Index = 1
        Me.MnuCopyOpen.Text = "&Open"
        '
        'MnuCopyPreferences
        '
        Me.MnuCopyPreferences.Index = 2
        Me.MnuCopyPreferences.Text = "P&references"
        '
        'MnuCopyMark
        '
        Me.MnuCopyMark.Index = 3
        Me.MnuCopyMark.Text = "&Mark"
        '
        'MnuCopySave
        '
        Me.MnuCopySave.Index = 4
        Me.MnuCopySave.Text = "&Save"
        '
        'MnuCopyPrint
        '
        Me.MnuCopyPrint.Index = 5
        Me.MnuCopyPrint.Text = "&Print"
        '
        'CmdResult
        '
        Me.CmdResult.Index = 3
        Me.CmdResult.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuResultChange, Me.MnuResultHistory})
        Me.CmdResult.Text = "&Result"
        '
        'MnuResultChange
        '
        Me.MnuResultChange.Index = 0
        Me.MnuResultChange.Text = "&Change Name"
        '
        'MnuResultHistory
        '
        Me.MnuResultHistory.Index = 1
        Me.MnuResultHistory.Text = "&History"
        '
        'MnuHelp
        '
        Me.MnuHelp.Index = 4
        Me.MnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuHelpContents, Me.MnuHelpAbout})
        Me.MnuHelp.Text = "&Help"
        '
        'MnuHelpContents
        '
        Me.MnuHelpContents.Index = 0
        Me.MnuHelpContents.Text = "&Contents"
        '
        'MnuHelpAbout
        '
        Me.MnuHelpAbout.Index = 1
        Me.MnuHelpAbout.Text = "&About"
        '
        'Printer
        '
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'FrmTT
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(573, 387)
        Me.Controls.Add(Me.TxtCopy)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(83, 139)
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmTT"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TypeTest"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FrmTT
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmTT
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmTT()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    'Used in printpage event handler
    Private printFont As Font
    Private streamToPrint As StreamReader 'this holds the data output in Printer_PrintPage 
    Private Shared filePath As String

    'Menu options
    Public Sub MnuCopyOpen_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyOpen.Popup
        MnuCopyOpen_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuCopyOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyOpen.Click
        NewCopy() 'clear and turn off time
        ChooseCopy() 'select previously typed copy
    End Sub

    Public Sub MnuCopyPreferences_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyPreferences.Popup
        MnuCopyPreferences_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuCopyPreferences_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyPreferences.Click
        NewCopy() 'end any current test
        DlgOptions.DefInstance.ShowDialog()
    End Sub

    Public Sub MnuCopyPrint_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyPrint.Popup
        MnuCopyPrint_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuCopyPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyPrint.Click

        PrintDialog1.Document = Printer
        PrintDialog1.PrinterSettings = Printer.PrinterSettings
        PrintDialog1.AllowSomePages = True

        If PrintDialog1.ShowDialog = DialogResult.OK Then
            Printer.PrinterSettings = PrintDialog1.PrinterSettings

            If Me.TxtCopy.Text <> "" Then
                StopTime() 'end current test
                filePath = VB6.GetPath & "\temp"
                Try
                    streamToPrint = New StreamReader(filePath)
                    printFont = New Font("Arial", 10)
                Catch ex As Exception
                    MsgBox("MnuCopyPrint_Click: OpenFile: " & ex.Message)
                End Try
                Try
                    Me.Printer.Print() 'See PrintPage event handler
                Catch ex As Exception
                    MsgBox("MnuCopyPrint_Click: Printing: " & ex.Message)
                End Try
                streamToPrint.Close()
            End If
        End If

    End Sub

    Public Sub MnuCopySave_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopySave.Popup
        MnuCopySave_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuCopySave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopySave.Click
        Dim fileSaveName As String = ""
        Dim fd As OpenFileDialog = New OpenFileDialog()

        If Me.TxtCopy.Text <> "" Then
            StopTime()

            fd.Title = "Save copy"
            fd.InitialDirectory = VB6.GetPath 'look in same directory
            fd.Filter = "All files (*.*)|*.*"
            fd.FilterIndex = 1 'test*.*
            fd.FileName = "temp"
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                fileSaveName = fd.FileName
                If fileSaveName <> "" Then
                    FileOpen(3, fileSaveName, OpenMode.Output)
                    On Error GoTo diskerr
                    Print(3, LTrim(CStr(tesLen / 60)))
                    PrintLine(3, FrmTT.DefInstance.TxtCopy.Text)
                    On Error GoTo 0
                    FileClose(3)
                End If
            End If

        End If
        GoTo endSave
diskerr:
        MsgBox("Error saving copy.")
        Resume Next
endSave:
    End Sub

    Public Sub MnuFileExit_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFileExit.Popup
        MnuFileExit_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuFileExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFileExit.Click
        End
    End Sub

    Public Sub MnuCopyMark_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyMark.Popup
        MnuCopyMark_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuCopyMark_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyMark.Click
        If Len(Me.TxtCopy.Text) > 0 Then 'they've typed something
            StopTime()
            If fil = "" Then 'no original specified
                ChooseOriginal()
            End If
            If fil <> "" Then 'did select original
                MarkTest() 'remark test
            End If
        End If
    End Sub

    Public Sub MnuCopyNew_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyNew.Popup
        MnuCopyNew_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuCopyNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCopyNew.Click
        NewCopy()
    End Sub

    Public Sub MnuHelpAbout_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuHelpAbout.Popup
        MnuHelpAbout_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuHelpAbout.Click
        NewCopy()
        FrmAbout.ShowDialog()
    End Sub

    Public Sub MnuOriginalHide_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOriginalHide.Popup
        MnuOriginalHide_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuOriginalHide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOriginalHide.Click
        If fil <> "" Then
            If VB6.PixelsToTwipsY(Me.Top) > VB6.PixelsToTwipsY(FrmOrig.DefInstance.Top) Then 'not already hidden
                Me.Top = FrmOrig.DefInstance.Top 'choose heighest aTop 'resize frmTT
                If VB6.PixelsToTwipsY(FrmOrig.DefInstance.Height) + VB6.PixelsToTwipsY(Me.Height) > VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - szTitle Then
                    Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - szTitle)
                Else
                    Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FrmOrig.DefInstance.Height) + VB6.PixelsToTwipsY(Me.Height)) 'aHeight
                End If
                FrmOrig.DefInstance.Left = Me.Left
                FrmOrig.DefInstance.Width = Me.Width
                FrmOrig.DefInstance.Height = Me.Height
                FrmOrig.DefInstance.Top = Me.Top
            End If
        End If
    End Sub

    Public Sub MnuOrigPrint_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOrigPrint.Popup
        MnuOrigPrint_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuOrigPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOrigPrint.Click
        PrintDialog1.Document = Printer
        PrintDialog1.PrinterSettings = Printer.PrinterSettings
        PrintDialog1.AllowSomePages = True

        If PrintDialog1.ShowDialog = DialogResult.OK Then
            Printer.PrinterSettings = PrintDialog1.PrinterSettings
            If fil <> "" Then
                filePath = fil
                Try
                    streamToPrint = New StreamReader(filePath)
                    printFont = New Font("Arial", 10)
                Catch ex As Exception
                    MsgBox("MnuOrigPrint_Click: OpenFile: " & ex.Message)
                End Try
                Try
                    Me.Printer.Print() 'See PrintPage event handler
                Catch ex As Exception
                    MsgBox("MnuOrigPrint_Click: Printing: " & ex.Message)
                End Try
                streamToPrint.Close()
            End If
        End If

    End Sub

    Public Sub MnuOrigView_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOrigView.Popup
        MnuOrigView_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuOrigView_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOrigView.Click
        'aTop and aHeight are global
        'they're initialised in the load
        'they contain the original top and height
        'Me.WindowState = 0 'prevent maximized window
        If fil = "" Then ChooseOriginal()
        If fil <> "" Then 'chose something
            If VB6.PixelsToTwipsY(FrmOrig.DefInstance.Top) = VB6.PixelsToTwipsY(FrmTT.DefInstance.Top) Then 'probably hidden
                FrmOrig.DefInstance.Top = Me.Top 'aTop
                FrmOrig.DefInstance.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) * 0.75) 'aHeight * 0.75
                FrmOrig.DefInstance.Left = Me.Left
                FrmOrig.DefInstance.Width = Me.Width
                Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FrmOrig.DefInstance.Top) + VB6.PixelsToTwipsY(FrmOrig.DefInstance.Height))
                Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(FrmOrig.DefInstance.Height)) 'aHeight - FrmOrig.Height
            Else 'already different
                MsgBox("You can move the windows by dragging them to where you would like them.")
            End If
        End If
    End Sub

    Public Sub MnuOrigChoose_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOrigChoose.Popup
        MnuOrigChoose_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuOrigChoose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOrigChoose.Click
        ChooseOriginal()
    End Sub

    'Form events

    Private Sub FrmTT_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        Application.Exit() '   End 'should unload everything
    End Sub

    Private Sub FrmTT_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            FrmTitle.ShowDialog()
            allowEdits = My.Settings.allowEdits 'allow backspace
            tesLen = My.Settings.testLength
            Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width * 0.75 ' Set width of form.
            If Width > 640 Then Width = 640 'originally designed for vga screen (actually before that it was text based 80 char by ? lines)
            Height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height * 0.75 ' Set height of form.
            If Height > 480 Then Height = 480
            CentreForm(FrmTT.DefInstance)
            TxtCopy.Font = VB6.FontChangeSize(TxtCopy.Font, VB6.PixelsToTwipsX(TxtCopy.Width) / 20 / 65 * 1.5)
        Catch ex As Exception
            MsgBox("FrmTT_Load: " & ex.Message)
        End Try
    End Sub

    Private Sub FrmTT_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        'Fires e.g. on returning from Title dialog box
        Try
            If Me.WindowState = FormWindowState.Maximized Then 'maximizedtxtCopy.Top = 0
                WindowState = System.Windows.Forms.FormWindowState.Normal 'normal
                Me.Top = 0 'problem with maximizing
                Me.Left = 0 'when viewing original window
                Me.Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
                Me.Height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - szTitle 'for taskbar
            End If
            TxtCopy.Font = VB6.FontChangeSize(TxtCopy.Font, VB6.PixelsToTwipsX(TxtCopy.Width) / 20 / 65 * 1.5)
        Catch ex As Exception
            MsgBox("FrmTT_Resize: Adjusting font: " & ex.Message)
        End Try
    End Sub

    Public Sub MnuResultChange_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuResultChange.Popup
        MnuResultChange_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuResultChange_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuResultChange.Click
        NewCopy() 'Clear any existing typing test attempt
        FrmTitle.DefInstance.ShowDialog()
    End Sub

    Public Sub MnuResultHistory_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuResultHistory.Popup
        MnuResultHistory_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuResultHistory_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuResultHistory.Click
        NewCopy()
        FrmHistory.DefInstance.ShowDialog()
    End Sub

    Public Sub MnuHelpContents_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuHelpContents.Popup
        MnuHelpContents_Click(eventSender, eventArgs)
    End Sub

    Public Sub MnuHelpContents_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuHelpContents.Click
        Try
            Dim AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
            System.Diagnostics.Process.Start(AppPath + "TTHelp.html")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TxtCopy_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCopy.TextChanged
        If Timer1.Enabled Then 'test in progress
            If TxtCopy.Text = "" Then ' stop time if no text
                NewCopy()
            End If
        Else 'timer not enabled - test not in progress
            If Len(TxtCopy.Text) = 1 Then ' they've typed the first letter.
                startTime()
                FrmTT.DefInstance.TxtCopy.Focus()
            Else 'timer off lots of text added
                If Len(TxtCopy.Text) > 1 Then 'Copy Open
                    startTime() 'start editing test
                    FrmTT.DefInstance.TxtCopy.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        If tRemains > 0 Then
            tRemains -= 1 'decrement 1 second
            FrmTime.DefInstance.Gauge1.Value = tRemains
        Else 'end of test
            If FrmTT.DefInstance.TxtCopy.Text = "" Then
                MsgBox("woops")
            Else
                StopTime()
                FrmTimesUp.DefInstance.ShowDialog()
            End If
        End If
    End Sub

    Private Sub TxtCopy_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtCopy.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If Me.TxtCopy.ReadOnly Then
            MsgBox("You can't modify this copy. " & "Use the Copy | New to start a new test " & "or Copy | Mark to remark it.")
        Else 'test in progress or about to start
            If Not allowEdits And IsEditing(KeyAscii) Then 'don't allow backspaces
                Beep()
                KeyAscii = 0 'Cancel the keystroke
            End If
            If KeyAscii = Asc(vbTab) Then
                System.Windows.Forms.SendKeys.Send("     ")
                KeyAscii = 0
            End If
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub Printer_PrintPage(ByVal sender As System.Object, ByVal ev As System.Drawing.Printing.PrintPageEventArgs) Handles Printer.PrintPage
        'Prints data held in streamToPrint
        Dim linesPerPage As Single
        Dim yPos As Single
        Dim count As Integer = 0
        Dim leftMargin As Single = ev.MarginBounds.Left
        Dim topMargin As Single = ev.MarginBounds.Top
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics)

        ' Iterate over the file, printing each line.
        While count < linesPerPage
            line = streamToPrint.ReadLine()
            If line Is Nothing Then
                Exit While
            End If
            yPos = topMargin + count * printFont.GetHeight(ev.Graphics)
            ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin,
                yPos, New StringFormat)
            count += 1
        End While

        ' If more lines exist, print another page.
        If Not (line Is Nothing) Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
        End If
    End Sub
End Class