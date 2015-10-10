Public Class FTestApp
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.

    Friend WithEvents panTitlebar As System.Windows.Forms.Panel
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents tbcDocsCont As System.Windows.Forms.TabControl
    Friend WithEvents tbMain As System.Windows.Forms.TabPage
    Friend WithEvents gboxCreate As System.Windows.Forms.GroupBox
    Friend WithEvents gboxOpen As System.Windows.Forms.GroupBox
    Friend WithEvents lblKbTitle As System.Windows.Forms.Label
    Friend WithEvents lblInto As System.Windows.Forms.Label
    Friend WithEvents lblPickNew As System.Windows.Forms.Label
    Friend WithEvents rbtnNewWord As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNewExcel As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNewPPT As System.Windows.Forms.RadioButton
    Friend WithEvents lblOpenFile As System.Windows.Forms.Label
    Friend WithEvents lblWordDocs As System.Windows.Forms.Label
    Friend WithEvents lblPPTDocs As System.Windows.Forms.Label
    Friend WithEvents lblExcelDocs As System.Windows.Forms.Label
    Friend WithEvents OFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnCreateNew As System.Windows.Forms.Button
    Friend WithEvents btnOpenFile As System.Windows.Forms.Button
    Friend WithEvents tbDoc1 As System.Windows.Forms.TabPage
    Friend WithEvents tbDoc2 As System.Windows.Forms.TabPage
    Friend WithEvents tbDoc3 As System.Windows.Forms.TabPage
    Friend WithEvents tbDoc4 As System.Windows.Forms.TabPage
    Friend WithEvents axFramer1 As AxDSOFramer.AxFramerControl
    Friend WithEvents axFramer2 As AxDSOFramer.AxFramerControl
    Friend WithEvents axFramer3 As AxDSOFramer.AxFramerControl
    Friend WithEvents axFramer4 As AxDSOFramer.AxFramerControl

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FTestApp))
        Me.panTitlebar = New System.Windows.Forms.Panel
        Me.lblVersion = New System.Windows.Forms.Label
        Me.lblTitle = New System.Windows.Forms.Label
        Me.picLogo = New System.Windows.Forms.PictureBox
        Me.tbcDocsCont = New System.Windows.Forms.TabControl
        Me.tbMain = New System.Windows.Forms.TabPage
        Me.gboxOpen = New System.Windows.Forms.GroupBox
        Me.lblPPTDocs = New System.Windows.Forms.Label
        Me.lblExcelDocs = New System.Windows.Forms.Label
        Me.lblWordDocs = New System.Windows.Forms.Label
        Me.btnOpenFile = New System.Windows.Forms.Button
        Me.lblOpenFile = New System.Windows.Forms.Label
        Me.gboxCreate = New System.Windows.Forms.GroupBox
        Me.rbtnNewPPT = New System.Windows.Forms.RadioButton
        Me.rbtnNewExcel = New System.Windows.Forms.RadioButton
        Me.rbtnNewWord = New System.Windows.Forms.RadioButton
        Me.lblPickNew = New System.Windows.Forms.Label
        Me.btnCreateNew = New System.Windows.Forms.Button
        Me.lblInto = New System.Windows.Forms.Label
        Me.lblKbTitle = New System.Windows.Forms.Label
        Me.tbDoc4 = New System.Windows.Forms.TabPage
        Me.axFramer4 = New AxDSOFramer.AxFramerControl
        Me.tbDoc1 = New System.Windows.Forms.TabPage
        Me.axFramer1 = New AxDSOFramer.AxFramerControl
        Me.tbDoc2 = New System.Windows.Forms.TabPage
        Me.axFramer2 = New AxDSOFramer.AxFramerControl
        Me.tbDoc3 = New System.Windows.Forms.TabPage
        Me.axFramer3 = New AxDSOFramer.AxFramerControl
        Me.OFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.panTitlebar.SuspendLayout()
        Me.tbcDocsCont.SuspendLayout()
        Me.tbMain.SuspendLayout()
        Me.gboxOpen.SuspendLayout()
        Me.gboxCreate.SuspendLayout()
        Me.tbDoc4.SuspendLayout()
        CType(Me.axFramer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbDoc1.SuspendLayout()
        CType(Me.axFramer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbDoc2.SuspendLayout()
        CType(Me.axFramer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbDoc3.SuspendLayout()
        CType(Me.axFramer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'panTitlebar
        '
        Me.panTitlebar.BackColor = System.Drawing.Color.Gold
        Me.panTitlebar.Controls.Add(Me.lblVersion)
        Me.panTitlebar.Controls.Add(Me.lblTitle)
        Me.panTitlebar.Dock = System.Windows.Forms.DockStyle.Top
        Me.panTitlebar.Location = New System.Drawing.Point(0, 0)
        Me.panTitlebar.Name = "panTitlebar"
        Me.panTitlebar.Size = New System.Drawing.Size(572, 56)
        Me.panTitlebar.TabIndex = 1
        '
        'lblVersion
        '
        Me.lblVersion.Location = New System.Drawing.Point(168, 32)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(152, 16)
        Me.lblVersion.TabIndex = 1
        Me.lblVersion.Text = "Version 1.3.1323"
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(160, 8)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(272, 24)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "DsoFramer VB.NET Sample"
        '
        'picLogo
        '
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(0, 0)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(160, 56)
        Me.picLogo.TabIndex = 0
        Me.picLogo.TabStop = False
        '
        'tbcDocsCont
        '
        Me.tbcDocsCont.Controls.Add(Me.tbMain)
        Me.tbcDocsCont.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcDocsCont.Location = New System.Drawing.Point(0, 56)
        Me.tbcDocsCont.Name = "tbcDocsCont"
        Me.tbcDocsCont.SelectedIndex = 0
        Me.tbcDocsCont.Size = New System.Drawing.Size(572, 454)
        Me.tbcDocsCont.TabIndex = 2
        '
        'tbMain
        '
        Me.tbMain.Controls.Add(Me.gboxOpen)
        Me.tbMain.Controls.Add(Me.gboxCreate)
        Me.tbMain.Controls.Add(Me.lblInto)
        Me.tbMain.Controls.Add(Me.lblKbTitle)
        Me.tbMain.Location = New System.Drawing.Point(4, 22)
        Me.tbMain.Name = "tbMain"
        Me.tbMain.Size = New System.Drawing.Size(564, 428)
        Me.tbMain.TabIndex = 1
        Me.tbMain.Text = "Main"
        '
        'gboxOpen
        '
        Me.gboxOpen.Controls.Add(Me.lblPPTDocs)
        Me.gboxOpen.Controls.Add(Me.lblExcelDocs)
        Me.gboxOpen.Controls.Add(Me.lblWordDocs)
        Me.gboxOpen.Controls.Add(Me.btnOpenFile)
        Me.gboxOpen.Controls.Add(Me.lblOpenFile)
        Me.gboxOpen.Location = New System.Drawing.Point(240, 112)
        Me.gboxOpen.Name = "gboxOpen"
        Me.gboxOpen.Size = New System.Drawing.Size(312, 208)
        Me.gboxOpen.TabIndex = 4
        Me.gboxOpen.TabStop = False
        Me.gboxOpen.Text = "Open Existing Document"
        '
        'lblPPTDocs
        '
        Me.lblPPTDocs.Location = New System.Drawing.Point(32, 112)
        Me.lblPPTDocs.Name = "lblPPTDocs"
        Me.lblPPTDocs.Size = New System.Drawing.Size(256, 24)
        Me.lblPPTDocs.TabIndex = 4
        Me.lblPPTDocs.Text = "- PowerPoint Presentations (*.ppt, *.pptx, *.pptm)"
        '
        'lblExcelDocs
        '
        Me.lblExcelDocs.Location = New System.Drawing.Point(32, 88)
        Me.lblExcelDocs.Name = "lblExcelDocs"
        Me.lblExcelDocs.Size = New System.Drawing.Size(256, 24)
        Me.lblExcelDocs.TabIndex = 3
        Me.lblExcelDocs.Text = "- Excel Workbooks (*.xls, *.xlsx, *.xlsm, *.xlsb)"
        '
        'lblWordDocs
        '
        Me.lblWordDocs.Location = New System.Drawing.Point(32, 64)
        Me.lblWordDocs.Name = "lblWordDocs"
        Me.lblWordDocs.Size = New System.Drawing.Size(256, 24)
        Me.lblWordDocs.TabIndex = 2
        Me.lblWordDocs.Text = "- Word Documents (*.doc, *.docx, *.docm)"
        '
        'btnOpenFile
        '
        Me.btnOpenFile.Location = New System.Drawing.Point(80, 160)
        Me.btnOpenFile.Name = "btnOpenFile"
        Me.btnOpenFile.Size = New System.Drawing.Size(152, 32)
        Me.btnOpenFile.TabIndex = 1
        Me.btnOpenFile.Text = "Open File"
        '
        'lblOpenFile
        '
        Me.lblOpenFile.Location = New System.Drawing.Point(16, 24)
        Me.lblOpenFile.Name = "lblOpenFile"
        Me.lblOpenFile.Size = New System.Drawing.Size(280, 32)
        Me.lblOpenFile.TabIndex = 0
        Me.lblOpenFile.Text = "Click this button to find and open a file from the local drive.  By default, the " & _
        "dialog will display: "
        '
        'gboxCreate
        '
        Me.gboxCreate.Controls.Add(Me.rbtnNewPPT)
        Me.gboxCreate.Controls.Add(Me.rbtnNewExcel)
        Me.gboxCreate.Controls.Add(Me.rbtnNewWord)
        Me.gboxCreate.Controls.Add(Me.lblPickNew)
        Me.gboxCreate.Controls.Add(Me.btnCreateNew)
        Me.gboxCreate.Location = New System.Drawing.Point(8, 112)
        Me.gboxCreate.Name = "gboxCreate"
        Me.gboxCreate.Size = New System.Drawing.Size(224, 208)
        Me.gboxCreate.TabIndex = 3
        Me.gboxCreate.TabStop = False
        Me.gboxCreate.Text = "Create New Document"
        '
        'rbtnNewPPT
        '
        Me.rbtnNewPPT.Location = New System.Drawing.Point(16, 112)
        Me.rbtnNewPPT.Name = "rbtnNewPPT"
        Me.rbtnNewPPT.Size = New System.Drawing.Size(200, 24)
        Me.rbtnNewPPT.TabIndex = 6
        Me.rbtnNewPPT.Text = "Microsoft PowerPoint Presentation"
        '
        'rbtnNewExcel
        '
        Me.rbtnNewExcel.Location = New System.Drawing.Point(16, 80)
        Me.rbtnNewExcel.Name = "rbtnNewExcel"
        Me.rbtnNewExcel.Size = New System.Drawing.Size(200, 24)
        Me.rbtnNewExcel.TabIndex = 5
        Me.rbtnNewExcel.Text = "Microsoft Excel Workbook"
        '
        'rbtnNewWord
        '
        Me.rbtnNewWord.Checked = True
        Me.rbtnNewWord.Location = New System.Drawing.Point(16, 48)
        Me.rbtnNewWord.Name = "rbtnNewWord"
        Me.rbtnNewWord.Size = New System.Drawing.Size(200, 24)
        Me.rbtnNewWord.TabIndex = 4
        Me.rbtnNewWord.TabStop = True
        Me.rbtnNewWord.Text = "Microsoft Word Document"
        '
        'lblPickNew
        '
        Me.lblPickNew.Location = New System.Drawing.Point(8, 24)
        Me.lblPickNew.Name = "lblPickNew"
        Me.lblPickNew.Size = New System.Drawing.Size(208, 24)
        Me.lblPickNew.TabIndex = 3
        Me.lblPickNew.Text = "Pick a new document type to open."
        '
        'btnCreateNew
        '
        Me.btnCreateNew.Location = New System.Drawing.Point(24, 160)
        Me.btnCreateNew.Name = "btnCreateNew"
        Me.btnCreateNew.Size = New System.Drawing.Size(152, 32)
        Me.btnCreateNew.TabIndex = 2
        Me.btnCreateNew.Text = "Create"
        '
        'lblInto
        '
        Me.lblInto.Location = New System.Drawing.Point(8, 40)
        Me.lblInto.Name = "lblInto"
        Me.lblInto.Size = New System.Drawing.Size(544, 56)
        Me.lblInto.TabIndex = 1
        Me.lblInto.Text = "This sample host is used to test having more than one instance of the DsoFramer c" & _
        "ontrol on a WinForm application window. Click the button to add or open a new do" & _
        "cument.  The document will open in a new tab.  This sample only supports opening" & _
        " 4 documents at any one time."
        '
        'lblKbTitle
        '
        Me.lblKbTitle.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKbTitle.Location = New System.Drawing.Point(8, 8)
        Me.lblKbTitle.Name = "lblKbTitle"
        Me.lblKbTitle.Size = New System.Drawing.Size(544, 32)
        Me.lblKbTitle.TabIndex = 0
        Me.lblKbTitle.Text = "KB311765 : DsoFramer ActiveX Control WinForm Host"
        '
        'tbDoc4
        '
        Me.tbDoc4.Controls.Add(Me.axFramer4)
        Me.tbDoc4.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc4.Name = "tbDoc4"
        Me.tbDoc4.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc4.TabIndex = 0
        Me.tbDoc4.Tag = 4
        Me.tbDoc4.Text = "Document4"
        '
        'axFramer4
        '
        Me.axFramer4.ContainingControl = Me
        Me.axFramer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.axFramer4.Enabled = True
        Me.axFramer4.Location = New System.Drawing.Point(0, 0)
        Me.axFramer4.Name = "axFramer4"
        Me.axFramer4.OcxState = CType(resources.GetObject("axFramer4.OcxState"), System.Windows.Forms.AxHost.State)
        Me.axFramer4.Size = New System.Drawing.Size(564, 451)
        Me.axFramer4.TabIndex = 0
        '
        'tbDoc1
        '
        Me.tbDoc1.Controls.Add(Me.axFramer1)
        Me.tbDoc1.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc1.Name = "tbDoc1"
        Me.tbDoc1.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc1.TabIndex = 0
        Me.tbDoc1.Tag = 1
        Me.tbDoc1.Text = "Document1"
        '
        'axFramer1
        '
        Me.axFramer1.ContainingControl = Me
        Me.axFramer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.axFramer1.Enabled = True
        Me.axFramer1.Location = New System.Drawing.Point(0, 0)
        Me.axFramer1.Name = "axFramer1"
        Me.axFramer1.OcxState = CType(resources.GetObject("axFramer1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.axFramer1.Size = New System.Drawing.Size(564, 451)
        Me.axFramer1.TabIndex = 0
        '
        'tbDoc2
        '
        Me.tbDoc2.Controls.Add(Me.axFramer2)
        Me.tbDoc2.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc2.Name = "tbDoc2"
        Me.tbDoc2.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc2.TabIndex = 0
        Me.tbDoc2.Tag = 2
        Me.tbDoc2.Text = "Document2"
        '
        'axFramer2
        '
        Me.axFramer2.ContainingControl = Me
        Me.axFramer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.axFramer2.Enabled = True
        Me.axFramer2.Location = New System.Drawing.Point(0, 0)
        Me.axFramer2.Name = "axFramer2"
        Me.axFramer2.OcxState = CType(resources.GetObject("axFramer2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.axFramer2.Size = New System.Drawing.Size(564, 451)
        Me.axFramer2.TabIndex = 0
        '
        'tbDoc3
        '
        Me.tbDoc3.Controls.Add(Me.axFramer3)
        Me.tbDoc3.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc3.Name = "tbDoc3"
        Me.tbDoc3.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc3.TabIndex = 0
        Me.tbDoc3.Tag = 3
        Me.tbDoc3.Text = "Document3"
        '
        'axFramer3
        '
        Me.axFramer3.ContainingControl = Me
        Me.axFramer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.axFramer3.Enabled = True
        Me.axFramer3.Location = New System.Drawing.Point(0, 0)
        Me.axFramer3.Name = "axFramer3"
        Me.axFramer3.OcxState = CType(resources.GetObject("axFramer3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.axFramer3.Size = New System.Drawing.Size(564, 451)
        Me.axFramer3.TabIndex = 0
        '
        'OFileDialog
        '
        Me.OFileDialog.Filter = "Microsoft Office Files|*.doc;*.docx;*.docm;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ppt;*.ppt" & _
        "x;*.pptm|All Files|*.*"
        '
        'FTestApp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(572, 510)
        Me.Controls.Add(Me.tbcDocsCont)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.panTitlebar)
        Me.MinimumSize = New System.Drawing.Size(580, 500)
        Me.Name = "FTestApp"
        Me.Text = "DsoFramer VB7 Sample Application"
        Me.panTitlebar.ResumeLayout(False)
        Me.tbcDocsCont.ResumeLayout(False)
        Me.tbMain.ResumeLayout(False)
        Me.gboxOpen.ResumeLayout(False)
        Me.gboxCreate.ResumeLayout(False)
        Me.tbDoc4.ResumeLayout(False)
        CType(Me.axFramer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbDoc1.ResumeLayout(False)
        CType(Me.axFramer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbDoc2.ResumeLayout(False)
        CType(Me.axFramer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbDoc3.ResumeLayout(False)
        CType(Me.axFramer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim m_iCounter As Integer
    Dim m_bFilesOpen(4) As Boolean

    ' ============================================================================
    '  GetOpenSlot - Helper Function
    '
    '   Returns free slot for tab page and framer control. You could implement a
    '   a dynamic control array and support up to the max number of framer controls,
    '   but we are keeping this simple and limiting to just 4 open docs at a time.
    '
    ' ============================================================================
    Private Function GetOpenSlot() As Integer
        Dim i As Integer
        For i = 0 To 3
            If m_bFilesOpen(i) = False Then
                Return i + 1
            End If
        Next
        Return 0
    End Function

    ' ============================================================================
    '  GetTabPageFromIdx - Helper Function
    '
    '   Returns tab page control for the given slot.
    '
    ' ============================================================================
    Private Function GetTabPageFromIdx(ByVal idx As Integer) As TabPage
        Select Case idx
            Case 1
                Return Me.tbDoc1
            Case 2
                Return Me.tbDoc2
            Case 3
                Return Me.tbDoc3
            Case 4
                Return Me.tbDoc4
            Case Else
                Throw New Exception("Invalid Index")
        End Select
    End Function

    ' ============================================================================
    '  GetFramerCtlFromIdx - Helper Function
    '
    '   Returns DsoFramer control for the given slot.
    '
    ' ============================================================================
    Private Function GetFramerCtlFromIdx(ByVal idx As Integer) As AxDSOFramer.AxFramerControl
        Select Case idx
            Case 1
                Return Me.axFramer1
            Case 2
                Return Me.axFramer2
            Case 3
                Return Me.axFramer3
            Case 4
                Return Me.axFramer4
            Case Else
                Throw New Exception("Invalid Index")
        End Select
    End Function

    ' ============================================================================
    '  AddTabAndActivate - Helper Subroutine
    '
    '   Adds tab page to the TabControl on our main form, and sets a default name
    '   for the tab (assuming this is new document). Then selects the tab to activate.
    '
    ' ============================================================================
    Private Sub AddTabAndActivate(ByVal idx As Integer)
        Dim tab As TabPage
        Dim axControl As AxDSOFramer.AxFramerControl
        m_iCounter = m_iCounter + 1

        ' Get the tab control and add it to the collection...
        tab = GetTabPageFromIdx(idx)
        tab.Text = "New Document " & m_iCounter
        Me.tbcDocsCont.Controls.Add(tab)
        Me.tbcDocsCont.SelectedTab = tab

        ' Get the Framer control and set some default properties since
        ' we don't want user to open/new without going to our main tab.
        axControl = GetFramerCtlFromIdx(idx)
        axControl.set_EnableFileCommand(DSOFramer.dsoFileCommandType.dsoFileNew, False)
        axControl.set_EnableFileCommand(DSOFramer.dsoFileCommandType.dsoFileOpen, False)
        ' We need to explicitly enable the event sinks. Due to a strange bug in .NET
        ' when the control is sited to tab page and not the main form, it is told to 
        ' freeze events (IOleControl) but never told to unfreeze. So events don't get
        ' fired correctly from tab strip. This sets the flag to re-enable the events.
        axControl.EventsEnabled = True
        axControl.Select()

        ' Since we just activated, mark this slot as occupied...
        m_bFilesOpen(idx - 1) = True
    End Sub

    ' ============================================================================
    '  RemoveTabAndSelectMain - Helper Subroutine
    '
    '   Removes the tab from the collection when the document is closed by user.
    '
    ' ============================================================================
    Private Sub RemoveTabAndSelectMain(ByVal idx As Integer)
        Me.tbcDocsCont.Controls.Remove(GetTabPageFromIdx(idx))
        Me.tbcDocsCont.SelectedIndex = 1
        m_bFilesOpen(idx - 1) = False
    End Sub


    ' ============================================================================
    '  btnOpenFile_Click - File Open Button Click Handler
    '
    '   Opens the file in DsoFramer control and appends it to a new tab in the
    '   tab strip control. We allow user to pick the file using File Open dialog.
    '
    ' ============================================================================
    Private Sub btnOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
        ' Ensure we have a free slot...
        Dim idx As Integer = GetOpenSlot()
        If idx = 0 Then
            MessageBox.Show("You can only have four documents open at a time. Close will need to close one to continue.", "Open File", _
                MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        ' Temporarily disable the buttons so we don't re-enter
        btnCreateNew.Enabled = False
        btnOpenFile.Enabled = False

        ' Ask user for the file to open...
        Dim r As DialogResult = OFileDialog.ShowDialog()
        If r = DialogResult.OK Then

            ' Add the tab to the collection and switch to the tab...
            AddTabAndActivate(idx)

            ' Get the framer control for that slot...
            Dim ctl As AxDSOFramer.AxFramerControl
            ctl = GetFramerCtlFromIdx(idx)
            Try
                ' Ask it to open the file...
                ctl.Open(OFileDialog.FileName)

                ' Get the tab page and change the title to the name of the file opened...
                Dim tp As TabPage
                tp = GetTabPageFromIdx(idx)
                tp.Text = ctl.DocumentName

            Catch ex As Exception
                ' Show the error to user...
                MessageBox.Show("Unable to open the file. " & ex.Message, "File Open", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' If we fail, remove the tab and take us back to the main page
                RemoveTabAndSelectMain(idx)
            End Try
        End If

        btnCreateNew.Enabled = True
        btnOpenFile.Enabled = True
    End Sub

    ' ============================================================================
    '  btnCreateNew_Click - Create New Button Click Handler
    '
    '   Creates new blank document of one of the types selected in Radio buttons.
    '
    ' ============================================================================
    Private Sub btnCreateNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateNew.Click
        ' Ensure we have a free slot...
        Dim idx As Integer = GetOpenSlot()
        If idx = 0 Then
            MessageBox.Show("You can only have four documents open at a time. Close will need to close one to continue.", "Create New", _
                MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        ' Temporarily disable the buttons so we don't re-enter
        btnCreateNew.Enabled = False
        btnOpenFile.Enabled = False

        ' Pick the ProgID from the Radio button selected...
        Dim sProgID As String
        If rbtnNewWord.Checked Then
            sProgID = "Word.Document.8"
        ElseIf rbtnNewExcel.Checked Then
            sProgID = "Excel.Sheet.8"
        Else
            sProgID = "PowerPoint.Show.8"
        End If

        ' Add the tab page to tab and make it visible...
        AddTabAndActivate(idx)

        ' Get the framer control for that free slot...
        Dim ctl As AxDSOFramer.AxFramerControl
        ctl = GetFramerCtlFromIdx(idx)

        Try
            ' Ask it to create the new object...
            ctl.CreateNew(sProgID)

        Catch ex As Exception
            ' Show error to user...
            MessageBox.Show("Unable to create the new document. " & ex.Message, "Create New", _
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' If we fail, remove the tab and take us back to the main page
            RemoveTabAndSelectMain(idx)
        End Try

        btnCreateNew.Enabled = True
        btnOpenFile.Enabled = True
    End Sub

    ' ============================================================================
    '  tbcDocsCont_SelectedIndexChanged 
    '
    '   When switching tabs, activate the framer control associated with that tab.
    '
    ' ============================================================================
    Private Sub tbcDocsCont_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbcDocsCont.SelectedIndexChanged
        Dim tab As TabPage = tbcDocsCont.SelectedTab
        Dim framer As AxDSOFramer.AxFramerControl
        Dim idx As Integer = tab.Tag
        If (idx >= 1 And idx <= 4) Then
            framer = GetFramerCtlFromIdx(idx)
            framer.Activate()
        End If
    End Sub

    ' ============================================================================
    '  axFramerX_OnDocumentClosed 
    '
    '  Control event handlers to remove the tab when the document(s) are closed.
    '
    ' ============================================================================
    Private Sub axFramer1_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer1.OnDocumentClosed
        RemoveTabAndSelectMain(1)
    End Sub

    Private Sub axFramer2_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer2.OnDocumentClosed
        RemoveTabAndSelectMain(2)
    End Sub

    Private Sub axFramer3_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer3.OnDocumentClosed
        RemoveTabAndSelectMain(3)
    End Sub

    Private Sub axFramer4_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer4.OnDocumentClosed
        RemoveTabAndSelectMain(4)
    End Sub

    ' ============================================================================
    '  axFramerX_OnSaveCompleted 
    '
    '  Control event handlers to rename tabs if document is saved by a new name.
    '
    ' ============================================================================
    Private Sub axFramer1_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer1.OnSaveCompleted
        Dim s As String = e.docName
        If s.Length > 1 Then tbDoc1.Text = e.docName
    End Sub

    Private Sub axFramer2_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer2.OnSaveCompleted
        Dim s As String = e.docName
        If s.Length > 1 Then tbDoc2.Text = e.docName
    End Sub

    Private Sub axFramer3_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer3.OnSaveCompleted
        Dim s As String = e.docName
        If s.Length > 1 Then tbDoc3.Text = e.docName
    End Sub

    Private Sub axFramer4_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer4.OnSaveCompleted
        Dim s As String = e.docName
        If s.Length > 1 Then tbDoc4.Text = e.docName
    End Sub

End Class

