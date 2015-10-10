VERSION 5.00
Object = "{00460180-9E5E-11D5-B7C8-B8269041DD57}#1.3#0"; "dsoframer.ocx"
Begin VB.Form FMainApplication 
   Caption         =   "VB6 Test Application for DsoFramer Control"
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   StartUpPosition =   3  'Windows Default
   Begin DSOFramer.FramerControl oFramer 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   11245
      BorderColor     =   -2147483632
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      TitlebarColor   =   -2147483635
      TitlebarTextColor=   -2147483634
      BorderStyle     =   1
      Titlebar        =   -1  'True
      Toolbars        =   -1  'True
      Menubar         =   -1  'True
   End
   Begin VB.Label lbCurrentFile 
      Caption         =   "Current File: [None]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   3
      Top             =   480
      Width           =   7335
   End
   Begin VB.Line lnTitle 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   56
      X2              =   552
      Y1              =   28
      Y2              =   28
   End
   Begin VB.Line lnTitle 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      Index           =   0
      X1              =   56
      X2              =   552
      Y1              =   28
      Y2              =   28
   End
   Begin VB.Label lbAppTitle 
      Caption         =   "DsoFramer Test Client"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   120
      Picture         =   "FTestApp.frx":0000
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lbAppVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version:  1.3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6450
      TabIndex        =   2
      Top             =   180
      Width           =   1800
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileWebOpen 
         Caption         =   "&Web Open..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "C&lose"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveCopyAs 
         Caption         =   "Save Cop&y As..."
      End
      Begin VB.Menu mnuFileSaveWeb 
         Caption         =   "Save to We&b..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "P&rint (Default)"
         Index           =   0
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Pri&nt to Target..."
         Index           =   2
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies..."
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit Application"
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show"
      Begin VB.Menu mnuShowCaption 
         Caption         =   "&Caption"
      End
      Begin VB.Menu mnuShowMenubar 
         Caption         =   "&Menubar"
      End
      Begin VB.Menu mnuShowToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuShowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorderStyleOut 
         Caption         =   "&Border Style"
         Begin VB.Menu mnuBorderStyle 
            Caption         =   "&None"
            Index           =   0
         End
         Begin VB.Menu mnuBorderStyle 
            Caption         =   "&Outline (default)"
            Index           =   1
         End
         Begin VB.Menu mnuBorderStyle 
            Caption         =   "3D &Frame"
            Index           =   2
         End
         Begin VB.Menu mnuBorderStyle 
            Caption         =   "3D Frame &Thin"
            Index           =   3
         End
      End
      Begin VB.Menu mnuShowSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowFileMenu 
         Caption         =   "&Disable File Menu Item"
         Begin VB.Menu mnuDisableItem 
            Caption         =   "&New"
            Index           =   0
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "&Open"
            Index           =   1
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "&Close"
            Index           =   2
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "&Save"
            Index           =   3
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "Save &As"
            Index           =   4
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "&Print"
            Index           =   5
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "Page Set&up"
            Index           =   6
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "Propert&ies"
            Index           =   7
         End
         Begin VB.Menu mnuDisableItem 
            Caption         =   "Print Pre&view"
            Index           =   8
         End
      End
      Begin VB.Menu mnuShowSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomCaption 
         Caption         =   "&Custom Caption..."
      End
   End
End
Attribute VB_Name = "FMainApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FMainApplication
'
'  VB6 Test Program for DsoFramer Control Sample KB 311765
'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main Form Events
'
Private Sub Form_Load()
' We'll use our our host name for the application
    oFramer.HostName = "VB6TestApp"
' Setup the default items...
    mnuShowMenubar.Checked = oFramer.Menubar
    mnuShowToolbar.Checked = oFramer.Toolbars
    mnuShowCaption.Checked = oFramer.Titlebar
    mnuBorderStyle(oFramer.BorderStyle).Checked = True
    Me.ScaleMode = 3 ' pixel
    EnableItems False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Check that document is closed before quit...
    If mnuFileClose.Enabled Then
        mnuFileClose_Click
        Cancel = mnuFileClose.Enabled
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
' On resize, we scale the framer control to fit the form...
    If (Me.ScaleWidth > 20) And (Me.ScaleHeight > 60) Then
        oFramer.Move 10, 52, Me.ScaleWidth - 20, Me.ScaleHeight - 60
        lnTitle(0).X2 = Me.ScaleWidth - 8
        lnTitle(1).X2 = Me.ScaleWidth - 7
        If (Me.ScaleWidth > 68) Then
            lbCurrentFile.Width = Me.ScaleWidth - 68
        End If
        If (Me.ScaleWidth > 264) Then
            lbAppVersion.Left = Me.ScaleWidth - 130
        End If
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File Menu Items
'
Private Sub mnuFileNew_Click()
' We just display the default New dialog to the user...
    On Error Resume Next
    oFramer.ShowDialog dsoDialogNew
    If Err.Number Then
        MsgBox "Unable to create new item." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
End Sub

Private Sub mnuFileOpen_Click()
    Dim vbPrompt As VbMsgBoxResult
    On Error Resume Next
    
 ' Just show built-in dialog for local open...
    oFramer.ShowDialog dsoDialogOpen

    If Err.Number Then
        MsgBox "Unable to open document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
    
End Sub

Private Sub mnuFileWebOpen_Click()
    Dim sFile As String, sProgId As String
    Dim fReadOnly As Boolean
    Dim oWeb As FWebOpen
    On Error Resume Next
    
 ' If they are opening from a web, ask for URL and if we should open read-write...
    Set oWeb = New FWebOpen
    oWeb.Show vbModal, Me
    sFile = oWeb.URL
    sProgId = oWeb.ProgID
    fReadOnly = Not oWeb.OpenWriteAccess
    Unload oWeb
    
 ' If they gave a URL, try to open it (with custom progid if given)...
    If Len(sFile) Then
        oFramer.Open sFile, fReadOnly, sProgId
    End If
    
    If Err.Number Then
        MsgBox "Unable to open document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
End Sub

Private Sub mnuFileClose_Click()
    On Error Resume Next
    oFramer.Close
End Sub

Private Sub mnuFileQuit_Click()
    mnuFileClose_Click
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    On Error Resume Next
    oFramer.Save
    If Err.Number Then
        MsgBox "Unable to save document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim vbPrompt As VbMsgBoxResult
    On Error Resume Next

 ' Just show built-in dialog for local save...
    oFramer.ShowDialog dsoDialogSave

    If Err.Number Then
        MsgBox "Unable to save document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    Else
     ' If we saved the file, get the (new) doc path for display...
        lbCurrentFile.Caption = "Current File: " & oFramer.DocumentFullName
    End If
End Sub

Private Sub mnuFileSaveCopyAs_Click()
    On Error Resume Next
    oFramer.ShowDialog dsoDialogSaveCopy
    If Err.Number Then
        MsgBox "Unable to save copy of document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
End Sub

Private Sub mnuFileSaveWeb_Click()
    Dim sFile As String
    Dim vbPrompt As VbMsgBoxResult
    On Error Resume Next
    
' If they are opening from a web, ask for URL and if we should open read-write...
    sFile = InputBox("Type the full URL to the location to save to (e.g., http://server/folder/mydoc.doc):", "Save to URL", "http://")
    If (Len(sFile) > 10) Then
        vbPrompt = MsgBox("If the file exists, do you want to overwrite it?", vbQuestion Or vbYesNo, "Overwrite?")
        oFramer.Save sFile, (vbPrompt = vbYes)
    End If
    
    If Err.Number Then
        MsgBox "Unable to save document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    Else
     ' If we saved the file, get the (new) doc path for display...
        lbCurrentFile.Caption = "Current File: " & oFramer.DocumentFullName
    End If
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    oFramer.ShowDialog dsoDialogPageSetup
    If Err.Number Then
        MsgBox "Unable to show page setup." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
End Sub

Private Sub mnuFilePrintPreview_Click()
    On Error Resume Next
    oFramer.PrintPreview
    If Err.Number Then
        MsgBox "Unable to go into print preview." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    Else
        EnableItems False
    End If
End Sub

Private Sub mnuFilePrint_Click(Index As Integer)
    On Error Resume Next
    Dim oPrintSettings As FPrinterSettings
    Dim sPrinter As String
    Dim iCopies As Integer
    Dim fPrompt As Boolean
    Dim vOutput, vFrom, vTo
    
 ' We will display custom print dialog for "Print to Target"...
    If (Index = 2) Then
        Set oPrintSettings = New FPrinterSettings
        oPrintSettings.Show vbModal, Me
        
     ' Get the specific printer user wants to print to...
        sPrinter = oPrintSettings.PrinterName
        
        iCopies = oPrintSettings.Copies
        fPrompt = oPrintSettings.DisplayPrintDialog
        
        If oPrintSettings.PrintToFile Then
            vOutput = oPrintSettings.OutputFile
        End If
        If oPrintSettings.PrintPages Then
            vFrom = oPrintSettings.FromPage
            vTo = oPrintSettings.ToPage
        End If
        
        Unload oPrintSettings
        Set oPrintSettings = Nothing
        
     ' Now print it (unless no printer name returned, which indicates cancel)...
        If (Len(sPrinter) And (iCopies > 0)) Then
            oFramer.PrintOutEx fPrompt, sPrinter, iCopies, vFrom, vTo, vOutput
        End If
        
    Else
        fPrompt = (Index > 0)
        oFramer.PrintOut fPrompt
    End If
    
    If Err.Number Then
        MsgBox "Unable to print document." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
    
End Sub

Private Sub mnuFileProperties_Click()
    On Error Resume Next
    oFramer.ShowDialog dsoDialogProperties
    If Err.Number Then
        MsgBox "Unable to show properties page." & vbCrLf & _
            "(" & Str(Err.Number) & "): " & Err.Description, _
            vbCritical, "Error"
        Err.Clear
    End If
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Show Menu Items
'
Private Sub mnuShowCaption_Click()
    On Error Resume Next
    oFramer.Titlebar = Not mnuShowCaption.Checked
    mnuShowCaption.Checked = oFramer.Titlebar
End Sub

Private Sub mnuShowMenubar_Click()
    On Error Resume Next
    oFramer.Menubar = Not mnuShowMenubar.Checked
    mnuShowMenubar.Checked = oFramer.Menubar
End Sub

Private Sub mnuShowToolbar_Click()
    On Error Resume Next
    oFramer.Toolbars = Not mnuShowToolbar.Checked
    mnuShowToolbar.Checked = oFramer.Toolbars
End Sub

Private Sub mnuBorderStyle_Click(Index As Integer)
    Dim i As Long, j As Long
    On Error Resume Next
    oFramer.BorderStyle = Index
    j = oFramer.BorderStyle
    For i = 0 To 3
        mnuBorderStyle(i).Checked = False
        If (i = j) Then mnuBorderStyle(i).Checked = True
    Next
End Sub

Private Sub mnuDisableItem_Click(Index As Integer)
    On Error Resume Next
    oFramer.EnableFileCommand(Index) = mnuDisableItem(Index).Checked
    mnuDisableItem(Index).Checked = Not oFramer.EnableFileCommand(Index)
End Sub

Private Sub mnuCustomCaption_Click()
    Dim sCustomCaption As String
    sCustomCaption = InputBox("Provide a custom caption for the framer titlebar:", "Caption")
    If (Len(sCustomCaption)) Then
        oFramer.Caption = sCustomCaption
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Framer Control Events
'

Private Sub oFramer_BeforeDocumentClosed(ByVal Document As Object, Cancel As Boolean)
    Dim vbPrompt As VbMsgBoxResult
    
    'oFramer.ExecOleCommand &H7003, False 'UnLock
    'oFramer.ExecOleCommand &H7003, True 'Lock

    On Error Resume Next
 ' If file is dirty, ask user if they want to save before close...
    If oFramer.IsDirty Then
        vbPrompt = MsgBox("Would you like to save the file before closing it?", vbQuestion Or vbYesNoCancel, "Save Changes?")
        If vbPrompt = vbCancel Then
            Cancel = True
        ElseIf vbPrompt = vbYes Then
         ' If file is read-only or new/unsaved file...
            If oFramer.IsReadOnly Or _
                Len(oFramer.DocumentFullName) = 0 Then
                ' Show SaveAs dialog...
                oFramer.ShowDialog dsoDialogSave
            Else ' Else save with no dialog...
                oFramer.Save
            End If
        End If
    End If
End Sub

Private Sub oFramer_OnFileCommand(ByVal Item As DSOFramer.dsoFileCommandType, Cancel As Boolean)
 ' Here is where you can override Framer Control File menu items. We
 ' will override the Open and SaveAs to prompt user if they want to save to
 ' web as well as get new document name on successful SaveAs...
    If Item = dsoFileOpen Then
        mnuFileOpen_Click ' We'll do our own Open routine...
        Cancel = True ' Cancel default since we handled it
    ElseIf Item = dsoFileSaveAs Then
        mnuFileSaveAs_Click ' We'll do our own SaveAs routine...
        Cancel = True ' Cancel default since we handled it
    End If
End Sub

Private Sub oFramer_OnDocumentOpened(ByVal File As String, ByVal Document As Object)
    ' When item is added/opened, enable items on form...
    EnableItems True
    If Len(File) Then
        lbCurrentFile.Caption = "Current File: " & File
    Else
        lbCurrentFile.Caption = "Current File: Unsaved Document"
    End If
End Sub

Private Sub oFramer_OnDocumentClosed()
    ' When item is closed, disable some items on form...
    EnableItems False
    lbCurrentFile.Caption = "Current File: [None]"
End Sub

Private Sub oFramer_OnPrintPreviewExit()
    ' Re-enable menu items after preview.
    EnableItems True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EnableItems
'
Private Sub EnableItems(fEnable As Boolean)
    mnuFileClose.Enabled = fEnable
    mnuFileSave.Enabled = fEnable
    mnuFileSaveAs.Enabled = fEnable
    mnuFileSaveCopyAs.Enabled = fEnable
    mnuFileSaveWeb.Enabled = fEnable
    mnuFilePageSetup.Enabled = fEnable
    mnuFilePrintPreview.Enabled = fEnable
    mnuFilePrint(0).Enabled = fEnable
    mnuFilePrint(1).Enabled = fEnable
    mnuFilePrint(2).Enabled = fEnable
    mnuFileProperties.Enabled = fEnable
End Sub

