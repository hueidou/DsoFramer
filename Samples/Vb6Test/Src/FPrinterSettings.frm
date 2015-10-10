VERSION 5.00
Begin VB.Form FPrinterSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Printer Settings"
   ClientHeight    =   2835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6060
   Icon            =   "FPrinterSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3180
      TabIndex        =   13
      Text            =   "C:\test.prn"
      Top             =   1860
      Width           =   2775
   End
   Begin VB.CheckBox chkPrintToFile 
      Caption         =   "Print to File:"
      Height          =   195
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkDisplay 
      Caption         =   "Display Server's Print Dialog"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2340
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox txtCopies 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4380
      TabIndex        =   10
      Text            =   "1"
      Top             =   1140
      Width           =   495
   End
   Begin VB.TextBox txtTo 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Text            =   "1"
      Top             =   780
      Width           =   495
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   255
      Left            =   4380
      TabIndex        =   6
      Text            =   "1"
      Top             =   780
      Width           =   495
   End
   Begin VB.OptionButton optPrintSel 
      Caption         =   "Print Pages"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   780
      Width           =   1515
   End
   Begin VB.OptionButton optPrintSel 
      Caption         =   "Print Entire Document"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   2955
   End
   Begin VB.ListBox lstPrinters 
      Height          =   1815
      Left            =   180
      TabIndex        =   2
      Top             =   420
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4740
      TabIndex        =   1
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Number of Copies:"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "to"
      Height          =   195
      Left            =   4980
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Select the printer to use and the settings to pass to the framer control:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5835
   End
End
Attribute VB_Name = "FPrinterSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PrinterName As String
Public OutputFile As String
Public PrintToFile As Boolean
Public DisplayPrintDialog As Boolean
Public PrintPages As Boolean
Public FromPage As Long
Public ToPage As Long
Public Copies As Long
Public Canceled As Boolean

Private Sub CancelButton_Click()
    Canceled = True
    Me.Hide
End Sub

Private Sub chkDisplay_Click()
    DisplayPrintDialog = chkDisplay.Value
End Sub

Private Sub chkPrintToFile_Click()
    txtFile.Enabled = chkPrintToFile.Value <> 0
    PrintToFile = chkPrintToFile.Value <> 0
End Sub

Private Sub Form_Load()
    Dim oPrinter As Printer
    For Each oPrinter In Printers
        lstPrinters.AddItem oPrinter.DeviceName
    Next oPrinter
    DisplayPrintDialog = True
End Sub

Private Sub lstPrinters_Click()
    If lstPrinters.SelCount > 0 Then
        OKButton.Enabled = True
    End If
End Sub

Private Sub OKButton_Click()

    If lstPrinters.SelCount <> 1 Then Exit Sub
    Canceled = False
    
    PrinterName = lstPrinters.List(lstPrinters.ListIndex)
    Copies = CLng(txtCopies.Text)
    If PrintPages Then
        FromPage = CLng(txtFrom.Text)
        ToPage = CLng(txtTo.Text)
    End If
    If PrintToFile Then
        OutputFile = txtFile.Text
    End If
    Me.Hide
End Sub

Private Sub optPrintSel_Click(Index As Integer)
    txtFrom.Enabled = (Index = 1)
    txtTo.Enabled = (Index = 1)
    PrintPages = (Index = 1)
End Sub
