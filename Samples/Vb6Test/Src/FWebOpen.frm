VERSION 5.00
Begin VB.Form FWebOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web URL"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "FWebOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProgID 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   4035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox chkOpenForWrite 
      Caption         =   "Open For Write Access (item in Web Folder)"
      Height          =   255
      Left            =   780
      TabIndex        =   3
      Top             =   900
      Width           =   3735
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Text            =   "http://"
      Top             =   480
      Width           =   5835
   End
   Begin VB.Label Label3 
      Caption         =   "ProgID (for HTML/ASP files):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Provide the URL to the file you want to open (i.e., http://myserver/myfolder/mydoc.doc):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "FWebOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public URL As String
Public OpenWriteAccess As Boolean
Public ProgID As String

Private Sub cmdCancel_Click()
    URL = ""
    Me.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim sURL As String
    sURL = txtURL.Text
    If Len(sURL) < 5 Or LCase(Left(sURL, 4)) <> "http" Then
        txtURL.SelStart = 0
        txtURL.SelLength = Len(sURL)
        Beep
        Exit Sub
    End If
    URL = sURL
    OpenWriteAccess = chkOpenForWrite.Value
    ProgID = txtProgID.Text
    Me.Visible = False
End Sub

Private Sub Form_Load()
    txtURL.SelStart = Len(txtURL.Text)
End Sub
