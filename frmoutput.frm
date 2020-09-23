VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmoutput 
   Caption         =   "Output File Location"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Save and return"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Back"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "&Change Path"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtpath 
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmoutput.frx":0000
      Top             =   720
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog dlgpath 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbloutput 
      Caption         =   "Location of Output File"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmoutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchange_Click()
  dlgpath.DialogTitle = "Output File"
  dlgpath.Filter = "Text Files (*.txt)|*.txt"
  dlgpath.ShowSave
  If Not dlgpath.filename = "" And Not dlgpath.filename = filename Then
    txtpath.Text = dlgpath.filename
  End If
End Sub

Private Sub Command1_Click()
  Me.Hide
End Sub

Private Sub Command2_Click()
  If txtpath = filename Then
    MsgBox "You have not made any alterations to the output file, please try again", vbOKOnly + vbInformation, "No change"
  Else
    If MsgBox("Are you sure you want to replace the currrent output file, " + filename + ", with the new one, " + txtpath.Text, vbQuestion + vbYesNo, "Replace") = vbYes Then
      Call filein(txtpath.Text)
      Me.Hide
    End If
  End If
End Sub

Private Sub Form_Load()
  txtpath.Text = filename
End Sub
