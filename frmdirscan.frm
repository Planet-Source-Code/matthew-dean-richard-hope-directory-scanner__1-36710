VERSION 5.00
Begin VB.Form frmdirscan 
   Caption         =   "Directory Scanner"
   ClientHeight    =   375
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdscan 
      Caption         =   "&Scan Directory"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuchangelog 
         Caption         =   "&Change location of output file"
      End
      Begin VB.Menu mnubreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnufielexit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmdirscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' force declaration of variables
Private Sub cmdscan_Click()
  Dim dialog As BROWSEINFO ' create all browseinfo variables
  Dim showit As Long ' create all long variables
  Dim position As Integer ' create all integer variables
  Dim directory As String, path As String ' create all string variables
  
  dialog.hOwner = Me.hWnd ' this centres the dialog box on the screen
  dialog.lpszTitle = "Select folder to scan" ' this sets the title of the dialog box
  showit = SHBrowseForFolder(dialog)  ' shows the dialogue box
  path = Space(512) 'sets the maximum characters of the path
  SHGetPathFromIDList ByVal showit, ByVal path ' this gets the selected path
  position = InStr(path, Chr(0)) 'finds how long the path is before it hits ascii code 0
  directory = Left(path, position - 1) 'removes the path from the path string
  If Not Right(directory, 1) = "\" Then ' if there is not a \ at the end of the dir name
    directory = directory + "\" ' add one to it
  End If
  Open filename For Output As #1 ' opens the output file
    Call scandir(directory) ' calls a sub to scan and output the results of the selected dir
    Call scansubs(directory) ' calls a sub to start scanning the subdirs
  Close #1 ' closes the output file
  Shell "notepad " + filename, vbNormalFocus ' opens the output file with notepad
End Sub
Sub scansubs(directory)
  Dim temp As String, pdir As String, myname As String, dir2 As String ' create all string variables
  pdir = directory ' set the pdir variable to hold the same data as the directory variable
  myname = Dir(directory, vbDirectory) ' scan the selected dir for subdirs
  Do While Not myname = "" ' do untill the myname variable holds no data
    ' Ignore the current directory and the encompassing directory.
    If Not myname = "." And Not myname = ".." Then
       ' check to ensure the result is a dir
      If (GetAttr(directory + myname) And vbDirectory) = vbDirectory Then
        dir2 = directory + myname ' set the dir2 variable to hold the directory veriable and add on to it the new dir found
        Call scandir(dir2 + "\") ' call the scandir sub with the path of the new dir
        If Not (dir2 + "\") = directory Then ' if the new dir is not the same as the current one
          Call scansubs(dir2 + "\") ' call this sub again using the new dir
        End If
      End If
    End If
    ' this section brings the directory back to its current place
    temp = Dir(pdir, vbDirectory)
    Do Until temp = myname
      temp = Dir
    Loop
    ' find the next dir
    myname = Dir
  Loop ' repeat the loop
End Sub
Sub scandir(directory)
  Dim dirs As Boolean ' creats all boolean variables
  Dim myname As String, myfile As String ' create all string variables
  dirs = False ' sets the value of dirs to false
  ' this section writes the intro to the dir scan
  Print #1, "File contents of " + directory
  Print #1, ""
  Print #1, ""
  ' this section is the header to the directorys section of the output file
  Print #1, "Directories:"
  Print #1, ""
  myname = Dir(directory, vbDirectory)  ' finds the first entry in the seclected dir
  Do While Not myname = ""   ' loops until the myname variable holds nothing
     ' Ignore the current directory and the encompassing directory.
     If Not myname = "." And Not myname = ".." Then
        ' check to ensure the result is a dir
        On Error GoTo error ' if there is an error goto the error lable
        If (GetAttr(directory & myname) And vbDirectory) = vbDirectory Then
           Print #1, myname  ' output the result to the file only if it is a dir
           dirs = True ' set the dirs variable to true
        End If
     End If
     myname = Dir   ' Get next entry.
  Loop ' repeat the loop
  If dirs = False Then ' if the dirs variable is set to false
    Print #1, "There are no sub directories within this directory" ' output that there are no subdirectories
  End If
  ' this section is the geader to the file section of the output file
  Print #1, ""
  Print #1, "Files:"
  Print #1, ""
  myfile = Dir(directory + "*.*") ' look for all the files in the directory
  If myfile = "" Then ' if the myname variable is empty
    Print #1, "There are no files in this directory" ' output that there are no files
    Print #1,
  Else ' if the myname variable does hold some data
    Do Until myfile = "" ' loop until it does not hold any data
      Print #1, myfile ' output the file name to the output file
      myfile = Dir ' find the next file
    Loop ' repeat the loop
    Print #1, ""
  End If
  Print #1, "----------------------------------------------------------------------"
error: ' if there is an error the sub will finish
End Sub

Private Sub Form_Load()
  Call filein("N:\desktop\log.txt") ' call the sub in the modual and pass the name of the output file
                                    ' so it can be stored in the moduals global variable
End Sub

Private Sub mnuchangelog_Click()
  frmoutput.Show vbModal ' show the form to change the output file
End Sub

Private Sub mnufielexit_Click()
  ' prompt the user, if they want to exit end the program
  If MsgBox("Are you sure you want to quit", vbQuestion + vbYesNo, "Quit") = vbYes Then End
End Sub
