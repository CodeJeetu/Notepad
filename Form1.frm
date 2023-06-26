VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "Notepad"
   ClientHeight    =   9270
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   17025
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dilag1 
      Left            =   9480
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text Files (*.txt)|*.txt|Graphic Files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|All Files (*.*)|*.*"
   End
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   16935
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnFormet 
      Caption         =   "Format"
      Begin VB.Menu mnWord 
         Caption         =   "Word worp"
      End
      Begin VB.Menu mnFont 
         Caption         =   "Font"
      End
   End
   Begin VB.Menu mnView 
      Caption         =   "View"
      Begin VB.Menu mnZoom 
         Caption         =   "Zoom "
      End
      Begin VB.Menu mnZoomOut 
         Caption         =   "Zoom Out"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "Help"
      Begin VB.Menu MnVhelp 
         Caption         =   "View help"
      End
      Begin VB.Menu mnfeedback 
         Caption         =   "Send Feedback"
      End
      Begin VB.Menu mnNotepad 
         Caption         =   "About Notepad"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnExit_Click()
End
End Sub

Private Sub mnFont_Click()
Form2.Visible = True
End Sub

Private Sub mnNew_Click()
Dim ret As Integer
If Text1.Text <> "" Then
ret = MsgBox("Do you want to save this file", vbYesNoCancel, "Attention")
End If
 If ret = 6 Then
  dilag1.ShowSave
  Text1.Text = " "
 ElseIf ret = 7 Then
  Text1.Text = " "
  ElseIf ret = 2 Then
  Text1.Text = Text1.Text
Else
Text1.Text = ""

 End If
End Sub

Private Sub mnNotepad_Click()
About.Visible = True
Form1.Visible = False
End Sub

Private Sub mnOpen_Click()

 dilag1.CancelError = True
 On Error GoTo ErrHandler

 ' Set filters
 dilag1.Filter = "Inf Files (*.Inf)|*Inf|" & _
 "Batch Files (*.bat)|*.bat|" & _
 "Modules (*.bas)|*.bas|" & _
 "HTML Files (*.Html)|*.Html|" & _
 "Javascript Files (*.js)|*.js|" & _
 "Ini Files (*.Ini)|*.Ini| " & _
 "Text Files (*.txt)|*.txt| " & _
 "Rtf Files (*.rtf)|*.rtf| " & _
 "File Opener Files (*.fop)|*.fop| " & _
 "Log Files (*.log)|*.log| " & _
 "H Files (*.h)|*.h| " & _
 "C Files (*.c)|*.c| " & _
 "All Files (*.*)|*.*| "
 
                                    ' Specify default filter to *.txt
 dilag1.FilterIndex = 7

                 ' Display the Open dialog box, and
                ' save the selected file in the
                ' variable strFileName
 dilag1.ShowOpen
FileName = dilag1.FileName

                        ' Read selected file into the Text Box.
Open FileName For Input As #1
 Text1.Text = Input(LOF(1), 1)
 Close #1
Form1.Caption = "File Opener - " & dilag1.FileTitle
 Exit Sub

ErrHandler:
 'User pressed the Cancel button

End Sub

Private Sub MnSave_Click()

            ' Set CancelError is True
            ' If user presses the cancel button,
            ' Common Dialog Control will generate
            ' a runtime error that can be caught

dilag1.CancelError = True
 On Error GoTo ErrHandler

            ' Set filters
 dilag1.Filter = "Inf Files (*.Inf)|*Inf|" & _
 "Batch File (*.bat)|*.bat|" & _
 "Module (*.bas)|*.bas|" & _
 "HTML File (*.Html)|*.Html|" & _
 "Javascript File (*.js)|*.js|" & _
 "Ini File (*.Ini)|*.Ini| " & _
 "Text File (*.txt)|*.txt| " & _
 "Rtf File (*.rtf)|*.rtf| " & _
 "Log Files (*.log)|*.log| " & _
 "H Files (*.h)|*.h| " & _
 "C Files (*.c)|*.c| " & _
 "File Opener File (*.fop)|*.fop| "
              ' Specify default filter to *.txt
 dilag1.FilterIndex = 7

            ' Display the SaveAs dialog box, and
            '  save the selected file in the
            ' variable strFileName
 dilag1.ShowSave
FileName = dilag1.FileName

            ' Save the Text Box content into the selected file.
 Open FileName For Output As #1
 Print #1, Text1.Text
 Close #1
 Form1.Caption = "File Opener - " & dilag1.FileTitle
 Exit Sub

ErrHandler:
            'User pressed the Cancel button

End Sub

Private Sub mnZoom_Click()
Text1.Height = Text1.Height + 100
Text1.Width = Text1.Width + 100
End Sub

Private Sub mnZoomOut_Click()
Text1.Height = Text1.Height - 100
Text1.Width = Text1.Width - 100
End Sub

Private Sub text1_Change()
Text1.FontName = Text1.FontName

End Sub
