VERSION 5.00
Begin VB.Form About 
   Caption         =   "Form3"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13290
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   6855
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBack 
      BackColor       =   &H8000000E&
      Caption         =   "Back"
      Height          =   495
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   3015
      Left            =   360
      Picture         =   "About.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   6960
      Picture         =   "About.frx":50A0AF
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBack_Click()
Form1.Visible = True
End Sub

Private Sub Form_Load()
Label1.Caption = "Hey Dear you are using Letest notepad, Developed by Jitendra Kushwaha "
End Sub

