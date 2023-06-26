VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6075
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      ItemData        =   "Form2.frx":0000
      Left            =   2400
      List            =   "Form2.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2400
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4800
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cencel"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "sample"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   3375
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "abcdef"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.ListBox List1 
      Height          =   1785
      ItemData        =   "Form2.frx":0004
      Left            =   0
      List            =   "Form2.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   1785
      ItemData        =   "Form2.frx":0008
      Left            =   4920
      List            =   "Form2.frx":000A
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Font style:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command1_Click()
Form2.Visible = False
Form1.Visible = True
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
List1.Clear
                                           'Add font name to first list
List1.AddItem "arial"
List1.AddItem "Consolas"
List1.AddItem "Mv boli"
List1.AddItem "Bel MT"
List1.AddItem "Calibri"
List1.AddItem "corbel"
List1.AddItem "century"
List1.AddItem "Constantia"
List1.AddItem "Montserrat"
List1.AddItem "High Tower Text"
List1.AddItem "MS Serif"
List1.AddItem "Mangal"
List1.AddItem "Sanskrit Text"
List1.AddItem "Zilla Slab"
List1.AddItem "Times New Roman"
List1.AddItem "Ink Free"
List1.AddItem "Segoe Script"
List1.AddItem "Roboto"
List1.AddItem "Courier"

'list1.ListIndex = 0
                                             'Add Style to Second list
            List2.AddItem "Narrow"
            List2.AddItem "Bold"
            List2.AddItem "Italic"
            List2.AddItem "Underline"
            'List2.ListIndex = 0
                                             'Add Size to Third list
                                       List3.AddItem "8"
                                       List3.AddItem "10"
                                       List3.AddItem "12"
                                       List3.AddItem "14"
                                       List3.AddItem "16"
                                       List3.AddItem "18"
                                       List3.AddItem "20"
                                       List3.AddItem "22"
                                       List3.AddItem "24"
                                       List3.AddItem "26"
                                    
                               
                                
      
End Sub

Private Sub List1_Click()
Dim name As Variant
name = Trim(List1.Text)
If name = "arial" Then
             Form1.Text1.FontName = "Arial"
            Label4.FontName = "Arial"
            
ElseIf name = "Consolas" Then
              Form1.Text1.FontName = "Consolas"
            Label4.FontName = "Consolas"
            
ElseIf name = "Mv boli" Then
               Form1.Text1.FontName = "MV Boli"
            Label4.FontName = "MV Boli"
            
ElseIf name = "Bel MT" Then
          Form1.Text1.FontName = "Bell MT"
            Label4.FontName = "Bell MT"
            
ElseIf name = "Calibri" Then
            Form1.Text1.FontName = "Calibri"
            Label4.FontName = "Calibri"
ElseIf name = "corbel" Then
             Form1.Text1.FontName = "Corbel"
            Label4.FontName = "Corbel"
 ElseIf name = "century" Then
           Form1.Text1.FontName = "Century"
           Label4.FontName = "Century"
 ElseIf name = "Constantia" Then
          Form1.Text1.FontName = "Constantia"
           Label4.FontName = "Constantia"
ElseIf name = "Montserrat" Then
          Form1.Text1.FontName = "Montserrat"
           Label4.FontName = "Montserrat"
ElseIf name = "High Tower Text" Then
          Form1.Text1.FontName = "High Tower Text"
           Label4.FontName = "High Tower Text"
           



End If


End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = List1
End Sub



Private Sub List2_Click()
Dim style As Variant
style = Trim(List2.Text)
If style = "Bold" Then
      Form1.Text1.FontBold = True
       Text2.FontUnderline = False
        Label4.FontUnderline = False
              Text2.FontItalic = False
              Label4.FontItalic = False

        Text2.FontBold = True
        Label4.FontBold = True

ElseIf style = "Italic" Then
 Form1.Text1.FontItalic = True
         Text2.FontBold = False
         Text2.FontUnderline = False
         Label4.FontUnderline = False
                        
                        Text2.FontItalic = True
                       Label4.FontItalic = True
ElseIf style = "Underline" Then
 Form1.Text1.FontUnderline = True
         Text2.FontBold = False
        Label4.FontBold = False
        Text2.FontItalic = False
        Label4.FontItalic = False
                  
                  Text2.FontUnderline = True
                 Label4.FontUnderline = True
 
                 
  End If
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.Text = List2.Text

End Sub


Private Sub List3_Click()
Dim size As Variant
size = Trim(List3.Text)
If size = "8" Then
            Form1.Text1.FontSize = 8
             Label4.FontSize = 8
ElseIf size = "10" Then
  Form1.Text1.FontSize = 10
               Label4.FontSize = 10
ElseIf size = "12" Then
  Form1.Text1.FontSize = 12
        Label4.FontSize = 12
ElseIf size = "14" Then
      Form1.Text1.FontSize = 14
                Label4.FontSize = 14
ElseIf size = "16" Then
        Form1.Text1.FontSize = 16
                Label4.FontSize = 16
ElseIf size = "18" Then
  Form1.Text1.FontSize = 18
                Label4.FontSize = 18
ElseIf size = "20" Then
      Form1.Text1.FontSize = 20
               Label4.FontSize = 20
ElseIf size = "22" Then
    Form1.Text1.FontSize = 22
               Label4.FontSize = 22
ElseIf size = "24" Then
      Form1.Text1.FontSize = 24
              Label4.FontSize = 24
ElseIf size = "26" Then
    Form1.Text1.FontSize = 26
              Label4.FontSize = 26

End If
End Sub

Private Sub List3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = List3.Text
End Sub

Private Sub Option1_Click()

End Sub
