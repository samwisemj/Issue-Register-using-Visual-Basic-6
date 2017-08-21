VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RCCIIT Computer Science and Engineering Department Issue Register"
   ClientHeight    =   6585
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "HP Simplified"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000B&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   6585
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   Departmental Issue                     Registry"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1080
      MouseIcon       =   "Form1.frx":BE87
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Temporary  Issue   Registry/Admin"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7080
      MouseIcon       =   "Form1.frx":C191
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":C49B
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   9855
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu more 
      Caption         =   "More...."
      Begin VB.Menu adev 
         Caption         =   "About Developers....."
      End
      Begin VB.Menu IssueRegister 
         Caption         =   "About Software....."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adev_Click()
Form10.Show
End Sub

Private Sub IssueRegister_Click()
Form11.Show
End Sub

Private Sub Label1_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub Label2_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Paint()
    Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\b5.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub

