VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Login Required"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   5730
   FillStyle       =   2  'Horizontal Line
   FontTransparent =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   4950
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H00800000&
      Caption         =   "Temporary"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Form3.frx":A323
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Departmental"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Form3.frx":A62D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   1920
      MouseIcon       =   "Form3.frx":A937
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Command2 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Form3.frx":AC41
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "admin" And Text2.Text = "admin" Then
If (Option1.Value = True) Then
Form4.Show
Form3.Hide
Else
Form9.Show
Form3.Hide
End If
Else
MsgBox "Invalid Login Credentials", vbCritical, "Incorrect Password"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Form1.Show
Me.Hide

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Paint()
Dim oPic As StdPicture
    
    Set oPic = LoadPicture("C:\Issue Register Project\Files\background-learner.jpg")
 
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
 
    Set oPic = Nothing
   
End Sub
