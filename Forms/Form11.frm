VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "About Software"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12675
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form11"
   ScaleHeight     =   6810
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form11.frx":0000
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Software...."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   6495
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\software.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub

