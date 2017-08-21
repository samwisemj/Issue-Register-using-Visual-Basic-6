VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Developers :)"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15930
   LinkTopic       =   "Form10"
   ScaleHeight     =   9360
   ScaleMode       =   0  'User
   ScaleWidth      =   15925
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\Dev\developers2.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub
