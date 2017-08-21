VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Create an Account"
   ClientHeight    =   4560
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   5190
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   240
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Issue Register Project\Database4.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Issue Register Project\Database4.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox confirmpassword 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2520
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Issue Register Project\Database4.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Issue Register Project\Database4.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox password 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox user 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   435
      Left            =   3360
      MouseIcon       =   "Form6.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   840
      MouseIcon       =   "Form6.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:-"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:-"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:-"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Len(user.Text) <> 0) And (password.Text = confirmpassword.Text) And (password.Text <> "") Then
Adodc1.RecordSource = "select * from Table1 where UserName='" + user.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
'use this adodc command to update individual field of a table in a database as per
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("UserName") = user.Text
Adodc1.Recordset.Fields("Password") = password.Text
Adodc1.Recordset.Update
MsgBox "Successfully Registered"
Form6.Hide
Form2.Show
Else
MsgBox "A same user is already present. Select a different username", vbCritical, "Same Username present"
password.Text = ""
confirmpassword.Text = ""
user.SetFocus
End If
Else
MsgBox "Incorrect Details..Password Mismatch or Empty field", vbCritical
End If
End Sub

Private Sub Command2_Click()
Form2.Show
Form6.Hide
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
user.Text = ""
password.Text = ""
confirmpassword.Text = ""
End Sub
Private Sub Form_Paint()
    Dim oPic As StdPicture
    
    Set oPic = LoadPicture("C:\Issue Register Project\Files\m3.jpg")
 
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
 
    Set oPic = Nothing
End Sub

