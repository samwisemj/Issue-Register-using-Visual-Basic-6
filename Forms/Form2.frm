VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faculty member Login"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   4815
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1635
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1635
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      RecordSource    =   "select * from Table1"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   3600
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   495
      TabIndex        =   4
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      MouseIcon       =   "Form2.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Command2 
      BackStyle       =   0  'Transparent
      Caption         =   "New User Register Here..."
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      MouseIcon       =   "Form2.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'type of search command to search the username and password pair.
Adodc1.RecordSource = "select * from table1 where username='" + Text1.Text + "' and password='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then 'if adodc controler moves end of file, correct password or username pair was not found
MsgBox "Incorrect login credentils", vbCritical
Else 'correct password or username pair was found
Form7.Text2.Text = Text1.Text
'the next lines load the adodc of form 7 from here itself...
Form7.Adodc1.RecordSource = "Select * from issues where Username='" + Text1.Text + "'"
Form7.Adodc1.Refresh
Form7.Adodc1.Caption = Adodc1.RecordSource
Form5.Text1.Text = Form2.Text1.Text 'ignore this...i make a txt box in form 5 to use it as a variable
Form7.Show
Form2.Hide
End If
End Sub

Private Sub Command2_Click()
Form6.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Paint()
    Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\m2.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub

Private Sub Label3_Click()
Me.Hide
Form1.Show

End Sub
