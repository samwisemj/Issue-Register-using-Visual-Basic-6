VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00C0E0FF&
   Caption         =   "User Issues"
   ClientHeight    =   8670
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   18735
   LinkTopic       =   "Form7"
   ScaleHeight     =   8670
   ScaleWidth      =   18735
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00C9ABC8&
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Text            =   "Search by any field"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ascending"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Descending"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   18480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7800
      TabIndex        =   1
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11880
      Top             =   7440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "select * from issues"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":0000
      Height          =   5775
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   19215
      _ExtentX        =   33893
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   14139350
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HP Simplified"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Issue"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   4800
      MouseIcon       =   "Form7.frx":0015
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   9000
      Width           =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFC0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFC0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   600
      X2              =   600
      Y1              =   1080
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   1080
      Y2              =   2040
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   3600
      MouseIcon       =   "Form7.frx":031F
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All My Issue"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   435
      Left            =   8640
      MouseIcon       =   "Form7.frx":0629
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make New Issue"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   1185
      MouseIcon       =   "Form7.frx":0933
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   9000
      Width           =   2145
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usernm As String

Private Sub Command1_Click()
Form5.Show
Form7.Hide
End Sub

Private Sub Command2_Click()
'delete of a record
If Adodc1.Recordset.RecordCount > 0 Then
Dim cnf As Integer
cnf = MsgBox("Are you Sure you want to delete?", vbYesNo, "Warning Message")
If cnf = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Issue successfully deleted", vbInformation, "Deleted Record Confirmation"
Else
MsgBox "Issue Not Deleted", vbInformation, "Not deleted"
End If
Else
MsgBox "No Data to be deleted", vbExclamation, "No Records"
End If

End Sub

Private Sub Command3_Click()
'special type of search...see carefully i have used username with every field combination such that when i filter the data from the common issue database i only get data related to the username who looged in
Adodc1.RecordSource = "Select * from issues where laptopid= '" + Text1.Text + "'and Username='" + Text2.Text + "'or projectorid= '" + Text1.Text + "'and Username='" + Text2.Text + "'or keyid= '" + Text1.Text + "'and Username='" + Text2.Text + "'or issuedate= '" + Text1.Text + "'and Username='" + Text2.Text + "'or issueday= '" + Text1.Text + "'and Username='" + Text2.Text + "'or timefrom= '" + Text1.Text + "'and Username='" + Text2.Text + "'or timeto= '" + Text1.Text + "'and Username='" + Text2.Text + "'or roomno= '" + Text1.Text + "' and Username='" + Text2.Text + "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Requery
DataGrid1.Refresh
'view all issue of logged in user
Adodc1.RecordSource = "Select * from issues where Username='" + Text2.Text + "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource

End Sub


Private Sub exit_Click()
End
End Sub


Private Sub Form_Load()
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
End Sub

Private Sub Form_Paint()
Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\bgp.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Logout_Click()
Me.Hide
Form2.Show
Form2.Text1 = ""
Form2.Text2 = ""

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Adodc1.Recordset.Sort = "id ASC"
Else
Adodc1.Recordset.Sort = "id DESC"
End If
End Sub
Private Sub Option4_Click()
If Option3.Value = True Then
Adodc1.Recordset.Sort = "id ASC"
Else
Adodc1.Recordset.Sort = "id DESC"
End If
End Sub


Private Sub Text1_Click()
'ignore this
If (Text1.Text = "Search by any field") Then
Text1.Text = ""
End If
End Sub
