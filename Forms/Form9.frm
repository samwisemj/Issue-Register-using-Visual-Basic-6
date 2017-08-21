VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Temporary Issue Database"
   ClientHeight    =   8745
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20250
   LinkTopic       =   "Form9"
   ScaleHeight     =   8745
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   15840
      TabIndex        =   6
      Text            =   "Search by any field"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Descending"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ascending"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E2D1E2&
      Caption         =   "View Mode"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1230
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E2D1E2&
      Caption         =   "Edit Mode"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1230
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form9.frx":0000
      Height          =   6135
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14864866
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HP Simplified"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12600
      Top             =   7680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      RecordSource    =   "select * from temporaryissues"
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
   Begin VB.Line Line4 
      BorderColor     =   &H00FF80FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   19800
      X2              =   19800
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF80FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   15600
      X2              =   15600
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF80FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   15600
      X2              =   19800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF80FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   15600
      X2              =   19800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   3600
      MouseIcon       =   "Form9.frx":0015
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   1080
      MouseIcon       =   "Form9.frx":031F
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   8760
      Width           =   1485
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   6000
      MouseIcon       =   "Form9.frx":0629
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   8760
      Width           =   1275
   End
   Begin VB.Label Command1 
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
      Height          =   435
      Left            =   18960
      MouseIcon       =   "Form9.frx":0933
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All Issues"
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
      Left            =   7080
      MouseIcon       =   "Form9.frx":0C3D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Menu Logout 
      Caption         =   "Logout"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'serching the databse for the string input by the user in the searchbox..notice the difference in 'and' & 'or'
Adodc1.RecordSource = "Select * from temporaryissues where returnedby='" + Text1.Text + "' or laptop= '" + Text1.Text + "'or projector= '" + Text1.Text + "'or key= '" + Text1.Text + "'or issuedate= '" + Text1.Text + "'or issueday= '" + Text1.Text + "'or timefrom= '" + Text1.Text + "'or timeto= '" + Text1.Text + "'or returndate= '" + Text1.Text + "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command2_Click()
'Adodc1.Recordset.Update
Adodc1.Recordset.Requery
DataGrid1.Refresh
'view all issues
Adodc1.RecordSource = "Select * from temporaryissues"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command3_Click()
Dim cnf As Integer
'deleting an issue in database
cnf = MsgBox("Are you Sure you want to delete?", vbYesNo, "Warning Message")
If cnf = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Issue successfully deleted", vbInformaion, "Deleted Record Confirmation"
Else
MsgBox "Issue Not Deleted", vbInformation, "Not deleted"
End If
End Sub

Private Sub Command5_Click()
DataGrid1.AllowAddNew = True 'this command allows a new entry into the table of the databse,which was otherwise kept disabled
Command5.Enabled = False 'after pressing addnew button, you cant press it again
Command4.Enabled = True
End Sub
Private Sub Command4_Click()
Adodc1.Recordset.Update
'Adodc1.RecordSource = "Select * from temporaryissues"
'Adodc1.Refresh
'Adodc1.Caption = Adodc1.RecordSource
'after pressing update , records get updated and addnew button is made available again
Command4.Enabled = True
Command5.Enabled = True
DataGrid1.AllowAddNew = False
End Sub

Private Sub exit_Click()
End
End Sub


Private Sub Logout_Click()
Me.Hide
Form3.Show
Form3.Text1.Text = ""
Form3.Text2.Text = ""
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
Private Sub Option1_Click()
'view mode.view mode disables editing options
DataGrid1.AllowAddNew = False
DataGrid1.AllowDelete = False
DataGrid1.AllowUpdate = False
Command5.Enabled = False
Command4.Enabled = False
Command3.Enabled = False

End Sub

Private Sub Option2_Click()
'mode selection...edit mode enables editing options
DataGrid1.AllowDelete = True
DataGrid1.AllowUpdate = True
Command5.Enabled = True
Command4.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Text1_Click()
If (Text1.Text = "Search by any field") Then
Text1.Text = ""
ElseIf Text1.Text = "" Then
Text1.Text = "Search by any field"
End If
End Sub

Private Sub Form_Paint()
    Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\bb1.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub

