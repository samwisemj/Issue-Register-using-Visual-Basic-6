VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Departmental Issues-Admin"
   ClientHeight    =   9120
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16455
   LinkTopic       =   "Form4"
   ScaleHeight     =   9120
   ScaleWidth      =   16455
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
      Left            =   16080
      TabIndex        =   6
      Text            =   "Search by any field"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FF8080&
      Caption         =   "Descending order"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FF8080&
      Caption         =   "Ascending order"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF8080&
      Caption         =   "Edit Mode"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "View Mode"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   13320
      Top             =   10320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Bindings        =   "Form4.frx":0000
      Height          =   6615
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   19575
      _ExtentX        =   34528
      _ExtentY        =   11668
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   15960
      X2              =   20160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   15960
      X2              =   20160
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   15960
      X2              =   15960
      Y1              =   840
      Y2              =   1800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   20160
      X2              =   20160
      Y1              =   840
      Y2              =   1800
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   675
      MouseIcon       =   "Form4.frx":0015
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   4995
      MouseIcon       =   "Form4.frx":031F
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   2910
      MouseIcon       =   "Form4.frx":0629
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   9480
      Width           =   1425
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Height          =   375
      Left            =   18720
      MouseIcon       =   "Form4.frx":0933
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All Issue Record"
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
      Height          =   345
      Left            =   6870
      MouseIcon       =   "Form4.frx":0C3D
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
   Begin VB.Menu Option 
      Caption         =   "More Option"
      Begin VB.Menu Addremoveproducts 
         Caption         =   "Add/Remove Products"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Addremoveproducts_Click()
Form8.Show
End Sub


Private Sub Form_Load()
DataGrid1.AllowUpdate = False
End Sub

Private Sub Form_Paint()
    Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\db3.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
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


Private Sub Command1_Click()
DataGrid1.AllowAddNew = True 'this command allows a new entry into the table of the databse,which was otherwise kept disabled
Command1.Enabled = False 'after pressing addnew button, you cant press it again
Command1.FontUnderline = False
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
Adodc1.Recordset.Resync
DataGrid1.Refresh
'after pressing update , records get updated and addnew button is made available again
Command1.Enabled = True
Command1.FontUnderline = True
'DataGrid1.AllowAddNew = False
Adodc1.Caption = Adodc1.RecordSource
'view mode.view mode disables editing options
DataGrid1.AllowAddNew = False
DataGrid1.AllowDelete = False
DataGrid1.AllowUpdate = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Option1.Value = True
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
Adodc1.Recordset.Update
Adodc1.Recordset.Requery
End If
End Sub


Private Sub Command4_Click()
'serching the databse for the string input by the user in the searchbox..notice the difference in 'and' & 'or'
Adodc1.RecordSource = "Select * from issues where Username='" + Text1.Text + "' or laptopid= '" + Text1.Text + "'or projectorid= '" + Text1.Text + "'or keyid= '" + Text1.Text + "'or issuedate= '" + Text1.Text + "'or issueday= '" + Text1.Text + "'or timefrom= '" + Text1.Text + "'or timeto= '" + Text1.Text + "'or roomno= '" + Text1.Text + "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub


Private Sub Command5_Click()
'Adodc1.Recordset.Update
Adodc1.Recordset.Requery
DataGrid1.Refresh
'view all issues
Adodc1.RecordSource = "Select * from issues"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub exit_Click()
End
End Sub
Private Sub Option2_Click()
'mode selection...edit mode enables editing options
DataGrid1.AllowDelete = True
DataGrid1.AllowUpdate = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
End Sub
Private Sub Option1_Click()
'view mode.view mode disables editing options
DataGrid1.AllowAddNew = False
DataGrid1.AllowDelete = False
DataGrid1.AllowUpdate = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Option5_Click()

End Sub

Private Sub Text1_Click()
'ignore this
If Text1.Text = "Search by any field" Then
Text1.Text = ""
End If
End Sub

