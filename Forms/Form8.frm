VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00404040&
   Caption         =   "Add/Remove Components"
   ClientHeight    =   6765
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9495
   LinkTopic       =   "Form8"
   ScaleHeight     =   6765
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3960
      Top             =   6360
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
      RecordSource    =   "select distinct roomkey from roomkeys"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2640
      Top             =   6480
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
      RecordSource    =   "Select distinct projector from projector"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3240
      Top             =   5880
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
      RecordSource    =   "select distinct LaptopName from laptop"
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   6240
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Issue Components"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   435
      Left            =   2880
      MouseIcon       =   "Form8.frx":0000
      TabIndex        =   17
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Command8 
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
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   1200
      MouseIcon       =   "Form8.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Command7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update All"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6720
      MouseIcon       =   "Form8.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   7680
      MouseIcon       =   "Form8.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   4680
      MouseIcon       =   "Form8.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   1800
      MouseIcon       =   "Form8.frx":0F32
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Enabled         =   0   'False
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
      Left            =   6360
      MouseIcon       =   "Form8.frx":123C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5040
      Width           =   525
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Enabled         =   0   'False
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
      Left            =   570
      MouseIcon       =   "Form8.frx":1546
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5040
      Width           =   525
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Enabled         =   0   'False
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
      Left            =   3360
      MouseIcon       =   "Form8.frx":1850
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5040
      Width           =   525
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keys:-"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Projectors:-"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Laptops:-"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Command1_Click()
List1.AddItem (Text1.Text)
Text1.Text = ""
End Sub

Private Sub Command2_Click()
List2.AddItem (Text2.Text)
Text2.Text = ""
End Sub

Private Sub Command3_Click()
List3.AddItem (Text3.Text)
Text3.Text = ""
End Sub

Private Sub Command4_Click()
Dim c As Integer
c = MsgBox("Sure You Want to remove the selected set?", vbYesNo, "Warning")
If c = vbYes Then
List1.RemoveItem (List1.ListIndex)
MsgBox "Item Removed", vbInformation, "Item Removed"
Else
MsgBox "Item Not Removed", vbInformation, "Item Not Removed"
End If
End Sub
Private Sub Command5_Click()
Dim c As Integer
c = MsgBox("Sure You Want to remove the selected set?", vbYesNo, "Warning")
If c = vbYes Then
List2.RemoveItem (List2.ListIndex)
MsgBox "Item Removed", vbInformation, "Item Removed"
Else
MsgBox "Item Not Removed", vbInformation, "Item Not Removed"
End If
End Sub
Private Sub Command6_Click()
Dim c As Integer
c = MsgBox("Sure You Want to remove the selected set?", vbYesNo, "Warning")
If c = vbYes Then
List3.RemoveItem (List3.ListIndex)
MsgBox "Item Removed", vbInformation, "Item Removed"
Else
MsgBox "Item Not Removed", vbInformation, "Item Not Removed"
End If
End Sub

Private Sub Command7_Click()
Dim i As Integer
'a with command releaves you to type the whole adodc.recordsource.etc etc ...just look and understand the command. its kind of a loop
Adodc3.Refresh
'first i use .delete to delete a record .moveNext moves the adodc control to next record...after deleting the whole table of components database then i will add all the elements that was in the list box...if this was not done then you would see that there are redundant set of copies in the database
With Adodc3.Recordset
Do Until .EOF
.Delete
.MoveNext
Loop:
End With
Adodc3.Refresh
i = 0
'i will keep on adding untill loop counter i reaches listcount(ie.no of item in the list)
With Adodc3.Recordset
Do Until (i = List3.ListCount)
.AddNew
.Fields("roomkey") = List3.List(i)
.Update
i = i + 1
Loop:
End With
'similar concept and command for other component like above
Adodc1.Refresh
With Adodc1.Recordset
Do Until .EOF
.Delete
.MoveNext
Loop:
End With
Adodc1.Refresh
i = 0
With Adodc1.Recordset
Do Until (i = List1.ListCount)
.AddNew
.Fields("laptopName") = List1.List(i)
.Update
i = i + 1
Loop:
End With


Adodc2.Refresh
With Adodc2.Recordset
Do Until .EOF
.Delete
.MoveNext
Loop:
End With
Adodc2.Refresh
i = 0
With Adodc2.Recordset
Do Until (i = List2.ListCount)
.AddNew
.Fields("projector") = List2.List(i)
.Update
i = i + 1
Loop:
End With
MsgBox "Data Updated", vbInformation


End Sub

Private Sub Command8_Click()
Form4.Show
Me.Hide
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
'when the form loads i will load all the components from the dbase to the listboxes...use three databse for three components
Adodc1.Refresh
'same commands for three databases
'frst component loading
With Adodc1.Recordset
Do Until .EOF
List1.AddItem ![LaptopName]
.MoveNext
Loop:
End With
'now the second component loading
Adodc2.Refresh
With Adodc2.Recordset
Do Until .EOF
List2.AddItem ![projector]
.MoveNext
Loop:
End With
'third component loading
Adodc3.Refresh
With Adodc3.Recordset
Do Until .EOF
List3.AddItem ![roomKey]
.MoveNext
Loop:
End With
'by default i select the first item of all the list.'(0)' indicates index..list index starts from 0
List1.Selected(0) = True
List2.Selected(0) = True
List3.Selected(0) = True

End Sub
Private Sub Form_Paint()
    Dim oPic As StdPicture
    
    Set oPic = LoadPicture("C:\Issue Register Project\Files\b7.jpg")
 
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
 
    Set oPic = Nothing
End Sub

Private Sub Label6_Click()

End Sub

'if the length of the string in text boxes are not zero.it means user wants to add a new item to list ...so we will enable the add option which i keep disabled from begining to avoid inclusion of blank records in the database table
Private Sub Text1_Change()
Command1.Enabled = (Len(Text1.Text) > 0)
End Sub
Private Sub Text2_Change()
Command2.Enabled = (Len(Text2.Text) > 0)
End Sub
Private Sub Text3_Change()
Command3.Enabled = (Len(Text3.Text) > 0)
End Sub


