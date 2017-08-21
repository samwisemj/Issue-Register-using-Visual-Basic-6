VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Issue Page"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10545
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo12 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "Form5.frx":0000
      Left            =   600
      List            =   "Form5.frx":0002
      TabIndex        =   26
      Text            =   "Select"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox Combo14 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "Form5.frx":0004
      Left            =   8280
      List            =   "Form5.frx":0006
      TabIndex        =   25
      Text            =   "Select"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox Combo13 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "Form5.frx":0008
      Left            =   4560
      List            =   "Form5.frx":000A
      TabIndex        =   24
      Text            =   "Select"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox hhf 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5760
      TabIndex        =   23
      Text            =   "HH"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox mmf 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6240
      TabIndex        =   22
      Text            =   "MM"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox mmt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8520
      TabIndex        =   21
      Text            =   "MM"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox hht 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   20
      Text            =   "HH"
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox Combo15 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form5.frx":000C
      Left            =   6720
      List            =   "Form5.frx":0016
      TabIndex        =   19
      Text            =   "AM"
      Top             =   2520
      Width           =   975
   End
   Begin VB.ComboBox Combo16 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form5.frx":0022
      Left            =   9000
      List            =   "Form5.frx":002C
      TabIndex        =   18
      Text            =   "AM"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   16
      Text            =   "YYYY"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   15
      Text            =   "MM"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   14
      Text            =   "DD"
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form5.frx":0038
      Left            =   8280
      List            =   "Form5.frx":005A
      TabIndex        =   13
      Text            =   "Select"
      Top             =   960
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Temporary"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entire Semester"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc keydb 
      Height          =   330
      Left            =   8040
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
   Begin MSAdodcLib.Adodc projectordb 
      Height          =   330
      Left            =   6600
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
      RecordSource    =   "select distinct projector from projector"
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
   Begin MSAdodcLib.Adodc laptopdb 
      Height          =   330
      Left            =   5040
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
      RecordSource    =   "select distinct laptopname from laptop"
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
      Left            =   1800
      Top             =   6600
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   6240
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
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   29
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Projector"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Laptop"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Sets"
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
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "IssueDate"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No."
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue For"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8760
      MouseIcon       =   "Form5.frx":009D
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5520
      Width           =   645
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   960
      MouseIcon       =   "Form5.frx":03A7
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5640
      Width           =   585
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt As String
Dim X As Integer
Private Sub Combo2_Click()
Combo13.Enabled = True
Combo14.Enabled = False
'disabling kry or projector as per room no
If Combo2.Text = "N305" Or Combo2.Text = "N308" Or Combo2.Text = "N310" Or Combo2.Text = "N311" Or Combo2.Text = "N318" Then
Combo13.Enabled = False
End If
If Combo2.Text = "N308" Or Combo2.Text = "N310" Then
Combo14.Enabled = True
End If
End Sub

Private Sub Command1_Click()
'back button...going back to form 7 but loading that from here itself
Form7.Adodc1.RecordSource = "Select * from issues where Username='" + Text1.Text + "'"
Form7.Adodc1.Refresh
Form7.Adodc1.Caption = Adodc1.RecordSource
Form5.Text1.Text = Form5.Text1.Text
Form7.Show
Form5.Hide
End Sub

Private Sub Command2_Click()
Dim looper As Integer
looper = 0
'used flag f ..checking if the entries made are legal or not...if everythingis ok then flag f=0..else f will be set 1 and we will not issue..
Dim day, dat, laptop, roomno, user, projector, keys As String
Dim time_from, time_to As String
Dim dd, mm, yy, f As Integer

f = 0
'checking rmno is selected or not
If (Combo2.Text = "Select") Then
MsgBox "Select Room No.", vbCritical, "No Room No Selected"
f = 1
End If
roomno = Combo2.Text

user = Text1.Text
'checking dates are filled or not
If f = 0 Then
If (Text5.Text = "DD" Or Text6.Text = "MM" Or Text7.Text = "YYYY") Then
MsgBox "Select Date", vbCritical, "No Date Selected"
f = 1
End If
If f = 0 Then
dd = CInt(Text5.Text)
mm = CInt(Text6.Text)
yy = CInt(Text7.Text)
End If
'checking if legal dates were entered
 If (mm = 2 And yy Mod 4 <> 0 And dd > 28) Then
    MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
    f = 1
End If
If (mm = 2 And yy Mod 4 = 0 And dd > 29) Then
    MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        f = 1
            End If
If (dd <= 31 And mm <= 12) Then
    If (dd > 30 And (mm = 4 Or mm = 6 Or mm = 9 Or mm = 11 Or mm = 2)) Then
        MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        f = 1
    End If
Else
MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        f = 1
    End If
If (dd < 10) Then
dat = "0" + CStr(dd) + "/"
Else
dat = CStr(dd) + "/"
End If
If (mm < 10) Then
dat = dat + "0" + CStr(mm) + "/" + CStr(yy)
Else
dat = dat + CStr(mm) + "/" + CStr(yy)
End If
End If
If f = 0 Then
Dim k As Integer
k = Weekday(dat)
Select Case k
Case 1: day = "Sunday"
Case 2: day = "Monday"
Case 3: day = "Tuesday"
Case 4: day = "Wednesday"
Case 5: day = "Thursday"
Case 6: day = "Friday"
Case 7: day = "Saturday"
End Select
If day = "Saturday" Or day = "Sunday" Then
MsgBox "You cant Issue Anything On a Saturday Or Sunday. If you still need to issue contact administrator", vbInformation, "Contact Admin"
f = 1
End If
End If

If f = 0 Then
'checking time was filled or not
If (hhf.Text = "HH" Or hht.Text = "HH" Or mmf.Text = "MM" Or mmt.Text = "MM") Then
MsgBox "Incorrect Time", vbCritical, "Incorrect time"
hhf.SetFocus
f = 1
End If
End If

If (f = 0) Then 'checking time was legal or not
If (CInt(hhf.Text) < 12 Or CInt(hhf.Text) > 1 Or CInt(hht.Text) < 12 Or CInt(hht.Text) > 1 Or CInt(mmf.Text) < 59 Or CInt(mmf.Text) > 0 Or CInt(mmt.Text) < 59 Or CInt(mmt.Text) > 0) Then
time_from = CStr(hhf) + "." + CStr(mmf) + Combo15.Text
time_to = CStr(hht) + "." + CStr(mmt) + Combo16.Text
End If
End If

'checking every component was selected or not..it it was disabled then we need not to worry..but if it was enabled and user didnt select a set then we will generate error msg
If f = 0 Then
If Combo13.Enabled = False Then
projector = ""
ElseIf (Combo13.Enabled = True And Combo13.Text = "Select") Then
MsgBox "Select a Set", vbCritical, "No set selected"
Combo13.SetFocus
f = 1
Else
projector = Combo13.Text
End If
End If
If f = 0 Then
'same as above
If Combo14.Enabled = False Then
keys = ""
ElseIf (Combo14.Enabled = True And Combo14.Text = "Select") Then
MsgBox "Select a Set", vbCritical, "No set selected"
Combo14.SetFocus
f = 1
Else
keys = Combo14.Text
End If
End If
'same as above
If f = 0 Then
If Combo12.Text = "select" Then
MsgBox "Select a Set", vbCritical, "No set selected"
Combo12.SetFocus
f = 1
Else
laptop = Combo12.Text
End If
End If



Dim dt As Integer
If (f = 0) Then 'this is a check that will prevent you to enter redundant entry into database..fr ex...there cant be same issue for two user on the same day and same room. or no one can issue a single set for a common time
    Adodc1.RecordSource = "select * from issues where RoomNo='" + roomno + "' and issuedate='" + dat + "' and timefrom='" + time_from + "'"
    Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
        Adodc2.RecordSource = "select * from issues where LaptopID='" + laptop + "' and issuedate='" + dat + "' and timefrom='" + time_from + "'"
        Adodc2.Refresh
        If Adodc2.Recordset.EOF Then
            If Option1.Value = True Then 'for whole semester...even semester...we will loop and enter data till the month
                If (mm >= 1) And (mm <= 6) Then
                    Do While (mm < 7)
                        dt = dd
                        Adodc1.Recordset.AddNew
                        Adodc1.Recordset.Fields("RoomNo") = roomno
                        Adodc1.Recordset.Fields("LaptopID") = laptop
                        Adodc1.Recordset.Fields("ProjectorID") = projector
                        Adodc1.Recordset.Fields("KeyID") = keys
                        Adodc1.Recordset.Fields("IssueDate") = dat
                        Adodc1.Recordset.Fields("IssueDay") = day
                        Adodc1.Recordset.Fields("TimeFrom") = time_from
                        Adodc1.Recordset.Fields("TimeTo") = time_to
                        Adodc1.Recordset.Fields("Username") = user
                        Adodc1.Recordset.Update
                        dt = dd
                        If (mm = 1 Or mm = 3 Or mm = 5 Or mm = 7 Or mm = 8 Or mm = 10 Or mm = 12) Then
                            dd = (dd + 7)
                            If dd > 30 Then
                                dd = dd Mod 30
                            End If
                            If (dt > dd) Then
                                mm = mm + 1
                            End If
                        ElseIf (mm = 4 Or mm = 6 Or mm = 9 Or mm = 11) Then
                            dd = (dd + 7)
                            If dd > 31 Then
                                dd = dd Mod 31
                            End If
                            If (dt > dd) Then
                                mm = mm + 1
                            End If
                        Else
                            If (yy Mod 4 = 0) Then
                            dd = (dd + 7)
                            If dd > 29 Then
                                dd = dd Mod 29
                            End If
                            If dt > dd Then
                                mm = mm + 1
                            End If
                            Else
                            dd = (dd + 7)
                            If dd > 28 Then
                                dd = dd Mod 28
                            End If
                            If (dt > dd) Then
                                mm = mm + 1
                            End If
                        End If
                    End If
                    If (dd < 10) Then
                        dat = "0" + CStr(dd) + "/"
                    Else
                        dat = CStr(dd) + "/"
                    End If
                    If (mm < 10) Then
                        dat = dat + "0" + CStr(mm) + "/" + CStr(yy)
                    Else
                        dat = dat + CStr(mm) + "/" + CStr(yy)
                    End If
                Loop:
            Else
                Do While (mm <= 12) 'whole semester...odd semester
                    dt = dd
                    Adodc1.Recordset.AddNew
                    Adodc1.Recordset.Fields("RoomNo") = roomno
                    Adodc1.Recordset.Fields("LaptopID") = laptop
                    Adodc1.Recordset.Fields("ProjectorID") = projector
                    Adodc1.Recordset.Fields("KeyID") = keys
                    Adodc1.Recordset.Fields("IssueDate") = dat
                    Adodc1.Recordset.Fields("IssueDay") = day
                    Adodc1.Recordset.Fields("TimeFrom") = time_from
                    Adodc1.Recordset.Fields("TimeTo") = time_to
                    Adodc1.Recordset.Fields("Username") = user
                    Adodc1.Recordset.Update
                    If (mm = 1 Or mm = 3 Or mm = 5 Or mm = 7 Or mm = 8 Or mm = 10 Or mm = 12) Then
                        dd = (dd + 7) Mod 31
                        If (dt > dd) Then
                            mm = mm + 1
                        End If
                    ElseIf (mm = 4 Or mm = 6 Or mm = 9 Or mm = 11) Then
                        dd = (dd + 7) Mod 30
                        If (dt > dd) Then
                            mm = mm + 1
                        End If
                    Else
                        If (yy Mod 4 = 0) Then
                            dd = (dd + 7) Mod 29
                            If dt > dd Then
                                mm = mm + 1
                            End If
                        Else
                            dd = (dd + 7) Mod 28
                            If (dt > dd) Then
                                 mm = mm + 1
                            End If
                        End If
                    End If
                    If (dd < 10) Then
                        dat = "0" + CStr(dd) + "/"
                    Else
                        dat = CStr(dd) + "/"
                    End If
                    If (mm < 10) Then
                        dat = dat + "0" + CStr(mm) + "/" + CStr(yy)
                    Else
                        dat = dat + CStr(mm) + "/" + CStr(yy)
                    End If
                Loop:
            End If
        Else
            Adodc1.Recordset.AddNew
            Adodc1.Recordset.Fields("RoomNo") = roomno
            Adodc1.Recordset.Fields("LaptopID") = laptop
            Adodc1.Recordset.Fields("ProjectorID") = projector
            Adodc1.Recordset.Fields("KeyID") = keys
            Adodc1.Recordset.Fields("IssueDate") = dat
            Adodc1.Recordset.Fields("IssueDay") = day
            Adodc1.Recordset.Fields("TimeFrom") = time_from
            Adodc1.Recordset.Fields("TimeTo") = time_to
            Adodc1.Recordset.Fields("Username") = user
            Adodc1.Recordset.Update
            End If
    MsgBox "Issue Added Successfully", vbInformation, "Item(s) issued"
    looper = 1
    Else
    MsgBox "This laptop set has already been issued for this particular time.Select a different laptop set", vbCritical, "Issue already present"
    Combo12.SetFocus
    End If
    Else
    MsgBox "An issue is already present for the same room no at the same date and time. Choose an alternative time.", vbCritical, "Select another time"
        Combo2.SetFocus
    End If
Else
    MsgBox "Please Try Again", vbInformation, "Try Again"
    Combo2.SetFocus
End If
If looper = 1 Then
Combo2.Text = "Select"
Text5.Text = "DD"
Text5.FontSize = 10

Text2.Text = ""
Text6.Text = "MM"
Text6.FontSize = 10

Combo12.Text = "Select"
Combo13.Text = "Select"
Combo14.Text = "Select"
hhf.Text = "HH"
hhf.FontSize = 10
mmf.Text = "MM"
mmt.Text = "MM"
hht.Text = "HH"
mmf.FontSize = 10
mmt.FontSize = 10
hht.FontSize = 10
Text7.Text = "YYYY"
Text7.FontSize = 10

'back button...going back to form 7 but loading that from here itself
Form7.Adodc1.RecordSource = "Select * from issues where Username='" + Text1.Text + "'"
Form7.Adodc1.Refresh
Form7.Adodc1.Caption = Adodc1.RecordSource
Form5.Text1.Text = Form5.Text1.Text
Form7.Show
Form5.Hide
End If
End Sub



Private Sub Form_Load()
laptopdb.Refresh

laptopdb.Recordset.Requery

With laptopdb.Recordset
Do Until .EOF
Combo12.AddItem ![LaptopName]
.MoveNext
Loop:
End With
projectordb.Refresh

projectordb.Recordset.Requery

With projectordb.Recordset
Do Until .EOF
Combo13.AddItem ![projector]
.MoveNext
Loop:
End With
keydb.Refresh

keydb.Recordset.Requery

With keydb.Recordset
Do Until .EOF
Combo14.AddItem ![roomKey]
.MoveNext
Loop:
End With
End Sub

Private Sub Form_Paint()
    Dim oPic As StdPicture
    Set oPic = LoadPicture("C:\Issue Register Project\Files\b3.jpg")
    PaintPicture oPic, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set oPic = Nothing
End Sub



Private Sub Text2_Click()
Dim d, m, Y As Integer
Dim dat, day As String
X = 0
If (Text5.Text = "DD" Or Text6.Text = "mm" Or Text7.Text = "YYYY") Then
MsgBox "Select Date", vbCritical, "No Date selected"
X = 1
End If
If X = 0 Then
d = CInt(Text5.Text)
m = CInt(Text6.Text)
Y = CInt(Text7.Text)
Else
Text5.SetFocus
End If
'checking if legal dates were entered
 If (m = 2 And Y Mod 4 <> 0 And d > 28) Then
    MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        X = 1
    End If
If (m = 2 And Y Mod 4 = 0 And d > 29) Then
    MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        X = 1
        End If
If (d <= 31 And m <= 12) Then
    If (d > 30 And (m = 4 Or m = 6 Or m = 9 Or m = 11 Or m = 2)) Then
        MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        X = 1
    End If
Else
MsgBox "Incorrect Date", vbCritical, "Incorrect Date"
        Text5.SetFocus
        X = 1
    End If
If X = 0 Then
If (d < 10) Then
dat = "0" + CStr(d) + "/"
Else
dat = CStr(d) + "/"
End If
If (m < 10) Then
dat = dat + "0" + CStr(m) + "/" + CStr(Y)
Else
dat = dat + CStr(m) + "/" + CStr(Y)
End If
Dim k As Integer
k = Weekday(dat)
Select Case k
Case 1: day = "Sunday"
Case 2: day = "Monday"
Case 3: day = "Tuesday"
Case 4: day = "Wednesday"
Case 5: day = "Thursday"
Case 6: day = "Friday"
Case 7: day = "Saturday"
End Select
Text2.Text = day
Text2.Enabled = False
Else
MsgBox "Incorrect Date Entered", vbCritical, "Incorrect Date"
End If
End Sub

Private Sub Text5_Click()
If Text5.Text = "DD" Then
Text5.Text = ""
End If
End Sub
Private Sub Text6_Click()
If Text6.Text = "MM" Then
Text6.Text = ""
End If
End Sub
Private Sub Text7_Click()
If Text7.Text = "YYYY" Then
Text7.Text = ""
End If
End Sub

Private Sub Text7_Change()
Text7.FontSize = 14
Text2.Enabled = True

End Sub
Private Sub Text5_Change()
Text5.FontSize = 14
Text2.Enabled = True
End Sub
Private Sub Text6_Change()
Text6.FontSize = 14
Text2.Enabled = True
End Sub
Private Sub hhf_Change()
hhf.FontSize = 14
End Sub
Private Sub hht_Change()
hht.FontSize = 14
End Sub
Private Sub mmf_Change()
mmf.FontSize = 14
End Sub
Private Sub mmt_Change()
mmt.FontSize = 14
End Sub
Private Sub hhf_Click()
hhf.Text = ""
mmf.Text = ""
End Sub

Private Sub hht_Click()
hht.Text = ""
mmt.Text = ""
End Sub

