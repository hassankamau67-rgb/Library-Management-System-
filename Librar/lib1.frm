VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Library"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8775
   Icon            =   "lib1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "&End"
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Delete Book"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete Member"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Issue"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add &Book "
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New Member"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Member Details"
      TabPicture(0)   =   "lib1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(1)=   "Label1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Book Details"
      TabPicture(1)   =   "lib1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "DBGrid2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Issue Details"
      TabPicture(2)   =   "lib1.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DBGrid3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "lib1.frx":0496
         Height          =   2295
         Left            =   1080
         OleObjectBlob   =   "lib1.frx":04AA
         TabIndex        =   3
         Top             =   1200
         Width           =   4935
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "lib1.frx":0E7D
         Height          =   2175
         Left            =   -74040
         OleObjectBlob   =   "lib1.frx":0E91
         TabIndex        =   2
         Top             =   1440
         Width           =   5535
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "lib1.frx":1864
         Height          =   1935
         Left            =   -73920
         OleObjectBlob   =   "lib1.frx":1878
         TabIndex        =   1
         Top             =   1440
         Width           =   4935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Book &Return"
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "ABC NationWide Library"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "ABC NationWide Library"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73320
         TabIndex        =   11
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "ABC NationWide Library"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73560
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form1.Hide
Form2.Show
Form2.Text1.Text = ""
Form2.Text2.Text = ""
Form2.Text3.Text = ""
Form2.Text4.Text = ""
Form2.Text5.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Hide
Form3.Show
Form3.Text1.Text = ""
Form3.Text2.Text = ""
Form3.Text3.Text = ""
Form3.Text4.Text = ""
Form3.Text5.Text = ""
Form3.Text6.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Hide
Form4.Show
Form4.Data1.RecordSource = "select * from Issue_mast"
Form4.Data1.Refresh
Form4.Text1.Text = ""
Form4.Text2.Text = ""
Form4.Text3.Text = ""
Form4.Text5.Text = ""
Form4.Text6.Text = ""
Form4.Text8.Text = ""
Form4.Text1.SetFocus
End Sub

Private Sub Command4_Click()
d = InputBox("Enter the member code to delete the record", "Delete")
	c = 0 ' Initialize counter
	Data3.RecordSource = "select * from issue_mast where mem_code='" + d + "'"
	Data3.Refresh
	Do While Not Data3.Recordset.EOF
	c = c + 1
Data3.Recordset.MoveNext
Loop
If c <> 0 Then
MsgBox "The member has not returned the book", vbInformation, "Library"
Data3.Recordset.Close
Else
Data1.RecordSource = "select * from mem_mast"
Data1.Refresh
Form1.Hide
Form5.Show
End If
End Sub

Private Sub Command5_Click()
Dim r As String
Dim qty As Integer
On Error Resume Next:
d1 = Date$
r = InputBox("Enter your member code", "Book Return")
Data3.RecordSource = "select * from issue_mast where mem_code='" + r + "'"
Data3.Refresh
qty = Data3.Recordset.Fields(6)
d1 = Data3.Recordset.Fields(5)
Data3.Recordset.Delete
Data3.Recordset.Close
Data3.RecordSource = "select * from issue_mast"
Data3.Refresh
b = InputBox("Enter your book code", "Book Return")
Data2.RecordSource = "select * from book_mast where book_code='" + b + "'"
Data2.Refresh
Data2.Recordset.Edit
s = Data2.Recordset.Fields(4)
cs = s + qty
' Data1.Recordset.Fields(0) = j ' Removed: Variable j is undefined and Data1 is the wrong data control
Data2.Recordset.Fields(4) = cs
Data2.Recordset.Update
Data2.Recordset.Close
Form7.Show
End Sub

Private Sub Command6_Click()
d = InputBox("Enter the book code to delete the book", "Delete")
	c = 0 ' Initialize counter
	Data3.RecordSource = "select * from issue_mast where book_code='" + d + "'"
	Data3.Refresh
	Do While Not Data3.Recordset.EOF
	c = c + 1
Data3.Recordset.MoveNext
Loop
If c <> 0 Then
MsgBox "Don't Delete, we have issued this book", vbInformation, "Library"
Data3.Recordset.Close
Else
Data2.RecordSource = "select * from book_mast"
Data2.Refresh
Form1.Hide
Form6.Show
End If
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Form_Load()
If SSTab1.Tab = 0 Then
Data1.RecordSource = "select * from mem_mast"
Data1.Refresh
Command1.Visible = True
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = False
Command6.Visible = False
End If
If SSTab1.Tab = 1 Then
Data2.RecordSource = "select * from Book_mast"
Data2.Refresh
Command1.Visible = False
Command2.Visible = True
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = True
End If
If SSTab1.Tab = 2 Then
Data3.RecordSource = "select * from Issue_mast"
Data3.Refresh
Command1.Visible = False
Command2.Visible = False
Command3.Visible = True
Command4.Visible = False
Command5.Visible = True
Command6.Visible = False
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
Data1.RecordSource = "select * from Mem_mast"
Data1.Refresh
Command1.Visible = True
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = False
Command6.Visible = False
ElseIf SSTab1.Tab = 1 Then
Data2.RecordSource = "select * from Book_mast"
Data2.Refresh
Command1.Visible = False
Command2.Visible = True
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = True
ElseIf SSTab1.Tab = 2 Then
Data3.RecordSource = "select * from Issue_mast"
Data3.Refresh
Command1.Visible = False
Command2.Visible = False
Command3.Visible = True
Command4.Visible = False
Command5.Visible = True
Command6.Visible = False
End If
End Sub

