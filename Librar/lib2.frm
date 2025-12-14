VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Book Details"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7410
   Icon            =   "lib2.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5535
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Return"
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Book"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Rack No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Publisher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Book Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.RecordSource = "select * from book_mast"
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = CInt(Text5.Text)
Data1.Recordset.Fields(5) = CInt(Text6.Text)
Data1.Recordset.Update
Data1.Recordset.Close
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Data2.RecordSource = "select * from book_mast"
Form1.Data2.Refresh
Form3.Hide
Form1.Show
End Sub

Private Sub Text1_LostFocus()
On Error GoTo errorhandle
Data1.RecordSource = "select * from book_mast where book_code='" + Text1.Text + "'"
Data1.Refresh
	c = 0 ' Initialize counter
	Do While Not Data1.Recordset.EOF
	c = c + 1
Data1.Recordset.MoveNext
Loop
If c <> 0 Then
MsgBox "Duplicate Code", vbExclamation, "Duplicate"
Text1.Text = " "
Text1.SetFocus
Else
Text2.SetFocus
End If
Exit Sub
errorhandle:
MsgBox "Error occurred!Wrong Book Code", vbInformation, "Error"
	Exit Sub ' Use Exit Sub instead of End to prevent application termination
End Sub

