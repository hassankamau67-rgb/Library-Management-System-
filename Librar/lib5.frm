VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Delete Book"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7920
   Icon            =   "lib5.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   5505
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Rack_No"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Stock"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "Publisher"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "Author"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "Book_Name"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "Book_code"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Book_mast"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Book Deletion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label6 
      Caption         =   "Rack No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   17
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Publisher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Book Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveFirst
Command1.Enabled = False
Command4.Enabled = False
Command3.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveNext
s = Data1.Recordset.RecordCount
Command1.Enabled = True
Command4.Enabled = True
If Data1.Recordset.AbsolutePosition = s - 1 Then
Command2.Enabled = False
Command3.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.MoveLast
Command1.Enabled = True
Command4.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Data1.Recordset.MovePrevious
s = Data1.Recordset.RecordCount
Command2.Enabled = True
Command3.Enabled = True
If Data1.Recordset.AbsolutePosition = s Then
Command1.Enabled = False
Command4.Enabled = False
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.Delete
If Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
Else
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command6_Click()
Form6.Hide
Form1.Show
Form1.Data2.RecordSource = "select * from book_mast"
Form1.Data2.Refresh
End Sub

