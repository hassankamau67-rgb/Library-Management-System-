VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Delete Member"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7380
   Icon            =   "lib6.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   5460
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
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
      Left            =   5880
      TabIndex        =   16
      Top             =   3720
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
      Left            =   5880
      TabIndex        =   15
      Top             =   3240
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
      Left            =   5880
      TabIndex        =   14
      Top             =   2760
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
      Left            =   5880
      TabIndex        =   13
      Top             =   2280
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
      Left            =   5880
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
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
      Left            =   5880
      TabIndex        =   11
      Top             =   1320
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
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mem_mast"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Ren_date"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "Mem_fees"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "Mem_code"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Renewal Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Member Fees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Member Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Delete Member"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form5"
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
Form5.Hide
Form1.Show
Form1.Data1.RecordSource = "select * from mem_mast"
Form1.Data1.Refresh
End Sub
