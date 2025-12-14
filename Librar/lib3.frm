VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Member Details"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7410
   Icon            =   "lib3.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Return"
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Member"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   4920
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
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Member Details"
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
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "Renewal Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Member fees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Member Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.RecordSource = "select * from mem_mast"
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = CInt(Text4.Text)
Data1.Recordset.Fields(4) = CDate(Text5.Text)
Data1.Recordset.Update
Data1.Recordset.Close
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Data1.RecordSource = "select * from mem_mast"
Form1.Data1.Refresh
Form2.Hide
Form1.Show
End Sub

Private Sub Text1_LostFocus()
On Error GoTo errorhandle
Data1.RecordSource = "select * from mem_mast where mem_code='" + Text1.Text + "'"
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
MsgBox "Error occurred!Wrong Member Code", vbInformation, "Error"
	Exit Sub ' Use Exit Sub instead of End to prevent application termination
End Sub

Private Sub Text4_GotFocus()
Text4.Text = "500"
End Sub
