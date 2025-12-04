VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Issue Book"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7410
   Icon            =   "lib4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5535
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "lib4.frx":0442
      Left            =   6000
      List            =   "lib4.frx":0444
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Return"
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Isuue"
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Text            =   "1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Issue Book"
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
      Left            =   2040
      TabIndex        =   20
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "Stock in hand"
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
      Left            =   4200
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Select the Book from the list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Date of Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Date of Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
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
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
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
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
i = 1
Data1.RecordSource = "select * from issue_mast"
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = CDate(Text5.Text)
Data1.Recordset.Fields(5) = CDate(Text6.Text)
Data1.Recordset.Fields(6) = i
Data1.Recordset.Update
Data1.Recordset.Close
Form1.Data2.RecordSource = "select * from book_mast where book_code='" + Text3.Text + "'"
Form1.Data2.Refresh
Form1.Data2.Recordset.Edit
Form4.Text8.Text = Form1.Data2.Recordset.Fields(4)
Text8.Text = Val(Text8.Text) - i
Form1.Data2.Recordset.Fields(4) = Val(Text8.Text)
Form1.Data2.Recordset.Update
Form1.Data2.Recordset.Close
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = i
End Sub

Private Sub Command2_Click()
Form1.Data3.RecordSource = "select * from issue_mast"
Form1.Data3.Refresh
Form4.Hide
Form1.Show
End Sub

Private Sub List1_Click()
Text4.Text = List1.Text
Form1.Data2.RecordSource = "select * from book_mast where book_name='" + List1.Text + "'"
Form1.Data2.Refresh
Do While Not Form1.Data2.Recordset.EOF
Form4.Text3.Text = Form1.Data2.Recordset.Fields(0)
Form4.Text8.Text = Form1.Data2.Recordset.Fields(4)
Form1.Data2.Recordset.MoveNext
Loop
Form1.Data2.Recordset.Close

End Sub

Private Sub Text1_LostFocus()
On Error GoTo errorhandle
Data1.RecordSource = "select * from issue_mast where mem_code='" + Text1.Text + "'"
Data1.Refresh
Do While Not Data1.Recordset.EOF
c = c + 1
Data1.Recordset.MoveNext
Loop
If c <> 0 Then
MsgBox "Please return the book first", vbInformation, "Book Return"
Text1.Text = ""
Form4.Hide
Form1.Show
Else
Form1.Data1.RecordSource = "select * from mem_mast where mem_code='" + Text1.Text + "'"
Form1.Data1.Refresh
Form1.Data1.Recordset.Edit
Text2.Text = Form1.Data1.Recordset.Fields(1)
Form1.Data1.Recordset.Update
Form1.Data1.Recordset.Close
List1.Clear
Form1.Data2.RecordSource = "select * from book_mast"
Form1.Data2.Refresh
Do While Not Form1.Data2.Recordset.EOF
Form4.List1.AddItem (Form1.Data2.Recordset.Fields(1))
Form1.Data2.Recordset.MoveNext
Loop
Form1.Data2.Recordset.Close
End If
Exit Sub
errorhandle:
MsgBox "Error occured!Wrong Member Code", vbInformation, "Error"
End
End Sub

Private Sub Text5_GotFocus()
Text5.Text = Date
End Sub

Private Sub Text7_LostFocus()
If Val(Text8.Text) > Val(Text7.Text) Then
Text8.Text = Val(Text8.Text) - Val(Text7.Text)
Command1.SetFocus
Else
MsgBox ("Please check your stock")
Text7.Text = " "
End If
End Sub
