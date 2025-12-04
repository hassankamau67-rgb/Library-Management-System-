VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form7"
   ScaleHeight     =   2490
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   0
      Top             =   -2280
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Double
Dim d2 As Date
d2 = Date$
i = DateDiff("d", d1, d2)
MsgBox d1
MsgBox d2
If i > 0 Then
Dim s As String
   s = (i * 5)
   Label2.Caption = "You have a fine of Rs " + s
   Else
   Label2.Caption = "You have no fine"
  End If
End Sub

