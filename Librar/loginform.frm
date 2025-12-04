VERSION 5.00
Begin VB.Form loginform 
   BackColor       =   &H0080C0FF&
   Caption         =   "Library management system, login form"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9285
   LinkTopic       =   "Form7"
   ScaleHeight     =   6255
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   9840
      TabIndex        =   1
      Top             =   5160
      Width           =   4455
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Login"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "User's name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   480
      Picture         =   "loginform.frx":0000
      Top             =   2760
      Width           =   8925
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"loginform.frx":CDBD
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   15855
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "abhinav" And Text2.Text = "12345" Then
Form1.Show
'Form7.Show
Unload Me
Else
MsgBox "please try again ,may be your user's name and password is wrong"
End If
End Sub

Private Sub Command2_Click()
End
End Sub
