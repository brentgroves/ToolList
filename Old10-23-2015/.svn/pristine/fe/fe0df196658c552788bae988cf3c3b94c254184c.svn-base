VERSION 5.00
Begin VB.Form EmailForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Email Addresses"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7020
   Begin VB.TextBox AdminTXT 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox BuyerTXT 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox ManagerTXT 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "USERNAMES "
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
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Admin"
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
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Buyer:"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Department Manager:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "EmailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    UpdateUsernames
    EmailForm.Hide
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Command2_Click()
    EmailForm.Hide
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Command3_Click()
    ProgressBar.Show
End Sub
