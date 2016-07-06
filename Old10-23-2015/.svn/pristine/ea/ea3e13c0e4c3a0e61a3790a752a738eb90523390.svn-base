VERSION 5.00
Begin VB.Form EmailForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Email Addresses"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7020
   Begin VB.TextBox Email4 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox Email3 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Email2 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Email1 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox EmailBuyer 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox EmailDept 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Email Addresses"
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
      TabIndex        =   14
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Engineering4 Email:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Engineering3 Email:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Engineering2 Email:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Engineering1 Email:"
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
      TabIndex        =   10
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Buyer Email:"
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
      TabIndex        =   9
      Top             =   1200
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
      TabIndex        =   8
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "EmailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    UpdateEmails
    EmailForm.Hide
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Command2_Click()
    EmailForm.Hide
    MDIForm1.RefreshMenuOptions
End Sub
