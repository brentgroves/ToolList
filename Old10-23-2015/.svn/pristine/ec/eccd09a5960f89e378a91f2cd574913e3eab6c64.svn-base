VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ActionDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Action Details"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ChangeIDTXT 
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   6600
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   6600
      TabIndex        =   14
      Top             =   1080
      Width           =   1695
   End
   Begin MSComctlLib.ListView ActionItemList 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Actions Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Action Text"
         Object.Width           =   14817
      EndProperty
   End
   Begin VB.TextBox Line4TXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox ActionTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Line1TXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox CommentsTXT 
      Height          =   615
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Line3TXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Line2TXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Actions:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Line4Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Usage:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Action:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Line2Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Item Group:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Line1Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Item Number:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   -240
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Line3Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Manufacturer:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "ActionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GWL_HWNDPARENT = (-8)
Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal wNewLong As Long) As Long
Private hParentWnd As Long

Private Sub Command1_Click()
    UpdateActionItems
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    hParentWnd = SetWindowLong(Me.hwnd, GWL_HWNDPARENT, MDIForm1.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetWindowLong(Me.hwnd, GWL_HWNDPARENT, hParentWnd)
End Sub
