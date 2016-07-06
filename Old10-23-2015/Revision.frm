VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RevisionForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Revision"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   3720
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox RevByTXT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker RevDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16449537
      CurrentDate     =   38816
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox RevDescTXT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   735
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   6495
   End
   Begin VB.TextBox RevNumTXT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   405
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Revision By:"
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
      Left            =   600
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Revision Description:"
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
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Revision Date:"
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
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Revision #:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "RevisionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelBTN_Click()
    MDIForm1.TabDock.FormHide "Revision"
End Sub

Private Sub Form_GotFocus()
    RevNumTXT.SetFocus
End Sub



Private Sub UpdateBTN_Click()
    If Not IsNumeric(RevNumTXT.Text) Then
        MsgBox ("Revision number must be numeric.")
        Exit Sub
    End If
    
    If revisionexists Then
        UpdateRevisionDetails
    Else
        AddRevisionSub
    End If
    ClearRevisionFields
    MDIForm1.TabDock.FormHide "REVISION"
    RefreshReport
End Sub


