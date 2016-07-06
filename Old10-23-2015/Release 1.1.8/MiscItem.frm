VERSION 5.00
Begin VB.Form MiscItem 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Miscellaneous / Hand Tools"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   3345
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox QtyOnHandTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox AdditionalNotesTXT 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CheckBox ConsumableCHECK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Consumable?"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox ToolLifeTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox QuantityTXT 
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox CuttingEdgesTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox CribNumberIDTXT 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox ItemNumberCOMBO 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "ItemNumberCOMBO"
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox ManufacturerTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox ItemGroupTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton CancelBTN 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   9360
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   615
      Left            =   7680
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Qty On Hand:"
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
      TabIndex        =   19
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Additional Notes:"
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
      Left            =   720
      TabIndex        =   18
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Tool Life Per Edge:"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Quantity:"
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
      Left            =   8400
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Number of Cutting Edges:"
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
      Left            =   6720
      TabIndex        =   15
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
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
      Left            =   960
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
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
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
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
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   12255
   End
End
Attribute VB_Name = "MiscItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBTN_Click()
    ClearMiscFields
    MDIForm1.TabDock.FormHide "Misc Details"
End Sub
Private Sub ConsumableCHECK_Click()
 If ConsumableCHECK.Value = 1 Then
    ToolLifeTXT.Enabled = True
    CuttingEdgesTXT.Enabled = True
    ToolLifeTXT.BackColor = &HFFFFFF
    CuttingEdgesTXT.BackColor = &HFFFFFF
Else
    ToolLifeTXT.Text = ""
    CuttingEdgesTXT.Text = ""
    ToolLifeTXT.Enabled = False
    CuttingEdgesTXT.Enabled = False
    ToolLifeTXT.BackColor = &H808080
    CuttingEdgesTXT.BackColor = &H808080
End If
End Sub


Private Sub ItemNumberCOMBO_LostFocus()
    ValidateMiscItemNumber
End Sub

Private Sub UpdateBTN_Click()
    If Not IsNumeric(QuantityTXT.Text) Then
        MsgBox ("Invalid Quantity")
        Exit Sub
    End If
    
    If ConsumableCHECK.Value Then
        If Not IsNumeric(CuttingEdgesTXT.Text) Or Val(CuttingEdgesTXT.Text) = 0 Then
            MsgBox ("Invalid Number of Cutting Edges")
            Exit Sub
        End If

        If Not IsNumeric(ToolLifeTXT.Text) Or Val(CuttingEdgesTXT.Text) = 0 Then
            MsgBox ("Invalid Number of Cutting Edges")
            Exit Sub
        End If
    End If
    
    If misctoolexists Then
        UpdateMiscDetails
    Else
        AddMiscSub
    End If
    ClearMiscFields
    MDIForm1.TabDock.FormHide "Misc Details"
    RefreshReport
End Sub
