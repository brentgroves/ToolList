VERSION 5.00
Begin VB.Form ItemAttri 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Item Attributes"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3300
   ScaleWidth      =   14595
   Begin VB.TextBox NumofRegrindsTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   12720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox ToolLifeRegrindTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   12720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox RegrindableChk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Regrindable?"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox QtyOnHandTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox CribNumberIDTXT 
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox ItemGroupTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox ManufacturerTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox CuttingEdgesTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox QuantityTXT 
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox ItemNumberCOMBO 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "ItemNumberCOMBO"
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton CancelBTN 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8400
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox ToolLifeTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1575
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
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox AdditionalNotesTXT 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Number of Regrinds:"
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
      Left            =   9720
      TabIndex        =   24
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Tool Life on Regrinds:"
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
      Left            =   9720
      TabIndex        =   23
      Top             =   1920
      Width           =   2895
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
      TabIndex        =   22
      Top             =   720
      Width           =   2175
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
      Left            =   1080
      TabIndex        =   17
      Top             =   1680
      Width           =   1575
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
      Left            =   5640
      TabIndex        =   16
      Top             =   1920
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
      Left            =   7320
      TabIndex        =   15
      Top             =   360
      Width           =   1215
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
      Left            =   5640
      TabIndex        =   14
      Top             =   1320
      Width           =   2895
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
      TabIndex        =   13
      Top             =   2160
      Width           =   1935
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
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   1575
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
      Left            =   1080
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2955
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   14235
   End
End
Attribute VB_Name = "ItemAttri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cl As New Class1

Private Sub CancelBTN_Click()
    ClearItemFields
    MDIForm1.TabDock.FormHide "Item Details"
End Sub

Private Sub ConsumableCHECK_Click()
 If ConsumableCHECK.Value = 1 Then
    ToolLifeTXT.Enabled = True
    CuttingEdgesTXT.Enabled = True
    ToolLifeTXT.BackColor = &HFFFFFF
    CuttingEdgesTXT.BackColor = &HFFFFFF
    RegrindableChk.Enabled = True
    ToolLifeTXT.TabStop = True
    CuttingEdgesTXT.TabStop = True
Else
    ToolLifeTXT.Text = ""
    CuttingEdgesTXT.Text = ""
    ToolLifeTXT.Enabled = False
    CuttingEdgesTXT.Enabled = False
    ToolLifeTXT.BackColor = &H80000000
    CuttingEdgesTXT.BackColor = &H80000000
    RegrindableChk.Enabled = False
    ToolLifeTXT.TabStop = False
    CuttingEdgesTXT.TabStop = False
End If
End Sub


Private Sub Form_GotFocus()
    ItemNumberCOMBO.SetFocus
End Sub


Private Sub ItemNumberCOMBO_GotFocus()
   Cl.ShowDropDownCombo ItemNumberCOMBO
End Sub

Private Sub RegrindableCHK_Click()
 If RegrindableChk.Value = 1 Then
    ToolLifeRegrindTXT.Enabled = True
    NumofRegrindsTXT.Enabled = True
    ToolLifeRegrindTXT.BackColor = &HFFFFFF
    NumofRegrindsTXT.BackColor = &HFFFFFF
    ToolLifeRegrindTXT.TabStop = True
    NumofRegrindsTXT.TabStop = True
Else
    ToolLifeRegrindTXT.Text = ""
    NumofRegrindsTXT.Text = ""
    ToolLifeRegrindTXT.Enabled = False
    NumofRegrindsTXT.Enabled = False
    ToolLifeRegrindTXT.BackColor = &H80000000
    NumofRegrindsTXT.BackColor = &H80000000
    ToolLifeRegrindTXT.TabStop = False
    NumofRegrindsTXT.TabStop = False
End If
End Sub

Private Sub ItemNumberCOMBO_LostFocus()
    ValidateItemNumber
End Sub

Private Sub UpdateBTN_Click()
    If Not ValidateItemNumber Then
        Exit Sub
    End If
    
    If Not IsNumeric(QuantityTXT.Text) Or Val(QuantityTXT.Text) = 0 Then
        MsgBox ("Invalid Quantity")
        Exit Sub
    End If
    
    If ConsumableCHECK.Value Then
        If Not IsNumeric(CuttingEdgesTXT.Text) Or Val(CuttingEdgesTXT.Text) = 0 Then
            MsgBox ("Invalid Number of Cutting Edges")
            Exit Sub
        End If

        If Not IsNumeric(ToolLifeTXT.Text) Or Val(ToolLifeTXT.Text) = 0 Then
            MsgBox ("Invalid Tool Life")
            Exit Sub
        End If
    End If
    If RegrindableChk.Value Then
        If Not IsNumeric(NumofRegrindsTXT.Text) Or Val(NumofRegrindsTXT.Text) = 0 Then
            MsgBox ("Invalid Number of Regrinds")
            Exit Sub
        End If

        If Not IsNumeric(ToolLifeRegrindTXT.Text) Or Val(ToolLifeRegrindTXT.Text) = 0 Then
            MsgBox ("Invalid Regrindable Tool Life")
            Exit Sub
        End If
    End If
    If itemexists Then
        UpdateItemDetails
    Else
        AddItemSub
    End If
    ClearItemFields
    MDIForm1.TabDock.FormHide "Item Details"
    RefreshReport
End Sub
