VERSION 5.00
Begin VB.Form KitAttri 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Item Attributes"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   4140
   ScaleWidth      =   14595
   Begin VB.CommandButton RefreshBTN 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox CribNumberIDTXT 
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox ItemNumberCOMBO 
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Text            =   "ItemNumberCOMBO"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Kit Number:"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4155
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   14355
   End
End
Attribute VB_Name = "KitAttri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cl As New Class1

Private Sub CancelBTN_Click()
    ClearKitFields
    MDIForm1.TabDock.FormHide "Add Kit"
End Sub

Private Sub Form_GotFocus()
    ItemNumberCOMBO.SetFocus
End Sub

Private Sub ItemNumberCOMBO_GotFocus()
        Cl.ShowDropDownCombo ItemNumberCOMBO
End Sub

Private Sub RefreshBTN_Click()
    PopulateKitList
End Sub

Private Sub ItemNumberCOMBO_LostFocus()
    ValidateKitNumber
End Sub

Private Sub UpdateBTN_Click()
    If Not ValidateKitNumber Then
        Exit Sub
    End If
    AddKit
    ClearKitFields
    MDIForm1.TabDock.FormHide "Add Kit"
    RefreshReport
End Sub
