VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ToolAttr 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tool Attributes"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   3690
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox OpDescTXT 
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
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox OffsetNumberTXT 
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
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox ToolLengthOffsetTXT 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   6840
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton TurretBOption 
      BackColor       =   &H8000000C&
      Caption         =   "Turret B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton TurretAOption 
      BackColor       =   &H8000000C&
      Caption         =   "Turret A"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox SequenceTxt 
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
      Left            =   12600
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox AdjustedVolume 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CheckBox PartNumberCheck 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Part # Specific:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ListBox SelectedPartsList 
      Height          =   1230
      Left            =   5760
      MultiSelect     =   2  'Extended
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.ListBox AllPartNumbersList 
      Height          =   1230
      Left            =   8400
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton AddPartBTN 
      Caption         =   "<"
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton RemovePartBTN 
      Caption         =   ">"
      Height          =   615
      Left            =   8040
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox AlternateCHECK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Alternate:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton CancelBTN 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   8160
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox ToolNumberTXT 
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
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin MSComctlLib.ListView SequenceList 
      Height          =   2055
      Left            =   10800
      TabIndex        =   19
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3625
      View            =   3
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gill Sans MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sequence"
         Object.Width           =   1782
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tool #"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Op Description"
         Object.Width           =   3263
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Offset Number:"
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
      TabIndex        =   25
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Tool Length Offset:"
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
      Left            =   4440
      TabIndex        =   24
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Tool Sequence #"
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
      Left            =   10560
      TabIndex        =   21
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Current Tool Sequence"
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
      Height          =   495
      Left            =   11400
      TabIndex        =   20
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Adjusted Volume:"
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
      Left            =   360
      TabIndex        =   18
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Selected Part Number"
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
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Part Number List"
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
      Height          =   495
      Left            =   8400
      TabIndex        =   16
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Operation Description:"
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
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Tool Number:"
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
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   14535
   End
End
Attribute VB_Name = "ToolAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBTN_Click()
    ClearProcessFields
    MDIForm1.TabDock.FormHide "Tool Details"
End Sub

Private Sub Form_GotFocus()
    ToolNumberTXT.SetFocus
End Sub


Private Sub PartNumberCheck_Click()
    If PartNumberCheck.Value = 1 Then
        EnableMultiPart
    Else
        DisableMultiPart
    End If
End Sub

Private Sub TurretAOption_Click()
    TurretAOption.Value = True
    TurretBOption.Value = False
End Sub

Private Sub TurretBOption_Click()
    TurretBOption.Value = True
    TurretAOption.Value = False
End Sub

Private Sub UpdateBTN_Click()
    If Not IsNumeric(AdjustedVolume.Text) And Len(AdjustedVolume.Text) > 0 Then
       MsgBox ("Invalid Adjusted Volume")
       Exit Sub
    End If
    If Len(OpDescTXT.Text) > 74 Then
        MsgBox ("Operation Description is tool long (max 75 characters)")
        Exit Sub
    End If
    If Not IsNumeric(ToolNumberTXT.Text) Then
       MsgBox ("Invalid Tool Number")
       Exit Sub
    End If
    If Not IsNumeric(SequenceTxt.Text) Then
       MsgBox ("Invalid Sequence Number")
       Exit Sub
    End If
    If toolexists Then
        UpdateToolDetails
    Else
        AddToolSub
    End If
    ClearToolFields
    MDIForm1.TabDock.FormHide "Tool Details"
    RefreshReport
End Sub
Public Sub DisableMultiPart()
    AdjustedVolume.Text = ""
    AdjustedVolume.Enabled = False
    AdjustedVolume.BackColor = &HC0C0C0
    SelectedPartsList.Enabled = False
    SelectedPartsList.BackColor = &HC0C0C0
    SelectedPartsList.Clear
    AllPartNumbersList.Enabled = False
    AllPartNumbersList.BackColor = &HC0C0C0
    AddPartBTN.Enabled = False
    RemovePartBTN.Enabled = False
End Sub

Public Sub EnableMultiPart()
    AdjustedVolume.Text = ""
    AdjustedVolume.Enabled = True
    AdjustedVolume.BackColor = &HFFFFFF
    SelectedPartsList.Enabled = True
    SelectedPartsList.BackColor = &HFFFFFF
    AllPartNumbersList.Enabled = True
    AllPartNumbersList.BackColor = &HFFFFFF
    AddPartBTN.Enabled = True
    RemovePartBTN.Enabled = True
    
End Sub
Private Sub RemovePartBTN_Click()
    Dim i As Integer
        While i < SelectedPartsList.ListCount
            If SelectedPartsList.Selected(i) Then
                SelectedPartsList.RemoveItem (i)
            End If
            i = i + 1
        Wend
End Sub
Private Sub AddPartBTN_Click()
    Dim i, j As Integer
    i = 0
    j = 0
    Dim InList As Boolean
    While i < AllPartNumbersList.ListCount
        If AllPartNumbersList.Selected(i) Then
            While j < SelectedPartsList.ListCount
                If AllPartNumbersList.List(i) = SelectedPartsList.List(j) Then
                    InList = True
                End If
                j = j + 1
            Wend
            If Not InList Then
                SelectedPartsList.AddItem (AllPartNumbersList.List(i))
            End If
            j = 0
        End If
        i = i + 1
    Wend
End Sub

Public Sub EnableMultiTurret()
    TurretAOption.Enabled = True
    TurretBOption.Enabled = True
    TurretAOption.Visible = True
    TurretBOption.Visible = True
    TurretAOption_Click
End Sub

Public Sub DisableMultiTurret()
    TurretAOption.Enabled = False
    TurretBOption.Enabled = False
    TurretAOption.Visible = False
    TurretBOption.Visible = False
End Sub
