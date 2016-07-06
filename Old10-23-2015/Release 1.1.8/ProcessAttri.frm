VERSION 5.00
Begin VB.Form ProcessAttr 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Process Attributes"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   4290
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox MultiTurretLathe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "2 Turret Lathe:"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CheckBox ApprovedCheck 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Approved:"
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
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton RemovePlantBTN 
      Caption         =   ">"
      Height          =   375
      Left            =   11400
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton AddPlantBTN 
      Caption         =   "<"
      Height          =   375
      Left            =   11400
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3000
      Width           =   255
   End
   Begin VB.ListBox AllPlantList 
      Height          =   1035
      Left            =   11760
      MultiSelect     =   2  'Extended
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2175
   End
   Begin VB.ListBox SelectedPlantsList 
      Height          =   1035
      Left            =   9120
      MultiSelect     =   2  'Extended
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox AnnualVolumeTXT 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
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
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox CustomerTXT 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CheckBox ObsoleteCheck 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Obsolete:"
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
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton CancelBTN 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton RemovePartBTN 
      Caption         =   ">"
      Height          =   615
      Left            =   11400
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton AddPartBTN 
      Caption         =   "<"
      Height          =   615
      Left            =   11400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.ListBox AllPartNumbersList 
      Height          =   2010
      Left            =   11760
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.ListBox SelectedPartsList 
      Height          =   2010
      Left            =   9120
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
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
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   6135
   End
   Begin VB.TextBox OpNumTXT 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox PartFamilyTXT 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Plant List"
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
      Left            =   11520
      TabIndex        =   23
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Selected Plants"
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
      Left            =   9000
      TabIndex        =   21
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Annual Volume:"
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
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Customer:"
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
      TabIndex        =   18
      Top             =   2160
      Width           =   1455
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
      Left            =   11760
      TabIndex        =   17
      Top             =   240
      Width           =   2175
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
      Left            =   9000
      TabIndex        =   16
      Top             =   240
      Width           =   2535
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
      Left            =   840
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Operation #:"
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
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Part Family:"
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
      Left            =   840
      TabIndex        =   9
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   14175
   End
End
Attribute VB_Name = "ProcessAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub AddPlantBTN_Click()
    Dim i, j As Integer
    i = 0
    j = 0
    Dim InList As Boolean
    While i < AllPlantList.ListCount
        If AllPlantList.Selected(i) Then
            While j < SelectedPlantsList.ListCount
                If AllPlantList.List(i) = AllPlantList.List(j) Then
                    InList = True
                End If
                j = j + 1
            Wend
            If Not InList Then
                SelectedPlantsList.AddItem (AllPlantList.List(i))
            End If
            j = 0
        End If
        i = i + 1
    Wend
End Sub

Private Sub CancelBTN_Click()
    ClearProcessFields
    MDIForm1.TabDock.FormHide "Process Details"
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Form_GotFocus()
    PartFamilyTXT.SetFocus
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
Private Sub RemovePlantBTN_Click()
    Dim i As Integer
        While i < SelectedPlantsList.ListCount
            If SelectedPlantsList.Selected(i) Then
                SelectedPlantsList.RemoveItem (i)
            End If
            i = i + 1
        Wend
End Sub

Private Sub UpdateBTN_Click()
    If Not IsNumeric(AnnualVolumeTXT.Text) Then
        MsgBox ("Invalid Annual Volume")
        Exit Sub
    End If
    UpdateProcessDetails
    UpdatePartNumbers
    UpdatePlants
    ClearProcessFields
    MDIForm1.TabDock.FormHide "Process Details"
    RefreshReport
End Sub

