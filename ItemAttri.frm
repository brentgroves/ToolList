VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ItemAttri 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Item Attributes"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   4710
   ScaleWidth      =   14595
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox cbDeletePic 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Delete Picture?"
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
      Left            =   7440
      TabIndex        =   39
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton SelectPicBTN 
      Caption         =   "..."
      Height          =   375
      Left            =   9480
      TabIndex        =   37
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtPicture 
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   3720
      Width           =   5895
   End
   Begin VB.CheckBox TBStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Force Toolboss Stock"
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
      Left            =   6480
      TabIndex        =   35
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton RefreshBTN 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recalc"
      Height          =   495
      Left            =   12480
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox MonthlyUsageTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox CostPerPartTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox CostTXT 
      BackColor       =   &H80000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox BinTxt 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox NumofRegrindsTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox ToolLifeRegrindTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox QtyOnHandTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox CribNumberIDTXT 
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox ItemGroupTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox ManufacturerTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox CuttingEdgesTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox QuantityTXT 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
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
      Left            =   12000
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton UpdateBTN 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   495
      Left            =   10560
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox ToolLifeTXT 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox AdditionalNotesTXT 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   2895
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
      Left            =   1750
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
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
      Left            =   1750
      TabIndex        =   6
      Top             =   2475
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808080&
      Caption         =   "Picture:"
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
      Left            =   2400
      TabIndex        =   38
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgItem 
      Height          =   1545
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1770
   End
   Begin VB.Shape Shape2 
      Height          =   2775
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Cost"
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
      TabIndex        =   30
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Monthly Usage"
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
      Left            =   9480
      TabIndex        =   33
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Cost Per Part"
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
      Left            =   9480
      TabIndex        =   32
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Bins:"
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
      TabIndex        =   28
      Top             =   600
      Width           =   615
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
      Left            =   360
      TabIndex        =   26
      Top             =   3350
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
      Left            =   360
      TabIndex        =   25
      Top             =   2920
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
      Left            =   9000
      TabIndex        =   24
      Top             =   240
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
      Left            =   9600
      TabIndex        =   19
      Top             =   1320
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
      Left            =   360
      TabIndex        =   18
      Top             =   2160
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
      Left            =   2040
      TabIndex        =   17
      Top             =   1080
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
      Left            =   360
      TabIndex        =   16
      Top             =   1800
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
      Left            =   1320
      TabIndex        =   15
      Top             =   480
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
      Left            =   1680
      TabIndex        =   14
      Top             =   120
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
      Left            =   9600
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4395
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   14355
   End
End
Attribute VB_Name = "ItemAttri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Private Declare Function GetDesktopWindow Lib "user32" () As Long
'     Private m_sSampleText As String

'      Const SW_SHOWNORMAL = 1
'
'  Const SE_ERR_FNF = 2&
'  Const SE_ERR_PNF = 3&
'  Const SE_ERR_ACCESSDENIED = 5&
'   Const SE_ERR_OOM = 8&
'   Const SE_ERR_DLLNOTFOUND = 32&
'   Const SE_ERR_SHARE = 26&
'   Const SE_ERR_ASSOCINCOMPLETE = 27&
'   Const SE_ERR_DDETIMEOUT = 28&
'   Const SE_ERR_DDEFAIL = 29&
'   Const SE_ERR_DDEBUSY = 30&
'    Const SE_ERR_NOASSOC = 31&
'     Const ERROR_BAD_FORMAT = 11&
Private WithEvents m_cHookDlg As cCommonDialog
Attribute m_cHookDlg.VB_VarHelpID = -1

Dim Cl As New Class1

Private Sub CancelBTN_Click()
    MDIForm1.TabDock.FormHide "Item Details"
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()
 ValidateItemNumber
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
    If Not itemexists Then
        Cl.ShowDropDownCombo ItemNumberCOMBO
    End If
End Sub

Private Sub RefreshBTN_Click()
    PopulateItemList
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
    If ItemNumberCOMBO.Text <> "" Then
        ValidateItemNumber
    End If
End Sub

Private Sub SelectPicBTN_Click()
On Error GoTo cmdClassError
    Dim StartDirectory As String
    StartDirectory = "c:\"
    
    With CommonDialog1
        .DialogTitle = "Choose Document To Load."
        .CancelError = True
        .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
        .InitDir = StartDirectory
        .Filter = "Supported Documents (*.JPG)|*.JPG"
        .FilterIndex = 1
        .ShowOpen
        If Len(.FileName) <> 0 Then
           Me.txtPicture.Text = .FileName
           Me.imgItem.Picture = LoadPicture(.FileName)
        End If
    End With

    
    Exit Sub

cmdClassError:
    If (Err.Number <> 20001) Then
        MsgBox "Error: " & Err.Description
    End If
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
