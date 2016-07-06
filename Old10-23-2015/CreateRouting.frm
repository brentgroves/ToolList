VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CreateRouting 
   Caption         =   "Create Routing"
   ClientHeight    =   10020
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CreateRouting.frx":0000
   ScaleHeight     =   10020
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton PrintNewToolListCMD 
      Caption         =   "Print New Tool List"
      Height          =   735
      Left            =   3000
      TabIndex        =   23
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton PrintOldToolListCMD 
      Caption         =   "Print Old Tool List"
      Height          =   735
      Left            =   1560
      TabIndex        =   22
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton PrintTaskListCMD 
      Caption         =   "Print Task List"
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton DenyChangesCMD 
      Caption         =   "Deny"
      Height          =   735
      Left            =   7080
      TabIndex        =   20
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton CompleteRoutingCMD 
      Caption         =   "Complete"
      Height          =   735
      Left            =   7080
      TabIndex        =   19
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton ApproveRoutingCMD 
      Caption         =   "Approve"
      Height          =   735
      Left            =   5160
      TabIndex        =   18
      Top             =   9120
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CreateRouting.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CreateRouting.frx":9F6B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox ReasonTxt 
      Height          =   615
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "CreateRouting.frx":13A6D
      Top             =   1320
      Width           =   8415
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   735
      Left            =   9000
      TabIndex        =   9
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton SendRoutingCMD 
      Caption         =   "Send Routing"
      Height          =   735
      Left            =   5160
      TabIndex        =   8
      Top             =   9120
      Width           =   1815
   End
   Begin MSComctlLib.ListView VolumeChangeList 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NEW VOLUME"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "OLD VOLUME"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "APPROVED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "COMPLETED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FILLER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "COMMENTS"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView StatusChangeList 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CHANGING STATUS TO"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "APPROVED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "COMPLETED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FILLER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FILLER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "COMMENTS"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView ToolingChangeList 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CRIBID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TOOL NUMBER"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ACTION"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "APPROVED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "COMPLETE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "COMMENTS"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView PlantChangeList 
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1720
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NEW PLANTS"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "OLD PLANTS"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "APPROVED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "COMPLETED"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FILLER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "COMMENTS"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label DateLBL 
      Alignment       =   2  'Center
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label UsernameLBL 
      Alignment       =   2  'Center
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label ToolListLBL 
      Alignment       =   2  'Center
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "REASON FOR CHANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   13
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "TOOL LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "PLANT CHANGES"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "VOLUME CHANGES"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TOOLING / FIXTURE CHANGES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   5880
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STATUS CHANGES"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   5055
   End
End
Attribute VB_Name = "CreateRouting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OpenType As String
Private ProcessChgID As Long
Private PrintOldProcessID As Long
Private PrintNewProcessID As Long

Private Sub ApproveRoutingCMD_Click()
    CreateRouting.Hide
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = False
    DoEvents
    ApproveRouting (ProcessChgID)
    SendNeedCompleteNotification (ProcessChgID)
    Reset
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
    MsgBox ("Routing has been issued to Buyer")
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Command2_Click()
    MDIForm1.CopyTool.Enabled = True
    CreateRouting.Hide
    If OpenType <> "Creation" Then
        MDIForm1.RefreshMenuOptions
        Reset
    End If
End Sub
Public Sub SetForCreation()
    ProcessChgID = 0
    OpenType = "Creation"
    ApproveRoutingCMD.Visible = False
    CompleteRoutingCMD.Visible = False
    PrintTaskListCMD.Visible = False
    DenyChangesCMD.Visible = False
    SendRoutingCMD.Visible = True
    PrintNewToolListCMD.Visible = False
    PrintOldToolListCMD.Visible = False
    ReasonTxt.Locked = False
    ReasonTxt.BackColor = &H80000005
'    ReasonTxt.BackColor = &H80000009
End Sub

Public Sub SetForApproval(ProcessChangeID As Long)
    ProcessChgID = ProcessChangeID
    OpenType = "Approval"
    ApproveRoutingCMD.Visible = True
    CompleteRoutingCMD.Visible = False
    DenyChangesCMD.Visible = True
    SendRoutingCMD.Visible = False
    PrintTaskListCMD.Visible = True
    PrintNewToolListCMD.Visible = True
    PrintOldToolListCMD.Visible = True
    ReasonTxt.Locked = True
    ReasonTxt.BackColor = &H8000000F
    PopulateRouting (ProcessChangeID)
    CreateRouting.Show
End Sub

Public Sub SetForCompletion(ProcessChangeID As Long)

    ProcessChgID = ProcessChangeID
    OpenType = "Completion"
    ApproveRoutingCMD.Visible = False
    CompleteRoutingCMD.Visible = True
    PrintTaskListCMD.Visible = True
    DenyChangesCMD.Visible = False
    SendRoutingCMD.Visible = False
    PrintNewToolListCMD.Visible = True
    PrintOldToolListCMD.Visible = True
    ReasonTxt.Locked = True
    ReasonTxt.BackColor = &H8000000F
    RefreshActionList (ProcessChangeID)
    If (bRefreshActionListError = False) Then
        PopulateRouting (ProcessChangeID)
        CreateRouting.Show
    Else
    ' Process just like the user pressed the close button from the CreateRouting form.
        CreateRouting.Hide
        If OpenType <> "Creation" Then
            MDIForm1.RefreshMenuOptions
            Reset
        End If
    End If
End Sub


Public Sub SetForViewing(ProcessChangeID As Long)
    ProcessChgID = ProcessChangeID
    OpenType = "Viewing"
    ApproveRoutingCMD.Visible = False
    CompleteRoutingCMD.Visible = False
    DenyChangesCMD.Visible = False
    SendRoutingCMD.Visible = False
    PrintTaskListCMD.Visible = True
    PrintNewToolListCMD.Visible = True
    PrintOldToolListCMD.Visible = True
    ReasonTxt.Locked = True
    ReasonTxt.BackColor = &H8000000F
    PopulateRouting (ProcessChangeID)
    CreateRouting.Show
End Sub

Private Sub CompleteRoutingCMD_Click()
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = False
    DoEvents
    If AttemptCompleteRouting(ProcessChgID) Then
        SendCompleteNotification (ProcessChgID)
        Reset
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
        MsgBox ("Routing is complete")
        MDIForm1.RefreshMenuOptions
    Else
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
        MsgBox ("All Action items are not complete")
    End If
End Sub

Private Sub DenyChangesCMD_Click()
    CreateRouting.Hide
    If IsInitialRelease(ProcessChgID) Then
        RemoveRevInProcess (ProcessID)
    Else
        RemoveRevInProcess (ProcessID)
        DeleteProcessSub (ProcessID)
    End If
    DeleteProcessChange (ProcessChgID)
    CreateRouting.Hide
    ClearFields
    MDIForm1.RefreshMenuOptions
    MsgBox ("Changes have been denied")
    Reset
End Sub

Private Sub PlantChangeList_DblClick()
    If PlantChangeList.SelectedItem Is Nothing Then
        Exit Sub
    End If
    If OpenType = "Completion" Then
        PopulateActionList (PlantChangeList.SelectedItem.SubItems(6))
    Else
        RoutingComments.setListInfo PlantChangeList, PlantChangeList.SelectedItem.Index, "Comments For Plant Change"
    End If
End Sub

Private Sub PrintNewToolListCMD_Click()
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    If (BUILDTYPE = TEST) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_TEST)
    ElseIf (BUILDTYPE = IND) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_IND)
    ElseIf (BUILDTYPE = AL) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_AL)
    End If
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (PrintNewProcessID)
    craxReport.PrintOut
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
    
End Sub

Private Sub PrintOldToolListCMD_Click()
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    If (BUILDTYPE = TEST) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_TEST)
    ElseIf (BUILDTYPE = IND) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_IND)
    ElseIf (BUILDTYPE = AL) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_AL)
    End If
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (PrintOldProcessID)
    craxReport.PrintOut
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
End Sub

Private Sub PrintTaskListCMD_Click()
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    
    If (BUILDTYPE = TEST) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_CHANGE_TASKS_TEST)
    ElseIf (BUILDTYPE = IND) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_CHANGE_TASKS_IND)
    ElseIf (BUILDTYPE = AL) Then
        Set craxReport = craxApp.OpenReport(TOOLLIST_CHANGE_TASKS_AL)
    End If
    
    craxReport.ParameterFields.GetItemByName("ProcessChangeID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessChangeID").AddCurrentValue (ProcessChgID)
    craxReport.PrintOut
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
End Sub

Private Sub SendRoutingCMD_Click()
    If ReasonTxt.Text = "" Then
        MsgBox ("You must enter a reason for your changes.")
        Exit Sub
    End If
    WriteRouting
End Sub

Private Sub StatusChangeList_DblClick()
    If StatusChangeList.SelectedItem Is Nothing Then
        Exit Sub
    End If
    If OpenType = "Completion" Then
        PopulateActionList (StatusChangeList.SelectedItem.SubItems(6))
    Else
        RoutingComments.setListInfo StatusChangeList, StatusChangeList.SelectedItem.Index, "Comments For Status Change"
    End If
End Sub

Private Sub ToolingChangeList_DblClick()
    If ToolingChangeList.SelectedItem Is Nothing Then
        Exit Sub
    End If
        'PopulateActionList (ToolingChangeList.SelectedItem.SubItems(6))
    PopulateItemChangeInfo ProcessChgID, ToolingChangeList.SelectedItem.Text
    ItemComments.setListInfo ToolingChangeList, ToolingChangeList.SelectedItem.Index, "Comments For: " + ToolingChangeList.SelectedItem.SubItems(1)
End Sub

Private Sub VolumeChangeList_DblClick()
    If VolumeChangeList.SelectedItem Is Nothing Then
        Exit Sub
    End If
    If OpenType = "Completion" Then
        PopulateActionList (VolumeChangeList.SelectedItem.SubItems(6))
    Else
        RoutingComments.setListInfo VolumeChangeList, VolumeChangeList.SelectedItem.Index, "Comments For Volume Change"
    End If
End Sub

Function GetViewingType() As String
    GetViewingType = OpenType
End Function

Public Sub ClearFields()
    CreateRouting.ToolingChangeList.ListItems.Clear
    CreateRouting.VolumeChangeList.ListItems.Clear
    CreateRouting.PlantChangeList.ListItems.Clear
    CreateRouting.StatusChangeList.ListItems.Clear
    CreateRouting.DateLBL.Caption = ""
    CreateRouting.UsernameLBL.Caption = ""
    CreateRouting.ToolListLBL.Caption = ""
    CreateRouting.ReasonTxt.Text = ""
End Sub
Public Sub SetProcessIDs(nid As Long, oid As Long)
    PrintOldProcessID = oid
    PrintNewProcessID = nid
End Sub
