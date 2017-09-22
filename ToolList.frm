VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ToolList 
   Caption         =   "Tool List"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6240
   ScaleWidth      =   9105
   Begin VB.CommandButton ReportBtn 
      Caption         =   "Display Report"
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton ReleaseBtn 
      Caption         =   "Submit For Initial Release"
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton ChgRoutingBtn 
      Caption         =   "Create Change Routing"
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   5520
      Width           =   2775
   End
   Begin MSComctlLib.TreeView TreeView4 
      Height          =   1935
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3413
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView3 
      Height          =   1935
      Left            =   3240
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3413
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3413
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1935
      Left            =   4920
      TabIndex        =   0
      Top             =   840
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3413
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tool List"
            Key             =   "Tool"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Misc Tooling"
            Key             =   "Misc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Revision"
            Key             =   "Rev"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fixture"
            Key             =   "Fixture"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu ProcessMenu 
      Caption         =   "Process Menu"
      Visible         =   0   'False
      Begin VB.Menu AddToolFromProcess 
         Caption         =   "Add Tool       F2"
         Index           =   1
      End
      Begin VB.Menu EditProcess 
         Caption         =   "Edit Process  F5"
      End
   End
   Begin VB.Menu ToolMenu 
      Caption         =   "Tool Menu"
      Visible         =   0   'False
      Begin VB.Menu AddToolfromTool 
         Caption         =   "Add Tool       F2"
         Index           =   2
      End
      Begin VB.Menu AddItemFromTool 
         Caption         =   "Add Item      F3"
      End
      Begin VB.Menu AddKitFromTool 
         Caption         =   "Add Kit         F4"
      End
      Begin VB.Menu DeleteTool 
         Caption         =   "Delete Tool"
      End
      Begin VB.Menu EditTool 
         Caption         =   "Edit Tool      F5"
      End
   End
   Begin VB.Menu ItemMenu 
      Caption         =   "ItemMenu"
      Visible         =   0   'False
      Begin VB.Menu AddItemfromItem 
         Caption         =   "Add Item      F3"
      End
      Begin VB.Menu AddKitFromItem 
         Caption         =   "Add Kit         F4"
      End
      Begin VB.Menu DeleteItem 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu EditItem 
         Caption         =   "Edit Item      F5"
      End
   End
   Begin VB.Menu RevisionProcessMenu 
      Caption         =   "Revision Process Menu"
      Visible         =   0   'False
      Begin VB.Menu AddRevisionFromProcess 
         Caption         =   "Add Revision"
      End
   End
   Begin VB.Menu RevisionMenu 
      Caption         =   "Revision Menu"
      Visible         =   0   'False
      Begin VB.Menu AddRevision 
         Caption         =   "Add Revision"
      End
      Begin VB.Menu DeleteRevision 
         Caption         =   "Delete Revision"
      End
   End
   Begin VB.Menu MiscTool 
      Caption         =   "Misc Tool"
      Visible         =   0   'False
      Begin VB.Menu AddMisc 
         Caption         =   "Add Misc Tool"
      End
      Begin VB.Menu DeleteMisc 
         Caption         =   "Delete Misc Tool"
      End
   End
   Begin VB.Menu MiscProcessMenu 
      Caption         =   "Misc Process Menu"
      Visible         =   0   'False
      Begin VB.Menu AddMiscFromProcess 
         Caption         =   "Add Misc Tool"
      End
   End
   Begin VB.Menu FixtureTool 
      Caption         =   "Fixture Tool"
      Visible         =   0   'False
      Begin VB.Menu AddFixture 
         Caption         =   "Add Fixture Tool"
      End
      Begin VB.Menu DeleteFixture 
         Caption         =   "Delete Fixture Tool"
      End
   End
   Begin VB.Menu FixtureProcessMenu 
      Caption         =   "Fixture Process Menu"
      Visible         =   0   'False
      Begin VB.Menu AddFixtureFromProcess 
         Caption         =   "Add Fixture Tool"
      End
   End
End
Attribute VB_Name = "ToolList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub AddItemfromItem_Click()
HideAllForms
    ClearItemFields
    toolID = Val(Trim(Right(TreeView1.SelectedItem.Parent.Key, Len(TreeView1.SelectedItem.Parent.Key) - 4)))
    itemexists = False
    MDIForm1.TabDock.FormShow "Item Details"
End Sub

Private Sub AddItemFromTool_Click()
HideAllForms
    ClearItemFields
    toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    itemexists = False
    MDIForm1.TabDock.FormShow "Item Details"
End Sub

Private Sub AddKitFromItem_Click()
HideAllForms
    ClearKitFields
    toolID = Val(Trim(Right(TreeView1.SelectedItem.Parent.Key, Len(TreeView1.SelectedItem.Parent.Key) - 4)))
    MDIForm1.TabDock.FormShow "Add Kit"
End Sub

Private Sub AddKitFromTool_Click()
HideAllForms
    ClearKitFields
    toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    MDIForm1.TabDock.FormShow "Add Kit"
End Sub

Private Sub AddMisc_Click()
HideAllForms
    ClearMiscFields
    misctoolexists = False
    MDIForm1.TabDock.FormShow "Misc Details"
End Sub
Private Sub Addfixture_Click()
HideAllForms
    ClearFixtureFields
    fixturetoolexists = False
    MDIForm1.TabDock.FormShow "Fixture Details"
End Sub

Private Sub AddMiscFromProcess_Click()
HideAllForms
    ClearMiscFields
    misctoolexists = False
    MDIForm1.TabDock.FormShow "Misc Details"
End Sub
Private Sub AddFixtureFromProcess_Click()
HideAllForms
    ClearFixtureFields
    fixturetoolexists = False
    MDIForm1.TabDock.FormShow "Fixture Details"
End Sub

Private Sub AddRevision_Click()
HideAllForms
    ClearRevisionFields
    revisionexists = False
    MDIForm1.TabDock.FormShow "Revision"
End Sub

Private Sub AddRevisionFromProcess_Click()
HideAllForms
    ClearRevisionFields
    revisionexists = False
    MDIForm1.TabDock.FormShow "Revision"
End Sub

Private Sub AddToolFromProcess_Click(Index As Integer)
HideAllForms
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    DoEvents
    If MultiTurret Then
        ToolAttr.EnableMultiTurret
    Else
        ToolAttr.DisableMultiTurret
    End If
    ClearToolFields
    toolexists = False
    MDIForm1.TabDock.FormShow "Tool Details"
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
    
End Sub

Private Sub AddToolfromTool_Click(Index As Integer)
    HideAllForms
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    DoEvents
    If MultiTurret Then
        ToolAttr.EnableMultiTurret
    Else
        ToolAttr.DisableMultiTurret
    End If
    ClearToolFields
    toolexists = False
    MDIForm1.TabDock.FormShow "Tool Details"
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
End Sub



Private Sub ChgRoutingBtn_Click()
HideAllForms
    ReportForm.Hide
    ReportBtn.Caption = "Display Report"
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    reportOpened = False
    DoEvents
    ' Add the tool changes to the appropriate list and then erase the ToolChanges object
    PopulateChangesForRouting
    CreateRouting.Show
    CreateRouting.SetForCreation
    CreateRouting.ZOrder (0)
    MDIForm1.CopyTool.Enabled = False
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
End Sub

Private Sub DeleteItem_Click()
    itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    DeleteItemSub
End Sub

Private Sub DeleteMisc_Click()
    MiscToolID = Val(Trim(Right(TreeView2.SelectedItem.Key, Len(TreeView2.SelectedItem.Key) - 4)))
    DeleteMiscSub
End Sub
Private Sub DeleteFixture_Click()
    FixtureToolID = Val(Trim(Right(TreeView4.SelectedItem.Key, Len(TreeView4.SelectedItem.Key) - 4)))
    DeleteFixtureSub
End Sub
Private Sub DeleteRevision_Click()
    RevisionID = Val(Trim(Right(TreeView3.SelectedItem.Key, Len(TreeView3.SelectedItem.Key) - 3)))
    DeleteRevisionSub
End Sub

Private Sub DeleteTool_Click()
    toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    DeleteToolSub
End Sub



Private Sub EditItem_Click()
HideAllForms
    itemexists = True
    ClearItemFields
    itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    GetItemDetails
    MDIForm1.TabDock.FormShow "Item Details"
    On Error Resume Next
    If (True = reportOpened) Then
        ReportForm.CRViewer1.ShowFirstPage
        ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Parent.Text, Len(TreeView1.SelectedItem.Parent.Text) - 1 - InStr(TreeView1.SelectedItem.Parent.Text, "-")))
    End If
    
End Sub

Private Sub EditProcess_Click()
HideAllForms
    ClearProcessFields
    GetProcessDetails
    GetAllPartNumbers
    GetAssignedPartNumbers
    GetAllPlants
    GetAssignedPlant
    MDIForm1.TabDock.FormShow "Process Details"
End Sub

Private Sub EditTool_Click()
HideAllForms
    toolexists = True
    toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    ClearToolFields
    GetToolDetails
    MDIForm1.TabDock.FormShow "Tool Details"
    On Error Resume Next
    If (True = reportOpened) Then
        ReportForm.CRViewer1.ShowFirstPage
        ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Text, Len(TreeView1.SelectedItem.Text) - 1 - InStr(TreeView1.SelectedItem.Text, "-")))
    End If
    
End Sub

Private Sub Form_Load()
    Dim hwndTips As Long
    hwndTips = SendMessage(TreeView1.hwnd, TVM_GETTOOLTIPS, 0&, ByVal 0&)
    SendMessage hwndTips, TTM_ACTIVATE, 0&, ByVal 0&
End Sub

Private Sub Form_Resize()
    If MDIForm1.WindowState <> vbMinimized Then
        TreeView1.Top = 360
        TreeView1.Left = 0
        If ScaleHeight > 1000 Then
            TreeView1.Height = ScaleHeight - 1600
            TreeView2.Height = ScaleHeight - 1600
            TreeView3.Height = ScaleHeight - 1600
            TreeView4.Height = ScaleHeight - 1600
        End If
        TreeView1.Width = ScaleWidth
        TreeView2.Top = 360
        TreeView2.Left = 0
        TreeView2.Width = ScaleWidth
        TreeView3.Top = 360
        TreeView3.Left = 0
        TreeView3.Width = ScaleWidth
        TreeView4.Top = 360
        TreeView4.Left = 0
        TreeView4.Width = ScaleWidth
        TabStrip1.Top = 0
        TabStrip1.Left = 0
        If ScaleHeight > 1000 Then
            TabStrip1.Height = ScaleHeight - 1600
        End If
        TabStrip1.Width = ScaleWidth
        If ScaleHeight > 600 Then
            ChgRoutingBtn.Top = ScaleHeight - 1200
        End If
        ChgRoutingBtn.Left = 100
        ChgRoutingBtn.Width = ScaleWidth - 200
        
        If ScaleHeight > 600 Then
            ReportBtn.Top = ScaleHeight - 600
        End If
        ReportBtn.Left = 100
        ReportBtn.Width = ScaleWidth - 200
        
        ReleaseBtn.Left = 100
        ReleaseBtn.Width = ScaleWidth - 200
        If ScaleHeight > 600 Then
            ReleaseBtn.Top = ScaleHeight - 1200
        End If
        
    End If
End Sub



Private Sub ReleaseBtn_Click()
    If MsgBox("Releasing this tool list will require all future changes to go through a routing process. Are you sure you want to continue?", vbYesNo, "Submit for Relase?") = vbYes Then
        SubmitForInitialRelease (processId)
    End If
End Sub

Private Sub ReportBtn_Click()
'Note: If you close the toollist before the report finishes it will display no data because the toollist is deleted unless the
' the modifications are commited.
    Dim i As Integer
    i = 5
    ReportBtn.Enabled = False
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    DoEvents
    OpenCRViewer (processId)
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
'    ReportBtn.Caption = "Refresh Report"
    ReportBtn.Enabled = True
    


End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Key = "Tool" Then
        TreeView1.Visible = True
        TreeView2.Visible = False
        TreeView3.Visible = False
        TreeView4.Visible = False
    ElseIf TabStrip1.SelectedItem.Key = "Misc" Then
        TreeView1.Visible = False
        TreeView2.Visible = True
        TreeView3.Visible = False
        TreeView4.Visible = False
    ElseIf TabStrip1.SelectedItem.Key = "Rev" Then
        TreeView1.Visible = False
        TreeView2.Visible = False
        TreeView3.Visible = True
        TreeView4.Visible = False
    ElseIf TabStrip1.SelectedItem.Key = "Fixture" Then
        TreeView1.Visible = False
        TreeView2.Visible = False
        TreeView3.Visible = False
        TreeView4.Visible = True
    End If
End Sub

Private Sub TreeView1_DblClick()
HideAllForms
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    DoEvents
    If TreeView1.SelectedItem.Parent Is Nothing Then
        ClearProcessFields
        GetProcessDetails
        GetAllPartNumbers
        GetAssignedPartNumbers
        GetAllPlants
        GetAssignedPlant
        MDIForm1.TabDock.FormShow "Process Details"
    ElseIf TreeView1.SelectedItem.Parent.Parent Is Nothing Then
        toolexists = True
        toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        ClearToolFields
        GetToolDetails
        MDIForm1.TabDock.FormShow "Tool Details"
        On Error Resume Next
        If (True = reportOpened) Then
            ReportForm.CRViewer1.ShowFirstPage
            ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Text, Len(TreeView1.SelectedItem.Text) - 1 - InStr(TreeView1.SelectedItem.Text, "-")))
'            ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Parent.Text, Len(TreeView1.SelectedItem.Parent.Text) - 1 - InStr(TreeView1.SelectedItem.Parent.Text, "-")))
        End If
    ElseIf TreeView1.SelectedItem.Parent.Parent.Parent Is Nothing Then
        itemexists = True
        ClearItemFields
        itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        GetItemDetails
        MDIForm1.TabDock.FormShow "Item Details"
        On Error Resume Next
        If (True = reportOpened) Then
            ReportForm.CRViewer1.ShowFirstPage
            ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Parent.Text, Len(TreeView1.SelectedItem.Parent.Text) - 1 - InStr(TreeView1.SelectedItem.Parent.Text, "-")))
        End If
    End If
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
    
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Not TreeView1.SelectedItem.Parent Is Nothing Then
            If Not TreeView1.SelectedItem.Parent.Parent Is Nothing Then
                If MsgBox("Are you sure you want to delete this item?", vbYesNo, "Delete?") = vbYes Then
                    itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
                    DeleteItemSub
                End If
            ElseIf Not TreeView1.SelectedItem.Parent Is Nothing Then
                If MsgBox("Are you sure you want to delete this item?", vbYesNo, "Delete?") = vbYes Then
                    toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
                    DeleteToolSub
                End If
            End If
        End If
    End If
    If KeyCode = vbKeyF2 Then
        ClearToolFields
        toolexists = False
        MDIForm1.TabDock.FormShow "Tool Details"
    End If
    If KeyCode = vbKeyF3 Then
        ClearItemFields
        If TreeView1.SelectedItem.Parent Is Nothing Then
            toolID = 0
            Exit Sub
        End If
        If InStr(TreeView1.SelectedItem.Key, "TOOL") Then
            toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        ElseIf InStr(TreeView1.SelectedItem.Parent.Key, "TOOL") Then
            toolID = Val(Trim(Right(TreeView1.SelectedItem.Parent.Key, Len(TreeView1.SelectedItem.Parent.Key) - 4)))
        Else
            toolID = 0
                Exit Sub
        End If
        itemexists = False
        MDIForm1.TabDock.FormShow "Item Details"
    End If
    If KeyCode = vbKeyF4 Then
        ClearKitFields
        If TreeView1.SelectedItem.Parent Is Nothing Then
            toolID = 0
            Exit Sub
        End If
        If InStr(TreeView1.SelectedItem.Key, "TOOL") Then
            toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        ElseIf InStr(TreeView1.SelectedItem.Parent.Key, "TOOL") Then
            toolID = Val(Trim(Right(TreeView1.SelectedItem.Parent.Key, Len(TreeView1.SelectedItem.Parent.Key) - 4)))
        Else
            toolID = 0
            Exit Sub
        End If
        MDIForm1.TabDock.FormShow "Add Kit"
    End If
    If KeyCode = vbKeyF5 Then
        If TreeView1.SelectedItem.Parent Is Nothing Then
            ClearProcessFields
            GetProcessDetails
            GetAllPartNumbers
            GetAssignedPartNumbers
            GetAllPlants
            GetAssignedPlant
            MDIForm1.TabDock.FormShow "Process Details"
        ElseIf TreeView1.SelectedItem.Parent.Parent Is Nothing Then
            toolexists = True
            toolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
            ClearToolFields
            GetToolDetails
            MDIForm1.TabDock.FormShow "Tool Details"
            On Error Resume Next
            If (True = reportOpened) Then
                ReportForm.CRViewer1.ShowFirstPage
                ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Text, Len(TreeView1.SelectedItem.Text) - 1 - InStr(TreeView1.SelectedItem.Text, "-")))
            End If
        ElseIf TreeView1.SelectedItem.Parent.Parent.Parent Is Nothing Then
            itemexists = True
            ClearItemFields
            itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
            GetItemDetails
            MDIForm1.TabDock.FormShow "Item Details"
            On Error Resume Next
            If (True = reportOpened) Then
                ReportForm.CRViewer1.ShowFirstPage
                ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Parent.Text, Len(TreeView1.SelectedItem.Parent.Text) - 1 - InStr(TreeView1.SelectedItem.Parent.Text, "-")))
            End If
        End If
    End If
End Sub

Private Sub TreeView3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        ClearRevisionFields
        revisionexists = False
        MDIForm1.TabDock.FormShow "Revision"
    End If
End Sub



Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If TreeView1.SelectedItem.Parent Is Nothing Then
            PopupMenu ProcessMenu
        ElseIf TreeView1.SelectedItem.Parent.Parent Is Nothing Then
            PopupMenu ToolMenu
        ElseIf TreeView1.SelectedItem.Parent.Parent.Parent Is Nothing Then
            PopupMenu ItemMenu
        End If
    End If
End Sub

Private Sub TreeView2_DblClick()
HideAllForms
    If TreeView2.SelectedItem.Parent Is Nothing Then
        ProgressBar.Show
        ProgressBar.Timer1.Enabled = True
        DoEvents
        
        ClearProcessFields
        GetProcessDetails
        GetAllPartNumbers
        GetAssignedPartNumbers
        GetAllPlants
        GetAssignedPlant
        
        MDIForm1.TabDock.FormShow "Process Details"
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
        
    ElseIf TreeView2.SelectedItem.Parent.Parent Is Nothing Then
        misctoolexists = True
        MDIForm1.TabDock.FormShow "Misc Details"
        ClearMiscFields
        MiscToolID = Val(Trim(Right(TreeView2.SelectedItem.Key, Len(TreeView2.SelectedItem.Key) - 4)))
        GetMiscDetails
    End If
End Sub

Private Sub TreeView4_DblClick()
HideAllForms
    If TreeView4.SelectedItem.Parent Is Nothing Then
        ProgressBar.Show
        ProgressBar.Timer1.Enabled = True
        DoEvents
        ClearProcessFields
        GetProcessDetails
        GetAllPartNumbers
        GetAssignedPartNumbers
        GetAllPlants
        GetAssignedPlant
        MDIForm1.TabDock.FormShow "Process Details"
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
    ElseIf TreeView4.SelectedItem.Parent.Parent Is Nothing Then
        fixturetoolexists = True
        MDIForm1.TabDock.FormShow "Fixture Details"
        ClearMiscFields
        FixtureToolID = Val(Trim(Right(TreeView4.SelectedItem.Key, Len(TreeView4.SelectedItem.Key) - 4)))
        GetFixtureDetails
    End If
End Sub
Private Sub TreeView3_DblClick()
HideAllForms
    If TreeView3.SelectedItem.Parent Is Nothing Then
        ProgressBar.Show
        ProgressBar.Timer1.Enabled = True
        DoEvents
        ClearProcessFields
        GetProcessDetails
        GetAllPartNumbers
        GetAssignedPartNumbers
        GetAllPlants
        GetAssignedPlant
        MDIForm1.TabDock.FormShow "Process Details"
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
    ElseIf TreeView3.SelectedItem.Parent.Parent Is Nothing Then
        revisionexists = True
        MDIForm1.TabDock.FormShow "Revision"
        ClearRevisionFields
        RevisionID = Val(Trim(Right(TreeView3.SelectedItem.Key, Len(TreeView3.SelectedItem.Key) - 3)))
        GetRevisionDetails
    End If
End Sub
Private Sub TreeView3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If TreeView3.SelectedItem.Parent Is Nothing Then
            PopupMenu RevisionProcessMenu
        ElseIf TreeView3.SelectedItem.Parent.Parent Is Nothing Then
            PopupMenu RevisionMenu
        End If
    End If
End Sub


Private Sub TreeView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If TreeView2.SelectedItem.Parent Is Nothing Then
            PopupMenu MiscProcessMenu
        ElseIf TreeView2.SelectedItem.Parent.Parent Is Nothing Then
            PopupMenu MiscTool
        End If
    End If
End Sub

Private Sub TreeView4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If TreeView4.SelectedItem.Parent Is Nothing Then
            PopupMenu FixtureProcessMenu
        ElseIf TreeView4.SelectedItem.Parent.Parent Is Nothing Then
            PopupMenu FixtureTool
        End If
    End If
End Sub

Private Sub HideAllForms()
    MDIForm1.TabDock.FormHide "Item Details"
    MDIForm1.TabDock.FormHide "Tool Details"
    MDIForm1.TabDock.FormHide "Process Details"
    MDIForm1.TabDock.FormHide "Misc Details"
    MDIForm1.TabDock.FormHide "Revision"
    MDIForm1.TabDock.FormHide "Fixture Details"
End Sub

