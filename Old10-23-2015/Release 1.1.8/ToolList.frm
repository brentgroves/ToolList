VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ToolList 
   Caption         =   "Tool List"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6240
   ScaleWidth      =   8475
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
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   2880
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
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1935
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
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
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4260
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
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
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gill Sans MT"
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
   End
   Begin VB.Menu ToolMenu 
      Caption         =   "Tool Menu"
      Visible         =   0   'False
      Begin VB.Menu AddToolfromTool 
         Caption         =   "Add Tool       F2"
         Index           =   2
      End
      Begin VB.Menu DeleteTool 
         Caption         =   "Delete Tool"
      End
      Begin VB.Menu AddItemFromTool 
         Caption         =   "Add Item      F3"
      End
   End
   Begin VB.Menu ItemMenu 
      Caption         =   "ItemMenu"
      Visible         =   0   'False
      Begin VB.Menu AddItemfromItem 
         Caption         =   "Add Item      F3"
      End
      Begin VB.Menu DeleteItem 
         Caption         =   "Delete Item"
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
End
Attribute VB_Name = "ToolList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AddItemfromItem_Click()
    ClearItemFields
    ToolID = Val(Trim(Right(TreeView1.SelectedItem.Parent.Key, Len(TreeView1.SelectedItem.Parent.Key) - 4)))
    itemexists = False
    MDIForm1.TabDock.FormShow "Item Details"
End Sub

Private Sub AddItemFromTool_Click()
    ClearItemFields
    ToolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    itemexists = False
    MDIForm1.TabDock.FormShow "Item Details"
End Sub

Private Sub AddMisc_Click()
    ClearMiscFields
    misctoolexists = False
    MDIForm1.TabDock.FormShow "Misc Details"
End Sub

Private Sub AddMiscFromProcess_Click()
    ClearMiscFields
    misctoolexists = False
    MDIForm1.TabDock.FormShow "Misc Details"
End Sub

Private Sub AddRevision_Click()
    ClearRevisionFields
    revisionexists = False
    MDIForm1.TabDock.FormShow "Revision"
End Sub

Private Sub AddRevisionFromProcess_Click()
    ClearRevisionFields
    revisionexists = False
    MDIForm1.TabDock.FormShow "Revision"
End Sub

Private Sub AddToolFromProcess_Click(Index As Integer)
    If MultiTurret Then
        ToolAttr.EnableMultiTurret
    Else
        ToolAttr.DisableMultiTurret
    End If
    ClearToolFields
    toolexists = False
    MDIForm1.TabDock.FormShow "Tool Details"
End Sub

Private Sub AddToolfromTool_Click(Index As Integer)
    If MultiTurret Then
        ToolAttr.EnableMultiTurret
    Else
        ToolAttr.DisableMultiTurret
    End If
    ClearToolFields
    toolexists = False
    MDIForm1.TabDock.FormShow "Tool Details"
End Sub

Private Sub DeleteItem_Click()
    itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    DeleteItemSub
End Sub

Private Sub DeleteMisc_Click()
    MiscToolID = Val(Trim(Right(TreeView2.SelectedItem.Key, Len(TreeView2.SelectedItem.Key) - 4)))
    DeleteMiscSub
End Sub

Private Sub DeleteRevision_Click()
    RevisionID = Val(Trim(Right(TreeView3.SelectedItem.Key, Len(TreeView3.SelectedItem.Key) - 3)))
    DeleteRevisionSub
End Sub

Private Sub DeleteTool_Click()
    ToolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
    DeleteToolSub
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
        If ScaleHeight > 400 Then
            TreeView1.Height = ScaleHeight - 360
            TreeView2.Height = ScaleHeight - 360
            TreeView3.Height = ScaleHeight - 360
        End If
        TreeView1.Width = ScaleWidth
        TreeView2.Top = 360
        TreeView2.Left = 0
        TreeView2.Width = ScaleWidth
        TreeView3.Top = 360
        TreeView3.Left = 0
        TreeView3.Width = ScaleWidth
        TabStrip1.Top = 0
        TabStrip1.Left = 0
        TabStrip1.Height = ScaleHeight
        TabStrip1.Width = ScaleWidth
    End If
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Key = "Tool" Then
        TreeView1.Visible = True
        TreeView2.Visible = False
        TreeView3.Visible = False
    ElseIf TabStrip1.SelectedItem.Key = "Misc" Then
        TreeView1.Visible = False
        TreeView2.Visible = True
        TreeView3.Visible = False
    ElseIf TabStrip1.SelectedItem.Key = "Rev" Then
        TreeView1.Visible = False
        TreeView2.Visible = False
        TreeView3.Visible = True
    End If
End Sub

Private Sub TreeView1_DblClick()
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
        ToolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        ClearToolFields
        GetToolDetails
        MDIForm1.TabDock.FormShow "Tool Details"
        On Error Resume Next
        ReportForm.CRViewer1.ShowFirstPage
        ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Text, Len(TreeView1.SelectedItem.Text) - 1 - InStr(TreeView1.SelectedItem.Text, "-")))
    ElseIf TreeView1.SelectedItem.Parent.Parent.Parent Is Nothing Then
        itemexists = True
        ClearItemFields
        itemID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        GetItemDetails
        MDIForm1.TabDock.FormShow "Item Details"
        On Error Resume Next
        ReportForm.CRViewer1.ShowFirstPage
        ReportForm.CRViewer1.SearchForText (Right(TreeView1.SelectedItem.Parent.Text, Len(TreeView1.SelectedItem.Parent.Text) - 1 - InStr(TreeView1.SelectedItem.Parent.Text, "-")))
    End If
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        ClearToolFields
        toolexists = False
        MDIForm1.TabDock.FormShow "Tool Details"
    End If
    If KeyCode = vbKeyF3 Then
        ClearItemFields
        If InStr(TreeView1.SelectedItem.Key, "TOOL") Then
            ToolID = Val(Trim(Right(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 4)))
        Else
            If ToolID = 0 Then
                Exit Sub
            End If
        End If
        itemexists = False
        MDIForm1.TabDock.FormShow "Item Details"
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
    If TreeView2.SelectedItem.Parent Is Nothing Then
        ClearProcessFields
        GetProcessDetails
        GetAllPartNumbers
        GetAssignedPartNumbers
        GetAllPlants
        GetAssignedPlant
        MDIForm1.TabDock.FormShow "Process Details"
    ElseIf TreeView2.SelectedItem.Parent.Parent Is Nothing Then
        misctoolexists = True
        MDIForm1.TabDock.FormShow "Misc Details"
        ClearMiscFields
        MiscToolID = Val(Trim(Right(TreeView2.SelectedItem.Key, Len(TreeView2.SelectedItem.Key) - 4)))
        GetMiscDetails
    End If
End Sub

Private Sub TreeView3_DblClick()
    If TreeView3.SelectedItem.Parent Is Nothing Then
        ClearProcessFields
        GetProcessDetails
        GetAllPartNumbers
        GetAssignedPartNumbers
        GetAllPlants
        GetAssignedPlant
        MDIForm1.TabDock.FormShow "Process Details"
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
