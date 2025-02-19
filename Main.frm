VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AB31F9DB-A157-4577-BB66-8991776D9488}#1.0#0"; "sptbdock.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Busche Tool List"
   ClientHeight    =   4590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9540
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main.frx":0ECA
   WindowState     =   2  'Maximized
   Begin TabDock.TTabDock TabDock 
      Left            =   120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4335
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "TabDockHost v 1.0"
            TextSave        =   "TabDockHost v 1.0"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu ViewToolList 
         Caption         =   "View Tool List"
         Shortcut        =   {F6}
      End
      Begin VB.Menu NewToolList 
         Caption         =   "New Tool List"
         Shortcut        =   {F7}
      End
      Begin VB.Menu OpenToolList 
         Caption         =   "Open Tool List"
         Shortcut        =   {F8}
      End
      Begin VB.Menu DeleteToolList 
         Caption         =   "Delete Tool List"
      End
      Begin VB.Menu ChangesMNU 
         Caption         =   "View Pending Approvals / Changes"
      End
      Begin VB.Menu Configure 
         Caption         =   "Configure"
      End
      Begin VB.Menu CloseToolList 
         Caption         =   "Close"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu CopyTool 
      Caption         =   "Copy Tool"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu ViewToolTree 
         Caption         =   "View Tool Tree"
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &2"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &3"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &4"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &6"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &7"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewPanels 
         Caption         =   "Panels"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDocking 
      Caption         =   "&Dock"
      Visible         =   0   'False
      Begin VB.Menu mnuDockForm 
         Caption         =   "Form &1"
         Index           =   0
      End
      Begin VB.Menu mnuDockForm 
         Caption         =   "Form &2"
         Index           =   1
      End
      Begin VB.Menu mnuDockForm 
         Caption         =   "Form &6"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Sub ChangesMNU_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    DeleteToolList.Enabled = False
    CloseToolList.Enabled = False
    Configure.Enabled = False
    CopyTool.Enabled = False
    ViewToolList.Enabled = False
    PopulateMainRoutingList
    ChangeList.Show
    ChangeList.WindowState = 2
    WorkingLive = True
End Sub

Private Sub CloseToolList_Click()
    MDIForm1.TabDock.FormHide "Item Details"
    MDIForm1.TabDock.FormHide "Tool Details"
    MDIForm1.TabDock.FormHide "Process Details"
    MDIForm1.TabDock.FormHide "Misc Details"
    MDIForm1.TabDock.FormHide "Revision"
    MDIForm1.TabDock.FormHide "Fixture Details"

    If Not IsReadyToExit Then
        If MsgBox("You have made changes that are not recorded on a routing. Do you still wish to exit and lose your changes?", vbYesNo, "Lose Unrouted Changes?") = vbNo Then
            Exit Sub
        End If
        DeleteProcessSub (processId)
    End If
    Reset
    ExitLoop = True
    RefreshMenuOptions
    If (True = reportOpened) Then
        ReportForm.Hide
    End If
    reportOpened = False
    If (True = reportViewed) Then
        ViewProcess.ListView1.ListItems.Clear
        ViewProcess.Hide
        ViewProcess.Refresh
    End If
End Sub

Private Sub Configure_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    DeleteToolList.Enabled = False
    CloseToolList.Enabled = False
    ChangesMNU.Enabled = False
    Configure.Enabled = False
    CopyTool.Enabled = False
    ViewToolList.Enabled = False
    GetUsernames
    EmailForm.Show
End Sub

Private Sub CopyTool_Click()
    CarbonCopy.Show
    CarbonCopyOpen
End Sub


Private Sub DeleteToolList_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    DeleteToolList.Enabled = False
    ChangesMNU.Enabled = False
    CloseToolList.Enabled = False
    Configure.Enabled = False
    CopyTool.Enabled = False
    ViewToolList.Enabled = False
    PopulateDeleteView
    DeleteProcess.Show
    OldCribID = ""
End Sub


Private Sub MDIForm_Load()
    TabDock.GrabMain Me.hwnd
    TabDock.AddForm ToolList, tdDocked, tdAlignRight, "Tool List", tdShowInvisible
    TabDock.AddForm ProcessAttr, tdDocked, tdAlignBottom, "Process Details", tdShowInvisible
    TabDock.AddForm ItemAttri, tdDocked, tdAlignBottom, "Item Details", tdShowInvisible
    TabDock.AddForm KitAttri, tdDocked, tdAlignBottom, "Add Kit", tdShowInvisible
    TabDock.AddForm ToolAttr, tdDocked, tdAlignBottom, "Tool Details", tdShowInvisible
    TabDock.AddForm RevisionForm, tdDocked, tdAlignBottom, "Revision", tdShowInvisible
    TabDock.AddForm MiscItem, tdDocked, tdAlignBottom, "Misc Details", tdShowInvisible
    TabDock.AddForm FixtureItem, tdDocked, tdAlignBottom, "Fixture Details", tdShowInvisible
    TabDock.Panels(tdAlignBottom).Height = 5100
    TabDock.Panels(tdAlignRight).Width = 5000
    TabDock.Show
    Init
    Me.Caption = "Busche Tool List V" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
'    If Not CheckVersion Then
'        If (BUILDTYPE = TEST) Then
'            MsgBox ("You have an old verion of BuscheToolList.exe.  Copy the updated version from P:\MIS\Busche Software Installation Files\Tool List\Test into your c:\program files\busche toollist directory.")
'        ElseIf (BUILDTYPE = IND) Then
'            MsgBox ("You have an old verion of BuscheToolList.exe.  Copy the updated version from P:\MIS\Busche Software Installation Files\Tool List\Indiana into your c:\program files\busche toollist directory.")
'        ElseIf (BUILDTYPE = AL) Then
'            MsgBox ("You have an old verion of BuscheToolList.exe.  Copy the updated version from P:\MIS\Busche Software Installation Files\Tool List\Alabama into your c:\program files\busche toollist directory.")
'        End If
'        Unload Me
'    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsReadyToExit Then
        Cancel = 0
    Else
        If MsgBox("You have made changes that are not recorded on a routing. Do you still wish to exit and lose your changes?", vbYesNo, "Lose Unrouted Changes?") = vbNo Then
            Cancel = 1
        Else
            DeleteProcessSub (processId)
            Cancel = 0
        End If
    End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Unload ToolList
    Unload ToolAttr
    Unload RoutingComments
    Unload ProcessAttr
    Unload MiscItem
    Unload KitAttri
    Unload RevisionForm
    Unload FixtureItem
    Unload ItemComments
    Unload ActionDetails
    Unload ProgressBar
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub NewToolList_Click()
'    ProgressBar.Show
'    ProgressBar.Timer1.Enabled = True
    WorkingLive = True
    ToolList.ReleaseBtn.Visible = True
    ToolList.ChgRoutingBtn.Visible = False
    ToolList.ReportBtn.Visible = True
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    CloseToolList.Enabled = True
    DeleteToolList.Enabled = False
    ChangesMNU.Enabled = False
    Configure.Enabled = False
    CopyTool.Enabled = True
    ViewToolList.Enabled = False
    TabDock.FormShow "Tool List"
    AddProcess
    
    RefreshReport
    BuildToolList
    BuildRevList
    BuildMiscList
    BuildFixtureList
    GetAllPartNumbers
    GetAllPlants
  '  ReportForm.Show
    TabDock.FormShow "Process Details"
    OldCribID = ""
 '   ProgressBar.Hide
'    ProgressBar.Timer1.Enabled = False
End Sub

Private Sub OpenToolList_Click()
    openSQLStatement = "SELECT * FROM [TOOLLIST MASTER] WHERE [TOOLLIST MASTER].REVOFPROCESSID = 0 ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    CloseToolList.Enabled = False
    DeleteToolList.Enabled = False
    ViewToolList.Enabled = False
    ChangesMNU.Enabled = False
    Configure.Enabled = False
    OpenProcesses
    GetAllPlantsForFilter
    GetAllPartsForFilter
    OpenProcess.Show
    OldCribID = ""
End Sub

Public Sub RefreshMenuOptions()
    NewToolList.Enabled = True
    OpenToolList.Enabled = True
    ViewToolList.Enabled = False
    DeleteToolList.Enabled = True
    ChangesMNU.Enabled = True
    Configure.Enabled = True
    CloseToolList.Enabled = False
    ViewToolTree.Enabled = False
    ViewToolList.Enabled = True
    TabDock.FormHide "Process Details"
    TabDock.FormHide "Tool Details"
    TabDock.FormHide "Item Details"
    TabDock.FormHide "Tool List"
    'TabDock.FormHide "Fixture Details"
    TabDock.FormHide "Misc Details"
    CopyTool.Enabled = False
    WorkingLive = False
End Sub

Private Sub ViewToolList_Click()
    WorkingLive = True
    openSQLStatement = "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    CloseToolList.Enabled = False
    DeleteToolList.Enabled = False
    ChangesMNU.Enabled = False
    Configure.Enabled = False
    ViewToolTree.Enabled = False
    ViewToolList.Enabled = False
    ViewProcesses
    GetAllPlantsForFilterView
    GetAllPartsForFilterView
    ViewProcess.Show
    OldCribID = ""
End Sub

Private Sub ViewToolTree_Click()
    TabDock.FormShow "Tool List"
End Sub
