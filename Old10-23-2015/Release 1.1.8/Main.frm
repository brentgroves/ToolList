VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AB31F9DB-A157-4577-BB66-8991776D9488}#1.0#0"; "sptbdock.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Busche Tool List"
   ClientHeight    =   4590
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9540
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main.frx":0ECA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDock.TTabDock TabDock 
      Left            =   120
      Top             =   3720
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
      Begin VB.Menu NewToolList 
         Caption         =   "New Tool List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu CopyToolList 
         Caption         =   "Copy Tool List"
      End
      Begin VB.Menu OpenToolList 
         Caption         =   "Open Tool List"
         Shortcut        =   {F6}
      End
      Begin VB.Menu DeleteToolList 
         Caption         =   "Delete Tool List"
      End
      Begin VB.Menu Configure 
         Caption         =   "Configure"
      End
      Begin VB.Menu CloseToolList 
         Caption         =   "Close"
         Enabled         =   0   'False
         Shortcut        =   {F7}
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
      Visible         =   0   'False
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &1"
         Index           =   0
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &2"
         Index           =   1
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &3"
         Index           =   2
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &4"
         Index           =   3
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &6"
         Index           =   4
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &7"
         Index           =   5
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPanels 
         Caption         =   "Panels"
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



Private Sub CloseToolList_Click()
    ExitLoop = True
    RefreshMenuOptions
    ReportForm.Hide
    ProcessID = 0
    ToolID = 0
    LastToolDescription = ""
    If Len(NotificationMessage) > 0 Then
        NotificationForm.Show
        NotificationForm.Text1.Text = NotificationMessage
    End If
    
End Sub


Private Sub Configure_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    DeleteToolList.Enabled = False
    CloseToolList.Enabled = False
    Configure.Enabled = False
    CopyToolList.Enabled = False
    CopyTool.Enabled = False
    GetEmails
    EmailForm.Show
End Sub

Private Sub CopyTool_Click()
    CarbonCopy.Show
    CarbonCopyOpen
End Sub

Private Sub CopyToolList_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    DeleteToolList.Enabled = False
    CloseToolList.Enabled = False
    Configure.Enabled = False
    CopyToolList.Enabled = False
    CopyTool.Enabled = False
    PopulateCopyView
    CopyProcess.Show
    NotificationMessage = ""
    NotificationSent = False
    OldItemNumber = ""
End Sub

Private Sub DeleteToolList_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    DeleteToolList.Enabled = False
    CloseToolList.Enabled = False
    Configure.Enabled = False
    CopyToolList.Enabled = False
    CopyTool.Enabled = False
    PopulateDeleteView
    DeleteProcess.Show
    NotificationMessage = ""
    NotificationSent = False
    OldItemNumber = ""
End Sub


Private Sub MDIForm_Load()
    TabDock.GrabMain Me.hwnd
    TabDock.AddForm ToolList, tdDocked, tdAlignRight, "Tool List", tdShowInvisible
    TabDock.AddForm ProcessAttr, tdDocked, tdAlignBottom, "Process Details", tdShowInvisible
    TabDock.AddForm ItemAttri, tdDocked, tdAlignBottom, "Item Details", tdShowInvisible
    TabDock.AddForm ToolAttr, tdDocked, tdAlignBottom, "Tool Details", tdShowInvisible
    TabDock.AddForm RevisionForm, tdDocked, tdAlignBottom, "Revision", tdShowInvisible
    TabDock.AddForm MiscItem, tdDocked, tdAlignBottom, "Misc Details", tdShowInvisible
    TabDock.Panels(tdAlignBottom).Height = 4800
    TabDock.Panels(tdAlignRight).Width = 5000
    TabDock.Show
    Init
    Me.Caption = "Busche Tool List V" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub NewToolList_Click()
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    CloseToolList.Enabled = True
    DeleteToolList.Enabled = False
    CopyToolList.Enabled = False
    Configure.Enabled = False
    CopyTool.Enabled = True
    TabDock.FormShow "Tool List"
    AddProcess
    RefreshReport
    BuildToolList
    GetAllPartNumbers
    GetAllPlants
    ReportForm.Show
    TabDock.FormShow "Process Details"
    NotificationMessage = ""
    NotificationSent = False
    OldItemNumber = ""
End Sub

Private Sub OpenToolList_Click()
    openSQLStatement = "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    NewToolList.Enabled = False
    OpenToolList.Enabled = False
    CloseToolList.Enabled = False
    CopyToolList.Enabled = False
    DeleteToolList.Enabled = False
    Configure.Enabled = False
    OpenProcesses
    GetAllPlantsForFilter
    GetAllPartsForFilter
    OpenProcess.Show
    NotificationMessage = ""
    NotificationSent = False
    OldItemNumber = ""
End Sub

Public Sub RefreshMenuOptions()
    NewToolList.Enabled = True
    OpenToolList.Enabled = True
    CopyToolList.Enabled = True
    DeleteToolList.Enabled = True
    Configure.Enabled = True
    CloseToolList.Enabled = False
    TabDock.FormHide "Process Details"
    TabDock.FormHide "Tool Details"
    TabDock.FormHide "Item Details"
    TabDock.FormHide "Tool List"
    CopyTool.Enabled = False
End Sub
