VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form OpenProcess 
   Caption         =   "Select Process to Open"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   10140
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Filter Results"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox PartListCombo 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.ComboBox PlantListCombo 
      Height          =   315
      ItemData        =   "Select Process.frx":0000
      Left            =   360
      List            =   "Select Process.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Part Family"
         Object.Width           =   10760
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Operation Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Op Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Approved"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Obsolete"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Obsolete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   8
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Approved and Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   13440
      Top             =   9480
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   11400
      Top             =   9480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   8760
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number:"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Plant_Label 
      Caption         =   "Plant:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Not Approved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   9
      Top             =   9600
      Width           =   1815
   End
End
Attribute VB_Name = "OpenProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If (PlantListCombo.Text = "" Or PlantListCombo.Text = "All") And (PartListCombo.Text = "" Or PartListCombo.Text = "All") Then
        openSQLStatement = "SELECT * FROM [TOOLLIST MASTER] WHERE [TOOLLIST MASTER].REVOFPROCESSID = 0 ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    ElseIf (PlantListCombo.Text = "" Or PlantListCombo.Text = "All") And (PartListCombo.Text <> "" And PartListCombo.Text <> "All") Then
        openSQLStatement = "SELECT [TOOLLIST MASTER].PROCESSID, CUSTOMER, " & _
        "PARTFAMILY, OPERATIONDESCRIPTION, OPERATIONNUMBER, PLANT, RELEASED, OBSOLETE FROM " & _
        "[TOOLLIST MASTER],[TOOLLIST PLANT],[TOOLLIST PARTNUMBERS] " & _
        "WHERE [TOOLLIST MASTER].PROCESSID = [TOOLLIST PARTNUMBERS].PROCESSID " & _
        "AND [TOOLLIST MASTER].PROCESSID = [TOOLLIST PLANT].PROCESSID AND " & _
        "PARTNUMBERS = '" + Trim(PartListCombo.Text) + "' AND [TOOLLIST MASTER].REVOFPROCESSID = 0 ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    ElseIf (PartListCombo.Text = "" Or PartListCombo.Text = "All") And (PlantListCombo.Text <> "" And PlantListCombo.Text <> "All") Then
        openSQLStatement = "SELECT DISTINCT [TOOLLIST MASTER].PROCESSID, CUSTOMER, " & _
        "PARTFAMILY, OPERATIONDESCRIPTION, OPERATIONNUMBER, PLANT, RELEASED, OBSOLETE FROM " & _
        "[TOOLLIST MASTER],[TOOLLIST PLANT],[TOOLLIST PARTNUMBERS] " & _
        "WHERE [TOOLLIST MASTER].PROCESSID = [TOOLLIST PARTNUMBERS].PROCESSID " & _
        "AND [TOOLLIST MASTER].PROCESSID = [TOOLLIST PLANT].PROCESSID AND " & _
        "PLANT = " + Trim(PlantListCombo.Text) + " AND [TOOLLIST MASTER].REVOFPROCESSID = 0 ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    Else
        openSQLStatement = "SELECT [TOOLLIST MASTER].PROCESSID, CUSTOMER, " & _
        "PARTFAMILY, OPERATIONDESCRIPTION, OPERATIONNUMBER, PLANT, RELEASED, OBSOLETE FROM " & _
        "[TOOLLIST MASTER],[TOOLLIST PLANT],[TOOLLIST PARTNUMBERS] " & _
        "WHERE [TOOLLIST MASTER].PROCESSID = [TOOLLIST PARTNUMBERS].PROCESSID " & _
        "AND [TOOLLIST MASTER].PROCESSID = [TOOLLIST PLANT].PROCESSID AND " & _
        "PLANT = " + Trim(PlantListCombo.Text) + " AND PARTNUMBERS = '" + Trim(PartListCombo.Text) + "' AND [TOOLLIST MASTER].REVOFPROCESSID = 0 ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION"
    End If
    OpenProcesses
End Sub

Private Sub Command2_Click()
    OpenProcess.ListView1.ListItems.Clear
    OpenProcess.Hide
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Form_Resize()
    ListView1.Top = 0
    ListView1.Left = 0
    If ScaleHeight > 750 Then
        ListView1.Height = ScaleHeight - 750
    End If
    ListView1.Width = ScaleWidth
    Command2.Top = ScaleHeight - 650
    Command1.Top = ScaleHeight - 650
    Label1.Top = ScaleHeight - 650
    Plant_Label.Top = ScaleHeight - 650
    PlantListCombo.Top = ScaleHeight - 350
    PartListCombo.Top = ScaleHeight - 350

End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
oddEvenSort = oddEvenSort + 1
If oddEvenSort Mod 2 > 0 Then
    ListView1.SortOrder = lvwAscending
Else
    ListView1.SortOrder = lvwDescending
End If

ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Sorted = True

End Sub

Private Sub ListView1_DblClick()
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    Reset
    If ListView1.SelectedItem Is Nothing Then
        Exit Sub
    End If
    If RevInProcess(Val(ListView1.SelectedItem.Text)) Then
        MsgBox ("There is an uncompleted change on this tool list, It may not be opened until that is complete.")
        Exit Sub
    End If
    DeleteExtraProcess (Val(ListView1.SelectedItem.Text))
    If ListView1.SelectedItem.SubItems(5) = "True" Then
        ProcessID = CopyProcessForChanges(Val(ListView1.SelectedItem.Text))
        OldProcessID = Val(ListView1.SelectedItem.Text)
        WorkingLive = False
        ToolList.ReleaseBtn.Visible = False
        ToolList.ChgRoutingBtn.Visible = True
    Else
        ProcessID = Val(ListView1.SelectedItem.Text)
        OldProcessID = 0
        WorkingLive = True
        ToolList.ReleaseBtn.Visible = True
        ToolList.ChgRoutingBtn.Visible = False
    End If
    OpenProcess.ListView1.ListItems.Clear
    OpenProcess.Hide
    OpenProcess.Refresh
    DoEvents
    BuildRevList
    BuildMiscList
    BuildToolList
    BuildFixtureList
    MDIForm1.TabDock.FormShow "Tool List"
    ReportForm.Show
    RefreshReport
    MDIForm1.CloseToolList.Enabled = True
    MDIForm1.CopyTool.Enabled = True
    MDIForm1.ViewToolTree.Enabled = True
    SetMultiTurret
    PopulateOriginalTools
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
End Sub
Public Sub SortByCustomer()
    If oddEvenSort Mod 2 = 0 Or oddEvenSort = 0 Then
        ListView1_ColumnClick ListView1.ColumnHeaders(2)
    End If
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyReturn Then
        If Not IsNull(ListView1.SelectedItem) Then
            ListView1_DblClick
        End If
    End If
    i = 1
     While i <= ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).SubItems(1) <> "" Then
            If Asc(Left(ListView1.ListItems.Item(i).SubItems(1), 1)) = KeyCode Then
                ListView1.ListItems.Item(i).Selected = True
                ListView1.ListItems.Item(i).EnsureVisible
                Exit Sub
            End If
        End If
        i = i + 1
    Wend
End Sub
