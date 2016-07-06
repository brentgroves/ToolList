VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CarbonCopy 
   Caption         =   "Copy Tool To Active Process"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   10350
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Copy 
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   6960
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4895
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Part Family"
         Object.Width           =   6880
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Operation Description"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Op Number"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5318
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tool Number"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Operation Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ToolID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Tool"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Process"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "CarbonCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    CarbonCopy.Hide
End Sub

Private Sub Copy_Click()
    Dim i As Integer
    i = 1
    While i <= ListView2.ListItems.Count
        If ListView2.ListItems.Item(i).Checked = True Then
            CopyTool Val(ListView1.SelectedItem.Text), Val(ListView2.ListItems.Item(i).SubItems(2))
        End If
        i = i + 1
    Wend
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    CarbonCopy.Hide
    BuildToolList
    RefreshReport
End Sub



Private Sub Form_Load()
Dim i As Integer
i = 1
End Sub

Private Sub Form_Resize()
'    CRViewer1.Top = 0
'    CRViewer1.Left = 0

'    CRViewer1.Height = ScaleHeight
'    CRViewer1.Width = ScaleWidth

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
oddEvenSort = oddEvenSort + 1
If oddEvenSort Mod 2 > 0 Then
    ListView1.SortOrder = lvwAscending
Else
    ListView1.SortOrder = lvwDescending
End If

ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Sorted = True

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ListView2.ListItems.Clear
    LoadTools (Val(Item.Text))
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
oddEvenSort = oddEvenSort + 1
If oddEvenSort Mod 2 > 0 Then
    ListView1.SortOrder = lvwAscending
Else
    ListView1.SortOrder = lvwDescending
End If

ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Sorted = True

End Sub

Private Sub LoadTools(pID As Long)
    Dim itmx2 As ListItem
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID = " + Str(pID) + " ORDER BY TOOLNUMBER", sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        Set itmx2 = CarbonCopy.ListView2.ListItems.Add(, , sqlRS.Fields("ToolNumber"))
        itmx2.SubItems(1) = Trim(sqlRS.Fields("OpDescription"))
        itmx2.SubItems(2) = Trim(sqlRS.Fields("ToolID"))
        
        sqlRS.MoveNext
    Wend
    sqlRS.Close
End Sub

Private Sub CopyTool(pID As Long, tID As Long)
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID = " + Str(pID) + " AND TOOLID = " + Str(tID), sqlConn, adOpenKeyset, adLockReadOnly
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "[TOOLLIST TOOL]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    SQLRS2.AddNew
    SQLRS2.Fields("ProcessID") = ProcessID
    SQLRS2.Fields("ToolNumber") = sqlRS.Fields("ToolNumber")
    SQLRS2.Fields("OpDescription") = sqlRS.Fields("OpDescription")
    SQLRS2.Fields("OffsetNumber") = sqlRS.Fields("OffsetNumber")
    SQLRS2.Fields("ToolLength") = sqlRS.Fields("ToolLength")
    SQLRS2.Fields("Turret") = "A"
    SQLRS2.Fields("Alternate") = 0
    SQLRS2.Fields("PartSpecific") = 0
    SQLRS2.Fields("AdjustedVolume") = 0
    SQLRS2.Fields("ToolOrder") = 0
    SQLRS2.Update
    sqlRS.Close
    SQLRS2.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] ORDER BY TOOLID DESC", sqlConn, adOpenKeyset, adLockReadOnly
        toolID = sqlRS.Fields("TOOLID")
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE PROCESSID = " + Str(pID) + " AND TOOLID = " + Str(tID), sqlConn, adOpenKeyset, adLockReadOnly
    SQLRS2.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = ProcessID
        SQLRS2.Fields("ToolID") = toolID
        SQLRS2.Fields("ToolType") = sqlRS.Fields("ToolType")
        SQLRS2.Fields("ToolDescription") = sqlRS.Fields("ToolDescription")
        SQLRS2.Fields("Manufacturer") = sqlRS.Fields("Manufacturer")
        SQLRS2.Fields("Consumable") = sqlRS.Fields("Consumable")
        SQLRS2.Fields("QuantityPerCuttingEdge") = sqlRS.Fields("QuantityPerCuttingEdge")
        SQLRS2.Fields("AdditionalNotes") = sqlRS.Fields("AdditionalNotes")
        SQLRS2.Fields("NumberOfCuttingEdges") = sqlRS.Fields("NumberOfCuttingEdges")
        SQLRS2.Fields("Quantity") = sqlRS.Fields("Quantity")
        SQLRS2.Fields("CribToolID") = sqlRS.Fields("CribToolID")
        SQLRS2.Fields("NumOfRegrinds") = sqlRS.Fields("NumOfRegrinds")
        SQLRS2.Fields("QtyPerRegrind") = sqlRS.Fields("QtyPerRegrind")
        SQLRS2.Fields("Regrindable") = sqlRS.Fields("Regrindable")
        SQLRS2.Update
        ToolChanges(0, ToolChangeCntr) = "ADDTOOL"
        ToolChanges(1, ToolChangeCntr) = sqlRS.Fields("CribToolID")
        ToolChangeCntr = ToolChangeCntr + 1
        sqlRS.MoveNext
    Wend
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = 1
    While i <= ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).SubItems(1) <> "" Then
            If Asc(Left(ListView1.ListItems.Item(i).SubItems(1), 1)) = KeyCode Then
                ListView1.ListItems.Item(i).Selected = True
                ListView1.ListItems.Item(i).EnsureVisible
                LoadTools (Val(ListView1.ListItems.Item(i).Text))
                Exit Sub
            End If
        End If
        i = i + 1
    Wend
End Sub
