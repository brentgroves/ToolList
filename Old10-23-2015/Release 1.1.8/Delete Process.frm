VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DeleteProcess 
   Caption         =   "Select Process to Delete"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   7665
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6800
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process ID"
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
   End
End
Attribute VB_Name = "DeleteProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    DeleteProcess.ListView1.ListItems.Clear
    DeleteProcess.Hide
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Form_Resize()
ListView1.Top = 0
ListView1.Left = 0
ListView1.Height = ScaleHeight - 750
ListView1.Width = ScaleWidth
Command2.Top = ScaleHeight - 650
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
    ProcessID = Val(ListView1.SelectedItem.Text)
    If MsgBox("Are you sure you want to delete Process " + Str(ProcessID) + " - " + ListView1.SelectedItem.SubItems(1) + " - " + ListView1.SelectedItem.SubItems(2) + "?", vbOKCancel) = vbOK Then
        ListView1.ListItems.Clear
        DeleteProcess.Hide
        DeleteProcessSub
        MDIForm1.CloseToolList.Enabled = False
        MDIForm1.OpenToolList.Enabled = True
        MDIForm1.NewToolList.Enabled = True
        MDIForm1.DeleteToolList.Enabled = True
    Else
        ProcessID = 0
    End If
End Sub
