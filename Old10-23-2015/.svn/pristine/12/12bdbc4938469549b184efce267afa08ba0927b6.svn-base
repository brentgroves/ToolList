VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ChangeList 
   Caption         =   "Pending Approvals & Changes"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CHANGEID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Date Initiated"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Part Family"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Operation Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Reason"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Engineer"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Approved"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Complete"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ApprovedInt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "CompleteInt"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6480
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
            Picture         =   "ChangeList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChangeList.frx":9C29
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ChangeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    ChangeList.Hide
    MDIForm1.RefreshMenuOptions
End Sub

Private Sub Form_Load()
    ListView1.Top = 0
    ListView1.Left = 0
    If ScaleHeight > 751 Then
        ListView1.Height = ScaleHeight - 750
    End If
    ListView1.Width = ScaleWidth
    Command1.Top = ScaleHeight - 650
End Sub


Private Sub Form_Resize()
    ListView1.Top = 0
    ListView1.Left = 0
    If ScaleHeight > 751 Then
        ListView1.Height = ScaleHeight - 1000
    End If
    ListView1.Width = ScaleWidth
    Command1.Top = ScaleHeight - 800
    Command1.Left = ScaleWidth - 4000
End Sub

Private Sub ListView1_DblClick()
    Select Case GetUserType
        Case "ADMIN"
            If ListView1.SelectedItem.SubItems(9) = "False" And ListView1.SelectedItem.SubItems(10) = "False" Then
                CreateRouting.SetForApproval (Val(ListView1.SelectedItem.Text))
            ElseIf ListView1.SelectedItem.SubItems(9) = "True" And ListView1.SelectedItem.SubItems(10) = "False" Then
                CreateRouting.SetForCompletion (Val(ListView1.SelectedItem.Text))
            Else
                CreateRouting.SetForViewing (Val(ListView1.SelectedItem.Text))
            End If
        Case "ENGINEER"
                CreateRouting.SetForViewing (Val(ListView1.SelectedItem.Text))
        Case "BUYER"
            If ListView1.SelectedItem.SubItems(9) = "True" And ListView1.SelectedItem.SubItems(10) = "False" Then
                CreateRouting.SetForCompletion (Val(ListView1.SelectedItem.Text))
            Else
                CreateRouting.SetForViewing (Val(ListView1.SelectedItem.Text))
            End If
        Case "MANAGER"
            If ListView1.SelectedItem.SubItems(9) = "False" And ListView1.SelectedItem.SubItems(10) = "False" Then
                CreateRouting.SetForApproval (Val(ListView1.SelectedItem.Text))
            ElseIf ListView1.SelectedItem.SubItems(9) = "True" And ListView1.SelectedItem.SubItems(10) = "False" Then
                CreateRouting.SetForCompletion (Val(ListView1.SelectedItem.Text))
            Else
                CreateRouting.SetForViewing (Val(ListView1.SelectedItem.Text))
            End If
        Case Else
            MsgBox ("Invalid User")
    End Select
    If bRefreshActionListError <> True Then
        ChangeList.Hide
    Else
        bRefreshActionListError = False
    End If
    
End Sub
