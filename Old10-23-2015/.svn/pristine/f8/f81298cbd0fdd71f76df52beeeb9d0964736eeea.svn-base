VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ItemComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Details"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ChangeIDTXT 
      Height          =   375
      Left            =   4560
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox OldMonthlyUsageTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox TotalTXT 
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox ReworkCostTXT 
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox ReworkQtyTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox CombinedUsageTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox ActionTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox NewQtyTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox BinTxt 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox NewCostTXT 
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox MonthlyUsageTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox ItemGroupTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox ManufacturerTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox ItemNumberTXT 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   735
      Left            =   8880
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   7320
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox CommentsTXT 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "ItemComments.frx":0000
      Top             =   480
      Width           =   3975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3413
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483637
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gill Sans MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Job"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Usage"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView ActionItemList 
      Height          =   1695
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Actions Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Action Text"
         Object.Width           =   14817
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Monthly Usage:"
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
      Left            =   4200
      TabIndex        =   31
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Value On Hand:"
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
      Left            =   4680
      TabIndex        =   29
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Rework Cost:"
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
      Left            =   4680
      TabIndex        =   27
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Rework Qty:"
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
      Left            =   4200
      TabIndex        =   25
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Combined Usage:"
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
      Left            =   840
      TabIndex        =   22
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Other Jobs Used On:"
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
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Action:"
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
      Left            =   4800
      TabIndex        =   19
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Comments"
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
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "New Qty:"
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
      Left            =   4200
      TabIndex        =   16
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
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
      Left            =   5760
      TabIndex        =   15
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Monthly Usage:"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "New Cost:"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Line3Lbl 
      Alignment       =   1  'Right Justify
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
      Left            =   4200
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Line1Lbl 
      Alignment       =   1  'Right Justify
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
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Line2Lbl 
      Alignment       =   1  'Right Justify
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
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "ItemComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GWL_HWNDPARENT = (-8)
Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal wNewLong As Long) As Long
Private hParentWnd As Long
Private ListIndex As Integer
Private List As ListView

Private Sub Command1_Click()
    'MsgBox (CreateRouting.GetViewingType)
    If CreateRouting.GetViewingType = "Completion" Then
        UpdateActionItemsForTools
    End If
    If CreateRouting.GetViewingType = "Creation" Then
        List.ListItems.Item(ListIndex).SubItems(5) = Trim(CommentsTXT.Text)
        CommentsTXT.Text = ""
        Me.Hide
    Else
        List.ListItems.Item(ListIndex).SubItems(5) = Trim(CommentsTXT.Text)
        Me.Hide
        Set sqlRS = New ADODB.Recordset
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE ITEMCHANGEID =" + List.ListItems.Item(ListIndex).SubItems(6), sqlConn, adOpenKeyset, adLockOptimistic
        If Not sqlRS.EOF Then
            sqlRS.Fields("COMMENTS") = Trim(CommentsTXT.Text)
            CommentsTXT.Text = ""
            sqlRS.Update
        End If
        sqlRS.Close
        Set sqlRS = Nothing
    End If
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    hParentWnd = SetWindowLong(Me.hwnd, GWL_HWNDPARENT, MDIForm1.hwnd)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SetWindowLong(Me.hwnd, GWL_HWNDPARENT, hParentWnd)
End Sub
Public Sub setListInfo(getlist As ListView, getindex As Integer, Title As String)
    Set List = getlist
    ListIndex = getindex
    Me.Caption = Title
    Me.Show
    Me.CommentsTXT = getlist.ListItems.Item(getindex).SubItems(5)
End Sub
