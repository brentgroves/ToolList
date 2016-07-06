VERSION 5.00
Begin VB.Form RoutingComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Details"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox CommentsTXT 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "RoutingComments.frx":0000
      Top             =   480
      Width           =   3975
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
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "RoutingComments"
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
    If CreateRouting.GetViewingType = "Creation" Then
        List.ListItems.Item(ListIndex).SubItems(5) = Trim(CommentsTXT.Text)
        CommentsTXT.Text = ""
        RoutingComments.Hide
    Else
        List.ListItems.Item(ListIndex).SubItems(5) = Trim(CommentsTXT.Text)
        RoutingComments.Hide
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
    RoutingComments.Hide
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
    RoutingComments.Caption = Title
    RoutingComments.Show
    RoutingComments.CommentsTXT = getlist.ListItems.Item(getindex).SubItems(5)
End Sub

