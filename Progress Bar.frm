VERSION 5.00
Begin VB.Form ProgressBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PROCESSING...DO NOT CLICK...DO NOT CLICK...DO NOT CLICK...DO NOT CLICK...DO NOT CLICK..."
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6000
      Top             =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "If this action takes more than 60 seconds inform the IT Dept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9495
   End
   Begin VB.Label Label1 
      Caption         =   "PLEASE DO NOT CLICK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   16
      Left            =   4080
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   15
      Left            =   3840
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   14
      Left            =   3600
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   13
      Left            =   3360
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   12
      Left            =   3120
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   11
      Left            =   2880
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   10
      Left            =   2640
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   9
      Left            =   2400
      Top             =   90
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   8
      Left            =   2160
      Top             =   90
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   7
      Left            =   1920
      Top             =   90
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   6
      Left            =   1680
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   5
      Left            =   1440
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   4
      Left            =   1200
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   3
      Left            =   960
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   2
      Left            =   720
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   1
      Left            =   480
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurPosition As Integer
Option Explicit
Private Const GWL_HWNDPARENT = (-8)
Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal wNewLong As Long) As Long
Private hParentWnd As Long

Public Sub AdvanceShapes()
    Dim i
    For i = 0 To Shape1.Count - 1
        If i = CurPosition Or i = CurPosition + 1 Or i = CurPosition + 2 Or i = CurPosition - Shape1.Count + 1 Or i = CurPosition - Shape1.Count + 2 Then
            Shape1(i).FillColor = &HFF0000
            Shape1(i).Height = 160
            Shape1(i).Top = 90
        Else
            Shape1(i).FillColor = &HFF00&
            Shape1(i).Height = 105
            Shape1(i).Top = 120
        End If
    Next
    If CurPosition = 16 Then
        CurPosition = 0
    Else
        CurPosition = CurPosition + 1
    End If
End Sub

Private Sub Form_Load()
    CurPosition = 0
    hParentWnd = SetWindowLong(Me.hwnd, GWL_HWNDPARENT, MDIForm1.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetWindowLong(Me.hwnd, GWL_HWNDPARENT, hParentWnd)
End Sub

Private Sub Timer1_Timer()
    AdvanceShapes
End Sub
