VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form NotificationForm 
   Caption         =   "Notification Edit"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   8745
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"NotificationEdit.frx":0000
   End
   Begin VB.CommandButton DontSendCMD 
      Caption         =   "Don't Send"
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton SendCMD 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "NOTIFICATION TEXT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "NotificationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DontSendCMD_Click()
    NotificationMessage = ""
    NotificationSubject = ""
    OldItemNumber = ""
    NotificationForm.Hide
End Sub


Private Sub SendCMD_Click()
    GetSendTo
    NotificationMessage = Text1.Text
    SendEmail
    NotificationMessage = ""
    NotificationSubject = ""
    OldItemNumber = ""
    NotificationForm.Hide
End Sub
