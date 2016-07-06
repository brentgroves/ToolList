VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Paste the following code in Class Module declaration
'API's for auto drop down combo

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F

'Paste the following code in Class Module


Public Sub ShowDropDownCombo(Cmb As ComboBox) 'Public procedure to call API's
SendMessage Cmb.hwnd, CB_SHOWDROPDOWN, 1, 0
End Sub
