Attribute VB_Name = "PublicVar"

Public processId As Long
Public OldProcessID As Long
Public toolID As Long
Public itemID As Long
Public FixtureToolID As Long
Public MiscToolID As Long
Public RevisionID As Long

Public sqlConn As ADODB.Connection
Public M2MsqlConn As ADODB.Connection
Public sqlRS As ADODB.Recordset
Public SQLRS2 As ADODB.Recordset
Public SQLRS3 As ADODB.Recordset
Public SQLRS4 As ADODB.Recordset
Public sqlCMD As ADODB.Command
Public CribRS As ADODB.Recordset
Public CribConn As ADODB.Connection

Public colItemImages As New Collection
Dim inv As New Collection

Public craxReport As New CRAXDRT.Report
Public reportOpened As Boolean
Public reportViewed As Boolean


Public craxApp As New CRAXDRT.Application

Public bRefreshActionListError As Boolean

Public toolexists As Boolean
Public itemexists As Boolean
Public misctoolexists As Boolean
Public fixturetoolexists As Boolean
Public revisionexists As Boolean
Public processexists As Boolean
Public WorkingLive As Boolean

Public OldCribID As String
Public LastToolModified As String
Public PlantChange(10) As Integer

Public ToolChanges(6, 200) As String
Public ToolChangeCntr As Long
Public MiscToolChangeCntr As Long
Public OriginalTools(400) As String
Public OriginalPlant(10) As Integer
Public OriginalPics(400) As Boolean

Public OriginalVolume As Long
Public OriginalReleased As Boolean
Public OriginalObsolete As Boolean
Public EmailMessage As String

Public oddEvenSort As Integer
Public ExitLoop As Boolean
Public LastToolDescription As String

Public MultiTurret As Boolean
Public openSQLStatement As String


Public Const TOOLLIST_RPT_IND = "\\buschesv2\public\Report Files\toollist.rpt"
Public Const TOOLLIST_RPT_AL = "\\hartselle-public\Shared\Public\Report Files\toollist.rpt"
Public Const TOOLLIST_RPT_TEST = "\\hartselle-public\Shared\Public\Report Files\toollistTest.rpt"
Public Const TOOLLIST_CHANGE_TASKS_IND = "\\buschesv2\public\Report Files\Tool List Change Tasks.rpt"
Public Const TOOLLIST_CHANGE_TASKS_AL = "\\hartselle-public\Shared\Public\Report Files\Tool List Change Tasks.rpt"
Public Const TOOLLIST_CHANGE_TASKS_TEST = "\\hartselle-public\Shared\Public\Report Files\Tool List Change TasksTest.rpt"
Public Const WM_USER = &H400
Public Const TV_FIRST = &H1100
Public Const TTM_ACTIVATE = (WM_USER + 1)
Public Const TVM_GETTOOLTIPS = (TV_FIRST + 25)


Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
' http://www.vbforums.com/showthread.php?348138-Classic-VB-How-do-I-open-a-file-web-page-in-its-default-application
'in General-Declarations:
'https://msdn.microsoft.com/en-us/library/windows/desktop/bb762153(v=vs.85).aspx
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Private Const WAIT_INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

'Be careful have form of same name
'Private Declare Function OpenProcess Lib "kernel32" _
'  (ByVal dwDesiredAccess As Long, _
'   ByVal bInheritHandle As Long, _
'   ByVal dwProcessId As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" _
  (ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long
   
Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long
  
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103&

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private m_TargetString As String
Private m_TargetHwnd As Long

Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


' Fac is a build parameter indicating whether this program will be used
' in Indiana or Alabama
Public BUILDTYPE As Integer
Public Const IND = 1
Public Const AL = 2
Public Const TEST = 3
Public CRVIEWER As String
'Example code (same effect as the examples above)
'  Call OpenURL("http://www.VBForums.com")
'  Call OpenAFile("c:\My folder\Test.doc")
Public Sub Init()
    ' Fac is a build parameter indicating whether this program will be used
    ' in Indiana or Alabama
'    If (BUILDTYPE = TEST) Then
'    ElseIf (BUILDTYPE = IND) Then
'    ElseIf (BUILDTYPE = AL) Then
'    End If
' TOOLLIST_RPT_AL
    BUILDTYPE = IND
    If (BUILDTYPE = TEST) Then
        CRVIEWER = "\\buschesv2\public\MIS\Busche Software Installation Files\CRSilentViewerManifest\CRSilentViewer.appref-ms"
    ElseIf (BUILDTYPE = IND) Then
        CRVIEWER = "\\buschesv2\public\MIS\Busche Software Installation Files\CRSilentViewerManifest\CRSilentViewer.appref-ms"
    ElseIf (BUILDTYPE = AL) Then
        CRVIEWER = "\\hartselle-public\shared\Public\MIS\Busche Software Installation Files\CRSilentViewerManifest\CRSilentViewer.appref-ms"
    End If
    
    
    Set sqlConn = New ADODB.Connection
    If (BUILDTYPE = TEST) Then
        sqlConn.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=Busche Toolist09042013;" & _
               "User Id=sa;" & _
               "Password=buschecnc1;"
    ElseIf (BUILDTYPE = IND) Then
        sqlConn.Open "Provider=sqloledb;" & _
               "Data Source=busche-sql;" & _
               "Initial Catalog=Busche ToolList;" & _
               "User Id=sa;" & _
               "Password=buschecnc1;"
    ElseIf (BUILDTYPE = AL) Then
        sqlConn.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=Busche ToolList;" & _
               "User Id=sa;" & _
               "Password=buschecnc1;"
    End If
    
   
    Set CribConn = New ADODB.Connection
    If (BUILDTYPE = TEST) Then
        CribConn.Open "Provider=SQLOLEDB.1;" & _
                "Data Source=BUSCHE-SQL;" & _
                "Initial Catalog=Cribmaster;" & _
                "User Id=sa;" & _
                "Password=buschecnc1;"
    ElseIf (BUILDTYPE = IND) Then
        CribConn.Open "Provider=SQLOLEDB.1;" & _
                "Data Source=busche-sql;" & _
                "Initial Catalog=Cribmaster;" & _
                "User Id=sa;" & _
                "Password=buschecnc1;"
    ElseIf (BUILDTYPE = AL) Then
        CribConn.Open "Provider=sqloledb;" & _
                "Data Source=hartselle-sql-1;" & _
                "Initial Catalog=Cribmaster;" & _
                "User Id=sa;" & _
                "Password=buschecnc1;"
    End If
           
    Set M2MsqlConn = New ADODB.Connection
    If (BUILDTYPE = TEST) Then
        M2MsqlConn.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=m2mdata01;" & _
               "User Id=sa;" & _
               "Password=buschecnc1;"
    ElseIf (BUILDTYPE = IND) Then
        M2MsqlConn.Open "Provider=sqloledb;" & _
               "Data Source=busche-sql-1;" & _
               "Initial Catalog=m2mdata01;" & _
               "User Id=sa;" & _
               "Password=buschecnc1;"
    ElseIf (BUILDTYPE = AL) Then
        M2MsqlConn.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=m2mdata01;" & _
               "User Id=sa;" & _
               "Password=buschecnc1;"
    End If
   
    openSQLStatement = "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER"
    bRefreshActionListError = False
    reportOpened = False
'    InitializeReport
    ToolChangeCntr = 0
End Sub
' Check a returned task to see if it's the one we want.
Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
Dim buf As String
Dim Title As String
Dim length As Long

    ' Get the window's title.
    length = GetWindowTextLength(app_hWnd)
    buf = Space$(length)
    length = GetWindowText(app_hWnd, buf, length)
    Title = Left$(buf, length)

    ' See if the title contains the target string.
    If InStr(Title, m_TargetString) <> 0 Then
        ' This is the one we want.
        m_TargetHwnd = app_hWnd

        ' Stop searching.
        EnumCallback = 0
    Else
        ' Continue searching.
        EnumCallback = 1
    End If
End Function
' Find a window with a title containing target_string
' and return its hWnd.
Public Function FindWindowByPartialTitle(ByVal target_string As String) As Long
    m_TargetString = target_string
    m_TargetHwnd = 0

    ' Enumerate windows.
    EnumWindows AddressOf EnumCallback, 0

    ' Return the hWnd found (if any).
    FindWindowByPartialTitle = m_TargetHwnd
End Function 'Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Sub Pause(SecsDelay As Single)
   Dim TimeOut   As Single
   Dim PrevTimer As Single
   
   PrevTimer = Timer
   TimeOut = PrevTimer + SecsDelay
   Do While PrevTimer < TimeOut
  '    Sleep 4 '-- Timer is only updated every 1/64 sec = 15.625 millisecs.
      DoEvents
      If Timer < PrevTimer Then TimeOut = TimeOut - 86400 '-- pass midnight
      PrevTimer = Timer
   Loop
End Sub
 
Public Sub OpenCRViewer(processId As String)
   Dim hProcess As Long
   Dim taskId As Long
    Dim exitCode As Long
    Dim target_hwnd As Long
   taskId = ShellExecute(Screen.ActiveForm.hwnd, "open", CRVIEWER, processId, "C:\", ByVal 0)
' TRY TO MAKE THE CRSilentViewer Mode..No Success
'   Pause (5)
'    Do
 '       target_hwnd = FindWindowByPartialTitle(processId)
  '      If target_hwnd <> 0 Then
   '     Else
    '        Pause (1)
     '   End If
  '  Loop While target_hwnd = 0
 '  ViewProcess.Show
'   taskId = ShellExecute(Screen.ActiveForm.hwnd, "open", strURL, vbNullString, "C:\", ByVal 0)

End Sub
 
 
Public Sub RunDotNetReport()
 '   Call Shell(App.Path & "\" & App.EXEName & ".exe", 1)
 Call Shell(App.Path & "\" & crvtest, 1)
End Sub
'THIS FUNCTION IS NO LONGER USED
'Public Sub InitializeReport()
 '   Set craxReport = craxApp.OpenReport("\\buschesv2\public\Report Files\toollist.rpt")
  '  Set craxReport = craxApp.OpenReport("\\hartselle-public\Shared\Public\Report Files\toollist.rpt")
   ' craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    'craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (ProcessID)
'End Sub
Public Sub RunReport()
    Dim delay As Date
    ReportForm.Show
    delay = Time
    While DateAdd("s", 0.75, delay) > Time
        DoEvents
    Wend
    If reportOpened = False Then
        If (BUILDTYPE = TEST) Then
            Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_TEST)
        ElseIf (BUILDTYPE = IND) Then
            Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_IND)
        ElseIf (BUILDTYPE = AL) Then
            Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_AL)
        End If
        reportOpened = True
    End If
    
    craxReport.DiscardSavedData
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (processId)
    ReportForm.CRViewer1.ReportSource = craxReport
    ReportForm.CRViewer1.Refresh
    ReportForm.CRViewer1.ViewReport
    ReportForm.CRViewer1.Zoom 80
    delay = Time
    ExitLoop = False
    While DateAdd("s", 3, delay) > Time
        ToolList.SetFocus
        DoEvents
        If Screen.ActiveForm.Caption = "Tool List" Or ExitLoop Then
            ExitLoop = False
            Exit Sub
        End If
    Wend
End Sub

Public Sub RefreshReport()
    ToolList.ReportBtn.Enabled = True
End Sub
Public Sub OpenProcesses()
    Dim itmx2 As ListItem
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open openSQLStatement, sqlConn
    OpenProcess.ListView1.ListItems.Clear
    While Not sqlRS.EOF
        Set itmx2 = OpenProcess.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSID"))
        If Not IsNull(sqlRS.Fields("Customer")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("Customer"))
        Else
            itmx2.SubItems(1) = ""
        End If
        If Not IsNull(sqlRS.Fields("PartFamily")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("PartFamily"))
        Else
            itmx2.SubItems(2) = ""
        End If
        If Not IsNull(sqlRS.Fields("OperationDescription")) Then
            itmx2.SubItems(3) = Trim(sqlRS.Fields("OperationDescription"))
        Else
            itmx2.SubItems(3) = ""
        End If
        If Not IsNull(sqlRS.Fields("OperationNumber")) Then
            itmx2.SubItems(4) = Trim(sqlRS.Fields("OperationNumber"))
        Else
            itmx2.SubItems(4) = ""
        End If
        If Not IsNull(sqlRS.Fields("RELEASED")) Then
            itmx2.SubItems(5) = sqlRS.Fields("RELEASED")
        Else
            itmx2.SubItems(5) = ""
        End If
        If Not IsNull(sqlRS.Fields("Obsolete")) Then
            itmx2.SubItems(6) = Trim(sqlRS.Fields("Obsolete"))
        Else
            itmx2.SubItems(6) = ""
        End If
        sqlRS.MoveNext
        itmx2.ForeColor = vbRed
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    OldCribID = ""
    ColorRows OpenProcess.ListView1
    OpenProcess.SortByCustomer
End Sub
Public Sub ColorRows(lv As ListView)
Dim intindex As Integer
Dim rowindex As Integer
Dim itmx As ListItem
Dim lvSI As ListSubItem
Dim rowcolor As OLE_COLOR

For rowindex = 1 To lv.ListItems.Count
    Set itmx = lv.ListItems(rowindex)
    If itmx.ListSubItems(5) = "True" Then
        rowcolor = &H8000&
    Else
        rowcolor = vbRed
    End If
    If itmx.ListSubItems(6) = "True" Then
        rowcolor = &HC0C0C0
    End If
    For intindex = 1 To lv.ColumnHeaders.Count - 1
        Set lvSI = itmx.ListSubItems(intindex)
        lvSI.ForeColor = rowcolor
    Next
    itmx.ForeColor = rowcolor
    itmx.Selected = True
Next
If lv.ListItems.Count > 0 Then
    lv.ListItems(1).Selected = True
End If
Set itmx = Nothing
Set lvSI = Nothing
End Sub

Public Sub AddProcess()
    ClearProcessFields
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[TOOLLIST MASTER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("PartFamily") = ""
    sqlRS.Fields("OperationNumber") = 0
    sqlRS.Fields("OperationDescription") = ""
    sqlRS.Fields("ReleaseD") = 0
    sqlRS.Fields("Obsolete") = 0
    sqlRS.Fields("Customer") = ""
    sqlRS.Fields("AnnualVolume") = 0
    sqlRS.Fields("MultiTurret") = 0
    sqlRS.Fields("RevOfProcessID") = 0
    sqlRS.Fields("RevInProcess") = 0
    sqlRS.Update
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY PROCESSID DESC", sqlConn, adOpenKeyset, adLockReadOnly
    processId = sqlRS.Fields("ProcessID")
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("OriginalProcessID") = sqlRS.Fields("ProcessID")
    sqlRS.Update
    Set sqlRS = Nothing
    WorkingLive = True
End Sub

Public Sub GetAllPartNumbers()
    Dim x As ListItem
    Dim y As ColumnHeader
    Set y = ProcessAttr.AllPartNumbersList2.ColumnHeaders.Add(, , "Part #", ProcessAttr.AllPartNumbersList2.Width / 2)
    Set y = ProcessAttr.AllPartNumbersList2.ColumnHeaders.Add(, , "Rev #", ProcessAttr.AllPartNumbersList2.Width / 5)
    Set y = ProcessAttr.AllPartNumbersList2.ColumnHeaders.Add(, , "Status", ProcessAttr.AllPartNumbersList2.Width / 5)
    ProcessAttr.AllPartNumbersList2.BorderStyle = ccFixedSingle ' Set BorderStyle property.
    ProcessAttr.AllPartNumbersList2.View = lvwReport ' Set View property to Report.
    Set sqlRS = New ADODB.Recordset
   ' sqlRS.Open "SELECT * FROM INVCUR,INMAST WHERE INVCUR.FCPARTNO = INMAST.FPARTNO AND INVCUR.FLANYCUR = 1 ORDER BY INVCUR.FCPARTNO", M2MsqlConn
    sqlRS.Open "SELECT fpartno,frev,fcstscode FROM INVCUR,INMAST Where INVCUR.FCPARTNO = INMAST.FPARTNO And INVCUR.FLANYCUR = 1 ORDER BY INVCUR.FCPARTNO", M2MsqlConn
    'Populate Part Numbers and Revisions under columns
    While Not sqlRS.EOF
        Set x = ProcessAttr.AllPartNumbersList2.ListItems.Add(, , sqlRS.Fields("FPARTNO"))
        x.SubItems(1) = (sqlRS.Fields("FREV"))
        x.SubItems(2) = (sqlRS.Fields("FCSTSCODE"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set aqlRS = Nothing
End Sub

'Old PartNumber Population code with Rev #'s added

'Public Sub GetAllPartNumbers()
'    Set sqlRS = New ADODB.Recordset
'    sqlRS.Open "SELECT * FROM INVCUR,INMAST WHERE INVCUR.FCPARTNO = INMAST.FPARTNO AND INVCUR.FLANYCUR = 1 ORDER BY INVCUR.FCPARTNO", M2MsqlConn
'    While Not sqlRS.EOF
'        ProcessAttr.AllPartNumbersList.AddItem (sqlRS.Fields("FPARTNO")) & (sqlRS.Fields("FREV"))
'        sqlRS.MoveNext
'    Wend
'    sqlRS.Close
'    Set sqlRS = Nothing
'End Sub
Public Sub GetAllPlants()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT LIST] ORDER BY PLANT", sqlConn
    While Not sqlRS.EOF
        ProcessAttr.AllPlantList.AddItem (sqlRS.Fields("PLANT"))

        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAllPlantsForFilter()
    OpenProcess.PlantListCombo.Clear
    OpenProcess.PlantListCombo.AddItem ("All")
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT LIST] ORDER BY PLANT", sqlConn
    While Not sqlRS.EOF
        OpenProcess.PlantListCombo.AddItem (sqlRS.Fields("PLANT"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAllPlantsForFilterView()
    ViewProcess.PlantListCombo.Clear
    ViewProcess.PlantListCombo.AddItem ("All")
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT LIST] ORDER BY PLANT", sqlConn
    While Not sqlRS.EOF
        ViewProcess.PlantListCombo.AddItem (sqlRS.Fields("PLANT"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAllPartsForFilter()
    OpenProcess.PartListCombo.Clear
    OpenProcess.PartListCombo.AddItem ("All")
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT DISTINCT PARTNUMBERS FROM [TOOLLIST PARTNUMBERS]", sqlConn
    While Not sqlRS.EOF
        OpenProcess.PartListCombo.AddItem (sqlRS.Fields("PARTNUMBERS"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAllPartsForFilterView()
    ViewProcess.PartListCombo.Clear
    ViewProcess.PartListCombo.AddItem ("All")
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT DISTINCT PARTNUMBERS FROM [TOOLLIST PARTNUMBERS]", sqlConn
    While Not sqlRS.EOF
        ViewProcess.PartListCombo.AddItem (sqlRS.Fields("PARTNUMBERS"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAssignedPartNumbers()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(processId), sqlConn
    While Not sqlRS.EOF
        ProcessAttr.SelectedPartsList.AddItem (sqlRS.Fields("PartNumbers"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub GetToolPartNumbers()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOLPARTNUMBER] WHERE TOOLID =" + Str(toolID), sqlConn
    While Not sqlRS.EOF
        ToolAttr.SelectedPartsList.AddItem (sqlRS.Fields("PartNumber"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(processId), sqlConn
    While Not sqlRS.EOF
        ToolAttr.AllPartNumbersList.AddItem (sqlRS.Fields("PartNumbers"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub GetAvailableToolPartNumbers()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(processId), sqlConn
    While Not sqlRS.EOF
        ToolAttr.AllPartNumbersList.AddItem (sqlRS.Fields("PartNumbers"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAssignedPlant()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT] WHERE PROCESSID = " + Str(processId) + " ORDER BY PLANT", sqlConn
    While Not sqlRS.EOF
        ProcessAttr.SelectedPlantsList.AddItem (sqlRS.Fields("Plant"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetProcessDetails()
    Dim i As Integer
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn
    If Not IsNull(sqlRS.Fields("PartFamily")) Then
        ProcessAttr.PartFamilyTXT.Text = sqlRS.Fields("PartFamily")
    End If
    If Not IsNull(sqlRS.Fields("OperationNumber")) Then
        ProcessAttr.OpNumTXT.Text = sqlRS.Fields("OperationNumber")
    End If
    If Not IsNull(sqlRS.Fields("OperationDescription")) Then
        ProcessAttr.OpDescTXT.Text = sqlRS.Fields("OperationDescription")
    End If
    If Not IsNull(sqlRS.Fields("Customer")) Then
        ProcessAttr.CustomerTXT.Text = sqlRS.Fields("Customer")
    End If
    If Not IsNull(sqlRS.Fields("AnnualVolume")) Then
        ProcessAttr.AnnualVolumeTXT.Text = sqlRS.Fields("AnnualVolume")
    End If
    If Not IsNull(sqlRS.Fields("OBSOLETE")) Then
        If sqlRS.Fields("Obsolete") Then
            i = 1
        Else
            i = 0
        End If
        ProcessAttr.ObsoleteCheck.Value = i
    End If
    If Not IsNull(sqlRS.Fields("MultiTurret")) Then
        If sqlRS.Fields("MultiTurret") Then
            i = 1
        Else
            i = 0
        End If
        ProcessAttr.MultiTurretLathe.Value = i
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetToolDetails()
    Dim i As Integer
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(toolID), sqlConn
    ToolAttr.ToolNumberTXT.Text = sqlRS.Fields("ToolNumber")
    ToolAttr.OpDescTXT.Text = sqlRS.Fields("OpDescription")
    If (True = reportOpened) Then
        While ReportForm.CRViewer1.IsBusy
            DoEvents
        Wend
    End If
    
    Dim TEST As Boolean
    If sqlRS.Fields("Alternate") Then
        i = 1
    Else
        i = 0
    End If
    ToolAttr.AlternateCHECK.Value = i
    If sqlRS.Fields("PartSpecific") Then
        i = 1
    Else
        i = 0
    End If
    ToolAttr.PartNumberCheck.Value = i
    If i = 1 Then
        ToolAttr.EnableMultiPart
        ToolAttr.AdjustedVolume.Text = sqlRS.Fields("AdjustedVolume")
    Else
        ToolAttr.DisableMultiPart
    End If
    ToolAttr.ToolLengthOffsetTXT.Text = sqlRS.Fields("ToolLength")
    If MultiTurret Then
        ToolAttr.EnableMultiTurret
        If sqlRS.Fields("Turret") = "B" Then
            ToolAttr.TurretAOption.Value = False
            ToolAttr.TurretBOption.Value = True
        Else
            ToolAttr.TurretAOption.Value = True
            ToolAttr.TurretBOption.Value = False
        End If
    End If
    toolID = sqlRS.Fields("TOOLID")

    ToolAttr.SequenceTxt.Text = sqlRS.Fields("TOOLORDER")
    ToolAttr.OffsetNumberTXT.Text = sqlRS.Fields("OffsetNumber")
    sqlRS.Close
    Set sqlRS = Nothing
    If i = 1 Then
        GetToolPartNumbers
    End If
End Sub

Public Sub GetItemDetails()
    Dim i As Integer
    Dim strTB As String
    Dim strStream As ADODB.Stream
    Dim bytearray() As Byte

    
    
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID), sqlConn
    ItemAttri.CribNumberIDTXT.Text = sqlRS.Fields("CribToolID")
    OldCribID = sqlRS.Fields("CribToolid")
    ItemAttri.QuantityTXT.Text = sqlRS.Fields("Quantity")
    ItemAttri.CuttingEdgesTXT.Text = sqlRS.Fields("NumberOfCuttingEdges")
    ItemAttri.ToolLifeTXT.Text = sqlRS.Fields("QuantityPerCuttingEdge")
    ItemAttri.AdditionalNotesTXT.Text = sqlRS.Fields("AdditionalNotes")
    ItemAttri.NumofRegrindsTXT.Text = sqlRS.Fields("NumOfRegrinds")
    ItemAttri.ToolLifeRegrindTXT.Text = sqlRS.Fields("QtyPerRegrind")
    ItemAttri.txtPicture = ""
    If IsNull(sqlRS.Fields("ItemImage")) = False Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
        strStream.Write sqlRS.Fields("ItemImage")
        strStream.Flush
          ' rewind stream and read text
        strStream.Position = 0
        strStream.Type = adTypeBinary
      '  strStream.SaveToFile "c:\temppic.jpg", adSaveCreateOverWrite
        bytearray = strStream.Read()
        ItemAttri.imgItem.Picture = PictureFromBits(bytearray)
        strStream.Close
        Set strStream = Nothing
    End If
    toolID = sqlRS.Fields("ToolID")
    
    If IsNull(sqlRS.Fields("ToolbossStock")) = False Then
       If sqlRS.Fields("ToolbossStock") Then
         i = 1
       Else
         i = 0
       End If
    Else
       i = 0
    End If
    ItemAttri.TBStock.Value = i
    If sqlRS.Fields("Consumable") Then
        i = 1
    Else
        i = 0
    End If
    ItemAttri.ConsumableCHECK.Value = i
    If sqlRS.Fields("Regrindable") Then
        i = 1
    Else
        i = 0
    End If
    ItemAttri.RegrindableChk.Value = i
    sqlRS.Close
    Set sqlRS = Nothing
    
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT DESCRIPTION1, Manufacturer, ItemClass, [INVENTRY].ItemNumber, Cost FROM [INVENTRY] LEFT OUTER JOIN [ALTVENDOR] ON [INVENTRY].[ALTVENDORNO] = [ALTVENDOR].[RECNUMBER] WHERE [INVENTRY].ITEMNUMBER = '" + ItemAttri.CribNumberIDTXT.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        ItemAttri.ItemNumberCOMBO.Text = CribRS.Fields("Description1")
        If Not IsNull(CribRS.Fields("Manufacturer")) Then
            ItemAttri.ManufacturerTXT.Text = CribRS.Fields("Manufacturer")
        End If
        If Not IsNull(CribRS.Fields("ItemClass")) Then
            ItemAttri.ItemGroupTXT.Text = CribRS.Fields("ItemClass")
        End If
        If Not IsNull(CribRS.Fields("COST")) Then
            ItemAttri.CostTXT.Text = CribRS.Fields("Cost")
        End If
    Else
        ItemAttri.ItemNumberCOMBO.Text = ""
        ItemAttri.ManufacturerTXT.Text = ""
        ItemAttri.ItemGroupTXT.Text = ""
        MsgBox ("This is no longer a Valid Tool")
    End If
    CribRS.Close
    Set CribRS = Nothing
    GetQty
    CalculateCosts
End Sub
Public Sub GetQty()
    Dim sum As Integer
    Dim binstring As String
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT ITEM, CRIBBIN, BINQUANTITY FROM STATION WHERE ITEM = '" + ItemAttri.CribNumberIDTXT.Text + "' OR ITEM = '" + ItemAttri.CribNumberIDTXT.Text + "R'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        While Not CribRS.EOF
            sum = sum + CribRS.Fields("binquantity")
            binstring = CribRS.Fields("CribBin") + ", " + binstring
            CribRS.MoveNext
        Wend
        ItemAttri.QtyOnHandTXT.Text = sum
        ItemAttri.BinTxt.Text = binstring
    Else
        ItemAttri.QtyOnHandTXT.Text = 0
        ItemAttri.BinTxt.Text = ""
    End If
    CribRS.Close
    Set CribRS = Nothing
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST TOOLBOSS STOCK ITEMS] WHERE ITEMCLASS = '" + ItemAttri.ItemGroupTXT.Text + "'", sqlConn, adOpenKeyset

    If SQLRS2.RecordCount > 0 Then
        ItemAttri.TBStock.Enabled = False
    Else
        ItemAttri.TBStock.Enabled = True
    End If
    SQLRS2.Close
    Set SQLRS2 = Nothing
End Sub
Public Sub UpdatePartNumbers()
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(processId)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[ToolList PartNumbers]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Dim i As Integer
    i = 0
    While i < ProcessAttr.SelectedPartsList.ListCount
        sqlRS.AddNew
        sqlRS.Fields("ProcessID") = processId
        sqlRS.Fields("PartNumbers") = Trim(ProcessAttr.SelectedPartsList.List(i))
        sqlRS.Update
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub UpdatePlants()
    Dim PlantsChanged As Boolean
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST PLANT] WHERE PROCESSID =" + Str(processId)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[ToolList PLANT]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Dim i As Integer
    i = 0
    'SET ALL ELEMENTS OF THE ARRAY TO 0
    For i = 0 To 9
         PlantChange(i) = 0
    Next i
    i = 0
    While i < ProcessAttr.SelectedPlantsList.ListCount
        sqlRS.AddNew
        sqlRS.Fields("ProcessID") = processId
        sqlRS.Fields("Plant") = Trim(ProcessAttr.SelectedPlantsList.List(i))
        sqlRS.Update
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    i = 0
    
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [ToolList PLANT] WHERE PROCESSID =" + Str(processId) + " ORDER BY PLANT", sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        PlantChange(i) = sqlRS.Fields("Plant")
        i = i + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    i = 0
    While i < 10
        If PlantChange(i) <> OriginalPlant(i) Then

            PlantsChanged = True
        End If
        i = i + 1
    Wend
    If PlantsChanged Then
        Dim TEMP As String
        For i = 0 To 9
            If PlantChange(i) <> 0 Then
                TEMP = TEMP + Str(PlantChange(i)) + ", "
            End If
        Next
        ToolChanges(0, ToolChangeCntr) = "PLANT"
        ToolChanges(1, ToolChangeCntr) = TEMP
        TEMP = ""
        ToolChangeCntr = ToolChangeCntr + 1
    End If
End Sub


Public Sub UpdateProcessDetails()
    Dim i As Integer
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("PartFamily") = UCase(ProcessAttr.PartFamilyTXT.Text)
    sqlRS.Fields("OperationNumber") = Val(ProcessAttr.OpNumTXT.Text)
    sqlRS.Fields("OperationDescription") = UCase(ProcessAttr.OpDescTXT.Text)
    If IsNull(sqlRS.Fields("Obsolete")) Then
        sqlRS.Fields("Obsolete") = ProcessAttr.ObsoleteCheck.Value
    End If
    If sqlRS.Fields("Obsolete") Then
        i = 1
    Else
        i = 0
    End If
    If i <> ProcessAttr.ObsoleteCheck.Value Then
        sqlRS.Fields("Obsolete") = ProcessAttr.ObsoleteCheck.Value
        If ProcessAttr.ObsoleteCheck.Value = 1 Then
            ToolChanges(0, ToolChangeCntr) = "STATUS"
            ToolChanges(6, ToolChangeCntr) = "OBSOLETE"
            ToolChangeCntr = ToolChangeCntr + 1
        Else
            ToolChanges(0, ToolChangeCntr) = "STATUS"
            ToolChanges(6, ToolChangeCntr) = "ACTIVE"
            ToolChangeCntr = ToolChangeCntr + 1
        End If
    End If
    sqlRS.Fields("Customer") = UCase(ProcessAttr.CustomerTXT.Text)
    If IsNull(sqlRS.Fields("AnnualVolume")) Then
        
        sqlRS.Fields("AnnualVolume") = Val(ProcessAttr.AnnualVolumeTXT.Text)
    Else
        If sqlRS.Fields("AnnualVolume") <> Val(ProcessAttr.AnnualVolumeTXT.Text) Then
            ToolChanges(0, ToolChangeCntr) = "VOLUME"
            ToolChanges(1, ToolChangeCntr) = Val(ProcessAttr.AnnualVolumeTXT.Text)
            ToolChanges(2, ToolChangeCntr) = sqlRS.Fields("AnnualVolume")
            ToolChangeCntr = ToolChangeCntr + 1
            sqlRS.Fields("AnnualVolume") = Val(ProcessAttr.AnnualVolumeTXT.Text)
        End If
    End If
    sqlRS.Fields("MultiTurret") = ProcessAttr.MultiTurretLathe.Value
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    SetMultiTurret
End Sub

Public Sub UpdateToolDetails()
    Dim newseq As Long
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(toolID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("ToolNumber") = Val(ToolAttr.ToolNumberTXT.Text)
    sqlRS.Fields("OpDescription") = UCase(ToolAttr.OpDescTXT.Text)
    sqlRS.Fields("Alternate") = ToolAttr.AlternateCHECK.Value
    sqlRS.Fields("PartSpecific") = ToolAttr.PartNumberCheck.Value
    sqlRS.Fields("AdjustedVolume") = Val(ToolAttr.AdjustedVolume.Text)
    sqlRS.Fields("ToolOrder") = Val(ToolAttr.SequenceTxt.Text)
    sqlRS.Fields("ToolLength") = Val(ToolAttr.ToolLengthOffsetTXT.Text)
    sqlRS.Fields("OffsetNumber") = Val(ToolAttr.OffsetNumberTXT.Text)
    toolID = sqlRS.Fields("TOOLID")
    newseq = Val(ToolAttr.SequenceTxt.Text)
    If Not MultiTurret Then
        sqlRS.Fields("Turret") = "A"
    Else
        If ToolAttr.TurretBOption.Value = True Then
            sqlRS.Fields("Turret") = "B"
        Else
            sqlRS.Fields("Turret") = "A"
        End If
    End If
    sqlRS.Update
    sqlRS.Close
    ReSequenceTools newseq
    Set sqlRS = Nothing
    UpdateToolPartNumbers
    BuildToolList
End Sub
Public Function RetrieveOriginalProcId(procId As Integer) As Integer
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(procId), sqlConn, adOpenKeyset, adLockOptimistic
   If sqlRS.EOF Then
        RetrieveOriginalProcId = -1
    Else
       RetrieveOriginalProcId = sqlRS.Fields("OriginalProcessID")
    End If
End Function
Public Sub UpdateItemDetails()
    Dim strStream As ADODB.Stream
    Dim changed As Boolean
    changed = False
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID), sqlConn, adOpenKeyset, adLockOptimistic
    If OldCribID <> ItemAttri.CribNumberIDTXT.Text Then
        ToolChanges(0, ToolChangeCntr) = "ADDTOOL"
        ToolChanges(1, ToolChangeCntr) = UCase(ItemAttri.CribNumberIDTXT.Text)
        ToolChangeCntr = ToolChangeCntr + 1
        ToolChanges(0, ToolChangeCntr) = "REMOVETOOL"
        ToolChanges(1, ToolChangeCntr) = OldCribID
        ToolChangeCntr = ToolChangeCntr + 1
        
        changed = True
    End If
    If sqlRS.EOF Then
        Exit Sub
    End If
    sqlRS.Fields("ToolType") = UCase(ItemAttri.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = UCase(ItemAttri.ItemNumberCOMBO.Text)
    sqlRS.Fields("CribToolID") = ItemAttri.CribNumberIDTXT.Text
    sqlRS.Fields("Manufacturer") = UCase(ItemAttri.ManufacturerTXT.Text)
    If sqlRS.Fields("Quantity") <> ItemAttri.QuantityTXT.Text And Not changed Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "USAGE"
        ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    sqlRS.Fields("Quantity") = ItemAttri.QuantityTXT.Text
    sqlRS.Fields("Consumable") = ItemAttri.ConsumableCHECK.Value
    If sqlRS.Fields("NumberOfCuttingEdges") <> Val(ItemAttri.CuttingEdgesTXT.Text) And Not changed Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "USAGE"
        ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    sqlRS.Fields("NumberOfCuttingEdges") = Val(ItemAttri.CuttingEdgesTXT.Text)
    If sqlRS.Fields("QuantityPerCuttingEdge") <> Val(ItemAttri.ToolLifeTXT.Text) And Not changed Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "USAGE"
        ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    Dim i
    If IsNull(sqlRS.Fields("ToolbossStock")) = False Then
       If sqlRS.Fields("ToolbossStock") Then
         i = 1
       Else
         i = 0
       End If
    Else
        i = 0
    End If
    If i <> ItemAttri.TBStock.Value And OldCribID = ItemAttri.CribNumberIDTXT.Text Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "STOCK"
        ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    
    sqlRS.Fields("NumOfRegrinds") = Val(ItemAttri.NumofRegrindsTXT.Text)
    sqlRS.Fields("QtyPerRegrind") = Val(ItemAttri.ToolLifeRegrindTXT.Text)
    sqlRS.Fields("Regrindable") = ItemAttri.RegrindableChk.Value
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(ItemAttri.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(ItemAttri.AdditionalNotesTXT.Text)
    sqlRS.Fields("ToolbossStock") = ItemAttri.TBStock.Value
     'Add the image to the database
    Dim IsDuplicate As Boolean
    Dim k As Integer

    If ("" <> ItemAttri.txtPicture) And (ItemAttri.cbDeletePic.Value <> 1) Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
        strStream.LoadFromFile ItemAttri.txtPicture
        sqlRS.Fields("ItemImage").Value = strStream.Read
        strStream.Close
        Set strStream = Nothing
        changed = True
                    
        'CHECK IF PICTURE CHANGE HAS ALREADY BEEN ADDED TO THE LIST OF TOOLCHANGES
        IsDuplicate = False
        For k = 0 To ToolChangeCntr - 1
            If ToolChanges(0, k) = "PICTURES" And _
               ToolChanges(2, k) = Str(sqlRS.Fields("ItemID")) Then
                 IsDuplicate = True
            End If
        Next
        If IsDuplicate = False Then
            ToolChanges(0, ToolChangeCntr) = "PICTURES"
            ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
            ToolChanges(2, ToolChangeCntr) = Str(sqlRS.Fields("ItemID"))
            ToolChangeCntr = ToolChangeCntr + 1
        End If
    End If

    If (ItemAttri.cbDeletePic.Value = 1) Then
        ' Is there is a pic displayed now
        If (IsNull(sqlRS.Fields("ItemImage").Value) = False) Then
            sqlRS.Fields("ItemImage").Value = Null
        End If
        
        ' The colItemImages collection is not created if we are working live before initial submission
        If WorkingLive = False Then
            ' If there was a picture originally then add this tool change
            If colItemImages.Item(Str(sqlRS.Fields("ItemID"))) = "T" Then
                'CHECK IF PICTURE CHANGE HAS ALREADY BEEN ADDED TO THE LIST OF TOOLCHANGES
                IsDuplicate = False
                For k = 0 To ToolChangeCntr - 1
                    If ToolChanges(0, k) = "PICTURES" And _
                        ToolChanges(2, k) = Str(sqlRS.Fields("ItemID")) Then
                         IsDuplicate = True
                    End If
                Next
                'add to toolchange list picture item if not already there
                If IsDuplicate = False Then
                    ToolChanges(0, ToolChangeCntr) = "PICTURES"
                    ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
                    ToolChanges(2, k) = Str(sqlRS.Fields("ItemID"))
                    ToolChangeCntr = ToolChangeCntr + 1
                End If
            Else
              'If there was not a picture originally then remove from toolchange list if present
                IsDuplicate = False
                For k = 0 To ToolChangeCntr - 1
                    If ToolChanges(0, k) = "PICTURES" And _
                        ToolChanges(2, k) = Str(sqlRS.Fields("ItemID")) Then
                        IsDuplicate = True
                    End If
                Next
                If IsDuplicate = True Then
                 'Remove toolchange item
                  Dim NewToolChanges(6, 200) As String
                  Dim NewToolChangeCntr As Integer
                  Dim itc As Integer
                  ' Copy ToolChanges to temp array skipping item to be removed
                  For itc = 0 To ToolChangeCntr - 1
                      If Not (ToolChanges(0, itc) = "PICTURES" And _
                              ToolChanges(2, itc) = Str(sqlRS.Fields("ItemID"))) Then
                          NewToolChanges(0, itc) = ToolChanges(0, itc)
                          NewToolChanges(1, itc) = ToolChanges(1, itc)
                          NewToolChanges(2, itc) = ToolChanges(2, itc)
                          NewToolChanges(3, itc) = ToolChanges(3, itc)
                          NewToolChanges(4, itc) = ToolChanges(4, itc)
                          NewToolChanges(5, itc) = ToolChanges(5, itc)
                      End If
                  Next
                  ' There is now 1 less toolchanges
                  ToolChangeCntr = ToolChangeCntr - 1
                  ' Copy to global variable
                  For itc = 0 To ToolChangeCntr - 1
                      ToolChanges(0, itc) = NewToolChanges(0, itc)
                      ToolChanges(1, itc) = NewToolChanges(1, itc)
                      ToolChanges(2, itc) = NewToolChanges(2, itc)
                      ToolChanges(3, itc) = NewToolChanges(3, itc)
                      ToolChanges(4, itc) = NewToolChanges(4, itc)
                      ToolChanges(5, itc) = NewToolChanges(5, itc)
                  Next
                  ' Remove last tool change it is not valid
                  ToolChanges(0, itc) = ""
                  ToolChanges(1, itc) = ""
                  ToolChanges(2, itc) = ""
                  ToolChanges(3, itc) = ""
                  ToolChanges(4, itc) = ""
                  ToolChanges(5, itc) = ""
                End If
            End If
        End If ' Not needed if working live
        ItemAttri.cbDeletePic.Value = 0
    End If
    Dim i1 As Long
    i1 = sqlRS.Fields("ProcessID")
    sqlRS.Update
    toolID = sqlRS.Fields("toolid")
    sqlRS.Close
    Set sqlRS = Nothing
    BuildToolList

    OldCribID = ""
End Sub

Public Sub BuildToolList()
    ToolList.TreeView1.Nodes.Clear
    Set sqlRS = New ADODB.Recordset
    Set SQLRS2 = New ADODB.Recordset
    sqlRS.Open "SELECT TOOLID, TOOLNUMBER, OPDESCRIPTION FROM [TOOLLIST TOOL] WHERE PROCESSID = " + Str(processId) + " ORDER BY TOOLORDER", sqlConn
    ToolList.TreeView1.Nodes.Add , , "Process" + Trim(Str(processId)), "Process #" + Str(processId)
    ToolList.TreeView1.Nodes.Item("Process" + Trim(Str(processId))).Expanded = True
    While Not sqlRS.EOF
        ToolList.TreeView1.Nodes.Add "Process" + Trim(Str(processId)), tvwChild, "TOOL" + Trim(Str(sqlRS.Fields("TOOLID"))), "TOOL " + Trim(Str(sqlRS.Fields("TOOLNUMBER"))) + " - " + sqlRS.Fields("OPDESCRIPTION")
        If toolID = sqlRS.Fields("TOOLID") Then
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Expanded = True
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Selected = True
            LastToolDescription = sqlRS.Fields("OpDescription")
        Else
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Expanded = False
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Selected = False
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
   InitToolListItems
End Sub
Public Sub BuildFixtureList()
    Dim invItem As Inventry
    ToolList.TreeView4.Nodes.Clear
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT ITEMID, TOOLDESCRIPTION, CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID = " + Str(processId) + " ORDER BY ITEMID", sqlConn
    ToolList.TreeView4.Nodes.Add , , "Process" + Trim(Str(processId)), "Process #" + Str(processId)
    ToolList.TreeView4.Nodes.Item("Process" + Trim(Str(processId))).Expanded = True
    While Not sqlRS.EOF
        If (False = Exists(inv, sqlRS.Fields("CRIBTOOLID"))) Then
            Set invItem = New Inventry
            invItem.itemNumber = sqlRS.Fields("CRIBTOOLID")
            invItem.description1 = GetInvDescription(sqlRS.Fields("CRIBTOOLID"))
            inv.Add invItem, sqlRS.Fields("CRIBTOOLID")
        End If
        ToolList.TreeView4.Nodes.Add "Process" + Trim(Str(processId)), tvwChild, "FIXT" + Trim(Str(sqlRS.Fields("ITEMID"))), GetItemDescription(sqlRS.Fields("CRIBTOOLID"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub BuildMiscList()
    Dim invItem As Inventry
    ToolList.TreeView2.Nodes.Clear
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT ITEMID,CRIBTOOLID, TOOLDESCRIPTION FROM [TOOLLIST MISC] WHERE PROCESSID = " + Str(processId) + " ORDER BY ITEMID", sqlConn
    ToolList.TreeView2.Nodes.Add , , "Process" + Trim(Str(processId)), "Process #" + Str(processId)
    ToolList.TreeView2.Nodes.Item("Process" + Trim(Str(processId))).Expanded = True
    While Not sqlRS.EOF
        'Is it already in collection
        If (False = Exists(inv, sqlRS.Fields("CRIBTOOLID"))) Then
            Set invItem = New Inventry
            invItem.itemNumber = sqlRS.Fields("CRIBTOOLID")
            invItem.description1 = GetInvDescription(sqlRS.Fields("CRIBTOOLID"))
            inv.Add invItem, sqlRS.Fields("CRIBTOOLID")
        End If
        ToolList.TreeView2.Nodes.Add "Process" + Trim(Str(processId)), tvwChild, "MISC" + Trim(Str(sqlRS.Fields("ITEMID"))), GetItemDescription(sqlRS.Fields("CRIBTOOLID"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Function Exists(ByVal oCol As Collection, ByVal vKey As Variant) As Boolean

    On Error Resume Next
    oCol.Item vKey
    Exists = (Err.Number = 0)
    Err.Clear

End Function
Public Sub InitToolListItems()
    Dim invItem As Inventry
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT ITEMID, TOOLID, TOOLTYPE , CRIBTOOLID, TOOLDESCRIPTION FROM [TOOLLIST ITEM] WHERE PROCESSID = " + Str(processId), sqlConn
    While Not sqlRS.EOF
        If (False = Exists(inv, sqlRS.Fields("CRIBTOOLID"))) Then
            Set invItem = New Inventry
            invItem.itemNumber = sqlRS.Fields("CRIBTOOLID")
            invItem.description1 = GetInvDescription(sqlRS.Fields("CRIBTOOLID"))
            inv.Add invItem, sqlRS.Fields("CRIBTOOLID")
        End If
        ToolList.TreeView1.Nodes.Add "TOOL" + Trim(Str(sqlRS.Fields("TOOLID"))), tvwChild, "ITEM" + _
            Trim(Str(sqlRS.Fields("ITEMID"))), GetItemDescription(sqlRS.Fields("CRIBTOOLID"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing

End Sub
Public Function GetInvDescription(CribIdNumber As String) As String
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT DESCRIPTION1 FROM [INVENTRY] WHERE ITEMNUMBER = '" + CribIdNumber + "'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        GetInvDescription = CribRS.Fields("Description1")
    Else
        GetInvDescription = "Err"
    End If
    CribRS.Close
    Set CribRS = Nothing
End Function
        
        

Public Function GetItemDescription(CribIdNumber As String) As String
    Dim invItem As Inventry
    If (Exists(inv, CribIdNumber)) Then
        Set invItem = inv(CribIdNumber)
        GetItemDescription = invItem.description1
    Else
        GetItemDescription = "Err"
    End If
End Function


Public Sub ClearProcessFields()
    ProcessAttr.PartFamilyTXT.Text = ""
    ProcessAttr.OpNumTXT.Text = ""
    ProcessAttr.OpDescTXT.Text = ""
    ProcessAttr.ObsoleteCheck.Value = 0
    ProcessAttr.CustomerTXT.Text = ""
    ProcessAttr.SelectedPartsList.Clear
    ProcessAttr.AllPartNumbersList2.ListItems.Clear
    ProcessAttr.SelectedPlantsList.Clear
    ProcessAttr.AllPlantList.Clear
    ProcessAttr.AnnualVolumeTXT.Text = ""
    ProcessAttr.MultiTurretLathe.Value = 0
End Sub


Public Sub ClearToolFields()
    ToolAttr.ToolNumberTXT.Text = ""
    ToolAttr.OpDescTXT.Text = ""
    ToolAttr.AlternateCHECK.Value = 0
    ToolAttr.AdjustedVolume.Text = ""
    ToolAttr.PartNumberCheck.Value = 0
    ToolAttr.SelectedPartsList.Clear
    ToolAttr.AllPartNumbersList.Clear
    ToolAttr.DisableMultiPart
    ToolAttr.SequenceList.ListItems.Clear
    ToolAttr.SequenceTxt.Text = Str(GetNextSequence)
    ToolAttr.ToolLengthOffsetTXT.Text = ""
    ToolAttr.OffsetNumberTXT.Text = ""
    GetAvailableToolPartNumbers
    PopulateSequence
End Sub

Public Sub ClearItemFields()
    ItemAttri.ItemGroupTXT.Text = ""
    ItemAttri.ItemNumberCOMBO.Text = ""
    ItemAttri.ManufacturerTXT.Text = ""
    ItemAttri.AdditionalNotesTXT.Text = ""
    ItemAttri.QuantityTXT.Text = ""
    ItemAttri.ConsumableCHECK.Value = 0
    ItemAttri.CuttingEdgesTXT.Text = ""
    ItemAttri.ToolLifeTXT.Text = ""
    ItemAttri.CribNumberIDTXT.Text = ""
    ItemAttri.QtyOnHandTXT.Text = ""
    ItemAttri.NumofRegrindsTXT = ""
    ItemAttri.ToolLifeRegrindTXT = ""
    ItemAttri.RegrindableChk.Value = 0
    ItemAttri.TBStock.Value = 0
    ItemAttri.BinTxt = ""
    ItemAttri.CostPerPartTXT = ""
    ItemAttri.CostTXT = ""
    ItemAttri.MonthlyUsageTXT = ""
    ItemAttri.txtPicture = ""
    ItemAttri.imgItem = LoadPicture()
    
    OldCribID = ""
    If ItemAttri.ItemNumberCOMBO.ListCount = 0 Then
        PopulateItemList
    End If
End Sub

Public Sub AddToolSub()
    Dim newseq As Integer
    Set sqlRS = New ADODB.Recordset
    sqlRS.CursorLocation = adUseClient
    sqlRS.Open "[TOOLLIST TOOL]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Fields("ToolNumber") = Val(ToolAttr.ToolNumberTXT.Text)
    sqlRS.Fields("OpDescription") = UCase(ToolAttr.OpDescTXT.Text)
    sqlRS.Fields("Alternate") = ToolAttr.AlternateCHECK.Value
    sqlRS.Fields("PartSpecific") = ToolAttr.PartNumberCheck.Value
    sqlRS.Fields("AdjustedVolume") = Val(ToolAttr.AdjustedVolume.Text)
    sqlRS.Fields("ToolOrder") = Val(ToolAttr.SequenceTxt.Text)
    sqlRS.Fields("ToolLength") = Val(ToolAttr.ToolLengthOffsetTXT.Text)
    sqlRS.Fields("OffsetNumber") = Val(ToolAttr.OffsetNumberTXT.Text)
    If Not MultiTurret Then
        sqlRS.Fields("Turret") = "A"
    Else
        If ToolAttr.TurretBOption.Value = True Then
            sqlRS.Fields("Turret") = "B"
        Else
            sqlRS.Fields("Turret") = "A"
        End If
    End If
    newseq = Val(ToolAttr.SequenceTxt.Text)
    sqlRS.Update
    toolID = sqlRS.Fields("TOOLID")
    sqlRS.Close
    Set sqlRS = Nothing
    ReSequenceTools ((newseq))
    UpdateToolPartNumbers
    BuildToolList
End Sub

Public Sub UpdateToolPartNumbers()
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST ToolPARTNUMBER] WHERE TOOLID =" + Str(toolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[ToolList TOOLPartNumber]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Dim i As Integer
    i = 0
    While i < ToolAttr.SelectedPartsList.ListCount
        sqlRS.AddNew
        sqlRS.Fields("TOOLID") = toolID
        sqlRS.Fields("PartNumber") = Trim(ToolAttr.SelectedPartsList.List(i))
        sqlRS.Update
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub DeleteToolSub()
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(toolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE TOOLID =" + Str(toolID), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        ToolChanges(0, ToolChangeCntr) = "REMOVETOOL"
        ToolChanges(1, ToolChangeCntr) = sqlRS.Fields("CribToolID")
        ToolChangeCntr = ToolChangeCntr + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST ITEM] WHERE TOOLID =" + Str(toolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST TOOLPARTNUMBER] WHERE TOOLID =" + Str(toolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    MDIForm1.TabDock.FormHide "Tool Details"
    Set sqlCMD = Nothing
    BuildToolList
    RefreshReport
End Sub

Public Sub DeleteProcessSub(pID As Long)
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST REV] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST PLANT] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
End Sub

Public Sub AddItemSub()
    Set sqlRS = New ADODB.Recordset
    'If item already exists in the Tool List then do a USAGE tool change
    ' because the ADDTOOL tool change will get removed in the ConsolidateChanges function
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE CribToolID = '" + ItemAttri.CribNumberIDTXT.Text + "'", sqlConn, adOpenKeyset, adLockOptimistic
    If Not sqlRS.EOF Then
        'If item already exists then do a USAGE tool change
        ToolChanges(0, ToolChangeCntr) = "USAGE"
        ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    sqlRS.Close
    sqlRS.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ToolType") = UCase(ItemAttri.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = ItemAttri.ItemNumberCOMBO.Text
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Fields("ToolID") = toolID
    sqlRS.Fields("CribToolID") = ItemAttri.CribNumberIDTXT.Text
    sqlRS.Fields("Consumable") = ItemAttri.ConsumableCHECK.Value
    sqlRS.Fields("Manufacturer") = UCase(ItemAttri.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = ItemAttri.QuantityTXT.Text
    sqlRS.Fields("NumberOfCuttingEdges") = Val(ItemAttri.CuttingEdgesTXT.Text)
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(ItemAttri.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(ItemAttri.AdditionalNotesTXT.Text)
    sqlRS.Fields("NumOfRegrinds") = Val(ItemAttri.NumofRegrindsTXT.Text)
    sqlRS.Fields("QtyPerRegrind") = Val(ItemAttri.ToolLifeRegrindTXT.Text)
    sqlRS.Fields("Regrindable") = ItemAttri.RegrindableChk.Value
    sqlRS.Fields("ToolbossStock") = ItemAttri.TBStock.Value
    If ("" <> ItemAttri.txtPicture) And (ItemAttri.cbDeletePic.Value <> 1) Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
        strStream.LoadFromFile ItemAttri.txtPicture
        sqlRS.Fields("ItemImage").Value = strStream.Read
        strStream.Close
        Set strStream = Nothing
        changed = True
    End If
    If (ItemAttri.cbDeletePic.Value = 1) Then
        ' Is there is a pic displayed now
        If (IsNull(sqlRS.Fields("ItemImage").Value) = False) Then
            sqlRS.Fields("ItemImage").Value = Null
        End If
        ItemAttri.cbDeletePic.Value = 0
    End If
    sqlRS.Update
    sqlRS.Close
    sqlRS.Open "select * from [TOOLLIST ITEM] order by itemid desc", sqlConn, adOpenKeyset, adLockReadOnly
    newItemId = sqlRS.Fields("ItemID")
    sqlRS.Close
    'Initial Toollists do not go through the routing process but I still need to update the colItemImages
    'in case the item is edited after the initial add
    sqlRS.Open "select * from [TOOLLIST ITEM] where itemid = " + Str(newItemId), sqlConn, adOpenKeyset, adLockReadOnly
    If (IsNull(sqlRS.Fields("ItemImage")) = True) Then
        colItemImages.Add "F", Str(newItemId)
    Else
        colItemImages.Add "T", Str(newItemId)
    End If
    sqlRS.Close
    Set sqlRS = Nothing
    BuildToolList
    'ADDTOOL maybe removed if this item already exists in the ToolList
    ToolChanges(0, ToolChangeCntr) = "ADDTOOL"
    ToolChanges(1, ToolChangeCntr) = ItemAttri.CribNumberIDTXT.Text
    ToolChanges(2, ToolChangeCntr) = newItemId
    ToolChangeCntr = ToolChangeCntr + 1
    OldCribID = ""
End Sub

Public Sub DeleteItemSub()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID = " + Str(itemID), sqlConn, adOpenKeyset, adLockReadOnly
    If sqlRS.RecordCount > 0 Then
        OldCribID = sqlRS.Fields("CribToolID")
    End If
    sqlRS.Close
    ToolChanges(0, ToolChangeCntr) = "REMOVETOOL"
    ToolChanges(1, ToolChangeCntr) = OldCribID
    ToolChangeCntr = ToolChangeCntr + 1
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    
    'If item still exists in the Tool List then add a USAGE tool change entry
    ' because the REMOVETOOL tool change may get removed in the ConsolidateChanges function
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE CribToolID = '" + OldCribID + "'", sqlConn, adOpenKeyset, adLockOptimistic
    If Not sqlRS.EOF Then
        'If item still exists then do a USAGE tool change
        ToolChanges(0, ToolChangeCntr) = "USAGE"
        ToolChanges(1, ToolChangeCntr) = OldCribID
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    Set sqlRS = Nothing
    
    MDIForm1.TabDock.FormHide "Item Details"
    BuildToolList
    RefreshReport
End Sub

Public Sub PopulateDeleteView()
    Dim itmx2 As ListItem
    Set sqlRS = New ADODB.Recordset
    DeleteProcess.ListView1.ListItems.Clear
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION", sqlConn
    While Not sqlRS.EOF
        Set itmx2 = DeleteProcess.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSID"))
        If Not IsNull(sqlRS.Fields("CUSTOMER")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("CUSTOMER"))
        End If
        If Not IsNull(sqlRS.Fields("PARTFAMILY")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("PARTFAMILY"))
        End If
        If Not IsNull(sqlRS.Fields("OPERATIONDESCRIPTION")) Then
            itmx2.SubItems(3) = Trim(sqlRS.Fields("OPERATIONDESCRIPTION"))
        End If
        If Not IsNull(sqlRS.Fields("OPERATIONNUMBER")) Then
            itmx2.SubItems(4) = Trim(sqlRS.Fields("OPERATIONNUMBER"))
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub PopulateCopyView()
    Dim itmx2 As ListItem
    Set sqlRS = New ADODB.Recordset
    DeleteProcess.ListView1.ListItems.Clear
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION", sqlConn
    While Not sqlRS.EOF
        Set itmx2 = DeleteProcess.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSID"))
        If Not IsNull(sqlRS.Fields("CUSTOMER")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("CUSTOMER"))
        End If
        If Not IsNull(sqlRS.Fields("PARTFAMILY")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("PARTFAMILY"))
        End If
        If Not IsNull(sqlRS.Fields("OPERATIONDESCRIPTION")) Then
            itmx2.SubItems(3) = Trim(sqlRS.Fields("OPERATIONDESCRIPTION"))
        End If
        If Not IsNull(sqlRS.Fields("OPERATIONNUMBER")) Then
            itmx2.SubItems(4) = Trim(sqlRS.Fields("OPERATIONNUMBER"))
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub PopulateItemList()
    Dim string1, string2, string3 As String
    string1 = ItemAttri.ItemNumberCOMBO.Text
    string2 = MiscItem.ItemNumberCOMBO.Text
    string3 = FixtureItem.ItemNumberCOMBO.Text
    ItemAttri.ItemNumberCOMBO.Clear
    MiscItem.ItemNumberCOMBO.Clear
    FixtureItem.ItemNumberCOMBO.Clear
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT DISTINCT DESCRIPTION1 FROM [INVENTRY] WHERE DESCRIPTION1 is not NULL ORDER BY DESCRIPTION1", CribConn
    While Not CribRS.EOF
        ItemAttri.ItemNumberCOMBO.AddItem CribRS.Fields("DESCRIPTION1")
        MiscItem.ItemNumberCOMBO.AddItem CribRS.Fields("DESCRIPTION1")
        FixtureItem.ItemNumberCOMBO.AddItem CribRS.Fields("DESCRIPTION1")
        CribRS.MoveNext
    Wend
    CribRS.Close
    Set CribRS = Nothing
    ItemAttri.ItemNumberCOMBO.Text = string1
    MiscItem.ItemNumberCOMBO.Text = string2
    FixtureItem.ItemNumberCOMBO.Text = string3
End Sub

Public Sub BuildRevList()
    ToolList.TreeView3.Nodes.Clear
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT REVISIONID,REVISION, [REVISION DESCRIPTION],[REVISION DATE],[REVISION BY] FROM [TOOLLIST REV] WHERE PROCESSID = " + Str(processId) + " ORDER BY REVISION", sqlConn
    ToolList.TreeView3.Nodes.Add , , "Process" + Trim(Str(processId)), "Process #" + Str(processId)
    ToolList.TreeView3.Nodes.Item("Process" + Trim(Str(processId))).Expanded = True
    While Not sqlRS.EOF
        ToolList.TreeView3.Nodes.Add "Process" + Trim(Str(processId)), tvwChild, "REV" + Trim(Str(sqlRS.Fields("REVISIONID"))), "REVISION - " + Trim(Str(sqlRS.Fields("REVISION")))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub



Public Sub ClearMiscFields()
    MiscItem.ItemGroupTXT.Text = ""
    MiscItem.ItemNumberCOMBO.Text = ""
    MiscItem.ManufacturerTXT.Text = ""
    MiscItem.AdditionalNotesTXT.Text = ""
    MiscItem.QuantityTXT.Text = ""
    MiscItem.ConsumableCHECK.Value = 0
    MiscItem.TBStock.Value = 0
    MiscItem.CuttingEdgesTXT.Text = ""
    MiscItem.ToolLifeTXT.Text = ""
    MiscItem.CribNumberIDTXT.Text = ""
    MiscItem.QtyOnHandTXT.Text = ""
    OldCribID = ""
    If MiscItem.ItemNumberCOMBO.ListCount = 0 Then
        PopulateItemList
    End If
End Sub
Public Sub ClearFixtureFields()
    FixtureItem.ItemGroupTXT.Text = ""
    FixtureItem.ItemNumberCOMBO.Text = ""
    FixtureItem.ManufacturerTXT.Text = ""
    FixtureItem.AdditionalNotesTXT.Text = ""
    FixtureItem.QuantityTXT.Text = ""
    FixtureItem.CribNumberIDTXT.Text = ""
    FixtureItem.QtyOnHandTXT.Text = ""
    FixtureItem.DetailNumberTxt.Text = ""
    FixtureItem.TBStock.Value = 0
    
    OldCribID = ""
    If FixtureItem.ItemNumberCOMBO.ListCount = 0 Then
        PopulateItemList
    End If
End Sub


Public Sub ClearRevisionFields()
    RevisionForm.RevByTXT.Text = ""
    RevisionForm.RevDate = Date
    RevisionForm.RevDescTXT = ""
    RevisionForm.RevNumTXT = ""
End Sub

Public Sub GetMiscDetails()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID), sqlConn
    MiscItem.ItemGroupTXT.Text = sqlRS.Fields("ToolType")
    MiscItem.ItemNumberCOMBO.Text = sqlRS.Fields("ToolDescription")
    OldCribID = sqlRS.Fields("CribToolID")
    MiscItem.ManufacturerTXT.Text = sqlRS.Fields("Manufacturer")
    If Not IsNull(sqlRS.Fields("cribtoolid")) Then
        MiscItem.CribNumberIDTXT.Text = sqlRS.Fields("CribToolID")
    End If
    MiscItem.QuantityTXT.Text = sqlRS.Fields("Quantity")
    MiscItem.CuttingEdgesTXT.Text = sqlRS.Fields("NumberOfCuttingEdges")
    MiscItem.ToolLifeTXT.Text = sqlRS.Fields("QuantityPerCuttingEdge")
    MiscItem.AdditionalNotesTXT.Text = sqlRS.Fields("AdditionalNotes")
    If sqlRS.Fields("Consumable") Then
        i = 1
    Else
        i = 0
    End If
    MiscItem.ConsumableCHECK.Value = i
    
    If IsNull(sqlRS.Fields("ToolbossStock")) = False Then
       If sqlRS.Fields("ToolbossStock") Then
         i = 1
       Else
         i = 0
       End If
    Else
       i = 0
    End If
    
    MiscItem.TBStock.Value = i
    MiscQty
    sqlRS.Close
    Set sqlRS = Nothing

End Sub

Public Sub GetFixtureDetails()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST FIXTURE] WHERE ITEMID =" + Str(FixtureToolID), sqlConn
    FixtureItem.ItemGroupTXT.Text = sqlRS.Fields("ToolType")
    FixtureItem.ItemNumberCOMBO.Text = sqlRS.Fields("ToolDescription")
    OldCribID = sqlRS.Fields("CribToolID")
    FixtureItem.ManufacturerTXT.Text = sqlRS.Fields("Manufacturer")
    If Not IsNull(sqlRS.Fields("cribtoolid")) Then
        FixtureItem.CribNumberIDTXT.Text = sqlRS.Fields("CribToolID")
    End If
    FixtureItem.QuantityTXT.Text = sqlRS.Fields("Quantity")
    FixtureItem.AdditionalNotesTXT.Text = sqlRS.Fields("AdditionalNotes")
    FixtureItem.DetailNumberTxt.Text = sqlRS.Fields("DetailNumber")
    If IsNull(sqlRS.Fields("ToolbossStock")) = False Then
       If sqlRS.Fields("ToolbossStock") Then
         i = 1
       Else
         i = 0
       End If
    Else
       i = 0
    End If
   
    FixtureItem.TBStock.Value = i
    FixtureQty
    sqlRS.Close
    Set sqlRS = Nothing

End Sub

Public Sub GetRevisionDetails()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST REV] WHERE REVISIONID =" + Str(RevisionID), sqlConn
    RevisionForm.RevByTXT.Text = sqlRS.Fields("Revision By")
    RevisionForm.RevNumTXT.Text = sqlRS.Fields("Revision")
    RevisionForm.RevDescTXT.Text = sqlRS.Fields("Revision Description")
    RevisionForm.RevDate = sqlRS.Fields("Revision Date")
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub UpdateRevisionDetails()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST REV] WHERE REVISIONID =" + Str(RevisionID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("Revision By") = UCase(RevisionForm.RevByTXT.Text)
    sqlRS.Fields("Revision") = UCase(RevisionForm.RevNumTXT.Text)
    sqlRS.Fields("Revision Description") = UCase(RevisionForm.RevDescTXT.Text)
    sqlRS.Fields("Revision Date") = RevisionForm.RevDate
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildRevList
    RefreshReport
End Sub

Public Sub UpdateMiscDetails()
    Dim changed As Boolean
    changed = False
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID), sqlConn, adOpenKeyset, adLockOptimistic
    If OldCribID <> MiscItem.CribNumberIDTXT.Text Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "ADDTOOLM"
        ToolChanges(1, ToolChangeCntr) = MiscItem.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
        ToolChanges(0, ToolChangeCntr) = "REMOVETOOLM"
        ToolChanges(1, ToolChangeCntr) = OldCribID
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    Dim i
    If IsNull(sqlRS.Fields("ToolbossStock")) = False Then
       If sqlRS.Fields("ToolbossStock") Then
         i = 1
       Else
         i = 0
       End If
    Else
        i = 0
    End If
    If i <> MiscItem.TBStock.Value And OldCribID = MiscItem.CribNumberIDTXT.Text Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "STOCK"
        ToolChanges(1, ToolChangeCntr) = MiscItem.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("ToolType") = UCase(MiscItem.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = UCase(MiscItem.ItemNumberCOMBO.Text)
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Fields("CribToolID") = MiscItem.CribNumberIDTXT.Text
    sqlRS.Fields("Consumable") = MiscItem.ConsumableCHECK.Value
    sqlRS.Fields("ToolbossStock") = MiscItem.TBStock.Value
    sqlRS.Fields("Manufacturer") = UCase(MiscItem.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = MiscItem.QuantityTXT.Text
    sqlRS.Fields("NumberOfCuttingEdges") = Val(MiscItem.CuttingEdgesTXT.Text)
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(MiscItem.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(MiscItem.AdditionalNotesTXT.Text)
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildMiscList
    RefreshReport
End Sub

Public Sub UpdateFixtureDetails()
    Dim changed As Boolean
    changed = False
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST FIXTURE] WHERE ITEMID =" + Str(FixtureToolID), sqlConn, adOpenKeyset, adLockOptimistic
    If OldCribID <> FixtureItem.CribNumberIDTXT.Text Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "ADDTOOLF"
        ToolChanges(1, ToolChangeCntr) = FixtureItem.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
        ToolChanges(0, ToolChangeCntr) = "REMOVETOOLF"
        ToolChanges(1, ToolChangeCntr) = OldCribID
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    Dim i
    If IsNull(sqlRS.Fields("ToolbossStock")) = False Then
       If sqlRS.Fields("ToolbossStock") Then
         i = 1
       Else
         i = 0
       End If
    Else
        i = 0
    End If
    If i <> FixtureItem.TBStock.Value And OldCribID = FixtureItem.CribNumberIDTXT.Text Then
        changed = True
        ToolChanges(0, ToolChangeCntr) = "STOCK"
        ToolChanges(1, ToolChangeCntr) = FixtureItem.CribNumberIDTXT.Text
        ToolChangeCntr = ToolChangeCntr + 1
    End If
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST FIXTURE] WHERE ITEMID =" + Str(FixtureToolID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("ToolType") = UCase(FixtureItem.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = UCase(FixtureItem.ItemNumberCOMBO.Text)
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Fields("ToolbossStock") = FixtureItem.TBStock.Value
    sqlRS.Fields("CribToolID") = FixtureItem.CribNumberIDTXT.Text
    sqlRS.Fields("Manufacturer") = UCase(FixtureItem.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = FixtureItem.QuantityTXT.Text
    sqlRS.Fields("AdditionalNotes") = UCase(FixtureItem.AdditionalNotesTXT.Text)
    sqlRS.Fields("DetailNumber") = FixtureItem.DetailNumberTxt.Text
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildFixtureList
    RefreshReport
End Sub

Public Sub DeleteMiscSub()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE ITEMID = " + Str(MiscToolID), sqlConn, adOpenKeyset, adLockReadOnly
    If sqlRS.RecordCount > 0 Then
        OldCribID = sqlRS.Fields("CribToolID")
    End If
    sqlRS.Close
    Set sqlRS = Nothing
    ToolChanges(0, ToolChangeCntr) = "REMOVETOOLM"
    ToolChanges(1, ToolChangeCntr) = OldCribID
    ToolChangeCntr = ToolChangeCntr + 1
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    MDIForm1.TabDock.FormHide "Misc Details"
    Set sqlCMD = Nothing
    BuildMiscList
    RefreshReport
End Sub

Public Sub DeleteFixtureSub()
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST FIXTURE] WHERE ITEMID =" + Str(FixtureToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    BuildFixtureList
    RefreshReport
End Sub

Public Sub DeleteRevisionSub()
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST REV] WHERE REVISIONID =" + Str(RevisionID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    BuildRevList
    MDIForm1.TabDock.FormHide "Revision"
    RefreshReport
End Sub

Public Sub AddMiscSub()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[TOOLLIST MISC]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ToolType") = UCase(MiscItem.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = MiscItem.ItemNumberCOMBO.Text
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Fields("CribToolID") = MiscItem.CribNumberIDTXT.Text
    sqlRS.Fields("Consumable") = MiscItem.ConsumableCHECK.Value
    sqlRS.Fields("Manufacturer") = UCase(MiscItem.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = MiscItem.QuantityTXT.Text
    sqlRS.Fields("NumberOfCuttingEdges") = Val(MiscItem.CuttingEdgesTXT.Text)
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(MiscItem.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(MiscItem.AdditionalNotesTXT.Text)
    sqlRS.Fields("ToolbossStock") = MiscItem.TBStock.Value
    OldCribID = ""
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildMiscList
    ToolChanges(0, ToolChangeCntr) = "ADDTOOLM"
    ToolChanges(1, ToolChangeCntr) = MiscItem.CribNumberIDTXT.Text
    ToolChangeCntr = ToolChangeCntr + 1
    RefreshReport
End Sub
Public Sub AddFixtureSub()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[TOOLLIST FIXTURE]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ToolType") = UCase(FixtureItem.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = FixtureItem.ItemNumberCOMBO.Text
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Fields("CribToolID") = FixtureItem.CribNumberIDTXT.Text
    sqlRS.Fields("Manufacturer") = UCase(FixtureItem.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = FixtureItem.QuantityTXT.Text
    sqlRS.Fields("AdditionalNotes") = UCase(FixtureItem.AdditionalNotesTXT.Text)
    sqlRS.Fields("DetailNumber") = FixtureItem.DetailNumberTxt.Text
    sqlRS.Fields("ToolbossStock") = FixtureItem.TBStock.Value
    OldCribID = ""
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildFixtureList
    ToolChanges(0, ToolChangeCntr) = "ADDTOOLF"
    ToolChanges(1, ToolChangeCntr) = FixtureItem.CribNumberIDTXT.Text
    ToolChangeCntr = ToolChangeCntr + 1
    RefreshReport
End Sub

Public Sub AddRevisionSub()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[TOOLLIST REV]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("Revision By") = UCase(RevisionForm.RevByTXT.Text)
    sqlRS.Fields("Revision") = UCase(RevisionForm.RevNumTXT.Text)
    sqlRS.Fields("Revision Description") = UCase(RevisionForm.RevDescTXT.Text)
    sqlRS.Fields("Revision Date") = RevisionForm.RevDate
    sqlRS.Fields("ProcessID") = processId
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildRevList
    RefreshReport
End Sub



Function ValidateItemNumber() As Boolean
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT Manufacturer, ItemClass, [INVENTRY].ItemNumber, Cost FROM [INVENTRY] LEFT OUTER JOIN [ALTVENDOR] ON [INVENTRY].[ALTVENDORNO] = [ALTVENDOR].[RECNUMBER] WHERE DESCRIPTION1 = '" + ItemAttri.ItemNumberCOMBO.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        If Not IsNull(CribRS.Fields("Manufacturer")) Then
            ItemAttri.ManufacturerTXT.Text = CribRS.Fields("Manufacturer")
        End If
        If Not IsNull(CribRS.Fields("ItemClass")) Then
            ItemAttri.ItemGroupTXT.Text = CribRS.Fields("ItemClass")
        End If
        If Not IsNull(CribRS.Fields("ItemNumber")) Then
            ItemAttri.CribNumberIDTXT.Text = CribRS.Fields("ItemNumber")
        End If
        If Not IsNull(CribRS.Fields("Cost")) Then
            ItemAttri.CostTXT.Text = CribRS.Fields("Cost")
        Else
            ItemAttri.CostTXT.Text = "N/A"
            ItemAttri.CostPerPartTXT = "N/A"
        End If
        ValidateItemNumber = True
        GetQty
        CalculateCosts
    Else
        ItemAttri.ItemGroupTXT.Text = ""
        ItemAttri.ManufacturerTXT.Text = ""
        ItemAttri.ItemNumberCOMBO.Text = ""
        ItemAttri.BinTxt = ""
        ItemAttri.QtyOnHandTXT = ""
        ItemAttri.CostTXT = ""
        ItemAttri.CostPerPartTXT = ""
        ItemAttri.MonthlyUsageTXT = ""
        MsgBox ("Invalid Item Number (1)")
        ValidateItemNumber = False
    End If
End Function

Function ValidateMiscItemNumber() As Boolean
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT * FROM [INVENTRY] WHERE DESCRIPTION1= '" + MiscItem.ItemNumberCOMBO.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        If Not IsNull(CribRS.Fields("Manufacturer")) Then
            MiscItem.ManufacturerTXT.Text = CribRS.Fields("Manufacturer")
        End If
        If Not IsNull(CribRS.Fields("ItemClass")) Then
            MiscItem.ItemGroupTXT.Text = CribRS.Fields("ItemClass")
        End If
        If Not IsNull(CribRS.Fields("ItemNumber")) Then
            MiscItem.CribNumberIDTXT.Text = CribRS.Fields("ItemNumber")
        End If
        ValidateMiscItemNumber = True
    Else
        MiscItem.ItemGroupTXT.Text = ""
        MiscItem.ManufacturerTXT.Text = ""
        MiscItem.ItemNumberCOMBO.Text = ""
        MsgBox ("Invalid Item Number")
        ValidateMiscItemNumber = False
    End If
    MiscQty
End Function

Function ValidateFixtureItemNumber() As Boolean
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT * FROM [INVENTRY] WHERE DESCRIPTION1= '" + FixtureItem.ItemNumberCOMBO.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        If Not IsNull(CribRS.Fields("Manufacturer")) Then
            FixtureItem.ManufacturerTXT.Text = CribRS.Fields("Manufacturer")
        End If
        If Not IsNull(CribRS.Fields("ItemClass")) Then
            FixtureItem.ItemGroupTXT.Text = CribRS.Fields("ItemClass")
        End If
        If Not IsNull(CribRS.Fields("ItemNumber")) Then
            FixtureItem.CribNumberIDTXT.Text = CribRS.Fields("ItemNumber")
        End If
        ValidateFixtureItemNumber = True
    Else
        FixtureItem.ItemGroupTXT.Text = ""
        FixtureItem.ManufacturerTXT.Text = ""
        FixtureItem.ItemNumberCOMBO.Text = ""
        MsgBox ("Invalid Item Number")
        ValidateFixtureItemNumber = False
    End If
    FixtureQty
End Function

Public Sub MiscQty()
    Dim sum As Integer
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT ITEM, BIN, QUANTITY FROM STATION WHERE ITEM = '" + MiscItem.CribNumberIDTXT.Text + "' OR ITEM = '" + MiscItem.CribNumberIDTXT.Text + "R'", CribConn, adOpenKeyset, adLockReadOnly
    
    If CribRS.RecordCount > 0 Then
        While Not CribRS.EOF
            sum = sum + CribRS.Fields("quantity")
            CribRS.MoveNext
        Wend
        MiscItem.QtyOnHandTXT.Text = sum
    Else
        MiscItem.QtyOnHandTXT.Text = 0
    End If
    CribRS.Close
    Set CribRS = Nothing
  
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST TOOLBOSS STOCK ITEMS] WHERE ITEMCLASS = '" + MiscItem.ItemGroupTXT.Text + "'", sqlConn, adOpenKeyset
    If SQLRS2.RecordCount > 0 Then
        MiscItem.TBStock.Enabled = False
    Else
        MiscItem.TBStock.Enabled = True
    End If
    SQLRS2.Close
    Set SQLRS2 = Nothing
End Sub

Public Sub FixtureQty()
    Dim sum As Integer
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT ITEM, BIN, QUANTITY FROM STATION WHERE ITEM = '" + FixtureItem.CribNumberIDTXT.Text + "' OR ITEM = '" + FixtureItem.CribNumberIDTXT.Text + "R'", CribConn, adOpenKeyset, adLockReadOnly
    
    If CribRS.RecordCount > 0 Then
        While Not CribRS.EOF
            sum = sum + CribRS.Fields("quantity")
            CribRS.MoveNext
        Wend
        FixtureItem.QtyOnHandTXT.Text = sum
    Else
        FixtureItem.QtyOnHandTXT.Text = 0
    End If
    CribRS.Close
    Set CribRS = Nothing
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST TOOLBOSS STOCK ITEMS] WHERE ITEMCLASS = '" + FixtureItem.ItemGroupTXT.Text + "'", sqlConn, adOpenKeyset
    
    If SQLRS2.RecordCount > 0 Then
        FixtureItem.TBStock.Enabled = False
    Else
        FixtureItem.TBStock.Enabled = True
    End If
    SQLRS2.Close
    Set SQLRS2 = Nothing
End Sub

Public Function CheckForOtherUse(itemNumber As String) As Boolean
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT [TOOLLIST ITEM].TOOLDESCRIPTION,[TOOLLIST MASTER].PROCESSID,[TOOLLIST MASTER].CUSTOMER, [TOOLLIST MASTER].PARTFAMILY FROM [TOOLLIST ITEM] " & _
     "INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST ITEM].PROCESSID = [TOOLLIST MASTER].PROCESSID " & _
     "WHERE [TOOLLIST ITEM].TOOLDESCRIPTION = '" + itemNumber + "' AND [TOOLLIST MASTER].OBSOLETE = 0 ", sqlConn, adOpenKeyset, adLockReadOnly
    If SQLRS2.RecordCount > 0 Then
        CheckForOtherUse = True
    Else
        CheckForOtherUse = False
    End If
    SQLRS2.Close
    Set SQLRS2 = Nothing
End Function

Public Sub GetUsernames()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST USERS]", sqlConn, adOpenKeyset, adLockReadOnly
    If sqlRS.RecordCount > 0 Then
        EmailForm.AdminTXT.Text = sqlRS.Fields("ADMIN")
        EmailForm.BuyerTXT.Text = sqlRS.Fields("BUYER")
        EmailForm.ManagerTXT.Text = sqlRS.Fields("DEPTMGR")
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub UpdateUsernames()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST USERS]", sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("ADMIN") = EmailForm.AdminTXT.Text
        sqlRS.Fields("BUYER") = EmailForm.BuyerTXT.Text
        sqlRS.Fields("DEPTMGR") = EmailForm.ManagerTXT.Text
        sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
    
Public Sub GetSendTo()
    NotificationSendTo = ""
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST EMAIL]", sqlConn, adOpenKeyset, adLockReadOnly
    If sqlRS.RecordCount > 0 Then
        Dim i As Integer
        i = 0
        While i < 6
            If sqlRS.Fields(i) <> "" Then
                NotificationSendTo = sqlRS.Fields(i) + " ," + NotificationSendTo
            End If
            i = i + 1
        Wend
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub CarbonCopyOpen()
    Dim itmx2 As ListItem
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION", sqlConn
    While Not sqlRS.EOF
        Set itmx2 = CarbonCopy.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSID"))
        If Not IsNull(sqlRS.Fields("CUSTOMER")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("CUSTOMER"))
        End If
        If Not IsNull(sqlRS.Fields("PARTFAMILY")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("PARTFAMILY"))
        End If
        If Not IsNull(sqlRS.Fields("OPERATIONDESCRIPTION")) Then
            itmx2.SubItems(3) = Trim(sqlRS.Fields("OPERATIONDESCRIPTION"))
        End If
        If Not IsNull(sqlRS.Fields("OPERATIONNUMBER")) Then
            itmx2.SubItems(4) = Trim(sqlRS.Fields("OPERATIONNUMBER"))
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub PopulateSequence()
    Dim itmx2
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(processId) + " ORDER BY TOOLORDER ", sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        Set itmx2 = ToolAttr.SequenceList.ListItems.Add(, , sqlRS.Fields("ToolOrder"))
        If Not IsNull(sqlRS.Fields("ToolNumber")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("ToolNumber"))
        End If
        If Not IsNull(sqlRS.Fields("OpDescription")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("OpDescription"))
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Function GetNextSequence() As Long
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(processId) + " ORDER BY TOOLORDER ", sqlConn, adOpenKeyset, adLockReadOnly
    If Not sqlRS.EOF Then
        sqlRS.MoveLast
        GetNextSequence = sqlRS.Fields("ToolOrder") + 1
    Else
        GetNextSequence = 1
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Function
Function ReSequenceTools(CurSequence As Long)
    Set sqlRS = New ADODB.Recordset
    sqlRS.CursorLocation = adUseClient
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(processId) + " AND TOOLORDER >= " + Str(CurSequence) + " AND TOOLID <> " + Str(toolID) + " ORDER BY TOOLORDER", sqlConn, adOpenDynamic, adLockOptimistic
    While Not sqlRS.EOF
        CurSequence = CurSequence + 1
        sqlRS.Fields("ToolOrder") = CurSequence
        sqlRS.Update
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Function

Public Sub SetMultiTurret()
    Dim i As Integer
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockOptimistic
    i = 0
    If Not IsNull(sqlRS.Fields("MultiTurret")) Then
        If sqlRS.Fields("MultiTurret") Then
            i = 1
        Else
            i = 0
        End If
        MultiTurret = i
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub CalculateCosts()
    On Error Resume Next
    If ItemAttri.ConsumableCHECK.Value = 0 Then
        Exit Sub
    End If
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(toolID), sqlConn
    If sqlRS.Fields("PartSpecific") = 1 Then
        ItemAttri.MonthlyUsageTXT = Round((ItemAttri.QuantityTXT * (sqlRS.Fields("AdjustedVolume") / 12)) / (ItemAttri.ToolLifeTXT * ItemAttri.CuttingEdgesTXT), 3)
    Else
        sqlRS.Close
        sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn
        ItemAttri.MonthlyUsageTXT.Text = Round((Val(ItemAttri.QuantityTXT.Text) * (sqlRS.Fields("AnnualVolume") / 12)) / (ItemAttri.ToolLifeTXT * ItemAttri.CuttingEdgesTXT), 3)
    End If
    
    sqlRS.Close
    Set sqlRS = Nothing
    
    If ItemAttri.ConsumableCHECK.Value = 1 And ItemAttri.RegrindableChk.Value = 0 Then
        If Val(ItemAttri.ToolLifeTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.QuantityTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.CuttingEdgesTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.CostTXT.Text) = 0 Or ItemAttri.CostTXT.Text = "N/A" Then
            Exit Sub
        End If
        ItemAttri.CostPerPartTXT.Text = Round((ItemAttri.CostTXT * ItemAttri.QuantityTXT) / (ItemAttri.ToolLifeTXT * ItemAttri.CuttingEdgesTXT), 3)
    ElseIf ItemAttri.ConsumableCHECK.Value = 1 And ItemAttri.RegrindableChk.Value = 1 Then
        If Val(ItemAttri.ToolLifeTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.QuantityTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.CuttingEdgesTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.CostTXT) = 0 Or ItemAttri.CostTXT.Text = "N/A" Then
            Exit Sub
        End If
        If Val(ItemAttri.NumofRegrindsTXT) = 0 Then
            Exit Sub
        End If
        If Val(ItemAttri.ToolLifeRegrindTXT) = 0 Then
            Exit Sub
        End If
        ItemAttri.CostPerPartTXT.Text = Round(((ItemAttri.CostTXT + (ItemAttri.NumofRegrindsTXT * 25)) * ItemAttri.QuantityTXT) / (((ItemAttri.NumofRegrindsTXT * ItemAttri.ToolLifeRegrindTXT) + ItemAttri.ToolLifeTXT) * ItemAttri.CuttingEdgesTXT), 3)
    Else
        Exit Sub
    End If
End Sub

Public Sub PopulateChangesForRouting()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn
    Dim originalPlantText As String
    Dim j As Integer
    Dim lsi
'    CreateRouting.ToolingChangeList.ListItems.Clear
    
    While j < 10
        If OriginalPlant(j) <> 0 Then
            originalPlantText = originalPlantText + Trim(Str(OriginalPlant(j))) + ", "
        End If
        j = j + 1
    Wend
    Dim i As Integer
    Dim itmx2
    i = 0
    CreateRouting.UsernameLBL.Caption = Environ("USERNAME")
    CreateRouting.DateLBL.Caption = Date
    CreateRouting.ToolListLBL.Caption = Str(processId) + " - " + sqlRS.Fields("CUSTOMER") + " - " + sqlRS.Fields("PartFamily") + " - " + sqlRS.Fields("OperationDescription")
    While i < 200
        Select Case ToolChanges(0, i)
        Case "STATUS"
            CreateRouting.StatusChangeList.ListItems.Clear
            If sqlRS.Fields("RELEASED") <> OriginalReleased Then
                If sqlRS.Fields("RELEASED") = True Then
                    Set itmx2 = CreateRouting.StatusChangeList.ListItems.Add(, , "RELEASED")
                Else
                    Set itmx2 = CreateRouting.StatusChangeList.ListItems.Add(, , "UNRELEASED")
                End If
                itmx2.ListSubItems.Add , , "", 1
                itmx2.ListSubItems.Add , , "", 1
            End If
            If sqlRS.Fields("OBSOLETE") <> OriginalObsolete Then
                If sqlRS.Fields("OBSOLETE") = True Then
                    Set itmx2 = CreateRouting.StatusChangeList.ListItems.Add(, , "OBSOLETE")
                Else
                    Set itmx2 = CreateRouting.StatusChangeList.ListItems.Add(, , "ACTIVE")
                End If
                itmx2.ListSubItems.Add , , "", 1
                itmx2.ListSubItems.Add , , "", 1
            End If
        Case "PLANT"
            If CreateRouting.PlantChangeList.ListItems.Count = 0 Then
                Set itmx2 = CreateRouting.PlantChangeList.ListItems.Add(, , ToolChanges(1, i))
                itmx2.SubItems(1) = originalPlantText
            Else
                CreateRouting.PlantChangeList.ListItems.Item(1).Text = ToolChanges(1, i)
                CreateRouting.PlantChangeList.ListItems.Item(1).SubItems(1) = originalPlantText
            End If
        Case "VOLUME"
            If CreateRouting.VolumeChangeList.ListItems.Count = 0 Then
                Set itmx2 = CreateRouting.VolumeChangeList.ListItems.Add(, , ToolChanges(1, i))
                itmx2.SubItems(1) = OriginalVolume
            Else
                If Val(ToolChanges(1, i)) = OriginalVolume Then
                    CreateRouting.VolumeChangeList.ListItems.Remove (1)
                Else
                    CreateRouting.VolumeChangeList.ListItems.Item(1).Text = ToolChanges(1, i)
                    CreateRouting.VolumeChangeList.ListItems.Item(1).SubItems(1) = OriginalVolume
                End If
            End If
        Case "ADDTOOL"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "ADDED"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "ADDTOOLM"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "ADDEDM"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "ADDTOOLF"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "ADDEDF"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "REMOVETOOL"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "REMOVED"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "REMOVETOOLM"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "REMOVEDM"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "REMOVETOOLF"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "REMOVEDF"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "USAGE"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "USAGE CHANGE"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "STOCK"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "STOCK TOOLBOSS"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
        Case "PICTURES"
            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(1, i))
'            Set itmx2 = CreateRouting.ToolingChangeList.ListItems.Add(, , ToolChanges(2, i))
            itmx2.SubItems(1) = GetItemDescription(ToolChanges(1, i))
            itmx2.SubItems(2) = "PICTURE CHANGE"
            itmx2.ListSubItems.Add , , "", 1
            itmx2.ListSubItems.Add , , "", 1
            itmx2.SubItems(6) = ToolChanges(2, i)
'            itmx2.SubItems.Add , , ToolChanges(2, i), 1
        End Select
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing

    ConsolidateChanges

    Erase ToolChanges
    ToolChangeCntr = 0

End Sub
Public Sub ConsolidateChanges()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim IsStillInToolList As Boolean
    Dim InOriginalToolList As Boolean
    Dim IsDuplicate As Boolean
    j = 0
    i = 1
    k = 1
    While i <= CreateRouting.ToolingChangeList.ListItems.Count
        Select Case CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2)
        Case "REMOVED", "REMOVEDM", "REMOVEDF"
            j = 0
            'CHECK IF REMOVED TOOL WAS EVER IN THE PROCESS
            InOriginalToolList = False
            Do While j <= 400 'Max tools in a tool list
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = OriginalTools(j) Then
                    InOriginalToolList = True
                    Exit Do 'No need to continue
                End If
                j = j + 1
            Loop
            'CHECK IF REMOVED TOOL IS A DUPLICATE REMOVAL
            IsDuplicate = False
            k = 1
            Do While k <= CreateRouting.ToolingChangeList.ListItems.Count
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = CreateRouting.ToolingChangeList.ListItems.Item(k).Text _
                And i <> k And CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2) = _
                CreateRouting.ToolingChangeList.ListItems.Item(k).SubItems(2) Then
                    IsDuplicate = True
                    Exit Do 'No need to continue
                End If
                k = k + 1
            Loop
            'CHECK IF REMOVED TOOL STILL EXISTS ELSEWHERE IN THE PROCESS
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            IsStillInToolList = False
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
           
            'Do we really need this ToolingChangeList item
            If InOriginalToolList And Not IsDuplicate And Not IsStillInToolList Then
                ' This item can stay in the ToolingChangeList
                i = i + 1
            Else
                'This item is either not in the Original tool list, still in the new tool list,
                'or is duplicated in the Toolling Change List so we can remove it
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
            End If
        Case "ADDED"
            j = 0
            'CHECK IF ADDED TOOL ALREADY EXISTED IN THE PROCESS
            InOriginalToolList = False
            While j <= 400
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = OriginalTools(j) Then
                    InOriginalToolList = True
                End If
                j = j + 1
            Wend
            'CHECK IF ADDED TOOL IS A DUPLICATE
            IsDuplicate = False
            k = 1
            While k <= CreateRouting.ToolingChangeList.ListItems.Count
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = CreateRouting.ToolingChangeList.ListItems.Item(k).Text And i <> k And CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2) = CreateRouting.ToolingChangeList.ListItems.Item(k).SubItems(2) Then
                    IsDuplicate = True
                End If
                k = k + 1
            Wend
            'CHECK IF ADDED TOOL STILL EXISTS ELSEWHERE IN THE PROCESS
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            IsStillInToolList = False
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            If Not InOriginalToolList And IsStillInToolList And Not IsDuplicate Then
                i = i + 1
            Else
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
            End If
        Case "ADDEDM"
            j = 0
            'CHECK IF ADDED TOOL ALREADY EXISTED IN THE PROCESS
            InOriginalToolList = False
            While j <= 100
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = OriginalTools(j) Then
                    InOriginalToolList = True
                End If
                j = j + 1
            Wend
            'CHECK IF ADDED TOOL IS A DUPLICATE
            IsDuplicate = False
            k = 1
            While k <= CreateRouting.ToolingChangeList.ListItems.Count
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = CreateRouting.ToolingChangeList.ListItems.Item(k).Text And i <> k And CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2) = CreateRouting.ToolingChangeList.ListItems.Item(k).SubItems(2) Then
                    IsDuplicate = True
                End If
                k = k + 1
            Wend
            'CHECK IF ADDED TOOL STILL EXISTS ELSEWHERE IN THE PROCESS
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            IsStillInToolList = False
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            If Not InOriginalToolList And IsStillInToolList And Not IsDuplicate Then
                i = i + 1
            Else
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
            End If
        Case "ADDEDF"
            j = 0
            'CHECK IF ADDED TOOL ALREADY EXISTED IN THE PROCESS
            InOriginalToolList = False
            While j <= 100
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = OriginalTools(j) Then
                    InOriginalToolList = True
                End If
                j = j + 1
            Wend
            'CHECK IF ADDED TOOL IS A DUPLICATE
            IsDuplicate = False
            k = 1
            While k <= CreateRouting.ToolingChangeList.ListItems.Count
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = CreateRouting.ToolingChangeList.ListItems.Item(k).Text And i <> k And CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2) = CreateRouting.ToolingChangeList.ListItems.Item(k).SubItems(2) Then
                    IsDuplicate = True
                End If
                k = k + 1
            Wend
            'CHECK IF ADDED TOOL STILL EXISTS ELSEWHERE IN THE PROCESS
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            IsStillInToolList = False
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            If Not InOriginalToolList And IsStillInToolList And Not IsDuplicate Then
                i = i + 1
            Else
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
            End If
        Case "USAGE CHANGE"
            j = 0
            'CHECK IF ADDED TOOL IS A DUPLICATE
            IsDuplicate = False
            k = 1
            
            While k <= CreateRouting.ToolingChangeList.ListItems.Count
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = CreateRouting.ToolingChangeList.ListItems.Item(k).Text _
                And CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2) = CreateRouting.ToolingChangeList.ListItems.Item(k).SubItems(2) _
                And i <> k Then
                    IsDuplicate = True
                End If
                k = k + 1
            Wend
            
            'CHECK IF USAGE TOOL STILL EXISTS IN THE PROCESS
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
            IsStillInToolList = False
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            If IsStillInToolList And Not IsDuplicate Then
                i = i + 1
            Else
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
            End If
        Case "STOCK TOOLBOSS"
            j = 0
            'CHECK IF ADDED TOOL IS A DUPLICATE
            IsDuplicate = False
            k = 1
            While k <= CreateRouting.ToolingChangeList.ListItems.Count
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = CreateRouting.ToolingChangeList.ListItems.Item(k).Text And i <> k And CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2) = CreateRouting.ToolingChangeList.ListItems.Item(k).SubItems(2) Then
                    IsDuplicate = True
                End If
                k = k + 1
            Wend
            'CHECK IF TOOL IS STILL MARKED FOR STOCKING AND IS STILL IN TOOL LIST(COULD OF BEEN DELETED AFTER THE THE STOCK WAS CHECKED)
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId) + " AND TOOLBOSSSTOCK = 1", sqlConn, adOpenKeyset, adLockReadOnly
            IsStillInToolList = False
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(processId) + " AND TOOLBOSSSTOCK = 1", sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(processId) + " AND TOOLBOSSSTOCK = 1", sqlConn, adOpenKeyset, adLockReadOnly
            While Not sqlRS.EOF
                If CreateRouting.ToolingChangeList.ListItems.Item(i).Text = sqlRS.Fields("CRIBTOOLID") Then
                    IsStillInToolList = True
                End If
                sqlRS.MoveNext
            Wend
            sqlRS.Close
            If IsStillInToolList And Not IsDuplicate Then
                i = i + 1
            Else
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
                
            End If
        Case "PICTURE CHANGE"
            Dim strItemId As String
            strItemId = CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(6)
'            CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(3) = ""
            Set sqlRS = New ADODB.Recordset
            sqlRS.Open "SELECT ItemImage FROM [TOOLLIST ITEM] WHERE ItemId =" + strItemId, sqlConn, adOpenKeyset, adLockReadOnly
        
            ' If there was no picture originally and none has been added then remove this tool change
            If colItemImages.Item(strItemId) = "F" And IsNull(sqlRS.Fields("ItemImage")) Then
                CreateRouting.ToolingChangeList.ListItems.Remove (i)
            Else
                 i = i + 1
            End If
            sqlRS.Close
            Set sqlRS = Nothing
        Case Else
            i = i + 1
        End Select
    Wend
    
End Sub

Public Sub ClearKitFields()
    KitAttri.ItemNumberCOMBO.Text = ""
    If KitAttri.ItemNumberCOMBO.ListCount = 0 Then
        PopulateKitList
    End If
End Sub
Public Sub PopulateKitList()
    KitAttri.ItemNumberCOMBO.Clear
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT DISTINCT DESCRIPTION1 FROM [INVENTRY] WHERE DESCRIPTION1 is not NULL AND ITEMTYPE = 4 ORDER BY DESCRIPTION1", CribConn
    While Not CribRS.EOF
        KitAttri.ItemNumberCOMBO.AddItem CribRS.Fields("DESCRIPTION1")
        CribRS.MoveNext
    Wend
    CribRS.Close
    Set CribRS = Nothing
End Sub

Public Sub AddKit()
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT KIT.ITEMNUMBER, KIT.QUANTITY, INVENTRY.DESCRIPTION1, INVENTRY.MANUFACTURER, INVENTRY.ITEMCLASS FROM KIT INNER JOIN [INVENTRY] ON [KIT].ITEMNUMBER = [INVENTRY].ITEMNUMBER WHERE KITNUMBER = '" + KitAttri.CribNumberIDTXT.Text + "'", CribConn
    While Not CribRS.EOF
        sqlRS.AddNew
        sqlRS.Fields("ToolType") = UCase(CribRS.Fields("ITEMCLASS"))
        sqlRS.Fields("ToolDescription") = UCase(CribRS.Fields("DESCRIPTION1"))
        sqlRS.Fields("ProcessID") = processId
        sqlRS.Fields("ToolID") = toolID
        sqlRS.Fields("CribToolID") = UCase(CribRS.Fields("ITEMNUMBER"))
        sqlRS.Fields("Consumable") = 0
        sqlRS.Fields("Manufacturer") = UCase(CribRS.Fields("MANUFACTURER"))
        sqlRS.Fields("Quantity") = UCase(CribRS.Fields("QUANTITY"))
        sqlRS.Fields("NumberOfCuttingEdges") = 0
        sqlRS.Fields("QuantityPerCuttingEdge") = 0
        sqlRS.Fields("AdditionalNotes") = ""
        sqlRS.Fields("NumOfRegrinds") = 0
        sqlRS.Fields("QtyPerRegrind") = 0
        sqlRS.Fields("Regrindable") = 0
        sqlRS.Fields("ToolbossStock") = 0
        sqlRS.Update
        ToolChanges(0, ToolChangeCntr) = "ADDTOOL"
        ToolChanges(1, ToolChangeCntr) = CribRS.Fields("ItemNumber")
        ToolChangeCntr = ToolChangeCntr + 1
        CribRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    CribRS.Close
    Set CribRS = Nothing
    BuildToolList
    OldCribID = ""
End Sub

Function ValidateKitNumber() As Boolean
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT INVENTRY.ItemNumber FROM INVENTRY WHERE DESCRIPTION1 = '" + KitAttri.ItemNumberCOMBO.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
    If CribRS.RecordCount > 0 Then
        If Not IsNull(CribRS.Fields("ItemNumber")) Then
            KitAttri.CribNumberIDTXT.Text = CribRS.Fields("ItemNumber")
        End If
        ValidateKitNumber = True

    Else
        KitAttri.ItemNumberCOMBO.Text = ""
        MsgBox ("Invalid Kit Number")
        ValidateKitNumber = False
    End If
End Function

Public Sub PopulateOriginalTools()
    Erase OriginalTools
    Dim i As Integer
    i = 0
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
    
    While Not sqlRS.EOF
        OriginalTools(i) = sqlRS.Fields("CRIBTOOLID")
        i = i + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        OriginalTools(i) = sqlRS.Fields("CRIBTOOLID")
        i = i + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    sqlRS.Open "SELECT CRIBTOOLID FROM [TOOLLIST FIXTURE] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        OriginalTools(i) = sqlRS.Fields("CRIBTOOLID")
        i = i + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT ANNUALVOLUME, RELEASED, OBSOLETE FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
    OriginalVolume = sqlRS.Fields("ANNUALVOLUME")
    OriginalReleased = sqlRS.Fields("RELEASED")
    OriginalObsolete = sqlRS.Fields("OBSOLETE")
    sqlRS.Close
    
    i = 0
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [ToolList PLANT] WHERE PROCESSID =" + Str(processId) + " ORDER BY PLANT", sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        OriginalPlant(i) = sqlRS.Fields("Plant")
        i = i + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub ClearOriginalTools()
    Erase OriginalTools
    OriginalVolume = 0
    OriginalReleased = False
    OriginalObsolete = False
    Erase OriginalPlant
    Erase PlantChange
    CreateRouting.ReasonTxt.Text = ""
End Sub
Public Sub ClearRoutingForm()
    CreateRouting.Hide
    CreateRouting.ToolListLBL = ""
    CreateRouting.UsernameLBL = ""
    CreateRouting.DateLBL = ""
    CreateRouting.ReasonTxt = ""
    CreateRouting.ToolingChangeList.ListItems.Clear
    CreateRouting.VolumeChangeList.ListItems.Clear
    CreateRouting.PlantChangeList.ListItems.Clear
    CreateRouting.StatusChangeList.ListItems.Clear
End Sub

Public Sub Reset()
    DoEvents
    ClearItemFields
    ClearKitFields
    DoEvents
    ClearMiscFields
    ClearRevisionFields
    DoEvents
    ClearProcessFields
    ClearToolFields
    DoEvents
    ClearOriginalTools
    ClearRoutingForm
    DoEvents
    processId = 0
    OldProcessID = 0
    toolID = 0
    itemID = 0
    MiscToolID = 0
    RevisionID = 0
    bRefreshActionListError = False
    DoEvents
    Set sqlRS = Nothing
    Set SQLRS2 = Nothing
    Set sqlCMD = Nothing
    Set CribRS = Nothing
    toolexists = False
    itemexists = False
    misctoolexists = False
    fixturetoolexists = False
    revisionexists = False
    processexists = False
    DoEvents
    OldCribID = ""
    LastToolModified = ""
    Erase ToolChanges
    ToolChangeCntr = 0
    LastToolDescription = ""
    MultiTurret = False
    DoEvents
    openSQLStatement = ""
    CreateRouting.ClearFields
    ClearOriginalTools
    WorkingLive = False
    DoEvents
End Sub

Function GetUserType() As String
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST USERS]", sqlConn, adOpenKeyset, adLockReadOnly
    If Not sqlRS.EOF Then
    'Debug: Comment out entire Select Case statement for debug and uncomment desired GetUserType and MDIForm1.DeleteToolList.Visible assignment
    Select Case LCase(Environ("USERNAME"))
      Case sqlRS.Fields("ADMIN")
            GetUserType = "ADMIN"
            MDIForm1.DeleteToolList.Visible = True
        Case sqlRS.Fields("BUYER")
            GetUserType = "BUYER"
            MDIForm1.DeleteToolList.Visible = False
        Case sqlRS.Fields("DEPTMGR")
            GetUserType = "MANAGER"
            MDIForm1.DeleteToolList.Visible = False
        Case Else
'            GetUserType = "ENGINEER"  TAKE ME OUT FOR PRODUCTION
            GetUserType = "BUYER"
            MDIForm1.DeleteToolList.Visible = False
       End Select
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Function

Public Sub WriteRouting()
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    DoEvents
    EmailMessage = ""
    Dim ProcessChangeID As Long
    Dim OldProcessID As Long
    If CreateRouting.ToolingChangeList.ListItems.Count > 0 Or CreateRouting.PlantChangeList.ListItems.Count > 0 Or CreateRouting.VolumeChangeList.ListItems.Count > 0 Or CreateRouting.StatusChangeList.ListItems.Count > 0 Then
        CreateRouting.Hide
        If (True = reportOpened) Then
            ReportForm.Hide
        End If
        Set sqlRS = New ADODB.Recordset
        sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Trim(Str(processId)), sqlConn, adOpenKeyset, adLockReadOnly
        OldProcessID = sqlRS.Fields("RevOfProcessID")
        Set sqlRS = New ADODB.Recordset
        sqlRS.CursorLocation = adUseClient
        sqlRS.Open "[TOOLLIST CHANGE MASTER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        sqlRS.AddNew
        sqlRS.Fields("PROCESSID") = processId
        sqlRS.Fields("COMPLETE") = False
        sqlRS.Fields("COMMENTS") = Trim(CreateRouting.ReasonTxt.Text)
        EmailMessage = vbCrLf + vbCrLf + "Reason For Change: " + Trim(CreateRouting.ReasonTxt.Text) + vbCrLf + vbCrLf
        sqlRS.Fields("ENGINEER") = Trim(Environ("USERNAME"))
        sqlRS.Fields("DATEINITIATED") = Date
        sqlRS.Fields("DATECOMPLETE") = #1/1/1900#
        sqlRS.Fields("APPROVED") = 0
        sqlRS.Fields("INITIALRELEASE") = 0
        sqlRS.Fields("OldProcessID") = OldProcessID
        sqlRS.Update
        sqlRS.Close
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] ORDER BY PROCESSCHANGEID DESC", sqlConn, adOpenKeyset, adLockReadOnly
        ProcessChangeID = sqlRS.Fields("ProcessChangeID")
        sqlRS.Close
        Dim i As Integer
        i = 1
        Set sqlRS = New ADODB.Recordset
        sqlRS.CursorLocation = adUseClient
        sqlRS.Open "[TOOLLIST CHANGE ITEMS]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        While i <= CreateRouting.ToolingChangeList.ListItems.Count
            sqlRS.AddNew
            sqlRS.Fields("ProcessChangeID") = ProcessChangeID
            sqlRS.Fields("Type") = Trim(CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2))
            sqlRS.Fields("CribmasterID") = Trim(CreateRouting.ToolingChangeList.ListItems.Item(i).Text)
            sqlRS.Fields("NewStatus") = ""
            sqlRS.Fields("NewPlants") = ""
            sqlRS.Fields("OldPlants") = ""
            sqlRS.Fields("NewVolume") = 0
            sqlRS.Fields("OldVolume") = 0
            sqlRS.Fields("DispositionMethod") = ""
            sqlRS.Fields("Comments") = Trim(CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(5))
            sqlRS.Fields("Completed") = 0
            sqlRS.Fields("APPROVED") = 0
            sqlRS.Update
            EmailMessage = EmailMessage + vbCrLf + Trim(CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(2)) + " - " + Trim(CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(1)) + " ---- " + Trim(CreateRouting.ToolingChangeList.ListItems.Item(i).SubItems(5))
            i = i + 1
        Wend
        i = 1
        While i <= CreateRouting.StatusChangeList.ListItems.Count
            sqlRS.AddNew
            sqlRS.Fields("ProcessChangeID") = ProcessChangeID
            sqlRS.Fields("Type") = "STATUS"
            sqlRS.Fields("CribmasterID") = ""
            sqlRS.Fields("NewStatus") = Trim(CreateRouting.StatusChangeList.ListItems.Item(i).Text)
            sqlRS.Fields("NewPlants") = ""
            sqlRS.Fields("OldPlants") = ""
            sqlRS.Fields("NewVolume") = 0
            sqlRS.Fields("OldVolume") = 0
            sqlRS.Fields("DispositionMethod") = ""
            sqlRS.Fields("Comments") = Trim(CreateRouting.StatusChangeList.ListItems.Item(i).SubItems(5))
            sqlRS.Fields("Completed") = 0
            sqlRS.Fields("APPROVED") = 0
            sqlRS.Update
            EmailMessage = EmailMessage + vbCrLf + "CHANGING STATUS TO: " + Trim(CreateRouting.StatusChangeList.ListItems.Item(i).Text) + " ---- " + Trim(CreateRouting.StatusChangeList.ListItems.Item(i).SubItems(5))
            i = i + 1
        Wend
        i = 1
        While i <= CreateRouting.PlantChangeList.ListItems.Count
            sqlRS.AddNew
            sqlRS.Fields("ProcessChangeID") = ProcessChangeID
            sqlRS.Fields("Type") = "PLANT"
            sqlRS.Fields("CribmasterID") = ""
            sqlRS.Fields("NewStatus") = ""
            sqlRS.Fields("NewPlants") = Trim(CreateRouting.PlantChangeList.ListItems.Item(i).Text)
            sqlRS.Fields("OldPlants") = Trim(CreateRouting.PlantChangeList.ListItems.Item(i).SubItems(1))
            sqlRS.Fields("NewVolume") = 0
            sqlRS.Fields("OldVolume") = 0
            sqlRS.Fields("DispositionMethod") = ""
            sqlRS.Fields("Comments") = Trim(CreateRouting.PlantChangeList.ListItems.Item(i).SubItems(5))
            sqlRS.Fields("Completed") = 0
            sqlRS.Fields("APPROVED") = 0
            sqlRS.Update
            EmailMessage = EmailMessage + vbCrLf + "OLDPLANTS: " + Trim(CreateRouting.PlantChangeList.ListItems.Item(i).SubItems(1)) + " ---- " + Trim(CreateRouting.PlantChangeList.ListItems.Item(i).SubItems(5))
            EmailMessage = EmailMessage + vbCrLf + "NEWPLANTS: " + Trim(CreateRouting.PlantChangeList.ListItems.Item(i).Text)
            i = i + 1
        Wend
        i = 1
        While i <= CreateRouting.VolumeChangeList.ListItems.Count
            sqlRS.AddNew
            sqlRS.Fields("ProcessChangeID") = ProcessChangeID
            sqlRS.Fields("Type") = "VOLUME"
            sqlRS.Fields("CribmasterID") = ""
            sqlRS.Fields("NewStatus") = ""
            sqlRS.Fields("NewPlants") = ""
            sqlRS.Fields("OldPlants") = ""
            sqlRS.Fields("NewVolume") = Trim(CreateRouting.VolumeChangeList.ListItems.Item(i).Text)
            sqlRS.Fields("OldVolume") = Trim(CreateRouting.VolumeChangeList.ListItems.Item(i).SubItems(1))
            sqlRS.Fields("DispositionMethod") = ""
            sqlRS.Fields("Comments") = Trim(CreateRouting.VolumeChangeList.ListItems.Item(i).SubItems(5))
            sqlRS.Fields("Completed") = 0
            sqlRS.Fields("APPROVED") = 0
            sqlRS.Update
            EmailMessage = EmailMessage + vbCrLf + "OLDVOLUME: " + Trim(CreateRouting.VolumeChangeList.ListItems.Item(i).SubItems(1)) + " ---- " + Trim(CreateRouting.VolumeChangeList.ListItems.Item(i).SubItems(5))
            EmailMessage = EmailMessage + vbCrLf + "NEWVOLUME: " + Trim(CreateRouting.VolumeChangeList.ListItems.Item(i).Text)
            i = i + 1
        Wend
        sqlRS.Close
        sqlRS.Open "SELECT * FROM [ToolList Change Items] WHERE PROCESSCHANGEID = '" + Str(ProcessChangeID) + "'", sqlConn, adOpenKeyset, adLockReadOnly
        Set SQLRS2 = New ADODB.Recordset
        SQLRS2.Open "[TOOLLIST CHANGE ACTIONS]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        
        Set SQLRS3 = New ADODB.Recordset
        SQLRS3.Open "SELECT * FROM [ToolList Toolboss Stock Items]", sqlConn, adOpenKeyset, adLockReadOnly
        
        Set CribRS = New ADODB.Recordset
'        CribRS.Open "SELECT ITEMNUMBER,ITEMCLASS,CRIBBIN FROM INVENTRY LEFT OUTER JOIN STATION ON INVENTRY.ITEMNUMBER = STATION.ITEM", CribConn, adOpenKeyset, adLockReadOnly
        
        Set SQLRS4 = New ADODB.Recordset
        
        While Not sqlRS.EOF
            Select Case Trim(sqlRS.Fields("Type"))
                Case "ADDED"
                    CribRS.Open "SELECT ITEMNUMBER,ITEMCLASS,CRIBBIN FROM INVENTRY " & _
                        "LEFT OUTER JOIN STATION ON INVENTRY.ITEMNUMBER = STATION.ITEM " & _
                        "Where ItemNumber = '" + sqlRS.Fields("CRIBMASTERID") + "'", CribConn, adOpenKeyset, adLockReadOnly
                    SQLRS4.Open "SELECT TOOLBOSSSTOCK FROM [TOOLLIST ITEM] WHERE TOOLBOSSSTOCK = 1 AND PROCESSID = " + Trim(Str(processId)), sqlConn, adOpenKeyset
                    If Not CribRS.EOF Then
                        SQLRS3.MoveFirst
                        SQLRS3.Find ("ITEMCLASS LIKE '" + CribRS.Fields("ITEMCLASS") + "'")
                        If Not SQLRS3.EOF Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 3
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        End If
                        If IsNull(CribRS.Fields("CRIBBIN")) Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 1
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        Else
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 2
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        End If
                        If SQLRS4.RecordCount > 0 Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 14
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        End If
                        SQLRS4.Close
                    End If
                    CribRS.Close
                    
                Case "REMOVED"
                    CribRS.Open "SELECT ITEMNUMBER,ITEMCLASS,CRIBBIN FROM INVENTRY " & _
                        "LEFT OUTER JOIN STATION ON INVENTRY.ITEMNUMBER = STATION.ITEM " & _
                        "Where ItemNumber = '" + sqlRS.Fields("CRIBMASTERID") + "'", CribConn, adOpenKeyset, adLockReadOnly
                    If Not CribRS.EOF Then
                        SQLRS3.MoveFirst
                        SQLRS3.Find ("ITEMCLASS LIKE '" + CribRS.Fields("ITEMCLASS") + "'")
                        SQLRS4.Open "SELECT * FROM [TOOLLIST ITEM] INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST ITEM].PROCESSID = [TOOLLIST MASTER].PROCESSID WHERE [TOOLLIST MASTER].OBSOLETE = 0 AND [TOOLLIST ITEM].CRIBTOOLID = '" + sqlRS.Fields("CRIBMASTERID") + "' AND [TOOLLIST MASTER].PROCESSID <> " + Str(OldProcessID), sqlConn, adOpenKeyset, adLockReadOnly
                        If SQLRS4.EOF And SQLRS3.EOF Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 6
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        ElseIf SQLRS4.EOF And Not SQLRS3.EOF Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 11
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 6
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        ElseIf Not SQLRS3.EOF Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 4
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        End If
                        SQLRS4.Close
                    End If
                    CribRS.Close
                Case "USAGE CHANGE"
                    CribRS.Open "SELECT ITEMNUMBER,ITEMCLASS,CRIBBIN FROM INVENTRY " & _
                        "LEFT OUTER JOIN STATION ON INVENTRY.ITEMNUMBER = STATION.ITEM " & _
                        "Where ItemNumber = '" + sqlRS.Fields("CRIBMASTERID") + "'", CribConn, adOpenKeyset, adLockReadOnly
                    If Not CribRS.EOF Then
                        SQLRS3.Find ("ITEMCLASS LIKE '" + CribRS.Fields("ITEMCLASS") + "'")
                        If SQLRS3.EOF Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 2
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        Else
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 12
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 2
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        End If
                        If IsNull(CribRS.Fields("CRIBBIN")) Then
                            SQLRS2.AddNew
                            SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                            SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                            SQLRS2.Fields("ACTIONITEM") = 1
                            SQLRS2.Fields("COMPLETE") = 0
                            SQLRS2.Update
                        End If
                    End If
                    CribRS.Close
                
                Case "STATUS"
                    If Trim(sqlRS.Fields("NEWSTATUS")) = "RELEASED" Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 10
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 9
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    ElseIf Trim(sqlRS.Fields("NEWSTATUS")) = "OBSOLETE" Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 8
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 13
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    ElseIf Trim(sqlRS.Fields("NEWSTATUS")) = "ACTIVE" Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 2
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 9
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 10
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                Case "PLANT"
                    SQLRS2.AddNew
                    SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                    SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                    SQLRS2.Fields("ACTIONITEM") = 7
                    SQLRS2.Fields("COMPLETE") = 0
                    SQLRS2.Update
                Case "VOLUME"
                    SQLRS2.AddNew
                    SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                    SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                    SQLRS2.Fields("ACTIONITEM") = 5
                    SQLRS2.Fields("COMPLETE") = 0
                    SQLRS2.Update
            End Select
            sqlRS.MoveNext
        Wend
        sqlRS.Close
        SQLRS2.Close
        SQLRS3.Close
        sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(processId), sqlConn, adOpenKeyset, adLockReadOnly
        SQLRS2.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(sqlRS.Fields("REVOFPROCESSID")), sqlConn, adOpenKeyset, adLockOptimistic
        SQLRS2.Fields("REVINPROCESS") = 1
        SQLRS2.Update
        sqlRS.Close
        SQLRS2.Close
        SendNeedApprovalNotification (ProcessChangeID)
        
        Reset
        ExitLoop = True
        MDIForm1.RefreshMenuOptions
        If (True = reportOpened) Then
            ReportForm.Hide
        End If
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
        MsgBox ("ROUTING SENT ON FOR APPROVAL")
    Else
        ProgressBar.Hide
        ProgressBar.Timer1.Enabled = False
        MsgBox ("No changes have been made")
    End If
End Sub


Public Function IsReadyToExit() As Boolean
    If Not WorkingLive And (0 <> processId) Then
        If CreateRouting.ToolingChangeList.ListItems.Count > 0 Or CreateRouting.StatusChangeList.ListItems.Count > 0 Or CreateRouting.ToolingChangeList.ListItems.Count > 0 Or CreateRouting.PlantChangeList.ListItems.Count > 0 Or ToolChanges(0, 0) <> "" Then
            IsReadyToExit = False
        Else
            IsReadyToExit = True
            DeleteProcessSub (processId)
        End If
    Else
        IsReadyToExit = True
    End If
End Function


Public Sub PopulateRouting(ProcessChangeID As Long)
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT [TOOLLIST CHANGE MASTER].PROCESSID, [TOOLLIST CHANGE MASTER].OLDPROCESSID, COMMENTS, ENGINEER, DATEINITIATED, CUSTOMER, PARTFAMILY, OPERATIONDESCRIPTION FROM [TOOLLIST CHANGE MASTER] INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST CHANGE MASTER].PROCESSID = [TOOLLIST MASTER].PROCESSID WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    CreateRouting.SetProcessIDs sqlRS.Fields("PROCESSID"), sqlRS.Fields("OLDPROCESSID")
    CreateRouting.ReasonTxt = sqlRS.Fields("Comments")
    CreateRouting.UsernameLBL = sqlRS.Fields("Engineer")
    CreateRouting.ToolListLBL = sqlRS.Fields("CUSTOMER") + " - " + sqlRS.Fields("PARTFAMILY") + " - " + sqlRS.Fields("OPERATIONDESCRIPTION")
    CreateRouting.DateLBL = sqlRS.Fields("DATEINITIATED")
    sqlRS.Close
    Dim itmx
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        Select Case Trim(sqlRS.Fields("TYPE"))
        Case "ADDED"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "STOCK TOOLBOSS"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "ADDEDM"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "ADDEDF"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "REMOVED"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "REMOVEDM"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "REMOVEDF"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "USAGE CHANGE"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "STATUS"
            Set itmx = CreateRouting.StatusChangeList.ListItems.Add(, , sqlRS.Fields("NEWSTATUS"))
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "PLANT"
            'sqlRS.AddNew
            Set itmx = CreateRouting.PlantChangeList.ListItems.Add(, , sqlRS.Fields("NEWPLANTS"))
            itmx.SubItems(1) = sqlRS.Fields("OldPlants")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
            'sqlRS.Update
        Case "VOLUME"
            Set itmx = CreateRouting.VolumeChangeList.ListItems.Add(, , sqlRS.Fields("NewVolume"))
            itmx.SubItems(1) = sqlRS.Fields("OldVolume")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case "PICTURE CHANGE"
            Set itmx = CreateRouting.ToolingChangeList.ListItems.Add(, , sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(1) = GetInvDescription(sqlRS.Fields("CRIBMASTERID"))
            itmx.SubItems(2) = sqlRS.Fields("Type")
            If sqlRS.Fields("APPROVED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            If sqlRS.Fields("COMPLETED") Then
                i = 1
            Else
                i = 0
            End If
            itmx.ListSubItems.Add , , "", i + 1
            itmx.SubItems(5) = sqlRS.Fields("Comments")
            itmx.SubItems(6) = sqlRS.Fields("ItemChangeID")
        Case Else
            MsgBox ("Invalid Type")
        End Select
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Function AttemptCompleteRouting(ProcessChangeID) As Boolean
    Dim ProcessIsDone As Boolean
    Dim OldProcessID As Long
    Dim NewProcessID As Long
    Dim InitialRelease As Boolean
    ProcessIsDone = True
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE PROCESSCHANGEID =" + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        If sqlRS.Fields("Completed") = 0 And Trim(sqlRS.Fields("Type")) <> "PICTURE CHANGE" Then
'            MsgBox (sqlRS.Fields("itemchangeid"))
            ProcessIsDone = False
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    If ProcessIsDone Then
        CreateRouting.Hide
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("COMPLETE") = 1
        InitialRelease = sqlRS.Fields("InitialRelease")
        NewProcessID = sqlRS.Fields("PROCESSID")
        sqlRS.Update
        sqlRS.Close
    Else
        ATTEMPTCOMPLETEPROCESS = False
 '       ProgressBar.Hide
 '       ProgressBar.Timer1.Enabled = False
        Exit Function
    End If
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(NewProcessID), sqlConn, adOpenKeyset, adLockOptimistic
    OldProcessID = sqlRS.Fields("REVOFPROCESSID")
    sqlRS.Fields("REVOFPROCESSID") = 0
    sqlRS.Fields("REVINPROCESS") = 0
    sqlRS.Fields("RELEASED") = 1
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    If Not InitialRelease Then
        DeleteProcessSub (OldProcessID)
    End If
    CreateRouting.ClearFields
    AttemptCompleteRouting = True
  '      ProgressBar.Hide
  '      ProgressBar.Timer1.Enabled = False
End Function

Function AttemptCompleteItem(ItemChangeID As Long) As Boolean
    Dim ItemIsDone As Boolean
    Dim ProcessChangeID As Long
    ItemIsDone = True
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ACTIONS] WHERE ITEMCHANGEID = " + Str(ItemChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        If sqlRS.Fields("Complete") = 0 Then
            ItemIsDone = False
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    If ItemIsDone Then
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE ITEMCHANGEID = " + Str(ItemChangeID), sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("COMPLETED") = 1
        sqlRS.Update
        ProcessChangeID = sqlRS.Fields("PROCESSCHANGEID")
    Else
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE ITEMCHANGEID = " + Str(ItemChangeID), sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("COMPLETED") = 0
        sqlRS.Update
        ProcessChangeID = sqlRS.Fields("PROCESSCHANGEID")
        AttemptCompleteItem = False
    End If
    CreateRouting.ClearFields
    PopulateRouting (ProcessChangeID)
    AttemptCompleteItem = True
End Function

Public Sub SendNeedApprovalNotification(ProcessChangeID)
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST CHANGE MASTER].PROCESSID = [TOOLLIST MASTER].PROCESSID WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST USERS]", sqlConn, adOpenKeyset, adLockReadOnly
    Dim EmailSession As OSSMTP.SMTPSession
    Set EmailSession = New OSSMTP.SMTPSession
    EmailSession.MessageSubject = "APPROVAL NOTIFICATION FOR " + sqlRS.Fields("CUSTOMER") + " - " + sqlRS.Fields("PARTFAMILY") + " - " + sqlRS.Fields("OPERATIONDESCRIPTION")
    EmailSession.MessageText = "TOOL LIST CHANGES ARE AWAING APPROVAL." + vbCrLf + vbCrLf + EmailMessage
    EmailSession.AuthenticationType = AuthNone
    EmailSession.Server = "buschecnc-com01e.mail.protection.outlook.com"
    'TODO use the NotifyMe table to create email sendto list
    'The engineer's want the BuyerCompleted email, Wes wants the DeptMgrApproval email, and Nancy wants the BuyerApproval email only.
    
    'select * from NotifyMe
    'While (Not EOF)
    '{
       'if(NotifyMe.User.ApprovalNotification == true)
       '{
          'SendEmail (NotifyMe.User)
      '}
    '}

    
    EmailSession.SendTo = SQLRS2.Fields("DEPTMGR") + "@buschegroup.com"
    EmailSession.MailFrom = "processchange@buschegroup.com"
    EmailSession.SendEmail
    sqlRS.Close
    SQLRS2.Close
    Set sqlRS = Nothing
    Set SQLRS2 = Nothing
End Sub

Public Sub SendNeedCompleteNotification(ProcessChangeID)
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST CHANGE MASTER].PROCESSID = [TOOLLIST MASTER].PROCESSID WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST USERS]", sqlConn, adOpenKeyset, adLockReadOnly
    Dim EmailSession As OSSMTP.SMTPSession
    Set EmailSession = New OSSMTP.SMTPSession
    EmailSession.MessageSubject = sqlRS.Fields("CUSTOMER") + " - " + sqlRS.Fields("PARTFAMILY") + " - " + sqlRS.Fields("OPERATIONDESCRIPTION")
    EmailSession.MessageText = "TOOL LIST CHANGES ARE AWAING COMPLETION FROM THE BUYER."
    EmailSession.AuthenticationType = AuthNone
    EmailSession.Server = "buschecnc-com01e.mail.protection.outlook.com"
    EmailSession.SendTo = SQLRS2.Fields("Buyer") + "@buschegroup.com"
    EmailSession.MailFrom = "processchange@buschegroup.com"
    EmailSession.SendEmail
    sqlRS.Close
    SQLRS2.Close
    Set sqlRS = Nothing
    Set SQLRS2 = Nothing
End Sub

Public Sub SendCompleteNotification(ProcessChangeID)
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST CHANGE MASTER].PROCESSID = [TOOLLIST MASTER].PROCESSID WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST USERS]", sqlConn, adOpenKeyset, adLockReadOnly
    Dim EmailSession As OSSMTP.SMTPSession
    Set EmailSession = New OSSMTP.SMTPSession
    EmailSession.MessageSubject = sqlRS.Fields("CUSTOMER") + " - " + sqlRS.Fields("PARTFAMILY") + " - " + sqlRS.Fields("OPERATIONDESCRIPTION")
    EmailSession.MessageText = "TOOL LIST CHANGE IS COMPLETE."
    EmailSession.AuthenticationType = AuthNone
    EmailSession.Server = "buschecnc-com01e.mail.protection.outlook.com"
    EmailSession.SendTo = sqlRS.Fields("Engineer") + "@buschegroup.com"
    EmailSession.MailFrom = "processchange@buschegroup.com"
    EmailSession.SendEmail
    Set EmailSession = Nothing
    sqlRS.Close
    SQLRS2.Close
    Set sqlRS = Nothing
    Set SQLRS2 = Nothing
    
End Sub

Public Sub PopulateActionList(ItemChangeID)
    ActionDetails.ActionItemList.ListItems.Clear
    Dim itmx
    Set SQLRS2 = New ADODB.Recordset
    SQLRS2.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE ITEMCHANGEID = " + ItemChangeID, sqlConn, adOpenKeyset, adLockReadOnly
    
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT ACTIONID, COMPLETE, ACTIONITEMTEXT FROM [TOOLLIST CHANGE ACTIONS] INNER JOIN [TOOLLIST CHANGE ACTION TEXT] ON [TOOLLIST CHANGE ACTIONS].ACTIONITEM = [TOOLLIST CHANGE ACTION TEXT].ACTIONITEMNUMBER WHERE ITEMCHANGEID = " + Str(ItemChangeID), sqlConn, adOpenKeyset, adLockReadOnly
    
    
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT * FROM INVENTRY WHERE ITEMNUMBER = '" + SQLRS2.Fields("CRIBMASTERID") + "'", CribConn, adOpenDynamic, adLockReadOnly
    
    While Not sqlRS.EOF
        Set itmx = ActionDetails.ActionItemList.ListItems.Add(, , sqlRS.Fields("ACTIONID"))
        itmx.Checked = sqlRS.Fields("COMPLETE")
        itmx.SubItems(1) = sqlRS.Fields("ACTIONITEMTEXT")
        sqlRS.MoveNext
    Wend
    Select Case Trim(SQLRS2.Fields("TYPE"))
        Case "ADDED"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "Item Number"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = True
            ActionDetails.Line3Lbl.Visible = True
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = True
            ActionDetails.Line4Lbl.Visible = True
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            If Not IsNull(CribRS.Fields("DESCRIPTION1")) Then
                ActionDetails.Line1TXT.Text = CribRS.Fields("DESCRIPTION1")
            End If
            If Not IsNull(CribRS.Fields("ITEMCLASS")) Then
                ActionDetails.Line2TXT.Text = CribRS.Fields("ITEMCLASS")
            End If
            If Not IsNull(CribRS.Fields("MANUFACTURER")) Then
                ActionDetails.Line3TXT.Text = CribRS.Fields("MANUFACTURER")
            End If
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "STOCK TOOLBOSS"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "Item Number"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = True
            ActionDetails.Line3Lbl.Visible = True
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = True
            ActionDetails.Line4Lbl.Visible = True
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            If Not IsNull(CribRS.Fields("DESCRIPTION1")) Then
                ActionDetails.Line1TXT.Text = CribRS.Fields("DESCRIPTION1")
            End If
            If Not IsNull(CribRS.Fields("ITEMCLASS")) Then
                ActionDetails.Line2TXT.Text = CribRS.Fields("ITEMCLASS")
            End If
            If Not IsNull(CribRS.Fields("MANUFACTURER")) Then
                ActionDetails.Line3TXT.Text = CribRS.Fields("MANUFACTURER")
            End If
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "ADDEDM"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "Item Number"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = True
            ActionDetails.Line3Lbl.Visible = True
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = True
            ActionDetails.Line4Lbl.Visible = True
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            If Not IsNull(CribRS.Fields("DESCRIPTION1")) Then
                ActionDetails.Line1TXT.Text = CribRS.Fields("DESCRIPTION1")
            End If
            If Not IsNull(CribRS.Fields("ITEMCLASS")) Then
                ActionDetails.Line2TXT.Text = CribRS.Fields("ITEMCLASS")
            End If
            If Not IsNull(CribRS.Fields("MANUFACTURER")) Then
                ActionDetails.Line3TXT.Text = CribRS.Fields("MANUFACTURER")
            End If
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "REMOVED"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "Item Number"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = True
            ActionDetails.Line3Lbl.Visible = True
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = False
            ActionDetails.Line4Lbl.Visible = False
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            ActionDetails.Line1TXT.Text = CribRS.Fields("DESCRIPTION1")
            ActionDetails.Line2TXT.Text = CribRS.Fields("ITEMCLASS")
            ActionDetails.Line3TXT.Text = CribRS.Fields("MANUFACTURER")
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "REMOVEDM"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "Item Number"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = True
            ActionDetails.Line3Lbl.Visible = True
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = False
            ActionDetails.Line4Lbl.Visible = False
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            ActionDetails.Line1TXT.Text = CribRS.Fields("DESCRIPTION1")
            ActionDetails.Line2TXT.Text = CribRS.Fields("ITEMCLASS")
            ActionDetails.Line3TXT.Text = CribRS.Fields("MANUFACTURER")
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "USAGE CHANGE"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "Item Number"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = True
            ActionDetails.Line3Lbl.Visible = True
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = True
            ActionDetails.Line4Lbl.Visible = True
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            ActionDetails.Line1TXT.Text = CribRS.Fields("DESCRIPTION1")
            ActionDetails.Line2TXT.Text = CribRS.Fields("ITEMCLASS")
            ActionDetails.Line3TXT.Text = CribRS.Fields("MANUFACTURER")
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "PLANT"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "New Plants"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Old Plants"
            ActionDetails.Line3TXT.Visible = False
            ActionDetails.Line3Lbl.Visible = False
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = False
            ActionDetails.Line4Lbl.Visible = False
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            ActionDetails.Line1TXT.Text = SQLRS2.Fields("NEWPLANTS")
            ActionDetails.Line2TXT.Text = SQLRS2.Fields("OLDPLANTS")
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "STATUS"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "New Status"
            ActionDetails.Line2TXT.Visible = False
            ActionDetails.Line2Lbl.Visible = False
            ActionDetails.Line2Lbl.Caption = "Item Group"
            ActionDetails.Line3TXT.Visible = False
            ActionDetails.Line3Lbl.Visible = False
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = False
            ActionDetails.Line4Lbl.Visible = False
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            ActionDetails.Line1TXT.Text = SQLRS2.Fields("NEWSTATUS")
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case "VOLUME"
            ActionDetails.Line1TXT.Visible = True
            ActionDetails.Line1Lbl.Visible = True
            ActionDetails.Line1Lbl.Caption = "New Volume"
            ActionDetails.Line2TXT.Visible = True
            ActionDetails.Line2Lbl.Visible = True
            ActionDetails.Line2Lbl.Caption = "Old Volume"
            ActionDetails.Line3TXT.Visible = False
            ActionDetails.Line3Lbl.Visible = False
            ActionDetails.Line3Lbl.Caption = "Manufacturer"
            ActionDetails.Line4TXT.Visible = False
            ActionDetails.Line4Lbl.Visible = False
            ActionDetails.Line4Lbl.Caption = "Usage"
            
            ActionDetails.ActionTXT.Text = SQLRS2.Fields("TYPE")
            ActionDetails.Line1TXT.Text = SQLRS2.Fields("NEWVOLUME")
            ActionDetails.Line2TXT.Text = SQLRS2.Fields("OLDVOLUME")
            ActionDetails.CommentsTXT.Text = SQLRS2.Fields("COMMENTS")
            ActionDetails.ChangeIDTXT.Text = SQLRS2.Fields("ItemChangeID")
        Case Else
            MsgBox ("Invalid Type")
        End Select
    sqlRS.Close
    SQLRS2.Close
    CribRS.Close
    Set sqlRS = Nothing
    Set SQLRS2 = Nothing
    Set CribRS = Nothing
    ActionDetails.Show
End Sub
Public Sub ApproveRouting(ProcessChangeID)
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "UPDATE [TOOLLIST CHANGE ITEMS] SET APPROVED = 1 WHERE PROCESSCHANGEID = " + Str(ProcessChangeID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    sqlCMD.CommandText = "UPDATE [TOOLLIST CHANGE MASTER] SET APPROVED = 1 WHERE PROCESSCHANGEID = " + Str(ProcessChangeID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
End Sub

Public Sub PopulateMainRoutingList()
    Dim itmx2
    Dim i As Integer
    ChangeList.ListView1.ListItems.Clear
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT PROCESSCHANGEID, DATEINITIATED, ENGINEER, CUSTOMER, COMPLETE, [TOOLLIST CHANGE MASTER].APPROVED, PARTFAMILY, COMMENTS, OPERATIONDESCRIPTION FROM [TOOLLIST CHANGE MASTER] INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST CHANGE MASTER].PROCESSID = [TOOLLIST MASTER].PROCESSID WHERE [TOOLLIST CHANGE MASTER].APPROVED = 0 OR COMPLETE = 0", sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        Set itmx = ChangeList.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSCHANGEID"))
        itmx.SubItems(1) = sqlRS.Fields("DATEINITIATED")
        itmx.SubItems(2) = sqlRS.Fields("CUSTOMER")
        itmx.SubItems(3) = sqlRS.Fields("PARTFAMILY")
        itmx.SubItems(4) = sqlRS.Fields("OPERATIONDESCRIPTION")
        itmx.SubItems(5) = sqlRS.Fields("COMMENTS")
        itmx.SubItems(6) = sqlRS.Fields("ENGINEER")
        If sqlRS.Fields("APPROVED") Then
            i = 1
        Else
            i = 0
        End If
        itmx.ListSubItems.Add , , "", i + 1
        If sqlRS.Fields("COMPLETE") Then
            i = 1
        Else
            i = 0
        End If
        itmx.ListSubItems.Add , , "", i + 1
        itmx.SubItems(9) = sqlRS.Fields("APPROVED")
        itmx.SubItems(10) = sqlRS.Fields("COMPLETE")
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub UpdateActionItems()
    Dim i As Integer
    i = 1
    Set sqlRS = New ADODB.Recordset
    While i <= ActionDetails.ActionItemList.ListItems.Count
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ACTIONS] WHERE ACTIONID = " + ActionDetails.ActionItemList.ListItems.Item(i).Text, sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("COMPLETE") = ActionDetails.ActionItemList.ListItems.Item(i).Checked
        sqlRS.Update
        sqlRS.Close
        i = i + 1
    Wend
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE ITEMCHANGEID = " + ActionDetails.ChangeIDTXT.Text, sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("COMMENTS") = Trim(ActionDetails.CommentsTXT.Text)
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    AttemptCompleteItem (ActionDetails.ChangeIDTXT.Text)
End Sub
Public Sub UpdateActionItemsForTools()
    Dim i As Integer
    i = 1
    Set sqlRS = New ADODB.Recordset
    While i <= ItemComments.ActionItemList.ListItems.Count
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ACTIONS] WHERE ACTIONID = " + ItemComments.ActionItemList.ListItems.Item(i).Text, sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("COMPLETE") = ItemComments.ActionItemList.ListItems.Item(i).Checked
        sqlRS.Update
        sqlRS.Close
        i = i + 1
    Wend
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] WHERE ITEMCHANGEID = " + ItemComments.ChangeIDTXT.Text, sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("COMMENTS") = Trim(ItemComments.CommentsTXT.Text)
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    AttemptCompleteItem (ItemComments.ChangeIDTXT.Text)
End Sub
Public Sub CopyProcessMaster(pID As Long)
 Set cmd = New ADODB.Command
 cmd.ActiveConnection = sqlConn
 cmd.CommandType = adCmdStoredProc
 cmd.CommandText = "CopyProcessMaster"

 cmd.Parameters.Append cmd.CreateParameter("pid", adInteger, adParamInput, 3, pID)
' cmd.Parameters.Append cmd.CreateParameter("empid", adVarChar, adParamInput, 6, str_empid)

 Set rs = cmd.Execute

 'If Not rs.EOF Then
 ' txt_firstname = rs.Fields(0)
 ' txt_title = rs.Fields(1)
 ' txt_address = rs.Fields(2)
 'End If

 Set cmd.ActiveConnection = Nothing

End Sub

Function SPCopyProcessForChanges(pID As Long) As Long
    Dim NewProcIDObj As Object
    Dim newProcId As Long
    Dim sqlRS As New ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim itemID As Object
    Dim toolID As Object
    Dim itemImage As Object
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = sqlConn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "CopyProcessForChanges"
    
    cmd.Parameters.Append cmd.CreateParameter("pid", adInteger, adParamInput, 3, pID)
 
'return a recordset of all the [toollist item]s for all the tools in the new ToolList

    Set sqlRS = cmd.Execute
    Set NewProcIDObj = sqlRS.Fields("ProcessID")
    newProcId = NewProcIDObj
    While Not sqlRS.EOF
        Set itemID = sqlRS.Fields("itemId")
        Set toolID = sqlRS.Fields("toolId")
        If (0 = sqlRS("itemImage")) Then
            colItemImages.Add "F", Str(itemID)
        Else
            colItemImages.Add "T", Str(itemID)
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    SPCopyProcessForChanges = newProcId
    
'    sqlRS.Open "SELECT MAX(PROCESSID) FROM [TOOLLIST MASTER]", sqlConn, adOpenKeyset, adLockReadOnly
'    NewProcessID = sqlRS.Fields("ProcessID")
'    sqlRS.Close
 
 Set cmd.ActiveConnection = Nothing
End Function
Function CopyProcessForChanges(pID As Long) As Long
    Dim NewProcessID As Long
    Dim newToolID As Long
    Dim oldToolID As Long
    'Copy Process Master
    Set sqlRS = New ADODB.Recordset
    Set SQLRS2 = New ADODB.Recordset
    Set SQLRS3 = New ADODB.Recordset
    Set SQLRS4 = New ADODB.Recordset
    
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    SQLRS2.Open "[TOOLLIST MASTER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    
    SQLRS2.AddNew
    SQLRS2.Fields("CUSTOMER") = sqlRS.Fields("CUSTOMER")
    SQLRS2.Fields("PartFamily") = sqlRS.Fields("PartFamily")
    SQLRS2.Fields("OperationNumber") = sqlRS.Fields("OperationNumber")
    SQLRS2.Fields("OperationDescription") = sqlRS.Fields("OperationDescription")
    SQLRS2.Fields("Obsolete") = sqlRS.Fields("Obsolete")
    SQLRS2.Fields("RELEASED") = sqlRS.Fields("RELEASED")
    SQLRS2.Fields("AnnualVolume") = sqlRS.Fields("AnnualVolume")
    SQLRS2.Fields("MultiTurret") = sqlRS.Fields("MultiTurret")
    SQLRS2.Fields("RevOfProcessID") = sqlRS.Fields("ProcessID")
    SQLRS2.Fields("RevInProcess") = 0
    SQLRS2.Fields("OriginalProcessID") = sqlRS.Fields("OriginalProcessID")
    
    SQLRS2.Update
    SQLRS2.Close
    sqlRS.Close
    
    SQLRS2.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY PROCESSID DESC", sqlConn, adOpenKeyset, adLockReadOnly
    NewProcessID = SQLRS2.Fields("ProcessID")
    SQLRS2.Close
    'End Copy Process Master
    'Copy ToolList Plants
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    SQLRS2.Open "[TOOLLIST PLANT]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    
    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = NewProcessID
        SQLRS2.Fields("Plant") = sqlRS.Fields("Plant")
        sqlRS.MoveNext
        SQLRS2.Update
    Wend
    sqlRS.Close
    SQLRS2.Close
    
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    SQLRS2.Open "[TOOLLIST PARTNUMBERS]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
         
    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = NewProcessID
        SQLRS2.Fields("PartNumbers") = sqlRS.Fields("PartNumbers")
        sqlRS.MoveNext
        SQLRS2.Update
    Wend
    sqlRS.Close
    SQLRS2.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    SQLRS2.Open "[TOOLLIST MISC]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable

    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = NewProcessID
        SQLRS2.Fields("Manufacturer") = sqlRS.Fields("Manufacturer")
        SQLRS2.Fields("ToolType") = sqlRS.Fields("ToolType")
        SQLRS2.Fields("ToolDescription") = sqlRS.Fields("ToolDescription")
        SQLRS2.Fields("Consumable") = sqlRS.Fields("Consumable")
        SQLRS2.Fields("QuantityPerCuttingEdge") = sqlRS.Fields("QuantityPerCuttingEdge")
        SQLRS2.Fields("AdditionalNotes") = sqlRS.Fields("AdditionalNotes")
        SQLRS2.Fields("NumberOfCuttingEdges") = sqlRS.Fields("NumberOfCuttingEdges")
        SQLRS2.Fields("Quantity") = sqlRS.Fields("Quantity")
        SQLRS2.Fields("CribToolID") = sqlRS.Fields("CribToolID")
        SQLRS2.Fields("TOOLBOSSSTOCK") = sqlRS.Fields("TOOLBOSSSTOCK")
        sqlRS.MoveNext
        SQLRS2.Update
    Wend
    
   
    sqlRS.Close
    SQLRS2.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST FIXTURE] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    SQLRS2.Open "[TOOLLIST FIXTURE]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable

    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = NewProcessID
        SQLRS2.Fields("Manufacturer") = sqlRS.Fields("Manufacturer")
        SQLRS2.Fields("ToolType") = sqlRS.Fields("ToolType")
        SQLRS2.Fields("ToolDescription") = sqlRS.Fields("ToolDescription")
        SQLRS2.Fields("AdditionalNotes") = sqlRS.Fields("AdditionalNotes")
        SQLRS2.Fields("Quantity") = sqlRS.Fields("Quantity")
        SQLRS2.Fields("CribToolID") = sqlRS.Fields("CribToolID")
        SQLRS2.Fields("TOOLBOSSSTOCK") = sqlRS.Fields("TOOLBOSSSTOCK")
        sqlRS.MoveNext
        SQLRS2.Update
    Wend
    sqlRS.Close
    SQLRS2.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST REV] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    SQLRS2.Open "[TOOLLIST REV]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable

    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = NewProcessID
        SQLRS2.Fields("Revision") = sqlRS.Fields("Revision")
        SQLRS2.Fields("Revision Description") = sqlRS.Fields("Revision Description")
        SQLRS2.Fields("Revision Date") = sqlRS.Fields("Revision Date")
        SQLRS2.Fields("Revision By") = sqlRS.Fields("Revision By")
        sqlRS.MoveNext
        SQLRS2.Update
    Wend

    sqlRS.Close
    SQLRS2.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    
    
    While Not sqlRS.EOF
        SQLRS2.Open "[TOOLLIST TOOL]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = NewProcessID
        SQLRS2.Fields("ToolNumber") = sqlRS.Fields("ToolNumber")
        SQLRS2.Fields("OpDescription") = sqlRS.Fields("OpDescription")
        SQLRS2.Fields("Alternate") = sqlRS.Fields("Alternate")
        SQLRS2.Fields("PartSpecific") = sqlRS.Fields("PartSpecific")
        SQLRS2.Fields("AdjustedVolume") = sqlRS.Fields("AdjustedVolume")
        SQLRS2.Fields("ToolOrder") = sqlRS.Fields("ToolOrder")
        SQLRS2.Fields("Turret") = sqlRS.Fields("Turret")
        SQLRS2.Fields("ToolLength") = sqlRS.Fields("ToolLength")
        SQLRS2.Fields("OffsetNumber") = sqlRS.Fields("OffsetNumber")
        SQLRS2.Update
        SQLRS2.Close
        SQLRS2.Open "SELECT * FROM [TOOLLIST TOOL] ORDER BY TOOLID DESC", sqlConn, adOpenKeyset, adLockReadOnly
        newToolID = SQLRS2.Fields("ToolID")
        oldToolID = sqlRS.Fields("ToolID")
        SQLRS2.Close

        SQLRS3.Open "SELECT * FROM [TOOLLIST ITEM] WHERE TOOLID = " + Str(oldToolID), sqlConn, adOpenKeyset, adLockReadOnly
        SQLRS4.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        While Not SQLRS3.EOF
            SQLRS4.AddNew
            SQLRS4.Fields("ProcessID") = NewProcessID
            SQLRS4.Fields("ToolID") = newToolID
            SQLRS4.Fields("ToolType") = SQLRS3.Fields("ToolType")
            SQLRS4.Fields("ToolDescription") = SQLRS3.Fields("ToolDescription")
            SQLRS4.Fields("Manufacturer") = SQLRS3.Fields("Manufacturer")
            SQLRS4.Fields("Consumable") = SQLRS3.Fields("Consumable")
            SQLRS4.Fields("QuantityPerCuttingEdge") = SQLRS3.Fields("QuantityPerCuttingEdge")
            SQLRS4.Fields("AdditionalNotes") = SQLRS3.Fields("AdditionalNotes")
            SQLRS4.Fields("NumberOfCuttingEdges") = SQLRS3.Fields("NumberOfCuttingEdges")
            SQLRS4.Fields("Quantity") = SQLRS3.Fields("Quantity")
            SQLRS4.Fields("CribToolID") = SQLRS3.Fields("CribToolID")
            SQLRS4.Fields("Regrindable") = SQLRS3.Fields("Regrindable")
            SQLRS4.Fields("QtyPerRegrind") = SQLRS3.Fields("QtyPerRegrind")
            SQLRS4.Fields("NumOfRegrinds") = SQLRS3.Fields("NumOfRegrinds")
            SQLRS4.Fields("ParentItem") = SQLRS3.Fields("ParentItem")
            SQLRS4.Fields("TOOLBOSSSTOCK") = SQLRS3.Fields("TOOLBOSSSTOCK")
            SQLRS4.Fields("ItemImage") = SQLRS3.Fields("ItemImage")
            
            SQLRS4.Update
            SQLRS4.Close
            SQLRS4.Open "select * from [TOOLLIST ITEM] order by itemid desc", sqlConn, adOpenKeyset, adLockReadOnly
            newItemId = SQLRS4.Fields("ItemID")
            SQLRS4.Close
'return a recordset of all the [toollist item]s for all the tools in the new ToolList
            SQLRS4.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
            If (IsNull(SQLRS3.Fields("ItemImage")) = True) Then
                colItemImages.Add "F", Str(newItemId)
            Else
                colItemImages.Add "T", Str(newItemId)
            End If
            
            SQLRS3.MoveNext
        Wend
        SQLRS3.Close
        SQLRS4.Close
        
        SQLRS3.Open "SELECT * FROM [TOOLLIST TOOLPARTNUMBER] WHERE TOOLID = " + Str(oldToolID), sqlConn, adOpenKeyset, adLockReadOnly
        SQLRS4.Open "[TOOLLIST TOOLPARTNUMBER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        While Not SQLRS3.EOF
            SQLRS4.AddNew
            SQLRS4.Fields("ToolID") = newToolID
            SQLRS4.Fields("PartNumber") = SQLRS3.Fields("PartNumber")
            SQLRS4.Update
            SQLRS3.MoveNext
        Wend
        SQLRS3.Close
        SQLRS4.Close
        sqlRS.MoveNext
    Wend
    CopyProcessForChanges = NewProcessID
End Function

Public Sub SubmitForInitialRelease(pID As Long)
    Dim ProcessChangeID As Long
    Dim ItemChangeID As Long
    Set sqlRS = New ADODB.Recordset
    ProgressBar.Show
    ProgressBar.Timer1.Enabled = True
    reportOpened = False
    
    sqlRS.CursorLocation = adUseClient
    sqlRS.Open "[TOOLLIST CHANGE MASTER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("PROCESSID") = pID
    sqlRS.Fields("COMPLETE") = False
    sqlRS.Fields("COMMENTS") = "INITIAL RELEASE"
    sqlRS.Fields("ENGINEER") = Trim(Environ("USERNAME"))
    sqlRS.Fields("DATEINITIATED") = Date
    sqlRS.Fields("DATECOMPLETE") = #1/1/1900#
    sqlRS.Fields("APPROVED") = 0
    sqlRS.Fields("INITIALRELEASE") = 1
    sqlRS.Fields("OLDPROCESSID") = pID
    sqlRS.Update
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] ORDER BY PROCESSCHANGEID DESC", sqlConn, adOpenKeyset, adLockReadOnly
    ProcessChangeID = sqlRS.Fields("ProcessChangeID")
    sqlRS.Close
    
    sqlRS.Open "[TOOLLIST CHANGE ITEMS]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ProcessChangeID") = ProcessChangeID
    sqlRS.Fields("Type") = "STATUS"
    sqlRS.Fields("CribmasterID") = ""
    sqlRS.Fields("NewStatus") = "RELEASED"
    sqlRS.Fields("NewPlants") = ""
    sqlRS.Fields("OldPlants") = ""
    sqlRS.Fields("NewVolume") = 0
    sqlRS.Fields("OldVolume") = 0
    sqlRS.Fields("DispositionMethod") = ""
    sqlRS.Fields("Comments") = "RELEASED JOB FOR PRODUCTION"
    sqlRS.Fields("Completed") = 0
    sqlRS.Fields("APPROVED") = 0
    sqlRS.Update
    sqlRS.Close
    
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE ITEMS] ORDER BY ITEMCHANGEID DESC", sqlConn, adOpenKeyset, adLockReadOnly
    ItemChangeID = sqlRS.Fields("ItemChangeID")
    sqlRS.Close
       
    sqlRS.Open "[TOOLLIST CHANGE ACTIONS]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ITEMCHANGEID") = ItemChangeID
    sqlRS.Fields("PROCESSCHANGEID") = ProcessChangeID
    sqlRS.Fields("ACTIONITEM") = 10
    sqlRS.Fields("COMPLETE") = 0
    sqlRS.Update
    sqlRS.AddNew
    sqlRS.Fields("ITEMCHANGEID") = ItemChangeID
    sqlRS.Fields("PROCESSCHANGEID") = ProcessChangeID
    sqlRS.Fields("ACTIONITEM") = 9
    sqlRS.Fields("COMPLETE") = 0
    sqlRS.Update
    sqlRS.Close
        
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("REVINPROCESS") = 1
    sqlRS.Fields("REVOFPROCESSID") = sqlRS.Fields("PROCESSID")
    sqlRS.Update
    sqlRS.Close
    SendNeedApprovalNotification (ProcessChangeID)
    Reset
    ExitLoop = True
    MDIForm1.RefreshMenuOptions
    If (True = reportOpened) Then
        ReportForm.Hide
    End If
    CreateRouting.Hide
    ProgressBar.Hide
    ProgressBar.Timer1.Enabled = False
    MsgBox ("ROUTING SENT ON FOR APPROVAL")
End Sub

Function RevInProcess(pID As Long) As Boolean
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    RevInProcess = sqlRS.Fields("REVINPROCESS") = True
    sqlRS.Close
    Set sqlRS = Nothing
End Function

Function IsInitialRelease(pID As Long) As Boolean
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] WHERE PROCESSCHANGEID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    IsInitialRelease = sqlRS.Fields("InitialRelease")
    processId = sqlRS.Fields("ProcessID")
    sqlRS.Close
    Set sqlRS = Nothing
End Function

Public Sub RemoveRevInProcess(pID As Long)
    Dim OriginalProcessID As Long
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockOptimistic
    OriginalProcessID = sqlRS.Fields("REVOFPROCESSID")
    sqlRS.Fields("REVOFPROCESSID") = 0
    sqlRS.Update
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(OriginalProcessID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("REVINPROCESS") = 0
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub DeleteProcessChange(pID As Long)
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST CHANGE MASTER] WHERE PROCESSCHANGEID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST CHANGE ITEMS] WHERE PROCESSCHANGEID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New ADODB.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST CHANGE ACTIONS] WHERE PROCESSCHANGEID =" + Str(pID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
End Sub
Function SPCalcUsage(cmId As String, procId As Long) As Long
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = sqlConn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "CalcUsage"
    Dim trmCmId As String
    trmCmId = Trim(cmId)
'cmd.Parameters.Append cmd.CreateParameter_
 '   ("empid", adVarChar, adParamInput, 6, txt_empid.Text)
'cmd.Parameters.Append cmd.CreateParameter_
 '   ("firstname", adVarChar, adParamInput, 30, txt_firstname.Text)
'cmd.Parameters.Append cmd.CreateParameter_
 '   ("title", adVarChar, adParamInput, 30, txt_title.Text)
'cmd.Parameters.Append cmd.CreateParameter_
 '   ("address", adVarChar, adParamInput, 100, txt_address.Text)
'cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
    
    cmd.Parameters.Append cmd.CreateParameter("cmId", adVarChar, adParamInput, 10, trmCmId)
    cmd.Parameters.Append cmd.CreateParameter("procId", adInteger, adParamInput, 10, procId)
    cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
    cmd.Execute
    SPCalcUsage = cmd("result")
 Set cmd.ActiveConnection = Nothing
End Function
Function SPAllOtherUsage(cmId As String, newProcId As Long, oldProcId As Long) As Long
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = sqlConn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "AllOtherUsage"
    
    cmd.Parameters.Append cmd.CreateParameter("cmId", adVarChar, adParamInput, 10, Trim(cmId))
    cmd.Parameters.Append cmd.CreateParameter("newProcId", adInteger, adParamInput, 10, newProcId)
    cmd.Parameters.Append cmd.CreateParameter("oldProcId", adInteger, adParamInput, 10, oldProcId)
    cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
    cmd.Execute
    SPAllOtherUsage = cmd("result")
 Set cmd.ActiveConnection = Nothing
End Function
'pID = processChangeId
Public Sub PopulateItemChangeInfo(pID As Long, cmId As String)
    ClearItemCommentsFields
    Dim CRIBRS2 As ADODB.Recordset
    Set CRIBRS2 = New ADODB.Recordset
    Set sqlRS = New ADODB.Recordset
    Set SQLRS2 = New ADODB.Recordset
    Dim sum As Integer
    Dim binstring As String
    Dim Usage As Double
    ItemComments.ChangeIDTXT.Text = ""
    If CreateRouting.GetViewingType = "Completion" Then
        'Add the MRO TODO checkboxes
        ItemComments.Height = 7785
        ItemComments.ActionItemList.Visible = True
        ItemComments.ActionItemList.ListItems.Clear
        sqlRS.Open "SELECT ACTIONID, COMPLETE, ACTIONITEMTEXT FROM [TOOLLIST CHANGE ACTIONS] INNER JOIN [TOOLLIST CHANGE ACTION TEXT] ON [TOOLLIST CHANGE ACTIONS].ACTIONITEM = [TOOLLIST CHANGE ACTION TEXT].ACTIONITEMNUMBER WHERE ITEMCHANGEID = " + CreateRouting.ToolingChangeList.SelectedItem.SubItems(6), sqlConn, adOpenKeyset, adLockReadOnly
        While Not sqlRS.EOF
            Set itmx = ItemComments.ActionItemList.ListItems.Add(, , sqlRS.Fields("ACTIONID"))
            itmx.Checked = sqlRS.Fields("COMPLETE")
            itmx.SubItems(1) = sqlRS.Fields("ACTIONITEMTEXT")
            sqlRS.MoveNext
        Wend
        sqlRS.Close
        ItemComments.ChangeIDTXT.Text = CreateRouting.ToolingChangeList.SelectedItem.SubItems(6)
    Else
        'No checkboxes for Approvals,Creation
        ItemComments.Height = 5940
        ItemComments.ActionItemList.Visible = False
    End If


'tO DO
'TOOLLIST CHANGE TASKS -- REPORTING WRONG QTY OR NEEDS TO BE EXPANDED
'AND THERE IS 2 LINES IN THE VERIFY CRIBMASTER MINS AND MAXS LINE WHEN THERE SHOULD BE 1




    If CreateRouting.GetViewingType <> "Creation" Then
        sqlRS.Open "SELECT * FROM [TOOLLIST CHANGE MASTER] WHERE PROCESSCHANGEID = " + Str(pID), sqlConn, adOpenDynamic, adLockReadOnly
        processId = sqlRS.Fields("PROCESSID")
        sqlRS.Close
    End If
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(processId), sqlConn, adOpenDynamic, adLockReadOnly
    OldProcessID = sqlRS.Fields("REVOFPROCESSID")
    sqlRS.Close
    
    If ((actionType <> "REMOVED") And (actionType <> "REMOVEDM") And (actionType <> "REMOVEDF")) Then
        Dim mu As Long
        Dim omu As Long
        Dim aou As Long
        Dim usageSum As Long
        
        mu = SPCalcUsage(cmId, processId)
        omu = SPCalcUsage(cmId, OldProcessID)
        aou = SPAllOtherUsage(cmId, processId, OldProcessID)
        usageSum = mu + aou
        ItemComments.MonthlyUsageTXT.Text = Str(mu)
        ItemComments.OldMonthlyUsageTXT.Text = Str(omu)
        ItemComments.CombinedUsageTXT.Text = Str(usageSum)
    End If
    
  '  CONTINUE WITH THIS FUNCTION
    Set CribRS = New ADODB.Recordset
    CribRS.Open "SELECT DESCRIPTION1, Manufacturer, ItemClass, [INVENTRY].ItemNumber, Cost " & _
    "FROM [INVENTRY] LEFT OUTER JOIN [ALTVENDOR] ON [INVENTRY].[ALTVENDORNO] = [ALTVENDOR].[RECNUMBER] " & _
    "WHERE [INVENTRY].ITEMNUMBER = '" + cmId + "' OR [INVENTRY].ITEMNUMBER = '" + cmId + "R'", CribConn, adOpenKeyset, adLockReadOnly
    If Not CribRS.EOF Then
        ItemComments.ItemNumberTXT.Text = CribRS.Fields("Description1")
        If Not IsNull(CribRS.Fields("Manufacturer")) Then
           ItemComments.ManufacturerTXT.Text = CribRS.Fields("Manufacturer")
        End If
        ItemComments.ItemGroupTXT.Text = CribRS.Fields("ItemClass")
        If Not IsNull(CribRS.Fields("COST")) Then
            If InStr(CribRS.Fields("ItemNumber"), "R") > 0 Then
                ItemComments.ReworkCostTXT.Text = CribRS.Fields("Cost")
            Else
                ItemComments.NewCostTXT.Text = CribRS.Fields("Cost")
            End If
        End If
        
        While Not CribRS.EOF
            'For the item and rework item
            CRIBRS2.Open "SELECT ITEM, CRIBBIN, BINQUANTITY FROM STATION WHERE ITEM = '" + CribRS.Fields("ItemNumber") + "'", CribConn, adOpenKeyset, adLockReadOnly
            While Not CRIBRS2.EOF
                binstring = CRIBRS2.Fields("CribBin") + ", " + binstring
                If InStr(CribRS.Fields("ItemNumber"), "R") > 0 Then
                    ItemComments.ReworkQtyTXT.Text = CRIBRS2.Fields("BinQuantity")
                Else
                    ItemComments.NewQtyTXT.Text = CRIBRS2.Fields("BinQuantity")
                End If
                CRIBRS2.MoveNext
            Wend
            CRIBRS2.Close
            CribRS.MoveNext
        Wend
        ItemComments.TotalTXT.Text = Str((Val(ItemComments.ReworkCostTXT.Text) * Val(ItemComments.ReworkQtyTXT.Text)) + (Val(ItemComments.NewCostTXT.Text) * Val(ItemComments.NewQtyTXT.Text)))
        ItemComments.BinTxt.Text = binstring
    Else
        ItemComments.TotalTXT.Text = "0"
        ItemComments.BinTxt.Text = ""
    End If
    CribRS.Close
    Set CribRS = Nothing
    Set CRIBRS2 = Nothing
    Set sqlRS = Nothing
    Set SQLRS2 = Nothing

    ItemComments.ActionTXT.Text = CreateRouting.ToolingChangeList.SelectedItem.SubItems(2)
    ItemComments.Show
    
End Sub

Public Sub ClearItemCommentsFields()
    ItemComments.BinTxt.Text = ""
    ItemComments.ActionTXT.Text = ""
    ItemComments.CommentsTXT.Text = ""
    ItemComments.ItemGroupTXT.Text = ""
    ItemComments.ItemNumberTXT.Text = ""
    ItemComments.ManufacturerTXT.Text = ""
    ItemComments.MonthlyUsageTXT.Text = ""
    ItemComments.NewCostTXT.Text = ""
    ItemComments.NewQtyTXT.Text = ""
    ItemComments.ReworkCostTXT.Text = "N/A"
    ItemComments.ReworkQtyTXT.Text = "N/A"
    ItemComments.TotalTXT.Text = ""
    ItemComments.ListView1.ListItems.Clear
    ItemComments.CombinedUsageTXT.Text = ""
End Sub

Function DeleteExtraProcess(pID As Long)
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE REVOFPROCESSID = " + Str(pID), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        DeleteProcessSub (sqlRS.Fields("PROCESSID"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
End Function

Public Sub ViewProcesses()
    Dim itmx2 As ListItem
    Set sqlRS = New ADODB.Recordset
    sqlRS.Open openSQLStatement, sqlConn
    ViewProcess.ListView1.ListItems.Clear
    While Not sqlRS.EOF
        Set itmx2 = ViewProcess.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSID"))
        If Not IsNull(sqlRS.Fields("Customer")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("Customer"))
        Else
            itmx2.SubItems(1) = ""
        End If
        If Not IsNull(sqlRS.Fields("PartFamily")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("PartFamily"))
        Else
            itmx2.SubItems(2) = ""
        End If
        If Not IsNull(sqlRS.Fields("OperationDescription")) Then
            itmx2.SubItems(3) = Trim(sqlRS.Fields("OperationDescription"))
        Else
            itmx2.SubItems(3) = ""
        End If
        If Not IsNull(sqlRS.Fields("OperationNumber")) Then
            itmx2.SubItems(4) = Trim(sqlRS.Fields("OperationNumber"))
        Else
            itmx2.SubItems(4) = ""
        End If
        If Not IsNull(sqlRS.Fields("RELEASED")) Then
            itmx2.SubItems(5) = sqlRS.Fields("RELEASED")
        Else
            itmx2.SubItems(5) = ""
        End If
        If Not IsNull(sqlRS.Fields("Obsolete")) Then
            itmx2.SubItems(6) = Trim(sqlRS.Fields("Obsolete"))
        Else
            itmx2.SubItems(6) = ""
        End If
        sqlRS.MoveNext
        itmx2.ForeColor = vbRed
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    OldCribID = ""
    ColorRows ViewProcess.ListView1
    ViewProcess.SortByCustomer
End Sub
Public Sub RefreshReportForViewing()
    If reportOpened = False Then
        If (BUILDTYPE = TEST) Then
            Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_TEST)
        ElseIf (BUILDTYPE = IND) Then
            Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_IND)
        ElseIf (BUILDTYPE = AL) Then
            Set craxReport = craxApp.OpenReport(TOOLLIST_RPT_AL)
        End If
        reportOpened = True
    End If
    
    craxReport.DiscardSavedData
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (processId)
    ReportForm.CRViewer1.ReportSource = craxReport
    ReportForm.CRViewer1.Refresh
    ReportForm.CRViewer1.ViewReport
    ReportForm.CRViewer1.Zoom 80
    
End Sub
'Function CheckVersion() As Boolean
'    Set sqlRS = New ADODB.Recordset
'    sqlRS.Open "SELECT * FROM [TOOLLIST VERSION]", sqlConn, adOpenKeyset, adLockReadOnly
'    If Trim(sqlRS.Fields("VERSION")) = Trim(Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))) Then
'        CheckVersion = True
'    Else
'        CheckVersion = False
'    End If
'End Function

'sets bRefreshActionListError to true if there is a validation check error
Public Sub RefreshActionList(ProcessChangeID As Long)
        Dim Msg, Style, Title, Response
        bRefreshActionListError = False
  '      ProgressBar.Show
  '      ProgressBar.Timer1.Enabled = True
        DoEvents
        Dim OLDPID As Long
        Dim NEWPID As Long
        Set sqlRS = New ADODB.Recordset
        sqlRS.Open "SELECT PROCESSID FROM [TOOLLIST CHANGE MASTER] WHERE PROCESSCHANGEID = " + Str(ProcessChangeID), sqlConn, adOpenKeyset, adLockReadOnly
        Set SQLRS2 = New ADODB.Recordset
        NEWPID = sqlRS.Fields("PROCESSID")
        SQLRS2.Open "SELECT REVOFPROCESSID FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(sqlRS.Fields("PROCESSID")), sqlConn, adOpenKeyset, adLockReadOnly
        OLDPID = SQLRS2.Fields("REVOFPROCESSID")
        sqlRS.Close
        SQLRS2.Close
        Set sqlRS = New ADODB.Recordset
        sqlRS.Open "SELECT * FROM [ToolList Change Items] WHERE PROCESSCHANGEID = '" + Str(ProcessChangeID) + "'", sqlConn, adOpenKeyset, adLockReadOnly
        Set SQLRS2 = New ADODB.Recordset
        SQLRS2.Open "[TOOLLIST CHANGE ACTIONS]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
        
        Set SQLRS3 = New ADODB.Recordset
        Dim TBStockItems As String
        TBStockItems = "SELECT * FROM [ToolList Toolboss Stock Items] where ITEMCLASS LIKE "
        
        Set CribRS = New ADODB.Recordset
        Dim CribQry As String
        CribQry = "SELECT ITEMNUMBER,ITEMCLASS,CRIBBIN " + _
        "FROM INVENTRY LEFT OUTER JOIN STATION " + _
        "ON INVENTRY.ITEMNUMBER = STATION.ITEM " + _
        "where ItemNumber = "
                
        Set SQLRS4 = New ADODB.Recordset
        Set sqlCMD = New ADODB.Command
        Dim itemType As String
        
        'Loop through all [ToolList Change Items]
        Do While Not sqlRS.EOF
            ' Some changes do not require a crib id such as plant changes
            If Trim(sqlRS.Fields("CRIBMASTERID")) <> "" Then
                CribRS.Open CribQry + "'" + Trim(sqlRS.Fields("CRIBMASTERID")) + "'", CribConn, adOpenKeyset, adLockReadOnly
                'START "CANT CONTINUE" ERROR CHECKS
                'if [ToolList Change Items] item is not in Crib the tool list change can not be completed
                'if an item is to be added to the Crib or ToolBoss it has already been added to the crib
                'because it can not be added to the ToolList unless it is already in the Crib
                'If an item is to be removed it must also be in the CribMaster until it is completed by the buyer
                'So something is WRONG if the item is not at least in the Crib before completing a REMOVED change
                If CribRS.EOF Then
                   Msg = "Item Number not in CribMaster: "
                   Style = vbCritical + vbOKOnly  ' Define buttons.
                   Title = "No Item in Crib Error"   ' Define title.
      '             ProgressBar.Hide
     '              ProgressBar.Timer1.Enabled = False
                   Response = MsgBox(Msg, Style, Title)
                   bRefreshActionListError = True
                   Exit Do
                End If
            ' If cribmaster item does not contain a group the tool list change can not be completed
'            If (Trim(sqlRS.Fields("Type")) = "ADDED") Or (Trim(sqlRS.Fields("Type")) = "ADDEDM") Or (Trim(sqlRS.Fields("Type")) = "ADDEDF") _
 '              Or (Trim(sqlRS.Fields("Type")) = "REMOVED") Or (Trim(sqlRS.Fields("Type")) = "REMOVEDM") Or (Trim(sqlRS.Fields("Type")) = "REMOVEDF") _
  '             Or (Trim(sqlRS.Fields("Type")) = "USAGE CHANGE") Then
                If IsNull(CribRS.Fields("ITEMCLASS")) Then
                    Msg = "Please add a group to CribMaster Item: " + CribRS.Fields("ITEMNUMBER") ' Define message.
                    Style = vbCritical + vbOKOnly  ' Define buttons.
                    Title = "No Item Group Error"   ' Define title.
 '                   ProgressBar.Hide
  '                  ProgressBar.Timer1.Enabled = False
                    Response = MsgBox(Msg, Style, Title)
                    bRefreshActionListError = True
                    Exit Do
                Else
                    ' Check to determine if the item's class is in the [ToolList Toolboss Stock Items] table
                    SQLRS3.Open TBStockItems + "'" + Trim(CribRS.Fields("ITEMCLASS")) + "'", sqlConn, adOpenKeyset, adLockReadOnly
                End If
            End If
   '         End If
            'END CAN'T CONTINUE ERROR CHECKS

            ' Delete all [TOOLLIST CHANGE ACTIONS] for this [ToolList Change Items] item
            sqlCMD.CommandText = "DELETE FROM [TOOLLIST CHANGE ACTIONS] WHERE ITEMCHANGEID = " + Str(sqlRS.Fields("ITEMCHANGEID")) + " AND COMPLETE = 0"
            sqlCMD.ActiveConnection = sqlConn
            sqlCMD.Execute
            
            ' rebuild [TOOLLIST CHANGE ACTIONS] records for this [ToolList Change Items] item
            
            ' Common to ADDED, ADDEDM, and ADDEDF tool list change actions
            If (Trim(sqlRS.Fields("Type")) = "ADDED") Or (Trim(sqlRS.Fields("Type")) = "ADDEDM") _
                Or (Trim(sqlRS.Fields("Type")) = "ADDEDF") Then
                'Is Location assigned?
                If IsNull(CribRS.Fields("CRIBBIN")) Then
                    'CREATE CRIBMASTER BIN / SET MIN'S AND MAX'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 1) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 1
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                Else
                    'VERIFY CRIBMASTER MIN'S AND MAX'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 2) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 2
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
                'Is this [ToolList Change Items] item a ToolBoss stocked class?
                If Not SQLRS3.EOF Then
                    'ADD ITEM TO TOOLBOSS
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 3) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 3
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
            
                'Add item type specific code
                Select Case Trim(sqlRS.Fields("Type"))
                    Case "ADDED"
                        itemType = "[TOOLLIST ITEM]"
                    Case "ADDEDM"
                        itemType = "[TOOLLIST MISC]"
                    Case "ADDEDF"
                        itemType = "[TOOLLIST FIXTURE]"
                End Select
                ' Are we going to force this [ToolList Change Items] item to be stocked in a ToolBoss
                ' i.e. class in [ToolList Toolboss Stock Items]
                Dim forceTBStock As String
                forceTBStock = "SELECT TOOLBOSSSTOCK FROM " + itemType + " WHERE TOOLBOSSSTOCK = 1 AND PROCESSID = " + _
                Trim(Str(NEWPID)) + " AND CribToolID = '" + Trim(sqlRS.Fields("CribmasterID")) + "'"
                SQLRS4.Open forceTBStock, sqlConn, adOpenKeyset
                'Force ToolBoss stocked item?
                If SQLRS4.RecordCount > 0 Then
                    ' STOCK THIS ITEM IN TOOLBOSS. THIS NORMALLY UNSTOCKED ITEM NEEDS TO BE PUT IN THE TOOLBOSS.
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 14) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 14
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
                SQLRS4.Close
            End If
            'END CODE SPECIFIC TO ADDED,ADDEDM, AND ADDEDF TYPE [TOOLLIST CHANGE ITEMS]
                
                
            'START CODE SPECIFIC TO REMOVED, REMOVEDF,AND REMOVEDF TYPE [TOOLLIST CHANGE ITEMS]
            If (Trim(sqlRS.Fields("Type")) = "REMOVED") Or (Trim(sqlRS.Fields("Type")) = "REMOVEDM") _
                Or (Trim(sqlRS.Fields("Type")) = "REMOVEDF") Then
                Select Case Trim(sqlRS.Fields("Type"))
                    Case "REMOVED"
                        itemType = "[TOOLLIST ITEM]"
                    Case "REMOVEDM"
                        itemType = "[TOOLLIST MISC]"
                    Case "REMOVEDF"
                        itemType = "[TOOLLIST FIXTURE]"
                End Select
                Dim toolListItem As String
                toolListItem = "SELECT * FROM " + itemType + " INNER JOIN [TOOLLIST MASTER] " + _
                    "ON " + itemType + ".PROCESSID = [TOOLLIST MASTER].PROCESSID " + _
                    "WHERE [TOOLLIST MASTER].OBSOLETE = 0 AND " + _
                    itemType + ".CRIBTOOLID = '" + sqlRS.Fields("CRIBMASTERID") + "' " + _
                    "AND [TOOLLIST MASTER].PROCESSID <> " + Str(OLDPID)
                'Does this cribmaster item exist in any other ToolLists?
                SQLRS4.Open toolListItem, sqlConn, adOpenKeyset, adLockReadOnly
            
                'This cribmaster item does not exist in any other ToolLists
                'and the item does Not have a ToolBoss stocked class
                'Attempt to return item
                If SQLRS3.EOF And SQLRS4.EOF Then
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 6) Then
                        'add to [TOOLLIST CHANGE ACTIONS]
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 6
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                'This cribmaster item DOES NOT exist in any other ToolList
                'and the item DOES have a ToolBoss stocked class
                'Attemp to return item and remove from the ToolBoss
                ElseIf Not SQLRS3.EOF And SQLRS4.EOF Then
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 11) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 11
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 6) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 6
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                'This cribmaster item DOES exist in another ToolList
                'and the item DOES have a ToolBoss stocked class
                'REMOVE ITEM FROM TOOLBOSS RESTRICTIONS
                ElseIf Not SQLRS3.EOF And Not SQLRS4.EOF Then
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 4) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 4
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
                SQLRS4.Close
            End If
         
            'START CODE SPECIFIC TO "USAGE CHANGE" TYPE [TOOLLIST CHANGE ITEMS]
            If (Trim(sqlRS.Fields("Type")) = "USAGE CHANGE") Then
                'Not ToolBoss stocked class
                If SQLRS3.EOF Then
                    'VERIFY CRIBMASTER MIN'S AND MAX'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 2) Then
                        '[TOOLLIST CHANGE ACTIONS]
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 2
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                Else
                    'ToolBoss stocked class
                    'VERIFY TOOLBOSS MIN'S AND MAX'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 12) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 12
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                    'VERIFY CRIBMASTER MIN'S AND MAX'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 2) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 2
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
                'Location assigned
                'CREATE CRIBMASTER BIN / SET MIN'S AND MAX'S
                If IsNull(CribRS.Fields("CRIBBIN")) Then
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 1) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 1
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
            End If
            
            'START CODE SPECIFIC TO "STATUS" TYPE [TOOLLIST CHANGE ITEMS]
            If (Trim(sqlRS.Fields("Type")) = "STATUS") Then
                '[TOOLLIST CHANGE MASTER]
                If Trim(sqlRS.Fields("NEWSTATUS")) = "RELEASED" Then
                    'CREATE BINS FOR ALL ITEMS / SET MIN'S AND MAX'S FOR ALL ITEMS
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 10) Then
                        'add to [TOOLLIST CHANGE ACTIONS]
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 10
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                    'ADD JOB TO TOOLBOSS
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 9) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 9
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                ElseIf Trim(sqlRS.Fields("NEWSTATUS")) = "OBSOLETE" Then
                    'REMOVE JOB FROM ALL TOOLBOSS'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 8) Then
                       'add to [TOOLLIST CHANGE ACTIONS]
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 8
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                    'ATTEMPT TO RETURN ALL ITEMS
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 13) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 13
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                ElseIf Trim(sqlRS.Fields("NEWSTATUS")) = "ACTIVE" Then
                    'VERIFY CRIBMASTER MIN'S AND MAX'S
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 2) Then
                        'add to [TOOLLIST CHANGE ACTIONS]
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 2
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                    'ADD JOB TO TOOLBOSS
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 9) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 9
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                    'CREATE BINS FOR ALL ITEMS / SET MIN'S AND MAX'S FOR ALL ITEMS
                    If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 10) Then
                        SQLRS2.AddNew
                        SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                        SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                        SQLRS2.Fields("ACTIONITEM") = 10
                        SQLRS2.Fields("COMPLETE") = 0
                        SQLRS2.Update
                    End If
                End If
            End If
            'START CODE SPECIFIC TO "PLANT" TYPE [TOOLLIST CHANGE ITEMS]
            If (Trim(sqlRS.Fields("Type")) = "PLANT") Then
                'ADD / REMOVE JOB FROM TOOLBOSS BASED ON PLANT CHANGES.
                If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 7) Then
                    'add to [TOOLLIST CHANGE ACTIONS]
                    SQLRS2.AddNew
                    SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                    SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                    SQLRS2.Fields("ACTIONITEM") = 7
                    SQLRS2.Fields("COMPLETE") = 0
                    SQLRS2.Update
                End If
            End If
            'START CODE SPECIFIC TO "VOLUME" TYPE [TOOLLIST CHANGE ITEMS]
            If (Trim(sqlRS.Fields("Type")) = "VOLUME") Then
                'VERIFIY ALL TOOLS HAVE SUFFICIENT MIN'S AND MAX'S
                If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 5) Then
                    'add to [TOOLLIST CHANGE ACTIONS]
                    SQLRS2.AddNew
                    SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                    SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                    SQLRS2.Fields("ACTIONITEM") = 5
                    SQLRS2.Fields("COMPLETE") = 0
                    SQLRS2.Update
                End If
            End If
            'START CODE SPECIFIC TO "STOCK TOOLBOSS" TYPE [TOOLLIST CHANGE ITEMS]
            If (Trim(sqlRS.Fields("Type")) = "STOCK TOOLBOSS") Then
                'STOCK THIS ITEM IN TOOLBOSS. THIS NORMALLY UNSTOCKED ITEM NEEDSTO BE PUT IN THE TOOLBOSS.
                If Not ActionItemExists(sqlRS.Fields("ITEMCHANGEID"), 14) Then
                     'add to [TOOLLIST CHANGE ACTIONS]
                     SQLRS2.AddNew
                     SQLRS2.Fields("ITEMCHANGEID") = sqlRS.Fields("ITEMCHANGEID")
                     SQLRS2.Fields("PROCESSCHANGEID") = sqlRS.Fields("PROCESSCHANGEID")
                     SQLRS2.Fields("ACTIONITEM") = 14
                     SQLRS2.Fields("COMPLETE") = 0
                     SQLRS2.Update
                 End If
            End If
            If 1 = CribRS.State Then
                CribRS.Close 'Item in Crib
            End If
            If 1 = SQLRS3.State Then
                SQLRS3.Close 'Does item have ToolBoss stocked class
            End If
            sqlRS.MoveNext
        Loop
        'CribRS is closed after each loop iteration
        sqlRS.Close '[TOOLLIST CHANGE ITEMS]
        SQLRS2.Close '[TOOLLIST CHANGE ACTIONS] - finished adding records
        'SQLRS3 is closed after each loop iteration
        Set sqlRS = Nothing
        Set SQLRS2 = Nothing
        Set SQLRS3 = Nothing
        Set SQLRS4 = Nothing
        Set sqlCMD = Nothing
   '     ProgressBar.Hide
   '     ProgressBar.Timer1.Enabled = False
End Sub

Function ActionItemExists(IcID As Long, ChangeNumber As Integer) As Boolean
    Dim SQLrs5
    Set SQLrs5 = New ADODB.Recordset
    SQLrs5.Open "SELECT * FROM [ToolList Change Actions] WHERE ITEMCHANGEID = " + Trim(Str(IcID)) + " AND ACTIONITEM = " + Trim(Str(ChangeNumber)), sqlConn, adOpenKeyset, adLockReadOnly
    'MsgBox (SQLrs5.RecordCount)
    If SQLrs5.RecordCount > 0 Then
        ActionItemExists = True
    Else
        ActionItemExists = False
    End If
    SQLrs5.Close
    Set SQLrs5 = Nothing
End Function

