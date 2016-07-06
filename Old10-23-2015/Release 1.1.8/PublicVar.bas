Attribute VB_Name = "PublicVar"
Public ProcessID As Integer
Public ToolID As Integer
Public itemID As Integer
Public MiscToolID As Integer
Public RevisionID As Integer
Public sqlConn As adodb.Connection
Public M2MsqlConn As adodb.Connection
Public sqlRS As adodb.Recordset
Public SQLRS2 As adodb.Recordset
Public sqlCMD As adodb.Command
Public oddEvenSort As Integer
Public craxReport As New CRAXDRT.Report
Public craxApp As New CRAXDRT.Application
Public toolexists As Boolean
Public itemexists As Boolean
Public misctoolexists As Boolean
Public revisionexists As Boolean
Public processexists As Boolean
Public CribRS As adodb.Recordset
Public CribConn As adodb.Connection
Public NotificationSent As Boolean
Public NotificationMessage As String
Public NotificationSubject As String
Public NotificationSendTo As String
Public OldItemNumber As String
Public ItemIsUsedElsewhere As Boolean
Public LastToolModified As String
Public PlantOriginal(10) As Integer
Public PlantChange(10) As Integer
Public ExitLoop As Boolean
Public LastToolDescription As String
Public openSQLStatement As String
Public MultiTurret As Boolean
Public Const WM_USER = &H400
Public Const TV_FIRST = &H1100
Public Const TTM_ACTIVATE = (WM_USER + 1)
Public Const TVM_GETTOOLTIPS = (TV_FIRST + 25)
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long




Public Sub Init()
    Set sqlConn = New adodb.Connection
    sqlConn.Open "Provider=sqloledb;" & _
           "Data Source=busche-sql;" & _
           "Initial Catalog=busche toollist;" & _
           "User Id=sa;" & _
           "Password=buschecnc1"
    Set CribConn = New adodb.Connection
    CribConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=\\busche-sql\crbmaster\CRIBMSTR.MDB;"
    Set M2MsqlConn = New adodb.Connection
    M2MsqlConn.Open "Provider=sqloledb;" & _
           "Data Source=busche-sql;" & _
           "Initial Catalog=m2mdata01;" & _
           "User Id=sa;" & _
           "Password=buschecnc1;"
    openSQLStatement = "SELECT * FROM [TOOLLIST MASTER] ORDER BY CUSTOMER"
    InitializeReport
End Sub
Public Sub OpenProcesses()
    Dim itmx2 As ListItem
    Set sqlRS = New adodb.Recordset
    sqlRS.Open openSQLStatement, sqlConn
    OpenProcess.ListView1.ListItems.Clear
    While Not sqlRS.EOF
        Set itmx2 = OpenProcess.ListView1.ListItems.Add(, , sqlRS.Fields("PROCESSID"))
        If Not IsNull(sqlRS.Fields("Customer")) Then
            itmx2.SubItems(1) = Trim(sqlRS.Fields("Customer"))
        End If
        If Not IsNull(sqlRS.Fields("PartFamily")) Then
            itmx2.SubItems(2) = Trim(sqlRS.Fields("PartFamily"))
        End If
        If Not IsNull(sqlRS.Fields("OperationDescription")) Then
            itmx2.SubItems(3) = Trim(sqlRS.Fields("OperationDescription"))
        End If
        If Not IsNull(sqlRS.Fields("OperationNumber")) Then
            itmx2.SubItems(4) = Trim(sqlRS.Fields("OperationNumber"))
        End If
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    OldItemNumber = ""
    NotificationMessage = ""
    NotificationSent = False
    OpenProcess.SortByCustomer
End Sub

Public Sub AddProcess()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[TOOLLIST MASTER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("PartFamily") = " "
    sqlRS.Update
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY PROCESSID DESC", sqlConn, adOpenKeyset, adLockReadOnly
    ProcessID = sqlRS.Fields("ProcessID")
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAllPartNumbers()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM INVCUR,INMAST WHERE INVCUR.FCPARTNO = INMAST.FPARTNO AND (INMAST.FGROUP = 'SACAST' OR INMAST.FGROUP = 'SALFOG' OR INMAST.FGROUP = 'VALADD' OR INMAST.FGROUP = 'BUYIN' OR INMAST.FGROUP = 'CUSTIN') AND INVCUR.FLANYCUR = 1 ORDER BY INVCUR.FCPARTNO", M2MsqlConn
    While Not sqlRS.EOF
        ProcessAttr.AllPartNumbersList.AddItem (sqlRS.Fields("FCPARTNO"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub GetAllPlants()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT LIST]", sqlConn
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
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT LIST]", sqlConn
    While Not sqlRS.EOF
        OpenProcess.PlantListCombo.AddItem (sqlRS.Fields("PLANT"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAllPartsForFilter()
    OpenProcess.PartListCombo.Clear
    OpenProcess.PartListCombo.AddItem ("All")
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT DISTINCT PARTNUMBERS FROM [TOOLLIST PARTNUMBERS]", sqlConn
    While Not sqlRS.EOF
        OpenProcess.PartListCombo.AddItem (sqlRS.Fields("PARTNUMBERS"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAssignedPartNumbers()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(ProcessID), sqlConn
    While Not sqlRS.EOF
        ProcessAttr.SelectedPartsList.AddItem (sqlRS.Fields("PartNumbers"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub GetToolPartNumbers()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOLPARTNUMBER] WHERE TOOLID =" + Str(ToolID), sqlConn
    While Not sqlRS.EOF
        ToolAttr.SelectedPartsList.AddItem (sqlRS.Fields("PartNumber"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(ProcessID), sqlConn
    While Not sqlRS.EOF
        ToolAttr.AllPartNumbersList.AddItem (sqlRS.Fields("PartNumbers"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub GetAvailableToolPartNumbers()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(ProcessID), sqlConn
    While Not sqlRS.EOF
        ToolAttr.AllPartNumbersList.AddItem (sqlRS.Fields("PartNumbers"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetAssignedPlant()
    Dim i As Integer
    'SET ALL ELEMENTS OF THE ARRAY TO 0
    For i = 0 To 9
         PlantOriginal(i) = 0
    Next i
    i = 0
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT] WHERE PROCESSID = " + Str(ProcessID), sqlConn
    While Not sqlRS.EOF
        ProcessAttr.SelectedPlantsList.AddItem (sqlRS.Fields("Plant"))
        PlantOriginal(i) = Val(sqlRS.Fields("Plant"))
        i = i + 1
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub GetProcessDetails()
    Dim i As Integer
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(ProcessID), sqlConn
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
    If Not IsNull(sqlRS.Fields("APPROVED")) Then
        If sqlRS.Fields("Approved") Then
            i = 1
        Else
            i = 0
        End If
        ProcessAttr.ApprovedCheck.Value = i
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
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(ToolID), sqlConn
    ToolAttr.ToolNumberTXT.Text = sqlRS.Fields("ToolNumber")
    ToolAttr.OpDescTXT.Text = sqlRS.Fields("OpDescription")
    While ReportForm.CRViewer1.IsBusy
        DoEvents
    Wend
    Dim test As Boolean
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
    ToolID = sqlRS.Fields("TOOLID")

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
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID), sqlConn
    ItemAttri.ItemGroupTXT.Text = sqlRS.Fields("ToolType")
    ItemAttri.ItemNumberCOMBO.Text = sqlRS.Fields("ToolDescription")
    OldItemNumber = sqlRS.Fields("ToolDescription")
    ItemAttri.ManufacturerTXT.Text = sqlRS.Fields("Manufacturer")
    If Not IsNull(sqlRS.Fields("cribtoolid")) Then
        ItemAttri.CribNumberIDTXT.Text = sqlRS.Fields("CribToolID")
    End If
    ItemAttri.QuantityTXT.Text = sqlRS.Fields("Quantity")
    ItemAttri.CuttingEdgesTXT.Text = sqlRS.Fields("NumberOfCuttingEdges")
    ItemAttri.ToolLifeTXT.Text = sqlRS.Fields("QuantityPerCuttingEdge")
    ItemAttri.AdditionalNotesTXT.Text = sqlRS.Fields("AdditionalNotes")
    ItemAttri.NumofRegrindsTXT.Text = sqlRS.Fields("NumOfRegrinds")
    ItemAttri.ToolLifeRegrindTXT.Text = sqlRS.Fields("QtyPerRegrind")
    GetQty
    
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
End Sub

Public Sub UpdatePartNumbers()
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[ToolList PartNumbers]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Dim i As Integer
    i = 0
    While i < ProcessAttr.SelectedPartsList.ListCount
        sqlRS.AddNew
        sqlRS.Fields("ProcessID") = ProcessID
        sqlRS.Fields("PartNumbers") = Trim(ProcessAttr.SelectedPartsList.List(i))
        sqlRS.Update
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
Public Sub UpdatePlants()
    Dim PlantsChanged As Boolean
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST PLANT] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[ToolList PLANT]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Dim i As Integer
    i = 0
    While i < ProcessAttr.SelectedPlantsList.ListCount
        sqlRS.AddNew
        sqlRS.Fields("ProcessID") = ProcessID
        sqlRS.Fields("Plant") = Trim(ProcessAttr.SelectedPlantsList.List(i))
        sqlRS.Update
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = New adodb.Recordset
    i = 0
    'SET ALL ELEMENTS OF THE ARRAY TO 0
    For i = 0 To 9
         PlantChange(i) = 0
    Next i
    i = 0
    sqlRS.Open "SELECT * FROM [TOOLLIST PLANT] WHERE PROCESSID = " + Str(ProcessID), sqlConn
    While Not sqlRS.EOF
        PlantChange(i) = Val(sqlRS.Fields("Plant"))
        i = i + 1
        sqlRS.MoveNext
    Wend
    i = 0
    Dim j As Integer
    j = o
    While i < 10
        If PlantChange(i) <> PlantOriginal(i) Then
            PlantsChanged = True
        End If
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    
    If PlantsChanged Then
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "This Process used to run at plants: "
        For i = 0 To 9
            If PlantOriginal(i) <> 0 Then
                NotificationMessage = NotificationMessage + Str(PlantOriginal(i)) + ", "
            End If
        Next
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "This Process now runs at plants: "
        For i = 0 To 9
            If PlantChange(i) <> 0 Then
                NotificationMessage = NotificationMessage + Str(PlantChange(i)) + ", "
            End If
        Next
    End If
End Sub


Public Sub UpdateProcessDetails()
    Dim i As Integer
    Set sqlRS = New adodb.Recordset
    NotificationSubject = "Process #" + Str(ProcessID) + " - " + UCase(ProcessAttr.PartFamilyTXT.Text)
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(ProcessID), sqlConn, adOpenKeyset, adLockOptimistic
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
            NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "Process is now OBSOLETE."
        Else
            NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "Process is now ACTIVE."
        End If
    End If
    
    If IsNull(sqlRS.Fields("Approved")) Then
        sqlRS.Fields("approved") = ProcessAttr.ApprovedCheck.Value
    End If
    If sqlRS.Fields("Approved") Then
        i = 1
    Else
        i = 0
    End If
    If i <> ProcessAttr.ApprovedCheck.Value Then
        sqlRS.Fields("Approved") = ProcessAttr.ApprovedCheck.Value
        If ProcessAttr.ApprovedCheck.Value = 1 Then
            NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "Process is now APPROVED."
        Else
            NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "Process is now UNAPPROVED."
        End If
    End If
    sqlRS.Fields("Customer") = UCase(ProcessAttr.CustomerTXT.Text)
    If IsNull(sqlRS.Fields("AnnualVolume")) Then
        
        sqlRS.Fields("AnnualVolume") = Val(ProcessAttr.AnnualVolumeTXT.Text)
    Else
        If sqlRS.Fields("AnnualVolume") <> Val(ProcessAttr.AnnualVolumeTXT.Text) Then
            sqlRS.Fields("AnnualVolume") = Val(ProcessAttr.AnnualVolumeTXT.Text)
            NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + "Tool Usage has changed.."
        End If
    End If
    sqlRS.Fields("MultiTurret") = ProcessAttr.MultiTurretLathe.Value
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    SetMultiTurret
    BuildToolList
    BuildRevList
    BuildMiscList
End Sub

Public Sub UpdateToolDetails()
    Dim newseq As Integer
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(ToolID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("ToolNumber") = Val(ToolAttr.ToolNumberTXT.Text)
    sqlRS.Fields("OpDescription") = UCase(ToolAttr.OpDescTXT.Text)
    sqlRS.Fields("Alternate") = ToolAttr.AlternateCHECK.Value
    sqlRS.Fields("PartSpecific") = ToolAttr.PartNumberCheck.Value
    sqlRS.Fields("AdjustedVolume") = Val(ToolAttr.AdjustedVolume.Text)
    sqlRS.Fields("ToolOrder") = Val(ToolAttr.SequenceTxt.Text)
    sqlRS.Fields("ToolLength") = Val(ToolAttr.ToolLengthOffsetTXT.Text)
    sqlRS.Fields("OffsetNumber") = Val(ToolAttr.OffsetNumberTXT.Text)
    ToolID = sqlRS.Fields("TOOLID")
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
    BuildRevList
    BuildMiscList
End Sub

Public Sub UpdateItemDetails()
    Dim changed As Boolean
    changed = False
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID), sqlConn, adOpenKeyset, adLockOptimistic
    If OldItemNumber <> ItemAttri.ItemNumberCOMBO.Text Then
        ItemIsUsedElsewhere = CheckForOtherUse(OldItemNumber, itemID)
        changed = True
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(ItemAttri.ItemNumberCOMBO.Text) + " has been added to the process."
    End If
    sqlRS.Fields("ToolType") = UCase(ItemAttri.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = ItemAttri.ItemNumberCOMBO.Text
    sqlRS.Fields("CribToolID") = ItemAttri.CribNumberIDTXT.Text
    sqlRS.Fields("Manufacturer") = UCase(ItemAttri.ManufacturerTXT.Text)
    If sqlRS.Fields("Quantity") <> ItemAttri.QuantityTXT.Text And Not changed Then
        changed = True
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(ItemAttri.ItemNumberCOMBO.Text) + " usage has changed."
    End If
    sqlRS.Fields("Quantity") = ItemAttri.QuantityTXT.Text
    If sqlRS.Fields("Consumable") <> ItemAttri.ConsumableCHECK.Value And Not changed Then
        changed = True
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(ItemAttri.ItemNumberCOMBO.Text) + " usage has changed."
    End If
    sqlRS.Fields("Consumable") = ItemAttri.ConsumableCHECK.Value
    If sqlRS.Fields("NumberOfCuttingEdges") <> Val(ItemAttri.CuttingEdgesTXT.Text) And Not changed Then
        changed = True
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(ItemAttri.ItemNumberCOMBO.Text) + " usage has changed."
    End If
    sqlRS.Fields("NumberOfCuttingEdges") = Val(ItemAttri.CuttingEdgesTXT.Text)
    If sqlRS.Fields("QuantityPerCuttingEdge") <> Val(ItemAttri.ToolLifeTXT.Text) And Not changed Then
        changed = True
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(ItemAttri.ItemNumberCOMBO.Text) + " usage has changed."
    End If
    sqlRS.Fields("NumOfRegrinds") = Val(ItemAttri.NumofRegrindsTXT.Text)
    sqlRS.Fields("QtyPerRegrind") = Val(ItemAttri.ToolLifeRegrindTXT.Text)
    sqlRS.Fields("Regrindable") = ItemAttri.RegrindableChk.Value
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(ItemAttri.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(ItemAttri.AdditionalNotesTXT.Text)
    sqlRS.Update
    ToolID = sqlRS.Fields("toolid")
    sqlRS.Close
    Set sqlRS = Nothing
    BuildToolList
    BuildRevList
    BuildMiscList
    OldItemNumber = ""
    ClearItemFields
End Sub

Public Sub BuildToolList()
    ToolList.TreeView1.Nodes.Clear
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT TOOLID, TOOLNUMBER, OPDESCRIPTION FROM [TOOLLIST TOOL] WHERE PROCESSID = " + Str(ProcessID) + " ORDER BY TOOLORDER", sqlConn
    ToolList.TreeView1.Nodes.Add , , "Process" + Trim(Str(ProcessID)), "Process #" + Str(ProcessID)
    ToolList.TreeView1.Nodes.Item("Process" + Trim(Str(ProcessID))).Expanded = True
    While Not sqlRS.EOF
        ToolList.TreeView1.Nodes.Add "Process" + Trim(Str(ProcessID)), tvwChild, "TOOL" + Trim(Str(sqlRS.Fields("TOOLID"))), "TOOL " + Trim(Str(sqlRS.Fields("TOOLNUMBER"))) + " - " + sqlRS.Fields("OPDESCRIPTION")
        If ToolID = sqlRS.Fields("TOOLID") Then
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Expanded = True
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Selected = True
            LastToolDescription = sqlRS.Fields("OpDescription")
        Else
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Expanded = False
            ToolList.TreeView1.Nodes.Item("TOOL" + Trim(Str(sqlRS.Fields("TOOLID")))).Selected = False
        End If
        sqlRS.MoveNext
    Wend
        Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT ITEMID, TOOLID, TOOLTYPE , TOOLDESCRIPTION FROM [TOOLLIST ITEM] WHERE PROCESSID = " + Str(ProcessID), sqlConn
    While Not sqlRS.EOF
        ToolList.TreeView1.Nodes.Add "TOOL" + Trim(Str(sqlRS.Fields("TOOLID"))), tvwChild, "ITEM" + Trim(Str(sqlRS.Fields("ITEMID"))), Trim(sqlRS.Fields("TOOLDESCRIPTION"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub


Public Sub ClearProcessFields()
    ProcessAttr.PartFamilyTXT.Text = ""
    ProcessAttr.OpNumTXT.Text = ""
    ProcessAttr.OpDescTXT.Text = ""
    ProcessAttr.ObsoleteCheck.Value = 0
    ProcessAttr.CustomerTXT.Text = ""
    ProcessAttr.SelectedPartsList.Clear
    ProcessAttr.AllPartNumbersList.Clear
    ProcessAttr.SelectedPlantsList.Clear
    ProcessAttr.AllPlantList.Clear
    ProcessAttr.AnnualVolumeTXT.Text = ""
    ProcessAttr.MultiTurretLathe.Value = 0
End Sub

Public Sub InitializeReport()
    Set craxReport = craxApp.OpenReport("\\busche-sql\m2m\Report Files\toollist.rpt")
    For n = 1 To craxReport.Database.Tables.Count
        craxReport.Database.Tables(n).SetLogOnInfo "busche-sql", "Busche Toollist", "sa", "buschecnc1"
    Next n
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (ProcessID)
End Sub

Public Sub RefreshReport()
    Dim delay As Date
    delay = Time
    While DateAdd("s", 0.75, delay) > Time
        DoEvents
    Wend
    craxReport.DiscardSavedData
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (ProcessID)
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
    OldItemNumber = ""
    PopulateItemList
End Sub

Public Sub AddToolSub()
    Dim newseq As Integer
    Set sqlRS = New adodb.Recordset
    sqlRS.CursorLocation = adUseClient
    sqlRS.Open "[TOOLLIST TOOL]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ProcessID") = ProcessID
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
    ToolID = sqlRS.Fields("TOOLID")
    sqlRS.Close
    Set sqlRS = Nothing
    ReSequenceTools newseq
    UpdateToolPartNumbers
    BuildRevList
    BuildToolList
    BuildMiscList
End Sub

Public Sub UpdateToolPartNumbers()
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST ToolPARTNUMBER] WHERE TOOLID =" + Str(ToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[ToolList TOOLPartNumber]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    Dim i As Integer
    i = 0
    While i < ToolAttr.SelectedPartsList.ListCount
        sqlRS.AddNew
        sqlRS.Fields("TOOLID") = ToolID
        sqlRS.Fields("PartNumber") = Trim(ToolAttr.SelectedPartsList.List(i))
        sqlRS.Update
        i = i + 1
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub DeleteToolSub()
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST TOOL] WHERE TOOLID =" + Str(ToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE TOOLID =" + Str(ToolID), sqlConn, adOpenKeyset, adLockReadOnly
    While Not sqlRS.EOF
        ItemIsUsedElsewhere = CheckForOtherUse(sqlRS.Fields("TOOLDESCRIPTION"), sqlRS.Fields("ITEMID"))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST ITEM] WHERE TOOLID =" + Str(ToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST TOOLPARTNUMBER] WHERE TOOLID =" + Str(ToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
    RefreshReport
End Sub

Public Sub DeleteProcessSub()
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST ITEM] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST PARTNUMBERS] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST REV] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST MISC] WHERE PROCESSID =" + Str(ProcessID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
End Sub

Public Sub AddItemSub()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ToolType") = UCase(ItemAttri.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = ItemAttri.ItemNumberCOMBO.Text
    sqlRS.Fields("ProcessID") = ProcessID
    sqlRS.Fields("ToolID") = ToolID
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
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
    NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(ItemAttri.ItemNumberCOMBO.Text) + " has been added to the process."
    OldItemNumber = ""
End Sub

Public Sub DeleteItemSub()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID = " + Str(itemID), sqlConn, adOpenKeyset, adLockReadOnly
    If sqlRS.RecordCount > 0 Then
        OldItemNumber = sqlRS.Fields("ToolDescription")
    End If
    sqlRS.Close
    Set sqlRS = Nothing
    CheckForOtherUse OldItemNumber, itemID
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE  FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
    RefreshReport
End Sub

Public Sub PopulateDeleteView()
    Dim itmx2 As ListItem
    Set sqlRS = New adodb.Recordset
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
    Set sqlRS = New adodb.Recordset
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
    ItemAttri.ItemNumberCOMBO.Clear
    MiscItem.ItemNumberCOMBO.Clear
    Set CribRS = New adodb.Recordset
    CribRS.Open "SELECT DISTINCT DESCRIPTION1 FROM [INVENTRY] WHERE DESCRIPTION1 <> NULL ORDER BY DESCRIPTION1", CribConn
    While Not CribRS.EOF
        ItemAttri.ItemNumberCOMBO.AddItem CribRS.Fields("DESCRIPTION1")
        MiscItem.ItemNumberCOMBO.AddItem CribRS.Fields("DESCRIPTION1")
        CribRS.MoveNext
    Wend
    CribRS.Close
    Set CribRS = Nothing
End Sub

Public Sub BuildRevList()
    ToolList.TreeView3.Nodes.Clear
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT REVISIONID,REVISION, [REVISION DESCRIPTION],[REVISION DATE],[REVISION BY] FROM [TOOLLIST REV] WHERE PROCESSID = " + Str(ProcessID) + " ORDER BY REVISION", sqlConn
    ToolList.TreeView3.Nodes.Add , , "Process" + Trim(Str(ProcessID)), "Process #" + Str(ProcessID)
    ToolList.TreeView3.Nodes.Item("Process" + Trim(Str(ProcessID))).Expanded = True
    While Not sqlRS.EOF
        ToolList.TreeView3.Nodes.Add "Process" + Trim(Str(ProcessID)), tvwChild, "REV" + Trim(Str(sqlRS.Fields("REVISIONID"))), "REVISION - " + Trim(Str(sqlRS.Fields("REVISION")))
        sqlRS.MoveNext
    Wend
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub BuildMiscList()
    ToolList.TreeView2.Nodes.Clear
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT ITEMID, TOOLDESCRIPTION FROM [TOOLLIST MISC] WHERE PROCESSID = " + Str(ProcessID) + " ORDER BY ITEMID", sqlConn
    ToolList.TreeView2.Nodes.Add , , "Process" + Trim(Str(ProcessID)), "Process #" + Str(ProcessID)
    ToolList.TreeView2.Nodes.Item("Process" + Trim(Str(ProcessID))).Expanded = True
    While Not sqlRS.EOF
        ToolList.TreeView2.Nodes.Add "Process" + Trim(Str(ProcessID)), tvwChild, "MISC" + Trim(Str(sqlRS.Fields("ITEMID"))), Trim(sqlRS.Fields("TOOLDESCRIPTION"))
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
    MiscItem.CuttingEdgesTXT.Text = ""
    MiscItem.ToolLifeTXT.Text = ""
    MiscItem.CribNumberIDTXT.Text = ""
    MiscItem.QtyOnHandTXT.Text = ""
    OldItemNumber = ""
    PopulateItemList
End Sub

Public Sub ClearRevisionFields()
    RevisionForm.RevByTXT.Text = ""
    RevisionForm.RevDate = Date
    RevisionForm.RevDescTXT = ""
    RevisionForm.RevNumTXT = ""
End Sub

Public Sub GetMiscDetails()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID), sqlConn
    MiscItem.ItemGroupTXT.Text = sqlRS.Fields("ToolType")
    MiscItem.ItemNumberCOMBO.Text = sqlRS.Fields("ToolDescription")
    OldItemNumber = sqlRS.Fields("ToolDescription")
    MiscItem.ManufacturerTXT.Text = sqlRS.Fields("Manufacturer")
    If Not IsNull(sqlRS.Fields("cribtoolid")) Then
        MiscItem.CribNumberIDTXT.Text = sqlRS.Fields("CribToolID")
    End If
    MiscItem.QuantityTXT.Text = sqlRS.Fields("Quantity")
    MiscItem.CuttingEdgesTXT.Text = sqlRS.Fields("NumberOfCuttingEdges")
    MiscItem.ToolLifeTXT.Text = sqlRS.Fields("QuantityPerCuttingEdge")
    MiscItem.AdditionalNotesTXT.Text = sqlRS.Fields("AdditionalNotes")
    MiscQty
    If sqlRS.Fields("Consumable") Then
        i = 1
    Else
        i = 0
    End If
    MiscItem.ConsumableCHECK.Value = i
    sqlRS.Close
    Set sqlRS = Nothing

End Sub

Public Sub GetRevisionDetails()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST REV] WHERE REVISIONID =" + Str(RevisionID), sqlConn
    RevisionForm.RevByTXT.Text = sqlRS.Fields("Revision By")
    RevisionForm.RevNumTXT.Text = sqlRS.Fields("Revision")
    RevisionForm.RevDescTXT.Text = sqlRS.Fields("Revision Description")
    RevisionForm.RevDate = sqlRS.Fields("Revision Date")
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub UpdateRevisionDetails()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST REV] WHERE REVISIONID =" + Str(RevisionID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("Revision By") = UCase(RevisionForm.RevByTXT.Text)
    sqlRS.Fields("Revision") = UCase(RevisionForm.RevNumTXT.Text)
    sqlRS.Fields("Revision Description") = UCase(RevisionForm.RevDescTXT.Text)
    sqlRS.Fields("Revision Date") = RevisionForm.RevDate
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildToolList
    BuildRevList
    BuildMiscList
End Sub

Public Sub UpdateMiscDetails()
    Dim changed As Boolean
    changed = False
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE ITEMID =" + Str(itemID), sqlConn, adOpenKeyset, adLockOptimistic
    If OldItemNumber <> MiscItem.ItemNumberCOMBO.Text Then
        ItemIsUsedElsewhere = CheckForOtherUse(OldItemNumber, itemID)
        changed = True
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(MiscItem.ItemNumberCOMBO.Text) + " has been added to the process."
    End If
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID), sqlConn, adOpenKeyset, adLockOptimistic
    sqlRS.Fields("ToolType") = UCase(MiscItem.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = MiscItem.ItemNumberCOMBO.Text
    sqlRS.Fields("ProcessID") = ProcessID
    sqlRS.Fields("CribToolID") = MiscItem.CribNumberIDTXT.Text
    sqlRS.Fields("Consumable") = MiscItem.ConsumableCHECK.Value
    sqlRS.Fields("Manufacturer") = UCase(MiscItem.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = MiscItem.QuantityTXT.Text
    sqlRS.Fields("NumberOfCuttingEdges") = Val(MiscItem.CuttingEdgesTXT.Text)
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(MiscItem.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(MiscItem.AdditionalNotesTXT.Text)
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildToolList
    BuildRevList
    BuildMiscList
End Sub

Public Sub DeleteMiscSub()
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST MISC] WHERE ITEMID =" + Str(MiscToolID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
    RefreshReport
End Sub

Public Sub DeleteRevisionSub()
    Set sqlCMD = New adodb.Command
    sqlCMD.CommandText = "DELETE FROM [TOOLLIST REV] WHERE REVISIONID =" + Str(RevisionID)
    sqlCMD.ActiveConnection = sqlConn
    sqlCMD.Execute
    Set sqlCMD = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
    RefreshReport
End Sub

Public Sub AddMiscSub()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[TOOLLIST MISC]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("ToolType") = UCase(MiscItem.ItemGroupTXT.Text)
    sqlRS.Fields("ToolDescription") = MiscItem.ItemNumberCOMBO.Text
    sqlRS.Fields("ProcessID") = ProcessID
    sqlRS.Fields("CribToolID") = MiscItem.CribNumberIDTXT.Text
    sqlRS.Fields("Consumable") = MiscItem.ConsumableCHECK.Value
    sqlRS.Fields("Manufacturer") = UCase(MiscItem.ManufacturerTXT.Text)
    sqlRS.Fields("Quantity") = MiscItem.QuantityTXT.Text
    sqlRS.Fields("NumberOfCuttingEdges") = Val(MiscItem.CuttingEdgesTXT.Text)
    sqlRS.Fields("QuantityPerCuttingEdge") = Val(MiscItem.ToolLifeTXT.Text)
    sqlRS.Fields("AdditionalNotes") = UCase(MiscItem.AdditionalNotesTXT.Text)
    NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + UCase(MiscItem.ItemNumberCOMBO.Text) + " has been added to the process."
    OldItemNumber = ""
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
End Sub

Public Sub AddRevisionSub()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "[TOOLLIST REV]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    sqlRS.AddNew
    sqlRS.Fields("Revision By") = UCase(RevisionForm.RevByTXT.Text)
    sqlRS.Fields("Revision") = UCase(RevisionForm.RevNumTXT.Text)
    sqlRS.Fields("Revision Description") = UCase(RevisionForm.RevDescTXT.Text)
    sqlRS.Fields("Revision Date") = RevisionForm.RevDate
    sqlRS.Fields("ProcessID") = ProcessID
    sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
    BuildRevList
    BuildToolList
    BuildMiscList
End Sub

Public Sub SendEmail()
    Dim test As OSSMTP.SMTPSession
    Set test = New OSSMTP.SMTPSession
    test.MessageSubject = NotificationSubject
    test.MessageText = NotificationMessage
    test.AuthenticationType = AuthNone
    test.Server = "10.1.2.13"
    test.SendTo = NotificationSendTo
    test.MailFrom = "processchange@busche-cnc.com"
    test.SendEmail
    NotificationSent = True
End Sub

Function ValidateItemNumber() As Boolean
    Set CribRS = New adodb.Recordset
    CribRS.Open "SELECT * FROM [INVENTRY] WHERE DESCRIPTION1 = '" + ItemAttri.ItemNumberCOMBO.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
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
        ValidateItemNumber = True
    Else
        ItemAttri.ItemGroupTXT.Text = ""
        ItemAttri.ManufacturerTXT.Text = ""
        ItemAttri.ItemNumberCOMBO.Text = ""
        MsgBox ("Invalid Item Number")
        ValidateItemNumber = False
    End If
    GetQty
End Function

Public Sub ValidateMiscItemNumber()
    Set CribRS = New adodb.Recordset
    CribRS.Open "SELECT * FROM [INVENTRY] WHERE DESCRIPTION1 = '" + MiscItem.ItemNumberCOMBO.Text + "'", CribConn, adOpenKeyset, adLockReadOnly
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
        
    Else
        MiscItem.ItemGroupTXT.Text = ""
        MiscItem.ManufacturerTXT.Text = ""
        MiscItem.ItemNumberCOMBO.Text = ""
        MsgBox ("Invalid Item Number")
    End If
    MiscQty
End Sub
Public Sub MiscQty()
    Dim sum As Integer
    Set CribRS = New adodb.Recordset
    CribRS.Open "SELECT ITEM, BIN, QUANTITY FROM STATION WHERE ITEM = '" + ItemAttri.CribNumberIDTXT.Text + "' OR ITEM = '" + MiscItem.CribNumberIDTXT.Text + "R'", CribConn, adOpenKeyset, adLockReadOnly
    
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
End Sub

Public Sub GetQty()
    Dim sum As Integer
    Set CribRS = New adodb.Recordset
    CribRS.Open "SELECT ITEM, BIN, QUANTITY FROM STATION WHERE ITEM = '" + ItemAttri.CribNumberIDTXT.Text + "' OR ITEM = '" + ItemAttri.CribNumberIDTXT.Text + "R'", CribConn, adOpenKeyset, adLockReadOnly
    
    If CribRS.RecordCount > 0 Then
        While Not CribRS.EOF
            sum = sum + CribRS.Fields("quantity")
            CribRS.MoveNext
        Wend
        ItemAttri.QtyOnHandTXT.Text = sum
    Else
        ItemAttri.QtyOnHandTXT.Text = 0
    End If
    CribRS.Close
    Set CribRS = Nothing
End Sub

Public Function CheckForOtherUse(ItemNumber As String, itemID2 As Integer) As Boolean
    Set SQLRS2 = New adodb.Recordset
    SQLRS2.Open "SELECT [TOOLLIST ITEM].TOOLDESCRIPTION,[TOOLLIST MASTER].PROCESSID,[TOOLLIST MASTER].CUSTOMER, [TOOLLIST MASTER].PARTFAMILY FROM [TOOLLIST ITEM] " & _
     "INNER JOIN [TOOLLIST MASTER] ON [TOOLLIST ITEM].PROCESSID = [TOOLLIST MASTER].PROCESSID " & _
     "WHERE [TOOLLIST ITEM].TOOLDESCRIPTION = '" + ItemNumber + "' AND [TOOLLIST MASTER].OBSOLETE = 0 And [TOOLLIST ITEM].ITEMID <> " + Str(itemID2), sqlConn, adOpenKeyset, adLockReadOnly
    If SQLRS2.RecordCount > 0 Then
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + ItemNumber + " has been removed from this tool list, however still is an active tool on the following jobs:"
        While Not SQLRS2.EOF
            NotificationMessage = NotificationMessage + vbCrLf + Str(SQLRS2.Fields("ProcessiD")) + " - " + SQLRS2.Fields("Customer") + " - " + SQLRS2.Fields("PartFamily")
            SQLRS2.MoveNext
        Wend
    Else
        NotificationMessage = NotificationMessage + vbCrLf + vbCrLf + ItemNumber + " needs to be dispositioned."
    End If
    SQLRS2.Close
    Set SQLRS2 = Nothing
End Function

Public Sub GetEmails()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST EMAIL]", sqlConn, adOpenKeyset, adLockReadOnly
    If sqlRS.RecordCount > 0 Then
        EmailForm.EmailDept.Text = sqlRS.Fields("deptemail")
        EmailForm.EmailBuyer.Text = sqlRS.Fields("buyeremail")
        EmailForm.Email1.Text = sqlRS.Fields("engemail1")
        EmailForm.Email2.Text = sqlRS.Fields("engemail2")
        EmailForm.Email3.Text = sqlRS.Fields("engemail3")
        EmailForm.Email4.Text = sqlRS.Fields("engemail4")
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Sub

Public Sub UpdateEmails()
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST EMAIL]", sqlConn, adOpenKeyset, adLockOptimistic
        sqlRS.Fields("deptemail") = EmailForm.EmailDept.Text
        sqlRS.Fields("buyeremail") = EmailForm.EmailBuyer.Text
        sqlRS.Fields("engemail1") = EmailForm.Email1.Text
        sqlRS.Fields("engemail2") = EmailForm.Email2.Text
        sqlRS.Fields("engemail3") = EmailForm.Email3.Text
        sqlRS.Fields("engemail4") = EmailForm.Email4.Text
        sqlRS.Update
    sqlRS.Close
    Set sqlRS = Nothing
End Sub
    
Public Sub GetSendTo()
    NotificationSendTo = ""
    Set sqlRS = New adodb.Recordset
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
    Set sqlRS = New adodb.Recordset
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
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(ProcessID) + " ORDER BY TOOLORDER ", sqlConn, adOpenKeyset, adLockReadOnly
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
Function GetNextSequence() As Integer
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(ProcessID) + " ORDER BY TOOLORDER ", sqlConn, adOpenKeyset, adLockReadOnly
    If Not sqlRS.EOF Then
        sqlRS.MoveLast
        GetNextSequence = sqlRS.Fields("ToolOrder") + 1
    Else
        GetNextSequence = 1
    End If
    sqlRS.Close
    Set sqlRS = Nothing
End Function
Function ReSequenceTools(CurSequence As Integer)
    Set sqlRS = New adodb.Recordset
    sqlRS.CursorLocation = adUseClient
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID =" + Str(ProcessID) + " AND TOOLORDER >= " + Str(CurSequence) + " AND TOOLID <> " + Str(ToolID) + " ORDER BY TOOLORDER", sqlConn, adOpenDynamic, adLockOptimistic
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
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID =" + Str(ProcessID), sqlConn, adOpenKeyset, adLockOptimistic
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



Private Sub CopyProcess()
    Dim PROCESSID2 As Integer
    Set sqlRS = New adodb.Recordset
    sqlRS.Open "SELECT * FROM [TOOLLIST MASTER] WHERE PROCESSID = " + Str(ProcessID), sqlConn, adOpenKeyset, adLockReadOnly
    
    Set SQLRS2 = New adodb.Recordset
    SQLRS2.Open "[TOOLLIST MASTER]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    
    SQLRS2.AddNew
    SQLRS2.Fields("PartFamily") = sqlRS.Fields("PartFamily")
    SQLRS2.Fields("Customer") = sqlRS.Fields("Customer")
    SQLRS2.Fields("OperationNumber") = sqlRS.Fields("OperationNumber")
    SQLRS2.Fields("OperationDescription") = sqlRS.Fields("OperationDescription")
    SQLRS2.Fields("Obsolete") = sqlRS.Fields("Obsolete")
    SQLRS2.Fields("AnnualVolume") = sqlRS.Fields("AnnualVolume")
    SQLRS2.Fields("Approved") = sqlRS.Fields("Approved")
    SQLRS2.Fields("MultiTurret") = sqlRS.Fields("MultiTurret")
    SQLRS2.Update
    SQLRS2.Close
    
    SQLRS2.Open "SELECT * FROM [TOOLLIST MASTER] ORDER BY PROCESSID DESC", sqlConn, adOpenKeyset, adLockReadOnly
    PROCESSID2 = SQLRS2.Fields("ProcessID")
    SQLRS2.Close
    
    sqlRS.Close
    
    sqlRS.Open "SELECT * FROM [TOOLLIST TOOL] WHERE PROCESSID = " + Str(ProcessID), sqlConn, adOpenKeyset, adLockReadOnly
    
    While Not sqlRS.EOF
    SQLRS2.AddNew
    SQLRS2.Fields("ProcessID") = ProcessID
    SQLRS2.Fields("ToolNumber") = tNum
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
        ToolID = sqlRS.Fields("TOOLID")
    sqlRS.Close
    sqlRS.Open "SELECT * FROM [TOOLLIST ITEM] WHERE PROCESSID = " + Str(pID) + " AND TOOLID = " + Str(tID), sqlConn, adOpenKeyset, adLockReadOnly
    SQLRS2.Open "[TOOLLIST ITEM]", sqlConn, adOpenKeyset, adLockOptimistic, adCmdTable
    While Not sqlRS.EOF
        SQLRS2.AddNew
        SQLRS2.Fields("ProcessID") = ProcessID
        SQLRS2.Fields("ToolID") = ToolID
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
        sqlRS.MoveNext
    Wend
End Sub

