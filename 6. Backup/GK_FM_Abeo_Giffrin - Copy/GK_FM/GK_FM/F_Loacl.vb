Imports System.Diagnostics.Process
Imports System.Threading
Imports System.IO
Public Class F_Loacl
    Dim rowDelete As Integer
    Dim matrixUID As String
    Dim AWBForm As SAPbouiCOM.Form = Nothing
    Dim hawbForm As SAPbouiCOM.Form = Nothing
    Dim mawbForm As SAPbouiCOM.Form = Nothing
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Dim oF_PiecesWeight As F_PiecesWeight
    Dim oform1 As SAPbouiCOM.Form
    Dim oF_AWBParameter As F_AWBParameter
    Public ShowFolderBrowserThread As Threading.Thread
    Dim strpath As String
    Dim FilePath As String
    Dim FileName As String
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub

    Public Sub Local_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.DataTables.Add("PVLO")
            oForm.DataSources.DataTables.Add("DOLO")
            oForm.DataSources.DataTables.Add("BILO") 'REFLO
            oForm.DataSources.DataTables.Add("REFLO") '
            oForm.PaneLevel = 1
            DocNumber_Local()
            oEdit = oForm.Items.Item("SIJ18").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            CFL_Item_Vessel(oForm, SBO_Application)
            oEdit = oForm.Items.Item("SIJ27").Specific
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemCode"
            ooption = oForm.Items.Item("85").Specific
            ooption.GroupWith("86")
            CFL_SalesOrder(oForm, SBO_Application, "LC")
            oEdit = oForm.Items.Item("SIJ4").Specific
            oEdit.ChooseFromListUID = "ORDR"
            oEdit.ChooseFromListAlias = "DocNum"
            oForm.DataBrowser.BrowseBy = "1000004"
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:AE_Job_Bind" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Shared Sub LoadGrid_PV(ByVal oForm As SAPbouiCOM.Form)
        Try
            oGrid = oForm.Items.Item("PVGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum], T0.[U_VOVCode] as 'Vendor Code', T0.[U_VOVName] as 'Vendor Name', T0.[U_VOType] as 'Payment Type', T0.[U_JobNo] as 'Job No', T0.[U_VONo] as 'Voucher No', T0.[U_VODt] as 'Voucher Date', T0.[U_VOTotAmt] as 'Amount' FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 WHERE T0.[U_JobNo]  ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("PVLO").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("PVLO")
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_PV" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_DO(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("DOGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum],T0.[U_JobNo] 'JobNo',  T0.[U_CardCode] 'Card Code', T0.[U_CardName] 'Card Name', T0.[U_VesselNo] 'Vessel', T0.[U_MAWBNo] 'OBL No', T0.[U_TaxDate] as 'Date', T0.[U_ANSRecNo] as 'DO No' FROM [dbo].[@AIGI]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("DOLO").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("DOLO")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
        End Try
    End Sub


    Public Shared Sub LoadGrid_BI(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("BIGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = "'" & oEdit.String & "'"
            oEdit = oForm.Items.Item("1000003").Specific
            If oEdit.String <> "" And oEdit.String <> "NA" Then
                JobNo = JobNo & "," & "'" & oEdit.String & "'"
            End If
            Dim str As String = "SELECT DocEntry 'DocNum','DraftInvoice' DocumentType ,T0.[DocDate] 'BIDate', T0.[CardCode] 'Customer Code', T0.[CardName] 'Customer Name', T0.[U_AB_JobNo] 'Job No', T0.[DocTotal] 'Document Total' FROM ODRF T0 WHERE T0.[ObjType] =13 and  T0.[DocStatus] ='O' and  T0.[U_AB_JobNo] in ( " & JobNo & ") union all SELECT DocEntry 'DocNum','Invoice' DocumentType , T0.[DocDate] 'BIDate', T0.[CardCode] 'Customer Code', T0.[CardName] 'Customer Name', T0.[U_AB_JobNo] 'Job No', T0.[DocTotal] 'Document Total' FROM OINV T0 WHERE   T0.[U_AB_JobNo] in (" & JobNo & ")"

            oForm.DataSources.DataTables.Item("BILO").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("BILO")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
        End Try
    End Sub

    Public Sub DocNumber_Local()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
           
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_LOCAL_HEADER]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "LO" & Format(Now.Date, "yy") & "J" & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "LO" & Format(Now.Date, "yy") & "J" & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "LO" & Format(Now.Date, "yy") & "J" & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "LO" & Format(Now.Date, "yy") & "J" & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "LO" & Format(Now.Date, "yy") & "J" & "0" & DocNum
            ElseIf DocNumLen >= 5 Then
                oEdit.String = "LO" & Format(Now.Date, "yy") & "J" & DocNum
            End If

        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:DocNumber_AI" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Public Sub LoadJobOrder(ByVal SQNO As String)
        Try
            'SELECT T0.[CardCode], T0.[CardName], T0.[U_AB_Divison], T0.[U_AB_TransTo], T0.[U_AB_Trnst], T0.[U_AB_VessCode], T0.[U_AB_VessName], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_Desc], T0.[U_AB_Validity], T0.[U_AB_Ttime], T0.[U_AB_Freq], T0.[U_AB_CARTotQt], T0.[U_AB_CARTotWt] FROM ORDR T0 WHERE T0.[DocNum] ='1'
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.DocNum,T0.[CardCode], T0.[CardName], T1.[Name],  T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_OrginCodeC], T0.[U_AB_OriginNameC], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_DestCodeC], T0.[U_AB_DestNameC],T0.[U_AB_Desc],T0.[U_AB_TotPkg], T0.[U_AB_TotWT], T0.[U_AB_TotVol],T0.[U_AB_Divsion1], T0.[U_AB_JobNo] FROM ORDR T0 Left JOIN OCPR T1 ON T0.CntctCode = T1.CntctCode WHERE  T0.[U_AB_Divsion] ='LC' and isnull( T0.[U_AB_Status] ,'')='Open' and T0.DocNum='" & SQNO & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oEdit = oForm.Items.Item("SIJ4").Specific
            oEdit.String = oRecordSet1.Fields.Item("DocNum").Value
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.String = oRecordSet1.Fields.Item("CardCode").Value
            oEdit = oForm.Items.Item("SIJ8").Specific
            oEdit.String = oRecordSet1.Fields.Item("CardName").Value
            oEdit = oForm.Items.Item("LSIJ157").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OrginCodeC").Value
            oEdit = oForm.Items.Item("156").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OriginNameC").Value
            oEdit = oForm.Items.Item("LSIJ100001").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestCodeC").Value
            oEdit = oForm.Items.Item("40").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestNameC").Value
            oEdit = oForm.Items.Item("34").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_Desc").Value


            Try

                oItem = oForm.Items.Item("lje12")
                oItem.Enabled = True
                oItem = oForm.Items.Item("1000003")
                oItem.Enabled = True
                oEdit = oForm.Items.Item("lje12").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_Divsion1").Value
                End If
                oEdit = oForm.Items.Item("1000003").Specific
                If oRecordSet1.Fields.Item("U_AB_JobNo").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                End If
                oCombo = oForm.Items.Item("1000008").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "IN" Or oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "PR" Then
                    oCombo.Select("Approved", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    oCombo.Select("NA", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                oItem = oForm.Items.Item("lje12")
                oItem.Enabled = False
                oItem = oForm.Items.Item("1000003")
                oItem.Enabled = False

            Catch ex As Exception

            End Try
            'T0.[U_AB_TotPkg], T0.[U_AB_TotWT], T0.[U_AB_TotVol]
            '  oRecordSet1.DoQuery("SELECT sum( T0.[U_Vol]) cwt,Sum(T0.[U_M3]) m3, sum(T0.[U_Wt]) wt ,Sum(T0.[U_PKg]) pkg FROM [dbo].[@AB_SALESORDER_CARGO]  T0 WHERE T0.[U_SONo] ='" & SQNO & "'")
            oEdit = oForm.Items.Item("71").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_TotPkg").Value
            oEdit = oForm.Items.Item("73").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_TotWT").Value
            oEdit = oForm.Items.Item("75").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_TotVol").Value
            oEdit = oForm.Items.Item("SJI10").Specific
            oEdit.String = oRecordSet1.Fields.Item(3).Value

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormUID = "LOC_JOB" Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                    oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    LoadGrid_DO(oForm)
                    LoadGrid_PV(oForm)
                    LoadGrid_BI(oForm)
                End If
            End If

        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:SBO_Application_FormDataEvent" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Shared Sub LoadGrid_REF_ATTACH(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("REFATT").Specific
            oEdit = oForm.Items.Item("1000003").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            If JobNo.Contains("IN") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(2500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            ElseIf JobNo.Contains("PR") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(2500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_PRO_HEADER]  T0 , [dbo].[@AB_PRO_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            End If
            oForm.DataSources.DataTables.Item("REFLO").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("REFLO")
            oGrid.Columns.Item("RowsHeader").Width = 30

        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_REF_ATTACH" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If ((pVal.FormUID = "LOC_JOB" And pVal.ItemUID = "REFATTFOL" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                LoadGrid_REF_ATTACH(oForm) '137
            ElseIf ((pVal.FormUID = "LOC_JOB" And pVal.ItemUID = "REFDIS" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                Dim DocNum As String = ""
                oGrid = oForm.Items.Item("REFATT").Specific
                For F = 0 To oGrid.Rows.Count - 1
                    If oGrid.Rows.IsSelected(F) = True Then
                        DocNum = oGrid.DataTable.GetValue("File Name", F)
                        Exit For
                    End If
                Next
                If DocNum <> "" Then
                    Loadfile(DocNum)
                End If
            End If
            If ((pVal.FormUID = "LOC_JOB" And pVal.ItemUID = "ATTMAT" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                oColumns = oMatrix1.Columns
                Dim i As Integer
                For i = 1 To oMatrix1.RowCount
                    If oMatrix1.IsRowSelected(i) Then
                        oItem = oForm.Items.Item("1000006")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("155")
                        oItem.Enabled = True
                    End If
                Next
            End If
            If ((pVal.FormUID = "LOC_JOB" And pVal.ItemUID = "1000006" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                oColumns = oMatrix1.Columns
                Dim i As Integer
                For i = 1 To oMatrix1.RowCount
                    If oMatrix1.IsRowSelected(i) Then
                        oMatrix1.DeleteRow(i)
                    End If
                Next
                oItem = oForm.Items.Item("1000006")
                oItem.Enabled = False
                oItem = oForm.Items.Item("155")
                oItem.Enabled = False
                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If
            If ((pVal.FormUID = "LOC_JOB" And pVal.ItemUID = "155" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                oColumns = oMatrix1.Columns
                Dim i As Integer
                For i = 1 To oMatrix1.RowCount
                    If oMatrix1.IsRowSelected(i) Then
                        Dim Str As String = ""
                        oEdit = oMatrix1.Columns.Item("V_2").Cells.Item(i).Specific
                        Str = oEdit.String
                        oEdit = oMatrix1.Columns.Item("V_1").Cells.Item(i).Specific
                        Str = Str & "\" & oEdit.String
                        Loadfile(Str)
                    End If
                Next
                oItem = oForm.Items.Item("155")
                oItem.Enabled = False
                oItem = oForm.Items.Item("1000006")
                oItem.Enabled = False
            End If
            If ((pVal.FormUID = "LOC_JOB" And pVal.ItemUID = "1000005" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                Try
                    oForm = SBO_Application.Forms.Item("LOC_JOB")

                    ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
                    If ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Unstarted Then
                        ShowFolderBrowserThread.SetApartmentState(Threading.ApartmentState.STA)
                        ShowFolderBrowserThread.Start()
                        ShowFolderBrowserThread.Join()
                    Else
                        ShowFolderBrowserThread.Abort()
                    End If
                    'MsgBox(strpath)
                    'MsgBox(FilePath)
                    'MsgBox(FileName)
                    oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                    oColumns = oMatrix1.Columns
                    If FileName <> "" Then
                        oMatrix1.AddRow()
                        oMatrix1.ClearRowData(oMatrix1.RowCount)
                        oEdit = oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific
                        oEdit.String = ""
                    End If
                    If FileName <> "" Then
                        oEdit = oMatrix1.Columns.Item("V_2").Cells.Item(oMatrix1.RowCount).Specific
                        If oEdit.String = "" Then
                            oEdit = oMatrix1.Columns.Item("V_2").Cells.Item(oMatrix1.RowCount).Specific
                            oEdit.String = FilePath
                            oEdit = oMatrix1.Columns.Item("V_1").Cells.Item(oMatrix1.RowCount).Specific
                            oEdit.String = FileName
                            oEdit = oMatrix1.Columns.Item("V_0").Cells.Item(oMatrix1.RowCount).Specific
                            oEdit.String = Format(Now.Date, "dd/MM/yy")
                            oEdit = oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific
                            oEdit.String = oMatrix1.RowCount
                        End If
                    End If
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    ' Dim i As Integer = 0
                    '   For i = 1 To oMatrix1.RowCount

                    '  Next

                    ShowFolderBrowserThread.Abort()
                    'End If
                    FileName = ""
                    FilePath = ""
                    strpath = ""
                Catch ex As Exception
                    SBO_Application.MessageBox(ex.Message)
                End Try

            End If
            '-----------------Attachment----------
            '-----------------Load Job------------------
            If pVal.FormUID = "AI_FI_GoodsIssue" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("LOC_JOB")
                If oForm.Visible = True Then
                    LoadGrid_DO(oForm)
                End If
            End If
            If pVal.FormUID = "AB_PV" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("LOC_JOB")
                If oForm.Visible = True Then
                    LoadGrid_PV(oForm)
                End If
            End If
            If pVal.FormType = 133 And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("LOC_JOB")
                If oForm.Visible = True Then
                    LoadGrid_BI(oForm)
                End If
            End If
            '-----------------Local Job------------------
            If pVal.FormUID = "LOC_JOB" Then
                oForm = SBO_Application.Forms.Item("LOC_JOB")
                Try
                    Try
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            oItem = oForm.Items.Item("SIJ16")
                            oItem.Enabled = True
                            oItem = oForm.Items.Item("1000004")
                            oItem.Enabled = True
                        ElseIf pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            oItem = oForm.Items.Item("SIJ16")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("1000004")
                            oItem.Enabled = False
                        End If
                    Catch ex As Exception
                    End Try
                    oItem = oForm.Items.Item("DOButt")
                    If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oItem.Enabled = True Then
                        oItem = oForm.Items.Item("DOButt")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("149")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("153")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("PVButton")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("SIJPV")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("PrintBI")
                        oItem.Enabled = False
                    ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And oItem.Enabled = False Then
                        oItem = oForm.Items.Item("PVButton")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("SIJPV")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("153")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("DOButt")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("149")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("PrintBI")
                        oItem.Enabled = True
                    End If
                Catch ex As Exception
                End Try
                '---------------Item Pressed EVent........................
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "89" Then
                            oForm.PaneLevel = 1
                        ElseIf pVal.ItemUID = "SIJ125VOU" Then
                            oForm.PaneLevel = 4
                        ElseIf pVal.ItemUID = "DO1000001" Then
                            oForm.PaneLevel = 3
                        ElseIf pVal.ItemUID = "ATTACH" Then
                            oForm.PaneLevel = 5
                        ElseIf pVal.ItemUID = "BIFolder" Then
                            oForm.PaneLevel = 6
                        ElseIf pVal.ItemUID = "DOButt" Then
                            LoadDeliveryOrder(oForm)
                        ElseIf pVal.ItemUID = "153" Then
                            oEdit = oForm.Items.Item("SIJ16").Specific
                            LoadDraftInvoice(oEdit.String)
                        ElseIf pVal.ItemUID = "PrintBI" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf BI_Report)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "PVButton" Then
                            LoadPaymentVoucher(oForm)
                        ElseIf pVal.ItemUID = "SIJPV" Then
                            'PaymentVoucher
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf Payment_Vocher)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "149" Then
                            'Delivery Order
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf DOReport)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                DocNumber_Local()
                                oEdit = oForm.Items.Item("SIJ18").Specific
                                oEdit.String = Format(Now.Date, "dd/MM/yy")
                            End If
                        End If
                    ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "1" Then
                            DocNumber_Local()
                        End If
                    ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "1" Then
                            oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                            Dim i As Integer
                            Dim st As String = ""
                            Dim sourcePath As String = ""
                            For i = 1 To oMatrix1.RowCount
                                oEdit = oMatrix1.Columns.Item("V_2").Cells.Item(i).Specific
                                st = oEdit.String
                                oEdit = oMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific
                                oEdit.String = i
                                oEdit = oMatrix1.Columns.Item("V_3").Cells.Item(i).Specific

                                If st <> "" And oEdit.String = "Open" Then
                                    oEdit = oMatrix1.Columns.Item("V_1").Cells.Item(i).Specific
                                    sourcePath = st & "\" & oEdit.String
                                    oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT AttachPath from OADP")
                                    'SELECT AttachPath from OADP
                                    Dim destPath As String = oRecordSet.Fields.Item("AttachPath").Value.ToString
                                    If Not Directory.Exists(destPath) Then
                                        SBO_Application.MessageBox("Destination Path Not Found!")
                                        Exit Sub
                                    End If

                                    Dim getn As Array
                                    ' For Each file__1 As String In Directory.GetFiles(Path.GetDirectoryName(sourcePath))
                                    Dim FileName As String = Path.GetFileNameWithoutExtension(sourcePath) '& Now.ToString("ddMMyyyyhhmmssffff")
                                    Dim FileExten As String = Path.GetExtension(sourcePath)
                                    Dim K As Integer = 1
10:                                 If System.IO.File.Exists(destPath & FileName & FileExten) Then
                                        ' MsgBox("THis Name Existsts")
                                        FileName = Path.GetFileNameWithoutExtension(sourcePath) & "_" & K
                                        K = K + 1
                                        GoTo 10
                                    End If
                                    Dim dest As String = Path.Combine(destPath, FileName & FileExten)
                                    File.Copy(sourcePath, dest, False)
                                    oEdit = oMatrix1.Columns.Item("V_2").Cells.Item(i).Specific
                                    oEdit.String = destPath
                                    oEdit = oMatrix1.Columns.Item("V_1").Cells.Item(i).Specific
                                    oEdit.String = FileName & FileExten
                                    oEdit = oMatrix1.Columns.Item("V_3").Cells.Item(i).Specific
                                    oEdit.String = "Closed"
                                    'Next
                                    'For Each folder As String In Directory.GetDirectories(Path.GetDirectoryName(sourcePath))
                                    '    Dim dest As String = Path.Combine(destPath, Path.GetFileName(folder))
                                    '    CopyDirectory(folder, dest)
                                    'Next
                                End If
                            Next
                        End If

                    ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If pVal.ItemUID = "1" Then
                            oCombo = oForm.Items.Item("1000008").Specific
                            If oCombo.Selected.Value = "Done" Then
                                oEdit = oForm.Items.Item("1000003").Specific
                                Dim BaseJobNo As String = oEdit.String
                                oEdit = oForm.Items.Item("SIJ16").Specific
                                Dim JobNo As String = oEdit.String
                                If BaseJobNo.Substring(0, 2) = "IN" Or BaseJobNo.Substring(0, 2) = "PR" Then
                                    oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet1.DoQuery("UPDATE ODRF SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                                    oRecordSet1.DoQuery("UPDATE OINV SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                                    LoadGrid_BI(oForm)
                                End If
                            End If
                        End If
                    End If
                    '-------------------Validation---------------
                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "SIJ4" Then
                            oEdit = oForm.Items.Item("SIJ4").Specific
                            If oEdit.String <> "" Then
                                LoadJobOrder(oEdit.String)
                            End If
                        End If
                    End If
                    '---------Double Click Event----
                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                    Try
                        'Mitra
                        If pVal.BeforeAction = False And pVal.ItemUID = "BIGrid" Then
                            oGrid = oForm.Items.Item("BIGrid").Specific
                            'oEdit = oForm.Items.Item("64").Specific
                            For F = 0 To oGrid.Rows.Count - 1
                                If oGrid.Rows.IsSelected(F) = True Then
                                    Dim DocNum As String = oGrid.DataTable.GetValue("DocNum", F)
                                    Dim DocType As String = oGrid.DataTable.GetValue("DocumentType", F)
                                    If DocType = "DraftInvoice" Then
                                        oEdit = oForm.Items.Item("BI93").Specific
                                        oEdit.String = DocNum
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        BIAddButt = "Yes"
                                        oItem = oForm.Items.Item("94")
                                        oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked)
                                    ElseIf DocType = "Invoice" Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        SBO_Application.ActivateMenuItem("2053")
                                        oform1 = SBO_Application.Forms.ActiveForm
                                        oform1.Title = "Billing Instruction"
                                        oform1.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        oEdit = oform1.Items.Item("8").Specific
                                        oEdit.String = DocNum
                                        oItem = oform1.Items.Item("1")
                                        oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oItem.Visible = False

                                        ' SBO_Application.StatusBar.SetText("Invoice Can't Be Open", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                            Next
                        End If
                        If pVal.BeforeAction = False And pVal.ItemUID = "PVGrid" Then
                            oGrid = oForm.Items.Item("PVGrid").Specific
                            'oEdit = oForm.Items.Item("64").Specific

                            For F = 0 To oGrid.Rows.Count - 1
                                If oGrid.Rows.IsSelected(F) = True Then
                                    Dim DocNum As String = oGrid.DataTable.GetValue("DocNum", F)
                                    LoadFromXML("PaymentVoucher.srf", SBO_Application)
                                    oForm = SBO_Application.Forms.Item("AB_PV")
                                    oEdit = oForm.Items.Item("22").Specific

                                    oForm.EnableMenu("1282", False)  '// Add New Record
                                    oForm.EnableMenu("1288", False)  '// Next Record
                                    oForm.EnableMenu("1289", False)  '// Pevious Record
                                    oForm.EnableMenu("1290", False)  '// First Record
                                    oForm.EnableMenu("1291", False)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oItem = oForm.Items.Item("22")
                                    oItem.Enabled = True
                                    oEdit.Value = DocNum
                                    oItem = oForm.Items.Item("1")
                                    oEdit = oForm.Items.Item("4").Specific
                                    oEdit.String = oEdit.String
                                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oItem = oForm.Items.Item("22")
                                    oItem.Enabled = False
                                End If
                            Next
                        End If

                        If pVal.BeforeAction = False And pVal.ItemUID = "DOGrid" Then
                            oGrid = oForm.Items.Item("DOGrid").Specific
                            'oEdit = oForm.Items.Item("64").Specific

                            For F = 0 To oGrid.Rows.Count - 1
                                If oGrid.Rows.IsSelected(F) = True Then
                                    Dim DocNum As String = oGrid.DataTable.GetValue("DocNum", F)
                                    LoadFromXML("GoodsIssue.srf", SBO_Application)
                                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
                                    oEdit = oForm.Items.Item("12").Specific

                                    oForm.EnableMenu("1282", False)  '// Add New Record
                                    oForm.EnableMenu("1288", False)  '// Next Record
                                    oForm.EnableMenu("1289", False)  '// Pevious Record
                                    oForm.EnableMenu("1290", False)  '// First Record
                                    oForm.EnableMenu("1291", False)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oItem = oForm.Items.Item("12")
                                    oItem.Enabled = True
                                    oEdit.Value = DocNum
                                    oItem = oForm.Items.Item("1")
                                    oEdit = oForm.Items.Item("20").Specific
                                    oEdit.String = oEdit.String
                                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oItem = oForm.Items.Item("12")
                                    oItem.Enabled = False
                                End If
                            Next
                        End If

                    Catch ex As Exception
                        Functions.WriteLog("Class:F_SE_JobOrder" + " Function:ItemEvent" + " Error Message:" + ex.ToString)
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                    '---------------CFL -----------------------
                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    oForm = SBO_Application.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Try
                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "SIJ27" Then
                                Try
                                    oEdit = oForm.Items.Item("24").Specific
                                    oEdit.String = oDataTable.GetValue("ItemName", 0)
                                    oEdit = oForm.Items.Item("SIJ27").Specific
                                    oEdit.String = oDataTable.GetValue("ItemCode", 0)
                                Catch ex As Exception
                                End Try
                            ElseIf pVal.ItemUID = "SIJ4" Then
                                Try
                                    oEdit = oForm.Items.Item("SIJ4").Specific
                                    oEdit.String = oDataTable.GetValue("DocNum", 0)
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                End If

            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DOReport()
        Try
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("DOGrid").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next

            '  Exit Sub
            'oEdit = oForm.Items.Item("DO4").Specific
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select Delivery Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table

            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(sPath & "\GK_FM\DO_WHMS_AI.rpt")
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = Convert.ToInt32(DocNum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item("@DocKey")
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = Ocompany.Server
            Dim DB As String = Ocompany.CompanyDB
            Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\GK_FM\" & "Pwd.txt", True)
            Dim pwd As String
            pwd = file.ReadLine()

            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next


            Dim RptFrm As MY_Report
            RptFrm = New MY_Report
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            RptFrm.Text = "Delivery Order Report"
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub BI_Report()
        Try
            Dim DocType As String = ""
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("BIGrid").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    DocType = oGrid.DataTable.GetValue("DocumentType", F)
                    Exit For
                End If
            Next
            'AB_RP007_BI_INVOICE

            '  Exit Sub
            'oEdit = oForm.Items.Item("DO4").Specific
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select Billing Instruction No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            SBO_Application.StatusBar.SetText("Retrieving Data!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            If DocType = "Invoice" Then
                cryRpt.Load(sPath & "\GK_FM\AB_RP007_BI_INVOICE.rpt")
            Else
                cryRpt.Load(sPath & "\GK_FM\AB_RP007_BI_DRAFT.rpt")
            End If
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")
            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = Convert.ToInt32(DocNum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item("@DocKey")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = Ocompany.Server
            Dim DB As String = Ocompany.CompanyDB
            Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\GK_FM\" & "Pwd.txt", True)
            Dim pwd As String
            pwd = file.ReadLine()
            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With
            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next
            Dim RptFrm As MY_Report
            RptFrm = New MY_Report
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            RptFrm.Text = "Billing Instruction Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub Payment_Vocher()
        Try

            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("PVGrid").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next

            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Document Number can't be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'SBO_Application.StatusBar.SetText("report is generating")
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim ERRPT As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim objConInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim oLogonInfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim intCounter As Integer
            '  Dim Formula As String
            Dim OneMore As Boolean = False
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(sPath & "\GK_FM\AB_RP005_PV_LC.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@DocKey"
            Dim ParaValue As String = DocNum
            Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
            Dim Para As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim MyArr1 As Array = ParaName.Split(";")
            Dim MyArr2 As Array = ParaValue.Split(";")
            For i As Integer = 0 To MyArr1.Length - 1
                Para.Value = MyArr2(i)
                pvCollection.Add(Para)
                cryRpt.DataDefinition.ParameterFields(MyArr1(i)).ApplyCurrentValues(pvCollection)
            Next
            Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\GK_FM\" & "Pwd.txt", True)
            Dim pwd As String = ""
            pwd = file.ReadLine()
            ConInfo.ConnectionInfo.UserID = "sa"
            ConInfo.ConnectionInfo.Password = pwd
            ConInfo.ConnectionInfo.ServerName = Ocompany.Server
            ConInfo.ConnectionInfo.DatabaseName = Ocompany.CompanyDB
            For intCounter = 0 To cryRpt.Database.Tables.Count - 1
                cryRpt.Database.Tables(intCounter).ApplyLogOnInfo(ConInfo)
            Next
            Dim RptFrm As MY_Report
            RptFrm = New MY_Report
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.Refresh()
            RptFrm.Text = "Payment Voucher Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadDeliveryOrder(ByVal oform As SAPbouiCOM.Form)
        Try
            Dim oform1 As SAPbouiCOM.Form
            LoadFromXML("GoodsIssue.srf", SBO_Application)
            oform1 = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
            oEdit = oform.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            oEdit = oform1.Items.Item("35").Specific
            oEdit.String = JobNo

            oEdit = oform.Items.Item("24").Specific
            Dim VesslNo As String = oEdit.String
            oEdit = oform1.Items.Item("20").Specific
            oEdit.String = VesslNo

            oEdit = oform.Items.Item("SIJ6").Specific
            Dim BPCode As String = oEdit.String
            oEdit = oform1.Items.Item("GI4").Specific
            oEdit.String = BPCode
            oEdit = oform1.Items.Item("6").Specific
            If BPCode <> "" Then
                oEdit.String = BPName(BPCode, Ocompany)
                oEdit = oform1.Items.Item("37").Specific
                oEdit.String = BPName(BPCode, Ocompany)
            End If
            oEdit = oform1.Items.Item("8").Specific
            If BPCode <> "" Then
                oEdit.String = ContactPerson(BPCode, Ocompany)
            End If
            oEdit = oform1.Items.Item("33").Specific
            If BPCode <> "" Then
                oEdit.String = BPAddress(BPCode, Ocompany)
            End If
            Dim oDO As New F_GoodsIssue
            oDO.GoodsIssue_Bind(oform1, SBO_Application)

            'oColumn = oColumns.Item("V_0")
            'oColumn.ChooseFromListUID = "OITM"
            'oColumn.ChooseFromListAlias = "ItemCode"
            oMatrix = oform1.Items.Item("29").Specific
            oCombo = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            ComboLoad_Unit(oform, oCombo)
            ''    End Try

            ''    oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(12).Value
            ''    oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(13).Value
            ''    oEdit = oMatrix.Columns.Item("V_4").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(14).Value
            ''    oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(15).Value
            ''    Try
            ''        oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
            ''        oEdit.String = oRecordSet_GR.Fields.Item(16).Value
            ''    Catch ex As Exception

            ''    End Try

            ''    oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(18).Value
            ''    oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(17).Value
            ''    oEdit = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(20).Value
            ''    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(21).Value

            ''    oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
            ''    If oEdit.String = "" Then
            ''        oEdit.String = oRecordSet_GR.Fields.Item(6).Value
            ''    End If
            ''    oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(7).Value
            ''    oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
            ''    oEdit.String = oRecordSet_GR.Fields.Item(0).Value
            ''    oMatrix.AddRow()
            ''    oRecordSet_GR.MoveNext()
            ''Next

        Catch ex As Exception
            Functions.WriteLog("Class:F_SI_JobOrder" + " Function:LoadDeliveryOrder" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub ComboLoad_Unit(ByVal Oform As SAPbouiCOM.Form, ByVal oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@AIUNIT]  T0")
            Dim it As Integer = 1
            For it = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(it, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            Do While Not oRecordSet.EoF
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Loop
            oRecordSet = Nothing
            GC.Collect()
        Catch ex As Exception
            '  SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadPaymentVoucher(ByVal oform As SAPbouiCOM.Form)
        Try
            Dim oform1 As SAPbouiCOM.Form
            LoadFromXML("PaymentVoucher.srf", SBO_Application)
            oform1 = SBO_Application.Forms.Item("AB_PV")
            oEdit = oform.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            oEdit = oform1.Items.Item("32").Specific
            oEdit.String = JobNo
            Dim oPV As New PaymentVoucher
            oPV.PV_Bind(oform1, SBO_Application, "AI", Ocompany)

        Catch ex As Exception
            Functions.WriteLog("Class:F_SI_JobOrder" + " Function:LoadPaymentVoucher" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Public Sub LoadDraftInvoice(ByVal JobNo As String)
        Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt], T0.[U_VessVoyage], T0.[U_OBL], T0.[U_HBL],T0.U_VessVoyage [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_GrssWt],T0.[U_ItemDesc],T0.[U_ETD]  FROM [dbo].[@AB_LOCAL_HEADER] T0 WHERE T0.[U_JobNo]  ='" & JobNo & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            SBO_Application.ActivateMenuItem("2053")
            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
            ' oform1.Freeze(True)
            oform1.Title = "Billing Instruction"
            oItem = oform1.Items.Item("1")
            oItem.Visible = False

            Try
                Dim oNewItem As SAPbouiCOM.Item
                oNewItem = oform1.Items.Add("ADD", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem = oform1.Items.Item("1")
                oNewItem.Top = oItem.Top
                oNewItem.Height = oItem.Height
                oNewItem.Width = oItem.Width '+ 10
                oNewItem.Left = oItem.Left
                oButton = oNewItem.Specific
                oButton.Caption = "Add BI"
            Catch ex As Exception
            End Try

            oEdit = oform1.Items.Item("4").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_CCode").Value
            oEdit = oform1.Items.Item("54").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_CName").Value
            oEdit = oform1.Items.Item("14").Specific
            Try
                oEdit.String = oRecordSet1.Fields.Item("U_GKBNo").Value
            Catch ex As Exception

            End Try
            Try
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-133, 1)
            Catch ex As Exception
                SBO_Application.ActivateMenuItem("6913")
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-133, 1)
            End Try
            ' oform1.Freeze(True)
            oCombo = oform1.Items.Item("U_AB_Divsion").Specific
            oCombo.Select("LC", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oEdit = oform1.Items.Item("U_AB_JobNo").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_JobNo").Value
            oEdit = oform1.Items.Item("U_AB_OriginNameC").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_LoadPortNC").Value
            oEdit = oform1.Items.Item("U_AB_DestNameC").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_DisPortNC").Value
            ' T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_TotPkg], T0.[U_ChrgWt], T0.[U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_F1] FROM [dbo].[@AB_INT_HEADER]  T0 WHERE T0.[U_JobNo]   ='" & JobNo & "'")
            oEdit = oform1.Items.Item("U_AB_TotPkg").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_TotPkg").Value
            oEdit = oform1.Items.Item("U_AB_TotVol").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_ChrgWt").Value
            oEdit = oform1.Items.Item("U_AB_MAWB").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_OBL").Value
            oEdit = oform1.Items.Item("U_AB_HAWB").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_HBL").Value
            oEdit = oform1.Items.Item("U_AB_FLT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_F1").Value
            oEdit = oform1.Items.Item("U_AB_TotWT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_GrssWt").Value
            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessName").Value
            oEdit = oform1.Items.Item("U_AB_Desc").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_ItemDesc").Value
            oEdit = oform1.Items.Item("U_AB_ETDETA").Specific
            oEdit.String = Format(oRecordSet1.Fields.Item("U_ETD").Value, "dd/MM/yy")
            'U_VessName
            Dim QTNo As String = oRecordSet1.Fields.Item("U_QNo").Value
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[Dscription], T1.[Quantity], T1.[Price], T0.[DocCur],T1.[U_AB_Vendor],T1.[unitMsr],T1.U_AB_Cost,T1.unitMsr,T1.FreeTxt FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.[DocNum] ='" & QTNo & "'")
            'oform1.Freeze(False)
            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
            oCombo = oform1.Items.Item("63").Specific
            oCombo.Select(oRecordSet.Fields.Item("DocCur").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oMatrix2 = oform1.Items.Item("38").Specific
            oColumns = oMatrix2.Columns
            Dim i As Integer = 0
            For i = 1 To oRecordSet.RecordCount
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & oRecordSet.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oMatrix2.AddRow()
                oEdit = oMatrix2.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet.Fields.Item("ItemCode").Value
                oEdit = oMatrix2.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet.Fields.Item("Dscription").Value
                oEdit = oMatrix2.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet.Fields.Item("Quantity").Value
                oEdit = oMatrix2.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet.Fields.Item("Price").Value
                oEdit = oMatrix2.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet.Fields.Item("U_AB_Vendor").Value
                oEdit = oMatrix2.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet.Fields.Item("U_AB_Cost").Value
                Try
                    'T1.[unitMsr]
                    oEdit = oMatrix2.Columns.Item("212").Cells.Item(i).Specific
                    oEdit.String = oRecordSet.Fields.Item("unitMsr").Value
                    oEdit = oMatrix2.Columns.Item("163").Cells.Item(i).Specific
                    oEdit.String = oRecordSet.Fields.Item("FreeTxt").Value
                Catch ex As Exception

                End Try
                oRecordSet.MoveNext()
            Next
            SBO_Application.StatusBar.SetText("Data Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oform1.Freeze(False)
            oform1.Title = "Billing Instruction"
            'SELECT T0.[U_CCode], T0.[U_CName], T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_JobNo], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ETD] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] =''")
        Catch ex As Exception
            oform1.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.MenuUID = "1282" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "LOC_JOB" Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = False
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        DocNumber_Local()
                        oEdit = oForm.Items.Item("SIJ18").Specific
                        oEdit.String = Format(Now.Date, "dd/MM/yy")
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ShowFolderBrowser()
        ' Try
        Dim MyTest As New OpenFileDialog
        Dim MyProcs() As Process
        MyProcs = Process.GetProcessesByName("SAP Business One")
        Dim i As Integer = 0
        If MyProcs.Length >= 1 Then
            'For i As Integer = 0 To MyProcs.Length - 1
            Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
            MyTest.InitialDirectory = "C:\"
            MyTest.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            'MyTest.ShowDialog(MyWindow)

            If MyTest.ShowDialog(MyWindow) = DialogResult.OK Then
                strpath = MyTest.FileName
                FilePath = Path.GetDirectoryName(MyTest.FileName)
                FileName = Path.GetFileName(MyTest.FileName)
                'ShowFolderBrowserThread.Abort()
            Else
                'ShowFolderBrowserThread.Abort()
            End If

            ' Next
        Else
            SBO_Application.MessageBox("No SBO instances found.")
            'ShowFolderBrowserThread.Abort()
        End If
        ShowFolderBrowserThread.Abort()
        'Catch ex As Exception
        '    SBO_Application.MessageBox(ex.Message)
        '    ShowFolderBrowserThread.Abort()
        'End Try
    End Sub
End Class
