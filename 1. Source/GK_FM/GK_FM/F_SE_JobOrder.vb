Imports System.Diagnostics.Process
Imports System.Threading
Imports System.IO
Public Class F_SE_JobOrder
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Dim oform1 As SAPbouiCOM.Form
    Public ShowFolderBrowserThread As Threading.Thread
    Dim strpath As String
    Dim FilePath As String
    Dim FileName As String
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Public Sub SE_Job_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.DataTables.Add("DOSE")
            oForm.DataSources.DataTables.Add("OBL")
            oForm.DataSources.DataTables.Add("PVSE")
            oForm.DataSources.DataTables.Add("BILO") 'REFSE
            oForm.DataSources.DataTables.Add("REFSE") 'REFSE
            oForm.PaneLevel = 1
            DocNumber_SE()
            ShippigNameLoad()
            oItem = oForm.Items.Item("153")
            oItem.Enabled = False
            oEdit = oForm.Items.Item("SIJ18").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            ooption = oForm.Items.Item("28").Specific
            ooption.GroupWith("SIJ1000001")
            ooption = oForm.Items.Item("85").Specific
            ooption.GroupWith("86")
            '---------
            ooption = oForm.Items.Item("1000010").Specific
            ooption.GroupWith("1000011")
            'ooption = oForm.Items.Item("171").Specific
            'ooption.GroupWith("172")
            '------------
            'oCombo = oForm.Items.Item("42").Specific
            'ComboLoad_Whsc(oForm, oCombo)
            oCombo = oForm.Items.Item("30").Specific
            ComboLoad_ContainerType(oForm, oCombo)
            ''oCombo = oForm.Items.Item("1000002").Specific
            ''ComboLoad_Currency(oForm, oCombo)
            ''oCombo = oForm.Items.Item("1000003").Specific
            ''oCombo.ValidValues.Add("Cheque", "C")
            ''oCombo.ValidValues.Add("Cash", "Ch")
            ''oCombo.ValidValues.Add("CC", "CC")
            ''oCombo.ValidValues.Add("Online", "O")


            ' ''-----DO
            ''CFL_BP_Supplier2(oForm, SBO_Application)
            ''oMatrix = oForm.Items.Item("SIJDOMAT").Specific
            ''oColumns = oMatrix.Columns
            ''oMatrix.AddRow()
            ''oColumn = oColumns.Item("V_13")
            ''oColumn.ChooseFromListUID = "CFLBPV1"
            ''oColumn.ChooseFromListAlias = "CardCode"
            ''CFL_Item(oForm, SBO_Application)
            ''oColumn = oColumns.Item("V_15")
            ''oColumn.ChooseFromListUID = "OITM"
            ''oColumn.ChooseFromListAlias = "ItemCode"


            ' ''------------VO
            ''CFL_Item1(oForm, SBO_Application)
            ''oMatrix1 = oForm.Items.Item("148").Specific
            ''oColumns = oMatrix1.Columns
            ''oMatrix1.AddRow()
            ''oColumn = oColumns.Item("V_8")
            ''oColumn.ChooseFromListUID = "1OITM"
            ''oColumn.ChooseFromListAlias = "ItemCode"
            ' ''---- goods Receipt

            '------------Container Matrix
            oMatrix4 = oForm.Items.Item("SEContMat").Specific
            oColumns = oMatrix4.Columns
            oMatrix4.AddRow(7)


            'oForm.DataSources.UserDataSources.Add("V_0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oMatrix3 = oForm.Items.Item("SEJGR").Specific
            'oColumns = oMatrix3.Columns
            'oColumn = oColumns.Item("V_0")
            'oColumn.DataBind.SetBound(True, "", "V_0")
            'oItem = oForm.Items.Item("SEJGR")
            'oItem.Width = 150
            'oItem.Height = 30
            'oColumn.Width = 130
            'oMatrix3.AddRow()

            'CFL
            CFL_Item_Vessel(oForm, SBO_Application)
            oEdit = oForm.Items.Item("SIJ27").Specific
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemCode"

            CFL_BP_Customer(oForm, SBO_Application)
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.ChooseFromListUID = "CFLBPC"
            oEdit.ChooseFromListAlias = "CardCode"

            'CFL_BP_Supplier(oForm, SBO_Application)
            'oEdit = oForm.Items.Item("127").Specific
            'oEdit.ChooseFromListUID = "CFLBPV"
            'oEdit.ChooseFromListAlias = "CardCode"

            CFL_SalesOrder(oForm, SBO_Application, "SE")
            oEdit = oForm.Items.Item("SEJ4").Specific
            oEdit.ChooseFromListUID = "ORDR"
            oEdit.ChooseFromListAlias = "DocNum"
            oEdit = oForm.Items.Item("SEJ4").Specific
            oEdit.String = ""

            '  ComboLoad_PaymentType(oForm, oCombo)
            oForm.DataBrowser.BrowseBy = "1000004"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#Region "ComboLoad"
    Private Sub ComboLoad_City(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM OCRY T0 order by T0.COde")
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ComboLoad_Currency(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[CurrCode], T0.[CurrName] FROM OCRN T0 ORDER BY T0.[CurrCode]")
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ComboLoad_ContainerType(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM [@AB_SEAI_CONTTYPE] T0 order by T0.COde")
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ComboLoad_Whsc(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 ORDER BY T0.[WhsCode]")
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ComboLoad_PaymentType(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 ORDER BY T0.[WhsCode]")
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub DocNumber_SE()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_SEAE_JOB_H]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "SE" & Format(Now.Date, "yy") & "J" & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "SE" & Format(Now.Date, "yy") & "J" & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "SE" & Format(Now.Date, "yy") & "J" & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "SE" & Format(Now.Date, "yy") & "J" & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "SE" & Format(Now.Date, "yy") & "J" & "0" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = "SE" & Format(Now.Date, "yy") & "J" & DocNum
            End If
          


        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub

    Public Sub DocNumber_OBL(ByVal oform As SAPbouiCOM.Form, ByVal POD As String)
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_SEAE_OBL]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oform.Items.Item("4").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "GH" & POD & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "GH" & POD & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "GH" & POD & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "GH" & POD & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "GH" & POD & "0" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = "GH" & POD & DocNum
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Public Sub ShippigNameLoad()
        Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[CompnyName],T1.[Street],T1.[ZipCode], (select top 1 name from OCRY where Code=T1.[Country]) 'Country' FROM OADM T0 , ADM1 T1")
            oEdit = oForm.Items.Item("94").Specific
            Dim str As String = oRecordSet1.Fields.Item(0).Value.ToString.Trim & vbCrLf & oRecordSet1.Fields.Item(1).Value.ToString.Trim & vbCrLf & oRecordSet1.Fields.Item(3).Value.ToString.Trim & " " & oRecordSet1.Fields.Item(2).Value.ToString.Trim
            oEdit.String = str.ToUpper
            'oEdit = oForm.Items.Item("101").Specific
            'oEdit.String = oRecordSet1.Fields.Item(1).Value.ToString
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
#End Region

    Public Shared Sub LoadGrid_PV(ByVal oForm As SAPbouiCOM.Form)
        Try
            oGrid = oForm.Items.Item("PVGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum], T0.[U_VOVCode] as 'Vendor Code', T0.[U_VOVName] as 'Vendor Name', T0.[U_VOType] as 'Payment Type', T0.[U_JobNo] as 'Job No', T0.[U_VONo] as 'Voucher No', T0.[U_VODt] as 'Voucher Date', T0.[U_VOTotAmt] as 'Amount' FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 WHERE T0.[U_JobNo]  ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("PVSE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("PVSE")
            'oGrid.DataTable = Nothing
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_PV" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    'Public Shared Sub LoadGrid_PV(ByVal oForm As SAPbouiCOM.Form)
    '    Try



    '        oGrid = oForm.Items.Item("PVGrid").Specific
    '        oEdit = oForm.Items.Item("SIJ16").Specific
    '        Dim JobNo As String = oEdit.String
    '        Dim str As String = "SELECT T0.[DocNum], T0.[U_VOVCode] as 'Vendor Code', T0.[U_VOVName] as 'Vendor Name', T0.[U_VOType] as 'Payment Type', T0.[U_JobNo] as 'Job No', T0.[U_VONo] as 'Voucher No', T0.[U_VODt] as 'Voucher Date', T0.[U_VOTotAmt] as 'Amount' FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 WHERE T0.[U_JobNo]  ='" & JobNo & "'"
    '        oForm.DataSources.DataTables.Item("PVSE").ExecuteQuery(str)
    '        oGrid.DataTable = oForm.DataSources.DataTables.Item("PVSE")
    '    Catch ex As Exception
    '        Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_PV" + " Error Message:" + ex.ToString)
    '    End Try
    'End Sub
    Public Shared Sub LoadGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            'OBLGrd
            oGrid = oForm.Items.Item("OBLGrd").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum], T0.[U_JobNo] as 'JobNo', T0.[U_BLno] as 'BL No', T0.[U_OBLNo] as 'OBL No', T0.[U_Shipper] as 'Shipper', T0.[U_Cons] as 'Consignee', T0.[U_PortDel] as 'Port of Delivery', T0.[U_POD] as 'Port of Discharge' FROM [dbo].[@AB_SEAE_OBL]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "' order By DocNum"
            oForm.DataSources.DataTables.Item("OBL").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("OBL")
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_BI(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("BIGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = "'" & oEdit.String & "'"
            oEdit = oForm.Items.Item("137").Specific
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
    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormUID = "SEAE_JOB" Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                    oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    LoadGrid(oForm)
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
            oEdit = oForm.Items.Item("137").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            If JobNo.Contains("IN") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            ElseIf JobNo.Contains("PR") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_PRO_HEADER]  T0 , [dbo].[@AB_PRO_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            End If
            oForm.DataSources.DataTables.Item("REFSE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("REFSE")
            oGrid.Columns.Item("RowsHeader").Width = 30
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_REF_ATTACH" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try


            If ((pVal.FormUID = "SEAE_JOB" And pVal.ItemUID = "REFATTFOL" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                LoadGrid_REF_ATTACH(oForm) '137
            ElseIf ((pVal.FormUID = "SEAE_JOB" And pVal.ItemUID = "REFDIS" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "SEAE_JOB" And pVal.ItemUID = "ATTMAT" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "SEAE_JOB" And pVal.ItemUID = "1000006" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "SEAE_JOB" And pVal.ItemUID = "155" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "SEAE_JOB" And pVal.ItemUID = "1000005" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                Try
                    oForm = SBO_Application.Forms.Item("SEAE_JOB")

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
            If pVal.FormUID = "AI_FI_GoodsIssue" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("SEAE_JOB")
                If oForm.Visible = True Then
                    LoadGrid_DO(oForm)
                End If
            End If
            If pVal.FormType = 133 And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("SEAE_JOB")
                If oForm.Visible = True Then
                    LoadGrid_BI(oForm)
                End If
            End If
            If pVal.FormUID = "AB_PV" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("SEAE_JOB")
                If oForm.Visible = True Then
                    LoadGrid_PV(oForm)

                End If
            End If
            ''-----------Inovice Draft------------
            'If pVal.FormType = 133 Then
            '    If pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
            '        If pVal.ItemUID = "ADD" Then
            '            Try
            '                If SBO_Application.MessageBox("Would you Like to Create PO?", 1, "Yes", "No") = 1 Then
            '                    Try
            '                        CreatePO()
            '                    Catch ex As Exception
            '                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '                    End Try
            '                End If
            '                SBO_Application.ActivateMenuItem("5907")
            '            Catch ex As Exception
            '                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '            End Try
            '        End If
            '    End If
            'End If

            '----------load marix
            '------laod matrix
            If pVal.FormType = 2000080 Then
                'If (pVal.ItemUID = "1" And pVal.Before_Action = False And pVal.InnerEvent = False And SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Or (pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                If (pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) Then
                    oForm = SBO_Application.Forms.Item("SEAE_JOB")
                    'oMatrix1 = oForm.Items.Item("1000001").Specific
                    'oEdit = oMatrix1.Columns.Item("V_0").Cells.Item(1).Specific
                    'oEdit.String = ""
                    'oEdit = oForm.Items.Item("8").Specific
                    'oEdit.String = oEdit.String
                End If
                If (pVal.Before_Action = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) Then
                    'MsgBox("Hi")
                    Try
                        oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        oMatrix2 = oForm.Items.Item("4").Specific
                        Dim i As Integer
                        Dim No(oMatrix2.RowCount) As String
                        Dim VenCode(oMatrix2.RowCount) As String
                        Dim PO(oMatrix2.RowCount) As String
                        Dim k As Integer = 0
                        For i = 1 To oMatrix2.RowCount
                            If oMatrix2.IsRowSelected(i) = True Then
                                oEdit = oMatrix2.Columns.Item("COL1").Cells.Item(i).Specific
                                No(k) = oEdit.String
                                oEdit = oMatrix2.Columns.Item("COL9").Cells.Item(i).Specific
                                VenCode(k) = oEdit.String
                                oEdit = oMatrix2.Columns.Item("COL4").Cells.Item(i).Specific
                                PO(k) = oEdit.String
                                k = k + 1
                            End If

                        Next
                        'k = omatrix2.RowCount
                        oForm.Visible = False

                        For i = 0 To k + 1
                            Try
                                If No(i) <> "" Then
                                    Try
                                        oForm.Freeze(True)
                                        MatrixLoad(No(i), VenCode(i), PO(i))
                                        oForm.Freeze(False)
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try

                                End If
                            Catch ex As Exception
                            End Try

                        Next
                    Catch ex As Exception

                    End Try
                End If
            End If

            If pVal.FormUID = "SEAE_OBL" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("SEAE_JOB")
                LoadGrid(oForm)
            End If


            If pVal.FormUID = "SEAE_JOB" Then
                oForm = SBO_Application.Forms.Item("SEAE_JOB")
                Try
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oItem = oForm.Items.Item("OBLBtt")
                        oItem.Enabled = True
                    ElseIf pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oItem = oForm.Items.Item("OBLBtt")
                        oItem.Enabled = False
                    End If
                Catch ex As Exception

                End Try
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                    Try

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
                                        oItem = oForm.Items.Item("142")
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

                                        'BIAddButt_Disable = "Yes"
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

                        If pVal.BeforeAction = False And pVal.ItemUID = "OBLGrd" Then
                            oGrid = oForm.Items.Item("OBLGrd").Specific
                            'oEdit = oForm.Items.Item("64").Specific

                            For F = 0 To oGrid.Rows.Count - 1
                                If oGrid.Rows.IsSelected(F) = True Then
                                    Dim DocNum As String = oGrid.DataTable.GetValue("DocNum", F)
                                    LoadFromXML("OceanBillofLading.srf", SBO_Application)
                                    oForm = SBO_Application.Forms.Item("SEAE_OBL")
                                    oEdit = oForm.Items.Item("64").Specific

                                    oForm.EnableMenu("1282", False)  '// Add New Record
                                    oForm.EnableMenu("1288", False)  '// Next Record
                                    oForm.EnableMenu("1289", False)  '// Pevious Record
                                    oForm.EnableMenu("1290", False)  '// First Record
                                    oForm.EnableMenu("1291", False)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oItem = oForm.Items.Item("64")
                                    oItem.Enabled = True
                                    oEdit.Value = DocNum
                                    oItem = oForm.Items.Item("1")
                                    oEdit = oForm.Items.Item("62").Specific
                                    oEdit.String = oEdit.String
                                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                    oItem = oForm.Items.Item("64")
                                    oItem.Enabled = False
                                End If
                            Next
                        End If
                    Catch ex As Exception
                        Functions.WriteLog("Class:F_SE_JobOrder" + " Function:ItemEvent" + " Error Message:" + ex.ToString)
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                End If
                '--------Load Matrix
                If pVal.ItemUID = "SEJGR" And pVal.ColUID = "V_0" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    Try


                        oEdit = oForm.Items.Item("24").Specific
                        Dim Vessel As String = oEdit.String

                        oMatrix3 = oForm.Items.Item("SEJGR").Specific
                        oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                        If oEdit.String <> "" Then
                            oEdit = oForm.Items.Item("SIJ6").Specific
                            If oEdit.String <> "" Then
                                oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet1.DoQuery("SELECT distinct T0.[DocNum], T0.[U_CardCode], T0.[U_CardName], T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_TaxDate] , T0.[U_ANSRecNo], T1.[U_VenName] FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and T0.[U_Status]='Open' and  T1.[U_Decript] <>'' and T0.[U_CardCode] ='" & oEdit.String & "' and  T0.[U_VesselNo] ='" & Vessel & "' and  T1.U_OpenQty>0 ORDER BY T0.[DocNum]")
                                If oRecordSet1.RecordCount = 1 Then
                                    Try
                                        oForm.Freeze(True)
                                        MatrixLoad(oRecordSet1.Fields.Item(0).Value, oRecordSet1.Fields.Item(8).Value, oRecordSet1.Fields.Item(3).Value)
                                        oForm.Freeze(False)
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try

                                End If
                            End If
                        End If
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                End If
                '---------------Item Event-----------------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "89" Then
                            oForm.PaneLevel = 1
                        ElseIf pVal.ItemUID = "SO90" Then
                            oForm.PaneLevel = 2 '
                        ElseIf pVal.ItemUID = "REFATTFOL" Then
                            oForm.PaneLevel = 11
                        ElseIf pVal.ItemUID = "DO1000001" Then
                            oForm.PaneLevel = 3
                        ElseIf pVal.ItemUID = "SIJ125VOU" Then
                            oForm.PaneLevel = 4
                        ElseIf pVal.ItemUID = "ATTACH" Then
                            oForm.PaneLevel = 5
                        ElseIf pVal.ItemUID = "OBL" Then
                            oForm.PaneLevel = 6
                        ElseIf pVal.ItemUID = "193" Then
                            oForm.PaneLevel = 7
                        ElseIf pVal.ItemUID = "BIFolder" Then
                            oForm.PaneLevel = 10
                        ElseIf pVal.ItemUID = "DOButt" Then
                            LoadDeliveryOrder(oForm)
                        ElseIf pVal.ItemUID = "PVButton" Then
                            'oEdit = oForm.Items.Item("SIJ16").Specific
                            'LoadDraftPaymentVouher(oEdit.String)
                            LoadPaymentVoucher(oForm)
                        ElseIf pVal.ItemUID = "OBLBtt" Then
                            'aswin
                            LoadOBL(oForm)
                        ElseIf pVal.ItemUID = "153" Then

                            oEdit = oForm.Items.Item("SIJ16").Specific
                            LoadDraftInvoice(oEdit.String)
                        ElseIf pVal.ItemUID = "152" Then
                            'Booking sheet
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf Booking_Sheet)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "SIJPSO" Then
                            'Shipping Order
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf SHipping_Order)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "SIJPV" Then
                            'PaymentVoucher
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf Payment_Voucher)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "152" Then
                            'print Booking Sheet
                        ElseIf pVal.ItemUID = "SIJPV" Then
                            'Shipping Order
                        ElseIf pVal.ItemUID = "149" Then
                            'Delivery Order
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf DOReport)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "PrintBI" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf BI_Report)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "139" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf MasterBL)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "138" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf HouseBL)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "150" Then
                            'Packing List
                        ElseIf pVal.ItemUID = "151" Then
                            'Tally SHeet
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf Tally_sheet)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                            'AB_RP006_TS

                        ElseIf pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            DocNumber_SE()
                            oEdit = oForm.Items.Item("SIJ18").Specific
                            oEdit.String = Format(Now.Date, "dd/MM/yy")
                            ShippigNameLoad()
                            'oMatrix3 = oForm.Items.Item("SEJGR").Specific
                            'oMatrix3.AddRow()
                            'oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                            'oMatrix.AddRow()
                            'oMatrix1 = oForm.Items.Item("148").Specific
                            'oMatrix1.AddRow()
                        End If
                    ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "1" Then
                            DocNumber_SE()
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
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                Try
                                    oCombo = oForm.Items.Item("sejj2").Specific
                                    If oCombo.Selected.Value = "Done" Then
                                        oEdit = oForm.Items.Item("137").Specific
                                        Dim BaseJobNo As String = oEdit.String
                                        oEdit = oForm.Items.Item("SIJ16").Specific
                                        Dim JobNo As String = oEdit.String
                                        If BaseJobNo.Substring(0, 2) = "IN" Or BaseJobNo.Substring(0, 2) = "PR" Then
                                            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet1.DoQuery("UPDATE ODRF SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "' and ObjType =13")
                                            oRecordSet1.DoQuery("UPDATE OINV SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                                            LoadGrid_BI(oForm)
                                        End If
                                    End If

                                Catch ex As Exception
                                End Try
                            End If
                        End If

                        'ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        '    If pVal.ItemUID = "1" Then

                        '    End If
                    End If
                End If
                '---------------Combo Select-----------------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "36" Then
                            oCombo = oForm.Items.Item("36").Specific
                            oEdit = oForm.Items.Item("39").Specific
                            oEdit.String = oCombo.Selected.Description
                        ElseIf pVal.ItemUID = "38" Then
                            oCombo = oForm.Items.Item("38").Specific
                            oEdit = oForm.Items.Item("40").Specific
                            oEdit.String = oCombo.Selected.Description
                        ElseIf pVal.ItemUID = "42" Then
                            oCombo = oForm.Items.Item("42").Specific
                            oEdit = oForm.Items.Item("43").Specific
                            oEdit.String = oCombo.Selected.Description
                        End If
                    End If

                End If
                '---------------Validate Event-----------------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "148" And (pVal.ColUID = "V_1" Or pVal.ColUID = "V_5") Then
                            Try
                                oMatrix1 = oForm.Items.Item("148").Specific
                                Dim i As Integer = 1
                                Dim amt As Double = 0
                                Dim TotAmt As Double = 0
                                Dim TotGST As Double = 0
                                Dim LineAmt As Double = 0
                                Dim TaxtAmt As Double = 0
                                For i = 1 To oMatrix1.RowCount
                                    oEdit = oMatrix1.Columns.Item("V_8").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        oEdit = oMatrix1.Columns.Item("V_1").Cells.Item(i).Specific
                                        amt = amt + oEdit.Value
                                        LineAmt = oEdit.Value
                                        oEdit = oMatrix1.Columns.Item("V_5").Cells.Item(i).Specific
                                        TaxtAmt = TAXPer(oEdit.String, Ocompany) * LineAmt * (1 / 100)
                                        TotGST = TotGST + TaxtAmt
                                    End If
                                Next
                                oEdit = oForm.Items.Item("139").Specific
                                oEdit.Value = amt
                                oEdit = oForm.Items.Item("143").Specific
                                oEdit.Value = TotGST
                                TotAmt = TotGST + amt
                                oEdit = oForm.Items.Item("145").Specific
                                oEdit.Value = TotAmt
                                'mitra
                            Catch ex As Exception
                            End Try

                        End If
                    End If
                    If pVal.Before_Action = False And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        'If pVal.ItemUID = "SEJ4" Then
                        '    oEdit = oForm.Items.Item("SEJ4").Specific
                        '    If oEdit.String <> "" Then
                        '        LoadJobOrder(oEdit.String)
                        '    End If
                        'End If

                        If pVal.ItemUID = "SEJ4" Then
                            oEdit = oForm.Items.Item("SEJ4").Specific
                            If oEdit.String <> "" Then
                                LoadJobOrder(oEdit.String)
                            End If
                        End If

                        If pVal.ItemUID = "SIJDOMAT" And pVal.ColUID = "V_14" Then
                            oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                            oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix.AddRow()
                                oMatrix.ClearRowData(oMatrix.RowCount)
                                oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                                oEdit.String = ""
                            End If
                        End If
                        If pVal.ItemUID = "148" And pVal.ColUID = "V_8" Then
                            oMatrix1 = oForm.Items.Item("148").Specific
                            oEdit = oMatrix1.Columns.Item("V_8").Cells.Item(oMatrix1.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix1.AddRow()
                                oMatrix1.ClearRowData(oMatrix1.RowCount)
                                oEdit = oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific
                                oEdit.String = ""
                            End If
                        End If
                        'aswin
                        If pVal.ItemUID = "SEContMat" And pVal.ColUID = "V_0" Then
                            oMatrix4 = oForm.Items.Item("SEContMat").Specific
                            oEdit = oMatrix4.Columns.Item("V_0").Cells.Item(oMatrix4.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix4.AddRow()
                                oMatrix4.ClearRowData(oMatrix4.RowCount)
                                oEdit = oMatrix4.Columns.Item("V_-1").Cells.Item(oMatrix4.RowCount).Specific
                                oEdit.String = ""
                            End If
                        End If

                        If pVal.ItemUID = "SEJ1000013" Then
                            oEdit = oForm.Items.Item("SEJ1000013").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("39").Specific
                                oEdit.String = Country_Code(ContCode, Ocompany)
                            End If
                        End If
                        If pVal.ItemUID = "SEJ190" Then
                            oEdit = oForm.Items.Item("SEJ190").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("40").Specific
                                oEdit.String = Country_Code(ContCode, Ocompany)
                            End If
                        End If
                        If pVal.ItemUID = "SEJ1000016" Then
                            oEdit = oForm.Items.Item("SEJ1000016").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("188").Specific
                                oEdit.String = City_Code(ContCode, Ocompany)
                            End If
                        End If
                        If pVal.ItemUID = "SEJ1000017" Then
                            oEdit = oForm.Items.Item("SEJ1000017").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("191").Specific
                                oEdit.String = City_Code(ContCode, Ocompany)
                            End If
                        End If


                    End If

                End If


                ''-----DO
                'oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                'oColumns = oMatrix.Columns
                'oMatrix.AddRow()
                ''------------VO
                'oMatrix1 = oForm.Items.Item("148").Specific
                'oColumns = oMatrix1.Columns
                'oMatrix1.AddRow()

                '--------------Form Resize-------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then
                    'Try
                    '    oForm = SBO_Application.Forms.Item("SEAE_JOB")
                    '    oForm.Freeze(True)
                    '    oMatrix3 = oForm.Items.Item("SEJGR").Specific
                    '    oColumns = oMatrix3.Columns
                    '    oColumn = oColumns.Item("V_0")
                    '    oItem = oForm.Items.Item("SEJGR")
                    '    oItem.Width = 150
                    '    oItem.Height = 30
                    '    oColumn.Width = 130
                    '    oForm.Freeze(False)
                    'Catch ex As Exception
                    '    oForm.Freeze(False)
                    'End Try
                    Try
                        If pVal.BeforeAction = False Then
                            oForm.Items.Item("136").Width = oForm.Width - 30
                            oForm.Items.Item("136").Height = oForm.Height - 220
                            oItem = oForm.Items.Item("se4612")
                            oItem.Left = 14
                            'oItem = oForm.Items.Item("70")
                            'oItem.Left = 20
                        End If
                    Catch ex As Exception
                    End Try
                End If
                '---------------CFL -----------------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
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
                            ElseIf pVal.ItemUID = "SEJ4" Then
                                Try
                                    oEdit = oForm.Items.Item("SEJ4").Specific
                                    oEdit.String = oDataTable.GetValue("DocNum", 0)
                                Catch ex As Exception
                                End Try
                            ElseIf pVal.ItemUID = "SIJ6" Then
                                Try
                                    oEdit = oForm.Items.Item("SIJ8").Specific
                                    oEdit.String = oDataTable.GetValue("CardName", 0)
                                    oEdit = oForm.Items.Item("SIJ6").Specific
                                    oEdit.String = oDataTable.GetValue("CardCode", 0)
                                Catch ex As Exception
                                End Try
                            ElseIf pVal.ItemUID = "127" Then
                                Try
                                    oEdit = oForm.Items.Item("129").Specific
                                    oEdit.String = oDataTable.GetValue("CardName", 0)
                                    oEdit = oForm.Items.Item("127").Specific
                                    oEdit.String = oDataTable.GetValue("CardCode", 0)
                                Catch ex As Exception
                                End Try
                            ElseIf pVal.ItemUID = "SIJDOMAT" And pVal.ColUID = "V_13" Then
                                oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardName", 0)
                                oEdit = oMatrix.Columns.Item("V_13").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardCode", 0)

                            ElseIf pVal.ItemUID = "SIJDOMAT" And pVal.ColUID = "V_15" Then
                                oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                                oEdit = oMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix.Columns.Item("V_15").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemCode", 0)
                            ElseIf pVal.ItemUID = "148" And pVal.ColUID = "V_8" Then
                                oMatrix1 = oForm.Items.Item("148").Specific
                                oEdit = oMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                oEdit.String = "SGD"
                                oEdit = oMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                oEdit.String = "1"
                                oEdit = oMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemCode", 0)
                            End If
                        End If
                    Catch ex As Exception
                        ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
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

                    If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oItem = oForm.Items.Item("153")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("DOButt")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("PVButton")
                        oItem.Enabled = False

                        oItem = oForm.Items.Item("149")
                        oItem.Enabled = False
                        'oItem = oForm.Items.Item("151")
                        'oItem.Enabled = False
                        oItem = oForm.Items.Item("SIJPSO")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("SIJPV")
                        oItem.Enabled = False

                        oItem = oForm.Items.Item("138")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("139")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("PrintBI")
                        oItem.Enabled = False
                    ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oItem = oForm.Items.Item("153")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("DOButt")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("PVButton")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("149")
                        oItem.Enabled = True
                        'oItem = oForm.Items.Item("151")
                        'oItem.Enabled = True
                        oItem = oForm.Items.Item("SIJPSO")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("SIJPV")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("138")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("139")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("PrintBI")
                        oItem.Enabled = True
                    End If
                Catch ex As Exception

                End Try
                '-----------End-----------------
            End If
        Catch ex As Exception
            If ex.Message <> "Form - Invalid Form" Then
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If (pVal.MenuUID = "43530" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "SEAE_JOB" Then
                    'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = True
                    'End If
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "SEAE_JOB" Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = False
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        DocNumber_SE()
                        oEdit = oForm.Items.Item("SIJ18").Specific
                        oEdit.String = Format(Now.Date, "dd/MM/yy")
                        ShippigNameLoad()
                        oMatrix4 = oForm.Items.Item("SEContMat").Specific
                        oColumns = oMatrix4.Columns
                        oMatrix4.AddRow(7)
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub CreatePO()
        Try
            Exit Sub
            Dim oform1 As SAPbouiCOM.Form
            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
            oMatrix2 = oform1.Items.Item("38").Specific
            oColumns = oMatrix2.Columns
            Dim i As Integer
            Dim k As Integer
            Dim oPO As SAPbobsCOM.Documents
            Dim VenCode As String = ""
            Dim ChrgCode As String = ""
            For i = 1 To oMatrix2.RowCount
                oEdit = oMatrix2.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                VenCode = oEdit.String
                oEdit = oMatrix2.Columns.Item("1").Cells.Item(i).Specific
                ChrgCode = oEdit.String
                If ChrgCode <> "" And VenCode <> "" Then
                    oPO = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    oPO.TaxDate = Now.Date
                    oPO.DocDate = Now.Date
                    oPO.CardCode = VenCode
                    oPO.Lines.ItemCode = ChrgCode
                    oPO.Lines.Add()
                    k = oPO.Add()
                    Dim st As String = ""
                    Ocompany.GetLastError(k, st)
                    If k = 0 Then
                        SBO_Application.StatusBar.SetText("Line No:" & i & " PO Created Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Else
                        SBO_Application.StatusBar.SetText("Line No:" & i & " PO Created Failed.Error Message:" & st & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            Next
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub LoadJobOrder(ByVal SQNO As String)
        Try

            'SELECT T0.[CardCode], T0.[CardName], T0.[U_AB_Divison], T0.[U_AB_TransTo], T0.[U_AB_Trnst], T0.[U_AB_VessCode], T0.[U_AB_VessName], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_Desc], T0.[U_AB_Validity], T0.[U_AB_Ttime], T0.[U_AB_Freq], T0.[U_AB_CARTotQt], T0.[U_AB_CARTotWt] FROM ORDR T0 WHERE T0.[DocNum] ='1'
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.DocNum,T0.[CardCode], T0.[CardName], T1.[Name], T0.[U_AB_VessCode], T0.[U_AB_VessName], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_Desc],T0.[Address],T0.U_AB_CaType,T0.[U_AB_OrginCodeC], T0.[U_AB_OriginNameC], T0.[U_AB_DestCodeC], T0.[U_AB_DestNameC],T0.[U_AB_Desc],T0.[U_AB_ContType],T0.[U_AB_TotWT], T0.[U_AB_TotVol], T0.[U_AB_TotPkg],T0.[U_AB_Divsion1], T0.[U_AB_JobNo]  FROM ORDR T0  Left JOIN OCPR T1 ON T0.CntctCode = T1.CntctCode WHERE T0.[DocNum] ='" & SQNO & "' and isnull(T0.U_AB_Status,'')='Open' and T0.[U_AB_Divsion]='SE'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            SBO_Application.StatusBar.SetText("Please Wait.Data is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            oEdit = oForm.Items.Item("SEJ4").Specific
            oEdit.String = oRecordSet1.Fields.Item(0).Value
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.String = oRecordSet1.Fields.Item(1).Value
            oEdit = oForm.Items.Item("SIJ8").Specific
            oEdit.String = oRecordSet1.Fields.Item(2).Value
            oEdit = oForm.Items.Item("SJI10").Specific
            oEdit.String = oRecordSet1.Fields.Item(3).Value
            oEdit = oForm.Items.Item("SIJ27").Specific
            oEdit.String = oRecordSet1.Fields.Item(4).Value
            oEdit = oForm.Items.Item("24").Specific
            oEdit.String = oRecordSet1.Fields.Item(5).Value

            'T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName],
            oEdit = oForm.Items.Item("SEJ1000013").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OrginCode").Value
            oEdit = oForm.Items.Item("SEJ1000016").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OrginCodeC").Value
            oEdit = oForm.Items.Item("188").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OriginNameC").Value

            oEdit = oForm.Items.Item("SEJ190").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestCode").Value

            oEdit = oForm.Items.Item("SEJ1000017").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestCodeC").Value

            oEdit = oForm.Items.Item("191").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestNameC").Value
            oEdit = oForm.Items.Item("34").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_Desc").Value
            oEdit = oForm.Items.Item("105").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_Desc").Value
            oEdit = oForm.Items.Item("39").Specific
            oEdit.String = oRecordSet1.Fields.Item(7).Value

            oEdit = oForm.Items.Item("40").Specific
            oEdit.String = oRecordSet1.Fields.Item(9).Value
            'T0.[Address]

            If oRecordSet1.Fields.Item("U_AB_CaType").Value = "1" Then
                ooption = oForm.Items.Item("28").Specific
                ooption.Selected = True
            End If

            If oRecordSet1.Fields.Item("U_AB_CaType").Value = "2" Then
                ooption = oForm.Items.Item("SIJ1000001").Specific
                ooption.Selected = True
            End If
            Try
                'U_AB_ContType
                oCombo = oForm.Items.Item("30").Specific
                oCombo.Select(oRecordSet1.Fields.Item("U_AB_ContType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try

            Try

                oItem = oForm.Items.Item("140")
                oItem.Enabled = True
                oItem = oForm.Items.Item("137")
                oItem.Enabled = True
                oEdit = oForm.Items.Item("140").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_Divsion1").Value
                End If
                oEdit = oForm.Items.Item("137").Specific
                If oRecordSet1.Fields.Item("U_AB_JobNo").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                End If
                oCombo = oForm.Items.Item("sejj2").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "IN" Or oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "PR" Then
                    oCombo.Select("Approved", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    oCombo.Select("NA", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                oItem = oForm.Items.Item("140")
                oItem.Enabled = False
                oItem = oForm.Items.Item("137")
                oItem.Enabled = False

                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "IN" Then
                    Try
                        Dim jno As String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("SELECT T0.[U_Shipp], T0.[U_COns], T0.[U_Rem] FROM [dbo].[@AB_INT_HEADER]  T0 WHERE T0.[U_JobNo] ='" & jno & "'")
                        oEdit = oForm.Items.Item("94").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_Shipp").Value
                        oEdit = oForm.Items.Item("96").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_COns").Value
                        oEdit = oForm.Items.Item("45").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_Rem").Value
                    Catch ex As Exception
                    End Try
                End If
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "PR" Then
                    Try
                        Dim jno As String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("SELECT T0.[U_Shipp], T0.[U_COns], T0.[U_Rem] FROM [dbo].[@AB_PRO_HEADER]  T0 WHERE T0.[U_JobNo] ='" & jno & "'")
                        oEdit = oForm.Items.Item("94").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_Shipp").Value
                        oEdit = oForm.Items.Item("96").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_COns").Value
                        oEdit = oForm.Items.Item("45").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_Rem").Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception

            End Try
            oEdit = oForm.Items.Item("73").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_TotWT").Value
            oEdit = oForm.Items.Item("75").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_TotVol").Value
            oEdit = oForm.Items.Item("71").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_TotPkg").Value

            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.String = oEdit.String
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Data Loading is Complected!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub MatrixLoad(ByVal DocNum As Integer, ByVal VenName As String, ByVal PONo As String)
        'oEdit = oForm.Items.Item("GI40").Specific
        'If oEdit.String <> "" Then
        'Dim DocNum As Integer = oEdit.String
        oForm = SBO_Application.Forms.Item("SEAE_JOB")
        Dim i As Integer
        oMatrix = oForm.Items.Item("SIJDOMAT").Specific
        Dim NewDocNum As Integer = 0
        Dim NewVenName As String = ""
        Dim NewPONo As String = ""
        For i = 1 To oMatrix.RowCount
            oEdit = oMatrix.Columns.Item("V_16").Cells.Item(i).Specific
            If oEdit.String <> "" Then
                NewDocNum = oEdit.String
            End If
            oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
            If oEdit.String <> "" Then
                NewVenName = oEdit.String
            End If
            oEdit = oMatrix.Columns.Item("V_11").Cells.Item(i).Specific
            If oEdit.String <> "" Then
                NewPONo = oEdit.String
            End If
            If NewDocNum = DocNum And NewVenName = VenName And NewPONo = PONo Then
                SBO_Application.StatusBar.SetText("This Record Already Entered..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Next
        Dim oRecordSet_GR As SAPbobsCOM.Recordset
        oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T1.[U_VenCode], T1.[U_VenName], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T0.[DocEntry], T1.[LineId],T0.U_CardCode FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  isnull(T1.[U_NumAtCar],'') ='" & PONo & "' and  isnull(T1.[U_VenName],'') ='" & VenName & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
        If oRecordSet_GR.RecordCount = 0 Then
            SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        'Dim NewBPCode As String = oRecordSet_GR.Fields.Item(22).Value.ToString.Trim
        'If CardCode <> NewBPCode Then
        '    SBO_Application.StatusBar.SetText("InValid BP Entered..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Exit Sub
        'End If
        'oEdit = oForm.Items.Item("10").Specific
        'If oEdit.String = "" Then
        '    oEdit.String = oRecordSet_GR.Fields.Item(0).Value
        'End If
        '' ''oEdit = oForm.Items.Item("20").Specific
        '' ''If oEdit.String = "" Then
        '' ''    oEdit.String = oRecordSet_GR.Fields.Item(1).Value
        '' ''End If
        '' ''oEdit = oForm.Items.Item("22").Specific
        '' ''If oEdit.String = "" Then
        '' ''    oEdit.String = oRecordSet_GR.Fields.Item(2).Value
        '' ''End If
        '' ''oEdit = oForm.Items.Item("24").Specific
        '' ''If oEdit.String = "" Then
        '' ''    oEdit.String = oRecordSet_GR.Fields.Item(3).Value
        '' ''End If
        ' '' ''oEdit = oForm.Items.Item("26").Specific
        ' '' ''oEdit.String = oRecordSet_GR.Fields.Item(4).Value
        '' ''oEdit = oForm.Items.Item("33").Specific
        '' ''If oEdit.String = "" Then
        '' ''    oEdit.String = oRecordSet_GR.Fields.Item(5).Value
        '' ''End If
        ' '' ''oEdit = oForm.Items.Item("35").Specific
        'If oEdit.String = "" Then
        '    oEdit.String = oRecordSet_GR.Fields.Item(6).Value
        'End If
        'oEdit = oForm.Items.Item("37").Specific
        'If oEdit.String = "" Then
        '    oEdit.String = oRecordSet_GR.Fields.Item(7).Value
        'End If
        Try
            oEdit = oForm.Items.Item("117").Specific
            If oEdit.String = "" Then
                oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
            Else
                oEdit.String = oEdit.String & " Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
            End If
        Catch ex As Exception
        End Try
        'oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], 
        '6T1.[U_VenCode], T1.[U_VenName], T0.[U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length],
        ' T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId],
        'T0.U_CardCode FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  isnull(T1.[U_NumAtCar],'') ='" & PONo & "' and  isnull(T1.[U_VenName],'') ='" & VenName & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
        'oColumns = oMatrix.Columns
        'oColumn = oColumns.Item("V_0")
        'oColumn.Editable = True
        For i = 1 To oRecordSet_GR.RecordCount
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow()
            End If
            'oMatrix.ClearRowData(oMatrix.RowCount)
            oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oMatrix.RowCount
            oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_ItemCode").Value
            oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Decript").Value

            Try
                'oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                'oCombo.Select(oRecordSet_GR.Fields.Item(11).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try

            oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_BinLoc").Value
            oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_VenName").Value
            oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_OpenQty").Value
            Try
                oEdit = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item("U_Weight").Value
            Catch ex As Exception
            End Try
            oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_NumAtCar").Value

            oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Length").Value & "x" & oRecordSet_GR.Fields.Item("U_Width").Value & "x" & oRecordSet_GR.Fields.Item("U_Height").Value
            oEdit = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("LineId").Value
            'oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(6).Value
            'oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(7).Value
            'oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(0).Value
            'If i <> oRecordSet_GR.RecordCount Then
            oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_VenCode").Value

            oEdit = oMatrix.Columns.Item("V_16").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("DocEntry").Value
            'End If
            oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
            If oRecordSet_GR.Fields.Item("U_Whsc").Value <> "" Then
                oEdit.String = oRecordSet_GR.Fields.Item("U_Whsc").Value
            Else
                oEdit.String = "01"
            End If

            oMatrix.AddRow()
            '  oMatrix.ClearRowData(oMatrix.RowCount)
            oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oMatrix.RowCount

            oRecordSet_GR.MoveNext()
        Next
        'oColmn.Editable = False

        oMatrix3 = oForm.Items.Item("SEJGR").Specific
        oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
        oEdit.String = ""
        'oEdit = oForm.Items.Item("GI40").Specific
        'oEdit.String = ""
        'End If
    End Sub
    Private Sub LoadOBL(ByVal oform As SAPbouiCOM.Form)
        Try
            Dim oform1 As SAPbouiCOM.Form
            LoadFromXML("OceanBillofLading.srf", SBO_Application)
            oform1 = SBO_Application.Forms.Item("SEAE_OBL")
            oCombo = oform1.Items.Item("36").Specific
            ComboLoad_ContainerType(oform1, oCombo)
            ooption = oform1.Items.Item("34").Specific
            ooption.GroupWith("35")

            Dim JobNo As String = ""
            oEdit = oform.Items.Item("SIJ16").Specific
            JobNo = oEdit.String
            oEdit = oform1.Items.Item("62").Specific
            oEdit.String = JobNo


            Dim Shipper As String = ""
            oEdit = oform.Items.Item("94").Specific
            Shipper = oEdit.String
            ' oEdit = oform.Items.Item("101").Specific
            Shipper = Shipper
            oEdit = oform1.Items.Item("9").Specific
            oEdit.String = Shipper

            Dim Consignee As String = ""
            oEdit = oform.Items.Item("96").Specific
            Consignee = oEdit.String
            oEdit = oform1.Items.Item("11").Specific
            oEdit.String = Consignee

            Dim NotifyParty As String = ""
            oEdit = oform.Items.Item("98").Specific
            NotifyParty = oEdit.String
            oEdit = oform1.Items.Item("13").Specific
            oEdit.String = NotifyParty
            Dim OBLno As String = ""
            oEdit = oform.Items.Item("63").Specific
            OBLno = oEdit.String
            oEdit = oform1.Items.Item("6").Specific
            oEdit.String = OBLno

            Dim Marking As String = ""
            oEdit = oform.Items.Item("103").Specific
            Marking = oEdit.String
            oEdit = oform1.Items.Item("29").Specific
            oEdit.String = Marking
            Dim Description As String = ""
            oEdit = oform.Items.Item("105").Specific
            Description = oEdit.String
            oEdit = oform1.Items.Item("55").Specific
            oEdit.String = Description
            Dim Placeofreceipt As String = ""
            oEdit = oform.Items.Item("188").Specific
            Placeofreceipt = oEdit.String
            oEdit = oform1.Items.Item("1000002").Specific
            oEdit.String = Placeofreceipt

            Dim VESSEL As String = ""
            oEdit = oform.Items.Item("49").Specific
            VESSEL = oEdit.String
            oEdit = oform1.Items.Item("17").Specific
            oEdit.String = VESSEL
            Dim PODCode As String = ""
            oEdit = oform.Items.Item("SEJ1000017").Specific
            PODCode = oEdit.String
            oEdit = oform1.Items.Item("19").Specific
            oEdit.String = PODCode
            oEdit = oform1.Items.Item("26").Specific
            oEdit.String = PODCode
            Dim PODName As String = ""
            oEdit = oform.Items.Item("191").Specific
            PODName = oEdit.String
            oEdit = oform1.Items.Item("24").Specific
            oEdit.String = PODName
            oEdit = oform1.Items.Item("27").Specific
            oEdit.String = PODName
            oEdit = oform1.Items.Item("53").Specific
            oEdit.String = PODName


            Dim LCL As Boolean = False
            ooption = oform.Items.Item("28").Specific
            If ooption.Selected = True Then
                LCL = True
            End If
            Dim FCL As Boolean = False
            ooption = oform.Items.Item("SIJ1000001").Specific
            If ooption.Selected = True Then
                FCL = True
            End If
            ooption = oform1.Items.Item("34").Specific
            If LCL = True Then
                ooption.Selected = True
            End If
            ooption = oform1.Items.Item("35").Specific
            If FCL = True Then
                ooption.Selected = True
            End If
            Dim Conttype As String = ""
            Try
                oCombo = oform.Items.Item("30").Specific
                Conttype = oCombo.Selected.Value
                oCombo = oform1.Items.Item("36").Specific
                oCombo.Select(Conttype, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try

            PODName = oEdit.String
            oEdit = oform1.Items.Item("24").Specific
            oEdit.String = PODName
            DocNumber_OBL(oform1, PODCode)

        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadOBL" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    ''Public Sub MatrixLoad(ByVal DocNum As Integer, ByVal VenName As String, ByVal PONo As String)
    ''    'oEdit = oForm.Items.Item("GI40").Specific
    ''    'If oEdit.String <> "" Then
    ''    'Dim DocNum As Integer = oEdit.String
    ''    oForm = SBO_Application.Forms.Item("SEAE_JOB")
    ''    Dim i As Integer
    ''    oMatrix = oForm.Items.Item("SIJDOMAT").Specific
    ''    Dim NewDocNum As Integer = 0
    ''    Dim NewVenName As String = ""
    ''    Dim NewPONo As String = ""
    ''    For i = 1 To oMatrix.RowCount
    ''        oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
    ''        If oEdit.String <> "" Then
    ''            NewDocNum = oEdit.String
    ''        End If
    ''        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
    ''        If oEdit.String <> "" Then
    ''            NewVenName = oEdit.String
    ''        End If
    ''        oEdit = oMatrix.Columns.Item("V_11").Cells.Item(i).Specific
    ''        If oEdit.String <> "" Then
    ''            NewPONo = oEdit.String
    ''        End If
    ''        If NewDocNum = DocNum And NewVenName = VenName And NewPONo = PONo Then
    ''            SBO_Application.StatusBar.SetText("This Record Already Entered..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    ''            Exit Sub
    ''        End If
    ''    Next
    ''    Dim oRecordSet_GR As SAPbobsCOM.Recordset
    ''    oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    ''    oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T1.[U_VenCode], T1.[U_VenName], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId],T0.U_CardCode FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  isnull(T1.[U_NumAtCar],'') ='" & PONo & "' and  isnull(T1.[U_VenName],'') ='" & VenName & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
    ''    If oRecordSet_GR.RecordCount = 0 Then
    ''        SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    ''        Exit Sub
    ''    End If
    ''    'Dim NewBPCode As String = oRecordSet_GR.Fields.Item(22).Value.ToString.Trim
    ''    'If CardCode <> NewBPCode Then
    ''    '    SBO_Application.StatusBar.SetText("InValid BP Entered..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    ''    '    Exit Sub
    ''    'End If
    ''    'oEdit = oForm.Items.Item("10").Specific
    ''    'If oEdit.String = "" Then
    ''    '    oEdit.String = oRecordSet_GR.Fields.Item(0).Value
    ''    'End If
    ''    '' ''oEdit = oForm.Items.Item("20").Specific
    ''    '' ''If oEdit.String = "" Then
    ''    '' ''    oEdit.String = oRecordSet_GR.Fields.Item(1).Value
    ''    '' ''End If
    ''    '' ''oEdit = oForm.Items.Item("22").Specific
    ''    '' ''If oEdit.String = "" Then
    ''    '' ''    oEdit.String = oRecordSet_GR.Fields.Item(2).Value
    ''    '' ''End If
    ''    '' ''oEdit = oForm.Items.Item("24").Specific
    ''    '' ''If oEdit.String = "" Then
    ''    '' ''    oEdit.String = oRecordSet_GR.Fields.Item(3).Value
    ''    '' ''End If
    ''    ' '' ''oEdit = oForm.Items.Item("26").Specific
    ''    ' '' ''oEdit.String = oRecordSet_GR.Fields.Item(4).Value
    ''    '' ''oEdit = oForm.Items.Item("33").Specific
    ''    '' ''If oEdit.String = "" Then
    ''    '' ''    oEdit.String = oRecordSet_GR.Fields.Item(5).Value
    ''    '' ''End If
    ''    ' '' ''oEdit = oForm.Items.Item("35").Specific
    ''    'If oEdit.String = "" Then
    ''    '    oEdit.String = oRecordSet_GR.Fields.Item(6).Value
    ''    'End If
    ''    'oEdit = oForm.Items.Item("37").Specific
    ''    'If oEdit.String = "" Then
    ''    '    oEdit.String = oRecordSet_GR.Fields.Item(7).Value
    ''    'End If
    ''    Try
    ''        oEdit = oForm.Items.Item("117").Specific
    ''        If oEdit.String = "" Then
    ''            oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
    ''        Else
    ''            oEdit.String = oEdit.String & " Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
    ''        End If
    ''    Catch ex As Exception
    ''    End Try
    ''    'oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], 
    ''    '6T1.[U_VenCode], T1.[U_VenName], T0.[U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length],
    ''    ' T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId],
    ''    'T0.U_CardCode FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  isnull(T1.[U_NumAtCar],'') ='" & PONo & "' and  isnull(T1.[U_VenName],'') ='" & VenName & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")

    ''    For i = 1 To oRecordSet_GR.RecordCount
    ''        If oMatrix.RowCount = 0 Then
    ''            oMatrix.AddRow()
    ''        End If
    ''        oMatrix.ClearRowData(oMatrix.RowCount)
    ''        oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oMatrix.RowCount
    ''        oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_ItemCode").Value
    ''        oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_Decript").Value
    ''        oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_Whsc").Value
    ''        Try
    ''            'oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
    ''            'oCombo.Select(oRecordSet_GR.Fields.Item(11).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
    ''        Catch ex As Exception

    ''        End Try

    ''        oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_BinLoc").Value
    ''        oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_VenCode").Value
    ''        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_VenName").Value
    ''        oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_OpenQty").Value
    ''        Try
    ''            oEdit = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
    ''            oEdit.String = oRecordSet_GR.Fields.Item("U_Weight").Value
    ''        Catch ex As Exception
    ''        End Try
    ''        oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_NumAtCar").Value

    ''        oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("U_Length").Value & "x" & oRecordSet_GR.Fields.Item("U_Width").Value & "x" & oRecordSet_GR.Fields.Item("U_Height").Value
    ''        oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("DocEntry").Value
    ''        oEdit = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
    ''        oEdit.String = oRecordSet_GR.Fields.Item("LineId").Value
    ''        'oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
    ''        'oEdit.String = oRecordSet_GR.Fields.Item(6).Value
    ''        'oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
    ''        'oEdit.String = oRecordSet_GR.Fields.Item(7).Value
    ''        'oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
    ''        'oEdit.String = oRecordSet_GR.Fields.Item(0).Value
    ''        If i <> oRecordSet_GR.RecordCount Then
    ''            oMatrix.AddRow()
    ''        End If

    ''        oRecordSet_GR.MoveNext()
    ''    Next

    ''    oMatrix1 = oForm.Items.Item("SIJDOMAT").Specific
    ''    oEdit = oMatrix1.Columns.Item("V_0").Cells.Item(1).Specific
    ''    oEdit.String = ""
    ''    'oEdit = oForm.Items.Item("GI40").Specific
    ''    'oEdit.String = ""
    ''    'End If
    ''End Sub
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
            oMatrix = oform1.Items.Item("29").Specific
            oCombo = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            ComboLoad_Unit(oform, oCombo)
            'Dim oRecordSet_GR As SAPbobsCOM.Recordset
            'oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T1.[U_VenCode], T1.[U_VenName], '' [U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId] FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
            'If oRecordSet_GR.RecordCount = 0 Then
            '    SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '    Exit Sub
            'End If

            'oEdit = oform1.Items.Item("20").Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(1).Value
            'oEdit = oform1.Items.Item("22").Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(2).Value
            'oEdit = oform1.Items.Item("24").Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(3).Value
            'oEdit = oform1.Items.Item("33").Specific
            'oEdit.String = oRecordSet_GR.Fields.Item(5).Value
            'Try
            '    oEdit = oform1.Items.Item("31").Specific
            '    If oEdit.String = "" Then
            '        oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
            '    Else
            '        oEdit.String = oEdit.String & "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
            '    End If
            'Catch ex As Exception
            'End Try
            'oform1 = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
            'oMatrix = oform1.Items.Item("29").Specific
            'Dim i As Integer
            'For i = 1 To oRecordSet_GR.RecordCount
            '    If oMatrix.RowCount = 0 Then
            '        oMatrix.AddRow()
            '    End If
            '    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(9).Value
            '    oEdit = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(19).Value
            '    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(10).Value
            '    Try
            '        oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
            '        oCombo.Select(oRecordSet_GR.Fields.Item(11).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            '    Catch ex As Exception

            '    End Try

            '    oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(12).Value
            '    oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(13).Value
            '    oEdit = oMatrix.Columns.Item("V_4").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(14).Value
            '    oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(15).Value
            '    Try
            '        oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
            '        oEdit.String = oRecordSet_GR.Fields.Item(16).Value
            '    Catch ex As Exception

            '    End Try

            '    oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(18).Value
            '    oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(17).Value
            '    oEdit = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(20).Value
            '    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(21).Value

            '    oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
            '    If oEdit.String = "" Then
            '        oEdit.String = oRecordSet_GR.Fields.Item(6).Value
            '    End If
            '    oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(7).Value
            '    oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
            '    oEdit.String = oRecordSet_GR.Fields.Item(0).Value
            '    oMatrix.AddRow()
            '    oRecordSet_GR.MoveNext()
            'Next

        Catch ex As Exception
            Functions.WriteLog("Class:F_SI_JobOrder" + " Function:LoadDeliveryOrder" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
            oPV.PV_Bind(oform1, SBO_Application, "SE", Ocompany)

        Catch ex As Exception
            Functions.WriteLog("Class:F_SI_JobOrder" + " Function:LoadPaymentVoucher" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Shared Sub LoadGrid_DO(ByVal oForm As SAPbouiCOM.Form)
        Try



            oGrid = oForm.Items.Item("DOGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum],T0.[U_JobNo] 'JobNo',  T0.[U_CardCode] 'Card Code', T0.[U_CardName] 'Card Name', T0.[U_VesselNo] 'Vessel', T0.[U_MAWBNo] 'OBL No', T0.[U_TaxDate] as 'Date', T0.[U_ANSRecNo] as 'DO No' FROM [dbo].[@AIGI]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("DOSE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("DOSE")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Sub LoadDraftInvoice(ByVal JobNo As String)
        Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],T0.U_VessName [U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_VessVoyage] [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_GrssWt],T0.[U_GKBNo], T0.[U_ItemDesc], T0.[U_ETD] FROM [dbo].[@AB_SEAE_JOB_H] T0   WHERE T0.[U_JobNo]   ='" & JobNo & "'")
            'T0.[U_GKBNo], T0.[U_ItemDesc], T0.[U_ETD]
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
            oCombo = oform1.Items.Item("U_AB_Divsion").Specific
            oCombo.Select("SE", SAPbouiCOM.BoSearchKey.psk_ByValue)
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
            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessVoyage").Value
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
            'T0.[U_GKBNo], T0.[U_ItemDesc], T0.[U_ETD]
            oEdit = oform1.Items.Item("U_AB_Desc").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_ItemDesc").Value
            oEdit = oform1.Items.Item("U_AB_ETDETA").Specific
            oEdit.String = Format(oRecordSet1.Fields.Item("U_ETD").Value, "dd/MM/yy")
            'U_VessName
            Dim QTNo As String = oRecordSet1.Fields.Item("U_QNo").Value
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[Dscription], T1.[Quantity], T1.[Price], T0.[DocCur],T1.[U_AB_Vendor],T1.U_AB_Cost,T1.unitMsr,T1.FreeTxt FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.[DocNum] ='" & QTNo & "'")
            'oform1.Freeze(False)
            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
            Try
                oCombo = oform1.Items.Item("63").Specific
                oCombo.Select(oRecordSet.Fields.Item("DocCur").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try

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
            ' oform1.Freeze(False)
            oform1.Title = "Billing Instruction"
            'SELECT T0.[U_CCode], T0.[U_CName], T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_JobNo], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ETD] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] =''")
        Catch ex As Exception
            'oform1.Freeze(False)
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
    Public Sub LoadDraftPaymentVouher(ByVal JobNo As String)
        Try
           oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],T0.U_VessName [U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_VessVoyage] [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_GrssWt] FROM [dbo].[@AB_SEAE_JOB_H] T0   WHERE T0.[U_JobNo]   ='" & JobNo & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            SBO_Application.ActivateMenuItem("2308")
            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(141, 1)
            ' oform1.Freeze(True)
            oform1.Title = "Payment Voucher"
            oItem = oform1.Items.Item("1")
            oItem.Visible = False

            Try
                Dim oNewItem As SAPbouiCOM.Item
                oNewItem = oform1.Items.Add("ADDPV", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem = oform1.Items.Item("1")
                oNewItem.Top = oItem.Top
                oNewItem.Height = oItem.Height
                oNewItem.Width = oItem.Width '+ 10
                oNewItem.Left = oItem.Left
                oButton = oNewItem.Specific
                oButton.Caption = "Add PV"
            Catch ex As Exception
            End Try

            'oEdit = oform1.Items.Item("4").Specific
            'oEdit.String = oRecordSet1.Fields.Item("U_CCode").Value
            'oEdit = oform1.Items.Item("54").Specific
            'oEdit.String = oRecordSet1.Fields.Item("U_CName").Value
            'oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-141, 1)
            'oform1.Freeze(True)

            Try
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-141, 1)
                ' oform1.Freeze(True)
            Catch ex As Exception
                SBO_Application.ActivateMenuItem("6913")
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-141, 1)
                'oform1.Freeze(True)
            End Try

            oCombo = oform1.Items.Item("U_AB_Divsion").Specific
            oCombo.Select("SE", SAPbouiCOM.BoSearchKey.psk_ByValue)
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
            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessVoyage").Value
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
            oEdit = oform1.Items.Item("U_AB_Consignee").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_CName").Value

            SBO_Application.StatusBar.SetText("Data Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            ' oform1.Freeze(False)
            oform1.Title = "Billing Instruction"
            'SELECT T0.[U_CCode], T0.[U_CName], T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_JobNo], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ETD] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] =''")
        Catch ex As Exception
            'oform1.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#Region "Reports"
    Private Sub Booking_Sheet()
        Try

            '1000004
            oEdit = oForm.Items.Item("1000004").Specific
            If oEdit.String = "" Then
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP002_BS.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            '' ''Dim ParaName As String = "@DocKey"
            '' ''Dim ParaValue As String = oEdit.Value
            '' ''Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
            '' ''Dim Para As New CrystalDecisions.Shared.ParameterDiscreteValue
            '' ''Dim MyArr1 As Array = ParaName.Split(";")
            '' ''Dim MyArr2 As Array = ParaValue.Split(";")
            '' ''For i As Integer = 0 To MyArr1.Length - 1
            '' ''    Para.Value = MyArr2(i)
            '' ''    pvCollection.Add(Para)
            '' ''    cryRpt.DataDefinition.ParameterFields(MyArr1(i)).ApplyCurrentValues(pvCollection)
            '' ''Next
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
            RptFrm.Text = "Booking Sheet"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub SHipping_Order()
        Try

            '1000004
            oEdit = oForm.Items.Item("1000004").Specific
            If oEdit.String = "" Then
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP003_SHO.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@DocKey"
            Dim ParaValue As String = oEdit.Value
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
            RptFrm.Text = "Shipping Order Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
            cryRpt.Load(sPath & "\GK_FM\DO_WHMS_SE.rpt")
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
    Private Sub Delivery_Order()
        Try

            '1000004
            oEdit = oForm.Items.Item("1000004").Specific
            If oEdit.String = "" Then
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP004_DO_SE.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@DocKey"
            Dim ParaValue As String = oEdit.Value
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
            RptFrm.Text = "Delivery Order Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Payment_Voucher()
        Try

            '1000004
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("PVGrid").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select Payment Voucher", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP005_PV_SE.rpt")
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

    Private Sub Tally_sheet()
        Try

            '1000004
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("DOGrid").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select Delivery Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP006_TS.rpt")
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
            RptFrm.Text = "Tally Sheet"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub MasterBL()
        Try
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("OBLGrd").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next

            '  Exit Sub
            'oEdit = oForm.Items.Item("DO4").Specific
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select Ocean Bill of Lading", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP005_BOLPrePrinted.rpt")
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
            RptFrm.Text = "Master Bill of Lading"
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub HouseBL()
        Try
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("OBLGrd").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next

            '  Exit Sub
            'oEdit = oForm.Items.Item("DO4").Specific
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select Ocean Bill of Lading", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            cryRpt.Load(sPath & "\GK_FM\SeaHBLReport.rpt")
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
            RptFrm.Text = "Draft Bill of Lading"
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
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
