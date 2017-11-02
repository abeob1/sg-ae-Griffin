Imports System.Diagnostics.Process
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports
Imports System.IO
Public Class F_AE_JobOrder
    Dim rowDelete As Integer
    Dim matrixUID As String
    Dim AWBForm As SAPbouiCOM.Form = Nothing
    Dim hawbForm As SAPbouiCOM.Form = Nothing
    Dim mawbForm As SAPbouiCOM.Form = Nothing
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Dim oF_PiecesWeight As F_PiecesWeight
    Dim oF_AWBParameter As F_AWBParameter
    Dim oform1 As SAPbouiCOM.Form


    Public ShowFolderBrowserThread As Threading.Thread
    Dim strpath As String
    Dim FilePath As String
    Dim FileName As String
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Public Sub AE_Job_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.DataTables.Add("DOAE")
            oForm.DataSources.DataTables.Add("PVAE")
            oForm.DataSources.DataTables.Add("AWB") 'REFAE
            oForm.DataSources.DataTables.Add("REFAE") '
            'BIAE
            oForm.DataSources.DataTables.Add("BIAE")
            oForm.PaneLevel = 1
            DocNumber_AI()
            oItem = oForm.Items.Item("153")
            oItem.Enabled = False
            oEdit = oForm.Items.Item("SIJ18").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            'ooption = oForm.Items.Item("85").Specific
            'ooption.GroupWith("86")
            oCombo = oForm.Items.Item("30").Specific
            ComboLoad_ContainerType(oForm, oCombo)
            'oCombo = oForm.Items.Item("1000011").Specific
            'ComboLoad_Carrier(oForm, oCombo)

            'CFL_BP_Supplier3(oForm, SBO_Application)
            'oMatrix4 = oForm.Items.Item("CargoMat").Specific
            'oColumns = oMatrix4.Columns
            'oMatrix4.AddRow()
            'oColumn = oColumns.Item("V_11")
            'oColumn.ChooseFromListUID = "3CFLBPV1"
            'oColumn.ChooseFromListAlias = "CardCode"

            'CFL

            CFL_BP_Customer(oForm, SBO_Application)
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.ChooseFromListUID = "CFLBPC"
            oEdit.ChooseFromListAlias = "CardCode"

            CFL_Item_Vessel(oForm, SBO_Application)
            oEdit = oForm.Items.Item("97").Specific
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemName"

            CFL_SalesOrder(oForm, SBO_Application, "AE")
            oEdit = oForm.Items.Item("AEJ4").Specific
            oEdit.ChooseFromListUID = "ORDR"
            oEdit.ChooseFromListAlias = "DocNum"

            oForm.DataBrowser.BrowseBy = "1000004"
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:AE_Job_Bind" + " Error Message:" + ex.ToString)
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_City" + " Error Message:" + ex.ToString)
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_Currency" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ComboLoad_ContainerType(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM [@AB_CARGOTYPE_AIR] T0 order by T0.COde")
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_ContainerType" + " Error Message:" + ex.ToString)
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_Whsc" + " Error Message:" + ex.ToString)
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_PaymentType" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub DocNumber_AI()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_AIRE_JOB_H]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "AE" & Format(Now.Date, "yy") & "J" & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "AE" & Format(Now.Date, "yy") & "J" & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "AE" & Format(Now.Date, "yy") & "J" & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "AE" & Format(Now.Date, "yy") & "J" & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "AE" & Format(Now.Date, "yy") & "J" & "0" & DocNum
            ElseIf DocNumLen >= 5 Then
                oEdit.String = "AE" & Format(Now.Date, "yy") & "J" & DocNum
            End If

        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:DocNumber_AI" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Public Sub ComboLoad_Carrier(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@AB_CARRIER]  T0 ORDER BY T0.[Name]")
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_Carrier" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub ComboLoad_FreightUnit(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@AB_FERIGHT]  T0 ORDER BY T0.[Name]")
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_FreightUnit" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub ComboLoad_VolumeUnit(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@AB_VO_UNITS]  T0 ORDER BY T0.[Name]")
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_VolumeUnit" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub ComboLoad_WeightUnit(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@AB_WT_UNITS]  T0 ORDER BY T0.[Name]")
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
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:ComboLoad_WeightUnit" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Shared Sub LoadGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            oGrid = oForm.Items.Item("AWBGRID").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.DocNum,'House' 'Type', T0.[U_AWBNo] 'MAWB No', T0.[U_B1] 'Flight No', T0.[U_FlighDate1] 'Flight Date', T0.[U_HAWBNo] 'HAWB No', T0.[U_ConsigneeNameAddr] 'Consignee Name' FROM [dbo].[@AB_AWB_H]  T0 where T0.[U_JobNo]='" & JobNo & "' Union all Select T0.DocNum,'Master' 'Type', T0.[U_AWBNo]  'MAWB No', T0.[U_B1] 'Flight No', T0.[U_FlighDate1] 'Flight Date', T0.[U_AWBNo] 'HAWB No', T0.[U_ConsigneeNameAddr] 'Consignee Name' FROM [dbo].[@AB_AWB_M]  T0 where T0.[U_JobNo]='" & JobNo & "' order by Type Desc"
            oForm.DataSources.DataTables.Item("AWB").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("AWB")
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_PV(ByVal oForm As SAPbouiCOM.Form)
        Try



            oGrid = oForm.Items.Item("PVGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum], T0.[U_VOVCode] as 'Vendor Code', T0.[U_VOVName] as 'Vendor Name', T0.[U_VOType] as 'Payment Type', T0.[U_JobNo] as 'Job No', T0.[U_VONo] as 'Voucher No', T0.[U_VODt] as 'Voucher Date', T0.[U_VOTotAmt] as 'Amount' FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 WHERE T0.[U_JobNo]  ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("PVAE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("PVAE")
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
            oForm.DataSources.DataTables.Item("DOAE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("DOAE")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_BI(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("BIGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = "'" & oEdit.String & "'"
            oEdit = oForm.Items.Item("aej101").Specific
            If oEdit.String <> "" And oEdit.String <> "NA" Then
                JobNo = JobNo & "," & "'" & oEdit.String & "'"
            End If
            Dim str As String = "SELECT DocEntry 'DocNum','DraftInvoice' DocumentType ,T0.[DocDate] 'BIDate', T0.[CardCode] 'Customer Code', T0.[CardName] 'Customer Name', T0.[U_AB_JobNo] 'Job No', T0.[DocTotal] 'Document Total' FROM ODRF T0 WHERE T0.[ObjType] =13 and  T0.[DocStatus] ='O' and  T0.[U_AB_JobNo] in ( " & JobNo & ") union all SELECT DocEntry 'DocNum','Invoice' DocumentType , T0.[DocDate] 'BIDate', T0.[CardCode] 'Customer Code', T0.[CardName] 'Customer Name', T0.[U_AB_JobNo] 'Job No', T0.[DocTotal] 'Document Total' FROM OINV T0 WHERE   T0.[U_AB_JobNo] in (" & JobNo & ")"
            oForm.DataSources.DataTables.Item("BIAE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("BIAE")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
        End Try
    End Sub
#End Region

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormUID = "AIRE_JOB" Then
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
    '
    Private Sub LoadfromHAWB(ByVal oform As SAPbouiCOM.Form)
        Try
            Dim JobNo As String = ""
            oEdit = oform.Items.Item("0_U_E").Specific
            JobNo = oEdit.String
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Str As String = "SELECT T1.U_VenCode, T1.U_VenName, T1.U_PONO, T1.U_PKg, T1.U_PkgType, T1.U_Wt, T1.U_Len, T1.U_Width, T1.U_Height, T1.U_M3, T1.U_Vol, T1.U_Desc FROM [dbo].[@AB_AWB_H]  T0 , [dbo].[@AB_AWB_H4]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[U_JobNo] ='" & JobNo & "' and isnull(T1.U_VenCode,'')<> '' ORDER BY T0.[DocEntry]"
            oRecordSet1.DoQuery(Str)
            Dim i As Integer = 0
            oMatrix = oform.Items.Item("CargoMAT").Specific
            oColumns = oMatrix.Columns
            oMatrix.Clear()
            For i = 1 To oRecordSet1.RecordCount
                oMatrix.AddRow()
                oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific
                oEdit.String = i
                'oEdit = oMatrix.Columns.Item("V_9").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_VenCode").Value
                'oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_VenName").Value
                'oEdit = oMatrix.Columns.Item("V_7").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_PONO").Value
                'oEdit = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_PKg").Value
                'oEdit = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_PkgType").Value
                'oEdit = oMatrix.Columns.Item("V_4").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_Wt").Value
                oEdit = oMatrix.Columns.Item("V_3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Len").Value
                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Width").Value
                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Height").Value
                oEdit = oMatrix.Columns.Item("V_11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_M3").Value
                oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_Vol").Value
                'oEdit = oMatrix.Columns.Item("V_012").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_Desc").Value

                oRecordSet1.MoveNext()
            Next

            Dim TotWt As Double = 0
            Dim TotPkg As Integer = 0
            Dim GrossWt As Double = 0
            'Dim i As Integer = 0
            For i = 1 To oMatrix.RowCount
                oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    TotWt = TotWt + oEdit.Value
                End If
                oEdit = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    TotPkg = TotPkg + oEdit.Value
                End If
                oEdit = oMatrix.Columns.Item("V_4").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    GrossWt = GrossWt + oEdit.Value
                End If
            Next
            oMatrix1 = oform.Items.Item("AWB_Mtr1").Specific
            oEdit = oMatrix1.Columns.Item("C_1_1").Cells.Item(1).Specific
            oEdit.String = TotPkg
            oEdit = oMatrix1.Columns.Item("C_1_2").Cells.Item(1).Specific
            oEdit.Value = GrossWt
            oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
            oEdit.String = TotWt

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Shared Sub LoadGrid_REF_ATTACH(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("REFATT").Specific
            oEdit = oForm.Items.Item("aej101").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            If JobNo.Contains("IN") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            ElseIf JobNo.Contains("PR") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_PRO_HEADER]  T0 , [dbo].[@AB_PRO_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            End If
            oForm.DataSources.DataTables.Item("REFAE").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("REFAE")
            oGrid.Columns.Item("RowsHeader").Width = 30
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_REF_ATTACH" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If ((pVal.FormUID = "AIRE_JOB" And pVal.ItemUID = "REFATTFOL" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                LoadGrid_REF_ATTACH(oForm)
            ElseIf ((pVal.FormUID = "AIRE_JOB" And pVal.ItemUID = "REFDIS" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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


            If ((pVal.FormUID = "AIRE_JOB" And pVal.ItemUID = "ATTMAT" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "AIRE_JOB" And pVal.ItemUID = "1000006" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then

                oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                oColumns = oMatrix1.Columns
                Dim i As Integer
                For i = 1 To oMatrix1.RowCount
                    If oMatrix1.IsRowSelected(i) Then
                        oMatrix1.DeleteRow(i)
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                        Exit For
                    End If
                Next
                oItem = oForm.Items.Item("1000006")
                oItem.Enabled = False
                oItem = oForm.Items.Item("155")
                oItem.Enabled = False

            End If
            If ((pVal.FormUID = "AIRE_JOB" And pVal.ItemUID = "155" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "AIRE_JOB" And pVal.ItemUID = "1000005" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                Try
                    oForm = SBO_Application.Forms.Item("AIRE_JOB")

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

            '---------
            If pVal.FormUID = "UDO_F_MAWB_D" Then
                Try
                    oForm = SBO_Application.Forms.Item("UDO_F_MAWB_D")

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.InnerEvent = False And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "Dim" Then
                            Dim i As Integer = 0
                            Dim l As Integer = 0
                            Dim w As Integer = 0
                            Dim h As Integer = 0
                            Dim Pkg As Integer = 0
                            Dim Dimen As String = ""
                            oEdit = oForm.Items.Item("149").Specific
                            oEdit.String = ""
                            Dim totVol As Double = 0
                            oMatrix = oForm.Items.Item("CargoMAT").Specific
                            oColumns = oMatrix.Columns
                            For i = 1 To oMatrix.RowCount
                                oEdit = oMatrix.Columns.Item("V_3").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    l = oEdit.Value
                                Else
                                    Exit For
                                    l = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    w = oEdit.Value
                                Else
                                    w = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    h = oEdit.Value
                                Else
                                    h = 0
                                End If
                                Dimen = l & "X" & w & "X" & h & " CM / "
                                oEdit = oForm.Items.Item("149").Specific
                                If oEdit.String <> "" Then
                                    oEdit.String = oEdit.String & vbCrLf & Dimen
                                Else
                                    oEdit.String = Dimen
                                End If
                                Try
                                    oEdit = oMatrix.Columns.Item("V_11").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        totVol = oEdit.Value + totVol
                                    Else
                                        totVol = totVol
                                    End If
                                Catch ex As Exception

                                End Try

                            Next

                            oEdit = oForm.Items.Item("1000019").Specific
                            oEdit.String = totVol
                            'Other charges
                            'LoadfromHAWB
                            oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                            oEdit = oMatrix1.Columns.Item("C_1_2").Cells.Item(1).Specific
                            Dim GrossWt As Double = oEdit.Value
                            If GrossWt > totVol Then
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                oEdit.String = Roundoff(GrossWt)
                            Else
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                oEdit.String = Roundoff(totVol)
                            End If
                            Try
                                oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                                Dim Rate As Double = 0
                                Dim Cwt As Double = 0
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Cwt = oEdit.Value
                                Else
                                    Cwt = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Rate = oEdit.Value
                                Else
                                    Exit Try
                                    Rate = 0
                                End If
                                Try
                                    oCombo = oMatrix1.Columns.Item("C_1_4").Cells.Item(1).Specific
                                    If oCombo.Selected.Value = "M" Then
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate
                                    Else
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate * Cwt
                                    End If
                                Catch ex As Exception
                                    oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                    oEdit.Value = Rate * Cwt
                                End Try

                                ooption = oForm.Items.Item("optionbtn1").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                ooption = oForm.Items.Item("optionbtn2").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0

                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "170" Then
                            LoadfromHAWB(oForm)
                        ElseIf (pVal.ItemUID = "optionbtn3" Or pVal.ItemUID = "optionbtn4") Then
                            Try
                                Dim i As Integer
                                Dim Amt As Double = 0.0
                                For i = 1 To oMatrix2.RowCount
                                    oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(i).Specific
                                    Amt = Amt + oEdit.Value
                                Next
                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = Amt
                                End If
                                oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                            Try
                                Dim Amt As Double = 0.0
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                Amt = Amt + oEdit.Value

                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = Amt
                                End If
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                            Dim WTChP As Double = 0
                            Dim RoChP As Double = 0
                            Dim TxChP As Double = 0
                            Dim DCChP As Double = 0
                            Dim DAChP As Double = 0


                            oEdit = oForm.Items.Item("94").Specific
                            WTChP = oEdit.String
                            oEdit = oForm.Items.Item("1000010").Specific
                            RoChP = oEdit.String
                            oEdit = oForm.Items.Item("98").Specific
                            TxChP = oEdit.String
                            oEdit = oForm.Items.Item("1000015").Specific
                            DCChP = oEdit.String
                            oEdit = oForm.Items.Item("99").Specific
                            DAChP = oEdit.String
                            oEdit = oForm.Items.Item("100").Specific
                            oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                            Dim WTChC As Double = 0
                            Dim RoChC As Double = 0
                            Dim TxChC As Double = 0
                            Dim DCChC As Double = 0
                            Dim DAChC As Double = 0
                            oEdit = oForm.Items.Item("ma101").Specific
                            WTChC = oEdit.String
                            oEdit = oForm.Items.Item("102").Specific
                            RoChC = oEdit.String
                            oEdit = oForm.Items.Item("103").Specific
                            TxChC = oEdit.String
                            oEdit = oForm.Items.Item("151").Specific
                            DCChC = oEdit.String
                            oEdit = oForm.Items.Item("104").Specific
                            DAChC = oEdit.String
                            oEdit = oForm.Items.Item("105").Specific
                            oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                            'oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                            'oEdit.String = oEdit.String
                        ElseIf (pVal.ItemUID = "optionbtn1" Or pVal.ItemUID = "optionbtn2") Then
                            Try
                                Dim Rate As Double = 0
                                Dim Cwt As Double = 0
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Cwt = oEdit.Value
                                Else
                                    Exit Try
                                    Cwt = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Rate = oEdit.Value
                                Else
                                    Exit Try
                                    Rate = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                oEdit.Value = Rate * Cwt
                                ooption = oForm.Items.Item("optionbtn1").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                ooption = oForm.Items.Item("optionbtn2").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = Rate * Cwt
                                End If

                            Catch ex As Exception
                            End Try
                            Dim WTChP As Double = 0
                            Dim RoChP As Double = 0
                            Dim TxChP As Double = 0
                            Dim DCChP As Double = 0
                            Dim DAChP As Double = 0


                            oEdit = oForm.Items.Item("94").Specific
                            WTChP = oEdit.String
                            oEdit = oForm.Items.Item("1000010").Specific
                            RoChP = oEdit.String
                            oEdit = oForm.Items.Item("98").Specific
                            TxChP = oEdit.String
                            oEdit = oForm.Items.Item("1000015").Specific
                            DCChP = oEdit.String
                            oEdit = oForm.Items.Item("99").Specific
                            DAChP = oEdit.String
                            oEdit = oForm.Items.Item("100").Specific
                            oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                            Dim WTChC As Double = 0
                            Dim RoChC As Double = 0
                            Dim TxChC As Double = 0
                            Dim DCChC As Double = 0
                            Dim DAChC As Double = 0
                            oEdit = oForm.Items.Item("ma101").Specific
                            WTChC = oEdit.String
                            oEdit = oForm.Items.Item("102").Specific
                            RoChC = oEdit.String
                            oEdit = oForm.Items.Item("103").Specific
                            TxChC = oEdit.String
                            oEdit = oForm.Items.Item("151").Specific
                            DCChC = oEdit.String
                            oEdit = oForm.Items.Item("104").Specific
                            DAChC = oEdit.String
                            oEdit = oForm.Items.Item("105").Specific
                            oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                            'oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                            'oEdit.String = oEdit.String
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = False And pVal.InnerEvent = False And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        If pVal.ItemUID = "17_U_E" Then
                            oEdit = oForm.Items.Item("17_U_E").Specific
                            Dim BPCode As String = ""
                            BPCode = oEdit.String
                            If oEdit.String <> "" Then
                                oEdit = oForm.Items.Item("18_U_E").Specific
                                oEdit.String = BPName(BPCode, Ocompany) & vbCrLf & BPAddress(BPCode, Ocompany)
                            End If
                        ElseIf (pVal.ItemUID = "ma101" Or pVal.ItemUID = "102" Or pVal.ItemUID = "103" Or pVal.ItemUID = "151" Or pVal.ItemUID = "104") Then
                            Try
                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                            Catch ex As Exception

                            End Try
                        ElseIf (pVal.ItemUID = "94" Or pVal.ItemUID = "1000010" Or pVal.ItemUID = "98" Or pVal.ItemUID = "1000015" Or pVal.ItemUID = "99") Then
                            Try
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0


                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP



                            Catch ex As Exception

                            End Try
                        ElseIf pVal.ItemUID = "150" And (pVal.ColUID = "V_2" Or pVal.ColUID = "V_3") Then
                            oMatrix2 = oForm.Items.Item("150").Specific
                            Dim Amt As Double = 0.0
                            Dim wt As Double = 0.0
                            oEdit = oMatrix2.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                            wt = oEdit.Value
                            If wt = 0 Then
                                Exit Sub
                            End If
                            oEdit = oMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                            Amt = oEdit.Value
                            If Amt = 0 Then
                                Exit Sub
                            End If
                            oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(pVal.Row).Specific
                            oEdit.Value = wt * Amt
                        ElseIf pVal.ItemUID = "150" And (pVal.ColUID = "V_0") Then
                            Try
                                oMatrix2 = oForm.Items.Item("150").Specific
                                Dim i As Integer
                                Dim Amt As Double = 0.0
                                For i = 1 To oMatrix2.RowCount
                                    oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(i).Specific
                                    Amt = Amt + oEdit.Value
                                Next
                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = Amt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0

                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC

                                oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "1000014" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_2") Then
                            oMatrix3 = oForm.Items.Item("1000014").Specific
                            Dim Amt As Double = 0.0
                            Dim wt As Double = 0.0
                            oEdit = oMatrix3.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                            wt = oEdit.Value
                            If wt = 0 Then
                                Exit Sub
                            End If
                            oEdit = oMatrix3.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                            Amt = oEdit.Value
                            If Amt = 0 Then
                                Exit Sub
                            End If
                            oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific
                            oEdit.Value = wt * Amt
                        ElseIf pVal.ItemUID = "1000014" And (pVal.ColUID = "V_0") Then
                            Try
                                oMatrix3 = oForm.Items.Item("1000014").Specific
                                Dim Amt As Double = 0.0
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                Amt = Amt + oEdit.Value

                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = Amt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0


                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "AWB_Mtr1" And (pVal.ColUID = "C_1_6" Or pVal.ColUID = "C_1_7") Then
                            Try
                                oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                                Dim Rate As Double = 0
                                Dim Cwt As Double = 0
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    Cwt = oEdit.Value
                                Else
                                    Cwt = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    Rate = oEdit.Value
                                Else
                                    Exit Try
                                    Rate = 0
                                End If
                                Try
                                    oCombo = oMatrix1.Columns.Item("C_1_4").Cells.Item(1).Specific
                                    If oCombo.Selected.Value = "M" Then
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate
                                    Else
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate * Cwt
                                    End If
                                Catch ex As Exception
                                    oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                    oEdit.Value = Rate * Cwt
                                End Try

                                ooption = oForm.Items.Item("optionbtn1").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                ooption = oForm.Items.Item("optionbtn2").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0

                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "CargoMAT" And (pVal.ColUID = "V_4" Or pVal.ColUID = "V_1") Then
                            Try
                                'oForm.Freeze(True)
                                Dim l As Integer = 0
                                Dim w As Integer = 0
                                Dim h As Integer = 0
                                Dim wt As Double = 0
                                Dim vol As Double = 0
                                Dim m3 As Double = 0
                                oEdit = oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    l = oEdit.Value
                                Else
                                    Exit Sub
                                    l = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    w = oEdit.Value
                                Else
                                    w = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    h = oEdit.Value
                                Else
                                    h = 0
                                End If
                                Dim Division As String = ""
                                Try
                                    'nath
                                    m3 = ((l * h * w) / 6000)
                                    ' m3 = ((l / 100) * (w / 100) * (h / 100))
                                Catch ex As Exception
                                    m3 = 0
                                End Try
                                oEdit = oMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific
                                oEdit.String = m3
                                oEdit = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    wt = oEdit.Value
                                Else
                                    wt = 0
                                End If
                                If m3 > wt Then
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                    oEdit.String = m3
                                Else
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                    oEdit.String = wt
                                End If

                                Dim TotWt As Double = 0
                                Dim TotPkg As Integer = 0
                                Dim GrossWt As Double = 0
                                Dim i As Integer = 0
                                For i = 1 To oMatrix.RowCount
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        TotWt = TotWt + oEdit.Value
                                    End If
                                    oEdit = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        TotPkg = TotPkg + oEdit.Value
                                    End If
                                    oEdit = oMatrix.Columns.Item("V_4").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        GrossWt = GrossWt + oEdit.Value
                                    End If
                                Next
                                oEdit = oMatrix1.Columns.Item("C_1_1").Cells.Item(1).Specific
                                oEdit.String = TotPkg
                                oEdit = oMatrix1.Columns.Item("C_1_2").Cells.Item(1).Specific
                                oEdit.Value = GrossWt
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                oEdit.String = TotWt
                                oEdit = oMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                oEdit.String = oEdit.String

                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "CargoMAT" And pVal.ColUID = "V_3" Then
                            oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix.AddRow()
                                oMatrix.ClearRowData(oMatrix.RowCount)
                                oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                                oEdit.String = ""
                            End If
                        ElseIf pVal.ItemUID = "20_U_E" Then
                            oEdit = oForm.Items.Item("20_U_E").Specific
                            Dim BPCode As String = ""
                            BPCode = oEdit.String
                            If oEdit.String <> "" Then
                                oEdit = oForm.Items.Item("21_U_E").Specific
                                oEdit.String = BPName(BPCode, Ocompany) & vbCrLf & BPAddress(BPCode, Ocompany)

                            End If
                        End If

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
                                If pVal.ItemUID = "17_U_E" Then
                                    Try
                                        oEdit = oForm.Items.Item("17_U_E").Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If
                                If pVal.ItemUID = "CargoMAT" And pVal.ColUID = "V_9" Then
                                    Try
                                        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                        oEdit.String = oDataTable.GetValue("CardName", 0)
                                        oEdit = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If

                                If pVal.ItemUID = "20_U_E" Then
                                    Try
                                        oEdit = oForm.Items.Item("20_U_E").Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception

                                    End Try
                                End If

                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception

                End Try
            End If
            '------------------------------**************************-----------------------------

            '**************************HAWB************************************************************
            If pVal.FormUID = "UDO_F_HAWB_D" Then
                Try
                    oForm = SBO_Application.Forms.Item("UDO_F_HAWB_D")
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.InnerEvent = False And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "Dim" Then
                            Dim i As Integer = 0
                            Dim l As Integer = 0
                            Dim w As Integer = 0
                            Dim h As Integer = 0
                            Dim Pkg As Integer = 0
                            Dim Dimen As String = ""
                            oEdit = oForm.Items.Item("149").Specific
                            oEdit.String = ""
                            Dim totVol As Double = 0
                            oMatrix = oForm.Items.Item("CargoMAT").Specific
                            oColumns = oMatrix.Columns
                            For i = 1 To oMatrix.RowCount
                                oEdit = oMatrix.Columns.Item("V_3").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    l = oEdit.Value
                                Else
                                    Exit For
                                    l = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    w = oEdit.Value
                                Else
                                    w = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    h = oEdit.Value
                                Else
                                    h = 0
                                End If
                                Dimen = l & "X" & w & "X" & h & " CM / "
                                oEdit = oForm.Items.Item("149").Specific
                                If oEdit.String <> "" Then
                                    oEdit.String = oEdit.String & vbCrLf & Dimen
                                Else
                                    oEdit.String = Dimen
                                End If
                                Try
                                    oEdit = oMatrix.Columns.Item("V_11").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        totVol = oEdit.Value + totVol
                                    Else
                                        totVol = totVol
                                    End If
                                Catch ex As Exception

                                End Try

                            Next

                            oEdit = oForm.Items.Item("1000019").Specific
                            oEdit.String = totVol
                            'Other charges
                            'LoadfromHAWB
                            oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                            oEdit = oMatrix1.Columns.Item("C_1_2").Cells.Item(1).Specific
                            Dim GrossWt As Double = oEdit.Value
                            If GrossWt > totVol Then
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                oEdit.String = Roundoff(GrossWt)
                            Else
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                oEdit.String = Roundoff(totVol)
                            End If
                            Try
                                oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                                Dim Rate As Double = 0
                                Dim Cwt As Double = 0
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Cwt = oEdit.Value
                                Else
                                    Cwt = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Rate = oEdit.Value
                                Else
                                    Exit Try
                                    Rate = 0
                                End If
                                Try
                                    oCombo = oMatrix1.Columns.Item("C_1_4").Cells.Item(1).Specific
                                    If oCombo.Selected.Value = "M" Then
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate
                                    Else
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate * Cwt
                                    End If
                                Catch ex As Exception
                                    oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                    oEdit.Value = Rate * Cwt
                                End Try

                                ooption = oForm.Items.Item("optionbtn1").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                ooption = oForm.Items.Item("optionbtn2").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0

                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "170" Then
                            LoadfromHAWB(oForm)
                        ElseIf (pVal.ItemUID = "optionbtn3" Or pVal.ItemUID = "optionbtn4") Then
                            Try
                                Dim i As Integer
                                Dim Amt As Double = 0.0
                                For i = 1 To oMatrix2.RowCount
                                    oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(i).Specific
                                    Amt = Amt + oEdit.Value
                                Next
                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = Amt
                                End If
                                oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                            Try
                                Dim Amt As Double = 0.0
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                Amt = Amt + oEdit.Value

                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = Amt
                                End If
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                            Dim WTChP As Double = 0
                            Dim RoChP As Double = 0
                            Dim TxChP As Double = 0
                            Dim DCChP As Double = 0
                            Dim DAChP As Double = 0


                            oEdit = oForm.Items.Item("94").Specific
                            WTChP = oEdit.String
                            oEdit = oForm.Items.Item("1000010").Specific
                            RoChP = oEdit.String
                            oEdit = oForm.Items.Item("98").Specific
                            TxChP = oEdit.String
                            oEdit = oForm.Items.Item("1000015").Specific
                            DCChP = oEdit.String
                            oEdit = oForm.Items.Item("99").Specific
                            DAChP = oEdit.String
                            oEdit = oForm.Items.Item("100").Specific
                            oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                            Dim WTChC As Double = 0
                            Dim RoChC As Double = 0
                            Dim TxChC As Double = 0
                            Dim DCChC As Double = 0
                            Dim DAChC As Double = 0
                            oEdit = oForm.Items.Item("ma101").Specific
                            WTChC = oEdit.String
                            oEdit = oForm.Items.Item("102").Specific
                            RoChC = oEdit.String
                            oEdit = oForm.Items.Item("103").Specific
                            TxChC = oEdit.String
                            oEdit = oForm.Items.Item("151").Specific
                            DCChC = oEdit.String
                            oEdit = oForm.Items.Item("104").Specific
                            DAChC = oEdit.String
                            oEdit = oForm.Items.Item("105").Specific
                            oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                            'oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                            'oEdit.String = oEdit.String
                        ElseIf (pVal.ItemUID = "optionbtn1" Or pVal.ItemUID = "optionbtn2") Then
                            Try
                                Dim Rate As Double = 0
                                Dim Cwt As Double = 0
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Cwt = oEdit.Value
                                Else
                                    Exit Try
                                    Cwt = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(1).Specific
                                If oEdit.String <> "" Then
                                    Rate = oEdit.Value
                                Else
                                    Exit Try
                                    Rate = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                oEdit.Value = Rate * Cwt
                                ooption = oForm.Items.Item("optionbtn1").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                ooption = oForm.Items.Item("optionbtn2").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = Rate * Cwt
                                End If

                            Catch ex As Exception
                            End Try
                            Dim WTChP As Double = 0
                            Dim RoChP As Double = 0
                            Dim TxChP As Double = 0
                            Dim DCChP As Double = 0
                            Dim DAChP As Double = 0


                            oEdit = oForm.Items.Item("94").Specific
                            WTChP = oEdit.String
                            oEdit = oForm.Items.Item("1000010").Specific
                            RoChP = oEdit.String
                            oEdit = oForm.Items.Item("98").Specific
                            TxChP = oEdit.String
                            oEdit = oForm.Items.Item("1000015").Specific
                            DCChP = oEdit.String
                            oEdit = oForm.Items.Item("99").Specific
                            DAChP = oEdit.String
                            oEdit = oForm.Items.Item("100").Specific
                            oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                            Dim WTChC As Double = 0
                            Dim RoChC As Double = 0
                            Dim TxChC As Double = 0
                            Dim DCChC As Double = 0
                            Dim DAChC As Double = 0
                            oEdit = oForm.Items.Item("ma101").Specific
                            WTChC = oEdit.String
                            oEdit = oForm.Items.Item("102").Specific
                            RoChC = oEdit.String
                            oEdit = oForm.Items.Item("103").Specific
                            TxChC = oEdit.String
                            oEdit = oForm.Items.Item("151").Specific
                            DCChC = oEdit.String
                            oEdit = oForm.Items.Item("104").Specific
                            DAChC = oEdit.String
                            oEdit = oForm.Items.Item("105").Specific
                            oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = False And pVal.InnerEvent = False And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "17_U_E" Then
                            oEdit = oForm.Items.Item("17_U_E").Specific
                            Dim BPCode As String = ""
                            BPCode = oEdit.String
                            If oEdit.String <> "" Then
                                oEdit = oForm.Items.Item("18_U_E").Specific
                                oEdit.String = BPName(BPCode, Ocompany) & vbCrLf & BPAddress(BPCode, Ocompany)
                            End If
                        ElseIf (pVal.ItemUID = "ma101" Or pVal.ItemUID = "102" Or pVal.ItemUID = "103" Or pVal.ItemUID = "151" Or pVal.ItemUID = "104") Then
                            Try
                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                            Catch ex As Exception

                            End Try
                        ElseIf (pVal.ItemUID = "94" Or pVal.ItemUID = "1000010" Or pVal.ItemUID = "98" Or pVal.ItemUID = "1000015" Or pVal.ItemUID = "99") Then
                            Try
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0


                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP



                            Catch ex As Exception

                            End Try
                        ElseIf pVal.ItemUID = "150" And (pVal.ColUID = "V_0") Then
                            Try
                                oMatrix2 = oForm.Items.Item("150").Specific
                                Dim i As Integer
                                Dim Amt As Double = 0.0
                                For i = 1 To oMatrix2.RowCount
                                    oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(i).Specific
                                    Amt = Amt + oEdit.Value
                                Next
                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("1000015").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("151").Specific
                                    oEdit.Value = Amt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0


                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC

                                oEdit = oMatrix2.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "1000014" And (pVal.ColUID = "V_0") Then
                            Try
                                oMatrix3 = oForm.Items.Item("1000014").Specific
                                Dim Amt As Double = 0.0
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                Amt = Amt + oEdit.Value

                                ooption = oForm.Items.Item("optionbtn3").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = Amt
                                End If
                                ooption = oForm.Items.Item("optionbtn4").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("99").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("104").Specific
                                    oEdit.Value = Amt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0


                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                                oEdit = oMatrix3.Columns.Item("V_0").Cells.Item(1).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "AWB_Mtr1" And (pVal.ColUID = "C_1_6" Or pVal.ColUID = "C_1_7") Then
                            Try
                                oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                                Dim Rate As Double = 0
                                Dim Cwt As Double = 0
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    Cwt = oEdit.Value
                                Else
                                    Cwt = 0
                                End If
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    Rate = oEdit.Value
                                Else
                                    Exit Try
                                    Rate = 0
                                End If
                                Try
                                    oCombo = oMatrix1.Columns.Item("C_1_4").Cells.Item(1).Specific
                                    If oCombo.Selected.Value = "M" Then
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate
                                    Else
                                        oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                        oEdit.Value = Rate * Cwt
                                    End If
                                Catch ex As Exception
                                    oEdit = oMatrix1.Columns.Item("C_1_8").Cells.Item(1).Specific
                                    oEdit.Value = Rate * Cwt
                                End Try

                                ooption = oForm.Items.Item("optionbtn1").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                ooption = oForm.Items.Item("optionbtn2").Specific
                                If ooption.Selected = True Then
                                    oEdit = oForm.Items.Item("94").Specific
                                    oEdit.Value = 0
                                    oEdit = oForm.Items.Item("ma101").Specific
                                    oEdit.Value = Rate * Cwt
                                End If
                                Dim WTChP As Double = 0
                                Dim RoChP As Double = 0
                                Dim TxChP As Double = 0
                                Dim DCChP As Double = 0
                                Dim DAChP As Double = 0

                                oEdit = oForm.Items.Item("94").Specific
                                WTChP = oEdit.String
                                oEdit = oForm.Items.Item("1000010").Specific
                                RoChP = oEdit.String
                                oEdit = oForm.Items.Item("98").Specific
                                TxChP = oEdit.String
                                oEdit = oForm.Items.Item("1000015").Specific
                                DCChP = oEdit.String
                                oEdit = oForm.Items.Item("99").Specific
                                DAChP = oEdit.String
                                oEdit = oForm.Items.Item("100").Specific
                                oEdit.String = WTChP + RoChP + TxChP + DCChP + DAChP

                                Dim WTChC As Double = 0
                                Dim RoChC As Double = 0
                                Dim TxChC As Double = 0
                                Dim DCChC As Double = 0
                                Dim DAChC As Double = 0
                                oEdit = oForm.Items.Item("ma101").Specific
                                WTChC = oEdit.String
                                oEdit = oForm.Items.Item("102").Specific
                                RoChC = oEdit.String
                                oEdit = oForm.Items.Item("103").Specific
                                TxChC = oEdit.String
                                oEdit = oForm.Items.Item("151").Specific
                                DCChC = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                DAChC = oEdit.String
                                oEdit = oForm.Items.Item("105").Specific
                                oEdit.String = WTChC + RoChC + TxChC + DCChC + DAChC
                                oEdit = oMatrix1.Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oEdit.String
                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "CargoMAT" And (pVal.ColUID = "V_4" Or pVal.ColUID = "V_1") Then
                            Try
                                'oForm.Freeze(True)
                                Dim l As Integer = 0
                                Dim w As Integer = 0
                                Dim h As Integer = 0
                                Dim wt As Double = 0
                                Dim vol As Double = 0
                                Dim m3 As Double = 0
                                oMatrix = oForm.Items.Item("CargoMAT").Specific
                                oEdit = oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    l = oEdit.Value
                                Else
                                    Exit Sub
                                    l = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    w = oEdit.Value
                                Else
                                    w = 0
                                End If
                                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    h = oEdit.Value
                                Else
                                    h = 0
                                End If
                                Dim Division As String = ""
                                Try
                                    'nath
                                    m3 = ((l * h * w) / 6000)
                                    'm3 = ((l / 100) * (w / 100) * (h / 100))
                                Catch ex As Exception
                                    m3 = 0
                                End Try
                                oEdit = oMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific
                                oEdit.String = m3
                                oEdit = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    wt = oEdit.Value
                                Else
                                    wt = 0
                                End If
                                If m3 > wt Then
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                    oEdit.String = m3
                                Else
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                    oEdit.String = wt
                                End If

                                Dim TotWt As Double = 0
                                Dim TotPkg As Integer = 0
                                Dim GrossWt As Double = 0
                                Dim i As Integer = 0
                                For i = 1 To oMatrix.RowCount
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        TotWt = TotWt + oEdit.Value
                                    End If
                                    oEdit = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        TotPkg = TotPkg + oEdit.Value
                                    End If
                                    oEdit = oMatrix.Columns.Item("V_4").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        GrossWt = GrossWt + oEdit.Value
                                    End If
                                Next
                                oEdit = oMatrix1.Columns.Item("C_1_1").Cells.Item(1).Specific
                                oEdit.String = TotPkg
                                oEdit = oMatrix1.Columns.Item("C_1_2").Cells.Item(1).Specific
                                oEdit.Value = GrossWt
                                oEdit = oMatrix1.Columns.Item("C_1_6").Cells.Item(1).Specific
                                oEdit.String = TotWt
                                oEdit = oMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                oEdit.String = oEdit.String

                            Catch ex As Exception
                            End Try
                        ElseIf pVal.ItemUID = "CargoMAT" And pVal.ColUID = "V_3" Then

                            oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix.AddRow()
                                oMatrix.ClearRowData(oMatrix.RowCount)
                                oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                                oEdit.String = ""
                            End If
                        ElseIf pVal.ItemUID = "20_U_E" Then
                            oEdit = oForm.Items.Item("20_U_E").Specific
                            Dim BPCode As String = ""
                            BPCode = oEdit.String
                            If oEdit.String <> "" Then
                                oEdit = oForm.Items.Item("21_U_E").Specific
                                oEdit.String = BPName(BPCode, Ocompany) & vbCrLf & BPAddress(BPCode, Ocompany)

                            End If
                        End If

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
                                If pVal.ItemUID = "17_U_E" Then
                                    Try
                                        oEdit = oForm.Items.Item("17_U_E").Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If
                                If pVal.ItemUID = "CargoMAT" And pVal.ColUID = "V_9" Then
                                    Try
                                        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                        oEdit.String = oDataTable.GetValue("CardName", 0)
                                        oEdit = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If

                                If pVal.ItemUID = "20_U_E" Then
                                    Try
                                        oEdit = oForm.Items.Item("20_U_E").Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If

                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception

                End Try
            End If
            '------------------------------**************************-----------------------------
            '***************************END HAWB********************************************************
            Try

                If FormUID = If(mawbForm Is Nothing, "", mawbForm.UniqueID) Then
                    AWBForm = mawbForm
                ElseIf FormUID = If(hawbForm Is Nothing, "", hawbForm.UniqueID) Then
                    AWBForm = hawbForm
                End If

                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            AWBForm = Nothing
                    End Select
                Else
                    BubbleEvent = True
                    If Not AWBForm Is Nothing Then
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                'Add row to matrix when press button Add which at bottom of matrix
                                If pVal.ItemUID = "btnAdd1" Then
                                    oMatrix = AWBForm.Items.Item("AWB_Mtr1").Specific
                                    oMatrix.AddRow(1, oMatrix.RowCount)
                                ElseIf pVal.ItemUID = "btnAdd2" Then
                                    oMatrix = AWBForm.Items.Item("AWB_Mtr2").Specific
                                    oMatrix.AddRow(1, oMatrix.RowCount)
                                ElseIf pVal.ItemUID = "CDManifest" Or pVal.ItemUID = "Manifest" Then
                                    oF_PiecesWeight = New F_PiecesWeight(Ocompany, SBO_Application, AWBForm, pVal.ItemUID)
                                ElseIf pVal.ItemUID = "btn_Print" Then
                                    oF_AWBParameter = New F_AWBParameter(Ocompany, SBO_Application, AWBForm, pVal.ItemUID)
                                End If
                                If pVal.BeforeAction = False Then
                                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        AWBForm.Close()
                                        AWBForm = Nothing
                                    End If
                                End If
                        End Select
                    End If
                End If

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

            '------laod matrix
            If pVal.FormType = 2000108 Then
                'If (pVal.ItemUID = "1" And pVal.Before_Action = False And pVal.InnerEvent = False And SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Or (pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                If (pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) Then
                    oForm = SBO_Application.Forms.Item("AIRE_JOB")
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
            If (pVal.FormUID = "UDO_F_HAWB_D" Or pVal.FormUID = "UDO_F_MAWB_D") And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Try
                    oForm = SBO_Application.Forms.Item("AIRE_JOB")
                    If oForm.Visible = True Then
                        LoadGrid(oForm)
                    End If
                Catch ex As Exception
                End Try
            End If

            If pVal.FormUID = "AI_FI_GoodsIssue" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("AIRE_JOB")
                If oForm.Visible = True Then
                    LoadGrid_DO(oForm)
                End If
            End If
            If pVal.FormUID = "AB_PV" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("AIRE_JOB")
                If oForm.Visible = True Then
                    LoadGrid_PV(oForm)
                End If
            End If
            If pVal.FormType = 133 And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("AIRE_JOB")
                If oForm.Visible = True Then
                    LoadGrid_BI(oForm)
                End If
            End If
            If pVal.FormUID = "AIRE_JOB" Then
                oForm = SBO_Application.Forms.Item("AIRE_JOB")

                '--------Load Matrix

                If pVal.ItemUID = "AECC107" And pVal.InnerEvent = False And pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("AECC107").Specific
                    Dim ContCode As String = oEdit.String
                    If ContCode <> "" Then
                        oEdit = oForm.Items.Item("61").Specific
                        oEdit.String = Carrier_Name(ContCode, Ocompany)
                    End If
                End If

                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.InnerEvent = False And pVal.Before_Action = False Then
                    Try
                        oItem = oForm.Items.Item("110")
                        If oItem.Enabled = False Then
                            oItem.Enabled = True
                            oItem = oForm.Items.Item("112")
                            oItem.Enabled = True
                            Exit Sub
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.InnerEvent = False And pVal.Before_Action = False Then
                    Try
                        oItem = oForm.Items.Item("110")
                        If oItem.Enabled = True Then
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("112")
                            oItem.Enabled = False
                            Exit Sub
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.InnerEvent = False Then
                    Try

                        oItem = oForm.Items.Item("HAWB")
                        If oItem.Enabled = True Then
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("MAWB")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("PVButton")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("SIJPV")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("153")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("DOButt")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("149")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("PrintBI")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("PrintAWB")
                            oItem.Enabled = False
                            oItem = oForm.Items.Item("PrintManf")
                            oItem.Enabled = False
                            Exit Sub
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.InnerEvent = False Then
                    Try
                        oItem = oForm.Items.Item("HAWB")
                        If oItem.Enabled = False Then
                            oItem.Enabled = True
                            oItem = oForm.Items.Item("MAWB")
                            oItem.Enabled = True
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
                            oItem = oForm.Items.Item("PrintAWB")
                            oItem.Enabled = True
                            oItem = oForm.Items.Item("PrintManf")
                            oItem.Enabled = True
                            Exit Sub
                        End If
                    Catch ex As Exception
                    End Try
                End If
                '---------------Item Event-----------------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "89" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 1
                        ElseIf pVal.ItemUID = "AWBinfo" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 2
                        ElseIf pVal.ItemUID = "REFATTFOL" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 11
                        ElseIf pVal.ItemUID = "Charge" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 6
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                        ElseIf pVal.ItemUID = "Cargo" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 7
                        ElseIf pVal.ItemUID = "DO1000001" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 3
                        ElseIf pVal.ItemUID = "SIJ125VOU" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 4
                        ElseIf pVal.ItemUID = "ATTACH" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 5
                        ElseIf pVal.ItemUID = "BIFolder" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 10
                        ElseIf pVal.ItemUID = "DOButt" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            LoadDeliveryOrder(oForm)
                        ElseIf pVal.ItemUID = "PVButton" Then
                            LoadPaymentVoucher(oForm)
                        ElseIf pVal.ItemUID = "HAWB" Then
                            'laod HAWB &AWB
                            oEdit = oForm.Items.Item("AEJ4").Specific
                            Dim QNo As String = oEdit.String
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                oEdit = oForm.Items.Item("SIJ6").Specific
                                Dim BPCode As String = oEdit.String

                                oEdit = oForm.Items.Item("104").Specific
                                Dim Dept As String = oEdit.String
                                oEdit = oForm.Items.Item("aej101").Specific
                                Dim BaseJobNo As String = oEdit.String

                                oEdit = oForm.Items.Item("SIJ16").Specific
                                LoadHAWB_MAWB("HAWB", oEdit.String, BPCode, BaseJobNo, Dept)
                            End If
                        ElseIf pVal.ItemUID = "MAWB" Then
                            oEdit = oForm.Items.Item("AEJ4").Specific
                            Dim QNo As String = oEdit.String
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                oEdit = oForm.Items.Item("SIJ6").Specific
                                Dim BPCode As String = oEdit.String
                                oEdit = oForm.Items.Item("104").Specific
                                Dim Dept As String = oEdit.String
                                oEdit = oForm.Items.Item("aej101").Specific
                                Dim BaseJobNo As String = oEdit.String

                                oEdit = oForm.Items.Item("SIJ16").Specific
                                LoadHAWB_MAWB("MAWB", oEdit.String, BPCode, BaseJobNo, Dept)
                            End If
                        ElseIf pVal.ItemUID = "185" Then

                            'Price Load
                            PriceLoad_AirExport(oForm)
                            ' LoadHandingCharge_AirImport(oForm, "AI")
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
                        ElseIf pVal.ItemUID = "PrintAWB" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf AWB_Report)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "PrintManf" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf MF_Report)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If

                        ElseIf pVal.ItemUID = "150" Then
                            'Packing List
                        ElseIf pVal.ItemUID = "151" Then
                            'Tally SHeet
                            'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            '    Dim trd As Threading.Thread
                            '    trd = New Threading.Thread(AddressOf Tally_sheet)
                            '    trd.IsBackground = True
                            '    trd.SetApartmentState(ApartmentState.STA)
                            '    trd.Start()
                            'End If
                            'AB_RP006_TS

                        ElseIf pVal.ItemUID = "PrintMAWB" Then
                            'Tally SHeet
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf PrintMAWB)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                            'AB_RP006_TS
                        ElseIf pVal.ItemUID = "PrintHAWB" Then
                            'Tally SHeet
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf PrintHAWB)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                            'AB_RP006_TS
                        ElseIf pVal.ItemUID = "PrintMF" Then
                            'Tally SHeet
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf PrintManifest)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                            'AB_RP006_TS

                        ElseIf pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            DocNumber_AI()
                            oEdit = oForm.Items.Item("SIJ18").Specific
                            oEdit.String = Format(Now.Date, "dd/MM/yy")
                            'oMatrix3 = oForm.Items.Item("SEJGR").Specific
                            'oMatrix3.AddRow()
                            'oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                            'oMatrix.AddRow()
                            'oMatrix1 = oForm.Items.Item("148").Specific
                            'oMatrix1.AddRow()
                        End If
                    End If
                    If pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "1" Then
                            DocNumber_AI()
                        End If
                    End If
                    If pVal.Before_Action = True And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
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
                                    oCombo = oForm.Items.Item("103").Specific
                                    If oCombo.Selected.Value = "Done" Then
                                        oEdit = oForm.Items.Item("aej101").Specific
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

                                Catch ex As Exception
                                End Try
                            End If
                        End If

                    End If
                    '------------
                    'ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '    If pVal.ItemUID = "1" Then
                    '        oCombo = oForm.Items.Item("103").Specific
                    '        If oCombo.Selected.Value = "Done" Then
                    '            oEdit = oForm.Items.Item("aej101").Specific
                    '            Dim BaseJobNo As String = oEdit.String
                    '            oEdit = oForm.Items.Item("SIJ16").Specific
                    '            Dim JobNo As String = oEdit.String
                    '            If BaseJobNo.Substring(0, 2) = "IN" Or BaseJobNo.Substring(0, 2) = "PR" Then
                    '                oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                oRecordSet1.DoQuery("UPDATE ODRF SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                    '                oRecordSet1.DoQuery("UPDATE OINV SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                    '                LoadGrid_BI(oForm)
                    '            End If
                    '        End If
                    '    End If
                    'End If
                End If

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
                                        oItem = oForm.Items.Item("126")
                                        oItem.Click(SAPbouiCOM.BoCellClickType.ct_Linked)
                                    ElseIf DocType = "Invoice" Then
                                        ' SBO_Application.StatusBar.SetText("Invoice Can't Be Open", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

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
                        If pVal.BeforeAction = False And pVal.ItemUID = "AWBGRID" Then
                            oGrid = oForm.Items.Item("AWBGRID").Specific
                            For F = 0 To oGrid.Rows.Count - 1
                                If oGrid.Rows.IsSelected(F) = True Then
                                    Dim Type As String = oGrid.DataTable.GetValue("Type", F)
                                    Dim HAWBNo As String = oGrid.DataTable.GetValue("HAWB No", F)
                                    Dim DocNum As String = oGrid.DataTable.GetValue("DocNum", F)
                                    If Type = "House" Then
                                        LoadFromXML("HAWB.srf", SBO_Application)
                                        oForm = SBO_Application.Forms.Item("UDO_F_HAWB_D")
                                        Try 'aswin
                                            CFL_BP_Supplier2(oForm, SBO_Application)
                                            oMatrix = oForm.Items.Item("CargoMAT").Specific
                                            oColumns = oMatrix.Columns
                                            oColumn = oColumns.Item("V_9")
                                            oColumn.ChooseFromListUID = "CFLBPV1"
                                            oColumn.ChooseFromListAlias = "CardCode"

                                            oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific

                                            oMatrix2 = oForm.Items.Item("150").Specific

                                            oMatrix3 = oForm.Items.Item("1000014").Specific

                                            ooption = oForm.Items.Item("optionbtn2").Specific
                                            ooption.GroupWith("optionbtn1")
                                            ooption = oForm.Items.Item("optionbtn4").Specific
                                            ooption.GroupWith("optionbtn3")

                                        Catch ex As Exception
                                        End Try
                                    ElseIf Type = "Master" Then
                                        LoadFromXML("MAWB.srf", SBO_Application)
                                        oForm = SBO_Application.Forms.Item("UDO_F_MAWB_D")
                                        Try 'aswin
                                            CFL_BP_Supplier2(oForm, SBO_Application)
                                            oMatrix = oForm.Items.Item("CargoMAT").Specific
                                            oColumns = oMatrix.Columns
                                            oColumn = oColumns.Item("V_9")
                                            oColumn.ChooseFromListUID = "CFLBPV1"
                                            oColumn.ChooseFromListAlias = "CardCode"
                                            oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                                            oMatrix2 = oForm.Items.Item("150").Specific
                                            oMatrix3 = oForm.Items.Item("1000014").Specific
                                            ooption = oForm.Items.Item("optionbtn2").Specific
                                            ooption.GroupWith("optionbtn1")
                                            ooption = oForm.Items.Item("optionbtn4").Specific
                                            ooption.GroupWith("optionbtn3")
                                        Catch ex As Exception
                                        End Try
                                    End If
                                    oForm.EnableMenu("1282", False)  '// Add New Record
                                    oForm.EnableMenu("1288", False)  '// Next Record
                                    oForm.EnableMenu("1289", False)  '// Pevious Record
                                    oForm.EnableMenu("1290", False)  '// First Record
                                    oForm.EnableMenu("1291", False)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oFolderItem = oForm.Items.Item("0_U_FD").Specific
                                    oFolderItem.Select()
                                    oItem = oForm.Items.Item("134")
                                    oItem.Enabled = True
                                    oEdit = oForm.Items.Item("134").Specific
                                    oEdit.Value = DocNum
                                    oEdit = oForm.Items.Item("17_U_E").Specific
                                    oEdit.String = oEdit.String


                                    oItem = oForm.Items.Item("134")
                                    oItem.Enabled = False
                                    oItem = oForm.Items.Item("1")
                                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If
                            Next

                        End If
                    Catch ex As Exception
                        Functions.WriteLog("Class:F_SE_JobOrder" + " Function:ItemEvent" + " Error Message:" + ex.ToString)
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
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
                        If pVal.ItemUID = "ChargeMat" And (pVal.ColUID = "V_8" Or pVal.ColUID = "V_6" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_0") Then
                            Try
                                oMatrix4 = oForm.Items.Item("ChargeMat").Specific

                                Dim UP As Double
                                Dim Qty As Integer
                                oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                Qty = oEdit.Value
                                oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(pVal.Row).Specific
                                UP = oEdit.Value
                                oEdit = oMatrix4.Columns.Item("V_0").Cells.Item(pVal.Row).Specific
                                oEdit.Value = (UP * Qty)

                                Dim i As Integer = 1
                                Dim amt As Double = 0
                                Dim TotAmt As Double = 0
                                Dim TotGST As Double = 0
                                Dim LineAmt As Double = 0
                                Dim TaxtAmt As Double = 0
                                For i = 1 To oMatrix4.RowCount
                                    oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        oEdit = oMatrix4.Columns.Item("V_0").Cells.Item(i).Specific
                                        amt = amt + oEdit.Value
                                        LineAmt = oEdit.Value
                                        oEdit = oMatrix4.Columns.Item("V_2").Cells.Item(i).Specific
                                        TaxtAmt = TAXPer(oEdit.String, Ocompany) * LineAmt * (1 / 100)
                                        TotGST = TotGST + TaxtAmt
                                    End If
                                Next
                                oEdit = oForm.Items.Item("1000013").Specific
                                oEdit.Value = amt
                                oEdit = oForm.Items.Item("1000015").Specific
                                oEdit.Value = TotGST
                                TotAmt = TotGST + amt
                                oEdit = oForm.Items.Item("1000017").Specific
                                oEdit.Value = TotAmt

                            Catch ex As Exception
                            End Try


                        ElseIf pVal.ItemUID = "CargoMat" And (pVal.ColUID = "V_6" Or pVal.ColUID = "V_3") Then
                            Try


                                Dim l As Integer = 0
                                Dim w As Integer = 0
                                Dim h As Integer = 0
                                Dim wt As Double = 0
                                Dim vol As Double = 0
                                Dim m3 As Double = 0
                                oMatrix5 = oForm.Items.Item("CargoMat").Specific

                                oEdit = oMatrix5.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    l = oEdit.Value
                                Else
                                    Exit Sub
                                    l = 0
                                End If
                                oEdit = oMatrix5.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    w = oEdit.Value
                                Else
                                    w = 0
                                End If
                                oEdit = oMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    h = oEdit.Value
                                Else
                                    h = 0
                                End If
                                Try
                                    m3 = ((l * w * h) / 6000)
                                Catch ex As Exception
                                    m3 = 0
                                End Try
                                oEdit = oMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                oEdit.String = m3
                                oEdit = oMatrix5.Columns.Item("V_6").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    wt = oEdit.Value
                                Else
                                    wt = 0
                                End If
                                If m3 > wt Then
                                    oEdit = oMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                                    oEdit.String = m3
                                Else
                                    oEdit = oMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                                    oEdit.String = wt
                                End If
                            Catch ex As Exception
                            End Try
                            Try
                                Dim i As Integer
                                Dim TotQty As Integer
                                Dim TotWt As Double
                                Dim TotVol As Double
                                oMatrix5 = oForm.Items.Item("CargoMat").Specific
                                For i = 1 To oMatrix5.RowCount
                                    oEdit = oMatrix5.Columns.Item("V_11").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        oEdit = oMatrix5.Columns.Item("V_8").Cells.Item(i).Specific
                                        TotQty = TotQty + oEdit.Value
                                        oEdit = oMatrix5.Columns.Item("V_6").Cells.Item(i).Specific
                                        TotWt = TotWt + oEdit.Value
                                        oEdit = oMatrix5.Columns.Item("V_1").Cells.Item(i).Specific
                                        TotVol = TotVol + oEdit.Value
                                    End If
                                    oEdit = oForm.Items.Item("154").Specific
                                    oEdit.Value = TotQty
                                    oEdit = oForm.Items.Item("1000021").Specific
                                    oEdit.Value = TotWt
                                    oEdit = oForm.Items.Item("1000022").Specific
                                    oEdit.Value = TotVol

                                Next
                            Catch ex As Exception

                            End Try
                        ElseIf pVal.ItemUID = "148" And (pVal.ColUID = "V_1" Or pVal.ColUID = "V_5") Then
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

                            Catch ex As Exception
                            End Try

                        End If
                    End If
                    If pVal.Before_Action = False And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        If pVal.ItemUID = "AEJ4" Then
                            oEdit = oForm.Items.Item("AEJ4").Specific
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
                            End If
                        End If
                        If pVal.ItemUID = "148" And pVal.ColUID = "V_8" Then
                            oMatrix1 = oForm.Items.Item("148").Specific
                            oEdit = oMatrix1.Columns.Item("V_8").Cells.Item(oMatrix1.RowCount).Specific
                            If oEdit.String <> "" Then
                                oEdit = oMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                oEdit.String = "SO"
                                oEdit = oMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                oEdit.String = "SGD"
                                oMatrix1.AddRow()
                                oMatrix1.ClearRowData(oMatrix1.RowCount)
                            End If
                        End If
                        If pVal.ItemUID = "ChargeMat" And pVal.ColUID = "V_10" Then
                            oMatrix4 = oForm.Items.Item("ChargeMat").Specific
                            oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(oMatrix4.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix4.AddRow()
                                oMatrix4.ClearRowData(oMatrix4.RowCount)
                            End If
                        End If
                        If pVal.ItemUID = "CargoMat" And pVal.ColUID = "V_11" Then
                            oMatrix5 = oForm.Items.Item("CargoMat").Specific
                            oEdit = oMatrix5.Columns.Item("V_11").Cells.Item(oMatrix5.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix5.AddRow()
                                oMatrix5.ClearRowData(oMatrix5.RowCount)
                            End If
                        End If
                        If pVal.ItemUID = "SIJ160" Then
                            oEdit = oForm.Items.Item("SIJ160").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("159").Specific
                                oEdit.String = Country_Code(ContCode, Ocompany)
                            End If
                        End If
                        If pVal.ItemUID = "SIJ1000009" Then
                            oEdit = oForm.Items.Item("SIJ1000009").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("39").Specific
                                oEdit.String = Country_Code(ContCode, Ocompany)
                            End If
                        End If
                        If pVal.ItemUID = "AEJ157" Then
                            oEdit = oForm.Items.Item("AEJ157").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("156").Specific
                                oEdit.String = City_Code(ContCode, Ocompany)
                            End If
                        End If
                        If pVal.ItemUID = "AEJ1000011" Then
                            oEdit = oForm.Items.Item("AEJ1000011").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("40").Specific
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
                    Try
                        oForm = SBO_Application.Forms.Item("AIRE_JOB")
                        oForm.Freeze(True)
                        oMatrix3 = oForm.Items.Item("AIJGR").Specific
                        oColumns = oMatrix3.Columns
                        oColumn = oColumns.Item("V_0")
                        oItem = oForm.Items.Item("AIJGR")
                        oItem.Width = 150
                        oItem.Height = 15
                        oColumn.Width = 130
                        oForm.Freeze(False)
                    Catch ex As Exception
                        oForm.Freeze(False)
                    End Try
                    Try
                        If pVal.BeforeAction = False Then
                            oForm.Items.Item("Rect").Width = oForm.Width - 50
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
                            End If
                            If pVal.ItemUID = "AEJ4" Then
                                Try
                                    oEdit = oForm.Items.Item("AEJ4").Specific
                                    oEdit.String = oDataTable.GetValue("DocNum", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "SIJ6" Then
                                Try
                                    oEdit = oForm.Items.Item("SIJ8").Specific
                                    oEdit.String = oDataTable.GetValue("CardName", 0)
                                    oEdit = oForm.Items.Item("SIJ6").Specific
                                    oEdit.String = oDataTable.GetValue("CardCode", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "127" Then
                                Try
                                    oEdit = oForm.Items.Item("129").Specific
                                    oEdit.String = oDataTable.GetValue("CardName", 0)
                                    oEdit = oForm.Items.Item("127").Specific
                                    oEdit.String = oDataTable.GetValue("CardCode", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "97" Then
                                Try
                                    oEdit = oForm.Items.Item("97").Specific
                                    oEdit.String = oDataTable.GetValue("ItemName", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "SIJDOMAT" And pVal.ColUID = "V_13" Then
                                oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardName", 0)
                                oEdit = oMatrix.Columns.Item("V_13").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardCode", 0)
                            End If
                            If pVal.ItemUID = "SIJDOMAT" And pVal.ColUID = "V_15" Then
                                oMatrix = oForm.Items.Item("SIJDOMAT").Specific
                                oEdit = oMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix.Columns.Item("V_15").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemCode", 0)
                            End If

                            If pVal.ItemUID = "ChargeMat" And pVal.ColUID = "V_10" Then
                                oMatrix4 = oForm.Items.Item("ChargeMat").Specific
                                oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                oEdit.String = "1"
                                oEdit = oMatrix4.Columns.Item("V_11").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("SalUnitMsr", 0)
                                oEdit = oMatrix4.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                                oEdit.String = "AI" 'oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemCode", 0)
                            End If
                            If pVal.ItemUID = "CargoMat" And pVal.ColUID = "V_11" Then
                                oMatrix5 = oForm.Items.Item("CargoMat").Specific
                                oEdit = oMatrix5.Columns.Item("V_10").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardName", 0)
                                oEdit = oMatrix5.Columns.Item("V_11").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardCode", 0)
                            End If

                            If pVal.ItemUID = "148" And pVal.ColUID = "V_8" Then
                                oMatrix1 = oForm.Items.Item("148").Specific
                                oEdit = oMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemCode", 0)
                            End If
                        End If
                    Catch ex As Exception
                        ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If

                '-----------End-----------------
            End If
        Catch ex As Exception
            If ex.Message <> "Form - Invalid Form" Then
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End Try
    End Sub
    'Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
    'End Sub
    Public Sub CreatePO()
        Try
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
    Public Sub LoadDraftInvoice(ByVal JobNo As String)
        Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet3 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim MAWBNo As String = ""
            Dim HAWBNo As String = ""
            oRecordSet1.DoQuery("SELECT T0.[U_AWBNo1] +T0.[U_AWBNo] FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
            MAWBNo = oRecordSet1.Fields.Item(0).Value.ToString
            oRecordSet1.DoQuery("SELECT T0.[U_HAWBNo] FROM [dbo].[@AB_AWB_H]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
            Dim F As Integer = 0
            For F = 1 To oRecordSet1.RecordCount
                If F = 1 Then
                    HAWBNo = oRecordSet1.Fields.Item(0).Value.ToString
                Else
                    HAWBNo = HAWBNo & "," & oRecordSet1.Fields.Item(0).Value.ToString
                End If
                oRecordSet1.MoveNext()
            Next
            oRecordSet3.DoQuery("SELECT T1.[U_ChWeight], T1.[U_Pieces], T0.[U_B1], T0.[U_FlighDate1], T0.[U_Nat] FROM [dbo].[@AB_AWB_M]  T0 , [dbo].[@AB_AWB_M2]  T1 WHERE T0.[DocEntry] = T1.[DocEntry] and T0.[U_JobNo]  ='" & JobNo & "'")
            oRecordSet1.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortN], T0.[U_TotPkg], T0.[U_ChrgWt],U_VessName [U_VessVoyage],'" & MAWBNo & "' [U_OBL],'" & HAWBNo & "' [U_HBL],U_FNo [U_F1],T0.[U_GKBNo],T0.[U_VessName] FROM [dbo].[@AB_AIRE_JOB_H] T0  WHERE T0.[U_JobNo]   ='" & JobNo & "'")
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
            'U_GKBNo
            oEdit = oform1.Items.Item("14").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_GKBNo").Value
            Try
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-133, 1)
                ' oform1.Freeze(True)
            Catch ex As Exception
                SBO_Application.ActivateMenuItem("6913")
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-133, 1)
                'oform1.Freeze(True)
            End Try
            oCombo = oform1.Items.Item("U_AB_Divsion").Specific
            oCombo.Select("AE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oEdit = oform1.Items.Item("U_AB_JobNo").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_JobNo").Value
            oEdit = oform1.Items.Item("U_AB_OriginNameC").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_LoadPortNC").Value
            oEdit = oform1.Items.Item("U_AB_DestNameC").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_DisPortN").Value
            ' T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_TotPkg], T0.[U_ChrgWt], T0.[U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_F1] FROM [dbo].[@AB_INT_HEADER]  T0 WHERE T0.[U_JobNo]   ='" & JobNo & "'")
            oEdit = oform1.Items.Item("U_AB_TotPkg").Specific
            oEdit.String = oRecordSet3.Fields.Item("U_Pieces").Value
            oEdit = oform1.Items.Item("U_AB_TotWT").Specific
            oEdit.String = oRecordSet3.Fields.Item("U_ChWeight").Value


            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessVoyage").Value
            oEdit = oform1.Items.Item("U_AB_MAWB").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_OBL").Value
            oEdit = oform1.Items.Item("U_AB_HAWB").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_HBL").Value
            oEdit = oform1.Items.Item("U_AB_FLT").Specific
            oEdit.String = oRecordSet3.Fields.Item("U_B1").Value
            Try
                'T0.[U_AB_ETDETA]
                oEdit = oform1.Items.Item("U_AB_ETDETA").Specific
                oEdit.String = oRecordSet3.Fields.Item("U_FlighDate1").Value
                oEdit = oform1.Items.Item("U_AB_Desc").Specific
                oEdit.String = oRecordSet3.Fields.Item("U_Nat").Value
            Catch ex As Exception

            End Try
            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessName").Value
            'U_VessName
            Dim QTNo As String = oRecordSet1.Fields.Item("U_QNo").Value
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[Dscription], T1.[Quantity], T1.[Price], T0.[DocCur],T1.[U_AB_Vendor],T1.U_AB_Cost,T1.[unitMsr] FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.[DocNum] ='" & QTNo & "'")
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
                oMatrix2.AddRow()
                SBO_Application.StatusBar.SetText("Please Waite Data is Loading.Line No-" & oRecordSet.RecordCount & " of-" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
                Catch ex As Exception

                End Try
                oRecordSet.MoveNext()
            Next
            SBO_Application.StatusBar.SetText("Data Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oform1.Title = "Billing Instruction"
            'SELECT T0.[U_CCode], T0.[U_CName], T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_JobNo], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ETD] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] =''")
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub LoadJobOrder(ByVal SQNO As String)
        Try
            'SELECT T0.[CardCode], T0.[CardName], T0.[U_AB_Divison], T0.[U_AB_TransTo], T0.[U_AB_Trnst], T0.[U_AB_VessCode], T0.[U_AB_VessName], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_Desc], T0.[U_AB_Validity], T0.[U_AB_Ttime], T0.[U_AB_Freq], T0.[U_AB_CARTotQt], T0.[U_AB_CARTotWt] FROM ORDR T0 WHERE T0.[DocNum] ='1'
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.DocNum,T0.[CardCode], T0.[CardName], T1.[Name], T0.[U_AB_CaType], T0.[U_SerLevel], T0.[U_AB_CarricerC], T0.[U_AB_CarrierN], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_OrginCodeC], T0.[U_AB_OriginNameC], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_DestCodeC], T0.[U_AB_DestNameC],T0.[U_AB_Divsion1], T0.[U_AB_JobNo],T0.[U_AB_VessName] FROM ORDR T0 Left JOIN OCPR T1 ON T0.CntctCode = T1.CntctCode WHERE  T0.[U_AB_Divsion] ='AE' and isnull( T0.[U_AB_Status] ,'')='Open' and T0.DocNum='" & SQNO & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oEdit = oForm.Items.Item("AEJ4").Specific
            oEdit.String = oRecordSet1.Fields.Item("DocNum").Value
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.String = oRecordSet1.Fields.Item("CardCode").Value
            oEdit = oForm.Items.Item("SIJ8").Specific
            oEdit.String = oRecordSet1.Fields.Item("CardName").Value
            oEdit = oForm.Items.Item("SJI10").Specific
            oEdit.String = oRecordSet1.Fields.Item("Name").Value
            Try
                oCombo = oForm.Items.Item("30").Specific
                oCombo.Select(oRecordSet1.Fields.Item("U_AB_CaType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000008").Specific
                oCombo.Select(oRecordSet1.Fields.Item("U_SerLevel").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
            End Try
            Try
                oEdit = oForm.Items.Item("AECC107").Specific
                oEdit.String = oRecordSet1.Fields.Item("U_AB_CarricerC").Value
            Catch ex As Exception
            End Try
            'T0.[U_AB_VessName]
            oEdit = oForm.Items.Item("97").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_VessName").Value

            oEdit = oForm.Items.Item("61").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_CarrierN").Value
            'T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName],
            oEdit = oForm.Items.Item("AEJ157").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OrginCodeC").Value
            oEdit = oForm.Items.Item("156").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_OriginNameC").Value
            oEdit = oForm.Items.Item("AEJ1000011").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestCodeC").Value
            oEdit = oForm.Items.Item("40").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_AB_DestNameC").Value

            Try

                oItem = oForm.Items.Item("104")
                oItem.Enabled = True
                oItem = oForm.Items.Item("aej101")
                oItem.Enabled = True
                oEdit = oForm.Items.Item("104").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_Divsion1").Value
                End If
                oEdit = oForm.Items.Item("aej101").Specific
                If oRecordSet1.Fields.Item("U_AB_JobNo").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                End If
                oCombo = oForm.Items.Item("103").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "IN" Or oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "PR" Then
                    oCombo.Select("Approved", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    oCombo.Select("NA", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                oItem = oForm.Items.Item("104")
                oItem.Enabled = False
                oItem = oForm.Items.Item("aej101")
                oItem.Enabled = False

            Catch ex As Exception

            End Try
            'oEdit = oForm.Items.Item("40").Specific
            'oEdit.String = oRecordSet1.Fields.Item(9).Value
            ''T0.[Address]
            'oEdit = oForm.Items.Item("96").Specific
            'oEdit.String = oRecordSet1.Fields.Item(2).Value & " " & oRecordSet1.Fields.Item("Address").Value



        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub MatrixLoad(ByVal DocNum As Integer, ByVal VenName As String, ByVal PONo As String)
        oForm = SBO_Application.Forms.Item("AIRE_JOB")
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
    'Private Sub LoadHandingCharge_AirImport(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
    '    Try
    '        'Dim DestCountry As String = ""
    '        'Dim DestCity As String = ""
    '        'oEdit = oForm.Items.Item("e13").Specific
    '        'DestCountry = oEdit.String
    '        'oEdit = oForm.Items.Item("ce13").Specific
    '        'DestCity = oEdit.String

    '        oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AI'")
    '        If oRecordSet1.RecordCount = 0 Then
    '            SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If
    '        oMatrix4 = oForm.Items.Item("ChargeMat").Specific
    '        oColumns = oMatrix4.Columns
    '        oMatrix4.Clear()
    '        SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '        oForm.Freeze(True)
    '        Dim i As Integer = 0
    '        For i = 1 To oRecordSet1.RecordCount
    '            oMatrix4.AddRow()

    '            'oEdit = oMatrix4.Columns.Item("U_Ven").Cells.Item(1).Specific
    '            'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
    '            oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
    '            oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
    '            oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(i).Specific
    '            oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
    '            oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Min").Value
    '            oEdit = oMatrix4.Columns.Item("V_11").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_PerKg").Value
    '            'U_Quantity
    '            oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
    '            oEdit = oMatrix4.Columns.Item("V_7").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value.ToString
    '            oRecordSet1.MoveNext()
    '        Next
    '        oForm.Freeze(False)
    '        SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Catch ex As Exception
    '        oForm.Freeze(False)
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP002_BS_AE.rpt")
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
            RptFrm.Text = "Shipping Order Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP004_DO_AE.rpt")
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



    Private Sub PrintMAWB()
        Try

            '1000004
            oEdit = oForm.Items.Item("SIJ16").Specific
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
            cryRpt.Load(sPath & "\GK_FM\AirMAWBReport.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@JobNo"
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
            RptFrm.Text = "MAWB Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub PrintHAWB()
        Try

            '1000004
            oEdit = oForm.Items.Item("SIJ16").Specific
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
            cryRpt.Load(sPath & "\GK_FM\AirHAWBReport.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@JobNo"
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
            RptFrm.Text = "HAWB Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub PrintManifest()
        Try

            '1000004
            oEdit = oForm.Items.Item("SIJ16").Specific
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
            cryRpt.Load(sPath & "\GK_FM\Consolidation Manifest.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@JobNo"
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
            RptFrm.Text = "Consolidation Manifest Report"
            RptFrm.TopMost = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            RptFrm.Refresh()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#End Region
#Region "Air Export Price"
    Private Sub PriceLoad_AirExport(ByVal oform As SAPbouiCOM.Form)

        Dim DestCountry As String = ""
        Dim DestCity As String = ""
        Dim ServiceLevel As String = ""
        Dim Carrier As String = ""

        Dim Cargo As String = ""
        Dim Weight As Double = 0
        Dim MinAmt As Double = 0
        oEdit = oform.Items.Item("SIJ160").Specific
        DestCountry = oEdit.String
        oEdit = oform.Items.Item("AEJ1000011").Specific
        DestCity = oEdit.String
        oEdit = oform.Items.Item("1000021").Specific
        Weight = oEdit.Value
        Dim BPCode As String = ""
        oEdit = oform.Items.Item("SIJ6").Specific
        BPCode = oEdit.Value
        Try
            oCombo = oform.Items.Item("30").Specific
            Cargo = oCombo.Selected.Value
        Catch ex As Exception
        End Try
        Try
            oCombo = oform.Items.Item("1000011").Specific
            Carrier = oCombo.Selected.Value
        Catch ex As Exception
        End Try
        Try
            oCombo = oform.Items.Item("1000008").Specific
            ServiceLevel = oCombo.Selected.Value
        Catch ex As Exception
        End Try
        Dim Str As String = "SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_MinBase], T1.[U_Neg45Base], T1.[U_45Base], T1.[U_100Base], T1.[U_300base], T1.[U_500Base], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500] FROM [dbo].[@AB_AIRSPECIAL_H]  T0 , [dbo].[@AB_AIRSPECIAL_D]  T1 WHERE T1.[Code] = T0.[Code]  and T0.[U_Division] ='AE' and  T1.[U_Country] ='" & DestCountry & "' and  T1.[U_Code] ='" & DestCity & "' and  T1.[U_Carrier] ='" & Carrier & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "' and  T1.[U_CargoType] ='" & Cargo & "' and T0.Code='" & BPCode & "'"
        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(Str)
        If oRecordSet.RecordCount <> 0 Then
            LoadHandingCharge_AirExport_SP(oform, "AE")
            LoadHandingCharge_AirImport(oform, "AE")
        Else
            LoadHandingCharge_AirExport(oform, "AE")
            LoadHandingCharge_AirImport(oform, "AE")
        End If


    End Sub
    Private Sub LoadHandingCharge_AirExport_SP(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            Dim DestCountry As String = ""
            Dim DestCity As String = ""
            Dim ServiceLevel As String = ""
            Dim Carrier As String = ""
            Dim Cargo As String = ""
            Dim Weight As Double = 0
            Dim MinAmt As Double = 0
            oEdit = oForm.Items.Item("SIJ160").Specific
            DestCountry = oEdit.String
            oEdit = oForm.Items.Item("AEJ1000011").Specific
            DestCity = oEdit.String
            oEdit = oForm.Items.Item("1000021").Specific
            Weight = oEdit.Value
            Try
                oCombo = oForm.Items.Item("30").Specific
                Cargo = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000011").Specific
                Carrier = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000008").Specific
                ServiceLevel = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Str As String = "SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_MinBase], T1.[U_Neg45Base], T1.[U_45Base], T1.[U_100Base], T1.[U_300base], T1.[U_500Base], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500] FROM [dbo].[@AB_AIRSPECIAL_H]  T0 , [dbo].[@AB_AIRSPECIAL_D]  T1 WHERE T1.[Code] = T0.[Code]  and T0.[U_Division] ='AE' and  T1.[U_Country] ='" & DestCountry & "' and  T1.[U_Code] ='" & DestCity & "' and  T1.[U_Carrier] ='" & Carrier & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "' and  T1.[U_CargoType] ='" & Cargo & "'"
            oRecordSet1.DoQuery(Str)
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix4 = oForm.Items.Item("ChargeMat").Specific
            oColumns = oMatrix4.Columns
            oMatrix4.Clear()
            SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix4.AddRow()
                oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(i).Specific
                oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
                If Weight < 45 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45").Value * Weight
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45Base").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_Neg45Base").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                ElseIf Weight >= 45 And Weight < 100 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45").Value * Weight
                    oEdit = oMatrix4.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45Base").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("V_3").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_45Base").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                ElseIf Weight >= 100 And Weight < 300 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100").Value * Weight
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100Base").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_100Base").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                ElseIf Weight >= 300 And Weight < 500 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300").Value * Weight
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300base").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_300base").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                ElseIf Weight >= 500 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500").Value * Weight
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500Base").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_500Base").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                End If

                oEdit = oMatrix4.Columns.Item("V_7").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_AirExport(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            Dim DestCountry As String = ""
            Dim DestCity As String = ""
            Dim ServiceLevel As String = ""
            Dim Carrier As String = ""
            Dim Cargo As String = ""
            Dim Weight As Double = 0
            Dim MinAmt As Double = 0
            oEdit = oForm.Items.Item("SIJ160").Specific
            DestCountry = oEdit.String
            oEdit = oForm.Items.Item("AEJ1000011").Specific
            DestCity = oEdit.String
            oEdit = oForm.Items.Item("1000021").Specific
            Weight = oEdit.Value
            Try
                oCombo = oForm.Items.Item("30").Specific
                Cargo = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000011").Specific
                Carrier = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000008").Specific
                ServiceLevel = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Str As String = "SELECT T0.[Code], T0.[Name], T1.[U_VendorCode], T1.[U_MIN], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500], T1.[U_Neg45Cost], T1.[U_45Cost], T1.[U_100Cost], T1.[U_300Cost], T1.[U_500Cost] FROM [dbo].[@AB_AIR_AIRFREIGHTH]  T0 , [dbo].[@AB_AIR_AIRFREIGHTD]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='AE' and  T1.[U_CargoType] ='" & Cargo & "' and  T1.[U_Carrier] ='" & Carrier & "' and  T1.[U_Country] ='" & DestCountry & "' and  T1.[U_Code] ='" & DestCity & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "'"
            oRecordSet1.DoQuery(Str)
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix4 = oForm.Items.Item("ChargeMat").Specific
            oColumns = oMatrix4.Columns
            oMatrix4.Clear()
            SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix4.AddRow()
                '    'oEdit = oMatrix4.Columns.Item("U_Ven").Cells.Item(1).Specific
                '    'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                '    oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                '    oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                '    oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(i).Specific
                '    oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                '    oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(i).Specific
                '    oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
                '    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                '    oEdit.String = oRecordSet1.Fields.Item("U_Min").Value
                '    oEdit = oMatrix4.Columns.Item("V_11").Cells.Item(i).Specific
                '    oEdit.String = oRecordSet1.Fields.Item("U_PerKg").Value
                '    'U_Quantity
                '    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                '    oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                '    oEdit = oMatrix4.Columns.Item("V_7").Cells.Item(i).Specific
                '    oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value.ToString

                'oEdit = oMatrix4.Columns.Item("U_Ven").Cells.Item(1).Specific
                'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("Code").Value
                oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("Name").Value
                oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(i).Specific
                oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
                If Weight < 45 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_Neg45").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45Cost").Value * Weight
                ElseIf Weight >= 45 And Weight < 100 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_45").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45Cost").Value * Weight
                ElseIf Weight >= 100 And Weight < 300 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_100").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100Cost").Value * Weight
                ElseIf Weight >= 300 And Weight < 500 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_300").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300Cost").Value * Weight
                ElseIf Weight >= 500 Then
                    oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_500").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500Cost").Value * Weight
                End If

                oEdit = oMatrix4.Columns.Item("V_7").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value.ToString
                oRecordSet1.MoveNext()
            Next

            '********************HANDLING CHARGES*************************
            Try



                oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AI'")
                If oRecordSet1.RecordCount = 0 Then
                    SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                oMatrix4 = oForm.Items.Item("ChargeMat").Specific
                oColumns = oMatrix4.Columns
                oMatrix4.Clear()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oForm.Freeze(True)
                Dim Min As Double = 0
                Dim Perkg As Double = 0
                Dim UnitPrice As Double = 0
                Dim wt As Double = 0
                oEdit = oForm.Items.Item("cce3").Specific
                wt = oEdit.Value

                For i = 1 To oRecordSet1.RecordCount
                    If oMatrix4.RowCount = 0 Then
                        oMatrix4.AddRow()
                    Else
                        oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                        If oEdit.String <> "" Then
                            oMatrix4.AddRow()
                        End If
                    End If
                    'oEdit = oMatrix4.Columns.Item("U_Ven").Cells.Item(1).Specific
                    'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                    oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(oMatrix4.RowCount()).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                    oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(oMatrix4.RowCount()).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                    oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(oMatrix4.RowCount()).Specific
                    oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
                    Min = oRecordSet1.Fields.Item("U_Min").Value
                    Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
                    If Perkg = 0 Then
                        UnitPrice = Min
                    Else
                        UnitPrice = Perkg * wt
                        If UnitPrice < Min Then
                            UnitPrice = Min
                        End If
                    End If
                    oEdit = oMatrix4.Columns.Item("V_14").Cells.Item(oMatrix4.RowCount()).Specific
                    oEdit.String = UnitPrice 'oRecordSet1.Fields.Item("U_PerKg").Value

                    oEdit = oMatrix4.Columns.Item("U_AB_Cost").Cells.Item(oMatrix4.RowCount()).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                    oEdit = oMatrix4.Columns.Item("U_AB_Vendor").Cells.Item(oMatrix4.RowCount()).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value.ToString
                    oRecordSet1.MoveNext()
                Next
            Catch ex As Exception

            End Try
            '********************END HANDLING CHARGES*************************
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_AirImport(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            'Dim DestCountry As String = ""
            'Dim DestCity As String = ""
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            'oEdit = oForm.Items.Item("ce13").Specific
            'DestCity = oEdit.String

            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AE'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix4 = oForm.Items.Item("ChargeMat").Specific
            oColumns = oMatrix4.Columns
            'oMatrix4.Clear()
            SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                If oMatrix4.RowCount = 0 Then
                    oMatrix4.AddRow()
                Else
                    oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        oMatrix4.AddRow()
                    End If
                End If

                'oEdit = oMatrix4.Columns.Item("U_Ven").Cells.Item(1).Specific
                'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix4.Columns.Item("V_6").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Min").Value
                oEdit = oMatrix4.Columns.Item("V_11").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_PerKg").Value
                'U_Quantity
                oEdit = oMatrix4.Columns.Item("V_3").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix4.Columns.Item("V_7").Cells.Item(oMatrix4.RowCount()).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#End Region

    Private Sub LoadHAWB_MAWB(ByVal Type As String, ByVal JobNo As String, ByVal BPCode As String, ByVal BaseJobNo As String, ByVal Dept As String)
        Try
            Dim formCreationParams As SAPbouiCOM.FormCreationParams
            Dim oXmlDoc As New Xml.XmlDocument
            Dim sPath As String
            Dim sFileName As String = ""
            Dim Fno As String = ""
            Dim FDate As String = ""
            Dim DestCode As String = ""
            Dim DestCity As String = ""
            Dim OrgCode As String = ""
            Dim OrgCity As String = ""
            Dim Carrier As String = ""
            Dim SHipType As String = ""
            Try
                oEdit = oForm.Items.Item("95").Specific
                Fno = oEdit.String
                oEdit = oForm.Items.Item("1000002").Specific
                FDate = oEdit.Value
                oEdit = oForm.Items.Item("AEJ1000011").Specific
                DestCode = oEdit.Value
                oEdit = oForm.Items.Item("40").Specific
                DestCity = oEdit.Value
                oEdit = oForm.Items.Item("AEJ157").Specific
                OrgCode = oEdit.Value
                oEdit = oForm.Items.Item("156").Specific
                OrgCity = oEdit.Value
                oEdit = oForm.Items.Item("AECC107").Specific
                Carrier = oEdit.String
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000018").Specific
                SHipType = oCombo.Selected.Description
            Catch ex As Exception
            End Try
            If SHipType = "" Then
                SBO_Application.StatusBar.SetText("Select Shipment Type in Home Page", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[DocEntry] FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
            If Type = "HAWB" Then
                If oRecordSet1.RecordCount = 0 Then
                    SBO_Application.StatusBar.SetText("First Create MAWB No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            End If
            If Type = "MAWB" Then
                If oRecordSet1.RecordCount > 0 Then
                    SBO_Application.StatusBar.SetText("MAWB Already Created", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            End If
            If Type = "HAWB" And SHipType = "D" Then
                SBO_Application.StatusBar.SetText("For Dierect Shipemt Type HAWB Not Applicable", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Try
                Select Case Type
                    Case "HAWB"
                        sFileName = "HAWB.srf"
                    Case "MAWB"
                        sFileName = "MAWB.srf"
                End Select
                sPath = IO.Directory.GetParent(Application.StartupPath).ToString
                oXmlDoc.Load(sPath & "\GK_FM\" & sFileName)
                formCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                formCreationParams.XmlData = oXmlDoc.InnerXml
                oForm = SBO_Application.Forms.AddEx(formCreationParams)
                'LoadFromXML(sFileName, SBO_Application)
                'formCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                'oForm = SBO_Application.Forms.AddEx(formCreationParams)
                'Dim oF_AWB As F_AWB
                'oF_AWB = New F_AWB(Ocompany, SBO_Application)
                oForm.EnableMenu("1282", False)  '// Add New Record
                oForm.EnableMenu("1288", False)  '// Next Record
                oForm.EnableMenu("1289", False)  '// Pevious Record
                oForm.EnableMenu("1290", False)  '// First Record
                oForm.EnableMenu("1291", False)  '// Last record
                'oF_AWB.AWB_Bind(oForm, JobNo)
                oForm.PaneLevel = 1
                AWBForm = oForm
                Select Case oForm.BusinessObject.Type
                    Case "HAWB"
                        hawbForm = oForm
                    Case "MAWB"
                        mawbForm = oForm
                End Select
                '0_U_E
                oEdit = oForm.Items.Item("0_U_E").Specific
                oEdit.String = JobNo


                ooption = oForm.Items.Item("optionbtn2").Specific
                ooption.GroupWith("optionbtn1")
                ooption = oForm.Items.Item("optionbtn1").Specific
                ooption.Selected = True
                ooption = oForm.Items.Item("optionbtn4").Specific
                ooption.GroupWith("optionbtn3")
                ooption = oForm.Items.Item("optionbtn3").Specific
                ooption.Selected = True


                oForm.Freeze(True)
                'oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRecordSet1.DoQuery("SELECT UPPER(T0.[CompnyName]), UPPER(T0.[CompnyAddr]) FROM OADM T0")
                'oEdit = oForm.Items.Item("18_U_E").Specific
                'oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString & vbCrLf & oRecordSet1.Fields.Item(1).Value.ToString
                'oEdit = oForm.Items.Item("21_U_E").Specific
                'oEdit.String = BPName(BPCode, Ocompany) & vbCrLf & BPAddress(BPCode, Ocompany)
                If Type = "MAWB" Then
                    oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet1.DoQuery("SELECT UPPER(T0.[CompnyName]), UPPER(T0.[CompnyAddr]),(Select Top 1 UPPER(T1.[Name]) from OCRY T1 where T1.Code=T0.Country) 'Country' FROM OADM T0")
                    oEdit = oForm.Items.Item("122").Specific
                    oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString
                    oEdit = oForm.Items.Item("23_U_E").Specific
                    oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString
                    oEdit = oForm.Items.Item("24_U_E").Specific
                    oEdit.String = oRecordSet1.Fields.Item(2).Value.ToString
                    oEdit = oForm.Items.Item("126").Specific
                    oEdit.String = oRecordSet1.Fields.Item(2).Value.ToString
                    oEdit = oForm.Items.Item("124").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    CFL_BP_Customer(oForm, SBO_Application)
                    oEdit = oForm.Items.Item("17_U_E").Specific
                    oEdit.ChooseFromListUID = "CFLBPC"
                    oEdit.ChooseFromListAlias = "CardCode"
                    CFL_BP_Customer1(oForm, SBO_Application)
                    oEdit = oForm.Items.Item("20_U_E").Specific
                    oEdit.ChooseFromListUID = "CFLBPC1"
                    oEdit.ChooseFromListAlias = "CardCode"
                    oEdit = oForm.Items.Item("131").Specific
                    oEdit.String = FightPreFix(Carrier, Ocompany)
                    oEdit = oForm.Items.Item("132").Specific
                    oEdit.String = "SIN"
                    'oEdit = oForm.Items.Item("19_U_E").Specific
                    'oEdit.String = DocNumber_MAWB()
                    oEdit = oForm.Items.Item("22_U_E").Specific
                    oEdit.String = Fightaddress(Carrier, Ocompany)
                    oPict = oForm.Items.Item("1000007").Specific
                    oPict.Picture = FightLogo(Carrier, Ocompany)
                    oEdit = oForm.Items.Item("35_U_E").Specific
                    oEdit.String = "SGD"
                    oEdit = oForm.Items.Item("30_U_E").Specific
                    oEdit.String = DestCode
                    oEdit = oForm.Items.Item("147").Specific
                    oEdit.String = Fno
                    oEdit = oForm.Items.Item("136").Specific
                    oEdit.Value = FDate
                    '29_U_E
                    oEdit = oForm.Items.Item("29_U_E").Specific
                    oEdit.String = OrgCity
                    oEdit = oForm.Items.Item("33_U_E").Specific
                    oEdit.String = DestCity
                    oEdit = oForm.Items.Item("128").Specific
                    oEdit.String = UserName(Ocompany.UserName.ToUpper, Ocompany)
                    CFL_BP_Supplier2(oForm, SBO_Application)
                    oMatrix = oForm.Items.Item("CargoMAT").Specific
                    oColumns = oMatrix.Columns
                    oMatrix.AddRow()
                    'oColumn = oColumns.Item("V_9")
                    'oColumn.ChooseFromListUID = "CFLBPV1"
                    'oColumn.ChooseFromListAlias = "CardCode"
                    oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                    oMatrix1.AddRow()
                    oMatrix2 = oForm.Items.Item("150").Specific
                    oMatrix2.AddRow(6)
                    oMatrix3 = oForm.Items.Item("1000014").Specific
                    oMatrix3.AddRow()
                    'puni
                    Try
                        'oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRecordSet1.DoQuery("SELECT T0.[U_AB_Divsion1],T0.U_AB_JobNo FROM ORDR T0  WHERE T0.[DocNum] ='" & SQNO & "' and T0.[U_AB_Divsion]='AE'")
                        'Dim jno As String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                        'If jno = "" Then
                        '    Exit Try
                        'End If
                        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If BaseJobNo.Substring(0, 2) = "IN" Then
                            oRecordSet.DoQuery("SELECT T0.[U_Shipp], T0.[U_COns], T0.[U_Rem] FROM [dbo].[@AB_INT_HEADER]  T0 WHERE T0.[U_JobNo] ='" & BaseJobNo & "'")
                        ElseIf BaseJobNo.Substring(0, 2) = "PR" Then
                            oRecordSet.DoQuery("SELECT T0.[U_Shipp], T0.[U_COns], T0.[U_Rem] FROM [dbo].[@AB_PRO_HEADER]  T0 WHERE T0.[U_JobNo] ='" & BaseJobNo & "'")
                        End If

                        oEdit = oForm.Items.Item("18_U_E").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_Shipp").Value
                        oEdit = oForm.Items.Item("21_U_E").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_COns").Value
                        'oEdit = oForm.Items.Item("45").Specific
                        'oEdit.String = oRecordSet.Fields.Item("U_Rem").Value
                    Catch ex As Exception

                    End Try
                End If
                If Type = "HAWB" Then
                    oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet1.DoQuery("SELECT T0.[U_AWBNo1], T0.[U_AWBNo2], T0.[U_AWBNo] FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
                    oEdit = oForm.Items.Item("131").Specific
                    oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString
                    oEdit = oForm.Items.Item("132").Specific
                    oEdit.String = oRecordSet1.Fields.Item(1).Value.ToString
                    oEdit = oForm.Items.Item("19_U_E").Specific
                    oEdit.String = oRecordSet1.Fields.Item(2).Value.ToString
                    oEdit = oForm.Items.Item("126").Specific
                    oEdit.String = oRecordSet1.Fields.Item(2).Value.ToString
                    oEdit = oForm.Items.Item("124").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet1.DoQuery("SELECT UPPER(T0.[CompnyName]), UPPER(T0.[CompnyAddr]),(Select Top 1 UPPER(T1.[Name]) from OCRY T1 where T1.Code=T0.Country) 'Country' FROM OADM T0")
                    oEdit = oForm.Items.Item("122").Specific
                    oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString
                    oEdit = oForm.Items.Item("23_U_E").Specific
                    oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString
                    oEdit = oForm.Items.Item("23_U_E").Specific
                    oEdit.String = oRecordSet1.Fields.Item(2).Value.ToString
                    CFL_BP_Customer(oForm, SBO_Application)
                    oEdit = oForm.Items.Item("17_U_E").Specific
                    oEdit.ChooseFromListUID = "CFLBPC"
                    oEdit.ChooseFromListAlias = "CardCode"
                    CFL_BP_Customer1(oForm, SBO_Application)
                    oEdit = oForm.Items.Item("20_U_E").Specific
                    oEdit.ChooseFromListUID = "CFLBPC1"
                    oEdit.ChooseFromListAlias = "CardCode"
                    'oEdit = oForm.Items.Item("131").Specific
                    'oEdit.String = FightPreFix(Carrier, Ocompany)
                    'oEdit = oForm.Items.Item("132").Specific
                    'oEdit.String = "SIN"
                    Carrier = "GK"
                    oEdit = oForm.Items.Item("1000016").Specific
                    oEdit.String = DocNumber_HAWB()
                    oEdit = oForm.Items.Item("22_U_E").Specific
                    oEdit.String = Fightaddress(Carrier, Ocompany)
                    oPict = oForm.Items.Item("1000007").Specific
                    oPict.Picture = FightLogo(Carrier, Ocompany)
                    oEdit = oForm.Items.Item("35_U_E").Specific
                    oEdit.String = "SGD"
                    oEdit = oForm.Items.Item("30_U_E").Specific
                    oEdit.String = DestCode
                    oEdit = oForm.Items.Item("147").Specific
                    oEdit.String = Fno
                    oEdit = oForm.Items.Item("136").Specific
                    oEdit.Value = FDate
                    oEdit = oForm.Items.Item("29_U_E").Specific
                    oEdit.String = OrgCity
                    oEdit = oForm.Items.Item("33_U_E").Specific
                    oEdit.String = DestCity

                    oEdit = oForm.Items.Item("128").Specific
                    oEdit.String = UserName(Ocompany.UserName.ToUpper, Ocompany)
                    CFL_BP_Supplier2(oForm, SBO_Application)
                    oMatrix = oForm.Items.Item("CargoMAT").Specific
                    oColumns = oMatrix.Columns
                    oMatrix.AddRow()
                    'oColumn = oColumns.Item("V_9")
                    'oColumn.ChooseFromListUID = "CFLBPV1"
                    'oColumn.ChooseFromListAlias = "CardCode"

                    oMatrix1 = oForm.Items.Item("AWB_Mtr1").Specific
                    oMatrix1.AddRow()
                    oMatrix2 = oForm.Items.Item("150").Specific
                    oMatrix2.AddRow(6)
                    oMatrix3 = oForm.Items.Item("1000014").Specific
                    oMatrix3.AddRow()

                End If

                oFolderItem = oForm.Items.Item("0_U_FD").Specific
                oFolderItem.Select()
                oForm.Freeze(False)
            Catch ex As Exception
                oForm.Freeze(False)
                SBO_Application.MessageBox(ex.Message)
            End Try
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SBO_Application_RghtClick(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
        If eventInfo.ItemUID = "AWB_Mtr1" Or eventInfo.ItemUID = "AWB_Mtr2" Then
            Try
                Exit Try
                If eventInfo.BeforeAction Then
                    Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                    Dim oMenus As SAPbouiCOM.Menus = Nothing

                    matrixUID = eventInfo.ItemUID
                    rowDelete = eventInfo.Row
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
                    oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oMenuItem = SBO_Application.Menus.Item("1280") 'Data
                    oMenus = oMenuItem.SubMenus

                    If Not SBO_Application.Menus.Exists("DeleteRow") Then
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "DeleteRow"
                        oCreationPackage.String = "Delete Row"
                        oCreationPackage.Enabled = True
                        oMenus.AddEx(oCreationPackage)
                    End If

                    If Not SBO_Application.Menus.Exists("ClearMatrix") Then
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "ClearMatrix"
                        oCreationPackage.String = "Clear Matrix"
                        oCreationPackage.Enabled = True
                        oMenus.AddEx(oCreationPackage)
                    End If
                Else
                    If SBO_Application.Menus.Exists("DeleteRow") Then
                        SBO_Application.Menus.RemoveEx("DeleteRow")
                    End If
                    If SBO_Application.Menus.Exists("ClearMatrix") Then
                        SBO_Application.Menus.RemoveEx("ClearMatrix")
                    End If
                End If
            Catch ex As Exception
                SBO_Application.MessageBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef menuEvent As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If (menuEvent.MenuUID = "43530" Or menuEvent.MenuUID = "1288" Or menuEvent.MenuUID = "1289" Or menuEvent.MenuUID = "1290" Or menuEvent.MenuUID = "1291") And menuEvent.InnerEvent = False And menuEvent.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "AIRE_JOB" Then
                    'If menuEvent.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = True
                    'End If



                End If
            End If
            If menuEvent.MenuUID = "1282" And menuEvent.InnerEvent = False And menuEvent.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "AIRE_JOB" Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = False
                    DocNumber_AI()
                    oEdit = oForm.Items.Item("SIJ18").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")

                End If
            End If
        Catch ex As Exception

        End Try
        If Not menuEvent.BeforeAction Then
            Try
                Dim matrix As SAPbouiCOM.Matrix
                If menuEvent.MenuUID = "DeleteRow" Then
                    matrix = AWBForm.Items.Item(matrixUID).Specific
                    If rowDelete <> 0 And rowDelete <> matrix.RowCount Then
                        matrix.DeleteRow(rowDelete)
                        rowDelete = 0
                        If AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                ElseIf menuEvent.MenuUID = "ClearMatrix" Then
                    matrix = AWBForm.Items.Item(matrixUID).Specific
                    matrix.Clear()
                    matrix.AddRow(1, 0)
                    matrix.FlushToDataSource()
                    If AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                End If

            Catch ex As Exception
                SBO_Application.MessageBox(ex.Message)

            End Try
        End If
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
    Private Sub LoadDeliveryOrder(ByVal oform As SAPbouiCOM.Form)
        Try
            Dim oform1 As SAPbouiCOM.Form
            LoadFromXML("GoodsIssue.srf", SBO_Application)
            oform1 = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
            oEdit = oform.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            oEdit = oform1.Items.Item("35").Specific
            oEdit.String = JobNo
            oEdit = oform.Items.Item("97").Specific
            Dim VesslNo As String = oEdit.String
            oEdit = oform1.Items.Item("20").Specific
            oEdit.String = VesslNo

            'oEdit = oform.Items.Item("1000036").Specific
            'Dim AWBNo As String = oEdit.String
            'oEdit = oform1.Items.Item("22").Specific
            'oEdit.String = VesslNo

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
            '1000036
            Dim MAWBNo As String = ""
            Try
                oCombo = oform.Items.Item("116").Specific
                If oCombo.Selected.Description = "H" Then
                    oEdit = oform.Items.Item("118").Specific
                    MAWBNo = oEdit.String
                    oEdit = oform1.Items.Item("22").Specific
                    oEdit.String = MAWBNo
                ElseIf oCombo.Selected.Description = "M" Then
                    oEdit = oform.Items.Item("1000036").Specific
                    MAWBNo = oEdit.String
                    oEdit = oform1.Items.Item("22").Specific
                    oEdit.String = MAWBNo
                End If
            Catch ex As Exception

            End Try

            oMatrix = oform1.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow()
            End If
            oCombo = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            ComboLoad_Unit(oform, oCombo)
            Dim oDO As New F_GoodsIssue
            oDO.GoodsIssue_Bind(oform1, SBO_Application)
            Exit Sub
            Dim oRecordSet_GR As SAPbobsCOM.Recordset
            oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T1.[U_VenCode], T1.[U_VenName], '' [U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId] FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[U_MAWBNo] ='" & MAWBNo & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
            If oRecordSet_GR.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            oEdit = oform1.Items.Item("20").Specific
            oEdit.String = oRecordSet_GR.Fields.Item(1).Value
            oEdit = oform1.Items.Item("22").Specific
            oEdit.String = oRecordSet_GR.Fields.Item(2).Value
            oEdit = oform1.Items.Item("24").Specific
            oEdit.String = oRecordSet_GR.Fields.Item(3).Value
            oEdit = oform1.Items.Item("33").Specific
            oEdit.String = oRecordSet_GR.Fields.Item(5).Value
            Try
                oEdit = oform1.Items.Item("31").Specific
                If oEdit.String = "" Then
                    oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
                Else
                    oEdit.String = oEdit.String & "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
                End If
            Catch ex As Exception
            End Try
            oform1 = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
            oMatrix = oform1.Items.Item("29").Specific
            Dim i As Integer
            For i = 1 To oRecordSet_GR.RecordCount
                If oMatrix.RowCount = 0 Then
                    oMatrix.AddRow()
                End If
                oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(9).Value
                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(19).Value
                oEdit = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(10).Value
                Try
                    oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                    oCombo.Select(oRecordSet_GR.Fields.Item(11).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try

                oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(12).Value
                oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(13).Value
                oEdit = oMatrix.Columns.Item("V_4").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(14).Value
                oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(15).Value
                Try
                    oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                    oEdit.String = oRecordSet_GR.Fields.Item(16).Value
                Catch ex As Exception

                End Try

                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(18).Value
                oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(17).Value
                oEdit = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(20).Value
                oEdit = oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(21).Value

                oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
                If oEdit.String = "" Then
                    oEdit.String = oRecordSet_GR.Fields.Item(6).Value
                End If
                oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(7).Value
                oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item(0).Value
                oMatrix.AddRow()
                oRecordSet_GR.MoveNext()
            Next

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
            oPV.PV_Bind(oform1, SBO_Application, "AI", Ocompany)

        Catch ex As Exception
            Functions.WriteLog("Class:F_SI_JobOrder" + " Function:LoadPaymentVoucher" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Function DocNumber_MAWB() As String
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT CONVERT(VARCHAR(25),DATEADD(dd,-(DAY(DATEADD(mm,1,getdate()))),DATEADD(mm,1,getdate())),101)")
            tdt = oRecordSet1.Fields.Item(0).Value
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            Dim DocNumLen As Integer
            Dim MAWBNo As String = ""
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                MAWBNo = "00000001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                MAWBNo = "0000000" & DocNum
            ElseIf DocNumLen = 2 Then
                MAWBNo = "000000" & DocNum
            ElseIf DocNumLen = 3 Then
                MAWBNo = "00000" & DocNum
            ElseIf DocNumLen = 4 Then
                MAWBNo = "0000" & DocNum
            ElseIf DocNumLen = 5 Then
                MAWBNo = "000" & DocNum
            ElseIf DocNumLen = 6 Then
                MAWBNo = "00" & DocNum
            ElseIf DocNumLen = 7 Then
                MAWBNo = "0" & DocNum
            Else
                MAWBNo = DocNum
            End If
            Return MAWBNo

        Catch ex As Exception
            Return ""
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Function

    Public Function DocNumber_HAWB() As String
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT CONVERT(VARCHAR(25),DATEADD(dd,-(DAY(DATEADD(mm,1,getdate()))),DATEADD(mm,1,getdate())),101)")
            tdt = oRecordSet1.Fields.Item(0).Value
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_AWB_H]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            Dim DocNumLen As Integer
            Dim MAWBNo As String = ""
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                MAWBNo = "GK" & Format(Now.Date, "yy") & "H" & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                MAWBNo = "GK" & Format(Now.Date, "yy") & "H" & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                MAWBNo = "GK" & Format(Now.Date, "yy") & "H" & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                MAWBNo = "GK" & Format(Now.Date, "yy") & "H" & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                MAWBNo = "GK" & Format(Now.Date, "yy") & "H" & "0" & DocNum
            ElseIf DocNumLen = 5 Then
                MAWBNo = "GK" & Format(Now.Date, "yy") & "H" & DocNum
            Else
                MAWBNo = DocNum
            End If
            Return MAWBNo

        Catch ex As Exception
            Return ""
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Function
#Region "Reports"
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
            cryRpt.Load(sPath & "\GK_FM\DO_WHMS_AE.rpt")
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
    Public Sub Payment_Voucher()
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP005_PV_AE.rpt")
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
    Private Sub AWB_Report()
        Try
            Dim DocType As String = ""
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("AWBGRID").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    DocType = oGrid.DataTable.GetValue("Type", F)
                    Exit For
                End If
            Next
            'AB_RP007_BI_INVOICE

            '  Exit Sub
            'oEdit = oForm.Items.Item("DO4").Specific
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select AWB No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            SBO_Application.StatusBar.SetText("Retrieving Data!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument


            Dim cryRpt1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            If DocType = "Master" Then
                cryRpt.Load(sPath & "\GK_FM\MAWB_PREPRINT_2.rpt")
            Else
                cryRpt.Load(sPath & "\GK_FM\MAWB_PREPRINT_3.rpt")
            End If


            'cryRpt.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
            '  cryRpt.PrintToPrinter(1, True, 1, 1)

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

            'cryRpt.SetParameterValue("@DocKey", DocNum, "MAWB_PREPRINT_1")

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
            '-----R&D
            cryRpt.PrintOptions.PrinterName = "Tally Dascom 1430"

            Dim doctoprint As New System.Drawing.Printing.PrintDocument()
            doctoprint.PrinterSettings.PrinterName = "Tally Dascom 1430"

            Dim rawKind As Integer = 0
            Dim i As Integer
            For i = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
                If doctoprint.PrinterSettings.PaperSizes(i).PaperName = "Fanfold 210 mm x 12 in" Then
                    rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("kind", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
                    Exit For
                End If

            Next
            cryRpt.PrintOptions.PaperSize = rawKind



            '-----R&D
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            RptFrm.Text = "AWB Report"
            RptFrm.TopMost = True
            RptFrm.Activate()



            'RptFrm.CrystalReportViewer1.PrintReport()
            '   System.Drawing.Printing.PrintDocument(printDoc = New System.Drawing.Printing.PrintDocument())
            'Try
            '    Dim rawKind As Integer = 0

            
            'Catch ex As Exception

            'End Try
            'cryRpt.PrintToPrinter(1, False, 0, 0)
            RptFrm.ShowDialog()

            '
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub MF_Report()
        Try
            Dim DocNum As String = ""
            oGrid = oForm.Items.Item("AWBGRID").Specific
            For F = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(F) = True Then
                    DocNum = oGrid.DataTable.GetValue("DocNum", F)
                    Exit For
                End If
            Next
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("Select MAWB No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

            cryRpt.Load(sPath & "\GK_FM\CargoManifest.rpt")

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
            RptFrm.Text = "CargoManifest Report"
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
