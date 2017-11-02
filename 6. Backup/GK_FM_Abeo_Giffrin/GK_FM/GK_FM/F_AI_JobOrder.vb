Imports System.Diagnostics.Process
Imports System.Threading
Imports System.IO
Public Class F_AI_JobOrder
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
    Public Sub AI_Job_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try

            oForm.DataSources.DataTables.Add("DOAI")
            oForm.DataSources.DataTables.Add("PVAI")
            oForm.DataSources.DataTables.Add("BIAI")
            oForm.DataSources.DataTables.Add("REFAI")

            oForm.PaneLevel = 1
            DocNumber_AI()
            ShippigNameLoad()
            oItem = oForm.Items.Item("153")
            oItem.Enabled = False
            oEdit = oForm.Items.Item("SIJ18").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            ''ooption = oForm.Items.Item("28").Specific
            ''ooption.GroupWith("SIJ1000001")1000011
            ooption = oForm.Items.Item("85").Specific
            ooption.GroupWith("86")
            '---------
            'ooption = oForm.Items.Item("1000010").Specific
            'ooption.GroupWith("1000011")
            'ooption = oForm.Items.Item("171").Specific
            'ooption.GroupWith("172")
            '------------
            'oCombo = oForm.Items.Item("42").Specific
            'ComboLoad_Whsc(oForm, oCombo)
            oCombo = oForm.Items.Item("30").Specific
            ComboLoad_ContainerType(oForm, oCombo)
            'oCombo = oForm.Items.Item("1000002").Specific
            'ComboLoad_Currency(oForm, oCombo)
            'oCombo = oForm.Items.Item("1000011").Specific
            'ComboLoad_Carrier(oForm, oCombo)

            oCombo = oForm.Items.Item("163").Specific
            ComboLoad_WeightUnit(oForm, oCombo)
            oCombo = oForm.Items.Item("169").Specific
            ComboLoad_WeightUnit(oForm, oCombo)

            'oCombo = oForm.Items.Item("178").Specific
            'ComboLoad_WeightUnit(oForm, oCombo)
            oCombo = oForm.Items.Item("166").Specific
            ComboLoad_VolumeUnit(oForm, oCombo)
            'oCombo = oForm.Items.Item("175").Specific
            'ComboLoad_VolumeUnit(oForm, oCombo)
            '-------------
            oCombo = oForm.Items.Item("1000031").Specific
            ComboLoad_FreightUnit(oForm, oCombo)
            oCombo = oForm.Items.Item("184").Specific
            ComboLoad_FreightUnit(oForm, oCombo)
            '--------
            'oCombo = oForm.Items.Item("1000003").Specific
            'oCombo.ValidValues.Add("Cheque", "C")
            'oCombo.ValidValues.Add("Cash", "Ch")
            'oCombo.ValidValues.Add("CC", "CC")
            'oCombo.ValidValues.Add("Online", "O")


            ''-----DO
            'CFL_BP_Supplier2(oForm, SBO_Application)
            'oMatrix = oForm.Items.Item("SIJDOMAT").Specific
            'oColumns = oMatrix.Columns
            'oMatrix.AddRow()
            'oColumn = oColumns.Item("V_13")
            'oColumn.ChooseFromListUID = "CFLBPV1"
            'oColumn.ChooseFromListAlias = "CardCode"
            'CFL_Item(oForm, SBO_Application)
            'oColumn = oColumns.Item("V_15")
            'oColumn.ChooseFromListUID = "OITM"
            'oColumn.ChooseFromListAlias = "ItemCode"


            ''------------VO
            'CFL_Item1(oForm, SBO_Application)
            'oMatrix1 = oForm.Items.Item("148").Specific
            'oColumns = oMatrix1.Columns
            'oMatrix1.AddRow()
            'oColumn = oColumns.Item("V_8")
            'oColumn.ChooseFromListUID = "1OITM"
            'oColumn.ChooseFromListAlias = "ItemCode"

            '------------Charge
            CFL_Item2(oForm, SBO_Application)
            oMatrix3 = oForm.Items.Item("ChargeMat").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.AddRow()
            oColumn = oColumns.Item("V_10")
            oColumn.ChooseFromListUID = "21OITM"
            oColumn.ChooseFromListAlias = "ItemCode"
            ''------------Chargo
            'CFL_BP_Supplier3(oForm, SBO_Application)
            'oMatrix4 = oForm.Items.Item("CargoMat").Specific
            'oColumns = oMatrix4.Columns
            'oMatrix4.AddRow()
            'oColumn = oColumns.Item("V_11")
            'oColumn.ChooseFromListUID = "3CFLBPV1"
            'oColumn.ChooseFromListAlias = "CardCode"


            ''---- goods Receipt

            'oForm.DataSources.UserDataSources.Add("V_0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oMatrix3 = oForm.Items.Item("AIJGR").Specific
            'oColumns = oMatrix3.Columns
            'oColumn = oColumns.Item("V_0")
            'oColumn.DataBind.SetBound(True, "", "V_0")
            'oItem = oForm.Items.Item("AIJGR")
            'oItem.Width = 150
            'oItem.Height = 15
            'oColumn.Width = 130
            'oMatrix3.AddRow()

            'CFL
            CFL_Item_Vessel(oForm, SBO_Application)
            oEdit = oForm.Items.Item("AIJ1000042").Specific
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemName"

            CFL_BP_Customer(oForm, SBO_Application)
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.ChooseFromListUID = "CFLBPC"
            oEdit.ChooseFromListAlias = "CardCode"

            'CFL_BP_Supplier(oForm, SBO_Application)
            'oEdit = oForm.Items.Item("127").Specific
            'oEdit.ChooseFromListUID = "CFLBPV"
            'oEdit.ChooseFromListAlias = "CardCode"
            ' oForm.PaneLevel = 1
            CFL_SalesOrder(oForm, SBO_Application, "AI")
            oEdit = oForm.Items.Item("AIJ4").Specific
            oEdit.ChooseFromListUID = "ORDR"
            oEdit.ChooseFromListAlias = "DocNum"
            '  ComboLoad_PaymentType(oForm, oCombo)
            oForm.DataBrowser.BrowseBy = "1000004"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Shared Sub LoadGrid_REF_ATTACH(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("REFATT").Specific
            oEdit = oForm.Items.Item("123").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            If JobNo.Contains("IN") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_INT_HEADER]  T0 , [dbo].[@AB_INT_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            ElseIf JobNo.Contains("PR") = True Then
                str = "SELECT Cast(T1.[U_Path]  as varchar(1500))+ T1.[U_File]  'File Name' FROM [dbo].[@AB_PRO_HEADER]  T0 , [dbo].[@AB_PRO_ATT]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and  T0.[U_JobNo] ='" & JobNo & "' and  T1.[U_St] ='Closed'  and Cast(T1.[U_Path]  as varchar(1500)) <>''"
            End If
            oForm.DataSources.DataTables.Item("REFAI").ExecuteQuery(Str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("REFAI")
            oGrid.Columns.Item("RowsHeader").Width = 30
            '   oGrid.AutoResizeColumns()
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_REF_ATTACH" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_PV(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("PVGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum], T0.[U_VOVCode] as 'Vendor Code', T0.[U_VOVName] as 'Vendor Name', T0.[U_VOType] as 'Payment Type', T0.[U_JobNo] as 'Job No', T0.[U_VONo] as 'Voucher No', T0.[U_VODt] as 'Voucher Date', T0.[U_VOTotAmt] as 'Amount' FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 WHERE T0.[U_JobNo]  ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("PVAI").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("PVAI")
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:LoadGrid_PV" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_BI(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("BIGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = "'" & oEdit.String & "'"
            oEdit = oForm.Items.Item("123").Specific
            If oEdit.String <> "" And oEdit.String <> "NA" Then
                JobNo = JobNo & "," & "'" & oEdit.String & "'"
            End If
            Dim str As String = "SELECT DocEntry 'DocNum','DraftInvoice' DocumentType ,T0.[DocDate] 'BIDate', T0.[CardCode] 'Customer Code', T0.[CardName] 'Customer Name', T0.[U_AB_JobNo] 'Job No', T0.[DocTotal] 'Document Total' FROM ODRF T0 WHERE T0.[ObjType] =13 and  T0.[DocStatus] ='O' and  T0.[U_AB_JobNo] in ( " & JobNo & ") union all SELECT DocEntry 'DocNum','Invoice' DocumentType , T0.[DocDate] 'BIDate', T0.[CardCode] 'Customer Code', T0.[CardName] 'Customer Name', T0.[U_AB_JobNo] 'Job No', T0.[DocTotal] 'Document Total' FROM OINV T0 WHERE   T0.[U_AB_JobNo] in (" & JobNo & ")"
            oForm.DataSources.DataTables.Item("BIAI").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("BIAI")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Public Shared Sub LoadGrid_DO(ByVal oForm As SAPbouiCOM.Form)
        Try

            oGrid = oForm.Items.Item("DOGrid").Specific
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            Dim str As String = "SELECT T0.[DocNum],T0.[U_JobNo] 'JobNo',  T0.[U_CardCode] 'Card Code', T0.[U_CardName] 'Card Name', T0.[U_VesselNo] 'Vessel', T0.[U_MAWBNo] 'OBL No', T0.[U_TaxDate] as 'Date', T0.[U_ANSRecNo] as 'DO No' FROM [dbo].[@AIGI]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'"
            oForm.DataSources.DataTables.Item("DOAI").ExecuteQuery(str)
            oGrid.DataTable = oForm.DataSources.DataTables.Item("DOAI")
        Catch ex As Exception
            Functions.WriteLog("Class:F_SE_JobOrder" + " Function:LoadGrid_DO" + " Error Message:" + ex.ToString)
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
    Public Sub DocNumber_AI()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"
            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+3) as CountNo FROM [dbo].[@AB_AIRI_JOB_H]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("SIJ16").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "AI" & Format(Now.Date, "yy") & "J" & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "AI" & Format(Now.Date, "yy") & "J" & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "AI" & Format(Now.Date, "yy") & "J" & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "AI" & Format(Now.Date, "yy") & "J" & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "AI" & Format(Now.Date, "yy") & "J" & "0" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = "AI" & Format(Now.Date, "yy") & "J" & DocNum
            End If
            ''--------DO NO
            'oEdit = oForm.Items.Item("115").Specific
            'DocNumLen = DocNum.ToString.Length
            'If DocNum = 0 Then
            '    oEdit.String = "GKWAI" & Format(Now.Date, "yyyyMMdd") & "0001"
            'ElseIf DocNumLen = 1 And DocNum <> 0 Then
            '    oEdit.String = "GKWAI" & Format(Now.Date, "yyyyMMdd") & "000" & DocNum
            'ElseIf DocNumLen = 2 Then
            '    oEdit.String = "GKWAI" & Format(Now.Date, "yyyyMMdd") & "00" & DocNum
            'ElseIf DocNumLen = 3 Then
            '    oEdit.String = "GKWAI" & Format(Now.Date, "yyyyMMdd") & "0" & DocNum
            'ElseIf DocNumLen = 4 Then
            '    oEdit.String = "GKWAI" & Format(Now.Date, "yyyyMMdd") & "" & DocNum
            'ElseIf DocNumLen = 5 Then
            '    oEdit.String = "GKWAI" & Format(Now.Date, "yyyyMMdd") & DocNum
            'End If


        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Public Sub ShippigNameLoad()
        'Try
        '    oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    oRecordSet1.DoQuery("SELECT UPPER(T0.[CompnyName]), UPPER(T0.[CompnyAddr]) FROM OADM T0")
        '    oEdit = oForm.Items.Item("1000027").Specific
        '    oEdit.String = oRecordSet1.Fields.Item(0).Value.ToString
        '    oEdit = oForm.Items.Item("96").Specific
        '    oEdit.String = oRecordSet1.Fields.Item(1).Value.ToString
        'Catch ex As Exception
        '    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'End Try

    End Sub
    Public Sub ComboLoad_Carrier(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = oCOmpany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    'AB_FERIGHT
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormUID = "AIRI_JOB" Then
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
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            '-----------Inovice Draft------------
            'If pVal.FormType = 11133 Then
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
            If ((pVal.FormUID = "AIRI_JOB" And pVal.ItemUID = "REFATTFOL" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                LoadGrid_REF_ATTACH(oForm)
            ElseIf ((pVal.FormUID = "AIRI_JOB" And pVal.ItemUID = "REFDIS" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "AIRI_JOB" And pVal.ItemUID = "ATTMAT" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "AIRI_JOB" And pVal.ItemUID = "1000006" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then

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
            If ((pVal.FormUID = "AIRI_JOB" And pVal.ItemUID = "155" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
            If ((pVal.FormUID = "AIRI_JOB" And pVal.ItemUID = "1000005" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                Try
                    oForm = SBO_Application.Forms.Item("AIRI_JOB")

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
                oForm = SBO_Application.Forms.Item("AIRI_JOB")
                If oForm.Visible = True Then
                    LoadGrid_DO(oForm)
                End If
            End If
            If pVal.FormUID = "AB_PV" And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("AIRI_JOB")
                If oForm.Visible = True Then
                    LoadGrid_PV(oForm)
                End If
            End If
            If pVal.FormType = 133 And pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.Item("AIRI_JOB")
                If oForm.Visible = True Then
                    LoadGrid_BI(oForm)
                End If
            End If
            '----------load marix
            '------laod matrix
            If pVal.FormType = 2000107 Then
                'If (pVal.ItemUID = "1" And pVal.Before_Action = False And pVal.InnerEvent = False And SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Or (pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                If (pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) Then
                    oForm = SBO_Application.Forms.Item("AIRI_JOB")
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

            If pVal.FormUID = "AIRI_JOB" Then
                oForm = SBO_Application.Forms.Item("AIRI_JOB")
                If pVal.ItemUID = "AICRR" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("AICRR").Specific
                    Dim ContCode As String = oEdit.String
                    If ContCode <> "" Then
                        oEdit = oForm.Items.Item("61").Specific
                        oEdit.String = Carrier_Name(ContCode, Ocompany)
                    End If
                End If
                '---------------Item Event-----------------------
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "89" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 1
                        ElseIf pVal.ItemUID = "REFATTFOL" Then
                            oForm.PaneLevel = 10
                        ElseIf pVal.ItemUID = "AWBinfo" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 2
                        ElseIf pVal.ItemUID = "Charge" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 6
                            'ElseIf pVal.ItemUID = "Cargo" Then
                            '    oForm.PaneLevel = 7
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
                        ElseIf pVal.ItemUID = "128" Then
                            oEdit = oForm.Items.Item("SIJ12").Specific
                            oEdit.String = oEdit.String
                            oForm.PaneLevel = 7
                        ElseIf pVal.ItemUID = "DOButt" Then
                            LoadDeliveryOrder(oForm)
                        ElseIf pVal.ItemUID = "PVButton" Then
                            LoadPaymentVoucher(oForm)
                        ElseIf pVal.ItemUID = "185" Then

                            'Price Load
                            LoadHandingCharge_AirImport(oForm, "AI")
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
                        ElseIf pVal.ItemUID = "PrintBI" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf BI_Report)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()
                            End If
                        ElseIf pVal.ItemUID = "PrintBI" Then
                            BI_Report()

                        ElseIf pVal.ItemUID = "152" Then
                            'print Booking Sheet
                        ElseIf pVal.ItemUID = "SIJPV" Then
                            'Shipping Order
                        ElseIf pVal.ItemUID = "149" Then
                            'Delivery Order
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf Delivery_Order)
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
                            DocNumber_AI()
                            oEdit = oForm.Items.Item("SIJ18").Specific
                            oEdit.String = Format(Now.Date, "dd/MM/yy")
                            ShippigNameLoad()
                        End If
                    ElseIf pVal.Before_Action = True Then
                        If pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If pVal.ItemUID = "1" Then
                                DocNumber_AI()
                            End If
                        End If
                        '----------------------------
                        If pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
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
10:                                     If System.IO.File.Exists(destPath & FileName & FileExten) Then
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
                                        oCombo = oForm.Items.Item("125").Specific
                                        If oCombo.Selected.Value = "Done" Then
                                            oEdit = oForm.Items.Item("123").Specific
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
                        '---------------------

                        ElseIf pVal.Before_Action = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        'If pVal.ItemUID = "1" Then
                        '    oCombo = oForm.Items.Item("125").Specific
                        '    If oCombo.Selected.Value = "Done" Then
                        '        oEdit = oForm.Items.Item("123").Specific
                        '        Dim BaseJobNo As String = oEdit.String
                        '        oEdit = oForm.Items.Item("SIJ16").Specific
                        '        Dim JobNo As String = oEdit.String
                        '        If BaseJobNo.Substring(0, 2) = "IN" Or BaseJobNo.Substring(0, 2) = "PR" Then
                        '            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '            oRecordSet1.DoQuery("UPDATE ODRF SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                        '            oRecordSet1.DoQuery("UPDATE OINV SET U_AB_JOBNO='" & BaseJobNo & "',U_AB_Divsion='" & BaseJobNo.Substring(0, 2) & "' where U_AB_JobNo='" & JobNo & "'")
                        '            LoadGrid_BI(oForm)
                        '        End If
                        '    End If
                        'End If
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
                            End If
                        End If

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
                                            SBO_Application.StatusBar.SetText("Invoice Can't Be Open", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

                                    'Gopinath
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

                            If pVal.ItemUID = "AIJ4" Then
                                oEdit = oForm.Items.Item("AIJ4").Specific
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
                            If pVal.ItemUID = "AIJ157" Then
                                oEdit = oForm.Items.Item("AIJ157").Specific
                                Dim ContCode As String = oEdit.String
                                If ContCode <> "" Then
                                    oEdit = oForm.Items.Item("156").Specific
                                    oEdit.String = City_Code(ContCode, Ocompany)
                                End If
                            End If
                            If pVal.ItemUID = "AIJ1000011" Then
                                oEdit = oForm.Items.Item("AIJ1000011").Specific
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
                            oForm = SBO_Application.Forms.Item("AIRI_JOB")
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
                                'AIJ1000042
                                If pVal.ItemUID = "AIJ1000042" Then
                                    Try
                                        oEdit = oForm.Items.Item("AIJ1000042").Specific
                                        oEdit.String = oDataTable.GetValue("ItemName", 0)
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
                                If pVal.ItemUID = "AIJ4" Then
                                    Try
                                        oEdit = oForm.Items.Item("AIJ4").Specific
                                        oEdit.String = oDataTable.GetValue("DocNum", 0)
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
                                'nath
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
                            oItem = oForm.Items.Item("TransBI")
                            oItem.Enabled = False
                        ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
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
                            oItem = oForm.Items.Item("TransBI")
                            oItem.Enabled = False
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
                If oForm.UniqueID = "AIRI_JOB" Then
                    'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = True
                    'End If

                   

                End If
            End If
            If pVal.MenuUID = "1282" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "AIRI_JOB" Then
                    oItem = oForm.Items.Item("153")
                    oItem.Enabled = False
                    DocNumber_AI()
                    oEdit = oForm.Items.Item("SIJ18").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    ShippigNameLoad()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub CreatePO()
        Try
            Dim oform1 As SAPbouiCOM.Form
            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
            oMatrix2 = oform1.Items.Item("38").Specific
            oColumns = oMatrix2.Columns
            Dim i As Integer
            Dim k As Integer
            Dim k1 As Integer
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
                    k1 = oPO.DocNum

               

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
            'Exit Sub
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim MAWBNo As String = ""
            Dim HAWBNo As String = ""
            oRecordSet1.DoQuery("SELECT T0.[U_AWBNo1] +T0.[U_AWBNo] FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
            MAWBNo = oRecordSet1.Fields.Item(0).Value.ToString
            oRecordSet1.DoQuery("SELECT T0.[U_HAWBNo] FROM [dbo].[@AB_AWB_H]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
            HAWBNo = oRecordSet1.Fields.Item(0).Value.ToString

            oRecordSet1.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],U_VessName [U_VessVoyage],U_AWBNo [U_OBL],U_HAWBNo [U_HBL],U_FNo [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.U_ETA FROM [dbo].[@AB_AIRI_JOB_H] T0 WHERE T0.[U_JobNo]   ='" & JobNo & "'")
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
            Catch ex As Exception
                SBO_Application.ActivateMenuItem("6913")
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-133, 1)
            End Try
            oCombo = oform1.Items.Item("U_AB_Divsion").Specific
            oCombo.Select("AI", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oEdit = oform1.Items.Item("U_AB_JobNo").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_JobNo").Value
            oEdit = oform1.Items.Item("U_AB_OriginNameC").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_LoadPortNC").Value
            oEdit = oform1.Items.Item("U_AB_DestNameC").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_DisPortNC").Value
            ' T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_TotPkg], T0.[U_ChrgWt], T0.[U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_F1] FROM [dbo].[@AB_INT_HEADER]  T0 WHERE T0.[U_JobNo]   ='" & JobNo & "'")
            oEdit = oform1.Items.Item("U_AB_TotPkg").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_TotPkg").Value
            oEdit = oform1.Items.Item("U_AB_TotWT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_ChrgWt").Value
            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessVoyage").Value
            oEdit = oform1.Items.Item("U_AB_MAWB").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_OBL").Value
            oEdit = oform1.Items.Item("U_AB_HAWB").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_HBL").Value
            oEdit = oform1.Items.Item("U_AB_FLT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_F1").Value
            oEdit = oform1.Items.Item("U_AB_WT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_ChrgWt").Value
            oEdit = oform1.Items.Item("U_AB_SSIT").Specific
            oEdit.String = oRecordSet1.Fields.Item("U_VessName").Value
            'U_VessName
            Try
                oEdit = oform1.Items.Item("U_AB_ETDETA").Specific
                oEdit.String = Format(oRecordSet1.Fields.Item("U_ETA").Value, "dd/MM/yy")
            Catch ex As Exception

            End Try
            
            Dim QTNo As String = oRecordSet1.Fields.Item("U_QNo").Value
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[Dscription], T1.[Quantity], T1.[Price], T0.[DocCur],T1.[U_AB_Vendor],T1.U_AB_Cost,T1.[unitMsr],T1.FreeTxt FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.[DocNum] ='" & QTNo & "'")
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
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & oRecordSet.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
            oform1.Title = "Billing Instruction"
            'SELECT T0.[U_CCode], T0.[U_CName], T0.[U_LoadPortN], T0.[U_DisPortN], T0.[U_JobNo], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ETD] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] =''")
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub LoadJobOrder(ByVal SQNO As String)
        Try
            SBO_Application.StatusBar.SetText("Please Waite Data is Loading", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'SELECT T0.[CardCode], T0.[CardName], T0.[U_AB_Divison], T0.[U_AB_TransTo], T0.[U_AB_Trnst], T0.[U_AB_VessCode], T0.[U_AB_VessName], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_Desc], T0.[U_AB_Validity], T0.[U_AB_Ttime], T0.[U_AB_Freq], T0.[U_AB_CARTotQt], T0.[U_AB_CARTotWt] FROM ORDR T0 WHERE T0.[DocNum] ='1'
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.DocNum,T0.[CardCode], T0.[CardName], T1.[Name], T0.[U_AB_CaType], T0.[U_SerLevel], T0.[U_AB_CarricerC], T0.[U_AB_CarrierN], T0.[U_AB_OrginCode], T0.[U_AB_OriginName], T0.[U_AB_OrginCodeC], T0.[U_AB_OriginNameC], T0.[U_AB_DestCode], T0.[U_AB_DestName], T0.[U_AB_DestCodeC], T0.[U_AB_DestNameC],T0.[U_AB_Desc],T0.[U_AB_Divsion1], T0.[U_AB_JobNo],T0.[U_AB_TotPkg], T0.[U_AB_TotWT],T0.[U_AB_VessName]  FROM ORDR T0 Left JOIN OCPR T1 ON T0.CntctCode = T1.CntctCode WHERE  T0.[U_AB_Divsion] ='AI' and isnull( T0.[U_AB_Status] ,'')='Open' and T0.DocNum='" & SQNO & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oEdit = oForm.Items.Item("AIJ4").Specific
            oEdit.String = oRecordSet1.Fields.Item(0).Value
            oEdit = oForm.Items.Item("SIJ6").Specific
            oEdit.String = oRecordSet1.Fields.Item(1).Value
            oEdit = oForm.Items.Item("SIJ8").Specific
            oEdit.String = oRecordSet1.Fields.Item(2).Value

            Try
                oCombo = oForm.Items.Item("30").Specific
                oCombo.Select(oRecordSet1.Fields.Item(4).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("1000008").Specific
                oCombo.Select(oRecordSet1.Fields.Item(5).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
            End Try
            Try
                oEdit = oForm.Items.Item("1000011").Specific
                oEdit.String = oRecordSet1.Fields.Item(6).Value
            Catch ex As Exception
            End Try
            oEdit = oForm.Items.Item("61").Specific
            oEdit.String = oRecordSet1.Fields.Item(7).Value

            oEdit = oForm.Items.Item("1000027").Specific
            oEdit.String = oRecordSet1.Fields.Item(2).Value
            oEdit = oForm.Items.Item("96").Specific
            oEdit.String = BPAddress(oRecordSet1.Fields.Item(1).Value, Ocompany)
            oEdit = oForm.Items.Item("SJI10").Specific
            oEdit.String = oRecordSet1.Fields.Item(3).Value
            oCombo = oForm.Items.Item("125").Specific
            If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "IN" Or oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "PR" Then
                oCombo.Select("Approved", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Else
                oCombo.Select("NA", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            oItem = oForm.Items.Item("1000002")
            oItem.Enabled = False
            oItem = oForm.Items.Item("123")
            oItem.Enabled = False


            Try

                oItem = oForm.Items.Item("1000002")
                oItem.Enabled = True
                oItem = oForm.Items.Item("123")
                oItem.Enabled = True
                oEdit = oForm.Items.Item("1000002").Specific
                If oRecordSet1.Fields.Item("U_AB_Divsion1").Value = "" Then
                    oEdit.String = "NA"
                    oEdit = oForm.Items.Item("AIJ157").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_OrginCodeC").Value
                    oEdit = oForm.Items.Item("156").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_OriginNameC").Value
                    oEdit = oForm.Items.Item("AIJ1000011").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_DestCodeC").Value
                    oEdit = oForm.Items.Item("40").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_DestNameC").Value
                    oEdit = oForm.Items.Item("160").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_TotPkg").Value
                    oEdit = oForm.Items.Item("162").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_TotWT").Value
                
                    oEdit = oForm.Items.Item("168").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_TotWT").Value
                    oEdit = oForm.Items.Item("AIJ1000042").Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_VessName").Value
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_Divsion1").Value
                End If
                oEdit = oForm.Items.Item("123").Specific
                If oRecordSet1.Fields.Item("U_AB_JobNo").Value = "" Then
                    oEdit.String = "NA"
                Else
                    oEdit.String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                    Dim JobNo As String = oRecordSet1.Fields.Item("U_AB_JobNo").Value
                    If JobNo <> "" Then
                        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If JobNo.Substring(0, 2) = "IN" Then
                            oRecordSet.DoQuery("SELECT T0.[U_VessCode], T0.[U_VessName], T0.[U_ItemDesc], T0.[U_LoadPortCC],T0.[U_LoadPortNC], T0.[U_DisPortCC],T0.[U_DisPortNC], T0.[U_OBL], T0.[U_HBL], T0.[U_F1], T0.[U_ETA], T0.[U_ETD], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ChrgWt] FROM [dbo].[@AB_INT_HEADER]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
                        ElseIf JobNo.Substring(0, 2) = "PR" Then
                            oRecordSet.DoQuery("SELECT T0.[U_VessCode], T0.[U_VessName], T0.[U_ItemDesc], T0.[U_LoadPortCC],T0.[U_LoadPortNC], T0.[U_DisPortCC],T0.[U_DisPortNC], T0.[U_OBL], T0.[U_HBL], T0.[U_F1], T0.[U_ETA], T0.[U_ETD], T0.[U_TotPkg], T0.[U_GrssWt], T0.[U_ChrgWt] FROM [dbo].[@AB_PRO_HEADER]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
                        Else
                            Exit Try
                        End If
                        oEdit = oForm.Items.Item("AIJ157").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_LoadPortCC").Value
                        oEdit = oForm.Items.Item("156").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_LoadPortNC").Value
                        oEdit = oForm.Items.Item("AIJ1000011").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_DisPortCC").Value
                        oEdit = oForm.Items.Item("40").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_DisPortNC").Value
                        oEdit = oForm.Items.Item("1000032").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_F1").Value
                        Try
                            oEdit = oForm.Items.Item("1000034").Specific
                            oEdit.String = Format(oRecordSet.Fields.Item("U_ETD").Value, "dd/MM/yy")
                            oEdit = oForm.Items.Item("51").Specific
                            oEdit.String = Format(oRecordSet.Fields.Item("U_ETA").Value, "dd/MM/yy")
                        Catch ex As Exception

                        End Try
                        oEdit = oForm.Items.Item("1000036").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_OBL").Value
                        oEdit = oForm.Items.Item("118").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_HBL").Value

                        oEdit = oForm.Items.Item("160").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_TotPkg").Value
                        oEdit = oForm.Items.Item("162").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_GrssWt").Value
                        oEdit = oForm.Items.Item("165").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_ChrgWt").Value
                        oEdit = oForm.Items.Item("168").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_GrssWt").Value
                        oEdit = oForm.Items.Item("AIJ1000042").Specific
                        oEdit.String = oRecordSet.Fields.Item("U_VessName").Value

                    End If
                End If

            Catch ex As Exception
            End Try

            oForm.PaneLevel = 1
            SBO_Application.StatusBar.SetText("Data Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub MatrixLoad(ByVal DocNum As Integer, ByVal VenName As String, ByVal PONo As String)
        oForm = SBO_Application.Forms.Item("AIRI_JOB")
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
    Private Sub LoadHandingCharge_AirImport(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            'Dim DestCountry As String = ""
            'Dim DestCity As String = ""
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            'oEdit = oForm.Items.Item("ce13").Specific
            'DestCity = oEdit.String

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
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix4.AddRow()

                'oEdit = oMatrix4.Columns.Item("U_Ven").Cells.Item(1).Specific
                'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                oEdit = oMatrix4.Columns.Item("V_10").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix4.Columns.Item("V_9").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                oEdit = oMatrix4.Columns.Item("V_8").Cells.Item(i).Specific
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
                oEdit = oMatrix4.Columns.Item("V_14").Cells.Item(i).Specific
                oEdit.String = UnitPrice 'oRecordSet1.Fields.Item("U_PerKg").Value

                oEdit = oMatrix4.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix4.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
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
    'Private Sub Payment_Voucher()
    '    Try

    '        '1000004
    '        Dim DocNum As String = ""
    '        oGrid = oForm.Items.Item("PVGrid").Specific
    '        For F = 0 To oGrid.Rows.Count - 1
    '            If oGrid.Rows.IsSelected(F) = True Then
    '                DocNum = oGrid.DataTable.GetValue("DocNum", F)
    '                Exit For
    '            End If
    '        Next
    '        If DocNum = "" Then
    '            SBO_Application.StatusBar.SetText("Select Payment Voucher", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If
    '        'SBO_Application.StatusBar.SetText("report is generating")
    '        Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '        Dim ERRPT As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '        Dim objConInfo As New CrystalDecisions.Shared.ConnectionInfo
    '        Dim oLogonInfo As New CrystalDecisions.Shared.TableLogOnInfo
    '        Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
    '        Dim intCounter As Integer
    '        '  Dim Formula As String
    '        Dim OneMore As Boolean = False
    '        Dim sPath As String
    '        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
    '        'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
    '        cryRpt.Load(sPath & "\GK_FM\AB_RP005_PV_SE.rpt")
    '        ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
    '        Dim ParaName As String = "@DocKey"
    '        Dim ParaValue As String = DocNum
    '        Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
    '        Dim Para As New CrystalDecisions.Shared.ParameterDiscreteValue
    '        Dim MyArr1 As Array = ParaName.Split(";")
    '        Dim MyArr2 As Array = ParaValue.Split(";")
    '        For i As Integer = 0 To MyArr1.Length - 1
    '            Para.Value = MyArr2(i)
    '            pvCollection.Add(Para)
    '            cryRpt.DataDefinition.ParameterFields(MyArr1(i)).ApplyCurrentValues(pvCollection)
    '        Next
    '        Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\GK_FM\" & "Pwd.txt", True)
    '        Dim pwd As String = ""
    '        pwd = file.ReadLine()
    '        ConInfo.ConnectionInfo.UserID = "sa"
    '        ConInfo.ConnectionInfo.Password = pwd
    '        ConInfo.ConnectionInfo.ServerName = Ocompany.Server
    '        ConInfo.ConnectionInfo.DatabaseName = Ocompany.CompanyDB
    '        For intCounter = 0 To cryRpt.Database.Tables.Count - 1
    '            cryRpt.Database.Tables(intCounter).ApplyLogOnInfo(ConInfo)
    '        Next
    '        Dim RptFrm As MY_Report
    '        RptFrm = New MY_Report
    '        RptFrm.CrystalReportViewer1.ReportSource = cryRpt
    '        RptFrm.Refresh()
    '        RptFrm.Text = "Payment Voucher Report"
    '        RptFrm.TopMost = True
    '        RptFrm.Activate()
    '        RptFrm.ShowDialog()
    '        RptFrm.Refresh()
    '    Catch ex As Exception
    '        SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
    Private Sub Transfer_BI()
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
    Private Sub LoadDeliveryOrder(ByVal oform As SAPbouiCOM.Form)
        Try
            Dim oform1 As SAPbouiCOM.Form
            LoadFromXML("GoodsIssue.srf", SBO_Application)
            oform1 = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
            oEdit = oform.Items.Item("SIJ16").Specific
            Dim JobNo As String = oEdit.String
            oEdit = oform1.Items.Item("35").Specific
            oEdit.String = JobNo
            oEdit = oform.Items.Item("AIJ1000042").Specific
            Dim VesslNo As String = oEdit.String
            oEdit = oform1.Items.Item("20").Specific
            oEdit.String = VesslNo
            'CFL_Item_Vessel(oform, SBO_Application)

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
            If MAWBNo = "" Then
                Exit Sub
            End If
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP002_BS_AI.rpt")
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP004_DO_AI.rpt")
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP005_PV_AI.rpt")
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
            cryRpt.Load(sPath & "\GK_FM\AB_RP006_TS_AI.rpt")
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
