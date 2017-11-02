Imports System.Diagnostics.Process
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports
Imports System.IO
Public Class F_Invoice
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub

    Public Sub Form_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.UserDataSources.Add("oedit1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("oedit2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("oedit3", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("oedit4", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("oedit5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("oedit6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("oedit7", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("oedit8", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("oedit9", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("V_0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_1", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("V_2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("V_8", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("V_9", SAPbouiCOM.BoDataType.dt_PRICE)


            CFL_Item_Vessel(oForm, SBO_Application)
            ' oEdit = oForm.Items.Item("97").Specific
            oEdit = oForm.Items.Item("vn4").Specific
            oEdit.DataBind.SetBound(True, "", "oedit1")
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemName"
            CFL_BP_Customer(oForm, SBO_Application)
            oEdit = oForm.Items.Item("cc16").Specific
            oEdit.DataBind.SetBound(True, "", "oedit2")
            oEdit.ChooseFromListUID = "CFLBPC"
            oEdit.ChooseFromListAlias = "CardCode"

            oEdit = oForm.Items.Item("dn6").Specific
            oEdit.DataBind.SetBound(True, "", "oedit3")
            oEdit = oForm.Items.Item("dd18").Specific
            oEdit.DataBind.SetBound(True, "", "oedit4")
            oEdit = oForm.Items.Item("po24").Specific
            oEdit.DataBind.SetBound(True, "", "oedit9")
            oEdit = oForm.Items.Item("invser").Specific
            oEdit.DataBind.SetBound(True, "", "oedit5")
            oEdit = oForm.Items.Item("ic10").Specific
            oEdit.DataBind.SetBound(True, "", "oedit6")
            oEdit = oForm.Items.Item("cd20").Specific
            oEdit.DataBind.SetBound(True, "", "oedit7")
            oEdit = oForm.Items.Item("cdd").Specific
            oEdit.DataBind.SetBound(True, "", "oedit8")

            oMatrix = oForm.Items.Item("22").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("V_0")
            oColumn.DataBind.SetBound(True, "", "V_0")
            oColumn = oColumns.Item("V_1")
            oColumn.DataBind.SetBound(True, "", "V_1")
            oColumn = oColumns.Item("V_2")
            oColumn.DataBind.SetBound(True, "", "V_2")
            oColumn = oColumns.Item("V_3")
            oColumn.DataBind.SetBound(True, "", "V_3")
            oColumn = oColumns.Item("V_4")
            oColumn.DataBind.SetBound(True, "", "V_4")
            oColumn = oColumns.Item("V_5")
            oColumn.DataBind.SetBound(True, "", "V_5")
            oColumn = oColumns.Item("V_6")
            oColumn.DataBind.SetBound(True, "", "V_6")
            oColumn = oColumns.Item("V_8")
            oColumn.DataBind.SetBound(True, "", "V_8")
            oColumn = oColumns.Item("V_9")
            oColumn.DataBind.SetBound(True, "", "V_9")
            oForm.DataSources.DataTables.Add("OWHS")
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try


    End Sub
    Private Sub MatrixLoad(ByVal VesselName As String, ByVal CustName As String, ByVal PONo As String, ByVal Fdate As String, ByVal Tdate As String)

        If PONo = "" Then
            PONo = "%"
        End If
        Dim sqlstr As String = "SELECT T0.[DocCur], Case when T0.[DocCur] ='SGD' then T0.[DocTotal] else T0.[DocTotalFC] End as [DocTotal]  , 'Y' as 'Y', T0.[NumAtCard],DocEntry, T0.[U_AB_SSIT], T0.[CardName], T0.[DocDate], T0.[U_AB_JobNo], T0.[Comments] FROM OINV T0 WHERE T0.[U_AB_SSIT]='" & VesselName & "' and  T0.[CardCode] ='" & CustName & "' and  isnull(T0.[NumAtCard],'') like '" & PONo & "' and  T0.[DocStatus] ='O' and isnull(T0.U_C_INVNo,'') = '' and T0.DocDate between '" & Fdate & "' and '" & Tdate & "'"
        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sqlstr)
        If oRecordSet.RecordCount = 0 Then
            SBO_Application.StatusBar.SetText("No Records found!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm.DataSources.DataTables.Item("OWHS").ExecuteQuery(sqlstr)
        oMatrix = oForm.Items.Item("22").Specific
        oMatrix.Clear()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        oColumns = oMatrix.Columns
        oForm.Items.Item("22").Specific.Columns.item("V_5").DataBind.Bind("OWHS", "Y")
        oForm.Items.Item("22").Specific.Columns.item("V_4").DataBind.Bind("OWHS", "DocEntry")
        oForm.Items.Item("22").Specific.Columns.item("V_3").DataBind.Bind("OWHS", "U_AB_SSIT")
        oForm.Items.Item("22").Specific.Columns.item("V_2").DataBind.Bind("OWHS", "CardName")
        oForm.Items.Item("22").Specific.Columns.item("V_1").DataBind.Bind("OWHS", "DocDate")
        oForm.Items.Item("22").Specific.Columns.item("V_6").DataBind.Bind("OWHS", "U_AB_JobNo")
        oForm.Items.Item("22").Specific.Columns.item("V_8").DataBind.Bind("OWHS", "Comments")
        oForm.Items.Item("22").Specific.Columns.item("V_7").DataBind.Bind("OWHS", "NumAtCard")

        oForm.Items.Item("22").Specific.Columns.item("V_0").DataBind.Bind("OWHS", "DocCur")
        oForm.Items.Item("22").Specific.Columns.item("V_9").DataBind.Bind("OWHS", "DocTotal")
        oForm.Items.Item("22").Specific.Clear()
        oForm.Items.Item("22").Specific.LoadFromDataSource()
        oForm.Items.Item("22").Specific.AutoResizeColumns()

    End Sub
    Private Sub SBO_Application_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "InvoiceNew" Then
                Try

                    oForm = SBO_Application.Forms.Item("InvoiceNew")

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                        Try
                            If pVal.ItemUID = "SN" Then
                                Dim CustName, VesselName, PONo, Fdate, Tdate As String
                             
                                oEdit = oForm.Items.Item("vn4").Specific
                                VesselName = oEdit.String
                                If VesselName = "" Then
                                    SBO_Application.StatusBar.SetText("Vessel Name Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                                oEdit = oForm.Items.Item("cc16").Specific
                                CustName = oEdit.String
                                If CustName = "" Then
                                    SBO_Application.StatusBar.SetText("Customer Code Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                                oEdit = oForm.Items.Item("po24").Specific
                                PONo = oEdit.String
                                'If PONo = "" Then
                                '    SBO_Application.StatusBar.SetText("PO/Ref No. Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '    Exit Sub
                                'End If
                                oEdit = oForm.Items.Item("dn6").Specific
                                Fdate = oEdit.Value
                                If Fdate = "" Then
                                    SBO_Application.StatusBar.SetText("From Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                                oEdit = oForm.Items.Item("dd18").Specific
                                Tdate = oEdit.Value
                                If Tdate = "" Then
                                    SBO_Application.StatusBar.SetText("To Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                                MatrixLoad(VesselName, CustName, PONo, Fdate, Tdate)
                            ElseIf pVal.ItemUID = "Gen" Then
                              

                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf ConslInvReport1)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()

                                

                                ' ConslInvoiceReport(DocNum)
                            ElseIf pVal.ItemUID = "FN" Then
                               
                                Dim trd As Threading.Thread
                                trd = New Threading.Thread(AddressOf ConslInvReport)
                                trd.IsBackground = True
                                trd.SetApartmentState(ApartmentState.STA)
                                trd.Start()

                                ' ConslInvoiceReport(DocNum)
                            End If

                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                    ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "cd20" Then
                            Try
                                Dim CustName As String
                                oEdit = oForm.Items.Item("cc16").Specific
                                CustName = oEdit.String
                                If CustName = "" Then
                                    SBO_Application.StatusBar.SetText("Customer Code Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If


                                oEdit = oForm.Items.Item("cd20").Specific
                                Dim st As String = oEdit.String
                                If st = "" Then
                                    SBO_Application.StatusBar.SetText("Consolidate Invoice Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If

                                Dim sqlstr As String = "SELECT T1.[ExtraMonth], T1.[ExtraDays] FROM OCRD T0 INNER JOIN OCTG T1 ON T0.[GroupNum] = T1.[GroupNum] WHERE T0.[CardCode] ='" & CustName & "'"
                                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery(sqlstr)
                                Dim Mon As Integer = 0
                                Dim Days As Integer = 0
                                If oRecordSet.RecordCount <> 0 Then
                                    Mon = oRecordSet.Fields.Item(0).Value
                                    Days = oRecordSet.Fields.Item(1).Value
                                End If
                                Dim Dt As DateTime
                                Dt = Convert.ToDateTime(st)
                                Dt = Dt.AddDays((Mon * 30) + Days)
                                oEdit = oForm.Items.Item("cdd").Specific
                                oEdit.String = Dt.ToString("dd/MM/yy")
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try

                        End If

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
                                If pVal.ItemUID = "vn4" Then
                                    Try
                                        oEdit = oForm.Items.Item("vn4").Specific
                                        oEdit.String = oDataTable.GetValue("ItemName", 0)
                                    Catch ex As Exception
                                    End Try
                                End If

                                If pVal.ItemUID = "cc16" Then
                                    Try

                                        oEdit = oForm.Items.Item("cc16").Specific
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
        Catch ex As Exception

        End Try
    End Sub
    Public Sub Update_PONo(ByVal DocNum As String, ByVal JobNo As String)
        Try
            Dim sqlstr As String = "SELECT distinct(T1.[U_NumAtCar]) 'NumCard' FROM [dbo].[@AIGI]  T0 , [dbo].[@AIGI1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[U_JobNo] ='" & JobNo & "'"
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sqlstr)
            Dim i As Integer = 0
            Dim PoNo As String = ""
            For i = 1 To oRecordSet.RecordCount
                If PoNo = "" Then
                    PoNo = oRecordSet.Fields.Item(0).Value
                Else
                    PoNo = PoNo & "," & oRecordSet.Fields.Item(0).Value
                End If
                oRecordSet.MoveNext()
            Next
            PoNo = PoNo.Replace("'", "''")

            sqlstr = "update OINV set U_AB_Remarks='" & PoNo & "' where DocEntry='" & DocNum & "'"
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sqlstr)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End Try
    End Sub
    Private Sub ConslInvReport1()
        Try
            oForm = SBO_Application.Forms.Item("InvoiceNew")
            Dim CNo, ConDate, ConDueDt As String
            oMatrix = oForm.Items.Item("22").Specific
            oColumns = oMatrix.Columns
            If oMatrix.RowCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data found for Generate Consolidate Invoice!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oEdit = oForm.Items.Item("ic10").Specific
            CNo = oEdit.String
            If CNo = "" Then
                SBO_Application.StatusBar.SetText("Consolidate Invoice No. Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oEdit = oForm.Items.Item("cd20").Specific
            ConDate = oEdit.Value
            If ConDate = "" Then
                SBO_Application.StatusBar.SetText("Consolidate Invoice Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oEdit = oForm.Items.Item("cdd").Specific
            ConDueDt = oEdit.Value
            If ConDueDt = "" Then
                SBO_Application.StatusBar.SetText("Consolidate Invoice Due Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim DocEntry As String = ""
            Dim i As Integer = 0
            Dim DocNum As String = ""
            Dim DocCurr As String = ""
            Dim DocCurr1 As String = ""
            Dim JobNo As String = ""
            oMatrix = oForm.Items.Item("22").Specific
            For i = 1 To oMatrix.RowCount
                oCheck = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific
                If oCheck.Checked = True Then
                    oEdit = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific
                    JobNo = oEdit.String
                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                    DocCurr = oEdit.String
                    If DocCurr1 <> "" Then
                        If DocCurr1 <> DocCurr Then
                            SBO_Application.StatusBar.SetText("Multi Currency not Supported for Consolidate Invoice!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                    End If
                    DocCurr1 = DocCurr
                    oEdit = oMatrix.Columns.Item("V_4").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        Update_PONo(oEdit.String, JobNo)
                    End If
                    If DocNum <> "" Then
                        DocNum = DocNum & "," & oEdit.String

                    Else
                        DocNum = oEdit.String
                    End If
                End If
            Next
            If DocNum = "" Then
                SBO_Application.StatusBar.SetText("No Invoices Selected!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim sqlstr As String = "Select U_C_INVNo from OINV where U_C_INVNo='" & CNo & "'"
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sqlstr)
            If oRecordSet.RecordCount > 0 Then
                SBO_Application.StatusBar.SetText("This Invoices No. Already Exists!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            sqlstr = "update OINV set U_C_INVNo='" & CNo & "' , U_C_DocDate='" & ConDate & "',U_C_DocDue='" & ConDueDt & "' where DocEntry in (" & DocNum & ")"
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sqlstr)

            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table

            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(sPath & "\GK_FM\ConsolInvoice.rpt")
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = (DocNum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item("@DocEntry")
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
            RptFrm.Text = "Consolidated Invoice Report"
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ConslInvReport()
        Try
            oForm = SBO_Application.Forms.Item("InvoiceNew")
            Dim DocNum As String = ""
            Dim CNo As String = ""
            oEdit = oForm.Items.Item("invser").Specific
            CNo = oEdit.String
            If CNo = "" Then
                SBO_Application.StatusBar.SetText("Consolidate Invoice No. Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim i As Integer = 0
            Dim sqlstr As String = "Select DocEntry from OINV where U_C_INVNo='" & CNo & "'"
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sqlstr)
            If oRecordSet.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("This Invoices No. Not Exists!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            For i = 1 To oRecordSet.RecordCount
                If DocNum <> "" Then
                    DocNum = DocNum & "," & oRecordSet.Fields.Item(0).Value
                Else
                    DocNum = oRecordSet.Fields.Item(0).Value
                End If
                oRecordSet.MoveNext()
            Next


            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table

            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(sPath & "\GK_FM\ConsolInvoice.rpt")
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = (DocNum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item("@DocEntry")
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
            RptFrm.Text = "Consolidated Invoice Report"
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ConslInvoiceReport(ByVal DocNum As String)
        Try

            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table

            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(sPath & "\GK_FM\ConsolInvoice.rpt")
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = (DocNum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item("@DocEntry")
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
            RptFrm.Text = "Consolidated Invoice Report"
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
