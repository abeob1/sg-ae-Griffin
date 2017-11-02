Imports System.Diagnostics.Process
Imports System.Threading
Public Class F_GoodsRelease
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Public Sub GoodsRelease_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.UserDataSources.Add("V_0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            'oMatrix1 = oForm.Items.Item("1000001").Specific
            'oColumns = oMatrix1.Columns
            'oMatrix1.AddRow()
            'oColumn = oColumns.Item("V_0")
            'oColumn.DataBind.SetBound(True, "", "oedit5")

            'oMatrix1 = oForm.Items.Item("1000001").Specific
            'oColumns = oMatrix1.Columns
            'oColumn = oColumns.Item("V_0")
            'oColumn.DataBind.SetBound(True, "", "V_0")
            'oItem = oForm.Items.Item("1000001")
            'oItem.Width = 150
            'oItem.Height = 30

            'oColumn.Width = 130
            'oMatrix1.AddRow()



            '            CFL_BP_Customer(oForm, SBO_Application)
            'CFL_BP_Customer(oForm, SBO_Application)
            CFL_BP_Supplier(oForm, SBO_Application)
            CFL_Item(oForm, SBO_Application)
            CFL_Item_Vessel(oForm, SBO_Application)

            'oEdit = oForm.Items.Item("35").Specific
            'oEdit.ChooseFromListUID = "CFLBPV"
            'oEdit.ChooseFromListAlias = "CardCode"
            oEdit = oForm.Items.Item("16").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            oEdit = oForm.Items.Item("18").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            oMatrix = oForm.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow()
            End If

            oColumn = oColumns.Item("V_0")
            oColumn.ChooseFromListUID = "OITM"
            oColumn.ChooseFromListAlias = "ItemCode"

            oCombo = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            ComboLoad_Unit(oForm, oCombo)
            oEdit = oForm.Items.Item("20").Specific
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemName"
            'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
            'oEdit.String = oMatrix.RowCount

            oEdit = oForm.Items.Item("20").Specific
            oEdit.ChooseFromListUID = "OITM11"
            oEdit.ChooseFromListAlias = "ItemName"

            oMatrix = oForm.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("V_15")
            oColumn.ChooseFromListUID = "CFLBPV"
            oColumn.ChooseFromListAlias = "CardCode"

            DocNumber_GRE()

            oItem = oForm.Items.Item("12")
            oItem.Enabled = False
            ''oItem = oForm.Items.Item("GRE4")
            'oItem.Enabled = True
            oItem = oForm.Items.Item("6")
            oItem.Enabled = True
            oMatrix = oForm.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("V_0")
            oColumn.Editable = True
            oColumn = oColumns.Item("V_1")
            oColumn.Editable = True
            oColumn = oColumns.Item("V_8")
            oColumn.Editable = True
            oColumn = oColumns.Item("V_12")
            oColumn.Editable = True
            oItem = oForm.Items.Item("GINo10")
            oItem.Enabled = True
            oItem = oForm.Items.Item("20")
            oItem.Enabled = True
            oForm.DataBrowser.BrowseBy = "12"
            ' loop1 = 0
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Public Sub DocNumber_GRE()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy-MM-dd")
            fdt = fdt.Substring(0, 8) & "01"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT CONVERT(VARCHAR(25),DATEADD(dd,-(DAY(DATEADD(mm,1,getdate()))),DATEADD(mm,1,getdate())),101)")
            tdt = oRecordSet1.Fields.Item(0).Value
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AIGE]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("26").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "GKWRE" & Format(Now.Date, "yyyyMMdd") & "0001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "GKWRE" & Format(Now.Date, "yyyyMMdd") & "000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "GKWRE" & Format(Now.Date, "yyyyMMdd") & "00" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "GKWRE" & Format(Now.Date, "yyyyMMdd") & "0" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "GKWRE" & Format(Now.Date, "yyyyMMdd") & "" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = "GKWRE" & Format(Now.Date, "yyyyMMdd") & DocNum
            End If

        Catch ex As Exception

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
    Private Sub CallReport()
        Try

            oEdit = oForm.Items.Item("RO4").Specific
            If oEdit.String = "" Then
                SBO_Application.StatusBar.SetText("Enter Document Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = False
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table

            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(sPath & "\GK_FM\RO_WHMS.rpt")
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = Convert.ToInt32(oEdit.Value)
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
            RptFrm.Text = "Release Order Report"
            RptFrm.TopMost = True
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub CallReport1()
        Try
            oEdit = oForm.Items.Item("RO4").Specific
            If oEdit.String = "" Then
                SBO_Application.StatusBar.SetText("Enter Document Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = False
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
            cryRpt.Load(sPath & "\GK_FM\RO_WHMS.rpt")
            ' cryRpt.RecordSelectionFormula = "{SP_AI_DeliveryOrder;1.DocEntry} ='" & oEdit.Value.ToString & "'"
            Dim ParaName As String = "@DocKey"
            Dim ParaValue As String = oEdit.String
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
            Dim pwd As String
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
            RptFrm.Text = "Release Order Report"
            RptFrm.TopMost = True
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = True
            RptFrm.Activate()
            RptFrm.ShowDialog()
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = True
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = True
        End Try
    End Sub
    Public Sub ROReport()
        Try

            oEdit = oForm.Items.Item("RO4").Specific
            If oEdit.String = "" Then
                SBO_Application.StatusBar.SetText("Enter Document Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = False
            Dim DocNum As Integer = oEdit.Value
            Dim Sqlstr As String

            Try
                SBO_Application.StatusBar.SetText("Retrieving Data", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Sqlstr = "SELECT T0.[U_VesselNo],T0.[U_Comments],T0.[U_CardCode], T0.[U_CardName], T0.[U_CntctCode], T0.[U_ShipTo], T0.[U_ANSRecNo] as DocNum, T0.[U_TaxDate], T1.[U_ItemCode], T1.[U_Decript], T1.[U_Qyt], T1.[U_Unit], T1.[U_Weight],( Cast(T1.[U_Length] as varchar) +'X'+ cast(T1.[U_Width] as varchar) + 'X' +Cast(T1.[U_Height] as Varchar)) as U_Dimen, T1.[U_Volume] FROM [dbo].[@AIGE]  T0 , [dbo].[@AIGE1]  T1 WHERE T1.[U_ItemCode] <>'' and T1.[DocEntry] = T0.[DocEntry]  and  T0.[DocNum] ='" & DocNum & "'"

                Dim frm As MY_Report
                frm = New MY_Report
                'frm.DO_Report(Sqlstr, Ocompany)
                frm.Text = "Release Order Report"
                frm.TopMost = True
                oItem = oForm.Items.Item("Print")
                oItem.Enabled = True
                frm.Activate()
                frm.ShowDialog()
            Catch ex As Exception
                oItem = oForm.Items.Item("Print")
                oItem.Enabled = True
            End Try
        Catch ex As Exception
            oItem = oForm.Items.Item("Print")
            oItem.Enabled = True
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try

            If (pVal.FormType = 0 And pVal.ItemUID = "1" And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                Try
                    Dim oOrderForm As SAPbouiCOM.Form
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(0, pVal.FormTypeCount)
                    oItem = oOrderForm.Items.Item("3")
                    If oItem.Visible = True Then

                        If oForm.UniqueID = "AI_FI_GoodsRelease" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim i As Integer
                            Try
                                DocNumber_GRE()
                                'If SBO_Application.MessageBox("You Cannot Change this Document after you have add it.Continue?", 1, "Yes", "No") = 2 Then
                                '    BubbleEvent = False
                                '    Exit Sub
                                'End If

                                oEdit = oForm.Items.Item("GRE4").Specific
                                If oEdit.String = "" Then
                                    SBO_Application.StatusBar.SetText("Customer Code can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oEdit = oForm.Items.Item("18").Specific
                                If oEdit.String = "" Then
                                    SBO_Application.StatusBar.SetText("Document Date can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oMatrix = oForm.Items.Item("29").Specific
                                oEdit = oMatrix.Columns.Item("V_0").Cells.Item(1).Specific
                                If oEdit.String = "" Then
                                    SBO_Application.StatusBar.SetText("Item Code Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oEdit = oMatrix.Columns.Item("V_8").Cells.Item(1).Specific
                                If oEdit.String = "" Then
                                    SBO_Application.StatusBar.SetText("Quantity Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(1).Specific
                                If oEdit.String = "" Then
                                    SBO_Application.StatusBar.SetText("Wareshouse Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oMatrix = oForm.Items.Item("29").Specific
                                For i = 1 To oMatrix.RowCount
                                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                                    If oEdit.String <> "" Then
                                        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
                                        If oEdit.String = "" Then
                                            SBO_Application.StatusBar.SetText("Wareshouse Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        Try
                                            oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                                            If oEdit.String = "" Or oEdit.Value = "0" Then
                                                SBO_Application.StatusBar.SetText("Quantity Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Catch ex As Exception
                                        End Try
                                    End If
                                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                                    If oEdit.String = "" Then
                                        oEdit = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific
                                        If oEdit.String <> "" Then
                                            SBO_Application.StatusBar.SetText("Item Code Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                Next

                                'Dim oIGN As SAPbobsCOM.Documents
                                'oIGN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                                'oIGN.DocDate = Now.Date
                                'oIGN.TaxDate = Now.Date
                                'oEdit = oForm.Items.Item("GRE4").Specific
                                'oIGN.CardCode = oEdit.String

                                'oMatrix = oForm.Items.Item("29").Specific
                                'For i = 1 To oMatrix.RowCount
                                '    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                                '    If oEdit.String <> "" Then
                                '        oIGN.Lines.ItemCode = oEdit.String
                                '        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
                                '        oIGN.Lines.WarehouseCode = oEdit.String
                                '        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                                '        oIGN.Lines.Quantity = oEdit.Value
                                '        oIGN.Lines.Add()
                                '    End If

                                'Next
                                'oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oRecordSet2.DoQuery("SELECT max(T0.[DocNum]) +1 FROM [dbo].[@AIGI]  T0")
                                'oIGN.Comments = " Based on Goods Receipt No: " & oRecordSet2.Fields.Item(0).Value & ""
                                'Dim RetCode As Integer = oIGN.Add()
                                'Dim SerrorMsg As String = ""
                                'Ocompany.GetLastError(RetCode, SerrorMsg)
                                'If RetCode <> 0 Then
                                '    SBO_Application.StatusBar.SetText(Ocompany.GetLastErrorDescription & Ocompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '    BubbleEvent = False
                                '    Exit Sub
                                'End If
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                            ' Dim i As Integer
                            For i = 1 To oMatrix.RowCount
                                Try
                                    Dim oldQty As Integer = 0
                                    Dim NewQty As Integer = 0
                                    Dim qty As Integer = 0
                                    Dim LineID As Integer = 0
                                    Dim DocEntry As Integer = 0
                                    oMatrix = oForm.Items.Item("29").Specific
                                    oEdit = oMatrix.Columns.Item("V_9").Cells.Item(i).Specific
                                    DocEntry = oEdit.String
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                                    LineID = oEdit.String
                                    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                                    qty = oEdit.String
                                    Dim oRecordSet_OPU As SAPbobsCOM.Recordset
                                    oRecordSet_OPU = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim oRecordSet_OP As SAPbobsCOM.Recordset
                                    oRecordSet_OP = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet_OP.DoQuery("SELECT T0.[U_OpenQty] FROM [dbo].[@AIGR1]  T0 WHERE T0.[DocEntry] ='" & DocEntry & "' and  T0.[LineId] ='" & LineID & "'")
                                    If oRecordSet_OP.RecordCount <> 0 Then
                                        oldQty = oRecordSet_OP.Fields.Item(0).Value
                                        NewQty = oldQty - qty
                                        If NewQty < 0 Then
                                            NewQty = 0
                                            'SBO_Application.StatusBar.SetText("Line ID : " & LineID & " Quantity falls Below Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            'BubbleEvent = False
                                            'Exit Sub
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Next
                            For i = 1 To oMatrix.RowCount
                                Try

                                    Dim oldQty As Integer = 0
                                    Dim NewQty As Integer = 0
                                    Dim qty As Integer = 0
                                    Dim LineID As Integer = 0
                                    Dim DocEntry As Integer = 0
                                    oMatrix = oForm.Items.Item("29").Specific
                                    oEdit = oMatrix.Columns.Item("V_9").Cells.Item(i).Specific
                                    DocEntry = oEdit.String
                                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                                    LineID = oEdit.String
                                    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                                    qty = oEdit.String
                                    Dim oRecordSet_OPU As SAPbobsCOM.Recordset
                                    oRecordSet_OPU = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim oRecordSet_OP As SAPbobsCOM.Recordset
                                    oRecordSet_OP = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet_OP.DoQuery("SELECT T0.[U_OpenQty] FROM [dbo].[@AIGR1]  T0 WHERE T0.[DocEntry] ='" & DocEntry & "' and  T0.[LineId] ='" & LineID & "'")
                                    If oRecordSet_OP.RecordCount <> 0 Then
                                        oldQty = oRecordSet_OP.Fields.Item(0).Value
                                        NewQty = oldQty - qty
                                        If NewQty < 0 Then
                                            NewQty = 0
                                            'SBO_Application.StatusBar.SetText("Line ID : " & LineID & " Quantity falls Below Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            'BubbleEvent = False
                                            'Exit Sub
                                        End If
                                        oRecordSet_OPU.DoQuery("UPDATE [@AIGR1] SET [U_OpenQty]='" & NewQty & "' WHERE [DocEntry] ='" & DocEntry & "' and  [LineId] ='" & LineID & "' ")
                                        oRecordSet_OPU = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet_OPU.DoQuery("SELECT T0.[U_OpenQty] FROM [dbo].[@AIGR1]  T0 WHERE T0.[U_OpenQty]   <> 0 and  T0.[DocEntry] ='" & DocEntry & "'")
                                        If oRecordSet_OPU.RecordCount = 0 Then
                                            oRecordSet_OP = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet_OP.DoQuery("UPDATE [@AIGR] SET U_Status='Closed' where [DocEntry] ='" & DocEntry & "'")
                                        End If
                                        oEdit = oForm.Items.Item("GINo10").Specific
                                        Dim DocEn As String = oEdit.String
                                        'Nath

                                        Dim DONo As String = DocEn.Substring(0, 4)
                                        If DONo = "GKWI" Then
                                            oRecordSet_OP.DoQuery("Update [@AB_SEAI_JOB_H] set U_DOStatus='Closed' where U_DONo ='" & DocEn & "'")
                                        ElseIf DONo = "GKWE" Then
                                            oRecordSet_OP.DoQuery("Update [@AB_SEAE_JOB_H] set U_DOStatus='Closed' where U_DONo ='" & DocEn & "'")
                                        End If

                                    End If
                                Catch ex As Exception
                                End Try
                            Next
                        End If

                    End If
                Catch ex As Exception

                End Try
            End If
            '=====================
            Try
                If pVal.FormType = 2000036 Then
                    'If (pVal.ItemUID = "1" And pVal.Before_Action = False And pVal.InnerEvent = False And SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Or (pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    If (pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) Then
                        oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
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
                                        'MatrixLoad(No(i), VenCode(i), PO(i))
                                    End If
                                Catch ex As Exception
                                End Try

                            Next
                        Catch ex As Exception

                        End Try
                    End If
                End If
                If pVal.FormUID = "GRNo" Then
                    oForm = SBO_Application.Forms.Item("GRNo")
                    If pVal.ItemUID = "5" And pVal.Before_Action = False And pVal.InnerEvent = False And SAPbouiCOM.BoEventTypes.et_FORMAT_SEARCH_COMPLETED Then
                        'MsgBox("Hi")

                    End If
                    If pVal.ItemUID = "Choose" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        oMatrix = oForm.Items.Item("3").Specific
                        Dim i As Integer
                        Dim No(oMatrix.RowCount) As String
                        Dim VenCode(oMatrix.RowCount) As String
                        Dim PO(oMatrix.RowCount) As String
                        Dim k As Integer = 0
                        For i = 1 To oMatrix.RowCount
                            If oMatrix.IsRowSelected(i) = True Then
                                oEdit = oMatrix.Columns.Item("V_11").Cells.Item(i).Specific
                                No(k) = oEdit.String
                                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
                                VenCode(k) = oEdit.String
                                oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                                PO(k) = oEdit.String
                                k = k + 1
                            End If

                        Next
                        'k = oMatrix.RowCount
                        oForm.Close()

                        For i = 0 To k + 1
                            Try
                                If No(i) <> "" Then
                                    '  MatrixLoad(No(i), VenCode(i), PO(i))
                                End If
                            Catch ex As Exception
                            End Try

                        Next

                    End If


                End If
            Catch ex As Exception
                '  SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
            '========================
            '===============


            If pVal.FormUID = "RO_Report" And pVal.ItemUID = "Print" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                oForm = SBO_Application.Forms.Item("RO_Report")


                Dim trd As Threading.Thread
                trd = New Threading.Thread(AddressOf CallReport)
                trd.IsBackground = True
                trd.SetApartmentState(ApartmentState.STA)
                trd.Start()
            End If



            '====================

            If pVal.FormUID = "AI_FI_GoodsRelease" Then
                oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")

                'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = True
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = False
                '    'oItem = oForm.Items.Item("GRE4")
                '    oItem.Enabled = True
                '    oItem = oForm.Items.Item("6")
                '    oItem.Enabled = True
                '    oMatrix = oForm.Items.Item("29").Specific
                '    oColumns = oMatrix.Columns
                '    oColumn = oColumns.Item("V_0")
                '    oColumn.Editable = True
                '    oColumn = oColumns.Item("V_1")
                '    oColumn.Editable = True
                '    oColumn = oColumns.Item("V_8")
                '    oColumn.Editable = True
                '    oColumn = oColumns.Item("V_12")
                '    oColumn.Editable = True
                'End If
                'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '    'oItem = oForm.Items.Item("GRE4")
                '    oItem.Enabled = False
                '    oItem = oForm.Items.Item("6")
                '    oItem.Enabled = False
                '    oMatrix = oForm.Items.Item("29").Specific
                '    oColumns = oMatrix.Columns
                '    oColumn = oColumns.Item("V_0")
                '    oColumn.Editable = False
                '    oColumn = oColumns.Item("V_1")
                '    oColumn.Editable = False
                '    oColumn = oColumns.Item("V_8")
                '    oColumn.Editable = False
                '    oColumn = oColumns.Item("V_12")
                '    oColumn.Editable = False
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = False
                'End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then
                    'oForm.Freeze(True)
                    'oMatrix1 = oForm.Items.Item("1000001").Specific
                    'oColumns = oMatrix1.Columns
                    'oColumn = oColumns.Item("V_0")
                    'oItem = oForm.Items.Item("1000001")
                    'oItem.Width = 150
                    'oItem.Height = 30
                    'oColumn.Width = 130
                    'oForm.Freeze(False)
                End If
                'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                '    'oItem = oform.Items.Item("1")
                '    'oItem.Enabled = True
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = False
                '    'oItem = oForm.Items.Item("GRE4")
                '    oItem.Enabled = True
                '    oItem = oForm.Items.Item("6")
                '    oItem.Enabled = True
                '    oMatrix = oForm.Items.Item("29").Specific
                '    oColumns = oMatrix.Columns
                '    oColumn = oColumns.Item("V_0")
                '    oColumn.Editable = True
                '    oColumn = oColumns.Item("V_1")
                '    oColumn.Editable = True
                '    oColumn = oColumns.Item("V_8")
                '    oColumn.Editable = True
                '    oColumn = oColumns.Item("V_12")
                '    oColumn.Editable = True
                '    oMatrix1 = oForm.Items.Item("1000001").Specific
                '    'If oMatrix.RowCount = 0 Then

                '    'End If
                '    oEdit = oForm.Items.Item("26").Specific
                '    If oEdit.String = "" Then
                '        DocNumber_GI()
                '    End If
                'End If
                'If pVal.ItemUID = "40" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                '    ' CopyFromGoodsReceipt()
                '    SBO_Application.ActivateMenuItem(7425)
                'End If
                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oEdit = oForm.Items.Item("16").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oEdit = oForm.Items.Item("18").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oMatrix = oForm.Items.Item("29").Specific
                    oColumns = oMatrix.Columns
                    oMatrix.AddRow()
                    DocNumber_GRE()
                    'oMatrix1 = oForm.Items.Item("1000001").Specific
                    'oMatrix1.AddRow()

                End If
                'Mitra
                If pVal.ItemUID = "GINo10" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("GINo10").Specific
                    Dim GINO As String = oEdit.String
                    If GINO <> "" Then
                        MatrixLoad(GINO)
                    End If


                End If
                If pVal.ItemUID = "1000001" And pVal.ColUID = "V_0" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("20").Specific
                    Dim Vessel As String = oEdit.String

                    oMatrix1 = oForm.Items.Item("1000001").Specific
                    oEdit = oMatrix1.Columns.Item("V_0").Cells.Item(1).Specific
                    If oEdit.String <> "" Then
                        oEdit = oForm.Items.Item("GRE4").Specific
                        If oEdit.String <> "" Then

                            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet1.DoQuery("SELECT distinct T0.[DocNum], T0.[U_CardCode], T0.[U_CardName], T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_TaxDate] , T0.[U_ANSRecNo], T1.[U_VenName] FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry]  and T0.[U_Status]='Open' and  T1.[U_Decript] <>'' and T0.[U_CardCode] ='" & oEdit.String & "' and  T0.[U_VesselNo] ='" & Vessel & "' and  T1.U_OpenQty>0 ORDER BY T0.[DocNum]")
                            If oRecordSet1.RecordCount = 1 Then
                                ' MatrixLoad(oRecordSet1.Fields.Item(0).Value, oRecordSet1.Fields.Item(8).Value, oRecordSet1.Fields.Item(3).Value)
                            End If
                        End If
                    End If
                End If
                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim i As Integer
                    Try
                        DocNumber_GRE()
                        If SBO_Application.MessageBox("You Cannot Change this Document after you have add it.Continue?", 1, "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        oEdit = oForm.Items.Item("GRE4").Specific
                        If oEdit.String = "" Then
                            SBO_Application.StatusBar.SetText("Customer Code can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oEdit = oForm.Items.Item("18").Specific
                        If oEdit.String = "" Then
                            SBO_Application.StatusBar.SetText("Document Date can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oMatrix = oForm.Items.Item("29").Specific
                        oEdit = oMatrix.Columns.Item("V_0").Cells.Item(1).Specific
                        If oEdit.String = "" Then
                            SBO_Application.StatusBar.SetText("Item Code Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(1).Specific
                        If oEdit.String = "" Then
                            SBO_Application.StatusBar.SetText("Quantity Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(1).Specific
                        If oEdit.String = "" Then
                            SBO_Application.StatusBar.SetText("Wareshouse Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oMatrix = oForm.Items.Item("29").Specific
                        For i = 1 To oMatrix.RowCount
                            oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                            If oEdit.String <> "" Then
                                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
                                If oEdit.String = "" Then
                                    SBO_Application.StatusBar.SetText("Wareshouse Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                Try
                                    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                                    If oEdit.String = "" Or oEdit.Value = "0" Then
                                        SBO_Application.StatusBar.SetText("Quantity Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                Catch ex As Exception
                                End Try
                            End If
                            oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                            If oEdit.String = "" Then
                                oEdit = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    SBO_Application.StatusBar.SetText("Item Code Can't Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        Next

                        'Dim oIGN As SAPbobsCOM.Documents
                        'oIGN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        'oIGN.DocDate = Now.Date
                        'oIGN.TaxDate = Now.Date
                        'oEdit = oForm.Items.Item("GRE4").Specific
                        'oIGN.CardCode = oEdit.String

                        'oMatrix = oForm.Items.Item("29").Specific
                        'For i = 1 To oMatrix.RowCount
                        '    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                        '    If oEdit.String <> "" Then
                        '        oIGN.Lines.ItemCode = oEdit.String
                        '        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
                        '        oIGN.Lines.WarehouseCode = oEdit.String
                        '        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                        '        oIGN.Lines.Quantity = oEdit.Value
                        '        oIGN.Lines.Add()
                        '    End If

                        'Next
                        'oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRecordSet2.DoQuery("SELECT max(T0.[DocNum]) +1 FROM [dbo].[@AIGI]  T0")
                        'oIGN.Comments = " Based on Goods Receipt No: " & oRecordSet2.Fields.Item(0).Value & ""
                        'Dim RetCode As Integer = oIGN.Add()
                        'Dim SerrorMsg As String = ""
                        'Ocompany.GetLastError(RetCode, SerrorMsg)
                        'If RetCode <> 0 Then
                        '    SBO_Application.StatusBar.SetText(Ocompany.GetLastErrorDescription & Ocompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        Exit Sub
                    End Try
                    ' Dim i As Integer

                    For i = 1 To oMatrix.RowCount
                        Try
                            Dim oldQty As Integer = 0
                            Dim NewQty As Integer = 0
                            Dim qty As Integer = 0
                            Dim LineID As Integer = 0
                            Dim DocEntry As Integer = 0
                            oMatrix = oForm.Items.Item("29").Specific
                            oEdit = oMatrix.Columns.Item("V_9").Cells.Item(i).Specific
                            DocEntry = oEdit.String
                            oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                            LineID = oEdit.String
                            oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                            qty = oEdit.String
                            Dim oRecordSet_OPU As SAPbobsCOM.Recordset
                            oRecordSet_OPU = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim oRecordSet_OP As SAPbobsCOM.Recordset
                            oRecordSet_OP = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet_OP.DoQuery("SELECT T0.[U_OpenQty] FROM [dbo].[@AIGR1]  T0 WHERE T0.[DocEntry] ='" & DocEntry & "' and  T0.[LineId] ='" & LineID & "'")
                            If oRecordSet_OP.RecordCount <> 0 Then
                                oldQty = oRecordSet_OP.Fields.Item(0).Value
                                NewQty = oldQty - qty
                                If NewQty < 0 Then
                                    NewQty = 0
                                    'SBO_Application.StatusBar.SetText("Line ID : " & LineID & " Quantity falls Below Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    'BubbleEvent = False
                                    'Exit Sub
                                End If
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                    For i = 1 To oMatrix.RowCount
                        Try

                            Dim oldQty As Integer = 0
                            Dim NewQty As Integer = 0
                            Dim qty As Integer = 0
                            Dim LineID As Integer = 0
                            Dim DocEntry As Integer = 0
                            oMatrix = oForm.Items.Item("29").Specific
                            oEdit = oMatrix.Columns.Item("V_9").Cells.Item(i).Specific
                            DocEntry = oEdit.String
                            oEdit = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific
                            LineID = oEdit.String

                            oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                            qty = oEdit.String
                            Dim oRecordSet_OPU As SAPbobsCOM.Recordset
                            oRecordSet_OPU = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim oRecordSet_OP As SAPbobsCOM.Recordset
                            oRecordSet_OP = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet_OP.DoQuery("SELECT T0.[U_OpenQty] FROM [dbo].[@AIGR1]  T0 WHERE T0.[DocEntry] ='" & DocEntry & "' and  T0.[LineId] ='" & LineID & "'")
                            If oRecordSet_OP.RecordCount <> 0 Then
                                oldQty = oRecordSet_OP.Fields.Item(0).Value
                                NewQty = oldQty - qty
                                If NewQty < 0 Then
                                    NewQty = 0
                                End If
                                'mitra
                                oRecordSet_OPU.DoQuery("UPDATE [@AIGR1] SET [U_OpenQty]='" & NewQty & "' WHERE [DocEntry] ='" & DocEntry & "' and  [LineId] ='" & LineID & "' ")
                                oRecordSet_OPU = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet_OPU.DoQuery("SELECT T0.[U_OpenQty] FROM [dbo].[@AIGR1]  T0 WHERE T0.[U_OpenQty]   <> 0 and  T0.[DocEntry] ='" & DocEntry & "'")
                                If oRecordSet_OPU.RecordCount = 0 Then
                                    oRecordSet_OP = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet_OP.DoQuery("UPDATE [@AIGR] SET U_Status='Closed' where [DocEntry] ='" & DocEntry & "'")
                                End If
                                oEdit = oForm.Items.Item("GINo10").Specific
                                Dim DocEn As String = oEdit.String
                                Dim DONo As String = DocEn.Substring(0, 4)
                                If DONo = "GKWI" Then
                                    oRecordSet_OP.DoQuery("Update [@AB_SEAI_JOB_H] set U_DOStatus='Closed' where U_DONo ='" & DocEn & "'")
                                ElseIf DONo = "GKWE" Then
                                    oRecordSet_OP.DoQuery("Update [@AB_SEAE_JOB_H] set U_DOStatus='Closed' where U_DONo ='" & DocEn & "'")
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                End If
                If pVal.ItemUID = "GRE40" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("GRE4").Specific
                    Dim CardCode As String = oEdit.String.Trim
                    If CardCode <> "" Then
                        'oEdit = oForm.Items.Item("GRE40").Specific
                        'If oEdit.String <> "" Then
                        Dim DocNum As Integer = oEdit.String
                        Dim i As Integer
                        oMatrix = oForm.Items.Item("29").Specific
                        For i = 1 To oMatrix.RowCount
                            oEdit = oMatrix.Columns.Item("V_9").Cells.Item(i).Specific
                            Dim NewDocNum As Integer = 0
                            If oEdit.String <> "" Then
                                NewDocNum = oEdit.String
                            End If

                            If NewDocNum = DocNum Then
                                SBO_Application.StatusBar.SetText("This Record Already Entered..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        Next
                        Dim oRecordSet_GR As SAPbobsCOM.Recordset
                        oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet_GR.DoQuery("SELECT T0.[U_NumAtCard], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T0.[U_VenCode], T0.[U_VenName], T0.[U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId],T0.U_CardCode FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
                        If oRecordSet_GR.RecordCount = 0 Then
                            SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        Dim NewBPCode As String = oRecordSet_GR.Fields.Item(22).Value.ToString.Trim
                        If CardCode <> NewBPCode Then
                            SBO_Application.StatusBar.SetText("InValid BP Entered..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        oEdit = oForm.Items.Item("10").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(0).Value
                        oEdit = oForm.Items.Item("20").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(1).Value
                        oEdit = oForm.Items.Item("22").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(2).Value
                        oEdit = oForm.Items.Item("24").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(3).Value
                        'oEdit = oForm.Items.Item("26").Specific
                        'oEdit.String = oRecordSet_GR.Fields.Item(4).Value
                        oEdit = oForm.Items.Item("33").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(5).Value
                        oEdit = oForm.Items.Item("35").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(6).Value
                        oEdit = oForm.Items.Item("37").Specific
                        oEdit.String = oRecordSet_GR.Fields.Item(7).Value
                        Try
                            oEdit = oForm.Items.Item("31").Specific
                            If oEdit.String = "" Then
                                oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
                            Else
                                oEdit.String = oEdit.String & " Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
                            End If
                        Catch ex As Exception
                        End Try


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
                            oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                            oEdit.String = oRecordSet_GR.Fields.Item(16).Value
                            oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
                            oEdit.String = oRecordSet_GR.Fields.Item(18).Value
                            oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
                            oEdit.String = oRecordSet_GR.Fields.Item(17).Value
                            oEdit = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
                            oEdit.String = oRecordSet_GR.Fields.Item(20).Value
                            oEdit = oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific
                            oEdit.String = oRecordSet_GR.Fields.Item(21).Value
                            oMatrix.AddRow()
                            oRecordSet_GR.MoveNext()
                        Next
                        'oEdit = oForm.Items.Item("GRE40").Specific
                        'oEdit.String = ""
                        'End If
                    End If
                End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_0" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    If oEdit.String <> "" Then
                        oMatrix.AddRow()
                        'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                        'oEdit.String = oMatrix.RowCount
                    End If
                End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_3" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    Try

                        Dim l As Integer = 0
                        Dim b As Integer = 0
                        Dim w As Integer = 0
                        Dim vol As Double = 0.0
                        oEdit = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                        l = oEdit.Value
                        oEdit = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                        b = oEdit.Value
                        oEdit = oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                        w = oEdit.Value
                        oEdit = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                        oEdit.Value = ((l * b * w) / 1000000)
                    Catch ex As Exception

                    End Try

                End If
                If pVal.ItemUID = "GRE4" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("GRE4").Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oForm.Items.Item("6").Specific
                    If BPCode <> "" Then
                        oEdit.String = BPName(BPCode, Ocompany)
                    End If
                    ' End If
                    'If pVal.ItemUID = "GRE4" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    'oEdit = oForm.Items.Item("GRE4").Specific
                    'Dim BPCode As String = oEdit.String
                    oEdit = oForm.Items.Item("8").Specific
                    If BPCode <> "" Then
                        oEdit.String = ContactPerson(BPCode, Ocompany)
                    End If
                    'End If
                    'If pVal.ItemUID = "GRE4" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    'oEdit = oForm.Items.Item("GRE4").Specific
                    'Dim BPCode As String = oEdit.String
                    oEdit = oForm.Items.Item("33").Specific
                    If BPCode <> "" Then
                        oEdit.String = BPAddress(BPCode, Ocompany)
                    End If
                    'oEdit = oMatrix1.Columns.Item("V_-1").Cells.Item(1).Specific
                    'oEdit.String = ""
                End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_15" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oMatrix.Columns.Item("V_15").Cells.Item(pVal.Row).Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oMatrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific
                    If BPCode <> "" Then
                        oEdit.String = BPName(BPCode, Ocompany)
                    End If
                End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_0" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                    If BPCode <> "" Then
                        oEdit.String = ItemName(BPCode, Ocompany)
                    End If

                    'oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                    'oEdit.String = Deafault_Whsc(Ocompany)
                    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                    oEdit.String = "1"
                End If

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
                            If pVal.ItemUID = "20" Then
                                oEdit = oForm.Items.Item("20").Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                            End If
                            If pVal.ItemUID = "GRE4" Then
                                'oEdit = oForm.Items.Item("6").Specific
                                'oEdit.String = oDataTable.GetValue("CardName", 0)
                                'oEdit = oForm.Items.Item("8").Specific
                                'oEdit.String = ContactPerson(oDataTable.GetValue("CardCode", 0), Ocompany)
                                'oEdit = oForm.Items.Item("33").Specific
                                'oEdit.String = BPAddress(oDataTable.GetValue("CardCode", 0), Ocompany)
                                oEdit = oForm.Items.Item("GRE4").Specific
                                oEdit.String = oDataTable.GetValue("CardCode", 0)
                            End If
                            If pVal.ItemUID = "35" Then
                                'oEdit = oForm.Items.Item("37").Specific
                                'oEdit.String = oDataTable.GetValue("CardName", 0)
                                'oEdit = oForm.Items.Item("35").Specific
                                'oEdit.String = oDataTable.GetValue("CardCode", 0)
                            End If
                            If pVal.ItemUID = "29" And pVal.ColUID = "V_15" Then
                                'oEdit = oMatrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific
                                'oEdit.String = oDataTable.GetValue("CardName", 0)
                                oEdit = oMatrix.Columns.Item("V_15").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardCode", 0)
                            End If
                            If pVal.ItemUID = "29" And pVal.ColUID = "V_0" Then
                                'oEdit = oMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific
                                'oEdit.String = oDataTable.GetValue("ItemName", 0)
                                Try
                                    'Dim oRecordSet_Ow As SAPbobsCOM.Recordset
                                    'oRecordSet_Ow = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'oRecordSet_Ow.DoQuery("SELECT T0.[BLength1], T0.[BWidth1], T0.[BHeight1], T0.[BVolume], T0.[BWeight1] FROM OITM T0 WHERE T0.[ItemCode] ='" & oDataTable.GetValue("ItemCode", 0) & "'")
                                    'oEdit = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = "1"
                                    'oEdit = oMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = oRecordSet_Ow.Fields.Item(4).Value
                                    'oEdit = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = oRecordSet_Ow.Fields.Item(0).Value
                                    'oEdit = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = oRecordSet_Ow.Fields.Item(1).Value
                                    'oEdit = oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = oRecordSet_Ow.Fields.Item(2).Value
                                    'oEdit = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = oRecordSet_Ow.Fields.Item(3).Value
                                    'oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = Deafault_Whsc(Ocompany)
                                    'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(pVal.Row).Specific
                                    'oEdit.String = pVal.Row
                                Catch ex As Exception
                                End Try
                                Try
                                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific
                                    oEdit.String = oDataTable.GetValue("ItemCode", 0)
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                    Catch ex As Exception
                        ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If

        Catch ex As Exception
            ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
        Try


            If eventInfo.FormUID = "AI_FI_GoodsRelease" Then
                If (eventInfo.BeforeAction = True) Then
                    'Dim oMenuItem As SAPbouiCOM.MenuItem
                    'Dim oMenus As SAPbouiCOM.Menus
                    'oMenuItem.UID = ""
                    'oMenuItem.Enabled = True
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus


                    Try
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        'oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                        'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        'oCreationPackage.UniqueID = "OnlyOnRC"
                        'oCreationPackage.String = "Add Row"
                        'oCreationPackage.Enabled = True

                        'oMenuItem = SBO_Application.Menus.Item("1280") 'Data'
                        'oMenus = oMenuItem.SubMenus
                        'oMenus.AddEx(oCreationPackage)

                        ' Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "OnlyOnRC1"
                        oCreationPackage.String = "Delete Row"
                        oCreationPackage.Enabled = True

                        oMenuItem = SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try


            If pVal.MenuUID = "OnlyOnRC1" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If oForm.UniqueID = "AI_FI_GoodsRelease" Then
                            oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
                            oMatrix = oForm.Items.Item("29").Specific
                            Dim i As Integer = 0
                            If oMatrix.RowCount = 1 Then
                                oMatrix.DeleteRow(1)
                                oMatrix.AddRow()
                                Exit Sub
                            End If
                            For i = 1 To oMatrix.RowCount
                                If oMatrix.IsRowSelected(i) = True Then
                                    oMatrix.DeleteRow(i)
                                End If
                            Next
                        End If

                    End If
                Catch ex As Exception
                End Try
            End If
            If pVal.MenuUID = "OnlyOnRC" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If oForm.UniqueID = "AI_FI_GoodsRelease" Then
                            oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
                            oMatrix = oForm.Items.Item("29").Specific
                            oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix.AddRow()
                            End If
                        End If

                    End If
                Catch ex As Exception
                End Try
            End If
            If pVal.MenuUID = "1281" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    If oForm.UniqueID = "AI_FI_GoodsRelease" Then
                        oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
                        oItem = oForm.Items.Item("12")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("GRE4")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("6")
                        oItem.Enabled = True
                        oMatrix = oForm.Items.Item("29").Specific
                        oColumns = oMatrix.Columns
                        oColumn = oColumns.Item("V_0")
                        oColumn.Editable = True
                        oColumn = oColumns.Item("V_1")
                        oColumn.Editable = True
                        oColumn = oColumns.Item("V_8")
                        oColumn.Editable = True
                        oColumn = oColumns.Item("V_12")
                        oColumn.Editable = True
                    End If
                Catch ex As Exception

                End Try
            End If
            'addmode
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                If oForm.UniqueID = "AI_FI_GoodsRelease" Then
                    oItem = oForm.Items.Item("GRE4")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("GINo10")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("20")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("6")
                    oItem.Enabled = False
                    oMatrix = oForm.Items.Item("29").Specific
                    oColumns = oMatrix.Columns
                    oColumn = oColumns.Item("V_0")
                    oColumn.Editable = False
                    oColumn = oColumns.Item("V_1")
                    oColumn.Editable = False
                    oColumn = oColumns.Item("V_8")
                    oColumn.Editable = False
                    oColumn = oColumns.Item("V_12")
                    oColumn.Editable = False
                    oItem = oForm.Items.Item("12")
                    oItem.Enabled = False
                End If
            End If

            If pVal.MenuUID = "1282" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    If oForm.UniqueID = "AI_FI_GoodsRelease" Then
                        oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
                        oEdit = oForm.Items.Item("16").Specific
                        oEdit.String = Format(Now.Date, "dd/MM/yy")
                        oEdit = oForm.Items.Item("18").Specific
                        oEdit.String = Format(Now.Date, "dd/MM/yy")
                        oMatrix = oForm.Items.Item("29").Specific
                        oColumns = oMatrix.Columns
                        oMatrix.AddRow()
                        'oMatrix1 = oForm.Items.Item("1000001").Specific
                        'oMatrix1.AddRow()
                        oItem = oForm.Items.Item("12")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("GRE4")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("GINo10")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("20")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("6")
                        oItem.Enabled = True
                        oMatrix = oForm.Items.Item("29").Specific
                        oColumns = oMatrix.Columns
                        oColumn = oColumns.Item("V_0")
                        oColumn.Editable = True
                        oColumn = oColumns.Item("V_1")
                        oColumn.Editable = True
                        oColumn = oColumns.Item("V_8")
                        oColumn.Editable = True
                        oColumn = oColumns.Item("V_12")
                        oColumn.Editable = True
                        'oMatrix1 = oForm.Items.Item("1000001").Specific
                        'If oMatrix.RowCount = 0 Then

                        'End If
                        oEdit = oForm.Items.Item("26").Specific
                        If oEdit.String = "" Then
                            DocNumber_GRE()
                        End If
                    End If

                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub MatrixLoad(ByVal GINo As String)

        Dim oRecordSet_GR As SAPbobsCOM.Recordset
        oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_GR.DoQuery("SELECT T0.[U_ShipTo], T0.[U_Comments],T0.[U_CardCode], T0.[U_CardName], T0.[U_CntctCode], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T1.[U_ItemCode], T1.[U_Decript], T1.[U_Qyt], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BaseKey], T1.[U_BaseRow], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_VenCode], T1.[U_VenName], T1.[U_NumAtCar] FROM [dbo].[@AIGI]  T0 , [dbo].[@AIGI1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & GINo & "' and T0. U_Status='Open'")
        If oRecordSet_GR.RecordCount = 0 Then
            SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oEdit = oForm.Items.Item("GRE4").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_CardCode").Value
        oEdit = oForm.Items.Item("6").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_CardName").Value

        oEdit = oForm.Items.Item("8").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_CntctCode").Value
        oEdit = oForm.Items.Item("20").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_VesselNo").Value
        oEdit = oForm.Items.Item("22").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_MAWBNo").Value
        oEdit = oForm.Items.Item("24").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_POL").Value
        oEdit = oForm.Items.Item("31").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_Comments").Value
        oEdit = oForm.Items.Item("33").Specific
        oEdit.String = oRecordSet_GR.Fields.Item("U_ShipTo").Value
        Dim i As Integer = 0
        oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
        oMatrix = oForm.Items.Item("29").Specific
        oMatrix.Clear()
        For i = 1 To oRecordSet_GR.RecordCount
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow()
            End If
            'SELECT T0.[U_CardCode], T0.[U_CardName], T0.[U_CntctCode], T0.[U_VesselNo], T0.[U_MAWBNo], 
            'T0.[U_POL], T1.[U_ItemCode], T1.[U_Decript], T1.[U_Qyt], T1.[U_Unit], T1.[U_Weight], 
            'T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BaseKey], T1.[U_BaseRow],
            ' T1.[U_BinLoc], T1.[U_Whsc], T1.[U_VenCode], T1.[U_VenName], T1.[U_NumAtCar] FROM [dbo].[@AIGI]  T0 , [dbo].[@AIGI1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] =[%0] and T0. U_Status=
            oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_ItemCode").Value
            oEdit = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Decript").Value
            oEdit = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Qyt").Value
            Try
                oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                oCombo.Select(oRecordSet_GR.Fields.Item("U_Unit").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try

            oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Weight").Value
            oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Length").Value
            oEdit = oMatrix.Columns.Item("V_4").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Width").Value
            oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Height").Value
            Try
                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet_GR.Fields.Item("U_Volume").Value
            Catch ex As Exception
            End Try
            oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_Whsc").Value
            oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_BinLoc").Value
            oEdit = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_BaseKey").Value
            oEdit = oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_BaseRow").Value
            oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_VenCode").Value
            oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_VenName").Value
            oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
            oEdit.String = oRecordSet_GR.Fields.Item("U_NumAtCar").Value

            oMatrix.AddRow()
            oRecordSet_GR.MoveNext()
        Next
    End Sub
End Class
