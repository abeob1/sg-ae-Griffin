Public Class PaymentVoucher
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Sub New()
        ' TODO: Complete member initialization 
    End Sub
    Public Sub PV_Bind(ByVal oForm As SAPbouiCOM.Form, ByVal SBO_Application As SAPbouiCOM.Application, ByVal TypeDiv As String, ByVal ocompany As SAPbobsCOM.Company)
        Try

            DocNumber_PV(TypeDiv, oForm, ocompany)
            oCombo = oForm.Items.Item("10").Specific
            'oCombo.ValidValues.Add("Cheque", "C")
            'oCombo.ValidValues.Add("Cash", "Ch")
            oCombo.Select("CQ", SAPbouiCOM.BoSearchKey.psk_ByDescription)
            oCombo = oForm.Items.Item("18").Specific
            ComboLoad_Currency(oForm, oCombo, ocompany)
            oCombo.Select("SGD", SAPbouiCOM.BoSearchKey.psk_ByValue)

            '-----PV
            CFL_Item1(oForm, SBO_Application)
            oMatrix = oForm.Items.Item("23").Specific
            oColumns = oMatrix.Columns
            oMatrix.AddRow()
            oColumn = oColumns.Item("V_0")
            oColumn.ChooseFromListUID = "1OITM"
            oColumn.ChooseFromListAlias = "ItemCode"
            oEdit = oForm.Items.Item("13").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            oEdit = oForm.Items.Item("17").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            CFL_BP_Supplier(oForm, SBO_Application)
            oEdit = oForm.Items.Item("4").Specific
            oEdit.ChooseFromListUID = "CFLBPV"
            oEdit.ChooseFromListAlias = "CardCode"
            ' oForm.DataBrowser.BrowseBy = "SIJ16"
        Catch ex As Exception
            Functions.WriteLog("Class:F_AE_JobOrder" + " Function:AE_Job_Bind" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub DocNumber_PV(ByVal Type As String, ByVal oform As SAPbouiCOM.Form, ByVal ocompany As SAPbobsCOM.Company)
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy") & "-01-01"

            tdt = Format(Now.Date, "yyyy") & "-12-31"
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oform.Items.Item("20").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = Type & Format(Now.Date, "yy") & "V" & "00001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = Type & Format(Now.Date, "yy") & "V" & "0000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = Type & Format(Now.Date, "yy") & "V" & "000" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = Type & Format(Now.Date, "yy") & "V" & "00" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = Type & Format(Now.Date, "yy") & "V" & "0" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = Type & Format(Now.Date, "yy") & "V" & DocNum
            End If



        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Private Sub ComboLoad_Currency(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox, ByVal ocompany As SAPbobsCOM.Company)
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
            oCombo.Select("SGD", SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try

            If pVal.FormUID = "AB_PV" Then
                oForm = SBO_Application.Forms.Item("AB_PV")
                If pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oForm.Close()
                        oForm = Nothing
                    End If
                End If
                '----------validate
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    If pVal.ItemUID = "23" And pVal.ColUID = "V_0" Then
                        oMatrix = oForm.Items.Item("23").Specific
                        oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                        If oEdit.String <> "" Then
                            oMatrix.AddRow()
                            oMatrix.ClearRowData(oMatrix.RowCount)
                            oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                            oEdit.String = ""
                        End If
                    End If

                    If pVal.ItemUID = "23" And (pVal.ColUID = "V_1" Or pVal.ColUID = "V_5") And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try
                            oMatrix = oForm.Items.Item("23").Specific
                            Dim i As Integer = 1
                            Dim amt As Double = 0
                            Dim TotAmt As Double = 0
                            Dim TotGST As Double = 0
                            Dim LineAmt As Double = 0
                            Dim TaxtAmt As Double = 0
                            For i = 1 To oMatrix.RowCount
                                oEdit = oMatrix.Columns.Item("V_7").Cells.Item(i).Specific
                                If oEdit.String <> "" Then
                                    oEdit = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific
                                    amt = amt + oEdit.Value
                                    LineAmt = oEdit.Value
                                    oEdit = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific
                                    TaxtAmt = TAXPer(oEdit.String, Ocompany) * LineAmt * (1 / 100)
                                    TotGST = TotGST + TaxtAmt
                                End If
                            Next
                            oEdit = oForm.Items.Item("15").Specific
                            oEdit.Value = amt
                            oEdit = oForm.Items.Item("25").Specific
                            oEdit.Value = TotGST
                            TotAmt = TotGST + amt
                            oEdit = oForm.Items.Item("27").Specific
                            oEdit.Value = TotAmt
                            'mitra
                        Catch ex As Exception
                        End Try

                    End If
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

                            If pVal.ItemUID = "23" And pVal.ColUID = "V_0" Then
                                oMatrix = oForm.Items.Item("23").Specific
                                oEdit = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemName", 0)
                                oEdit = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                oEdit.String = "SGD"
                                oEdit = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                oEdit.String = "ZP"
                                oEdit = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                oEdit.String = "1"
                                oEdit = oMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("ItemCode", 0)
                            ElseIf pVal.ItemUID = "4" Then
                                oEdit = oForm.Items.Item("6").Specific
                                oEdit.String = oDataTable.GetValue("CardName", 0)
                                oEdit = oForm.Items.Item("4").Specific
                                oEdit.String = oDataTable.GetValue("CardCode", 0)
                            End If
                        End If
                    Catch ex As Exception
                        ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class
