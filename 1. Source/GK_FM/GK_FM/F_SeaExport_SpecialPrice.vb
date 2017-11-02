
Public Class F_SeaExport_SpecialPrice
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try

            If pVal.FormTypeEx = "UDO_FT_SeaExport_SP" Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "Item_0" Then
                            oForm = SBO_Application.Forms.ActiveForm
                            Load_SeaExportSpectailPrice(oForm)
                        End If


                    End If
                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "0_U_E" Then
                            oForm = SBO_Application.Forms.ActiveForm
                            oEdit = oForm.Items.Item("0_U_E").Specific
                            Dim BPCode As String = oEdit.String
                            If BPCode <> "" Then
                                oEdit = oForm.Items.Item("1_U_E").Specific
                                oEdit.String = BPName(BPCode, Ocompany)
                                oEdit = oForm.Items.Item("15_U_E").Specific
                                oEdit.String = BPMarkupPer(BPCode, Ocompany)
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            Functions.WriteLog("Class:F_SeaExport_SpecialPrice" + " Function:SBO_Application_ItemEvent" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Load_SeaExportSpectailPrice(ByVal oForm As SAPbouiCOM.Form)
        Try
            'Sea Export handling charge FCL
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.[U_VendorCode], T1.[U_VendorName], T1.[U_ChargeCode], T1.[U_ChargeDesc], T1.[U_CargoType], T1.[U_Quantity], T1.[U_Unit], T1.[U_Cost], T1.[U_SellingP], T1.[U_ContainerType], T1.[U_Remarks] FROM [dbo].[@AB_SEA_HC]  T0 , [dbo].[@AB_SEA_HCFCL]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Division] ='SE'")
            'Sea Export handling charge LCL
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_ChargeCode, T1.U_ChargeDesc, T1.U_Quantity, T1.U_Unit, T1.U_CargoType, T1.U_Cost, T1.U_SellingP, T1.U_Remarks FROM [dbo].[@AB_SEA_HC]  T0, [dbo].[@AB_SEA_HCLCL]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Division] ='SE'")
            If (oRecordSet1.RecordCount = 0 And oRecordSet.RecordCount = 0) Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            'oForm.Freeze(True)
            '---Laod Data to Matrix
            oMatrix = oForm.Items.Item("0_U_G").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("C_0_1")
            oColumn.Width = 0
            'oColumn.DataBind.SetBound(True, "@AB_SEAE_SPRICEFCL", "LineId")
            'oMatrix.Clear()
            For F = 1 To oRecordSet1.RecordCount
                If F = 1 Then
                    If oMatrix.RowCount <> 0 Then
                        oEdit = oMatrix.Columns.Item("C_0_7").Cells.Item(oMatrix.RowCount).Specific
                        If oEdit.String <> "" Then
                            oMatrix.AddRow()
                        End If
                    Else

                        oMatrix.AddRow()
                    End If
                Else
                    oMatrix.AddRow()
                End If
             
                oMatrix.ClearRowData(oMatrix.RowCount)
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading FCL - " & F & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix.Columns.Item("C_0_6").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value
                oEdit = oMatrix.Columns.Item("C_0_7").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value
                oEdit = oMatrix.Columns.Item("C_0_8").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix.Columns.Item("C_0_9").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeDesc").Value
                oEdit = oMatrix.Columns.Item("C_0_10").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix.Columns.Item("C_0_12").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                oEdit = oMatrix.Columns.Item("C_0_11").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix.Columns.Item("C_0_13").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
                Try
                    oCombo = oMatrix.Columns.Item("C_0_3").Cells.Item(oMatrix.RowCount).Specific
                    oCombo.Select(oRecordSet1.Fields.Item("U_ContainerType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception
                End Try
                oEdit = oMatrix.Columns.Item("C_0_17").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Remarks").Value
                oRecordSet1.MoveNext()
            Next
            For F = 1 To oMatrix.RowCount
                oEdit = oMatrix.Columns.Item("C_0_1").Cells.Item(F).Specific
                oEdit.Value = F
            Next
            '------LCL

            oMatrix = oForm.Items.Item("1_U_G").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("C_1_1")
            oColumn.Width = 0
            For F = 1 To oRecordSet.RecordCount

                If F = 1 Then
                    If oMatrix.RowCount <> 0 Then
                        oEdit = oMatrix.Columns.Item("C_1_6").Cells.Item(oMatrix.RowCount).Specific
                        If oEdit.String <> "" Then
                            oMatrix.AddRow()
                        End If
                    Else
                        oMatrix.AddRow()
                    End If
                Else

                    oMatrix.AddRow()
                End If
                oMatrix.ClearRowData(oMatrix.RowCount)
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading LCL - " & F & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix.Columns.Item("C_1_5").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_VendorCode").Value
                oEdit = oMatrix.Columns.Item("C_1_6").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_VendorName").Value
                oEdit = oMatrix.Columns.Item("C_1_7").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix.Columns.Item("C_1_8").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_ChargeDesc").Value
                oEdit = oMatrix.Columns.Item("C_1_9").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_Quantity").Value
                oEdit = oMatrix.Columns.Item("C_1_10").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_Unit").Value
                oEdit = oMatrix.Columns.Item("C_1_11").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_Cost").Value
                oEdit = oMatrix.Columns.Item("C_1_12").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_SellingP").Value
            
                oEdit = oMatrix.Columns.Item("C_1_20").Cells.Item(oMatrix.RowCount).Specific
                oEdit.String = oRecordSet.Fields.Item("U_Remarks").Value
                oRecordSet.MoveNext()
            Next
            For F = 1 To oMatrix.RowCount
                oEdit = oMatrix.Columns.Item("C_1_1").Cells.Item(F).Specific
                oEdit.String = F
            Next
            SBO_Application.StatusBar.SetText("Please Loading complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Functions.WriteLog("Class:F_SeaExport_SpecialPrice" + " Function:Load_SeaExportSpectailPrice" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
End Class
