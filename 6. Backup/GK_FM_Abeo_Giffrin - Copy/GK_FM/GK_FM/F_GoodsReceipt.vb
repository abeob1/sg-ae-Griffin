Imports System.Diagnostics.Process
Imports System.Threading
Imports System.IO

Public Class F_GoodsReceipt
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Public ShowFolderBrowserThread As Threading.Thread
    Dim strpath As String
    Dim FilePath As String
    Dim FileName As String
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Dim DefSupp As Boolean = False
    Dim DefPoNo As Boolean = False
    Public Sub GoodsReceipt_Bind(ByVal oform As SAPbouiCOM.Form)
        Try
            oform = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
            oform.PaneLevel = 1
            CFL_BP_Customer(oform, SBO_Application)
            CFL_BP_Supplier(oform, SBO_Application)
            CFL_Item(oform, SBO_Application)
            CFL_Item_Vessel(oform, SBO_Application)
            CFL_BP_WareHouse(oform, SBO_Application)
            'oEdit = oform.Items.Item("35").Specific
            'oEdit.ChooseFromListUID = "CFLBPV"
            'oEdit.ChooseFromListAlias = "CardCode"
            oEdit = oform.Items.Item("16").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")
            oEdit = oform.Items.Item("18").Specific
            oEdit.String = Format(Now.Date, "dd/MM/yy")

           

            oMatrix = oform.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            oMatrix.AddRow()
            oColumn = oColumns.Item("V_0")
            oColumn.ChooseFromListUID = "OITM"
            oColumn.ChooseFromListAlias = "ItemCode"
           

            oCombo = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            ComboLoad_Unit(oform, oCombo)
            'oEdit = oform.Items.Item("20").Specific
            'oEdit.ChooseFromListUID = "OITM11"
            'oEdit.ChooseFromListAlias = "ItemName"
            'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
            'oEdit.String = oMatrix.RowCount
           

            
            oEdit = oform.Items.Item("4").Specific
            oEdit.ChooseFromListUID = "CFLBPC"
            oEdit.ChooseFromListAlias = "CardCode"
            oMatrix = oform.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("V_15")
            oColumn.ChooseFromListUID = "CFLBPV"
            oColumn.ChooseFromListAlias = "CardCode"


            oColumn = oColumns.Item("V_12")
            oColumn.ChooseFromListUID = "CFLWSC"
            oColumn.ChooseFromListAlias = "WhsCode"


            oform.DataBrowser.BrowseBy = "12"
            DocNumber_GR()
            'oMatrix = oform.Items.Item("29").Specific
            oItem = oform.Items.Item("CopyTo")
            oItem.Enabled = False
            oItem = oform.Items.Item("12")
            oItem.Enabled = False
            oItem = oform.Items.Item("4")
            oItem.Enabled = True
            oItem = oform.Items.Item("6")
            oItem.Enabled = True
            oMatrix = oform.Items.Item("29").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("V_0")
            oColumn.Editable = True
            oColumn = oColumns.Item("V_1")
            oColumn.Editable = True
            oColumn = oColumns.Item("V_8")
            oColumn.Editable = True
            oColumn = oColumns.Item("V_12")
            oColumn.Editable = True
            oMatrix1 = oform.Items.Item("ATTMAT").Specific
            oColumns = oMatrix1.Columns
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub
    Public Sub Loadfile(ByVal FileName As String)
        Try
            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(FileName)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'SBO_Application.StatusBar.SetText("Path Name Is Empty!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
    'Public Sub GoodsIssue_Load()
    '    Try
    '        oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
    '        oEdit = oForm.Items.Item("4").Specific
    '        Dim CardCode As String = oEdit.String
    '        oEdit = oForm.Items.Item("6").Specific
    '        Dim CardName As String = oEdit.String
    '        oEdit = oForm.Items.Item("8").Specific
    '        Dim CardContactPerson As String = oEdit.String
    '        If CardCode <> "" Then

    '            oEdit = oForm.Items.Item("12").Specific
    '            If oEdit.String <> "" Then
    '                Dim DocNum As Integer = oEdit.String
    '                oEdit = oForm.Items.Item("14").Specific
    '                If oEdit.String = "Closed" Then
    '                    Exit Sub
    '                End If
    '                LoadFromXML("GoodsIssue.srf", SBO_Application)
    '                oForm = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
    '                Dim oGoodsIssue As New F_GoodsIssue(Ocompany, SBO_Application)
    '                oGoodsIssue.GoodsIssue_Bind(oForm)
    '                oEdit = oForm.Items.Item("4").Specific
    '                oEdit.String = CardCode
    '                oEdit = oForm.Items.Item("6").Specific
    '                oEdit.String = CardName
    '                oEdit = oForm.Items.Item("8").Specific
    '                oEdit.String = CardContactPerson

    '                Dim oRecordSet_GR As SAPbobsCOM.Recordset
    '                oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRecordSet_GR.DoQuery("SELECT T0.[U_NumAtCard], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T0.[U_VenCode], T0.[U_VenName], T0.[U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId] FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
    '                If oRecordSet_GR.RecordCount = 0 Then
    '                    SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                    Exit Sub
    '                End If
    '                oEdit = oForm.Items.Item("10").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(0).Value
    '                oEdit = oForm.Items.Item("20").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(1).Value
    '                oEdit = oForm.Items.Item("22").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(2).Value
    '                oEdit = oForm.Items.Item("24").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(3).Value
    '                oEdit = oForm.Items.Item("26").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(4).Value
    '                oEdit = oForm.Items.Item("33").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(5).Value
    '                oEdit = oForm.Items.Item("35").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(6).Value
    '                oEdit = oForm.Items.Item("37").Specific
    '                oEdit.String = oRecordSet_GR.Fields.Item(7).Value
    '                Try
    '                    oEdit = oForm.Items.Item("31").Specific
    '                    If oEdit.String = "" Then
    '                        oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
    '                    Else
    '                        oEdit.String = oEdit.String & "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
    '                    End If
    '                Catch ex As Exception
    '                End Try


    '                For i = 1 To oRecordSet_GR.RecordCount
    '                    If oMatrix.RowCount = 0 Then
    '                        oMatrix.AddRow()
    '                    End If
    '                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(9).Value
    '                    oEdit = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(19).Value
    '                    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(10).Value
    '                    Try
    '                        oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
    '                        oCombo.Select(oRecordSet_GR.Fields.Item(11).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
    '                    Catch ex As Exception

    '                    End Try

    '                    oEdit = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(12).Value
    '                    oEdit = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(13).Value
    '                    oEdit = oMatrix.Columns.Item("V_4").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(14).Value
    '                    oEdit = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(15).Value
    '                    oEdit = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(16).Value
    '                    oEdit = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(18).Value
    '                    oEdit = oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(17).Value
    '                    oEdit = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(20).Value
    '                    oEdit = oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific
    '                    oEdit.String = oRecordSet_GR.Fields.Item(21).Value
    '                    oMatrix.AddRow()
    '                    oRecordSet_GR.MoveNext()
    '                Next

    '            End If
    '        End If
    '    Catch ex As Exception
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
    Public Sub GoodsIssue_Load()
        Try
            oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
            oEdit = oForm.Items.Item("4").Specific
            Dim CardCode As String = oEdit.String
            oEdit = oForm.Items.Item("6").Specific
            Dim CardName As String = oEdit.String
            oEdit = oForm.Items.Item("8").Specific
            Dim CardContactPerson As String = oEdit.String
            If CardCode <> "" Then

                oEdit = oForm.Items.Item("12").Specific
                If oEdit.String <> "" Then
                    oEdit = oForm.Items.Item("12").Specific
                    Dim DocNum As Integer = oEdit.Value
                    oEdit = oForm.Items.Item("14").Specific
                    If oEdit.String = "Closed" Then
                        Exit Sub
                    End If
                    'oForm.Close()
                    LoadFromXML("GoodsIssue.srf", SBO_Application)
                    ' Dim oform As SAPbouiCOM.Forms
                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsIssue")

                    oEdit = oForm.Items.Item("GI4").Specific
                    oEdit.String = CardCode
                    oEdit = oForm.Items.Item("6").Specific
                    oEdit.String = CardName
                    oEdit = oForm.Items.Item("8").Specific
                    oEdit.String = CardContactPerson

                    Dim oRecordSet_GR As SAPbobsCOM.Recordset
                    oRecordSet_GR = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet_GR.DoQuery("SELECT T1.[U_NumAtCar], T0.[U_VesselNo], T0.[U_MAWBNo], T0.[U_POL], T0.[U_ANSRecNo], T0.[U_ShipTo], T1.[U_VenCode], T1.[U_VenName], T0.[U_Drivname], T1.[U_ItemCode], T1.[U_OpenQty], T1.[U_Unit], T1.[U_Weight], T1.[U_Length], T1.[U_Width], T1.[U_Height], T1.[U_Volume], T1.[U_BinLoc], T1.[U_Whsc], T1.[U_Decript], T1.[DocEntry], T1.[LineId] FROM [dbo].[@AIGR]  T0 , [dbo].[@AIGR1]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  T0.[DocEntry] ='" & DocNum & "' and  T1.[U_ItemCode] <>'' and  T1.[U_OpenQty]  >0")
                    If oRecordSet_GR.RecordCount = 0 Then
                        SBO_Application.StatusBar.SetText("No Data Found..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    'oEdit = oForm.Items.Item("10").Specific
                    'oEdit.String = oRecordSet_GR.Fields.Item(0).Value
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
                    'oEdit = oForm.Items.Item("35").Specific
                    'oEdit.String = oRecordSet_GR.Fields.Item(6).Value
                    'oEdit = oForm.Items.Item("37").Specific
                    'oEdit.String = oRecordSet_GR.Fields.Item(7).Value
                    Try
                        oEdit = oForm.Items.Item("31").Specific
                        If oEdit.String = "" Then
                            oEdit.String = "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
                        Else
                            oEdit.String = oEdit.String & "Based on Goods Recipt No " & oRecordSet_GR.Fields.Item(20).Value & ""
                        End If
                    Catch ex As Exception
                    End Try
                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
                    oMatrix = oForm.Items.Item("29").Specific
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
                    '-------------------
                    Try
                        oForm.DataSources.UserDataSources.Add("V_0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

                        'oMatrix1 = oForm.Items.Item("1000001").Specific
                        'oColumns = oMatrix1.Columns
                        'oMatrix1.AddRow()
                        'oColumn = oColumns.Item("V_0")
                        'oColumn.DataBind.SetBound(True, "", "oedit5")

                        oMatrix1 = oForm.Items.Item("1000001").Specific
                        oColumns = oMatrix1.Columns
                        oColumn = oColumns.Item("V_0")
                        oColumn.DataBind.SetBound(True, "", "V_0")
                        oItem = oForm.Items.Item("1000001")
                        oItem.Width = 150
                        oItem.Height = 30

                        oColumn.Width = 130
                        oMatrix1.AddRow()



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

                        DocNumber_GI()

                        oItem = oForm.Items.Item("12")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("GI4")
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

                        oForm.DataBrowser.BrowseBy = "12"
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                    '--------------------
                    ' oGoodsIssue.GoodsIssue_Bind(oform)
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub DocNumber_GR()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy-MM-dd")
            fdt = fdt.Substring(0, 8) & "01"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT CONVERT(VARCHAR(25),DATEADD(dd,-(DAY(DATEADD(mm,1,getdate()))),DATEADD(mm,1,getdate())),101)")
            tdt = oRecordSet1.Fields.Item(0).Value
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AIGR]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("26").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "GKWR" & Format(Now.Date, "yyyyMMdd") & "0001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "GKWR" & Format(Now.Date, "yyyyMMdd") & "000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "GKWR" & Format(Now.Date, "yyyyMMdd") & "00" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "GKWR" & Format(Now.Date, "yyyyMMdd") & "0" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "GKWR" & Format(Now.Date, "yyyyMMdd") & "" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = "GKWR" & Format(Now.Date, "yyyyMMdd") & DocNum
            End If

        Catch ex As Exception

        End Try
       
    End Sub

    Public Sub DocNumber_GI()
        Try
            Dim fdt As String = ""
            Dim tdt As String = ""
            fdt = Format(Now.Date, "yyyy-MM-dd")
            fdt = fdt.Substring(0, 8) & "01"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT CONVERT(VARCHAR(25),DATEADD(dd,-(DAY(DATEADD(mm,1,getdate()))),DATEADD(mm,1,getdate())),101)")
            tdt = oRecordSet1.Fields.Item(0).Value
            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet2.DoQuery("SELECT (count(*)+1) as CountNo FROM [dbo].[@AIGI]  T0 WHERE T0.[CreateDate]  between '" & fdt & "' and '" & tdt & "'")
            Dim DocNum As Integer = oRecordSet2.Fields.Item(0).Value
            oEdit = oForm.Items.Item("26").Specific
            Dim DocNumLen As Integer
            DocNumLen = DocNum.ToString.Length
            If DocNum = 0 Then
                oEdit.String = "GKWI" & Format(Now.Date, "yyyyMMdd") & "0001"
            ElseIf DocNumLen = 1 And DocNum <> 0 Then
                oEdit.String = "GKWI" & Format(Now.Date, "yyyyMMdd") & "000" & DocNum
            ElseIf DocNumLen = 2 Then
                oEdit.String = "GKWI" & Format(Now.Date, "yyyyMMdd") & "00" & DocNum
            ElseIf DocNumLen = 3 Then
                oEdit.String = "GKWI" & Format(Now.Date, "yyyyMMdd") & "0" & DocNum
            ElseIf DocNumLen = 4 Then
                oEdit.String = "GKWI" & Format(Now.Date, "yyyyMMdd") & "" & DocNum
            ElseIf DocNumLen = 5 Then
                oEdit.String = "GKWI" & Format(Now.Date, "yyyyMMdd") & DocNum
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            'If (pVal.FormType = 0 And pVal.ItemUID = "1" And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
            '    If SBO_Application.MessageBox("You Cannot Change this Document after you have add it.Continue?", 1, "Yes", "No") = 2 Then
            '        BubbleEvent = False
            '        Exit Sub
            '    End If
            'End If
            If (pVal.FormType = 0 And pVal.ItemUID = "1" And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                Try
                    Dim oOrderForm As SAPbouiCOM.Form
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(0, pVal.FormTypeCount)
                    oItem = oOrderForm.Items.Item("3")
                    If oItem.Visible = True Then
                        If oForm.UniqueID = "AI_FI_GoodsReceipt" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oMatrix = oForm.Items.Item("29").Specific
                            'If SBO_Application.MessageBox("You Cannot Change this Document after you have add it.Continue?", 1, "Yes", "No") = 2 Then
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If
                            DocNumber_GR()
                            oEdit = oForm.Items.Item("4").Specific
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
                            Dim i As Integer
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
                            ''Dim oIGN As SAPbobsCOM.Documents
                            ''oIGN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                            ''oIGN.DocDate = Now.Date
                            ''oIGN.TaxDate = Now.Date
                            ''oEdit = oForm.Items.Item("4").Specific
                            ''oIGN.CardCode = oEdit.String

                            ''oMatrix = oForm.Items.Item("29").Specific
                            ''For i = 1 To oMatrix.RowCount
                            ''    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific
                            ''    If oEdit.String <> "" Then
                            ''        oIGN.Lines.ItemCode = oEdit.String
                            ''        oEdit = oMatrix.Columns.Item("V_12").Cells.Item(i).Specific
                            ''        oIGN.Lines.WarehouseCode = oEdit.String
                            ''        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific
                            ''        oIGN.Lines.Quantity = oEdit.Value
                            ''        oIGN.Lines.Add()
                            ''    End If
                            ''Next
                            ''oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ''oRecordSet2.DoQuery("SELECT max(T0.[DocNum]) +1 FROM [dbo].[@AIGR]  T0")
                            ''oIGN.Comments = "Based on Goods Receipt No: " & oRecordSet2.Fields.Item(0).Value & ""
                            ''Dim RetCode As Integer = oIGN.Add()
                            ''Dim SerrorMsg As String = ""
                            ''Ocompany.GetLastError(RetCode, SerrorMsg)
                            ''If RetCode <> 0 Then
                            ''    SBO_Application.StatusBar.SetText(Ocompany.GetLastErrorDescription & Ocompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            ''    BubbleEvent = False
                            ''    Exit Sub
                            ''End If
                        End If
                    End If
                Catch ex As Exception
                    'SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'BubbleEvent = False
                    'Exit Sub
                End Try
            End If
            If pVal.FormUID = "AI_FI_GoodsReceipt" Then
                oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
                If pVal.ItemUID = "GRIT" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oForm.PaneLevel = 1
                ElseIf pVal.ItemUID = "GRATT" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oForm.PaneLevel = 2
                End If
                If ((pVal.FormUID = "AI_FI_GoodsReceipt" And pVal.ItemUID = "ATTMAT" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                    oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                    oColumns = oMatrix1.Columns
                    Dim i As Integer
                    For i = 1 To oMatrix1.RowCount
                        If oMatrix1.IsRowSelected(i) Then
                            oItem = oForm.Items.Item("38")
                            oItem.Enabled = True
                            oItem = oForm.Items.Item("39")
                            oItem.Enabled = True
                        End If
                    Next
                End If
                If ((pVal.FormUID = "AI_FI_GoodsReceipt" And pVal.ItemUID = "39" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                    oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                    oColumns = oMatrix1.Columns
                    Dim i As Integer
                    For i = 1 To oMatrix1.RowCount
                        If oMatrix1.IsRowSelected(i) Then
                            oMatrix1.DeleteRow(i)
                        End If
                    Next
                    oItem = oForm.Items.Item("38")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("39")
                    oItem.Enabled = False
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                End If
                If ((pVal.FormUID = "AI_FI_GoodsReceipt" And pVal.ItemUID = "38" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
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
                    oItem = oForm.Items.Item("38")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("39")
                    oItem.Enabled = False
                End If
                If ((pVal.FormUID = "AI_FI_GoodsReceipt" And pVal.ItemUID = "37" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.Before_Action = False)) Then
                    Try
                        oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")

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

                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oEdit = oForm.Items.Item("16").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oEdit = oForm.Items.Item("18").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oMatrix = oForm.Items.Item("29").Specific
                    oColumns = oMatrix.Columns
                    oMatrix.AddRow()
                    DocNumber_GR()
                End If
                If (pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                    Try
                        oMatrix = oForm.Items.Item("29").Specific
                        If SBO_Application.MessageBox("You Cannot Change this Document after you have add it.Continue?", 1, "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        DocNumber_GR()
                        oEdit = oForm.Items.Item("4").Specific
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
                        Dim i As Integer
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
                        oMatrix1 = oForm.Items.Item("ATTMAT").Specific
                        'Mithra

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
10:                             If System.IO.File.Exists(destPath & FileName & FileExten) Then
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
                        'Dim oIGN As SAPbobsCOM.Documents
                        'oIGN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                        'oIGN.DocDate = Now.Date
                        'oIGN.TaxDate = Now.Date
                        'oEdit = oForm.Items.Item("4").Specific
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
                        'oRecordSet2.DoQuery("SELECT max(T0.[DocNum]) +1 FROM [dbo].[@AIGR]  T0")
                        'oIGN.Comments = "Based on Goods Receipt No: " & oRecordSet2.Fields.Item(0).Value & ""
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
                End If
                If pVal.ItemUID = "CopyTo" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim trd As Threading.Thread
                    trd = New Threading.Thread(AddressOf GoodsIssue_Load)
                    trd.IsBackground = True
                    trd.SetApartmentState(ApartmentState.STA)
                    trd.Start()
                End If

                'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = True
                '    oItem = oForm.Items.Item("CopyTo")
                '    oItem.Enabled = False
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = False
                '    oItem = oForm.Items.Item("4")
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
                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    'oItem = oForm.Items.Item("4")
                    'oItem.Enabled = False
                    'oItem = oForm.Items.Item("6")
                    'oItem.Enabled = False
                    'oMatrix = oForm.Items.Item("29").Specific
                    'oColumns = oMatrix.Columns
                    'oColumn = oColumns.Item("V_0")
                    'oColumn.Editable = False
                    'oColumn = oColumns.Item("V_1")
                    'oColumn.Editable = False
                    'oColumn = oColumns.Item("V_8")
                    'oColumn.Editable = False
                    'oColumn = oColumns.Item("V_12")
                    'oColumn.Editable = False
                    'oItem = oForm.Items.Item("12")
                    'oItem.Enabled = False
                    oItem = oForm.Items.Item("CopyTo")
                    oItem.Enabled = False
                    oEdit = oForm.Items.Item("14").Specific
                    If oEdit.String = "Open" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oItem = oForm.Items.Item("CopyTo")
                        oItem.Enabled = True
                    End If
                End If

                'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                '    'oItem = oForm.Items.Item("1")
                '    'oItem.Enabled = True
                '    'DocNumber_GR()
                '    oItem = oForm.Items.Item("CopyTo")
                '    oItem.Enabled = False
                '    oItem = oForm.Items.Item("12")
                '    oItem.Enabled = False
                '    oItem = oForm.Items.Item("4")
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
                '    oEdit = oForm.Items.Item("26").Specific
                '    If oEdit.String = "" Then
                '        ' DocNumber_GR()
                '    End If


                'End If
                'If pVal.ItemUID = "29" And pVal.ColUID = "V_13" And pVal.Before_Action = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                '    If oMatrix.RowCount = 1 Then
                '        Exit Sub
                '    End If
                '    oEdit = oMatrix.Columns.Item("V_15").Cells.Item(1).Specific
                '    Dim Supp As String = oEdit.String
                '    oEdit = oMatrix.Columns.Item("V_14").Cells.Item(1).Specific
                '    Dim Supp1 As String = oEdit.String
                '    oEdit = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
                '    oEdit.String = Supp
                '    oEdit = oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific
                '    oEdit.String = Supp1
                '    oEdit = oMatrix.Columns.Item("V_13").Cells.Item(1).Specific
                '    Dim PO As String = oEdit.String
                '    oEdit = oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific
                '    oEdit.String = PO
                '    'oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                '    'If oEdit.String = "" Then

                '    'End If

                'End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_0" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    If oEdit.String <> "" Then
                        'If DefSupp = True Then


                        oMatrix.AddRow()
                        'End If

                        'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                        'oEdit.String = oMatrix.RowCount
                    End If
                End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_8" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                    If oEdit.String <> "" Then
                        Dim openQty As Integer = oEdit.String
                        oEdit = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                        oEdit.String = openQty
                        If Ocompany.UserName = "GK1" Then
                            oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                            oEdit.String = "SGJA"
                        ElseIf Ocompany.UserName = "GK2" Then
                            oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                            oEdit.String = "SGCH"
                        End If
                    End If
                End If
                If pVal.ItemUID = "29" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_8" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5") And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    Try
                        Dim l As Integer = 0
                        Dim b As Integer = 0
                        Dim w As Integer = 0
                        Dim vol As Double = 0.0
                        Dim qty As Integer = 0
                        oEdit = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                        If oEdit.String <> "" Then
                            qty = oEdit.Value
                        Else
                            Exit Try
                        End If
                        oEdit = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                        If oEdit.String <> "" Then
                            l = oEdit.Value
                        Else
                            Exit Try
                        End If

                        oEdit = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                        If oEdit.String <> "" Then
                            b = oEdit.Value
                        Else
                            Exit Try
                        End If

                        oEdit = oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                        If oEdit.String <> "" Then
                            w = oEdit.Value
                        Else
                            Exit Try
                        End If
                        oEdit = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                        oEdit.Value = ((l * b * w * qty) / 1000000)
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ItemUID = "4" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("4").Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oForm.Items.Item("6").Specific
                    If BPCode <> "" Then
                        oEdit.String = BPName(BPCode, Ocompany)
                    End If
                End If
                If pVal.ItemUID = "GR1000002" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Exit Sub
                    oEdit = oForm.Items.Item("GR1000002").Specific
                    Dim JobNo As String = oEdit.String
                    If JobNo <> "" Then
                        Load_From_JobOrder(JobNo, oForm)
                        'SELECT T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_LoadPortNC] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] like '%'
                        ' oEdit.String = BPName(BPCode, Ocompany)
                    End If
                End If
                If pVal.ItemUID = "4" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("4").Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oForm.Items.Item("8").Specific
                    If BPCode <> "" Then
                        oEdit.String = ContactPerson(BPCode, Ocompany)
                    End If
                End If
                If pVal.ItemUID = "4" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oForm.Items.Item("4").Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oForm.Items.Item("33").Specific
                    If BPCode <> "" Then
                        oEdit.String = BPAddress(BPCode, Ocompany)
                    End If
                End If
                If pVal.ItemUID = "29" And pVal.ColUID = "V_15" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    oEdit = oMatrix.Columns.Item("V_15").Cells.Item(pVal.Row).Specific
                    Dim BPCode As String = oEdit.String
                    oEdit = oMatrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific
                    If BPCode <> "" Then
                        oEdit.String = BPName(BPCode, Ocompany)
                    End If
                    'oEdit = oMatrix.Columns.Item("V_15").Cells.Item(1).Specific
                    'If pVal.Row = 1 Then
                    '    If SBO_Application.MessageBox("Do you want to update Same Vendor for all Rows?", 1, "Yes", "No") = 1 Then
                    '        DefSupp = True
                    '    End If
                    'End If
                End If
                'If pVal.ItemUID = "29" And pVal.ColUID = "V_13" And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                '    'oEdit = oMatrix.Columns.Item("V_15").Cells.Item(1).Specific
                '    If pVal.Row = 1 Then
                '        If SBO_Application.MessageBox("Do you want to update Same PO No. for all Rows?", 1, "Yes", "No") = 1 Then
                '            DefPoNo = True
                '        End If
                '    End If
                'End If
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
                    oEdit = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
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
                            If pVal.ItemUID = "4" Then
                                'oEdit = oForm.Items.Item("6").Specific
                                'oEdit.String = oDataTable.GetValue("CardName", 0)
                                'oEdit = oForm.Items.Item("8").Specific
                                'oEdit.String = ContactPerson(oDataTable.GetValue("CardCode", 0), Ocompany)
                                'oEdit = oForm.Items.Item("33").Specific
                                'oEdit.String = BPAddress(oDataTable.GetValue("CardCode", 0), Ocompany)
                                oEdit = oForm.Items.Item("4").Specific
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
                            If pVal.ItemUID = "29" And pVal.ColUID = "V_12" Then
                                'oEdit = oMatrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific
                                'oEdit.String = oDataTable.GetValue("CardName", 0)
                                oEdit = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("WhsCode", 0)
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

        End Try
    End Sub

  

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try

      
            If pVal.MenuUID = "OnlyOnRC1" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If oForm.UniqueID = "AI_FI_GoodsReceipt" Then
                            oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
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
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If oForm.UniqueID = "AI_FI_GoodsReceipt" Then
                            oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
                            oMatrix = oForm.Items.Item("29").Specific
                            oEdit = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix.AddRow()
                            End If


                            'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                            'oEdit.String = oMatrix.RowCount
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If
            If pVal.MenuUID = "1281" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    If oForm.UniqueID = "AI_FI_GoodsReceipt" Then
                        oItem = oForm.Items.Item("12")
                        oItem.Enabled = True
                        oItem = oForm.Items.Item("CopyTo")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("4")
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

            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    If oForm.UniqueID = "AI_FI_GoodsReceipt" Then
                        oItem = oForm.Items.Item("4")
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
                        oItem = oForm.Items.Item("CopyTo")
                        oItem.Enabled = False
                        oEdit = oForm.Items.Item("14").Specific
                        If oEdit.String = "Open" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oItem = oForm.Items.Item("CopyTo")
                            oItem.Enabled = True
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If
            If pVal.MenuUID = "1282" And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                Try
                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
                    oEdit = oForm.Items.Item("16").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oEdit = oForm.Items.Item("18").Specific
                    oEdit.String = Format(Now.Date, "dd/MM/yy")
                    oMatrix = oForm.Items.Item("29").Specific
                    oColumns = oMatrix.Columns
                    oMatrix.AddRow()
                    'oEdit = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific
                    'oEdit.String = oMatrix.RowCount
                    oItem = oForm.Items.Item("CopyTo")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("12")
                    oItem.Enabled = False
                    oItem = oForm.Items.Item("4")
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
                    oEdit = oForm.Items.Item("26").Specific
                    If oEdit.String = "" Then
                        DocNumber_GR()
                    End If
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
        Try

     
            If eventInfo.FormUID = "AI_FI_GoodsReceipt" Then
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
    Public Sub Load_From_JobOrder(ByVal JobNo As String, ByVal oForm As SAPbouiCOM.Form)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_VessName], T0.[U_OBL], T0.[U_HBL], T0.[U_LoadPortNC] FROM [dbo].[@AB_SEAI_JOB_H]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
            oEdit = oForm.Items.Item("4").Specific
            oEdit.String = oRecordSet.Fields.Item("U_CCode").Value
            oEdit = oForm.Items.Item("6").Specific
            oEdit.String = oRecordSet.Fields.Item("U_CName").Value
            oEdit = oForm.Items.Item("8").Specific
            oEdit.String = oRecordSet.Fields.Item("U_Atten").Value
            oEdit = oForm.Items.Item("20").Specific
            oEdit.String = oRecordSet.Fields.Item("U_VessName").Value
            oEdit = oForm.Items.Item("22").Specific
            oEdit.String = oRecordSet.Fields.Item("U_OBL").Value.ToString.Trim & " " & oRecordSet.Fields.Item("U_HBL").Value.ToString.Trim
            oEdit = oForm.Items.Item("24").Specific
            oEdit.String = oRecordSet.Fields.Item("U_LoadPortNC").Value
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
    Public Sub CopyDirectory(ByVal sourcePath As String, ByVal destPath As String)
        Try
            'SELECT AttachPath from OADP
            If Not Directory.Exists(destPath) Then
                ' Directory.CreateDirectory(destPath)
            End If

            For Each file__1 As String In Directory.GetFiles(Path.GetDirectoryName(sourcePath))
                Dim FileName As String = Path.GetFileNameWithoutExtension(file__1) & Now.ToString("ddMMyyyyhhmmssffff")
                Dim FileExten As String = Path.GetExtension(file__1)
                Dim dest As String = Path.Combine(destPath, FileName & FileExten)
                File.Copy(file__1, dest, False)
            Next

            For Each folder As String In Directory.GetDirectories(Path.GetDirectoryName(sourcePath))
                Dim dest As String = Path.Combine(destPath, Path.GetFileName(folder))
                CopyDirectory(folder, dest)
            Next

        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
    '    Public Function CopyDirectory(ByVal SrcPath As String, ByVal DestPath As String, Optional _
    'ByVal bQuiet As Boolean = False) As Boolean

    '        If Not System.IO.Directory.Exists(SrcPath) Then
    '            'Throw New System.IO.DirectoryNotFoundException("The directory " & SrcPath & " does not exists")
    '            SBO_Application.MessageBox("Path Not Found")
    '        End If
    '        If Not System.IO.Directory.Exists(DestPath) Then
    '            SBO_Application.MessageBox("Path Not Found")
    '        End If
    '        'Dim Files As String()
    '        'Files = System.IO.Directory.GetFileSystemEntries(SrcPath)

    '        Try
    '            My.Computer.FileSystem.CopyFile("C:\Path1\testFile.txt", "C:\Path1\testFile.txt", Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
    '        Catch ex As Exception
    '            SBO_Application.MessageBox("Path Not Found")
    '        End Try


    '    End Function
End Class
