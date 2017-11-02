Imports System.Diagnostics.Process
Imports System.Threading
Imports System.IO
Public Class F_SalesOrder
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub


    Public ShowFolderBrowserThread As Threading.Thread
    Dim strpath As String
    Dim FilePath As String
    Dim FileName As String
    Dim rowDelete As Integer
    Dim matrixUID As String
    Dim oform1 As SAPbouiCOM.Form
    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormTypeEx = 139 Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = True Then
                    oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    'Add Container details
                    '  Add_Container_details(oForm)
                    '  Add_Cargo_details(oForm)
                    
                End If
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = True Then
                    oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    ' Update Container details
                    ' Update_Container_details(oForm)
                    ' Update_Cargo_details(oForm)
                End If
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                    oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    'Get_Container_details(oForm)
                    ' Get_Cargo_details(oForm)
                    'Load Container details 
                End If

            End If


        Catch ex As Exception
            'SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try


            If (pVal.FormType = 133) Then
                If ((pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) And (pVal.Before_Action = True)) Then
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)

                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then

                        oItem = oOrderForm.Items.Item("2")
                        oNewItem = oOrderForm.Items.Add("Pr", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = oItem.Height
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Print Invoice"
                        oNewItem.LinkTo = "2"

                    End If
                End If
                If pVal.ItemUID = "Pr" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oEdit = oForm.Items.Item("8").Specific
                    oRecordSet.DoQuery("SELECT DocEntry from OINV where DocNum='" & oEdit.String & "'")
                    If oRecordSet.RecordCount > 0 Then
                        Dim DocNum As Integer = oRecordSet.Fields.Item(0).Value
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
                        cryRpt.Load(sPath & "\GK_FM\invoice.rpt")
                        'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

                        Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                        Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
                        Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
                        Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

                        crParameterDiscreteValue.Value = DocNum
                        crParameterFieldDefinitions = _
                    cryRpt.DataDefinition.ParameterFields
                        crParameterFieldDefinition = _
                    crParameterFieldDefinitions.Item("DocKey@")
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
                        RptFrm.Text = "Invoice Report"
                        RptFrm.TopMost = True

                        RptFrm.Activate()
                        RptFrm.ShowDialog()

                    Else
                        SBO_Application.StatusBar.SetText("No Data Found for Invoice to Print!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            End If

            If (pVal.FormType = 170 Or pVal.FormType = 426) Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True And pVal.InnerEvent = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("SELECT AttachPath from OADP")
                        Dim destPath As String = oRecordSet.Fields.Item("AttachPath").Value.ToString
                        If Not Directory.Exists(destPath) Then
                            SBO_Application.MessageBox("Destination Path Not Found!")
                            Exit Sub
                        End If
                        Dim sourcePath As String = ""
                        oEdit = oForm.Items.Item("e0").Specific
                        If oEdit.String <> "" Then
                            sourcePath = oEdit.String
                            If sourcePath.Contains(destPath) = False Then
                                Dim FileName As String = Path.GetFileNameWithoutExtension(sourcePath) '& Now.ToString("ddMMyyyyhhmmssffff")
                                Dim FileExten As String = Path.GetExtension(sourcePath)
                                Dim path1 As String = Path.GetDirectoryName(sourcePath)
                                'path1 = path1 & "\"
                                If path1 <> destPath Then
                                    Dim K As Integer = 1
10:                                 If System.IO.File.Exists(destPath & FileName & FileExten) Then
                                        ' MsgBox("THis Name Existsts")
                                        FileName = Path.GetFileNameWithoutExtension(sourcePath) & "_" & K
                                        K = K + 1
                                        GoTo 10
                                    End If
                                    Dim dest As String = Path.Combine(destPath, FileName & FileExten)
                                    File.Copy(sourcePath, dest, False)
                                    oEdit.String = dest
                                End If
                            End If
                        End If
                        oEdit = oForm.Items.Item("e1").Specific
                        If oEdit.String <> "" Then
                            sourcePath = oEdit.String
                            If sourcePath.Contains(destPath) = False Then
                                Dim FileName As String = Path.GetFileNameWithoutExtension(sourcePath) '& Now.ToString("ddMMyyyyhhmmssffff")
                                Dim FileExten As String = Path.GetExtension(sourcePath)
                                Dim path1 As String = Path.GetDirectoryName(sourcePath)
                                '  path1 = path1 & "\"
                                If path1 <> destPath Then
                                    Dim K As Integer = 1
11:                                 If System.IO.File.Exists(destPath & FileName & FileExten) Then
                                        ' MsgBox("THis Name Existsts")
                                        FileName = Path.GetFileNameWithoutExtension(sourcePath) & "_" & K
                                        K = K + 1
                                        GoTo 11
                                    End If
                                    Dim dest As String = Path.Combine(destPath, FileName & FileExten)
                                    File.Copy(sourcePath, dest, False)
                                    oEdit.String = dest
                                End If
                            End If
                        End If
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False And pVal.InnerEvent = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If pVal.ItemUID = "b0" Then
                        ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
                        If ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Unstarted Then
                            ShowFolderBrowserThread.SetApartmentState(Threading.ApartmentState.STA)
                            ShowFolderBrowserThread.Start()
                            ShowFolderBrowserThread.Join()
                        Else
                            ShowFolderBrowserThread.Abort()
                        End If
                        If FileName <> "" Then
                            oEdit = oForm.Items.Item("e0").Specific
                            oEdit.String = strpath
                        End If
                    ElseIf pVal.ItemUID = "b3" Then
                        ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
                        If ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Unstarted Then
                            ShowFolderBrowserThread.SetApartmentState(Threading.ApartmentState.STA)
                            ShowFolderBrowserThread.Start()
                            ShowFolderBrowserThread.Join()
                        Else
                            ShowFolderBrowserThread.Abort()
                        End If
                        If FileName <> "" Then
                            oEdit = oForm.Items.Item("e1").Specific
                            oEdit.String = strpath
                        End If
                    ElseIf pVal.ItemUID = "b1" Then
                        oEdit = oForm.Items.Item("e0").Specific
                        oEdit.String = ""
                    ElseIf pVal.ItemUID = "b4" Then
                        oEdit = oForm.Items.Item("e1").Specific
                        oEdit.String = ""

                    ElseIf pVal.ItemUID = "b2" Then
                        oEdit = oForm.Items.Item("e0").Specific
                        If oEdit.String <> "" Then
                            Loadfile(oEdit.String)
                        End If
                    ElseIf pVal.ItemUID = "b5" Then
                        oEdit = oForm.Items.Item("e1").Specific
                        If oEdit.String <> "" Then
                            Loadfile(oEdit.String)
                        End If

                    End If
                End If
                If ((pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) And (pVal.Before_Action = True)) Then
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then

                        oItem = oOrderForm.Items.Item("95")
                        oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                        oNewItem.Left = oItem.Left
                        oNewItem.Width = oItem.Width
                        oNewItem.Top = oItem.Top + 30
                        oNewItem.Height = 14
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        oEdit = oNewItem.Specific
                        If pVal.FormType = 170 Then
                            oEdit.DataBind.SetBound(True, "ORCT", "U_Att1")
                        Else
                            oEdit.DataBind.SetBound(True, "OVPM", "U_Att1")
                        End If
                        oNewItem.Enabled = False

                        oItem = oOrderForm.Items.Item("e0")
                        oNewItem = oOrderForm.Items.Add("b0", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = 65 'oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = 19 'oItem.Height
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Browse"
                        oNewItem.LinkTo = "e0"
                        oItem = oOrderForm.Items.Item("b0")
                        oNewItem = oOrderForm.Items.Add("b1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = 65 'oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = 19 'oItem.Height
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Delete"
                        oNewItem.LinkTo = "e0"
                        oItem = oOrderForm.Items.Item("b1")
                        oNewItem = oOrderForm.Items.Add("b2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = 65 'oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = 19 'oItem.Height
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Display"
                        oNewItem.LinkTo = "e0"
                        oItem = oOrderForm.Items.Item("96")
                        oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                        oNewItem.Left = oItem.Left
                        oNewItem.Width = oItem.Width '+ 85
                        oNewItem.Top = oItem.Top + 30
                        oNewItem.Height = 14
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        oNewItem.LinkTo = "e0"
                        oStatic = oNewItem.Specific
                        oStatic.Caption = "Atachment 1"
                        oItem = oOrderForm.Items.Item("e0")
                        oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                        oNewItem.Left = oItem.Left
                        oNewItem.Width = oItem.Width
                        oNewItem.Top = oItem.Top + 20
                        oNewItem.Height = 14
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        oEdit = oNewItem.Specific
                        If pVal.FormType = 170 Then
                            oEdit.DataBind.SetBound(True, "ORCT", "U_Att2")
                        Else
                            oEdit.DataBind.SetBound(True, "OVPM", "U_Att2")
                        End If
                        oNewItem.Enabled = False

                        oItem = oOrderForm.Items.Item("s0")
                        oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                        oNewItem.Left = oItem.Left
                        oNewItem.Width = oItem.Width '+ 85
                        oNewItem.Top = oItem.Top + 15
                        oNewItem.Height = 14
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        oNewItem.LinkTo = "e1"
                        oStatic = oNewItem.Specific
                        oStatic.Caption = "Atachment 2"

                        oItem = oOrderForm.Items.Item("e1")
                        oNewItem = oOrderForm.Items.Add("b3", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = 65 'oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = 19 'oItem.Height
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Browse"
                        oNewItem.LinkTo = "e1"

                        oItem = oOrderForm.Items.Item("b3")
                        oNewItem = oOrderForm.Items.Add("b4", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = 65 'oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = 19 'oItem.Height
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Delete"
                        oNewItem.LinkTo = "e1"
                        oItem = oOrderForm.Items.Item("b4")
                        oNewItem = oOrderForm.Items.Add("b5", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 10
                        oNewItem.Width = 65 'oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = 19 'oItem.Height
                        'oNewItem.FromPane = 12
                        'oNewItem.ToPane = 12
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Display"
                        oNewItem.LinkTo = "e1"
                    End If
                End If
            End If
            If pVal.FormType = 141 Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then
                    Try
                        oform1 = SBO_Application.Forms.GetFormByTypeAndCount(141, 1)
                        oItem = oform1.Items.Item("1")
                        oNewItem = oform1.Items.Item("ADDPV")
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = oItem.Height
                        oNewItem.Width = oItem.Width '+ 10
                        oNewItem.Left = oItem.Left
                    Catch ex As Exception

                    End Try
                End If
                If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
                    Try
                        oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-141, 1)
                    Catch ex As Exception
                        SBO_Application.ActivateMenuItem("6913")
                    End Try
                End If

                If ((pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) And (pVal.Before_Action = True)) Then
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then

                        If PVAddButt = "Yes" Then
                            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(141, 1)
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
                                PVAddButt = ""
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If

            End If
            '-----------Invoice
            If pVal.FormType = 133 Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then
                    Try
                        oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
                        oItem = oform1.Items.Item("1")
                        oNewItem = oform1.Items.Item("ADD")
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = oItem.Height
                        oNewItem.Width = oItem.Width '+ 10
                        oNewItem.Left = oItem.Left
                    Catch ex As Exception

                    End Try
                End If
                If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
                    Try
                        oform1 = SBO_Application.Forms.GetFormByTypeAndCount(-133, 1)
                    Catch ex As Exception
                        SBO_Application.ActivateMenuItem("6913")
                    End Try
                End If

                If ((pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) And (pVal.Before_Action = True)) Then
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then
                        If BIAddButt_Disable = "Yes" Then
                            'oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
                            'oform1.Title = "Billing Instruction"
                            'oItem = oform1.Items.Item("1")
                            'oItem.Enabled = False
                            'Try
                            '    Dim oNewItem As SAPbouiCOM.Item
                            '    oNewItem = oform1.Items.Add("ADD", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                            '    oItem = oform1.Items.Item("1")
                            '    oNewItem.Top = oItem.Top
                            '    oNewItem.Height = oItem.Height
                            '    oNewItem.Width = oItem.Width '+ 10
                            '    oNewItem.Left = oItem.Left
                            '    oButton = oNewItem.Specific
                            '    oButton.Caption = "Add BI"
                            '    BIAddButt = ""
                            'Catch ex As Exception
                            'End Try
                        End If
                        If BIAddButt = "Yes" Then
                            oform1 = SBO_Application.Forms.GetFormByTypeAndCount(133, 1)
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
                                BIAddButt = ""
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If
            End If
            If pVal.FormType = 139 Then
                oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                '*********************************************
                If ((pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) And (pVal.Before_Action = True)) Then
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)

                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then

                        oItem = oOrderForm.Items.Item("2")
                        oNewItem = oOrderForm.Items.Add("Pr1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + oItem.Width + 150
                        oNewItem.Width = oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = oItem.Height
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "Print SQ"
                        oNewItem.LinkTo = "2"

                    End If
                End If
                If pVal.ItemUID = "Pr1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oEdit = oForm.Items.Item("8").Specific
                    oRecordSet.DoQuery("SELECT DocEntry from ORDR where DocNum='" & oEdit.String & "'")
                    If oRecordSet.RecordCount > 0 Then
                        Dim DocNum As Integer = oRecordSet.Fields.Item(0).Value
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
                        cryRpt.Load(sPath & "\GK_FM\AB_SQ.rpt")
                        'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

                        Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                        Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
                        Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
                        Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

                        crParameterDiscreteValue.Value = DocNum
                        crParameterFieldDefinitions = _
                    cryRpt.DataDefinition.ParameterFields
                        crParameterFieldDefinition = _
                    crParameterFieldDefinitions.Item("DocKey@")
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
                        RptFrm.Text = "Sales Quoto"
                        RptFrm.TopMost = True

                        RptFrm.Activate()
                        RptFrm.ShowDialog()

                    Else
                        SBO_Application.StatusBar.SetText("No Data Found for SQ to Print!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
                '*********************************************
                If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    Try
                        'oEdit = oForm.Items.Item("e6").Specific
                        'oEdit.String = oEdit.String
                        'oItem = oForm.Items.Item("cce3")
                        'oItem.Enabled = False
                        'oItem = oForm.Items.Item("cce4")
                        'oItem.Enabled = False
                    Catch ex As Exception

                    End Try

                End If
                If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                    Try
                        oItem = oForm.Items.Item("e9")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("e10")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("e12")
                        oItem.Enabled = False
                        ooption = oForm.Items.Item("ce2").Specific
                        If ooption.Selected = False Then
                            oItem = oForm.Items.Item("cce3")
                            oItem.Enabled = False
                        End If
                    Catch ex As Exception

                    End Try

                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        If pVal.ItemUID = "e2" And pVal.InnerEvent = False Then
                            oCombo = oForm.Items.Item("e2").Specific
                            oEdit = oForm.Items.Item("e22").Specific
                            oEdit.String = oCombo.Selected.Description
                        ElseIf pVal.ItemUID = "e1" And pVal.InnerEvent = False Then
                            oCombo = oForm.Items.Item("e1").Specific
                            If oCombo.Selected.Value = "SE" Then
                                oEdit = oForm.Items.Item("e13").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("e14").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("ce13").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("ce14").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("e3").Specific
                                oEdit.String = "SG"
                                oEdit = oForm.Items.Item("e33").Specific
                                oEdit.String = "SINGAPORE"
                                oEdit = oForm.Items.Item("ce3").Specific
                                oEdit.String = "SIN"
                                oEdit = oForm.Items.Item("ce33").Specific
                                oEdit.String = "SINGAPORE"
                            ElseIf oCombo.Selected.Value = "AE" Then
                                oEdit = oForm.Items.Item("ce13").Specific
                                oEdit.String = ""
                                oEdit = oForm.Items.Item("ce14").Specific
                                oEdit.String = ""
                                oEdit = oForm.Items.Item("ce3").Specific
                                oEdit.String = "SIN"
                                oEdit = oForm.Items.Item("ce33").Specific
                                oEdit.String = "SINGAPORE"
                            ElseIf oCombo.Selected.Value = "LC" Then
                                oEdit = oForm.Items.Item("ce13").Specific
                                oEdit.String = "SIN"
                                oEdit = oForm.Items.Item("ce14").Specific
                                oEdit.String = "SINGAPORE"
                                oEdit = oForm.Items.Item("ce3").Specific
                                oEdit.String = "SIN"
                                oEdit = oForm.Items.Item("ce33").Specific
                                oEdit.String = "SINGAPORE"

                            ElseIf oCombo.Selected.Value = "AI" Then
                                oEdit = oForm.Items.Item("ce3").Specific
                                oEdit.String = ""
                                oEdit = oForm.Items.Item("ce33").Specific
                                oEdit.String = ""
                                oEdit = oForm.Items.Item("ce13").Specific
                                oEdit.String = "SIN"
                                oEdit = oForm.Items.Item("ce14").Specific
                                oEdit.String = "SINGAPORE"
                            ElseIf oCombo.Selected.Value = "SI" Then
                                oEdit = oForm.Items.Item("e3").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("e33").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("ce3").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("ce33").Specific
                                oEdit.String = String.Empty
                                oEdit = oForm.Items.Item("e13").Specific
                                oEdit.String = "SG"
                                oEdit = oForm.Items.Item("e14").Specific
                                oEdit.String = "SINGAPORE"
                                oEdit = oForm.Items.Item("ce13").Specific
                                oEdit.String = "SIN"
                                oEdit = oForm.Items.Item("ce14").Specific
                                oEdit.String = "SINGAPORE"
                            End If

                        End If
                    End If
                    'ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    '    If pVal.Before_Action = True And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then


                    '    End If
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
                            ' If (pVal.ItemUID = "e2" Or pVal.ItemUID = "ce2" Or pVal.ItemUID = "vce2" Or (pVal.ItemUID = "38" And pVal.ColUID = "U_AB_Vendor")) Then
                            If pVal.ItemUID = "e2" Then
                                Try
                                    oEdit = oForm.Items.Item("e2").Specific
                                    oEdit.String = oDataTable.GetValue("ItemCode", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "ve2" Then
                                Try
                                    oEdit = oForm.Items.Item("ve2").Specific
                                    oEdit.String = oDataTable.GetValue("ItemCode", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "vce3" Then
                                Try
                                    oEdit = oForm.Items.Item("vce3").Specific
                                    oEdit.String = oDataTable.GetValue("ItemName", 0)
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "38" And pVal.ColUID = "U_AB_Vendor" Then
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(pVal.Row).Specific
                                oEdit.String = oDataTable.GetValue("CardName", 0)
                            End If
                            'End If
                            'If pVal.ItemUID = "MATCargo" And pVal.ColUID = "SupCode" Then
                            '    oMatrix = oForm.Items.Item("MATCargo").Specific
                            '    oEdit = oMatrix.Columns.Item("SupName").Cells.Item(pVal.Row).Specific
                            '    oEdit.String = oDataTable.GetValue("CardName", 0)
                            '    oEdit = oMatrix.Columns.Item("SupCode").Cells.Item(pVal.Row).Specific
                            '    oEdit.String = oDataTable.GetValue("CardCode", 0)
                            'End If


                        End If
                    Catch ex As Exception
                    End Try


                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    If pVal.Before_Action = False And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.ItemUID = "MATCont" And pVal.ColUID = "ContNo" Then
                            oMatrix1 = oForm.Items.Item("MATCont").Specific
                            oEdit = oMatrix1.Columns.Item("ContNo").Cells.Item(oMatrix1.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix1.AddRow()
                                oMatrix1.ClearRowData(oMatrix1.RowCount)
                                oEdit = oMatrix1.Columns.Item("#").Cells.Item(oMatrix1.RowCount).Specific
                                oEdit.String = ""
                            End If
                        ElseIf pVal.ItemUID = "MATCargo" And pVal.ColUID = "SupCode" Then
                            oMatrix = oForm.Items.Item("MATCargo").Specific
                            oEdit = oMatrix.Columns.Item("SupCode").Cells.Item(pVal.Row).Specific
                            Dim Supp As String = oEdit.String
                            If oEdit.String <> "" Then
                                oEdit = oMatrix.Columns.Item("SupName").Cells.Item(pVal.Row).Specific
                                oEdit.String = BPName(Supp, Ocompany)
                            End If
                            oEdit = oMatrix.Columns.Item("SupCode").Cells.Item(oMatrix.RowCount).Specific
                            If oEdit.String <> "" Then
                                oMatrix.AddRow()
                                oMatrix.ClearRowData(oMatrix.RowCount)
                                oEdit = oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific
                                oEdit.String = ""
                            End If
                        ElseIf pVal.ItemUID = "e3" Then
                            oEdit = oForm.Items.Item("e3").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("e33").Specific
                                oEdit.String = Country_Code(ContCode, Ocompany)
                            End If
                        ElseIf pVal.ItemUID = "ce3" Then
                            oEdit = oForm.Items.Item("ce3").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("ce33").Specific
                                oEdit.String = City_Code(ContCode, Ocompany)
                            End If
                        ElseIf pVal.ItemUID = "e13" Then
                            oEdit = oForm.Items.Item("e13").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("e14").Specific
                                oEdit.String = Country_Code(ContCode, Ocompany)
                            End If
                        ElseIf pVal.ItemUID = "e2c" Then
                            oEdit = oForm.Items.Item("e2c").Specific
                            Dim Code As String = oEdit.String
                            If Code <> "" Then
                                oEdit = oForm.Items.Item("e22").Specific
                                oEdit.String = Carrier_Name(Code, Ocompany)
                            End If
                        ElseIf pVal.ItemUID = "ce13" Then
                            oEdit = oForm.Items.Item("ce13").Specific
                            Dim ContCode As String = oEdit.String
                            If ContCode <> "" Then
                                oEdit = oForm.Items.Item("ce14").Specific
                                oEdit.String = City_Code(ContCode, Ocompany)
                            End If
                        ElseIf pVal.BeforeAction = False And pVal.ItemUID = "MATCargo" And (pVal.ColUID = "Pkg" Or pVal.ColUID = "H" Or pVal.ColUID = "Wt") Then
                            Try
                                oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                                Dim l As Integer = 0
                                Dim w As Integer = 0
                                Dim h As Integer = 0
                                Dim wt As Double = 0
                                Dim vol As Double = 0
                                Dim m3 As Double = 0
                                oEdit = oMatrix.Columns.Item("L").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    l = oEdit.Value
                                Else
                                    Exit Sub
                                    l = 0
                                End If
                                oEdit = oMatrix.Columns.Item("W").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    w = oEdit.Value
                                Else
                                    w = 0
                                End If
                                oEdit = oMatrix.Columns.Item("H").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    h = oEdit.Value
                                Else
                                    h = 0
                                End If
                                Dim Division As String = ""
                                Try
                                    'nath

                                    oCombo = oForm.Items.Item("e1").Specific
                                    Division = oCombo.Selected.Value
                                    If Division = "SE" Or Division = "SI" Then
                                        m3 = ((l / 100) * (w / 100) * (h / 100))
                                    ElseIf Division = "AE" Or Division = "AI" Then
                                        m3 = ((l * w * h) / 6000)
                                    End If

                                Catch ex As Exception
                                    m3 = 0
                                End Try
                                oEdit = oMatrix.Columns.Item("m3").Cells.Item(pVal.Row).Specific
                                oEdit.String = m3
                                oEdit = oMatrix.Columns.Item("Wt").Cells.Item(pVal.Row).Specific
                                If oEdit.String <> "" Then
                                    wt = oEdit.Value
                                Else
                                    wt = 0
                                End If
                                If m3 > wt Then
                                    oEdit = oMatrix.Columns.Item("vol").Cells.Item(pVal.Row).Specific
                                    oEdit.String = m3
                                Else
                                    oEdit = oMatrix.Columns.Item("vol").Cells.Item(pVal.Row).Specific
                                    oEdit.String = wt
                                End If

                                oCombo = oForm.Items.Item("e1").Specific
                                Division = oCombo.Selected.Value
                                Dim TotWt As Double = 0
                                Dim TotPkg As Double = 0
                                If Division = "AI" Or Division = "AE" Then
                                    Dim i As Integer = 0
                                    For i = 1 To oMatrix.RowCount
                                        oEdit = oMatrix.Columns.Item("vol").Cells.Item(i).Specific
                                        If oEdit.String <> "" Then
                                            TotWt = TotWt + oEdit.Value
                                        End If
                                        oEdit = oMatrix.Columns.Item("Pkg").Cells.Item(i).Specific
                                        If oEdit.String <> "" Then
                                            TotPkg = TotPkg + oEdit.Value
                                        End If
                                    Next
                                End If
                                'oItem = oForm.Items.Item("cce3")
                                'oItem.Enabled = True
                                'oItem = oForm.Items.Item("cce4")
                                'oItem.Enabled = True
                                oEdit = oForm.Items.Item("cce4").Specific
                                oEdit.String = TotPkg
                                oEdit = oForm.Items.Item("cce3").Specific
                                oEdit.Value = TotWt
                                oEdit = oForm.Items.Item("e6").Specific
                                oEdit.String = oEdit.String
                                'oItem = oForm.Items.Item("cce3")
                                'oItem.Enabled = False
                                'oItem = oForm.Items.Item("cce4")
                                'oItem.Enabled = False
                            Catch ex As Exception

                            End Try
                        End If

                    End If
                End If
                'ce1
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If pVal.ItemUID = "ce1" Then
                        ooption = oForm.Items.Item("ce1").Specific
                        If ooption.Selected = True Then
                            oCombo = oForm.Items.Item("cce3").Specific
                            oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            oEdit = oForm.Items.Item("16").Specific
                            oEdit.String = oEdit.String
                            oItem = oForm.Items.Item("cce3")
                            oItem.Enabled = False
                        End If
                    ElseIf pVal.ItemUID = "1" Then
                        Try
                            oMatrix = oForm.Items.Item("MATCargo").Specific
                            oMatrix.AddRow()
                            oMatrix1 = oForm.Items.Item("MATCont").Specific
                            oMatrix1.AddRow()

                        Catch ex As Exception

                        End Try
                    ElseIf pVal.ItemUID = "ce2" Then
                        ooption = oForm.Items.Item("ce2").Specific
                        If ooption.Selected = True Then
                            oItem = oForm.Items.Item("cce3")
                            oItem.Enabled = True
                        End If
                    End If
                    If pVal.ItemUID = "Charge1" Then
                        oForm = SBO_Application.Forms.GetFormByTypeAndCount(139, pVal.FormTypeCount)
                        oForm.PaneLevel = 1
                        oFolder = oForm.Items.Item("112").Specific
                        oFolder.Select()
                        oEdit = oForm.Items.Item("4").Specific
                        Dim customercode As String = oEdit.String
                        If oEdit.String = "" Then
                            SBO_Application.StatusBar.SetText("Select Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        Dim Division As String = ""
                        Try
                            oCombo = oForm.Items.Item("e1").Specific
                            Division = oCombo.Selected.Value
                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText("Select Division", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End Try

                        '------------------Price Load for Air
                        If Division = "AI" Then
                            LoadHandingCharge_AirImport(oForm, "AI")
                        ElseIf Division = "LC" Then
                            oEdit = oForm.Items.Item("4").Specific
                            Dim BPCode As String = oEdit.String
                            Dim Str As String = "SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_ItemCode, T1.U_ItemDesc, T1.U_Quantity, T1.U_Cost, T1.U_SPrice, T1.U_MarkUp, T1.U_MarkedUpPrice, T1.U_Remarks FROM [dbo].[@AB_LOCAL_SPRICEH]  T0 , [dbo].[@AB_LOCAL_SPRICED]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='LC' and  T1.[Code] ='" & BPCode & "'"
                            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(Str)
                            If oRecordSet.RecordCount <> 0 Then
                                LoadHandingCharge_LocalCharges_SpecialPrice(oForm, "LC", BPCode)
                            Else
                                LoadHandingCharge_LocalCharges(oForm, "LC")
                            End If
                        ElseIf Division = "AE" Then
                            'chk Special Price

                            'Dim DestCountry As String = ""
                            Dim DestCity As String = ""
                            Dim ServiceLevel As String = ""
                            Dim Carrier As String = ""
                            Dim Cargo As String = ""
                            Dim Weight As Double = 0
                            Dim MinAmt As Double = 0
                            'oEdit = oForm.Items.Item("e13").Specific
                            'DestCountry = oEdit.String
                            oEdit = oForm.Items.Item("ce13").Specific
                            DestCity = oEdit.String
                            oEdit = oForm.Items.Item("cce3").Specific
                            Weight = oEdit.Value
                            Try
                                oCombo = oForm.Items.Item("ce1").Specific
                                Cargo = oCombo.Selected.Value
                            Catch ex As Exception
                            End Try
                            Try
                                oEdit = oForm.Items.Item("e2c").Specific
                                Carrier = oEdit.String
                            Catch ex As Exception
                            End Try
                            Try
                                oCombo = oForm.Items.Item("e11").Specific
                                ServiceLevel = oCombo.Selected.Value
                            Catch ex As Exception
                            End Try
                            oEdit = oForm.Items.Item("4").Specific
                            Dim BPCode As String = oEdit.String
                            Dim Str As String = "SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_MinBase], T1.[U_Neg45Base], T1.[U_45Base], T1.[U_100Base], T1.[U_300base], T1.[U_500Base], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500] FROM [dbo].[@AB_AIRSPECIAL_H]  T0 , [dbo].[@AB_AIRSPECIAL_D]  T1 WHERE T1.[Code] = T0.[Code]  and T0.[U_Division] ='AE' and    T1.[U_Code] ='" & DestCity & "' and  T1.[U_Carrier] ='" & Carrier & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "' and  T1.[U_CargoType] ='" & Cargo & "' and T0.Code='" & BPCode & "'"
                            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(Str)
                            If oRecordSet.RecordCount <> 0 Then
                                LoadHandingCharge_AirExport_SPECIAL(oForm, "AE")
                                ' LoadHandingCharge_AirImport_only(oForm, "AE")
                            Else
                                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                ' LoadHandingCharge_AirExport(oForm, "AE")
                                ' LoadHandingCharge_AirImport_only(oForm, "AE")
                            End If

                            '------------------Price Load for Sea
                        ElseIf Division = "SI" Then
                            Dim LCLFCL As Boolean = False
                            ooption = oForm.Items.Item("ce1").Specific
                            If ooption.Selected = True Then
                                LCLFCL = True
                                LoadHandingCharge_LCL_SI(oForm, oCombo.Selected.Value)
                            End If
                            ooption = oForm.Items.Item("ce2").Specific
                            If ooption.Selected = True Then
                                LCLFCL = True
                                LoadHandingCharge_FCL_SI(oForm, oCombo.Selected.Value)
                            End If
                            If LCLFCL = False Then
                                SBO_Application.StatusBar.SetText("Select Cargo Type FCL/LCL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                            '----------price List for 
                        ElseIf Division = "SE" Then
                            ooption = oForm.Items.Item("ce1").Specific

                            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM [dbo].[@AB_SEAE_SPRICE]  T0 WHERE T0.[Code] ='" & customercode & "'")
                            If oRecordSet.RecordCount <> 0 Then
                                'Special price
                                Dim LCLFCL As Boolean = False
                                ooption = oForm.Items.Item("ce1").Specific
                                If ooption.Selected = True Then
                                    LCLFCL = True
                                    LoadHandingCharge_LCL_SE_SpecialPrice(oForm, "SE", customercode)
                                End If
                                ooption = oForm.Items.Item("ce2").Specific
                                If ooption.Selected = True Then
                                    LCLFCL = True
                                    LoadHandingCharge_FCL_SE_SpecialPrice(oForm, "SE", customercode)
                                End If
                                If LCLFCL = False Then
                                    SBO_Application.StatusBar.SetText("Select Cargo Type FCL/LCL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                            Else
                                'Not special price
                                Dim LCLFCL As Boolean = False
                                ooption = oForm.Items.Item("ce1").Specific
                                If ooption.Selected = True Then
                                    LCLFCL = True
                                    LoadHandingCharge_LCL_SE_SeaPrice(oForm, "SE")
                                End If
                                ooption = oForm.Items.Item("ce2").Specific
                                If ooption.Selected = True Then
                                    LCLFCL = True
                                    LoadHandingCharge_FCL_SE_SeaPrice(oForm, "SE")
                                End If
                                If LCLFCL = False Then
                                    SBO_Application.StatusBar.SetText("Select Cargo Type FCL/LCL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                            End If
                        End If

                    End If
                End If

                If pVal.ItemUID = "General" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oForm.PaneLevel = 12
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        Try
                            oItem = oForm.Items.Item("1")
                            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Catch ex As Exception
                        End Try
                    End If
                End If
                If pVal.ItemUID = "Terms" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then
                    oForm.PaneLevel = 18

                ElseIf pVal.ItemUID = "Cargo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oForm.PaneLevel = 13

                ElseIf pVal.ItemUID = "Cont" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oForm.PaneLevel = 15
                End If

                If pVal.ItemUID = "e2" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.Before_Action = False And pVal.InnerEvent = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oEdit = oForm.Items.Item("e2").Specific
                    Dim ItemCode As String = oEdit.String
                    If ItemCode <> "" Then
                        oEdit = oForm.Items.Item("e22").Specific
                        oEdit.String = ItemName(ItemCode, Ocompany)
                    End If
                End If
                If pVal.ItemUID = "ve2" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.Before_Action = False And pVal.InnerEvent = False Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    oEdit = oForm.Items.Item("ve2").Specific
                    Dim ItemCode As String = oEdit.String
                    If ItemCode <> "" Then
                        oEdit = oForm.Items.Item("ve22").Specific
                        oEdit.String = ItemName(ItemCode, Ocompany)
                    End If
                End If
                If ((pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) And (pVal.Before_Action = True)) Then
                    oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then
                        '---------------TERMS
                        Try
                            oNewItem = oOrderForm.Items.Add("Terms", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                            oItem = oOrderForm.Items.Item("1320002137")
                            oNewItem.Top = oItem.Top
                            oNewItem.Height = oItem.Height
                            oNewItem.Width = oItem.Width
                            oNewItem.Left = oItem.Left + oItem.Width + 5
                            oFolderItem = oNewItem.Specific
                            oFolderItem.Caption = "Terms & Conditions"
                            oFolderItem.GroupWith("1320002137")
                            oFolderItem.Pane = 18
                            'oFolderItem.Select()
                            '===============================================
                            oItem = oOrderForm.Items.Item("38")
                            oNewItem = oOrderForm.Items.Add("TR", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                            oNewItem.Left = oItem.Left
                            oNewItem.Width = oItem.Width
                            oNewItem.Top = oItem.Top
                            oNewItem.Height = oItem.Height
                            oNewItem.FromPane = 18
                            oNewItem.ToPane = 18
                            oEdit = oNewItem.Specific
                            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Remarks")
                            oEdit.String = Terms(Ocompany)
                        Catch ex As Exception
                        End Try
                        oOrderForm.DataSources.DataTables.Add("OCARGO")
                        oOrderForm.DataSources.DataTables.Add("OCARGO1")
                        oItem = oOrderForm.Items.Item("114")
                        oItem.Visible = False
                        oItem = oOrderForm.Items.Item("138")
                        oItem.Visible = False
                        oEdit = oOrderForm.Items.Item("12").Specific
                        oEdit.String = Format(Now.Date, "dd/MM/yy")
                        oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        If Type = "Sea" Then
                            SystemFormManuplation_Sea(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)
                            Salese_Order_ContMatrix(oOrderForm)
                            Try
                                CFL_Item_Vessel(oOrderForm, SBO_Application)
                                oEdit = oForm.Items.Item("e2").Specific
                                oEdit.ChooseFromListUID = "OITM11"
                                oEdit.ChooseFromListAlias = "ItemCode"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                            Catch ex As Exception

                            End Try
                            Try

                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception

                            End Try
                            Type = ""
                        ElseIf Type = "Air" Then
                            'vce3
                            SystemFormManuplation_Air(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)
                            Try

                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"

                                CFL_Item_Vessel(oOrderForm, SBO_Application)
                                oEdit = oForm.Items.Item("vce3").Specific
                                oEdit.ChooseFromListUID = "OITM11"
                                oEdit.ChooseFromListAlias = "ItemName"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"



                            Catch ex As Exception
                            End Try
                            Type = ""
                        ElseIf Type = "Local" Then
                            SystemFormManuplation_Local(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)

                            Try
                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"


                            Catch ex As Exception

                            End Try

                            Type = ""
                        ElseIf Type = "International" And SubType = "Local" Then
                            SystemFormManuplation_Local_IN(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)

                            Try
                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"


                            Catch ex As Exception

                            End Try
                            Type = ""
                            SubType = ""
                        ElseIf Type = "International" And SubType = "Air" Then
                            SystemFormManuplation_Air_IN(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)

                            Try
                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception

                            End Try

                            Type = ""
                            SubType = ""
                        ElseIf Type = "International" And SubType = "Sea" Then
                            SystemFormManuplation_Sea_IN(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)
                            Salese_Order_ContMatrix(oOrderForm)
                            Try
                                CFL_Item_Vessel(oOrderForm, SBO_Application)
                                oEdit = oForm.Items.Item("e2").Specific
                                oEdit.ChooseFromListUID = "OITM11"
                                oEdit.ChooseFromListAlias = "ItemCode"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                            Catch ex As Exception

                            End Try
                            Try
                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception

                            End Try
                            Type = ""
                            SubType = ""
                        ElseIf Type = "Project" And SubType = "Air" Then
                            SystemFormManuplation_Air_PR(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)

                            Try
                                CFL_Item_Vessel(oOrderForm, SBO_Application)
                                oEdit = oForm.Items.Item("ve2").Specific
                                oEdit.ChooseFromListUID = "OITM11"
                                oEdit.ChooseFromListAlias = "ItemCode"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception
                            End Try

                            Type = ""
                            SubType = ""
                        ElseIf Type = "Project" And SubType = "Sea" Then
                            SystemFormManuplation_Sea_PR(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)
                            Salese_Order_ContMatrix(oOrderForm)
                            Try
                                CFL_Item_Vessel(oOrderForm, SBO_Application)
                                oEdit = oForm.Items.Item("e2").Specific
                                oEdit.ChooseFromListUID = "OITM11"
                                oEdit.ChooseFromListAlias = "ItemCode"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                            Catch ex As Exception

                            End Try
                            Try
                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception

                            End Try
                            Type = ""
                            SubType = ""
                        ElseIf Type = "Project" And SubType = "Local" Then
                            SystemFormManuplation_Local_PR(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)

                            Try
                                CFL_Item_Vessel(oOrderForm, SBO_Application)
                                oEdit = oForm.Items.Item("ve2").Specific
                                oEdit.ChooseFromListUID = "OITM11"
                                oEdit.ChooseFromListAlias = "ItemCode"

                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception

                            End Try
                            Type = ""
                            SubType = ""
                        ElseIf Type = "International1234" Then
                            SystemFormManuplation_International(oForm, oOrderForm)
                            Salese_Order_CargoMatrix(oOrderForm)

                            Try
                                CFL_BP_Supplier2(oForm, SBO_Application)
                                oMatrix = oForm.Items.Item("MATCargo").Specific
                                oColumns = oMatrix.Columns
                                oColumn = oColumns.Item("SupCode")
                                oColumn.ChooseFromListUID = "CFLBPV1"
                                oColumn.ChooseFromListAlias = "CardCode"

                                CFL_BP_Supplier(oForm, SBO_Application)
                                oMatrix3 = oForm.Items.Item("38").Specific
                                oColumns = oMatrix3.Columns
                                oColumn = oColumns.Item("U_AB_Vendor")
                                oColumn.ChooseFromListUID = "CFLBPV"
                                oColumn.ChooseFromListAlias = "CardName"
                            Catch ex As Exception

                            End Try
                            ' Salese_Order_ContMatrix(oOrderForm)
                            'Try
                            '    CFL_Item_Vessel(oOrderForm, SBO_Application)
                            '    oEdit = oForm.Items.Item("e2").Specific
                            '    oEdit.ChooseFromListUID = "OITM11"
                            '    oEdit.ChooseFromListAlias = "ItemCode"

                            'Catch ex As Exception

                            'End Try
                            Type = ""
                        End If
                        oOrderForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized


                    End If
                End If
            End If
        Catch ex As Exception
            If ex.Message <> "Form - Invalid Form" Then
                ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End Try
    End Sub
#Region "System Form Manupation"
    Public Sub SystemFormManuplation_Local_PR(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try
            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
           
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")
            oCombo.ValidValues.Add("PR", "Project")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            '===============================================
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            oCombo.ValidValues.Add("LC", "Local")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'oCombo.ValidValues.Add("SI", "Sea Export")
            ' ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
           

            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width '+ 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False

            '++++++++++++++++++
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("ve2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessCode")
            oItem = oOrderForm.Items.Item("ve2")
            oNewItem = oOrderForm.Items.Add("ve22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessName")

            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("vs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ve2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Vessel"
            '===============================================
            oItem = oOrderForm.Items.Item("ve2")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100 'oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("vs2")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin"
            '===============================================

            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination"
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100 'oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")



            '===============================================
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs13")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 30
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 50
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 50
            oItem = oOrderForm.Items.Item("e6")
            oNewItem.Left = oItem.Left ' + oItem.Width
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("cs33")
            oNewItem = oOrderForm.Items.Add("cs34", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs34")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Air_PR(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
           
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")
            oCombo.ValidValues.Add("PR", "Project")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            oCombo.ValidValues.Add("AI", "Air Import")
            oCombo.ValidValues.Add("AE", "Air Export")

            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
            '  Exit Sub
            '===============================================
            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("ve2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessCode")
            oItem = oOrderForm.Items.Item("ve2")
            oNewItem = oOrderForm.Items.Add("ve22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessName")

            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ve2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Vessel"

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("ve2")
            oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 0 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transfer To"
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = 70 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ComboLoad_TransferTo(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oItem.LinkTo = "e9"
            oItem = oOrderForm.Items.Item("e9")
            oItem.Enabled = False
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 10 + oItem.Width
            oNewItem.Width = oItem.Width + 20 '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transferred Status"
            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Width + oItem.Left
            oNewItem.Width = 60 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            oItem = oOrderForm.Items.Item("s10")
            oItem.LinkTo = "e10"
            oItem = oOrderForm.Items.Item("e10")
            oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Service Level"
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_SerLevel")
            ' ComboLoad_ShipType(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("e10")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False

            '' ''==============
            '' ''ComboLoad_Division(oForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            ' ''oNewItem.Width = oItem.Width + 17
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Destination Country"
            ' ''oItem = oOrderForm.Items.Item("e11")
            ' ''oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            '' ''GOPi
            '' ''ComboLoad_City(oOrderForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s13")
            ' ''oItem.LinkTo = "e13"
            '' ''  Exit Sub
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '' ''+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width + 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination Airport"
            oItem = oOrderForm.Items.Item("e11")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            '===============================================
            oItem = oOrderForm.Items.Item("ve2")
            oNewItem = oOrderForm.Items.Add("e2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CarricerC")
            ComboLoad_Carrier(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_CarrierN")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s2")
            oNewItem = oOrderForm.Items.Add("ss2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Carrier"
            '  Exit Sub

            ' '' ''===============================================
            '' ''oItem = oOrderForm.Items.Item("e2")
            '' ''oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = 80 'oItem.Width - 10
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            ' '' ''ComboLoad_City(oOrderForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("e3")
            '' ''oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left + oItem.Width + 2
            '' ''oNewItem.Width = 100
            '' ''oNewItem.Top = oItem.Top
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            ' '' ''ComboLoad_Division(oForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("s2")
            '' ''oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = oItem.Width '+ 85
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oNewItem.LinkTo = "e3"
            '' ''oStatic = oNewItem.Specific
            '' ''oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++
            '===============================================
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e22")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left '+ oItem.Width + 2
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("ss2")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin Airport"
            '===============================================
            'oItem = oOrderForm.Items.Item("ce3")
            'oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 300 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            'oItem = oOrderForm.Items.Item("cs3")
            'oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e3"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 15
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e6")
            oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15 '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15 '+ 14
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e7"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transit Time"


            oItem = oOrderForm.Items.Item("e7")
            oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            oItem = oOrderForm.Items.Item("s7")
            oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Frequency"
            oOrderForm.PaneLevel = 12

            oItem = oOrderForm.Items.Item("s8")
            oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 40
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Cargo Type"
            oOrderForm.PaneLevel = 12


            oItem = oOrderForm.Items.Item("e8")
            oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 40
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CaType")
            ComboLoad_CargoType(oForm, oCombo)
            'oItem = oOrderForm.Items.Item("ce1")
            'oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            'oNewItem.Left = oItem.Left + oItem.Width + 10
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top ' + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'ooption = oNewItem.Specific
            'ooption.Caption = "FCL"
            'ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            'ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("cs1")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Sea_PR(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
           
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")

            oCombo.ValidValues.Add("PR", "Project")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            oCombo.ValidValues.Add("SI", "Sea Import")
            oCombo.ValidValues.Add("SE", "Sea Export")
            'ComboLoad_Division(oForm, oCombo)

            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
            '  Exit Sub
            '===============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 0 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transfer To"
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = 70 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ComboLoad_TransferTo(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oItem.LinkTo = "e9"
            oItem = oOrderForm.Items.Item("e9")
            oItem.Enabled = False
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 10 + oItem.Width
            oNewItem.Width = oItem.Width + 20 '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transferred Status"
            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Width + oItem.Left
            oNewItem.Width = 60 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            oItem = oOrderForm.Items.Item("s10")
            oItem.LinkTo = "e10"
            oItem = oOrderForm.Items.Item("e10")
            oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Shipping Type"
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_ShipType")
            ComboLoad_ShipType(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("e10")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False
            '==============
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width + 17
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination Country"
            oItem = oOrderForm.Items.Item("e11")
            oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s13")
            oItem.LinkTo = "e13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s13")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination City"
            oItem = oOrderForm.Items.Item("e13")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e14")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("e2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessCode")
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessName")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Vessel"
            '  Exit Sub

            '===============================================
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e3")
            oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s2")
            oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++
            '===============================================
            oItem = oOrderForm.Items.Item("e3")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s3")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin City"
            '===============================================
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            '  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e6")
            oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15 '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15 '+ 14
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e7"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transit Time"


            oItem = oOrderForm.Items.Item("e7")
            oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            oItem = oOrderForm.Items.Item("s7")
            oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Frequency"
            oOrderForm.PaneLevel = 12

            oItem = oOrderForm.Items.Item("s8")
            oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Cargo Type"
            oOrderForm.PaneLevel = 12


            oItem = oOrderForm.Items.Item("e8")
            oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 40 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "LCL"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType")

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oNewItem.Width = 40 'oItem.Width - 10
            oNewItem.Top = oItem.Top ' + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "FCL"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("ce2")
            oNewItem = oOrderForm.Items.Add("bce3", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oNewItem.Width = 75 'oItem.Width - 10
            oNewItem.Top = oItem.Top ' + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "BulkLoad"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType2")
            ooption.GroupWith("ce1")


            oItem = oOrderForm.Items.Item("cs1")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Container Type"

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_ContType")
            ComboLoad_ContainerType(oForm, oCombo)
            oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("wcs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Packages"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            oNewItem = oOrderForm.Items.Item("wcs3")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("wcs3")
            oNewItem = oOrderForm.Items.Add("wcs4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            oNewItem = oOrderForm.Items.Item("wcs4")
            oNewItem.LinkTo = "cce5"

            oItem = oOrderForm.Items.Item("cce5")
            oNewItem = oOrderForm.Items.Add("wcs5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"
            ' oNewItem = oOrderForm.Items.Item("wcs5")
            oNewItem.LinkTo = "cce5"

            oItem = oOrderForm.Items.Item("wcs5")
            oNewItem = oOrderForm.Items.Add("cce6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + +oItem.Width + 10
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            oNewItem = oOrderForm.Items.Item("wcs5")
            oNewItem.LinkTo = "cce6"


        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Sea_IN(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")

            oCombo.ValidValues.Add("IN", "International")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            oCombo.ValidValues.Add("SI", "Sea Import")
            oCombo.ValidValues.Add("SE", "Sea Export")
            'ComboLoad_Division(oForm, oCombo)

            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
            '  Exit Sub
            '===============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 0 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transfer To"
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = 70 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ComboLoad_TransferTo(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oItem.LinkTo = "e9"
            oItem = oOrderForm.Items.Item("e9")
            oItem.Enabled = False
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 10 + oItem.Width
            oNewItem.Width = oItem.Width + 20 '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transferred Status"
            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Width + oItem.Left
            oNewItem.Width = 60 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            oItem = oOrderForm.Items.Item("s10")
            oItem.LinkTo = "e10"
            oItem = oOrderForm.Items.Item("e10")
            oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Shipping Type"
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_ShipType")
            ComboLoad_ShipType(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("e10")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False
            '==============
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width + 17
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination Country"
            oItem = oOrderForm.Items.Item("e11")
            oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s13")
            oItem.LinkTo = "e13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s13")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination City"
            oItem = oOrderForm.Items.Item("e13")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e14")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("e2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessCode")
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessName")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Vessel"
            '  Exit Sub

            '===============================================
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e3")
            oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s2")
            oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++
            '===============================================
            oItem = oOrderForm.Items.Item("e3")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s3")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin City"
            '===============================================
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            '  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e6")
            oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15 '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15 '+ 14
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e7"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transit Time"


            oItem = oOrderForm.Items.Item("e7")
            oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            oItem = oOrderForm.Items.Item("s7")
            oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Frequency"
            oOrderForm.PaneLevel = 12

            oItem = oOrderForm.Items.Item("s8")
            oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Cargo Type"
            oOrderForm.PaneLevel = 12


            oItem = oOrderForm.Items.Item("e8")
            oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "LCL"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType")

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left + oItem.Width + 10
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top ' + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "FCL"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("cs1")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Container Type"

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_ContType")
            ComboLoad_ContainerType(oForm, oCombo)
            oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("wcs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Packages"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            oNewItem = oOrderForm.Items.Item("wcs3")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("wcs3")
            oNewItem = oOrderForm.Items.Add("wcs4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            oNewItem = oOrderForm.Items.Item("wcs3")
            oNewItem.LinkTo = "cce5"

            oItem = oOrderForm.Items.Item("cce5")
            oNewItem = oOrderForm.Items.Add("wcs5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"
            ' oNewItem = oOrderForm.Items.Item("wcs5")
            oNewItem.LinkTo = "cce5"

            oItem = oOrderForm.Items.Item("wcs5")
            oNewItem = oOrderForm.Items.Add("cce6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + +oItem.Width + 10
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            oNewItem = oOrderForm.Items.Item("wcs5")
            oNewItem.LinkTo = "cce6"


        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Sea(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 0 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transfer To"
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = 70 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ComboLoad_TransferTo(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oItem.LinkTo = "e9"
            oItem = oOrderForm.Items.Item("e9")
            oItem.Enabled = False
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 10 + oItem.Width
            oNewItem.Width = oItem.Width + 20 '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transferred Status"
            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Width + oItem.Left
            oNewItem.Width = 60 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            oItem = oOrderForm.Items.Item("s10")
            oItem.LinkTo = "e10"
            oItem = oOrderForm.Items.Item("e10")
            oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Shipping Type"
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_ShipType")
            ComboLoad_ShipType(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("e10")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False
            '==============
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width + 17
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination Country"
            oItem = oOrderForm.Items.Item("e11")
            oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s13")
            oItem.LinkTo = "e13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s13")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination City"
            oItem = oOrderForm.Items.Item("e13")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e14")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("e2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessCode")
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessName")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Vessel"
            '  Exit Sub

            '===============================================
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e3")
            oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s2")
            oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++
            '===============================================
            oItem = oOrderForm.Items.Item("e3")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s3")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin City"
            '===============================================
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            '  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e6")
            oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e7"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transit Time"


            oItem = oOrderForm.Items.Item("e7")
            oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            oItem = oOrderForm.Items.Item("s7")
            oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Frequency"
            oOrderForm.PaneLevel = 12

            oItem = oOrderForm.Items.Item("s8")
            oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Cargo Type"
            oOrderForm.PaneLevel = 12


            oItem = oOrderForm.Items.Item("e8")
            oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "LCL"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType")

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            oNewItem.Left = oItem.Left + oItem.Width + 10
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top ' + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ooption = oNewItem.Specific
            ooption.Caption = "FCL"
            ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("cs1")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Container Type"

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_ContType")
            ComboLoad_ContainerType(oForm, oCombo)
            oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("wcs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Packages"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            oNewItem = oOrderForm.Items.Item("wcs3")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("wcs3")
            oNewItem = oOrderForm.Items.Add("wcs4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            oNewItem = oOrderForm.Items.Item("wcs3")
            oNewItem.LinkTo = "cce5"

            oItem = oOrderForm.Items.Item("wcs4")
            oNewItem = oOrderForm.Items.Add("wcs5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"

            oItem = oOrderForm.Items.Item("cce5")
            oNewItem = oOrderForm.Items.Add("cce6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            oNewItem = oOrderForm.Items.Item("wcs4")
            oNewItem.LinkTo = "cce6"


        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Air(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 0 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transfer To"
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = 70 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ComboLoad_TransferTo(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oItem.LinkTo = "e9"
            oItem = oOrderForm.Items.Item("e9")
            oItem.Enabled = False
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 10 + oItem.Width
            oNewItem.Width = oItem.Width + 20 '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transferred Status"
            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Width + oItem.Left
            oNewItem.Width = 60 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            oItem = oOrderForm.Items.Item("s10")
            oItem.LinkTo = "e10"
            oItem = oOrderForm.Items.Item("e10")
            oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Service Level"
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_SerLevel")
            ' ComboLoad_ShipType(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("e10")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False

            '' ''==============
            '' ''ComboLoad_Division(oForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            ' ''oNewItem.Width = oItem.Width + 17
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Destination Country"
            ' ''oItem = oOrderForm.Items.Item("e11")
            ' ''oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            '' ''GOPi
            '' ''ComboLoad_City(oOrderForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s13")
            ' ''oItem.LinkTo = "e13"
            '' ''  Exit Sub
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '' ''+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width + 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination Airport"
            oItem = oOrderForm.Items.Item("e11")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("e2c", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_CarricerC")
            'ComboLoad_Carrier(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e2c")
            oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_CarrierN")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e2c"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Carrier"
            '  Exit Sub

            ' '' ''===============================================
            '' ''oItem = oOrderForm.Items.Item("e2")
            '' ''oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = 80 'oItem.Width - 10
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            ' '' ''ComboLoad_City(oOrderForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("e3")
            '' ''oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left + oItem.Width + 2
            '' ''oNewItem.Width = 100
            '' ''oNewItem.Top = oItem.Top
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            ' '' ''ComboLoad_Division(oForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("s2")
            '' ''oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = oItem.Width '+ 85
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oNewItem.LinkTo = "e3"
            '' ''oStatic = oNewItem.Specific
            '' ''oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++
            '===============================================
            oItem = oOrderForm.Items.Item("e2c")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e22")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left '+ oItem.Width + 2
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s2")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin Airport"
            '===============================================
            'oItem = oOrderForm.Items.Item("ce3")
            'oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 300 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            'oItem = oOrderForm.Items.Item("cs3")
            'oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e3"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"

            'Include vesse
            '
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("vce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left '+ oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_VessName")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("vcs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "vce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Vessel"
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("vce3")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 30
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("vcs3")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            'oItem = oOrderForm.Items.Item("e6")
            'oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15 + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            'oItem = oOrderForm.Items.Item("s6")
            'oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15 + 14
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e7"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Transit Time"


            'oItem = oOrderForm.Items.Item("e7")
            'oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            'oItem = oOrderForm.Items.Item("s7")
            'oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e8"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Frequency"
            'oOrderForm.PaneLevel = 12

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 40
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Cargo Type"
            oOrderForm.PaneLevel = 12


            oItem = oOrderForm.Items.Item("e6")
            oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 40
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CaType")
            ComboLoad_CargoType(oForm, oCombo)
            'oItem = oOrderForm.Items.Item("ce1")
            'oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            'oNewItem.Left = oItem.Left + oItem.Width + 10
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top ' + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'ooption = oNewItem.Specific
            'ooption.Caption = "FCL"
            'ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            'ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("cs1")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Air_IN(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")
            oCombo.ValidValues.Add("IN", "International")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            oCombo.ValidValues.Add("AI", "Air Import")
            oCombo.ValidValues.Add("AE", "Air Export")
            ' ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
            '  Exit Sub
            '===============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 0 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transfer To"
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = 70 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ComboLoad_TransferTo(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oItem.LinkTo = "e9"
            oItem = oOrderForm.Items.Item("e9")
            oItem.Enabled = False
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left + 10 + oItem.Width
            oNewItem.Width = oItem.Width + 20 '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transferred Status"
            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Width + oItem.Left
            oNewItem.Width = 60 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            oItem = oOrderForm.Items.Item("s10")
            oItem.LinkTo = "e10"
            oItem = oOrderForm.Items.Item("e10")
            oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s9")
            oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Service Level"
            oItem = oOrderForm.Items.Item("e9")
            oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_SerLevel")
            ' ComboLoad_ShipType(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("s10")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("e10")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False

            '' ''==============
            '' ''ComboLoad_Division(oForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            ' ''oNewItem.Width = oItem.Width + 17
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Destination Country"
            ' ''oItem = oOrderForm.Items.Item("e11")
            ' ''oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            '' ''GOPi
            '' ''ComboLoad_City(oOrderForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s13")
            ' ''oItem.LinkTo = "e13"
            '' ''  Exit Sub
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '' ''+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s11")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width + 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination Airport"
            oItem = oOrderForm.Items.Item("e11")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("e2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CarricerC")
            ComboLoad_Carrier(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 2
            oNewItem.Width = 100
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_CarrierN")

            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e2"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Carrier"
            '  Exit Sub

            ' '' ''===============================================
            '' ''oItem = oOrderForm.Items.Item("e2")
            '' ''oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = 80 'oItem.Width - 10
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            ' '' ''ComboLoad_City(oOrderForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("e3")
            '' ''oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left + oItem.Width + 2
            '' ''oNewItem.Width = 100
            '' ''oNewItem.Top = oItem.Top
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            ' '' ''ComboLoad_Division(oForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("s2")
            '' ''oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = oItem.Width '+ 85
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oNewItem.LinkTo = "e3"
            '' ''oStatic = oNewItem.Specific
            '' ''oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++
            '===============================================
            oItem = oOrderForm.Items.Item("e2")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("e22")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left '+ oItem.Width + 2
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s2")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin Airport"
            '===============================================
            'oItem = oOrderForm.Items.Item("ce3")
            'oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 300 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            'oItem = oOrderForm.Items.Item("cs3")
            'oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e3"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 30
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            oItem = oOrderForm.Items.Item("e6")
            oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15 + 25
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15 + 25
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e7"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Transit Time"


            oItem = oOrderForm.Items.Item("e7")
            oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            oItem = oOrderForm.Items.Item("s7")
            oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Frequency"
            oOrderForm.PaneLevel = 12

            oItem = oOrderForm.Items.Item("s8")
            oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e8"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Cargo Type"
            oOrderForm.PaneLevel = 12


            oItem = oOrderForm.Items.Item("e8")
            oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CaType")
            ComboLoad_CargoType(oForm, oCombo)
            'oItem = oOrderForm.Items.Item("ce1")
            'oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            'oNewItem.Left = oItem.Left + oItem.Width + 10
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top ' + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'ooption = oNewItem.Specific
            'ooption.Caption = "FCL"
            'ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            'ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("cs1")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("ce1")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Local_IN(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")
            oCombo.ValidValues.Add("IN", "International")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            '===============================================
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            oCombo.ValidValues.Add("LC", "Local")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'oCombo.ValidValues.Add("SI", "Sea Export")
            ' ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
           
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width '+ 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False


            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin"
            '===============================================

            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination"
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")



            '===============================================
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs13")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 30
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 50
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 50
            oItem = oOrderForm.Items.Item("e6")
            oNewItem.Left = oItem.Left ' + oItem.Width
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("cs33")
            oNewItem = oOrderForm.Items.Add("cs34", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs34")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_Local(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()
            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
         
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width '+ 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False


            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin"
            '===============================================

            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination"
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

          

            '===============================================
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs13")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 30
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
           
            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 50
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 50
            oItem = oOrderForm.Items.Item("e6")
            oNewItem.Left = oItem.Left ' + oItem.Width
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("cs33")
            oNewItem = oOrderForm.Items.Add("cs34", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs34")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub SystemFormManuplation_International(ByVal oForm As SAPbouiCOM.Form, ByVal oOrderForm As SAPbouiCOM.Form)
        Try

            oNewItem = oOrderForm.Items.Add("Charge1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oOrderForm.Items.Item("2")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width + 60
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oButton = oNewItem.Specific
            oButton.Caption = "Load Charge Code"
            'oNewItem = oOrderForm.Items.Add("Charge2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem = oOrderForm.Items.Item("Charge1")
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = oItem.Height
            'oNewItem.Width = oItem.Width + 20
            'oNewItem.Left = oItem.Left + oItem.Width + 5
            'oButton = oNewItem.Specific
            'oButton.Caption = "Load Handling Charge Code"
            oNewItem = oOrderForm.Items.Add("General", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("112")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left - 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "General"
            oFolderItem.GroupWith("112")
            oFolderItem.Pane = 12
            oFolderItem.Select()

            '===============================================
            oItem = oOrderForm.Items.Item("18")
            oNewItem = oOrderForm.Items.Add("e0", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion1")
            oCombo.ValidValues.Add("IN", "International")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("19")
            oNewItem = oOrderForm.Items.Add("s0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e0"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Division"
            '  Exit Sub
            '===============================================
            oItem = oOrderForm.Items.Item("e0")
            oNewItem = oOrderForm.Items.Add("e1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oCombo = oNewItem.Specific
            oCombo.DataBind.SetBound(True, "ORDR", "U_AB_Divsion")
            'oCombo.ValidValues.Add("IN", "International")
            'oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s0")
            oNewItem = oOrderForm.Items.Add("s1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Sub Division"
            '  Exit Sub
            '===============================================
            'ComboLoad_Division(oForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("e1")
            ' ''oNewItem = oOrderForm.Items.Add("s9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left + 0 + oItem.Width
            ' ''oNewItem.Width = oItem.Width '+ 85
            ' ''oNewItem.Top = oItem.Top
            ' ''oNewItem.Height = 14N
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Transfer To"
            ' ''oItem = oOrderForm.Items.Item("s9")
            ' ''oNewItem = oOrderForm.Items.Add("e9", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            ' ''oNewItem.Left = oItem.Width + oItem.Left + 20
            ' ''oNewItem.Width = 70 'oItem.Width
            ' ''oNewItem.Top = oItem.Top
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oCombo = oNewItem.Specific
            ' ''oCombo.DataBind.SetBound(True, "ORDR", "U_AB_TransTo")
            ' ''ComboLoad_TransferTo(oOrderForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s9")
            ' ''oItem.LinkTo = "e9"
            ' ''oItem = oOrderForm.Items.Item("e9")
            ' ''oItem.Enabled = False
            '' ''  Exit Sub
            ' ''oItem = oOrderForm.Items.Item("e9")
            ' ''oNewItem = oOrderForm.Items.Add("s10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left + 10 + oItem.Width
            ' ''oNewItem.Width = oItem.Width + 20 '+ 85
            ' ''oNewItem.Top = oItem.Top
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Transferred Status"
            ' ''oItem = oOrderForm.Items.Item("s10")
            ' ''oNewItem = oOrderForm.Items.Add("e10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Width + oItem.Left
            ' ''oNewItem.Width = 60 'oItem.Width
            ' ''oNewItem.Top = oItem.Top
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Trnst")
            ' ''oItem = oOrderForm.Items.Item("s10")
            ' ''oItem.LinkTo = "e10"
            ' ''oItem = oOrderForm.Items.Item("e10")
            ' ''oItem.Enabled = False
            '=============================================
            'ComboLoad_Division(oForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s9")
            ' ''oNewItem = oOrderForm.Items.Add("s11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            ' ''oNewItem.Width = oItem.Width '+ 85
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Job Type"
            ' ''oItem = oOrderForm.Items.Item("e9")
            ' ''oNewItem = oOrderForm.Items.Add("e11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oCombo = oNewItem.Specific
            ' ''oCombo.DataBind.SetBound(True, "ORDR", "U_SerLevel")
            '' '' ComboLoad_ShipType(oOrderForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oItem.LinkTo = "e11"
            '  Exit Sub

            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("s12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Width + oItem.Left + 20
            oNewItem.Width = oItem.Width ' + 20 '+ 85
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job No"
            oItem = oOrderForm.Items.Item("s12")
            oNewItem = oOrderForm.Items.Add("e12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width '+ 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobNo")
            oItem = oOrderForm.Items.Item("s12")
            oItem.LinkTo = "e12"
            oItem = oOrderForm.Items.Item("e12")
            oItem.Enabled = False

            '' ''==============
            '' ''ComboLoad_Division(oForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("s13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' ''oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            ' ''oNewItem.Width = oItem.Width + 17
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            '' '' oNewItem.LinkTo = "e1"
            ' ''oStatic = oNewItem.Specific
            ' ''oStatic.Caption = "Destination Country"
            ' ''oItem = oOrderForm.Items.Item("e11")
            ' ''oNewItem = oOrderForm.Items.Add("e13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCode")
            '' ''GOPi
            '' ''ComboLoad_City(oOrderForm, oCombo)
            ' ''oItem = oOrderForm.Items.Item("s13")
            ' ''oItem.LinkTo = "e13"
            '' ''  Exit Sub
            ' ''oItem = oOrderForm.Items.Item("s11")
            ' ''oNewItem = oOrderForm.Items.Add("e14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' ''oNewItem.Left = oItem.Left
            ' ''oNewItem.Width = oItem.Width
            ' ''oNewItem.Top = oItem.Top + 15
            ' ''oNewItem.Height = 14
            ' ''oNewItem.FromPane = 12
            ' ''oNewItem.ToPane = 12
            ' ''oEdit = oNewItem.Specific
            ' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestName")
            '' ''+++++++++++++++++Destination City
            'ComboLoad_Division(oForm, oCombo)


            '===============================================
            oItem = oOrderForm.Items.Item("e1")
            oNewItem = oOrderForm.Items.Add("ce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCodeC")
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginNameC")
            'ComboLoad_Division(oForm, oCombo)
            oItem = oOrderForm.Items.Item("s1")
            oNewItem = oOrderForm.Items.Add("cs3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "ce3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Origin"
            '===============================================

            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("cs13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ 10 + oItem.Width
            oNewItem.Width = oItem.Width '+ 50
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            ' oNewItem.LinkTo = "e1"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Destination"
            oItem = oOrderForm.Items.Item("ce3")
            oNewItem = oOrderForm.Items.Add("ce13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestCodeC")
            'GOPi
            'ComboLoad_City(oOrderForm, oCombo)
            oItem = oOrderForm.Items.Item("cs13")
            oItem.LinkTo = "ce13"
            '  Exit Sub
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("ce14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left + oItem.Width + 20
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top '+ 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_DestNameC")

            ''===============================================
            'oItem = oOrderForm.Items.Item("e1")
            'oNewItem = oOrderForm.Items.Add("e2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oCombo = oNewItem.Specific
            'oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CarricerC")
            'ComboLoad_Carrier(oOrderForm, oCombo)
            'oItem = oOrderForm.Items.Item("e2")
            'oNewItem = oOrderForm.Items.Add("e22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left + oItem.Width + 2
            'oNewItem.Width = 100
            'oNewItem.Top = oItem.Top
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_CarrierN")

            ''ComboLoad_Division(oForm, oCombo)
            'oItem = oOrderForm.Items.Item("s1")
            'oNewItem = oOrderForm.Items.Add("s2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e2"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Carrier"
            '  Exit Sub

            ' '' ''===============================================
            '' ''oItem = oOrderForm.Items.Item("e2")
            '' ''oNewItem = oOrderForm.Items.Add("e3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = 80 'oItem.Width - 10
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OrginCode")
            ' '' ''ComboLoad_City(oOrderForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("e3")
            '' ''oNewItem = oOrderForm.Items.Add("e33", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            '' ''oNewItem.Left = oItem.Left + oItem.Width + 2
            '' ''oNewItem.Width = 100
            '' ''oNewItem.Top = oItem.Top
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oEdit = oNewItem.Specific
            '' ''oEdit.DataBind.SetBound(True, "ORDR", "U_AB_OriginName")
            ' '' ''ComboLoad_Division(oForm, oCombo)
            '' ''oItem = oOrderForm.Items.Item("s2")
            '' ''oNewItem = oOrderForm.Items.Add("s3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            '' ''oNewItem.Left = oItem.Left
            '' ''oNewItem.Width = oItem.Width '+ 85
            '' ''oNewItem.Top = oItem.Top + 15
            '' ''oNewItem.Height = 14
            '' ''oNewItem.FromPane = 12
            '' ''oNewItem.ToPane = 12
            '' ''oNewItem.LinkTo = "e3"
            '' ''oStatic = oNewItem.Specific
            '' ''oStatic.Caption = "Origin Country"

            '=++++++++++++++++++++++++++++++++++++++++++++

            '===============================================
            oItem = oOrderForm.Items.Item("ce13")
            oNewItem = oOrderForm.Items.Add("e4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Desc")
            oItem = oOrderForm.Items.Item("cs3")
            oNewItem = oOrderForm.Items.Add("s4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e3"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Item Description"
            ''  Exit Sub
            '=====================================================
            oItem = oOrderForm.Items.Item("e4")
            oNewItem = oOrderForm.Items.Add("e5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Validity")
            oItem = oOrderForm.Items.Item("s4")
            oNewItem = oOrderForm.Items.Add("s5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width ' + 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e5"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Validity"
            '
            '=====================================================
            '  Exit SubSIN
            oItem = oOrderForm.Items.Item("e5")
            oNewItem = oOrderForm.Items.Add("e6", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 300 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 30
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_JobScope")
            oItem = oOrderForm.Items.Item("s5")
            oNewItem = oOrderForm.Items.Add("s6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width '+ 85
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oNewItem.LinkTo = "e6"
            oStatic = oNewItem.Specific
            oStatic.Caption = "Job Scope"
            '
            '=====================================================
            '  Exit Sub
            'oItem = oOrderForm.Items.Item("e6")
            'oNewItem = oOrderForm.Items.Add("e7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15 + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Ttime")
            'oItem = oOrderForm.Items.Item("s6")
            'oNewItem = oOrderForm.Items.Add("s7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15 + 14
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e7"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Transit Time"


            'oItem = oOrderForm.Items.Item("e7")
            'oNewItem = oOrderForm.Items.Add("e8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oEdit = oNewItem.Specific
            'oEdit.DataBind.SetBound(True, "ORDR", "U_AB_Freq")
            'oItem = oOrderForm.Items.Item("s7")
            'oNewItem = oOrderForm.Items.Add("s8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oNewItem.Left = oItem.Left
            'oNewItem.Width = oItem.Width '+ 85
            'oNewItem.Top = oItem.Top + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'oNewItem.LinkTo = "e8"
            'oStatic = oNewItem.Specific
            'oStatic.Caption = "Frequency"
            'oOrderForm.PaneLevel = 12

            ''oItem = oOrderForm.Items.Item("s6")
            ''oNewItem = oOrderForm.Items.Add("cs1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ''oNewItem.Left = oItem.Left
            ''oNewItem.Width = oItem.Width '+ 85
            ''oNewItem.Top = oItem.Top + 40
            ''oNewItem.Height = 14
            ''oNewItem.FromPane = 12
            ''oNewItem.ToPane = 12
            ' '' oNewItem.LinkTo = "e8"
            ''oStatic = oNewItem.Specific
            ''oStatic.Caption = "Cargo Type"
            ''oOrderForm.PaneLevel = 12


            ''oItem = oOrderForm.Items.Item("e6")
            ''oNewItem = oOrderForm.Items.Add("ce1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            ''oNewItem.Left = oItem.Left
            ''oNewItem.Width = 80 'oItem.Width - 10
            ''oNewItem.Top = oItem.Top + 40
            ''oNewItem.Height = 14
            ''oNewItem.FromPane = 12
            ''oNewItem.ToPane = 12
            ''oCombo = oNewItem.Specific
            ''oCombo.DataBind.SetBound(True, "ORDR", "U_AB_CaType")
            ''ComboLoad_CargoType(oForm, oCombo)
            'oItem = oOrderForm.Items.Item("ce1")
            'oNewItem = oOrderForm.Items.Add("ce2", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON)
            'oNewItem.Left = oItem.Left + oItem.Width + 10
            'oNewItem.Width = 80 'oItem.Width - 10
            'oNewItem.Top = oItem.Top ' + 15
            'oNewItem.Height = 14
            'oNewItem.FromPane = 12
            'oNewItem.ToPane = 12
            'ooption = oNewItem.Specific
            'ooption.Caption = "FCL"
            'ooption.DataBind.SetBound(True, "ORDR", "U_AB_CaType1")
            'ooption.GroupWith("ce1")

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cs2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 50
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Weight"

            oItem = oOrderForm.Items.Item("s6")
            oNewItem = oOrderForm.Items.Add("cce3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 50
            oItem = oOrderForm.Items.Item("e6")
            oNewItem.Left = oItem.Left ' + oItem.Width
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotWT")
            ' oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs2")
            oNewItem.LinkTo = "cce3"

            oItem = oOrderForm.Items.Item("cs2")
            oNewItem = oOrderForm.Items.Add("cs33", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Pkgs"

            oItem = oOrderForm.Items.Item("cce3")
            oNewItem = oOrderForm.Items.Add("cce4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotPkg")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs33")
            oNewItem.LinkTo = "cce4"

            oItem = oOrderForm.Items.Item("cs33")
            oNewItem = oOrderForm.Items.Add("cs34", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oNewItem.Left = oItem.Left '+ oItem.Width + 10
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oStatic = oNewItem.Specific
            oStatic.Caption = "Total Volume"

            oItem = oOrderForm.Items.Item("cce4")
            oNewItem = oOrderForm.Items.Add("cce5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.Left = oItem.Left
            oNewItem.Width = 80 'oItem.Width - 10
            oNewItem.Top = oItem.Top + 15
            oNewItem.Height = 14
            oNewItem.FromPane = 12
            oNewItem.ToPane = 12
            oEdit = oNewItem.Specific
            oEdit.DataBind.SetBound(True, "ORDR", "U_AB_TotVol")
            'oNewItem.Enabled = False
            oNewItem = oOrderForm.Items.Item("cs34")
            oNewItem.LinkTo = "cce4"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#End Region
    Private Sub ComboLoad_Division(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM [@AB_DIVISION] T0 order by T0.COde")
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
    Private Sub ComboLoad_CargoType(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@AB_CARGOTYPE_AIR]  T0 ORDER BY T0.[Name]")
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
            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM [dbo].[@AB_SEAI_CONTTYPE]  T0 WHERE T0.[U_Division] ='SI' ORDER BY T0.[Code]")
            Dim it As Integer = 1
            For it = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(it, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombo.ValidValues.Add("", "")
            Do While Not oRecordSet.EoF
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Loop

            'oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oRecordSet = Nothing
            GC.Collect()
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ComboLoad_TransferTo(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[Code], T0.[Name] FROM [@AB_TRANSFERTO] T0 order by T0.COde")
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
    Private Sub ComboLoad_ShipType(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[TrnspName], T0.[TrnspCode] FROM OSHP T0 ORDER BY T0.[TrnspName]")
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
    Public Sub Salese_Order_ContMatrix(ByVal Oform As SAPbouiCOM.Form)
        Try
            Exit Sub
            oNewItem = oOrderForm.Items.Add("Cont", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("Cargo")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "Container Details"
            oFolderItem.GroupWith("Cargo")
            oFolderItem.Pane = 15
            'oFolderItem.Select()
            oItem = oOrderForm.Items.Item("38")
            oNewItem = Oform.Items.Add("MATCont", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.FromPane = 15
            oNewItem.ToPane = 15
            oMatrix1 = oNewItem.Specific
            oColumns = oMatrix1.Columns
            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "#"
            oColumn.Width = 30
            oColumn.Editable = False
            Oform.DataSources.UserDataSources.Add("LineNo", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oColumn = oColumns.Item("#")
            oColumn.DataBind.SetBound(True, "", "LineNo")
            '

            oColumn = oColumns.Add("ContNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Container No"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("ContNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("ContNo")
            oColumn.DataBind.SetBound(True, "", "ContNo")

            oColumn = oColumns.Add("Size", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Size/Type"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("Size")
            oColumn.DataBind.SetBound(True, "", "Size")


            oColumn = oColumns.Add("Wt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Weight"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("Wt1", SAPbouiCOM.BoDataType.dt_MEASURE)
            oColumn = oColumns.Item("Wt")
            oColumn.DataBind.SetBound(True, "", "Wt1")
            oColumn = oColumns.Add("SealNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Seal No"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("SealNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("SealNo")
            oColumn.DataBind.SetBound(True, "", "SealNo")

            oMatrix1.AddRow()
            oMatrix1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            ' oMatrix1.Clear()

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub Salese_Order_CargoMatrix(ByVal Oform As SAPbouiCOM.Form)
        Try
            Exit Sub
            oNewItem = oOrderForm.Items.Add("Cargo", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oOrderForm.Items.Item("General")
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.Width = oItem.Width
            oNewItem.Left = oItem.Left + oItem.Width + 5
            oFolderItem = oNewItem.Specific
            oFolderItem.Caption = "Cargo Details"
            oFolderItem.GroupWith("General")
            oFolderItem.Pane = 13
            'oFolderItem.Select()


            oItem = oOrderForm.Items.Item("38")
            oNewItem = Oform.Items.Add("MATCargo", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            oNewItem.Left = oItem.Left
            oNewItem.Width = oItem.Width
            oNewItem.Top = oItem.Top
            oNewItem.Height = oItem.Height
            oNewItem.FromPane = 13
            oNewItem.ToPane = 13
            oMatrix = oNewItem.Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "#"
            oColumn.Width = 30
            oColumn.Editable = False
            Oform.DataSources.UserDataSources.Add("V_-1", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oColumn = oColumns.Item("#")
            oColumn.DataBind.SetBound(True, "", "V_-1")
            '

            oColumn = oColumns.Add("SupCode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Supplier Code"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("SupCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("SupCode")
            oColumn.DataBind.SetBound(True, "", "SupCode")

            oColumn = oColumns.Add("SupName", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Supplier Name"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("SupName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("SupName")
            oColumn.DataBind.SetBound(True, "", "SupName")


            oColumn = oColumns.Add("PONo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "PO No"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("PONo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("PONo")
            oColumn.DataBind.SetBound(True, "", "PONo")

            oColumn = oColumns.Add("Pkg", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Pcks #"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("Pkg", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oColumn = oColumns.Item("Pkg")
            oColumn.DataBind.SetBound(True, "", "Pkg")

            oColumn = oColumns.Add("PkgT", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Pkg. Type"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("PkgT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("PkgT")
            oColumn.DataBind.SetBound(True, "", "PkgT")

            oColumn = oColumns.Add("Wt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Wt(kgs)"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("Wt", SAPbouiCOM.BoDataType.dt_MEASURE)
            oColumn = oColumns.Item("Wt")
            oColumn.DataBind.SetBound(True, "", "Wt")

            oColumn = oColumns.Add("L", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "L(cms)"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("L", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oColumn = oColumns.Item("L")
            oColumn.DataBind.SetBound(True, "", "L")

            oColumn = oColumns.Add("W", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "W(cms)"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("W", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oColumn = oColumns.Item("W")
            oColumn.DataBind.SetBound(True, "", "W")

            oColumn = oColumns.Add("H", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "H(cms)"
            oColumn.Width = 40
            oColumn.Editable = True

            Oform.DataSources.UserDataSources.Add("H", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oColumn = oColumns.Item("H")
            oColumn.DataBind.SetBound(True, "", "H")

            oColumn = oColumns.Add("m3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "m3"
            oColumn.Width = 40
            oColumn.Editable = True

            Oform.DataSources.UserDataSources.Add("m3", SAPbouiCOM.BoDataType.dt_MEASURE)
            oColumn = oColumns.Item("m3")
            oColumn.DataBind.SetBound(True, "", "m3")

            oColumn = oColumns.Add("vol", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Vol/Wt"
            oColumn.Width = 40
            oColumn.Editable = True

            Oform.DataSources.UserDataSources.Add("vol", SAPbouiCOM.BoDataType.dt_MEASURE)
            oColumn = oColumns.Item("vol")
            oColumn.DataBind.SetBound(True, "", "vol")

            oColumn = oColumns.Add("Desc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Description"
            oColumn.Width = 40
            oColumn.Editable = True
            Oform.DataSources.UserDataSources.Add("Desc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oColumn = oColumns.Item("Desc")
            oColumn.DataBind.SetBound(True, "", "Desc")
            oMatrix.AddRow()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            'oMatrix.Clear()
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Add_Container_details(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = Ocompany.GetCompanyService
            Dim DocNum As String = ""
            Dim LineNo As Integer = 0
            Dim ContNo As String = ""
            Dim Size As String = ""
            Dim Wt As Double = 0.0
            Dim SealNo As String = ""
            oGeneralService = sCmp.GetGeneralService("AB_SALESORDER_CONT")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            Dim i As Integer = 1
            oMatrix1 = oForm.Items.Item("MATCont").Specific
            oColumns = oMatrix1.Columns
            oCombo = oForm.Items.Item("88").Specific
            Dim series As String = oCombo.Selected.Value
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[ObjectCode], T0.[Series], T0.[SeriesName], T0.[InitialNum], T0.[NextNumber], T0.[LastNum], T0.[BeginStr], T0.[EndStr] FROM NNM1 T0 WHERE T0.[ObjectCode] =17 and T0.[Series]='" & series & "'")
            DocNum = oRecordSet1.Fields.Item("NextNumber").Value
            For i = 1 To oMatrix1.RowCount
                oEdit = oMatrix1.Columns.Item("ContNo").Cells.Item(i).Specific
                ContNo = oEdit.String
                oEdit = oMatrix1.Columns.Item("Size").Cells.Item(i).Specific
                Size = oEdit.String
                oEdit = oMatrix1.Columns.Item("Wt").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    Wt = oEdit.Value
                Else
                    Wt = 0
                End If


                oEdit = oMatrix1.Columns.Item("SealNo").Cells.Item(i).Specific
                SealNo = oEdit.String
                If ContNo <> "" Then
                    oGeneralData.SetProperty("U_SONo", DocNum)
                    oGeneralData.SetProperty("U_LineNo", i)
                    oGeneralData.SetProperty("U_ContNo", ContNo)
                    oGeneralData.SetProperty("U_Size", Size)
                    oGeneralData.SetProperty("U_Wt", Wt)
                    oGeneralData.SetProperty("U_SealNo", SealNo)
                    Try
                        oGeneralService.Add(oGeneralData)
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText("Error: Add_Container_details-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If

            Next
            oMatrix1.Clear()
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Get_Container_details(ByVal oForm As SAPbouiCOM.Form)
        Try

            oForm.Freeze(True)
            oEdit = oForm.Items.Item("8").Specific
            Dim sqlstr As String = "SELECT T0.[U_LineNo]  ,T0.[U_ContNo] , T0.[U_Size] , T0.[U_Wt]  ,T0.[U_SealNo] FROM [dbo].[@AB_SALESORDER_CONT]  T0 WHERE T0.[U_SONo]='" & oEdit.String & "' ORDER BY T0.[U_LineNo], T0.[DocEntry]"
            oForm.DataSources.DataTables.Item("OCARGO").ExecuteQuery(sqlstr)
            oMatrix = oForm.Items.Item("MATCont").Specific
            oMatrix.Clear()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
            oColumns = oMatrix.Columns
            oForm.Items.Item("MATCont").Specific.Columns.item("#").DataBind.Bind("OCARGO", "U_LineNo")
            oForm.Items.Item("MATCont").Specific.Columns.item("ContNo").DataBind.Bind("OCARGO", "U_ContNo")
            oForm.Items.Item("MATCont").Specific.Columns.item("Size").DataBind.Bind("OCARGO", "U_Size")
            oForm.Items.Item("MATCont").Specific.Columns.item("Wt").DataBind.Bind("OCARGO", "U_Wt")
            oForm.Items.Item("MATCont").Specific.Columns.item("SealNo").DataBind.Bind("OCARGO", "U_SealNo")
            oForm.Items.Item("MATCont").Specific.Clear()
            oForm.Items.Item("MATCont").Specific.LoadFromDataSource()
            oForm.Items.Item("MATCont").Specific.AutoResizeColumns()
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Update_Container_details(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = Ocompany.GetCompanyService
            Dim DocNum As String = ""
            Dim LineNo As Integer = 0
            Dim ContNo As String = ""
            Dim Size As String = ""
            Dim Wt As Double = 0.0
            Dim SealNo As String = ""
            oGeneralService = sCmp.GetGeneralService("AB_SALESORDER_CONT")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            Dim oHeaderParams1 As SAPbobsCOM.GeneralDataParams
            oHeaderParams1 = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            Dim oRecordSet_EC1 As SAPbobsCOM.Recordset
            oRecordSet_EC1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
           
            Dim i As Integer = 1
            oMatrix1 = oForm.Items.Item("MATCont").Specific
            oColumns = oMatrix1.Columns
            oCombo = oForm.Items.Item("88").Specific
            Dim series As String = oCombo.Selected.Value
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[ObjectCode], T0.[Series], T0.[SeriesName], T0.[InitialNum], T0.[NextNumber], T0.[LastNum], T0.[BeginStr], T0.[EndStr] FROM NNM1 T0 WHERE T0.[ObjectCode] =17 and T0.[Series]='" & series & "'")
            DocNum = oRecordSet1.Fields.Item("NextNumber").Value
            oEdit = oForm.Items.Item("8").Specific
            Dim SODocNo As String = oEdit.String
            For i = 1 To oMatrix1.RowCount
                oRecordSet_EC1.DoQuery("SELECT T0.[DocEntry] FROM [dbo].[@AB_SALESORDER_CONT]  T0 WHERE T0.[U_SONo] ='" & SODocNo & "' and  T0.[U_LineNo] ='" & i & "'")
                If oRecordSet_EC1.RecordCount > 0 Then '--update udo
                    oHeaderParams1.SetProperty("DocEntry", oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim()) '
                    oGeneralData = oGeneralService.GetByParams(oHeaderParams1)
                    oEdit = oMatrix1.Columns.Item("ContNo").Cells.Item(i).Specific
                    ContNo = oEdit.String
                    oEdit = oMatrix1.Columns.Item("Size").Cells.Item(i).Specific
                    Size = oEdit.String
                    oEdit = oMatrix1.Columns.Item("Wt").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        Wt = oEdit.Value
                    Else
                        Wt = 0
                    End If

                    oEdit = oMatrix1.Columns.Item("SealNo").Cells.Item(i).Specific
                    SealNo = oEdit.String
                    ' If ContNo <> "" Then
                    'oGeneralData.SetProperty("U_SONo", DocNum)
                    'oGeneralData.SetProperty("U_LineNo", i)
                    oGeneralData.SetProperty("U_ContNo", ContNo)
                    oGeneralData.SetProperty("U_Size", Size)
                    oGeneralData.SetProperty("U_Wt", Wt)
                    oGeneralData.SetProperty("U_SealNo", SealNo)
                    Try
                        oGeneralService.Update(oGeneralData)
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText("Error: Add_Container_details-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                    'End If

                Else '---add udo
                    oEdit = oMatrix1.Columns.Item("ContNo").Cells.Item(i).Specific
                    ContNo = oEdit.String
                    oEdit = oMatrix1.Columns.Item("Size").Cells.Item(i).Specific
                    Size = oEdit.String
                    oEdit = oMatrix1.Columns.Item("Wt").Cells.Item(i).Specific
                    Wt = oEdit.Value
                    oEdit = oMatrix1.Columns.Item("SealNo").Cells.Item(i).Specific
                    SealNo = oEdit.String
                    If ContNo <> "" Then
                        oGeneralData.SetProperty("U_SONo", SODocNo)
                        oGeneralData.SetProperty("U_LineNo", i)
                        oGeneralData.SetProperty("U_ContNo", ContNo)
                        oGeneralData.SetProperty("U_Size", Size)
                        oGeneralData.SetProperty("U_Wt", Wt)
                        oGeneralData.SetProperty("U_SealNo", SealNo)
                        Try
                            oGeneralService.Add(oGeneralData)
                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText("Error: Add_Container_details-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If


                End If
            Next
            ' oMatrix1.Clear()
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Add_Cargo_details(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = Ocompany.GetCompanyService
            Dim DocNum As String = ""
            Dim LineNo As Integer = 0
            Dim SUppCOde As String = ""
            Dim SUppName As String = ""
            Dim PONo As String = ""
            Dim Pkg As Integer = 0
            Dim PType As String = ""
            Dim Wt As Double = 0
            Dim L As Integer = 0
            Dim w As Integer = 0
            Dim H As Integer = 0
            Dim m3 As Double = 0
            Dim vol As Double = 0
            Dim Desc As String = 0

            oGeneralService = sCmp.GetGeneralService("AB_SALESORDER_CARGO")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            Dim i As Integer = 1
            oMatrix = oForm.Items.Item("MATCargo").Specific
            oColumns = oMatrix.Columns
            oCombo = oForm.Items.Item("88").Specific
            Dim series As String = oCombo.Selected.Value
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[ObjectCode], T0.[Series], T0.[SeriesName], T0.[InitialNum], T0.[NextNumber], T0.[LastNum], T0.[BeginStr], T0.[EndStr] FROM NNM1 T0 WHERE T0.[ObjectCode] =17 and T0.[Series]='" & series & "'")
            DocNum = oRecordSet1.Fields.Item("NextNumber").Value
            For i = 1 To oMatrix.RowCount
                oEdit = oMatrix.Columns.Item("SupCode").Cells.Item(i).Specific
                SUppCOde = oEdit.String
                oEdit = oMatrix.Columns.Item("SupName").Cells.Item(i).Specific
                SUppName = oEdit.String
                oEdit = oMatrix.Columns.Item("PONo").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    PONo = oEdit.Value
                Else
                    PONo = ""
                End If

                oEdit = oMatrix.Columns.Item("Pkg").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    Pkg = oEdit.Value
                Else
                    Pkg = 0
                End If

                oEdit = oMatrix.Columns.Item("PkgT").Cells.Item(i).Specific
                PType = oEdit.String
                oEdit = oMatrix.Columns.Item("Wt").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    Wt = oEdit.Value
                Else
                    Wt = 0
                End If

                oEdit = oMatrix.Columns.Item("L").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    L = oEdit.Value
                Else
                    L = 0
                End If

                oEdit = oMatrix.Columns.Item("W").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    w = oEdit.Value
                Else
                    w = 0
                End If

                oEdit = oMatrix.Columns.Item("H").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    H = oEdit.Value
                Else
                    H = 0
                End If
                oEdit = oMatrix.Columns.Item("m3").Cells.Item(i).Specific
                If oEdit.String <> "" Then
                    m3 = oEdit.Value
                Else
                    m3 = 0
                End If

                oEdit = oMatrix.Columns.Item("vol").Cells.Item(i).Specific
                If (oEdit.String <> "") Then
                    vol = oEdit.Value
                Else
                    vol = 0
                End If
                oEdit = oMatrix.Columns.Item("Desc").Cells.Item(i).Specific
                Desc = oEdit.String
                If SUppCOde <> "" Then
                    oGeneralData.SetProperty("U_SONo", DocNum)
                    oGeneralData.SetProperty("U_LineNo", i)
                    oGeneralData.SetProperty("U_VenCode", SUppCOde)
                    oGeneralData.SetProperty("U_VenName", SUppName)
                    oGeneralData.SetProperty("U_PONO", PONo)
                    oGeneralData.SetProperty("U_PKg", Pkg)
                    oGeneralData.SetProperty("U_PkgType", PType)
                    oGeneralData.SetProperty("U_Wt", Wt)
                    oGeneralData.SetProperty("U_Len", L)
                    oGeneralData.SetProperty("U_Width", w)
                    oGeneralData.SetProperty("U_Height", H)
                    oGeneralData.SetProperty("U_M3", m3)
                    oGeneralData.SetProperty("U_Vol", vol)
                    oGeneralData.SetProperty("U_Desc", Desc)
                    Try
                        oGeneralService.Add(oGeneralData)
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText("Error: Add_Container_details-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If

            Next
            oMatrix.Clear()
        Catch ex As Exception
            '  SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Get_Cargo_details(ByVal oForm As SAPbouiCOM.Form)
        Try

            oForm.Freeze(True)
            oEdit = oForm.Items.Item("8").Specific
            Dim sqlstr As String = "SELECT T0.[U_LineNo], T0.[U_VenCode], T0.[U_VenName], T0.[U_PONO], T0.[U_PKg], T0.[U_PkgType], T0.[U_Wt], T0.[U_Len], T0.[U_Width], T0.[U_Height], T0.[U_M3], T0.[U_Vol], T0.[U_Desc] FROM [dbo].[@AB_SALESORDER_CARGO]  T0 WHERE T0.[U_SONo] ='" & oEdit.String & "' ORDER BY T0.[U_LineNo], T0.[DocEntry]"
            oForm.DataSources.DataTables.Item("OCARGO1").ExecuteQuery(sqlstr)
            oMatrix1 = oForm.Items.Item("MATCargo").Specific
            oMatrix1.Clear()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
            oColumns = oMatrix.Columns
            oForm.Items.Item("MATCargo").Specific.Columns.item("#").DataBind.Bind("OCARGO1", "U_LineNo")
            oForm.Items.Item("MATCargo").Specific.Columns.item("SupCode").DataBind.Bind("OCARGO1", "U_VenCode")
            oForm.Items.Item("MATCargo").Specific.Columns.item("SupName").DataBind.Bind("OCARGO1", "U_VenName")
            oForm.Items.Item("MATCargo").Specific.Columns.item("PONo").DataBind.Bind("OCARGO1", "U_PONO")
            oForm.Items.Item("MATCargo").Specific.Columns.item("Pkg").DataBind.Bind("OCARGO1", "U_PKg")

            oForm.Items.Item("MATCargo").Specific.Columns.item("PkgT").DataBind.Bind("OCARGO1", "U_PkgType")
            oForm.Items.Item("MATCargo").Specific.Columns.item("Wt").DataBind.Bind("OCARGO1", "U_Wt")
            oForm.Items.Item("MATCargo").Specific.Columns.item("L").DataBind.Bind("OCARGO1", "U_Len")
            oForm.Items.Item("MATCargo").Specific.Columns.item("W").DataBind.Bind("OCARGO1", "U_Width")
            oForm.Items.Item("MATCargo").Specific.Columns.item("H").DataBind.Bind("OCARGO1", "U_Height")

            oForm.Items.Item("MATCargo").Specific.Columns.item("m3").DataBind.Bind("OCARGO1", "U_M3")
            oForm.Items.Item("MATCargo").Specific.Columns.item("vol").DataBind.Bind("OCARGO1", "U_Vol")
            oForm.Items.Item("MATCargo").Specific.Columns.item("Desc").DataBind.Bind("OCARGO1", "U_Desc")

            oForm.Items.Item("MATCargo").Specific.Clear()
            oForm.Items.Item("MATCargo").Specific.LoadFromDataSource()
            oForm.Items.Item("MATCargo").Specific.AutoResizeColumns()
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            ' SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub Update_Cargo_details(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = Ocompany.GetCompanyService
            Dim DocNum As String = ""
            Dim LineNo As Integer = 0
            Dim SUppCOde As String = ""
            Dim SUppName As String = ""
            Dim PONo As String = ""
            Dim Pkg As Integer = 0
            Dim PType As String = ""
            Dim Wt As Double = 0
            Dim L As Integer = 0
            Dim w As Integer = 0
            Dim H As Integer = 0
            Dim m3 As Double = 0
            Dim vol As Double = 0
            Dim Desc As String = ""
            oGeneralService = sCmp.GetGeneralService("AB_SALESORDER_CARGO")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            Dim oHeaderParams1 As SAPbobsCOM.GeneralDataParams
            oHeaderParams1 = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            Dim oRecordSet_EC1 As SAPbobsCOM.Recordset
            oRecordSet_EC1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim i As Integer = 1
            oMatrix = oForm.Items.Item("MATCargo").Specific
            oColumns = oMatrix.Columns
            oCombo = oForm.Items.Item("88").Specific
            Dim series As String = oCombo.Selected.Value
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[ObjectCode], T0.[Series], T0.[SeriesName], T0.[InitialNum], T0.[NextNumber], T0.[LastNum], T0.[BeginStr], T0.[EndStr] FROM NNM1 T0 WHERE T0.[ObjectCode] =17 and T0.[Series]='" & series & "'")
            DocNum = oRecordSet1.Fields.Item("NextNumber").Value
            oEdit = oForm.Items.Item("8").Specific
            Dim SODocNo As String = oEdit.String
            For i = 1 To oMatrix.RowCount
                oRecordSet_EC1.DoQuery("SELECT T0.[DocEntry] FROM [dbo].[@AB_SALESORDER_CARGO]  T0 WHERE T0.[U_SONo] ='" & SODocNo & "' and  T0.[U_LineNo] ='" & i & "'")
                If oRecordSet_EC1.RecordCount > 0 Then '--update udo
                    oHeaderParams1.SetProperty("DocEntry", oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim()) '
                    oGeneralData = oGeneralService.GetByParams(oHeaderParams1)
                    oEdit = oMatrix.Columns.Item("SupCode").Cells.Item(i).Specific
                    SUppCOde = oEdit.String
                    oEdit = oMatrix.Columns.Item("SupName").Cells.Item(i).Specific
                    SUppName = oEdit.String
                    oEdit = oMatrix.Columns.Item("PONo").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        PONo = oEdit.Value
                    Else
                        PONo = ""
                    End If
                    oEdit = oMatrix.Columns.Item("Pkg").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        Pkg = oEdit.Value
                    Else
                        Pkg = 0
                    End If
                    oEdit = oMatrix.Columns.Item("PkgT").Cells.Item(i).Specific
                    PType = oEdit.String
                    oEdit = oMatrix.Columns.Item("Wt").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        Wt = oEdit.Value
                    Else
                        Wt = 0
                    End If
                    oEdit = oMatrix.Columns.Item("L").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        L = oEdit.Value
                    Else
                        L = 0
                    End If
                    oEdit = oMatrix.Columns.Item("W").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        w = oEdit.Value
                    Else
                        w = 0
                    End If
                    oEdit = oMatrix.Columns.Item("H").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        H = oEdit.Value
                    Else
                        H = 0
                    End If
                    oEdit = oMatrix.Columns.Item("m3").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        m3 = oEdit.Value
                    Else
                        m3 = 0
                    End If
                    oEdit = oMatrix.Columns.Item("vol").Cells.Item(i).Specific
                    If (oEdit.String <> "") Then
                        vol = oEdit.Value
                    Else
                        vol = 0
                    End If
                    oEdit = oMatrix.Columns.Item("Desc").Cells.Item(i).Specific
                    Desc = oEdit.String
                    If SUppCOde <> "" Then
                        'oGeneralData.SetProperty("U_SONo", SODocNo)
                        'oGeneralData.SetProperty("U_LineNo", i)
                        oGeneralData.SetProperty("U_VenCode", SUppCOde)
                        oGeneralData.SetProperty("U_VenName", SUppName)
                        oGeneralData.SetProperty("U_PONO", PONo)
                        oGeneralData.SetProperty("U_PKg", Pkg)
                        oGeneralData.SetProperty("U_PkgType", PType)
                        oGeneralData.SetProperty("U_Wt", Wt)
                        oGeneralData.SetProperty("U_Len", L)
                        oGeneralData.SetProperty("U_Width", w)
                        oGeneralData.SetProperty("U_Height", H)
                        oGeneralData.SetProperty("U_M3", m3)
                        oGeneralData.SetProperty("U_Vol", vol)
                        oGeneralData.SetProperty("U_Desc", Desc)
                        Try
                            oGeneralService.Update(oGeneralData)
                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText("Error: Add_Container_details-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Else '---add udo
                    oEdit = oMatrix.Columns.Item("SupCode").Cells.Item(i).Specific
                    SUppCOde = oEdit.String
                    oEdit = oMatrix.Columns.Item("SupName").Cells.Item(i).Specific
                    SUppName = oEdit.String
                    oEdit = oMatrix.Columns.Item("PONo").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        PONo = oEdit.Value
                    Else
                        PONo = ""
                    End If
                    oEdit = oMatrix.Columns.Item("Pkg").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        Pkg = oEdit.Value
                    Else
                        Pkg = 0
                    End If
                    oEdit = oMatrix.Columns.Item("PkgT").Cells.Item(i).Specific
                    PType = oEdit.String
                    oEdit = oMatrix.Columns.Item("Wt").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        Wt = oEdit.Value
                    Else
                        Wt = 0
                    End If
                    oEdit = oMatrix.Columns.Item("L").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        L = oEdit.Value
                    Else
                        L = 0
                    End If
                    oEdit = oMatrix.Columns.Item("W").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        w = oEdit.Value
                    Else
                        w = 0
                    End If
                    oEdit = oMatrix.Columns.Item("H").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        H = oEdit.Value
                    Else
                        H = 0
                    End If
                    oEdit = oMatrix.Columns.Item("m3").Cells.Item(i).Specific
                    If oEdit.String <> "" Then
                        m3 = oEdit.Value
                    Else
                        m3 = 0
                    End If

                    oEdit = oMatrix.Columns.Item("vol").Cells.Item(i).Specific
                    If (oEdit.String <> "") Then
                        vol = oEdit.Value
                    Else
                        vol = 0
                    End If
                    oEdit = oMatrix.Columns.Item("Desc").Cells.Item(i).Specific
                    Desc = oEdit.String
                    If SUppCOde <> "" Then
                        oGeneralData.SetProperty("U_SONo", SODocNo)
                        oGeneralData.SetProperty("U_LineNo", i)
                        oGeneralData.SetProperty("U_VenCode", SUppCOde)
                        oGeneralData.SetProperty("U_VenName", SUppName)
                        oGeneralData.SetProperty("U_PONO", PONo)
                        oGeneralData.SetProperty("U_PKg", Pkg)
                        oGeneralData.SetProperty("U_PkgType", PType)
                        oGeneralData.SetProperty("U_Wt", Wt)
                        oGeneralData.SetProperty("U_Len", L)
                        oGeneralData.SetProperty("U_Width", w)
                        oGeneralData.SetProperty("U_Height", H)
                        oGeneralData.SetProperty("U_M3", m3)
                        oGeneralData.SetProperty("U_Vol", vol)
                        oGeneralData.SetProperty("U_Desc", Desc)
                        Try
                            oGeneralService.Add(oGeneralData)
                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText("Error: Add_Container_details-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                End If
            Next
            ' oMatrix1.Clear()
        Catch ex As Exception
            '  SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    'Private Sub LoadHandingCharge_LCL_SI_SP(ByVal oForm As SAPbouiCOM.Form, ByVal CustCode As String)
    '    Try
    '        oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecordSet1.DoQuery("SELECT T1.U_CargoType, T1.U_VendorCode, T1.U_VendorName, T1.U_ChargeCode, T1.U_ChargeDesc, T1.U_Quantity, T1.U_Unit, T1.U_Cost, T1.U_SellingP, T1.U_MarkUp, T1.U_SpecialSellingP, T1.U_Remarks FROM [dbo].[@AB_SEAI_SPRICE]  T0 , [dbo].[@AB_SEAI_SPRICELCL]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Division] ='SI' and  T0.[Code] like '" & CustCode & "'")
    '        If oRecordSet1.RecordCount = 0 Then
    '            SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If
    '        oMatrix3 = oForm.Items.Item("38").Specific
    '        oColumns = oMatrix3.Columns
    '        oMatrix3.Clear()
    '        Dim i As Integer = 0
    '        For i = 1 To oRecordSet1.RecordCount
    '            oMatrix3.AddRow()
    '            oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_VendorCode").Value
    '            oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
    '            oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_ChargeDesc").Value
    '            oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
    '            oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
    '            'U_Quantity
    '            oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
    '            oRecordSet1.MoveNext()
    '        Next
    '    Catch ex As Exception
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
#Region "Sea Price Loading"
    Private Sub LoadHandingCharge_LCL_SI(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT  T1.U_VendorCode, T1.U_VendorName, T1.U_ChargeCode, T1.U_ChargeDesc, T1.U_Quantity, T1.U_Unit, T1.U_CargoType, T1.U_Cost, T1.U_SellingP, T1.U_Remarks FROM [dbo].[@AB_SEA_HC]  T0 , [dbo].[@AB_SEA_HCLCL]  T1 WHERE T0.[Code] = T1.[Code] and  T0.[U_Division] ='" & Division & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No-" & oRecordSet1.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oMatrix3.AddRow()
                'oEdit = oMatrix3.Columns.Item("U_Ven").Cells.Item(1).Specific
                'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
                oEdit = oMatrix3.Columns.Item("212").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                'U_Quantity
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_FCL_SI(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oCombo = oForm.Items.Item("cce3").Specific
            Dim ContType As String = ""
            Try
                ContType = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            If ContType = "" Then
                SBO_Application.StatusBar.SetText("Select Container Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_ChargeCode, T1.U_ChargeDesc, T1.U_CargoType, T1.U_Quantity, T1.U_Unit, T1.U_Cost, T1.U_SellingP, T1.U_Remarks, T1.U_ContainerType FROM [dbo].[@AB_SEA_HC]  T0 , [dbo].[@AB_SEA_HCFCL]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='" & Division & "' and  T1.[U_ContainerType]  like '" & ContType & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm.Freeze(True)
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & oRecordSet1.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oMatrix3.AddRow()
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
                'U_Quantity
                oEdit = oMatrix3.Columns.Item("212").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    'LoadHandingCharge_LCL_SE_SpecialPrice
    Private Sub LoadHandingCharge_LCL_SE_SpecialPrice(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String, ByVal CCode As String)
        Try
            Dim DestCountry As String = ""
            Dim DestCity As String = ""
            oEdit = oForm.Items.Item("e13").Specific
            DestCountry = oEdit.String
            oEdit = oForm.Items.Item("ce13").Specific
            DestCity = oEdit.String
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.[U_VendorCode], T1.[U_VendorName], T1.[U_ChargeCode], T1.[U_ChargeDesc], T1.[U_Quantity], T1.[U_Unit], T1.[U_Cost],T1.[U_SellingP] FROM [dbo].[@AB_SEAE_SPRICE]  T0 , [dbo].[@AB_SEAE_SPRICELCL]  T1 WHERE T0.[Code] = T1.[Code]  and  T1.[Code] ='" & CCode & "' and  T1.[U_Country] ='" & DestCountry & "' and  T1.[U_City] ='" & DestCity & "'  and  T0.[U_Division] ='" & Division & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            'SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'oEdit = oMatrix3.Columns.Item("U_Ven").Cells.Item(1).Specific
                'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
                oEdit = oMatrix3.Columns.Item("212").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                'U_Quantity
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    'LoadHandingCharge_FCL_SE_SpecialPrice
    Private Sub LoadHandingCharge_FCL_SE_SpecialPrice(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String, ByVal CCode As String)
        Try
            Dim DestCountry As String = ""
            Dim DestCity As String = ""
            oEdit = oForm.Items.Item("e13").Specific
            DestCountry = oEdit.String
            oEdit = oForm.Items.Item("ce13").Specific
            DestCity = oEdit.String

            '  SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oCombo = oForm.Items.Item("cce3").Specific
            Dim ContType As String = ""
            Try
                ContType = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            If ContType = "" Then
                SBO_Application.StatusBar.SetText("Select Container Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.[U_VendorCode], T1.[U_VendorName], T1.[U_ChargeCode], T1.[U_ChargeDesc], T1.[U_Quantity], T1.[U_Cost], T1.[U_Unit], T1.[U_SumSellingP] FROM [dbo].[@AB_SEAE_SPRICE]  T0 , [dbo].[@AB_SEAE_SPRICEFCL]  T1 WHERE T1.[Code] = T0.[Code] and T1.[U_ContainerType] ='" & ContType & "' and  T0.[Code] ='" & CCode & "' and  T1.[U_Country] ='" & DestCountry & "'and  T1.[U_City] ='" & DestCity & "' and   T0.[U_Division] ='" & Division & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm.Freeze(True)
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SumSellingP").Value
                'U_Quantity
                oEdit = oMatrix3.Columns.Item("212").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    'LoadHandingCharge_LCL_SE_SeaPrice
    Private Sub LoadHandingCharge_LCL_SE_SeaPrice(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            Dim DestCountry As String = ""
            Dim DestCity As String = ""
            oEdit = oForm.Items.Item("e13").Specific
            DestCountry = oEdit.String
            oEdit = oForm.Items.Item("ce13").Specific
            DestCity = oEdit.String
            oCombo = oForm.Items.Item("cce3").Specific
            Dim ContType As String = ""
            Try
                ContType = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            If ContType = "" Then
                ContType = "%"
            End If
            Dim Str As String = ""
            Str = "SELECT T0.[U_ChargeCode], T0.[U_Description],T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Quantity], T1.[U_Unit], T1.[U_Cost], T1.[U_SellingP] FROM [dbo].[@AB_SEA_FREIGHT]  T0 , [dbo].[@AB_SEA_FREIGHT_LCL]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='" & Division & "' and  T1.[U_City] ='" & DestCity & "' and  T0.[U_Country] ='" & DestCountry & "'" + " Union All " + "SELECT T1.[U_ChargeCode], T1.[U_ChargeDesc], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Quantity], T1.[U_Unit], T1.[U_Cost], T1.[U_SellingP] FROM [dbo].[@AB_SEA_HC]  T0 , [dbo].[@AB_SEA_HCLCL]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Division] ='SE'"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery(Str)
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oMatrix3.AddRow()
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
                'U_Quantity
                oEdit = oMatrix3.Columns.Item("212").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_FCL_SE_SeaPrice(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            Dim DestCountry As String = ""
            Dim DestCity As String = ""
            oEdit = oForm.Items.Item("e13").Specific
            DestCountry = oEdit.String
            oEdit = oForm.Items.Item("ce13").Specific
            DestCity = oEdit.String
            oCombo = oForm.Items.Item("cce3").Specific
            Dim ContType As String = ""
            Try
                ContType = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            If ContType = "" Then
                SBO_Application.StatusBar.SetText("Select Container Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim Str As String = ""
            Str = "SELECT T0.[U_ChargeCode], T0.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Quantity], T1.[U_Unit], T1.[U_Cost], T1.[U_SumSelllingP] as 'U_SellingP' FROM [dbo].[@AB_SEA_FREIGHT]  T0 , [dbo].[@AB_SEA_FREIGHT_FCL]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Division] ='SE' and  T0.[U_Country] ='" & DestCountry & "' and  T1.[U_City] ='" & DestCity & "' and  T1.[U_ContainerType] ='" & ContType & "'" + " Union All " + "SELECT T1.[U_ChargeCode], T1.[U_ChargeDesc], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Quantity], T1.[U_Unit], T1.[U_Cost], T1.[U_SellingP] FROM [dbo].[@AB_SEA_HC]  T0 , [dbo].[@AB_SEA_HCFCL]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Division] ='SE'"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery(Str)
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            ' SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & oRecordSet1.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oMatrix3.AddRow()
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_SellingP").Value
                'U_Quantity
                oEdit = oMatrix3.Columns.Item("212").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Unit").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
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
#Region "Air Price Loading"
    Private Sub LoadHandingCharge_AirImport(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            Dim Min As Double = 0
            Dim Perkg As Double = 0
            Dim UnitPrice As Double = 0
            Dim UnitPrice1 As Double = 0
            Dim cost As Double = 0
            Dim netcost As Double = 0
            Dim wt As Double = 0
            Dim CustCode As String = ""
            Dim i As Integer = 0
            oEdit = oForm.Items.Item("4").Specific
            CustCode = oEdit.Value
            CustCode = CustCode & "_AI" '.Replace("_AI", "").Trim()
            oEdit = oForm.Items.Item("cce3").Specific
            wt = oEdit.Value
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Min], T1.[U_PerKg], T1.[U_Cost] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Divsion] ='AI' and  T0.[Code] ='" & CustCode & "'")
            If oRecordSet1.RecordCount <> 0 Then
                oMatrix3 = oForm.Items.Item("38").Specific
                oColumns = oMatrix3.Columns
                oMatrix3.Clear()
                ' oForm.Freeze(True)
                For i = 1 To oRecordSet1.RecordCount
                    oMatrix3.AddRow()
                    SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                    'oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                    'oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                    Min = oRecordSet1.Fields.Item("U_Min").Value
                    Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
                    If Perkg = 0 Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.String = Min
                    Else
                        UnitPrice1 = Perkg * wt
                        If UnitPrice1 < Min Then
                            oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                            oEdit.String = 1
                            oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                            oEdit.String = Min
                        Else
                            oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                            oEdit.String = wt
                            oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                            oEdit.String = Perkg
                        End If
                      
                    End If
                    
                    cost = oRecordSet1.Fields.Item("U_Cost").Value
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.String = cost
                    oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                    oRecordSet1.MoveNext()
                Next
                ' oForm.Freeze(False)
                SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Exit Sub
            End If
            oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AI' and T0.[Code] ='AI'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            'oForm.Freeze(True)
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                'oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                Min = oRecordSet1.Fields.Item("U_Min").Value
                Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
                If Perkg = 0 Then
                    oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                    oEdit.String = 1
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.String = Min
                Else
                    UnitPrice1 = Perkg * wt
                    If UnitPrice1 < Min Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.String = Min
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = wt
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.String = Perkg
                    End If

                End If

                cost = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = cost
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            'oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_AirExport1(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            Dim Min As Double = 0
            Dim Perkg As Double = 0
            Dim UnitPrice As Double = 0
            Dim UnitPrice1 As Double = 0
            Dim cost As Double = 0
            Dim netcost As Double = 0
            Dim wt As Double = 0
            Dim CustCode As String = ""
            Dim i As Integer = 0
            Dim J As Integer = 0
            oEdit = oForm.Items.Item("4").Specific
            CustCode = oEdit.Value
            CustCode = CustCode & "_AE" '.Replace("_AE", "").Trim()
            oEdit = oForm.Items.Item("cce3").Specific
            wt = oEdit.Value
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Min], T1.[U_PerKg], T1.[U_Cost] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Divsion] ='AE' and  T0.[Code] ='" & CustCode & "'")
            If oRecordSet1.RecordCount <> 0 Then
                oMatrix3 = oForm.Items.Item("38").Specific
                oColumns = oMatrix3.Columns
                ' oMatrix3.Clear()
                ' oForm.Freeze(True)
                Dim Addrow As Boolean = False
                If oMatrix3.RowCount = 1 Then
                    Addrow = False
                Else
                    Addrow = True
                End If
                For J = 1 To oRecordSet1.RecordCount
                    oMatrix3.AddRow()
                    If Addrow = False Then
                        i = oMatrix3.RowCount
                    ElseIf Addrow = True Then
                        i = oMatrix3.RowCount - 1
                    End If
                    SBO_Application.StatusBar.SetText("Please Wait handling Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                    'oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                    'oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                    Min = oRecordSet1.Fields.Item("U_Min").Value
                    Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
                    If Perkg = 0 Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.String = Min
                    Else
                        UnitPrice1 = Perkg * wt
                        If UnitPrice1 < Min Then
                            oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                            oEdit.String = 1
                            oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                            oEdit.String = Min
                        Else
                            oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                            oEdit.String = wt
                            oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                            oEdit.String = Perkg
                        End If

                    End If

                    cost = oRecordSet1.Fields.Item("U_Cost").Value
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.String = cost
                    oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                    oRecordSet1.MoveNext()
                Next
                ' oForm.Freeze(False)
                SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Exit Sub
            End If
            oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AE' and T0.[Code] ='AE'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            'oForm.Freeze(True)
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                'oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                'oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                Min = oRecordSet1.Fields.Item("U_Min").Value
                Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
                If Perkg = 0 Then
                    oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                    oEdit.String = 1
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.String = Min
                Else
                    UnitPrice1 = Perkg * wt
                    If UnitPrice1 < Min Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.String = Min
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = wt
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.String = Perkg
                    End If

                End If

                cost = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = cost
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            'oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Private Sub LoadHandingCharge_AirExport_SPECIAL(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try

            Dim DestCity As String = ""
            Dim ServiceLevel As String = ""
            Dim Carrier As String = ""
            Dim Cargo As String = ""
            Dim Weight As Double = 0
            Dim MinAmt As Double = 0
            oEdit = oForm.Items.Item("ce13").Specific
            DestCity = oEdit.String
            oEdit = oForm.Items.Item("cce3").Specific
            Weight = oEdit.Value
            Try
                oCombo = oForm.Items.Item("ce1").Specific
                Cargo = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oEdit = oForm.Items.Item("e2c").Specific
                Carrier = oEdit.String
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("e11").Specific
                ServiceLevel = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            oEdit = oForm.Items.Item("4").Specific
            Dim BPCode As String = oEdit.String
            Dim Str As String = "SELECT T1.U_1000base,T1.[U_ChargeCode], T1.[U_Description], T1.U_VendorName,T1.[U_VendorCode], T1.[U_MinBase], T1.[U_Neg45Base], T1.[U_45Base], T1.[U_100Base], T1.[U_300base], T1.[U_500Base], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500],T1.U_Min,T1.U_1000 FROM [dbo].[@AB_AIRSPECIAL_H]  T0 , [dbo].[@AB_AIRSPECIAL_D]  T1 WHERE T1.[Code] = T0.[Code]  and T0.[U_Division] ='AE' and    T1.[U_Code] ='" & DestCity & "' and  T1.[U_Carrier] ='" & Carrier & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "' and  T1.[U_CargoType] ='" & Cargo & "' and T0.Code='" & BPCode & "'"
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery(Str)
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            '*************LOADING FREIGHT CHARGE*********************
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Freight Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
                If Weight < 45 Then
                    MinAmt = oRecordSet1.Fields.Item("U_Min").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_Neg45").Value * Weight) Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = MinAmt
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = Weight
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = oRecordSet1.Fields.Item("U_Neg45").Value
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45Base").Value * Weight

                ElseIf Weight >= 45 And Weight < 100 Then
                    MinAmt = oRecordSet1.Fields.Item("U_Min").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_45").Value * Weight) Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = MinAmt
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = Weight
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = oRecordSet1.Fields.Item("U_45").Value
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45Base").Value * Weight
                ElseIf Weight >= 100 And Weight < 300 Then
                    MinAmt = oRecordSet1.Fields.Item("U_Min").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_100").Value * Weight) Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = MinAmt
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = Weight
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = oRecordSet1.Fields.Item("U_100").Value

                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100Base").Value * Weight
                ElseIf Weight >= 300 And Weight < 500 Then
                    MinAmt = oRecordSet1.Fields.Item("U_Min").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_300").Value * Weight) Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = MinAmt
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = Weight
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = oRecordSet1.Fields.Item("U_300").Value

                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300base").Value * Weight
                ElseIf Weight >= 500 And Weight < 1000 Then
                    MinAmt = oRecordSet1.Fields.Item("U_Min").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_500").Value * Weight) Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = MinAmt
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = Weight
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = oRecordSet1.Fields.Item("U_500").Value

                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500base").Value * Weight
                ElseIf Weight >= 1000 Then
                    MinAmt = oRecordSet1.Fields.Item("U_Min").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_1000").Value * Weight) Then
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = 1
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = MinAmt
                    Else
                        oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                        oEdit.String = Weight
                        oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                        oEdit.Value = oRecordSet1.Fields.Item("U_1000").Value

                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_1000base").Value * Weight
                End If

                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            '*************LOADING HANDLING CHARGE*********************
            LoadHandingCharge_AirExport1(oForm, Division)
        Catch ex As Exception

        End Try
    End Sub
    'Private Sub LoadHandingCharge_AirExport_SP(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
    '    Try
    '        ' Dim DestCountry As String = ""
    '        Dim DestCity As String = ""
    '        Dim ServiceLevel As String = ""
    '        Dim Carrier As String = ""
    '        Dim Cargo As String = ""
    '        Dim Weight As Double = 0
    '        Dim MinAmt As Double = 0
    '        'oEdit = oForm.Items.Item("e13").Specific
    '        'DestCountry = oEdit.String
    '        oEdit = oForm.Items.Item("ce13").Specific
    '        DestCity = oEdit.String
    '        oEdit = oForm.Items.Item("cce3").Specific
    '        Weight = oEdit.Value
    '        Try
    '            oCombo = oForm.Items.Item("ce1").Specific
    '            Cargo = oCombo.Selected.Value
    '        Catch ex As Exception
    '        End Try
    '        Try
    '            oCombo = oForm.Items.Item("e2").Specific
    '            Carrier = oCombo.Selected.Value
    '        Catch ex As Exception
    '        End Try
    '        Try
    '            oCombo = oForm.Items.Item("e11").Specific
    '            ServiceLevel = oCombo.Selected.Value
    '        Catch ex As Exception
    '        End Try
    '        oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim Str As String = "SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode],T1.[U_VendorName], T1.[U_MinBase], T1.[U_Neg45Base], T1.[U_45Base], T1.[U_100Base], T1.[U_300base], T1.[U_500Base], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500] FROM [dbo].[@AB_AIRSPECIAL_H]  T0 , [dbo].[@AB_AIRSPECIAL_D]  T1 WHERE T1.[Code] = T0.[Code]  and T0.[U_Division] ='AE' and   T1.[U_Code] ='" & DestCity & "' and  T1.[U_Carrier] ='" & Carrier & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "' and  T1.[U_CargoType] ='" & Cargo & "'"
    '        oRecordSet1.DoQuery(Str)
    '        If oRecordSet1.RecordCount = 0 Then
    '            SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If
    '        oMatrix3 = oForm.Items.Item("38").Specific
    '        oColumns = oMatrix3.Columns
    '        oMatrix3.Clear()
    '        '  SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '        ' oForm.Freeze(True)
    '        Dim i As Integer = 0
    '        For i = 1 To oRecordSet1.RecordCount
    '            oMatrix3.AddRow()
    '            SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
    '            oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Description").Value
    '            oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
    '            oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
    '            If Weight < 45 Then
    '                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_Neg45").Value * Weight
    '                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_Neg45Base").Value * Weight
    '                MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
    '                If MinAmt > (oRecordSet1.Fields.Item("U_Neg45Base").Value * Weight) Then
    '                    oEdit.Value = MinAmt
    '                End If
    '            ElseIf Weight >= 45 And Weight < 100 Then
    '                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_45").Value * Weight
    '                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_45Base").Value * Weight
    '                MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
    '                If MinAmt > (oRecordSet1.Fields.Item("U_45Base").Value * Weight) Then
    '                    oEdit.Value = MinAmt
    '                End If
    '            ElseIf Weight >= 100 And Weight < 300 Then
    '                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_100").Value * Weight
    '                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_100Base").Value * Weight
    '                MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
    '                If MinAmt > (oRecordSet1.Fields.Item("U_100Base").Value * Weight) Then
    '                    oEdit.Value = MinAmt
    '                End If
    '            ElseIf Weight >= 300 And Weight < 500 Then
    '                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_300").Value * Weight
    '                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_300base").Value * Weight
    '                MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
    '                If MinAmt > (oRecordSet1.Fields.Item("U_300base").Value * Weight) Then
    '                    oEdit.Value = MinAmt
    '                End If
    '            ElseIf Weight >= 500 Then
    '                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_500").Value * Weight
    '                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
    '                oEdit.Value = oRecordSet1.Fields.Item("U_500Base").Value * Weight
    '                MinAmt = oRecordSet1.Fields.Item("U_MinBase").Value
    '                If MinAmt > (oRecordSet1.Fields.Item("U_500Base").Value * Weight) Then
    '                    oEdit.Value = MinAmt
    '                End If
    '            End If

    '            oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
    '            oRecordSet1.MoveNext()
    '        Next

    '        '********************HANDLING CHARGES*************************
    '        Try

    '            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AE'")
    '            If oRecordSet1.RecordCount = 0 Then
    '                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Exit Sub
    '            End If
    '            oMatrix3 = oForm.Items.Item("38").Specific
    '            oColumns = oMatrix3.Columns
    '            ' oMatrix3.Clear()
    '            SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            ' oForm.Freeze(True)
    '            Dim Min As Double = 0
    '            Dim Perkg As Double = 0
    '            Dim UnitPrice As Double = 0
    '            Dim wt As Double = 0
    '            Dim cost As Double = 0
    '            Dim netcost As Double = 0
    '            oEdit = oForm.Items.Item("cce3").Specific
    '            wt = oEdit.Value
    '            Dim K As Integer = 0
    '            For i = 1 To oRecordSet1.RecordCount
    '                K = oMatrix3.RowCount()
    '                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                If oMatrix3.RowCount = 0 Then
    '                    oMatrix3.AddRow()
    '                    oMatrix3.ClearRowData(oMatrix3.RowCount())
    '                Else
    '                    oEdit = oMatrix3.Columns.Item("1").Cells.Item(K).Specific
    '                    If oEdit.String <> "" Then
    '                        oMatrix3.AddRow()
    '                        oMatrix3.ClearRowData(oMatrix3.RowCount())
    '                    End If
    '                End If
    '                K = oMatrix3.RowCount()
    '                oEdit = oMatrix3.Columns.Item("1").Cells.Item(K).Specific
    '                oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
    '                oEdit = oMatrix3.Columns.Item("3").Cells.Item(K).Specific
    '                oEdit.String = oRecordSet1.Fields.Item("U_Description").Value

    '                'oEdit = oMatrix3.Columns.Item("11").Cells.Item(oMatrix3.RowCount()).Specific
    '                'oEdit.String = "1" 'oRecordSet1.Fields.Item("U_Description").Value

    '                Min = oRecordSet1.Fields.Item("U_Min").Value
    '                Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
    '                If Perkg = 0 Then
    '                    UnitPrice = Min
    '                Else
    '                    UnitPrice = Perkg * wt
    '                    If UnitPrice < Min Then
    '                        UnitPrice = Min
    '                    End If
    '                End If
    '                oMatrix3.DeleteRow(oMatrix3.RowCount())
    '                oEdit = oMatrix3.Columns.Item("14").Cells.Item(K).Specific
    '                oEdit.String = UnitPrice
    '                cost = oRecordSet1.Fields.Item("U_Cost").Value

    '                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(K).Specific
    '                oEdit.String = cost
    '                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(K).Specific
    '                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
    '                'oMatrix3.DeleteRow(oMatrix3.RowCount())

    '                oRecordSet1.MoveNext()

    '            Next
    '            oMatrix3.AddRow()
    '        Catch ex As Exception
    '            oForm.Freeze(False)
    '        End Try
    '        '********************END HANDLING CHARGES*************************
    '        oForm.Freeze(False)
    '        SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Catch ex As Exception
    '        oForm.Freeze(False)
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
    Private Sub LoadHandingCharge_AirExport(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            '  Dim DestCountry As String = ""
            Dim DestCity As String = ""
            Dim ServiceLevel As String = ""
            Dim Carrier As String = ""
            Dim Cargo As String = ""
            Dim Weight As Double = 0
            Dim MinAmt As Double = 0
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            oEdit = oForm.Items.Item("ce13").Specific
            DestCity = oEdit.String
            oEdit = oForm.Items.Item("cce3").Specific
            Weight = oEdit.Value
            Try
                oCombo = oForm.Items.Item("ce1").Specific
                Cargo = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("e2").Specific
                Carrier = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            Try
                oCombo = oForm.Items.Item("e11").Specific
                ServiceLevel = oCombo.Selected.Value
            Catch ex As Exception
            End Try
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Str As String = "SELECT T0.[Code], T0.[Name], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_MIN], T1.[U_Neg45], T1.[U_45], T1.[U_100], T1.[U_300], T1.[U_500], T1.[U_Neg45Cost], T1.[U_45Cost], T1.[U_100Cost], T1.[U_300Cost], T1.[U_500Cost] FROM [dbo].[@AB_AIR_AIRFREIGHTH]  T0 , [dbo].[@AB_AIR_AIRFREIGHTD]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='AE' and  T1.[U_CargoType] ='" & Cargo & "' and  T1.[U_Carrier] ='" & Carrier & "'  and  T1.[U_Code] ='" & DestCity & "' and  T1.[U_SrvLevel] ='" & ServiceLevel & "'"
            oRecordSet1.DoQuery(Str)
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            '  SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'oEdit = oMatrix3.Columns.Item("U_Ven").Cells.Item(1).Specific
                'oEdit.Value = "ABCT" 'oRecordSet1.Fields.Item("U_VendorCode").Value.ToString.Trim
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("Code").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("Name").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = 1 'oRecordSet1.Fields.Item("U_Quantity").Value
                If Weight < 45 Then
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_Neg45").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_Neg45Cost").Value * Weight
                ElseIf Weight >= 45 And Weight < 100 Then
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_45").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_45Cost").Value * Weight
                ElseIf Weight >= 100 And Weight < 300 Then
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_100").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_100Cost").Value * Weight
                ElseIf Weight >= 300 And Weight < 500 Then
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_300").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_300Cost").Value * Weight
                ElseIf Weight >= 500 Then
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500").Value * Weight
                    MinAmt = oRecordSet1.Fields.Item("U_MIN").Value
                    If MinAmt > (oRecordSet1.Fields.Item("U_500").Value * Weight) Then
                        oEdit.Value = MinAmt
                    End If
                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                    oEdit.Value = oRecordSet1.Fields.Item("U_500Cost").Value * Weight
                End If

                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            '********************HANDLING CHARGES*************************
            Try

                oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AE'")
                If oRecordSet1.RecordCount = 0 Then
                    SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                oMatrix3 = oForm.Items.Item("38").Specific
                oColumns = oMatrix3.Columns
                ' oMatrix3.Clear()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                ' oForm.Freeze(True)
                Dim Min As Double = 0
                Dim Perkg As Double = 0
                Dim UnitPrice As Double = 0
                Dim wt As Double = 0
                Dim cost As Double = 0
                Dim netcost As Double = 0
                oEdit = oForm.Items.Item("cce3").Specific
                wt = oEdit.Value
                Dim K As Integer = 0
                For i = 1 To oRecordSet1.RecordCount
                    K = oMatrix3.RowCount()
                    SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If oMatrix3.RowCount = 0 Then
                        oMatrix3.AddRow()
                        oMatrix3.ClearRowData(oMatrix3.RowCount())
                    Else
                        oEdit = oMatrix3.Columns.Item("1").Cells.Item(K).Specific
                        If oEdit.String <> "" Then
                            oMatrix3.AddRow()
                            oMatrix3.ClearRowData(oMatrix3.RowCount())
                        End If
                    End If
                    K = oMatrix3.RowCount()
                    oEdit = oMatrix3.Columns.Item("1").Cells.Item(K).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
                    oEdit = oMatrix3.Columns.Item("3").Cells.Item(K).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_Description").Value

                    'oEdit = oMatrix3.Columns.Item("11").Cells.Item(oMatrix3.RowCount()).Specific
                    'oEdit.String = "1" 'oRecordSet1.Fields.Item("U_Description").Value

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
                    oMatrix3.DeleteRow(oMatrix3.RowCount())
                    oEdit = oMatrix3.Columns.Item("14").Cells.Item(K).Specific
                    oEdit.String = UnitPrice
                    cost = oRecordSet1.Fields.Item("U_Cost").Value

                    oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(K).Specific
                    oEdit.String = cost
                    oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(K).Specific
                    oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                    'oMatrix3.DeleteRow(oMatrix3.RowCount())

                    oRecordSet1.MoveNext()

                Next
                oMatrix3.AddRow()
            Catch ex As Exception
                oForm.Freeze(False)
            End Try
            '********************END HANDLING CHARGES*************************
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    'Private Sub LoadHandingCharge_AirImport_only(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
    '    Try
    '        Exit Sub
    '        Dim DestCountry As String = ""
    '        Dim DestCity As String = ""
    '        oEdit = oForm.Items.Item("ce13").Specific
    '        DestCity = oEdit.String

    '        oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecordSet1.DoQuery("SELECT T1.[U_ChargeCode], T1.[U_Description], T1.[U_VendorCode], T1.[U_VendorName], T1.[U_Cost], T1.[U_Min], T1.[U_PerKg] FROM [dbo].[@AB_AIR_HC]  T0 , [dbo].[@AB_AIR_HCDETAILS]  T1 WHERE T1.[Code] = T0.[Code]  and  T0.[U_Divsion] ='AE'")
    '        If oRecordSet1.RecordCount = 0 Then
    '            SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If
    '        oMatrix3 = oForm.Items.Item("38").Specific
    '        oColumns = oMatrix3.Columns

    '        ' oForm.Freeze(True)
    '        Dim Min As Double = 0
    '        Dim Perkg As Double = 0
    '        Dim UnitPrice As Double = 0
    '        Dim cost As Double = 0
    '        Dim netcost As Double = 0
    '        Dim wt As Double = 0
    '        Dim i As Integer = 0
    '        For i = 1 To oRecordSet1.RecordCount
    '            SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            If oMatrix3.RowCount = 0 Then
    '                oMatrix3.AddRow()
    '                oMatrix3.ClearRowData(oMatrix3.RowCount())
    '            Else
    '                oEdit = oMatrix3.Columns.Item("1").Cells.Item(oMatrix3.RowCount()).Specific
    '                If oEdit.String <> "" Then
    '                    oMatrix3.AddRow()
    '                    oMatrix3.ClearRowData(oMatrix3.RowCount())
    '                End If
    '            End If

    '            oEdit = oMatrix3.Columns.Item("1").Cells.Item(oMatrix3.RowCount()).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_ChargeCode").Value
    '            oEdit = oMatrix3.Columns.Item("3").Cells.Item(oMatrix3.RowCount()).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_Description").Value

    '            'oEdit = oMatrix3.Columns.Item("11").Cells.Item(oMatrix3.RowCount()).Specific
    '            'oEdit.String = "1" 'oRecordSet1.Fields.Item("U_Description").Value

    '            Min = oRecordSet1.Fields.Item("U_Min").Value
    '            Perkg = oRecordSet1.Fields.Item("U_PerKg").Value
    '            If Perkg = 0 Then
    '                UnitPrice = Min
    '            Else
    '                UnitPrice = Perkg * wt
    '                If UnitPrice < Min Then
    '                    UnitPrice = Min
    '                End If
    '            End If
    '            oEdit = oMatrix3.Columns.Item("14").Cells.Item(oMatrix3.RowCount()).Specific
    '            oEdit.String = UnitPrice
    '            cost = oRecordSet1.Fields.Item("U_Cost").Value

    '            oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(oMatrix3.RowCount()).Specific
    '            oEdit.String = cost
    '            oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(oMatrix3.RowCount()).Specific
    '            oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
    '            oMatrix3.DeleteRow(oMatrix3.RowCount())

    '            oRecordSet1.MoveNext()

    '            '
    '            'oMatrix3.AddRow()

    '        Next
    '        oForm.Freeze(False)
    '        SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Catch ex As Exception
    '        oForm.Freeze(False)
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
#End Region
#Region "Local"
    Private Sub LoadHandingCharge_LocalCharges(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            ' Dim DestCountry As String = ""
            Dim DestCity As String = ""
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            'oEdit = oForm.Items.Item("ce13").Specific
            'DestCity = oEdit.String
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_Quantity, T1.U_Unit, T1.U_Cost, T1.U_Sprice, T1.U_Remarks, T1.U_ItemCode, T1.U_ItemDesc FROM [dbo].[@AB_LOCAL_PRICEH]  T0 , [dbo].[@AB_LOCAL_PRICED]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='LC'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim Min As Double = 0
            Dim Perkg As Double = 0
            Dim UnitPrice As Double = 0
            Dim cost As Double = 0
            Dim netcost As Double = 0
            Dim wt As Double = 0
            oEdit = oForm.Items.Item("cce3").Specific
            wt = oEdit.Value
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No-" & oRecordSet1.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Sprice").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_LocalCharges_SpecialPrice(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String, ByVal Customer As String)
        Try
            ' Dim DestCountry As String = ""
            Dim DestCity As String = ""
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            'oEdit = oForm.Items.Item("ce13").Specific
            'DestCity = oEdit.String

            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_ItemCode, T1.U_ItemDesc, T1.U_Quantity, T1.U_Cost, T1.U_SPrice, T1.U_MarkUp, T1.U_MarkedUpPrice, T1.U_Remarks FROM [dbo].[@AB_LOCAL_SPRICEH]  T0 , [dbo].[@AB_LOCAL_SPRICED]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='LC' and  T1.[Code] ='" & Customer & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim Min As Double = 0
            Dim Perkg As Double = 0
            Dim UnitPrice As Double = 0
            Dim cost As Double = 0
            Dim netcost As Double = 0
            Dim wt As Double = 0
            oEdit = oForm.Items.Item("cce3").Specific
            wt = oEdit.Value
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No-" & oRecordSet1.RecordCount & " of -" & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Sprice").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
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
#Region "International"
    Private Sub LoadHandingCharge_International(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String)
        Try
            ' Dim DestCountry As String = ""
            Dim DestCity As String = ""
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            'oEdit = oForm.Items.Item("ce13").Specific
            'DestCity = oEdit.String
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_Quantity, T1.U_Unit, T1.U_Cost, T1.U_Sprice, T1.U_Remarks, T1.U_ItemCode, T1.U_ItemDesc FROM [dbo].[@AB_LOCAL_PRICEH]  T0 , [dbo].[@AB_LOCAL_PRICED]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='LC'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim Min As Double = 0
            Dim Perkg As Double = 0
            Dim UnitPrice As Double = 0
            Dim cost As Double = 0
            Dim netcost As Double = 0
            Dim wt As Double = 0
            oEdit = oForm.Items.Item("cce3").Specific
            wt = oEdit.Value
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Sprice").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
                oRecordSet1.MoveNext()
            Next
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("Price Loading Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub LoadHandingCharge_International_SpecialPrice(ByVal oForm As SAPbouiCOM.Form, ByVal Division As String, ByVal Customer As String)
        Try
            ' Dim DestCountry As String = ""
            Dim DestCity As String = ""
            'oEdit = oForm.Items.Item("e13").Specific
            'DestCountry = oEdit.String
            'oEdit = oForm.Items.Item("ce13").Specific
            'DestCity = oEdit.String

            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T1.U_VendorCode, T1.U_VendorName, T1.U_ItemCode, T1.U_ItemDesc, T1.U_Quantity, T1.U_Cost, T1.U_SPrice, T1.U_MarkUp, T1.U_MarkedUpPrice, T1.U_Remarks FROM [dbo].[@AB_LOCAL_SPRICEH]  T0 , [dbo].[@AB_LOCAL_SPRICED]  T1 WHERE T1.[Code] = T0.[Code] and  T0.[U_Division] ='LC' and  T1.[Code] ='" & Customer & "'")
            If oRecordSet1.RecordCount = 0 Then
                SBO_Application.StatusBar.SetText("No Data Found!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oMatrix3 = oForm.Items.Item("38").Specific
            oColumns = oMatrix3.Columns
            oMatrix3.Clear()
            Dim Min As Double = 0
            Dim Perkg As Double = 0
            Dim UnitPrice As Double = 0
            Dim cost As Double = 0
            Dim netcost As Double = 0
            Dim wt As Double = 0
            oEdit = oForm.Items.Item("cce3").Specific
            wt = oEdit.Value
            oForm.Freeze(True)
            Dim i As Integer = 0
            For i = 1 To oRecordSet1.RecordCount
                oMatrix3.AddRow()
                SBO_Application.StatusBar.SetText("Please Wait Price is Loading! Line No " & i & " of " & oRecordSet1.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oEdit = oMatrix3.Columns.Item("1").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemCode").Value
                oEdit = oMatrix3.Columns.Item("3").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_ItemDesc").Value
                oEdit = oMatrix3.Columns.Item("11").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Quantity").Value
                oEdit = oMatrix3.Columns.Item("14").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Sprice").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Cost").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_Cost").Value
                oEdit = oMatrix3.Columns.Item("U_AB_Vendor").Cells.Item(i).Specific
                oEdit.String = oRecordSet1.Fields.Item("U_VendorName").Value.ToString
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
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If (pVal.MenuUID = "1282" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.InnerEvent = False And pVal.BeforeAction = False Then
                oForm = SBO_Application.Forms.ActiveForm
                'SBO_Application.Forms.GetFormByTypeAndCount(139, 1).UniqueID
                Dim oform1 As SAPbouiCOM.Form
                oform1 = SBO_Application.Forms.GetFormByTypeAndCount(139, 1)
                If oForm.UniqueID = oform1.UniqueID Then
                    Try
                        oForm = SBO_Application.Forms.ActiveForm
                        oMatrix = oForm.Items.Item("MATCargo").Specific
                        oMatrix.AddRow()
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        oEdit = oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific
                        oEdit.String = ""
                        oMatrix1 = oForm.Items.Item("MATCont").Specific
                        oMatrix1.AddRow()
                        oMatrix1.ClearRowData(oMatrix1.RowCount)
                        oEdit = oMatrix1.Columns.Item("#").Cells.Item(oMatrix1.RowCount).Specific
                        oEdit.String = ""

                        oItem = oForm.Items.Item("e9")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("e10")
                        oItem.Enabled = False
                        oItem = oForm.Items.Item("e12")
                        oItem.Enabled = False

                        ooption = oForm.Items.Item("ce2").Specific
                        If ooption.Selected = False Then
                            oItem = oForm.Items.Item("cce3")
                            oItem.Enabled = False
                        End If
                        
                    Catch ex As Exception
                    End Try
                End If
            End If


            If Not pVal.BeforeAction Then
                Try
                    oForm = SBO_Application.Forms.ActiveForm
                    Dim matrix As SAPbouiCOM.Matrix
                    If pVal.MenuUID = "DeleteRow" Then
                        matrix = oForm.Items.Item(matrixUID).Specific
                        If rowDelete <> 0 And rowDelete <> matrix.RowCount Then
                            matrix.DeleteRow(rowDelete)
                            rowDelete = 0
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If
                        End If
                    ElseIf pVal.MenuUID = "ClearMatrix" Then
                        matrix = oForm.Items.Item(matrixUID).Specific
                        matrix.Clear()
                        matrix.AddRow(1, 0)
                        matrix.FlushToDataSource()
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If

                Catch ex As Exception
                    SBO_Application.MessageBox(ex.Message)

                End Try
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
