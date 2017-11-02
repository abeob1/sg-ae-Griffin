Option Strict Off
Option Explicit On
Imports System.Diagnostics.Process
Imports System.Threading
Imports System.Net.Mail
Public Class Connection
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public ocompany As New SAPbobsCOM.Company
    Public SboGuiApi As New SAPbouiCOM.SboGuiApi
    Public sConnectionString As String
    Dim oF_SI_JobOrder As F_SI_JobOrder
    Dim oF_SalesOrder As F_SalesOrder
    Dim oGoodsReceipt As F_GoodsReceipt
    Dim oGoodsIssue As F_GoodsIssue
    Dim oGoodsRelease As F_GoodsRelease
    Dim oF_SE_JobOrder As F_SE_JobOrder
    Dim oF_SeaExport_SpecialPrice As F_SeaExport_SpecialPrice
    Dim oF_AI_JobOrder As F_AI_JobOrder
    Dim oF_AE_JobOrder As F_AE_JobOrder
    Dim oF_AWB As F_AWB
    Dim oPrepareAWB As PrepareAWB
    Dim oF_OBL As F_OBL
    Dim oPaymentVoucher As PaymentVoucher
    Dim oF_Loacl As F_Loacl
    Dim oInternational As International
    Dim oF_Project As F_Project
    Dim oF_JobOrder As F_JobOrder
    Dim oF_Invoice As F_Invoice
    Public Sub conn2()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Try
            sconn = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            SBO_Application = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = ocompany.GetContextCookie
            str = SBO_Application.Company.GetConnectionContext(scook)
            ret = ocompany.SetSboLoginContext(str)
            ocompany.Connect()
            ocompany.GetLastError(ret, str)
            If ret <> 0 Then
                SBO_Application.MessageBox("SAP Connection Failed :" & str)
            End If
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:conn2" + " Error Message:" + ex.ToString)
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
#Region "Connection"
    Public Sub New()
        MyBase.new()
        Try
            conn2()
            LoadFromXML("MyMenus.xml", SBO_Application)
            oF_SI_JobOrder = New F_SI_JobOrder(ocompany, SBO_Application)
            oF_SalesOrder = New F_SalesOrder(ocompany, SBO_Application)
            oGoodsReceipt = New F_GoodsReceipt(ocompany, SBO_Application)
            oGoodsIssue = New F_GoodsIssue(ocompany, SBO_Application)
            oGoodsRelease = New F_GoodsRelease(ocompany, SBO_Application)
            oF_SE_JobOrder = New F_SE_JobOrder(ocompany, SBO_Application)
            oF_SeaExport_SpecialPrice = New F_SeaExport_SpecialPrice(ocompany, SBO_Application)
            oF_AI_JobOrder = New F_AI_JobOrder(ocompany, SBO_Application)
            oF_AE_JobOrder = New F_AE_JobOrder(ocompany, SBO_Application)
            oF_AWB = New F_AWB(ocompany, SBO_Application)
            oF_OBL = New F_OBL(ocompany, SBO_Application)
            oPaymentVoucher = New PaymentVoucher(ocompany, SBO_Application)
            oF_Loacl = New F_Loacl(ocompany, SBO_Application)
            oInternational = New International(ocompany, SBO_Application)
            oF_Project = New F_Project(ocompany, SBO_Application)
            oF_JobOrder = New F_JobOrder(ocompany, SBO_Application)
            oF_Invoice = New F_Invoice(ocompany, SBO_Application)
            Try
                oMenuItem = SBO_Application.Menus.Item("FMMyMenu01") 'moudles'
                Dim sPath As String
                sPath = IO.Directory.GetParent(Application.StartupPath).ToString
                oMenuItem.Image = sPath & "\GK_FM\" & "WHMS.jpg"
            Catch ex As Exception
                '   SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
            SBO_Application.MessageBox("Welcome To FM...")
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:New" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub conn1()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Try
            sconn = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            SBO_Application = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = ocompany.GetContextCookie
            str = SBO_Application.Company.GetConnectionContext(scook)
            ret = ocompany.SetSboLoginContext(str)
            ocompany.Connect()
            ocompany.GetLastError(ret, str)
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:conn1" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Try
            If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()
            End If

            If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()
            End If

            If EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()

            End If

        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:SBO_Application_AppEvent" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = True Then
                If pVal.MenuUID = "FMMySubMenu01" Then
                    LoadFromXML("SEA_IMPORT_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("SEAI_JOB")
                    oF_SI_JobOrder.SI_Job_Bind(oForm)
                ElseIf pVal.MenuUID = "FMMySubMenu03" Then
                    LoadFromXML("SEA_EXPORT_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("SEAE_JOB")
                    oF_SE_JobOrder.SE_Job_Bind(oForm)
                ElseIf pVal.MenuUID = "FMMySubMenu02" Then
                    Type = "Sea"
                    SBO_Application.ActivateMenuItem("2050")

                ElseIf pVal.MenuUID = "MySubMenu01" Then
                    LoadFromXML("GoodsReceipt.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsReceipt")
                    oGoodsReceipt.GoodsReceipt_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu02" Then
                    LoadFromXML("GoodsIssue.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsIssue")
                    oGoodsIssue.GoodsIssue_Bind(oForm, SBO_Application)
                ElseIf pVal.MenuUID = "MySubMenu03" Then
                    LoadFromXML("GoodsRelease.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("AI_FI_GoodsRelease")
                    oGoodsRelease.GoodsRelease_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu05" Then
                    LoadFromXML("HoldingStockreport.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("HOLD_STOCK")
                    oForm.DataSources.UserDataSources.Add("oedit1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oEdit = oForm.Items.Item("HR4").Specific
                    oEdit.DataBind.SetBound(True, "", "oedit1")
                    oForm.DataSources.UserDataSources.Add("oedit2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oEdit = oForm.Items.Item("HR6").Specific
                    oEdit.DataBind.SetBound(True, "", "oedit2")
                ElseIf pVal.MenuUID = "MySubMenu10" Then
                    LoadFromXML("DOReport.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("DO_Report")
                    oForm.DataSources.UserDataSources.Add("oedit1", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                    oEdit = oForm.Items.Item("DO4").Specific
                    oEdit.DataBind.SetBound(True, "", "oedit1")
                ElseIf pVal.MenuUID = "MySubMenu06" Then
                    LoadFromXML("ROReport.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("RO_Report")
                    oForm.DataSources.UserDataSources.Add("oedit1", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                    oEdit = oForm.Items.Item("RO4").Specific
                    oEdit.DataBind.SetBound(True, "", "oedit1")
                    '------------------Air Menu Load---------------
                ElseIf pVal.MenuUID = "MySubMenu09" Then
                    Type = "Air"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu07" Then
                    LoadFromXML("AIR_IMPORT_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("AIRI_JOB")
                    oF_AI_JobOrder.AI_Job_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu08" Then
                    LoadFromXML("AIR_EXPORT_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("AIRE_JOB")
                    oF_AE_JobOrder.AE_Job_Bind(oForm)
                    '-----------------Local---------------
                ElseIf pVal.MenuUID = "MySubMenu11" Then
                    LoadFromXML("LOCAL_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("LOC_JOB")
                    oF_Loacl.Local_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu12" Then
                    LoadFromXML("INTERNATIONAL_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("INT_JOB")
                    oInternational.International_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu13" Then
                    Type = "Local"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu14" Then
                    Type = "International"
                    SubType = "Sea"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu15" Then
                    Type = "International"
                    SubType = "Air"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu16" Then
                    Type = "International"
                    SubType = "Local"
                    SBO_Application.ActivateMenuItem("2050")

                    '---------------Project
                ElseIf pVal.MenuUID = "MySubMenu21" Then
                    LoadFromXML("PROJECT_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("PRO_JOB")
                    oF_Project.Project_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu18" Then
                    Type = "Project"
                    SubType = "Sea"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu19" Then
                    Type = "Project"
                    SubType = "Air"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu20" Then
                    Type = "Project"
                    SubType = "Local"
                    SBO_Application.ActivateMenuItem("2050")
                ElseIf pVal.MenuUID = "MySubMenu22" Then
                    LoadFromXML("WHSC_JOBORDER.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("WHSC_JOB")
                    oF_JobOrder.WHSC_Bind(oForm)
                ElseIf pVal.MenuUID = "MySubMenu23" Then
                    LoadFromXML("InvNew.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("InvoiceNew")
                    oF_Invoice.Form_Bind(oForm)
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Functions.WriteLog("Class:Connection" + " Function:SBO_Application_MenuEvent" + " Error Message:" + ex.ToString)
        End Try
    End Sub
End Class
