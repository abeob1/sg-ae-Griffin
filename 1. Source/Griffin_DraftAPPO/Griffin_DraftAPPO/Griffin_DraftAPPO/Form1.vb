Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.IO
Public Class Form1
    Dim connect As New Connection()

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try

            If Connection.bConnect = False Then
                connect.setDB()
                If Not connect.connectDB() Then
                    'MsgBox("SAP Connection Failed")
                    Exit Sub
                Else
                    Create_APInvoice()
                    Create_PO()

                End If
            Else
                Create_APInvoice()
                Create_PO()
            End If
            PublicVariable.oCompany.Disconnect()
            PublicVariable.oCompany = Nothing
            GC.Collect()
            Environment.Exit(0)
        Catch ex As Exception
            connect.WriteLog(ex.Message)
            Environment.Exit(0)
        End Try

    End Sub


    Public Sub Create_APInvoice()
        Try
            Dim oPOR As SAPbobsCOM.Documents
            oPOR = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            Dim i As Integer = 0
            Dim K As Integer = 0
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet1 As SAPbobsCOM.Recordset
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet21 As SAPbobsCOM.Recordset
            oRecordSet21 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim J As Integer = 0
            Dim NewDocNum As String = ""
            Dim DocNum As String = ""
            Dim JobNo As String = ""
            Dim PVDocNum As String = ""
            oRecordSet1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[DocEntry],T1.[LineId],T0.[U_VODt],T0.[U_VOVCode], T0.[U_JobNo], T1.[U_ICode], T1.[U_IName],isnull(T1.[U_Qty],1) U_Qty, T1.[U_Amt] FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 , [dbo].[@AB_PAYMENTVOUCHER_L]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  isnull(T1.[U_ICode] ,'') <> '' and isnull(T0.U_APSt,'') <> 'Success'")
            For i = 1 To oRecordSet.RecordCount
                Try
                    oPOR = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    oPOR.DocDate = oRecordSet.Fields.Item("U_VODt").Value
                    DocNum = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    PVDocNum = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    oPOR.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                    oPOR.DocDueDate = oRecordSet.Fields.Item("U_VODt").Value
                    oPOR.TaxDate = oRecordSet.Fields.Item("U_VODt").Value
                    If oRecordSet.Fields.Item("U_VOVCode").Value.ToString.Length > 29 Then
                        oPOR.CardCode = oRecordSet.Fields.Item("U_VOVCode").Value.ToString.Substring(0, 30)
                    Else
                        oPOR.CardCode = oRecordSet.Fields.Item("U_VOVCode").Value.ToString
                    End If
                    oPOR.UserFields.Fields.Item("U_DocEntry").Value = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    'oPOR.UserFields.Fields.Item("U_LineNum").Value = oRecordSet.Fields.Item("LineNum").Value.ToString
                    oPOR.UserFields.Fields.Item("U_AB_JobNo").Value = oRecordSet.Fields.Item("U_JobNo").Value.ToString
                    JobNo = oRecordSet.Fields.Item("U_JobNo").Value.ToString
                    If oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 2) = "SI" Then
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],T0.U_VessName [U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_VessVoyage] [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_ItemDesc],T0.[U_GrssWt],T0.U_ETA FROM [dbo].[@AB_SEAI_JOB_H] T0   WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "ZI"
                            oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value
                            oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotPkg").Value = oRecordSet2.Fields.Item("U_TotPkg").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet2.Fields.Item("U_ChrgWt").Value
                            oPOR.UserFields.Fields.Item("U_AB_FLT").Value = oRecordSet2.Fields.Item("U_F1").Value
                            oPOR.UserFields.Fields.Item("U_AB_Desc").Value = oRecordSet2.Fields.Item("U_ItemDesc").Value
                            oPOR.UserFields.Fields.Item("U_AB_ETDETA").Value = Format(oRecordSet2.Fields.Item("U_ETA").Value, "dd/MM/yy")
                            oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                        End If
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 2) = "SE" Then
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],T0.U_VessName [U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_VessVoyage] [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_GrssWt],T0.U_ItemDesc FROM [dbo].[@AB_SEAE_JOB_H] T0   WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "SE" 'oRecordSet2.Fields.Item("U_AB_Divsion").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value
                            oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotPkg").Value = oRecordSet2.Fields.Item("U_TotPkg").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet2.Fields.Item("U_GrssWt").Value
                            oPOR.UserFields.Fields.Item("U_AB_Desc").Value = oRecordSet2.Fields.Item("U_ItemDesc").Value
                            'oPOR.UserFields.Fields.Item("U_AB_ETDETA").Value = Format(oRecordSet2.Fields.Item("U_AB_ETDETA").Value, "dd/MM/yy")
                            oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                        End If
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 2) = "AI" Then
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],U_VessName [U_VessVoyage],U_AWBNo [U_OBL],U_HAWBNo [U_HBL],U_FNo [U_F1],T0.[U_GKBNo],T0.[U_VessName] FROM [dbo].[@AB_AIRI_JOB_H] T0 WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "AI" 'oRecordSet2.Fields.Item("U_AB_Divsion").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortNC").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet2.Fields.Item("U_ChrgWt").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_FLT").Value = oRecordSet2.Fields.Item("U_F1").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                        End If
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "AE" Then
                        Dim MAWBNo As String = ""
                        Dim HAWBNo As String = ""
                        oRecordSet2.DoQuery("SELECT T0.[U_AWBNo1] +T0.[U_AWBNo] FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
                        MAWBNo = oRecordSet2.Fields.Item(0).Value.ToString
                        oRecordSet2.DoQuery("SELECT T0.[U_HAWBNo] FROM [dbo].[@AB_AWB_H]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
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
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortN], T0.[U_TotPkg], T0.[U_ChrgWt],U_VessName [U_VessVoyage],'" & MAWBNo & "' [U_OBL],'" & HAWBNo & "' [U_HBL],U_FNo [U_F1],T0.[U_GKBNo],T0.[U_VessName] FROM [dbo].[@AB_AIRE_JOB_H] T0  WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "AE" 'oRecordSet2.Fields.Item("U_AB_Divsion").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_JobNo").Value = oRecordSet2.Fields.Item("U_JobNo").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortN").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_TotPkg").Value = oRecordSet3.Fields.Item("U_Pieces").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet3.Fields.Item("U_ChWeight").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_FLT").Value = oRecordSet3.Fields.Item("U_B1").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_ETDETA").Value = oRecordSet3.Fields.Item("U_FlighDate1").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_Desc").Value = oRecordSet3.Fields.Item("U_Nat").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "IN" Then

                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "PR" Then

                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "LO" Then

                    End If



                    '                oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = oRecordSet.Fields.Item("U_AB_Divsion").Value.ToString
                    For J = 1 To 500
                        'oPOR.Lines.UserFields.Fields.Item("U_DocEntry").Value = oRecordSet.Fields.Item("DocEntry").Value.ToString
                        'oPOR.Lines.UserFields.Fields.Item("U_LineNum").Value = oRecordSet.Fields.Item("LineId").Value.ToString
                        oPOR.Lines.ItemCode = oRecordSet.Fields.Item("U_ICode").Value
                        oPOR.Lines.Quantity = oRecordSet.Fields.Item("U_Qty").Value '"1.0"
                        oPOR.Lines.UnitPrice = oRecordSet.Fields.Item("U_Amt").Value
                        oPOR.Lines.VatGroup = "ZI"
                        oPOR.Lines.Add()
                        If oRecordSet.BoF <> False Then
                            oRecordSet.MoveNext()
                        Else
                            Exit For
                        End If
                        NewDocNum = oRecordSet.Fields.Item("DocEntry").Value.ToString
                        If NewDocNum <> DocNum Then
                            Exit For
                        End If
                    Next
                    Dim str As String = "Select U_APSt from [@AB_PAYMENTVOUCHER] where [U_APSt]='Success' and DocEntry='" & PVDocNum & "'"
                    oRecordSet21.DoQuery(str)
                    If oRecordSet21.RecordCount = 0 Then
                        K = oPOR.Add


                    End If
                    Dim K1 As String
                    K1 = PublicVariable.oCompany.GetNewObjectKey()
                    Dim st As String = ""
                    PublicVariable.oCompany.GetLastError(K, st)
                    If K <> 0 Then
                        oRecordSet1.DoQuery("UPDATE [@AB_PAYMENTVOUCHER] SET [U_APSt]='" & st & "' where DocEntry='" & PVDocNum & "'")
                    Else
                        oRecordSet1.DoQuery("UPDATE [@AB_PAYMENTVOUCHER] SET [U_APSt]='Success' where DocEntry='" & PVDocNum & "'")
                    End If
                    'oRecordSet.MoveNext()

                Catch ex As Exception
                    connect.WriteLog(ex.Message & "-" & i)
                    'oRecordSet.MoveNext()
                End Try
            Next
        Catch ex As Exception
            connect.WriteLog(ex.Message)
            'MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Create_PO()
        Try
            Dim oPOR As SAPbobsCOM.Documents
            oPOR = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            Dim i As Integer = 0
            Dim K As Integer = 0
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet1 As SAPbobsCOM.Recordset
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet21 As SAPbobsCOM.Recordset
            oRecordSet21 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim J As Integer = 0
            Dim NewDocNum As String = ""
            Dim DocNum As String = ""
            Dim JobNo As String = ""
            Dim PVDocNum As String = ""
            oRecordSet1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.[DocEntry],T1.[LineId],T0.[U_VODt],T0.[U_VOVCode], T0.[U_JobNo], T1.[U_ICode], T1.[U_IName],isnull(T1.[U_Qty],1) U_Qty, T1.[U_Amt] FROM [dbo].[@AB_PAYMENTVOUCHER]  T0 , [dbo].[@AB_PAYMENTVOUCHER_L]  T1 WHERE T1.[DocEntry] = T0.[DocEntry] and  isnull(T1.[U_ICode] ,'') <> '' and isnull(T0.U_APSt,'') <> 'Success'")
            For i = 1 To oRecordSet.RecordCount
                Try
                    oPOR = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    oPOR.DocDate = oRecordSet.Fields.Item("U_VODt").Value
                    DocNum = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    PVDocNum = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    oPOR.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                    oPOR.DocDueDate = oRecordSet.Fields.Item("U_VODt").Value
                    oPOR.TaxDate = oRecordSet.Fields.Item("U_VODt").Value
                    If oRecordSet.Fields.Item("U_VOVCode").Value.ToString.Length > 29 Then
                        oPOR.CardCode = oRecordSet.Fields.Item("U_VOVCode").Value.ToString.Substring(0, 30)
                    Else
                        oPOR.CardCode = oRecordSet.Fields.Item("U_VOVCode").Value.ToString
                    End If
                    oPOR.UserFields.Fields.Item("U_DocEntry").Value = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    'oPOR.UserFields.Fields.Item("U_LineNum").Value = oRecordSet.Fields.Item("LineNum").Value.ToString
                    oPOR.UserFields.Fields.Item("U_AB_JobNo").Value = oRecordSet.Fields.Item("U_JobNo").Value.ToString
                    JobNo = oRecordSet.Fields.Item("U_JobNo").Value.ToString
                    If oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 2) = "SI" Then
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],T0.U_VessName [U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_VessVoyage] [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_ItemDesc],T0.[U_GrssWt],T0.U_ETA FROM [dbo].[@AB_SEAI_JOB_H] T0   WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "ZI"
                            oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value
                            oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotPkg").Value = oRecordSet2.Fields.Item("U_TotPkg").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet2.Fields.Item("U_ChrgWt").Value
                            oPOR.UserFields.Fields.Item("U_AB_FLT").Value = oRecordSet2.Fields.Item("U_F1").Value
                            oPOR.UserFields.Fields.Item("U_AB_Desc").Value = oRecordSet2.Fields.Item("U_ItemDesc").Value
                            oPOR.UserFields.Fields.Item("U_AB_ETDETA").Value = Format(oRecordSet2.Fields.Item("U_ETA").Value, "dd/MM/yy")
                            oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                        End If
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 2) = "SE" Then
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],T0.U_VessName [U_VessVoyage], T0.[U_OBL], T0.[U_HBL], T0.[U_VessVoyage] [U_F1],T0.[U_GKBNo],T0.[U_VessName],T0.[U_GrssWt],T0.U_ItemDesc FROM [dbo].[@AB_SEAE_JOB_H] T0   WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "SE" 'oRecordSet2.Fields.Item("U_AB_Divsion").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value
                            oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value
                            oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortNC").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotPkg").Value = oRecordSet2.Fields.Item("U_TotPkg").Value
                            oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet2.Fields.Item("U_GrssWt").Value
                            oPOR.UserFields.Fields.Item("U_AB_Desc").Value = oRecordSet2.Fields.Item("U_ItemDesc").Value
                            'oPOR.UserFields.Fields.Item("U_AB_ETDETA").Value = Format(oRecordSet2.Fields.Item("U_AB_ETDETA").Value, "dd/MM/yy")
                            oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                        End If
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 2) = "AI" Then
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortNC], T0.[U_TotPkg], T0.[U_ChrgWt],U_VessName [U_VessVoyage],U_AWBNo [U_OBL],U_HAWBNo [U_HBL],U_FNo [U_F1],T0.[U_GKBNo],T0.[U_VessName] FROM [dbo].[@AB_AIRI_JOB_H] T0 WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "AI" 'oRecordSet2.Fields.Item("U_AB_Divsion").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortNC").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet2.Fields.Item("U_ChrgWt").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_FLT").Value = oRecordSet2.Fields.Item("U_F1").Value.ToString
                            oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                        End If
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "AE" Then
                        Dim MAWBNo As String = ""
                        Dim HAWBNo As String = ""
                        oRecordSet2.DoQuery("SELECT T0.[U_AWBNo1] +T0.[U_AWBNo] FROM [dbo].[@AB_AWB_M]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
                        MAWBNo = oRecordSet2.Fields.Item(0).Value.ToString
                        oRecordSet2.DoQuery("SELECT T0.[U_HAWBNo] FROM [dbo].[@AB_AWB_H]  T0 WHERE T0.[U_JobNo] ='" & JobNo & "'")
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
                        oRecordSet2.DoQuery("SELECT T0.[U_QNo], T0.[U_CCode], T0.[U_CName], T0.[U_Atten], T0.[U_ConNo], T0.[U_JobNo], '' [U_RefJob], T0.[U_LoadPortNC], T0.[U_DisPortN], T0.[U_TotPkg], T0.[U_ChrgWt],U_VessName [U_VessVoyage],'" & MAWBNo & "' [U_OBL],'" & HAWBNo & "' [U_HBL],U_FNo [U_F1],T0.[U_GKBNo],T0.[U_VessName] FROM [dbo].[@AB_AIRE_JOB_H] T0  WHERE T0.[U_JobNo]   ='" & JobNo & "'")
                        oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = "AE" 'oRecordSet2.Fields.Item("U_AB_Divsion").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_JobNo").Value = oRecordSet2.Fields.Item("U_JobNo").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_OriginNameC").Value = oRecordSet2.Fields.Item("U_LoadPortNC").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_DestNameC").Value = oRecordSet2.Fields.Item("U_DisPortN").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_TotPkg").Value = oRecordSet3.Fields.Item("U_Pieces").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_TotWT").Value = oRecordSet3.Fields.Item("U_ChWeight").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_SSIT").Value = oRecordSet2.Fields.Item("U_VessName").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_MAWB").Value = oRecordSet2.Fields.Item("U_OBL").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_HAWB").Value = oRecordSet2.Fields.Item("U_HBL").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_FLT").Value = oRecordSet3.Fields.Item("U_B1").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_ETDETA").Value = oRecordSet3.Fields.Item("U_FlighDate1").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_Desc").Value = oRecordSet3.Fields.Item("U_Nat").Value.ToString
                        oPOR.UserFields.Fields.Item("U_AB_Consignee").Value = oRecordSet2.Fields.Item("U_CName").Value.ToString
                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "IN" Then

                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "PR" Then

                    ElseIf oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(0, 1) = "LO" Then

                    End If



                    '                oPOR.UserFields.Fields.Item("U_AB_Divsion").Value = oRecordSet.Fields.Item("U_AB_Divsion").Value.ToString
                    For J = 1 To 500
                        'oPOR.Lines.UserFields.Fields.Item("U_DocEntry").Value = oRecordSet.Fields.Item("DocEntry").Value.ToString
                        'oPOR.Lines.UserFields.Fields.Item("U_LineNum").Value = oRecordSet.Fields.Item("LineId").Value.ToString
                        oPOR.Lines.ItemCode = oRecordSet.Fields.Item("U_ICode").Value
                        oPOR.Lines.Quantity = oRecordSet.Fields.Item("U_Qty").Value '"1.0"
                        oPOR.Lines.UnitPrice = oRecordSet.Fields.Item("U_Amt").Value
                        oPOR.Lines.VatGroup = "ZI"
                        oPOR.Lines.Add()
                        If oRecordSet.BoF <> False Then
                            oRecordSet.MoveNext()
                        Else
                            Exit For
                        End If
                        NewDocNum = oRecordSet.Fields.Item("DocEntry").Value.ToString
                        If NewDocNum <> DocNum Then
                            Exit For
                        End If
                    Next
                    Dim str As String = "Select U_APSt from [@AB_PAYMENTVOUCHER] where [U_APSt]='Success' and DocEntry='" & PVDocNum & "'"
                    oRecordSet21.DoQuery(str)
                    If oRecordSet21.RecordCount = 0 Then
                        K = oPOR.Add


                    End If
                    Dim K1 As String
                    K1 = PublicVariable.oCompany.GetNewObjectKey()
                    Dim st As String = ""
                    PublicVariable.oCompany.GetLastError(K, st)
                    If K <> 0 Then
                        oRecordSet1.DoQuery("UPDATE [@AB_PAYMENTVOUCHER] SET [U_APSt]='" & st & "' where DocEntry='" & PVDocNum & "'")
                    Else
                        oRecordSet1.DoQuery("UPDATE [@AB_PAYMENTVOUCHER] SET [U_APSt]='Success' where DocEntry='" & PVDocNum & "'")
                    End If
                    'oRecordSet.MoveNext()

                Catch ex As Exception
                    connect.WriteLog(ex.Message & "-" & i)
                    'oRecordSet.MoveNext()
                End Try
            Next
        Catch ex As Exception
            connect.WriteLog(ex.Message)
            'MsgBox(ex.Message)
        End Try
    End Sub

End Class
