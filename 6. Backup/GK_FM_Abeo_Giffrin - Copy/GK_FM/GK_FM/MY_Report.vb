Imports System.Diagnostics.Process
Imports System.Threading
Imports System.Data
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Public Class MY_Report

    Inherits System.Windows.Forms.Form
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Dim DataSet1 As DataSet1
    Dim DataSet2 As DataSet2
    Public Sub conn(ByVal ocompany As SAPbobsCOM.Company)


        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\GK_FM\" & "Pwd.txt", True)
        Dim pwd As String
        pwd = file.ReadLine()
        Dim connectionString As String = ""
        connectionString = "Provider=SQLOLEDB;"
        connectionString += "Server=" + ocompany.Server + ";Database=" + ocompany.CompanyDB + ";"
        connectionString += "User ID=" & ocompany.DbUserName & ";Password=" & pwd & ""  'toshiba"
        adoOleDbConnection = New OleDbConnection(connectionString)
        Try
            file.Close()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub HoldingStock_Report(ByVal ct1 As String, ByVal ocompany As SAPbobsCOM.Company)
        Try
            ' InitializeComponent()
            conn(ocompany)
            adoOleDbDataAdapter = New OleDbDataAdapter(ct1, adoOleDbConnection)
            DataSet2 = New DataSet2
            Dim Report1 As HoldingStockReport
            adoOleDbDataAdapter.Fill(DataSet2, "HOSTOCK")
            Report1 = New HoldingStockReport
            Report1.SetDataSource(DataSet2)
            CrystalReportViewer1.ReportSource = Report1
            CrystalReportViewer1.Visible = True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub HoldingStock_Report_1(ByVal ct1 As String, ByVal ocompany As SAPbobsCOM.Company)
        Try
            ' InitializeComponent()
            conn(ocompany)
            adoOleDbDataAdapter = New OleDbDataAdapter(ct1, adoOleDbConnection)
            DataSet2 = New DataSet2
            Dim Report1 As HoldingStockReport1
            adoOleDbDataAdapter.Fill(DataSet2, "HOSTOCK")
            Report1 = New HoldingStockReport1
            Report1.SetDataSource(DataSet2)
            CrystalReportViewer1.ReportSource = Report1
            CrystalReportViewer1.Visible = True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class