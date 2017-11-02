Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data.SqlTypes
Public Class Connection
    Public Shared sConn As SqlConnection
    Public Shared sConnSAP As SqlConnection
    Public Shared bConnect As Boolean
    Public Sub setDB()
        Try
            Dim strConnect As String = ""
            Dim sCon As String = ""
            Dim SQLType As String = ""
            Dim MyArr As Array
            Dim sErrMsg As String = ""
            strConnect = "SAPConnect"
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString.trim()
            Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\" & "ServerConfiguration.txt", True)
            ' Dim sCon As String
            sCon = file.ReadLine()
            ' sCon = System.Configuration.ConfigurationSettings.AppSettings.Get(strConnect)
            MyArr = sCon.Split(";")
            If IsNothing(PublicVariable.oCompany) Then
                PublicVariable.oCompany = New SAPbobsCOM.Company
            End If
            PublicVariable.oCompany.CompanyDB = MyArr(0).ToString.Trim()
            PublicVariable.oCompany.UserName = MyArr(1).ToString.trim()
            PublicVariable.oCompany.Password = MyArr(2).ToString.trim()
            PublicVariable.oCompany.Server = MyArr(3).ToString.trim()
            PublicVariable.oCompany.DbUserName = MyArr(4).ToString.trim()
            PublicVariable.oCompany.DbPassword = MyArr(5).ToString.Trim()
            PublicVariable.oCompany.LicenseServer = MyArr(6)
            SQLType = MyArr(7)
            If SQLType = 2008 Then
                PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            Else
                PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            End If
            sCon = "server= " + MyArr(3).ToString.Trim() + ";database=" + MyArr(0).ToString.Trim() + " ;uid=" + MyArr(4).ToString.Trim() + "; pwd=" + MyArr(5).ToString.Trim() + ";"
            sConnSAP = New SqlConnection(sCon)
        Catch ex As Exception
            '   Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("C:\\SetDB.txt", True)
            '  file.WriteLine(ex)
            'file.Close()
        End Try
    End Sub
    Public Function connectDB() As Boolean
        Try
            Dim sErrMsg As String = ""
            Dim connectOk As Integer = 0
            If PublicVariable.oCompany.Connect() <> 0 Then
                PublicVariable.oCompany.GetLastError(connectOk, sErrMsg)
                WriteLog(sErrMsg)
                bConnect = False
                Return bConnect

            Else
                WriteLog("Connected")
                bConnect = True
                Return bConnect
            End If
        Catch ex As Exception
            'Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("C:\\connectDB.txt", True)
            'file.WriteLine(ex)
            'file.Close()
            Return False
        End Try
    End Function
    Public Sub WriteLog(ByVal Str As String)
        Dim oWrite As IO.StreamWriter
        Dim FilePath As String
        FilePath = Application.StartupPath + "\ErrorLog.txt"

        If IO.File.Exists(FilePath) Then
            oWrite = IO.File.AppendText(FilePath)
        Else
            oWrite = IO.File.CreateText(FilePath)
        End If
        oWrite.Write(Now.ToString.Trim() + ":" + Str + vbCrLf)
        oWrite.Close()
    End Sub
End Class
