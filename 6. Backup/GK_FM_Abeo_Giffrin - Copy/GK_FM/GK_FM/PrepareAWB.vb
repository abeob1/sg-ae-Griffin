Imports System.Xml
Public Class PrepareAWB
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
        CreateUDTsFromXML()
    End Sub
    Private Sub CreateUDTsFromXML()
        Try
            Dim document As XmlDocument = New XmlDocument()
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            document.Load(sPath & "\GK_FM\" & "Tables.xml")
            CreateUDTsFromXML(document)
            AddUDO()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub CreateUDTsFromXML(ByVal document As XmlDocument)
        Dim tableNodes As XmlNodeList = document.SelectNodes("/Tables/Table[@Name and (not(@DesignTime)or(@DesignTime!='true'))]")
        Dim attribute As XmlAttribute

        oUserTableMD = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try

            For Each node As XmlNode In tableNodes
                attribute = node.Attributes("Name")

                If (oUserTableMD.GetByKey(attribute.Value)) Then ' The table already exists?
                    Continue For
                End If

                oUserTableMD.TableName = attribute.Value

                attribute = node.Attributes("Description")

                If Not attribute Is Nothing Then
                    oUserTableMD.TableDescription = attribute.Value
                End If


                attribute = node.Attributes("Type")

                If Not attribute Is Nothing Then

                    oUserTableMD.TableType = [Enum].Parse(GetType(SAPbobsCOM.BoUTBTableType), attribute.Value)
                End If

                Dim nErrorCode As Int32 = oUserTableMD.Add

                If nErrorCode <> 0 Then

                    If nErrorCode <> -2035 Then
                        Throw New Exception("Failed to add UDT.")
                    End If
                End If
            Next
        Finally
            oUserTableMD = Nothing
        End Try

        'Create fields.
        oUserFieldsMD = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try

            For Each node As XmlNode In tableNodes
                attribute = node.Attributes("Name")

                oUserFieldsMD.TableName = attribute.Value
                Dim fieldNodes As XmlNodeList = node.SelectNodes("./Fields/Field[@Name]")
                For Each fieldNode As XmlNode In fieldNodes

                    oUserFieldsMD.Name = fieldNode.Attributes("Name").Value

                    attribute = fieldNode.Attributes("Description")
                    oUserFieldsMD.Description = If(attribute Is Nothing, Nothing, attribute.Value)

                    attribute = fieldNode.Attributes("Type")
                    oUserFieldsMD.Type = If(attribute Is Nothing, SAPbobsCOM.BoFieldTypes.db_Alpha, [Enum].Parse(GetType(SAPbobsCOM.BoFieldTypes), attribute.Value))

                    attribute = fieldNode.Attributes("SubType")
                    oUserFieldsMD.SubType = If(attribute Is Nothing, SAPbobsCOM.BoFldSubTypes.st_None, [Enum].Parse(GetType(SAPbobsCOM.BoFldSubTypes), attribute.Value))

                    attribute = fieldNode.Attributes("EditSize")
                    oUserFieldsMD.EditSize = If(attribute Is Nothing, 0, Int32.Parse(attribute.Value))

                    Dim nErrorCode As Int32 = oUserFieldsMD.Add

                    If nErrorCode <> 0 Then

                        If nErrorCode <> -2035 Then
                            'Throw New Exception("Failed to add UDT.")
                        End If
                    End If
                Next
            Next
        Finally

            oUserFieldsMD = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub AddUDO()
        Try
            Dim lRetCode As Integer
            Dim ErrorMss As String

            'Add UDO for House AWB
            oUserObjectMD = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUserObjectMD.GetByKey("MAWB") Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES

                oUserObjectMD.Code = "HAWB"
                oUserObjectMD.Name = "HAWB"
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
                oUserObjectMD.TableName = "AB_AWB_H"
                oUserObjectMD.ChildTables.TableName = "AB_AWB_H1"
                oUserObjectMD.ChildTables.Add()
                oUserObjectMD.ChildTables.TableName = "AB_AWB_H2"
                oUserObjectMD.ChildTables.Add()
                oUserObjectMD.ChildTables.TableName = "AB_AWB_H3"
                oUserObjectMD.ChildTables.Add()

                lRetCode = oUserObjectMD.Add()
                If lRetCode <> 0 Then
                    ErrorMss = ""
                    Ocompany.GetLastError(lRetCode, ErrorMss)
                    SBO_Application.StatusBar.SetText(ErrorMss, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
            oUserObjectMD = Nothing

            'Add UDO for Master AWB
            oUserObjectMD = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUserObjectMD.GetByKey("MAWB") Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES

                oUserObjectMD.Code = "MAWB"
                oUserObjectMD.Name = "MAWB"
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
                oUserObjectMD.TableName = "AB_AWB_M"
                oUserObjectMD.ChildTables.TableName = "AB_AWB_M1"
                oUserObjectMD.ChildTables.Add()
                oUserObjectMD.ChildTables.TableName = "AB_AWB_M2"
                oUserObjectMD.ChildTables.Add()
                oUserObjectMD.ChildTables.TableName = "AB_AWB_M3"
                oUserObjectMD.ChildTables.Add()

                lRetCode = oUserObjectMD.Add()
                If lRetCode <> 0 Then
                    ErrorMss = ""
                    Ocompany.GetLastError(lRetCode, ErrorMss)
                    SBO_Application.StatusBar.SetText(ErrorMss, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
            oUserObjectMD = Nothing
        Finally
            If Not oUserObjectMD Is Nothing Then
                oUserObjectMD = Nothing
            End If
            GC.Collect()
        End Try
    End Sub
End Class
