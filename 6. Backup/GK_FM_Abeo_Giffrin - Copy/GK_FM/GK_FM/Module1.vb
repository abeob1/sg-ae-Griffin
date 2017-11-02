Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.Threading.Thread
Imports System.Threading
Module Module1

    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public Bol As Boolean = False
    Public BP123 As String
    Public DocEtry As Integer
    Public F As Integer = 0
    Public adoOleDbConnection As OleDbConnection
    Public adoOleDbDataAdapter1 As OleDbDataAdapter
    Public adoOleDbDataAdapter2 As OleDbDataAdapter
    Public adoOleDbDataAdapter As OleDbDataAdapter
    Public DocNumPO As String = ""
    Public oForm As SAPbouiCOM.Form
    Public oOrderForm As SAPbouiCOM.Form
    Public oNewItem As SAPbouiCOM.Item
    Public oFolderItem As SAPbouiCOM.Folder
    Public oGrid As SAPbouiCOM.Grid
    Public oPict As SAPbouiCOM.PictureBox
    Public oCombo As SAPbouiCOM.ComboBox
    Public oCombo1 As SAPbouiCOM.ComboBox
    Public oCombo2 As SAPbouiCOM.ComboBox
    Public oCombo3 As SAPbouiCOM.ComboBox
    Public oCombo4 As SAPbouiCOM.ComboBox
    Public oChk1 As SAPbouiCOM.CheckBox
    Public oChk2 As SAPbouiCOM.CheckBox
    Public oChk3 As SAPbouiCOM.CheckBox
    Public oChk4 As SAPbouiCOM.CheckBox
    Public Vencode As String
    Public Type As String
    Public SubType As String
    Public BIAddButt As String
    Public BIAddButt_Disable As String
    Public PVAddButt As String
    Public oButton As SAPbouiCOM.Button
    Public oCheck As SAPbouiCOM.CheckBox
    Public ooption As SAPbouiCOM.OptionBtn
    Public oColumns As SAPbouiCOM.Columns
    Public oColumns1 As SAPbouiCOM.Columns
    Public oColumns2 As SAPbouiCOM.Columns
    Public oColumn As SAPbouiCOM.Column
    Public oColumn1 As SAPbouiCOM.Column
    Public oColumn2 As SAPbouiCOM.Column
    Public oEdit As SAPbouiCOM.EditText
    Public oEEdit As SAPbouiCOM.EditTextColumn
    Public oMenu As SAPbouiCOM.Menus
    Public oMenuItem As SAPbouiCOM.MenuItem
    Public oMenuParam As SAPbouiCOM.MenuCreationParams
    Public oStatic As SAPbouiCOM.StaticText
    Public oCFLs As SAPbouiCOM.ChooseFromListCollection
    Public Obutt As SAPbouiCOM.Button
    Public UserDS As SAPbouiCOM.UserDataSource
    Public UserDS1 As SAPbouiCOM.UserDataSource
    Public oMatrix As SAPbouiCOM.Matrix
    Public oMatrix1 As SAPbouiCOM.Matrix
    Public oMatrix2 As SAPbouiCOM.Matrix
    Public oMatrix3 As SAPbouiCOM.Matrix
    Public oMatrix4 As SAPbouiCOM.Matrix
    Public oMatrix5 As SAPbouiCOM.Matrix

    Public oCreation As SAPbouiCOM.FormCreationParams
    Public oItem As SAPbouiCOM.Item
    Public oRecordSet As SAPbobsCOM.Recordset
    Public oRecordSet1 As SAPbobsCOM.Recordset
    Public oRecordSet2 As SAPbobsCOM.Recordset
    Public oRecordSet3 As SAPbobsCOM.Recordset
    Public oRecordSet4 As SAPbobsCOM.Recordset
    Public oRecordSet5 As SAPbobsCOM.Recordset
    Public oDBDataSource As SAPbouiCOM.DBDataSource
    Public oUserDataSource As SAPbouiCOM.DBDataSource
    Public oUserTableMD As SAPbobsCOM.UserTablesMD
    Public oUserFieldsMD As SAPbobsCOM.UserFieldsMD
    Public oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Public format1 As New System.Globalization.CultureInfo("fr-FR", True)

    Public oFolder As SAPbouiCOM.Folder
    Public OCCode As String = ""
    Public OCName As String = ""
    Public OBiiSt As String = ""
    Public OZIp As String = ""
    Public OPh1 As String = ""
    Public OPh2 As String = ""
    Public OCell As String = ""
    Public OEmail As String = ""
    Public NCCode As String = ""
    Public NCName As String = ""
    Public NBiiSt As String = ""
    Public NZIp As String = ""
    Public NPh1 As String = ""
    Public NPh2 As String = ""
    Public NCell As String = ""
    Public NEmail As String = ""

    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Dim oXmlDoc As New Xml.XmlDocument
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oXmlDoc.Load(sPath & "\GK_FM\" & FileName)
        Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
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
        End Try
    End Sub
    Public Sub SaveAsXML(ByVal Form As SAPbouiCOM.Form, ByVal FileName As String)
        Dim oXmlDoc As New Xml.XmlDocument
        Dim sXmlString As String
        Dim sPath As String
        sXmlString = Form.GetAsXML
        oXmlDoc.LoadXml(sXmlString)
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oXmlDoc.Save(sPath & "\HE\" & FileName)
    End Sub

#Region "CFL"
    Public Sub CFL_BP_WareHouse(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbobsCOM.BoObjectTypes.oWarehouses
            oCFLCreationParams.UniqueID = "CFLWSC"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)

            'oCFLCreationParams.UniqueID = "CFLBPC1"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCFL = oCFLs.Add(oCFLCreationParams)
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Active"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

            '=================
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "DropShip"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            'oCon.BracketCloseNum = 1
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            'oCon = oCons.Add
            'oCon.BracketOpenNum = 1
            'oCon.Alias = "U_Active"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2 '[/code]
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "CFLWHSC"
            oCFL = oCFLs.Add(oCFLCreationParams)

            '=================
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_BP_Customer(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFLBPC"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)

            'oCFLCreationParams.UniqueID = "CFLBPC1"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCFL = oCFLs.Add(oCFLCreationParams)
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Active"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

            '=================
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2 '[/code]
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "1CFLBPC1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            '=================
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_BP_Customer1(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFLBPC1"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)

            'oCFLCreationParams.UniqueID = "CFLBPC1"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCFL = oCFLs.Add(oCFLCreationParams)
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Active"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

            '=================
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
            oCon.CondVal = "C"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2 '[/code]
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "1CFLBPC11"
            oCFL = oCFLs.Add(oCFLCreationParams)

            '=================
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_SalesOrder_International(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application, ByVal Division As String) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "ORDR"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            ''oCon = oCons.Add()
            ''oCon.Alias = "U_AB_Divsion"
            ''oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''oCon.CondVal = Division


            ''oCFL.SetConditions(oCons)
            ''oCFLCreationParams.UniqueID = "ORDRA1"
            ''oCFL = oCFLs.Add(oCFLCreationParams)

            ''oCons = oCFL.GetConditions()
            ''oCon = oCons.Add()
            ''oCon.Alias = "U_AB_JobNo"
            ''oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''oCon.CondVal = Nothing
            ''oCFL.SetConditions(oCons)
            ''oCFLCreationParams.UniqueID = "ORDRA2"
            ''oCFL = oCFLs.Add(oCFLCreationParams)

            oCon = oCons.Add
            oCon.BracketOpenNum = 3
            oCon.Alias = "U_AB_Divsion1"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Division
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_AB_Status1"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Open"
            oCon.BracketCloseNum = 2
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCon.BracketCloseNum = 3
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "ORDRA1"
            oCFL = oCFLs.Add(oCFLCreationParams)


        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_SalesOrder(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application, ByVal Division As String) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList

            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "ORDR"


            Try
                oCFL = oCFLs.Add(oCFLCreationParams)
            Catch ex As Exception

            End Try
            oCons = oCFL.GetConditions()
            ''oCon = oCons.Add()
            ''oCon.Alias = "U_AB_Divsion"
            ''oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''oCon.CondVal = Division


            ''oCFL.SetConditions(oCons)
            ''oCFLCreationParams.UniqueID = "ORDRA1"
            ''oCFL = oCFLs.Add(oCFLCreationParams)

            ''oCons = oCFL.GetConditions()
            ''oCon = oCons.Add()
            ''oCon.Alias = "U_AB_JobNo"
            ''oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''oCon.CondVal = Nothing
            ''oCFL.SetConditions(oCons)
            ''oCFLCreationParams.UniqueID = "ORDRA2"
            ''oCFL = oCFLs.Add(oCFLCreationParams)

            oCon = oCons.Add
            oCon.BracketOpenNum = 3
            oCon.Alias = "U_AB_Divsion"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Division
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_AB_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Open"
            oCon.BracketCloseNum = 2
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCon.BracketCloseNum = 3
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "ORDRA1"
            oCFL = oCFLs.Add(oCFLCreationParams)


        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_BP_Supplier(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFLBPV"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "CFLBPC11"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_BP_Supplier2(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFLBPV1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "CFLBPC111"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_BP_Supplier3(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2" 'SAPbobsCOM.BoObjectTypes.oOrders
            oCFLCreationParams.UniqueID = "3CFLBPV1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "3CFLBPC111"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_OpenSalesOrder(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application, ByVal Cond As String) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbobsCOM.BoObjectTypes.oOrders
            oCFLCreationParams.UniqueID = "ORDR1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_AB_Divsion"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Cond '"S"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "CONORDR1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCon.Alias = "U_AB_Divsion"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "" '"S"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "CONORDR2"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_Item(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "OITM"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItmsGrpCod"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "103"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "OOITM1"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_Item1(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "1OITM"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItmsGrpCod"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "102"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "1OOITM1"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_Item2(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "21OITM"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItmsGrpCod"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "102"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "21OOITM1"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_Item_Vessel(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "OITM11"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItmsGrpCod"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "102"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "OITMM1"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_Item_Vessel2(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "OITM112"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItmsGrpCod"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "102"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "OITMM12"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_Item_Vessel3(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "OITM112"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItmsGrpCod"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "102"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "OITMM12"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
    Public Sub CFL_BP_Employee(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Employee
            oCFLCreationParams.UniqueID = "CFLEMP"
            oCFL = oCFLs.Add(oCFLCreationParams)
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)
            'oCFLCreationParams.UniqueID = "CFLBPC1"
            'oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
#End Region
#Region "Function"
    Public Function BPPhoneNo(ByVal CardCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[Phone1] FROM OCRD T0  WHERE  T0.[CardCode] ='" & CardCode & "'")
        x = oRecordSet_UN.Fields.Item(0).Value
        Return x
    End Function
    Public Function UserName(ByVal UserCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[U_NAME] FROM OUSR T0 WHERE T0.[USER_CODE] ='" & UserCode & "'")
        x = oRecordSet_UN.Fields.Item(0).Value
        Return x
    End Function
    Public Function FightPreFix(ByVal Carrier As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[U_PreFix], T0.[U_Address] FROM [dbo].[@AB_CARRIER]  T0 WHERE T0.[Code] ='" & Carrier & "'")
        x = oRecordSet_UN.Fields.Item(0).Value
        Return x
    End Function
    Public Function Fightaddress(ByVal Carrier As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[U_PreFix], T0.[U_Address],T0.[Name] FROM [dbo].[@AB_CARRIER]  T0 WHERE T0.[Code] ='" & Carrier & "'")
        x = oRecordSet_UN.Fields.Item(2).Value & vbCrLf & oRecordSet_UN.Fields.Item(1).Value
        Return x.ToUpper
    End Function
    Public Function FightLogo(ByVal Carrier As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[U_Logo] FROM [dbo].[@AB_CARRIER]  T0 WHERE T0.[Code] ='" & Carrier & "'")
        x = oRecordSet_UN.Fields.Item(0).Value
        oRecordSet_UN.DoQuery("SELECT BitmapPath from OADP")
        x = oRecordSet_UN.Fields.Item(0).Value & x
        Return x

    End Function
    Public Function CardCode(ByVal Cardname As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[CardCode] FROM OCRD T0  WHERE  T0.[Cardname] ='" & Cardname & "'")
        x = oRecordSet_UN.Fields.Item(0).Value
        Return x
    End Function
    Public Function CardCode_Supplier(ByVal Cardname As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[CardCode] FROM OCRD T0  WHERE  T0.[Cardname] ='" & Cardname & "' and T0.CardType='S'")
        x = oRecordSet_UN.Fields.Item(0).Value
        Return x
    End Function
    Public Function ContactPerson(ByVal CardCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[CntctPrsn] FROM OCRD T0 where T0.[CardCode] ='" & CardCode & "'")
        x = oRecordSet_UN.Fields.Item(0).Value
        Return x.ToUpper
    End Function
    Public Function TAXPer(ByVal TaxCode As String, ByVal Ocompany As SAPbobsCOM.Company) As Double
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[Code], T0.[Name], T0.[Category], T0.[Rate] FROM OVTG T0 WHERE T0.[Category] ='I' and T0.[Code]='" & TaxCode & "' ")
        x = oRecordSet_UN.Fields.Item("Rate").Value
        Return x
    End Function
    Public Function Country_Code(ByVal Code As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[Name] FROM OCRY T0 WHERE T0.[Code] ='" & Code & "'")
        x = oRecordSet_UN.Fields.Item(0).Value.ToString.ToUpper
        Return x.ToUpper
    End Function
    Public Function Roundoff(ByVal intThisTemperature As Double) As Double
        Try
            Return FormatNumber((Math.Round(intThisTemperature * 2) / 2), 1)
        Catch ex As Exception
            Return 0
        End Try

    End Function
    Public Function Carrier_Name(ByVal Code As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[Name] FROM [dbo].[@AB_CARRIER]  T0 WHERE T0.[Code] ='" & Code & "'")
        x = oRecordSet_UN.Fields.Item(0).Value.ToString.ToUpper
        Return x.ToUpper
    End Function

    Public Function City_Code(ByVal Code As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_UN As SAPbobsCOM.Recordset
        oRecordSet_UN = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_UN.DoQuery("SELECT T0.[Name] FROM [dbo].[@AB_CITY]  T0 WHERE T0.[Code]  ='" & Code & "'")
        x = oRecordSet_UN.Fields.Item(0).Value.ToString.ToUpper
        Return x.ToUpper
    End Function
    Public Function BPNameAddress(ByVal BpCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        'T0.[StreetNo], T0.[City]
        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("SELECT T1.[Street],T1.[Block], T1.[Building], T1.[ZipCode], (select  Name from ocry where code=T1.[Country]),T0.CardName,T1.[StreetNo], T1.[City] FROM [dbo].[OCRD]  T0 Left JOIN CRD1 T1 ON T0.CardCode = T1.CardCode WHERE T1.[Address] = T0.[BillToDef]  and  T0.[CardCode]  ='" & BpCode & "'")
        x = oRecordSet.Fields.Item(0).Value.ToString.Trim & "  " & oRecordSet.Fields.Item("StreetNo").Value.ToString.Trim & " " & oRecordSet.Fields.Item(1).Value.ToString.Trim & vbCrLf & oRecordSet.Fields.Item("City").Value.ToString.Trim & "  " & oRecordSet.Fields.Item(3).Value.ToString.Trim & vbCrLf & oRecordSet.Fields.Item(4).Value.ToString.Trim
        Return x.Trim().ToUpper
    End Function
    Public Function BPAddress(ByVal BpCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("SELECT T1.[Street],T1.[Block], T1.[Building], T1.[ZipCode], (select  Name from ocry where code=T1.[Country]),T0.CardName,T1.[StreetNo], T1.[City] FROM [dbo].[OCRD]  T0 Left JOIN CRD1 T1 ON T0.CardCode = T1.CardCode WHERE T1.[Address] = T0.[BillToDef]  and  T0.[CardCode]  ='" & BpCode & "'")
        x = oRecordSet.Fields.Item(0).Value.ToString.Trim & "  " & oRecordSet.Fields.Item("StreetNo").Value.ToString.Trim & " " & oRecordSet.Fields.Item(1).Value.ToString.Trim & vbCrLf & oRecordSet.Fields.Item("City").Value.ToString.Trim & "  " & oRecordSet.Fields.Item(3).Value.ToString.Trim & vbCrLf & oRecordSet.Fields.Item(4).Value.ToString.Trim
        Return x.Trim().ToUpper
    End Function
    Public Function Deafault_Whsc(ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        'oRecordSet3 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRecordSet3.DoQuery("SELECT T0.[DfltWhs] FROM OADM T0")
        'x = oRecordSet3.Fields.Item(0).Value
        Return x
    End Function
    Public Function Terms(ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        x = "Rates quoted are strictly based on the given cargo specifications." & vbCrLf & "Not applicable to odd-Sized cargo." & vbCrLf & "Above quotation is subject to space availability upon carrier's booking confirmation." & vbCrLf & "Duties & Taxes Excluded." & vbCrLf & vbCrLf & "We Strongly recommend that you insure your shipment for the full value against loss(s) or damage(s).  If you do not have goods in transit policy covering all your shipments, we can arrange insurance for your shipments at  your request." & vbCrLf & "We look forward to receiving your favourable confirmation. Kindly confirm acceptance by endorsing below."
        Return x
    End Function
    Public Function BPName(ByVal BPCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_Ow As SAPbobsCOM.Recordset

        oRecordSet_Ow = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_Ow.DoQuery("SELECT T0.[CardName] FROM OCRD T0 WHERE T0.[CardCode]='" & BPCode & "'")
        x = oRecordSet_Ow.Fields.Item(0).Value
        Return x
    End Function
    Public Function BPMarkupPer(ByVal BPCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_Ow As SAPbobsCOM.Recordset

        oRecordSet_Ow = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_Ow.DoQuery("SELECT T0.[U_AB_Markup] FROM OCRD T0 WHERE T0.[CardCode]='" & BPCode & "'")
        x = oRecordSet_Ow.Fields.Item(0).Value
        Return x
    End Function
    Public Function ItemName(ByVal ItemCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_Ow As SAPbobsCOM.Recordset

        oRecordSet_Ow = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_Ow.DoQuery("SELECT T0.[ItemName] FROM OITM T0 WHERE T0.[ItemCode] ='" & ItemCode & "'")
        x = oRecordSet_Ow.Fields.Item(0).Value
        Return x
    End Function
    Public Function WhscCode2(ByVal WhscName As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_Ow As SAPbobsCOM.Recordset

        oRecordSet_Ow = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_Ow.DoQuery("SELECT T0.[WhsCode] FROM OWHS T0 WHERE T0.[WhsName]='" & WhscName & "'")
        x = oRecordSet_Ow.Fields.Item(0).Value
        Return x
    End Function
    Public Function PSize(ByVal PetCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        Dim oRecordSet_PZ As SAPbobsCOM.Recordset
        oRecordSet_PZ = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet_PZ.DoQuery("SELECT T0.[U_AI_Size] FROM [dbo].[@AI_PETMASTER]  T0 WHERE T0.[U_PetCode] ='" & PetCode & "'")
        x = oRecordSet_PZ.Fields.Item(0).Value
        Return x
    End Function
    Public Function ProgName(ByVal ProgCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet2.DoQuery("SELECT T0.[Name] FROM [dbo].[@AI_PROGRAM]  T0 WHERE T0.[Code]  ='" & ProgCode & "'")
        x = oRecordSet2.Fields.Item(0).Value
        Return x
    End Function
    Public Function MobileUserName(ByVal UserCode As String, ByVal Ocompany As SAPbobsCOM.Company) As String
        Dim x As String = ""
        oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet2.DoQuery("SELECT T0.Name FROM [dbo].[@AI_USER]  T0 WHERE T0.[Code] ='" & UserCode & "'")
        x = oRecordSet2.Fields.Item(0).Value
        Return x
    End Function

#End Region
#Region "ComboLoad"
    Public Sub ComboLoad_Carrier(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox, ByVal oCOmpany As SAPbouiCOM.Company, ByVal SBO_Application As SAPbouiCOM.Application)
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
#End Region


End Module
