Imports System.Diagnostics.Process
Imports System.Threading
Public Class F_AWB
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Dim rowDelete As Integer
    Dim matrixUID As String
    Dim hawbForm As SAPbouiCOM.Form = Nothing
    Dim mawbForm As SAPbouiCOM.Form = Nothing
    Dim childTop As Integer
    Dim childLeft As Integer
    Dim AWBForm As SAPbouiCOM.Form = Nothing
    Dim ModalForm As SAPbouiCOM.Form = Nothing
    Dim oF_PiecesWeight As F_PiecesWeight
    Dim oF_AWBParameter As F_AWBParameter

    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        'SBO_Application = sbo_application1
        'Ocompany = ocompany1
    End Sub
    'Public Sub AWB_Bind(ByVal oForm As SAPbouiCOM.Form, ByVal JobNo As String)
    '    AWBForm = oForm
    '    Select Case oForm.BusinessObject.Type
    '        Case "HAWB"
    '            hawbForm = oForm
    '        Case "MAWB"
    '            mawbForm = oForm
    '    End Select
    '    '0_U_E
    '    oEdit = oForm.Items.Item("0_U_E").Specific
    '    oEdit.String = JobNo

    '    ooption = oForm.Items.Item("optionbtn2").Specific
    '    ooption.GroupWith("optionbtn1")

    '    ooption = oForm.Items.Item("optionbtn4").Specific
    '    ooption.GroupWith("optionbtn3")



    '    oForm.Freeze(True)
    '    oFolderItem = oForm.Items.Item("0_U_FD").Specific
    '    oFolderItem.Select()
    '    oForm.Freeze(False)
    'End Sub
    ''Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
    ''    Try

    ''        If FormUID = If(mawbForm Is Nothing, "", mawbForm.UniqueID) Then
    ''            AWBForm = mawbForm
    ''        ElseIf FormUID = If(hawbForm Is Nothing, "", hawbForm.UniqueID) Then
    ''            AWBForm = hawbForm
    ''        End If

    ''        If pVal.BeforeAction Then
    ''            Select Case pVal.EventType
    ''                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
    ''                    AWBForm = Nothing
    ''            End Select
    ''        Else
    ''            BubbleEvent = True
    ''            If Not AWBForm Is Nothing Then
    ''                Select Case pVal.EventType
    ''                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
    ''                        'Add row to matrix when press button Add which at bottom of matrix
    ''                        If pVal.ItemUID = "btnAdd1" Then
    ''                            oMatrix = AWBForm.Items.Item("AWB_Mtr1").Specific
    ''                            oMatrix.AddRow(1, oMatrix.RowCount)
    ''                        ElseIf pVal.ItemUID = "btnAdd2" Then
    ''                            oMatrix = AWBForm.Items.Item("AWB_Mtr2").Specific
    ''                            oMatrix.AddRow(1, oMatrix.RowCount)
    ''                        ElseIf pVal.ItemUID = "CDManifest" Or pVal.ItemUID = "Manifest" Then
    ''                            oF_PiecesWeight = New F_PiecesWeight(Ocompany, SBO_Application, AWBForm, pVal.ItemUID)
    ''                        ElseIf pVal.ItemUID = "btn_Print" Then
    ''                            oF_AWBParameter = New F_AWBParameter(Ocompany, SBO_Application, AWBForm, pVal.ItemUID)
    ''                        End If
    ''                        If pVal.BeforeAction = False Then
    ''                            If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
    ''                                AWBForm.Close()
    ''                                AWBForm = Nothing
    ''                            End If
    ''                        End If
    ''                End Select
    ''            End If
    ''        End If

    ''    Catch ex As Exception
    ''        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ''    End Try
    ''End Sub
    'Private Sub SBO_Application_RghtClick(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
    '    If eventInfo.ItemUID = "AWB_Mtr1" Or eventInfo.ItemUID = "AWB_Mtr2" Then
    '        Try
    '            If eventInfo.BeforeAction Then
    '                Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
    '                Dim oMenus As SAPbouiCOM.Menus = Nothing

    '                matrixUID = eventInfo.ItemUID
    '                rowDelete = eventInfo.Row
    '                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
    '                oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
    '                oMenuItem = SBO_Application.Menus.Item("1280") 'Data
    '                oMenus = oMenuItem.SubMenus

    '                If Not SBO_Application.Menus.Exists("DeleteRow") Then
    '                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
    '                    oCreationPackage.UniqueID = "DeleteRow"
    '                    oCreationPackage.String = "Delete Row"
    '                    oCreationPackage.Enabled = True
    '                    oMenus.AddEx(oCreationPackage)
    '                End If

    '                If Not SBO_Application.Menus.Exists("ClearMatrix") Then
    '                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
    '                    oCreationPackage.UniqueID = "ClearMatrix"
    '                    oCreationPackage.String = "Clear Matrix"
    '                    oCreationPackage.Enabled = True
    '                    oMenus.AddEx(oCreationPackage)
    '                End If
    '            Else
    '                If SBO_Application.Menus.Exists("DeleteRow") Then
    '                    SBO_Application.Menus.RemoveEx("DeleteRow")
    '                End If
    '                If SBO_Application.Menus.Exists("ClearMatrix") Then
    '                    SBO_Application.Menus.RemoveEx("ClearMatrix")
    '                End If
    '            End If
    '        Catch ex As Exception
    '            SBO_Application.MessageBox(ex.Message)
    '        End Try
    '    End If
    'End Sub

    'Private Sub SBO_Application_MenuEvent(ByRef menuEvent As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
    '    If Not menuEvent.BeforeAction Then
    '        Try
    'Dim matrix As SAPbouiCOM.Matrix
    '            If menuEvent.MenuUID = "DeleteRow" Then
    '                matrix = AWBForm.Items.Item(matrixUID).Specific
    '                If rowDelete <> 0 And rowDelete <> matrix.RowCount Then
    '                    matrix.DeleteRow(rowDelete)
    '                    rowDelete = 0
    '                    If AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                        AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '                    End If
    '                End If
    '            ElseIf menuEvent.MenuUID = "ClearMatrix" Then
    '                matrix = AWBForm.Items.Item(matrixUID).Specific
    '                matrix.Clear()
    '                matrix.AddRow(1, 0)
    '                matrix.FlushToDataSource()
    '                If AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                    AWBForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '                End If
    '            End If

    '        Catch ex As Exception
    '            SBO_Application.MessageBox(ex.Message)

    '        End Try
    '    End If
    'End Sub
End Class
