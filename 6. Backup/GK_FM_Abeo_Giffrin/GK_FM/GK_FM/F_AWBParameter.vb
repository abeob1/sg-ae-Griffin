Public Class F_AWBParameter
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Dim parentForm As SAPbouiCOM.Form = Nothing
    Dim itemCallUID As String
    Private oForm As SAPbouiCOM.Form = Nothing

    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application, ByVal awbForm As SAPbouiCOM.Form, ByVal itemUID As String)
        Try
            SBO_Application = sbo_application1
            Ocompany = ocompany1
            parentForm = awbForm
            itemCallUID = itemUID

            Dim formCreationParams As SAPbouiCOM.FormCreationParams
            Dim oXmlDoc As New Xml.XmlDocument
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            oXmlDoc.Load(sPath & "\GK_FM\" & "AWBParameterForm.srf")
            formCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            formCreationParams.UniqueID = parentForm.UniqueID + "-" + "AWBParameter"

            formCreationParams.XmlData = oXmlDoc.InnerXml

            oForm = SBO_Application.Forms.AddEx(formCreationParams)
            oForm.Visible = False
            oForm.Top = parentForm.Top + (parentForm.Height / 2) - (oForm.Height / 2)
            oForm.Left = parentForm.Left + (parentForm.Width / 2) - (oForm.Width / 2)
            oForm.Visible = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.BeforeAction Then
                If If(oForm Is Nothing, True, FormUID <> oForm.UniqueID) Then
                    If FormUID = parentForm.UniqueID And Not oForm Is Nothing Then
                        BubbleEvent = False
                        oForm.Select()
                    End If
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            oForm = Nothing
                            BubbleEvent = True
                    End Select
                End If
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If Not oForm Is Nothing And FormUID = oForm.UniqueID Then
                            Select Case pVal.ItemUID
                                Case "Cancel"
                                    oForm.Close()
                                    oForm = Nothing
                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub
    'Private Sub SBO_Application_MenuEvent(ByRef menuEvent As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
    '    If menuEvent.BeforeAction Then
    '        BubbleEvent = False
    '    End If
    'End Sub
End Class
