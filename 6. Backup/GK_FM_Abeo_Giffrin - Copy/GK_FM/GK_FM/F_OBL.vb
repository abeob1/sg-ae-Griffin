Imports System.Diagnostics.Process
Imports System.Threading
Public Class F_OBL
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "SEAE_OBL" Then
                oForm = SBO_Application.Forms.Item("SEAE_OBL")
                If pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oForm.Close()
                        oForm = Nothing
                    End If
                End If
            End If
            If pVal.FormUID = "ARRI_NOT" Then
                oForm = SBO_Application.Forms.Item("ARRI_NOT")
                If pVal.BeforeAction = False And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oForm.Close()
                        oForm = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
