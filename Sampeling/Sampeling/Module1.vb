Imports SAPbouiCOM.Framework

Namespace Sampeling
    Module Module1

        <STAThread()>
        Sub Main(ByVal args() As String)

            Try

                Dim oApp As Application
                If (args.Length < 1) Then
                    oApp = New Application
                Else
                    oApp = New Application(args(0))
                End If

                Dim MyMenu As Menu
                MyMenu = New Menu()
                MyMenu.AddMenuItems()
                AddHandler Application.SBO_Application.AppEvent, AddressOf SBO_Application_AppEvent
                oApp.Run()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End Sub

        Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes)
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    System.Windows.Forms.Application.Exit()
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                Case SAPbouiCOM.BoAppEventTypes.aet_FontChanged
                Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
            End Select
        End Sub

    End Module
End Namespace